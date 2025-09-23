/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument } from 'pdf-lib';
import * as fs from 'fs/promises';

/**
 * Switch: PDF Page Range Extractor / Remover
 *
 * Purpose: Create a new PDF from an input PDF by either KEEPING or REMOVING
 * specified page ranges.
 *
 * Private Data (set on the job):
 *  - pageRanges: REQUIRED. Accepted formats (1-based, inclusive):
 *      • String list: "1-3, 5, 7-9"  (commas/spaces ok; en/em dashes ok)
 *      • Open-ended string ranges:
 *          - "3-" or "3," or "[3,]" → pages 3..end
 *          - "-5" or ",5" or "[,5]" → pages 1..5
 *          - "3-end" → pages 3..end
 *      • JSON array of numbers: "[1,3,5]"
 *      • JSON array of tuples:  "[[1,3],[5,5],[7,9]]"
 *      • JSON open-ended tuples: "[[3,null]]" or "[[null,5]]" or "[[3,\"end\"]]"
 *      • "all" (same as 1..N)
 *  - pageMode:   OPTIONAL. "keep" | "remove"  (default: "keep")
 *      • keep   → output contains ONLY pages listed by pageRanges
 *      • remove → output contains all pages EXCEPT those listed by pageRanges
 *
 * Examples:
 *  - pageRanges = "1-3,7,9-12"; pageMode = "keep"
 *      → Output pages: 1,2,3,7,9,10,11,12
 *  - pageRanges = "[[2,4],[10,10]]"; pageMode = "remove"
 *      → Output pages: all except 2,3,4,10
 *  - pageRanges = "[3,]" (or "3-" or "3-end"); pageMode = "keep"
 *      → Output pages: 3..end
 *
 * Notes:
 *  - Page numbers are 1-based in private data; internally converted to 0-based.
 *  - Invalid tokens, negative/zero indices, and reversed ranges are sanitized.
 *  - Out-of-bounds pages are ignored with a warning in the Switch log.
 */

/** Utility to fetch private data as string */
async function pd(job: Job, key: string): Promise<string> {
  return (await job.getPrivateData(key)) as string;
}

/** Normalize various dash characters to hyphen */
function normalizeDashes(s: string): string {
  // Replace: figure dash (‒), en dash (–), em dash (—), minus sign (−)
  return String(s || '').replace(/[‒–—−]/g, '-');
}

/** Convert mixed inputs to a page index, or special markers */
function toIndexOrMarker(v: any, _pageCount: number): number | null | 'END' {
  if (v == null || v === '') return null;
  if (typeof v === 'string') {
    const t = v.trim();
    if (/^end$/i.test(t)) return 'END';
    if (/^\d+$/.test(t)) return Math.floor(parseInt(t, 10));
    return NaN as any;
  }
  if (Number.isFinite(v)) return Math.floor(v as number);
  return NaN as any;
}

/** Parse pageRanges PD into a Set of 1-based page numbers */
function parsePageRanges(raw: string, pageCount: number, logger: (msg: string) => Promise<void>): Promise<Set<number>> {
  return new Promise(async (resolve) => {
    const out = new Set<number>();
    if (!raw || !String(raw).trim()) { resolve(out); return; }

    const text = String(raw).trim();

    // "all" shortcut
    if (/^all$/i.test(text)) {
      for (let i = 1; i <= pageCount; i++) out.add(i);
      resolve(out);
      return;
    }

    // --- Try JSON first (numbers, tuples, and open-ended tuples) ---
    try {
      const j = JSON.parse(text);
      const addNum = async (n: number) => {
        if (!Number.isFinite(n) || n <= 0) { await logger(`Ignoring invalid page index: ${n}`); return; }
        if (n > pageCount) { await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`); return; }
        out.add(Math.floor(n));
      };

      if (typeof j === 'number') { await addNum(j); resolve(out); return; }

      if (Array.isArray(j)) {
        for (const item of j) {
          if (typeof item === 'number') { await addNum(item); continue; }

          if (Array.isArray(item) && (item.length === 2 || item.length === 1)) {
            const a0 = item[0];
            const b0 = item.length === 2 ? item[1] : item[0];

            let a = toIndexOrMarker(a0, pageCount);
            let b = toIndexOrMarker(b0, pageCount);

            if (a === 'END') a = pageCount; // odd but supported
            if (b === 'END') b = pageCount;

            if (a == null) a = 1; // [,b] → 1..b
            if (b == null) b = pageCount; // [a,] → a..end

            if (!Number.isFinite(a as number) || !Number.isFinite(b as number)) {
              await logger(`Ignoring invalid range: ${JSON.stringify(item)}`);
              continue;
            }

            let A = Math.max(1, Math.floor(a as number));
            let B = Math.floor(b as number);
            if (A > B) { const t = A; A = B; B = t; }

            for (let n = A; n <= B; n++) {
              if (n > pageCount) { await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`); break; }
              out.add(n);
            }
            continue;
          }

          await logger(`Ignoring unrecognized JSON token in pageRanges: ${JSON.stringify(item)}`);
        }
        resolve(out);
        return;
      }
      // Fallthrough to string parsing if JSON is some other type
    } catch { /* not JSON, continue */ }

    // --- String list forms ---
    // Normalize dashes, and also allow a single surrounding pair of brackets like "[3,]"
    const cleanedWhole = normalizeDashes(text).replace(/^\[\s*|\s*\]$/g, '');

    // Special case: entire value like "[3,]" or "[,5]" after bracket strip becomes "3," or ",5"
    if (/^\d+\s*[,-]\s*$/.test(cleanedWhole)) {
      const m = cleanedWhole.match(/^(\d+)\s*[,-]\s*$/);
      if (m) {
        const a = Math.max(1, parseInt(m[1], 10));
        for (let n = a; n <= pageCount; n++) out.add(n);
        resolve(out);
        return;
      }
    }
    if (/^[,-]\s*\d+$/.test(cleanedWhole)) {
      const m = cleanedWhole.match(/^[,-]\s*(\d+)$/);
      if (m) {
        const b = Math.min(pageCount, parseInt(m[1], 10));
        for (let n = 1; n <= b; n++) out.add(n);
        resolve(out);
        return;
      }
    }

    // General comma-separated tokens
    const parts = cleanedWhole.split(',');
    for (let token of parts) {
      token = token.trim().replace(/^\[|\]$/g, ''); // drop per-token brackets if any
      if (!token) continue;

      // open-ended to END: "3-" or "3- end" or "3-" (also allow trailing comma caught as separate token)
      let m = token.match(/^(\d+)\s*\-\s*(?:end)?$/i);
      if (!m) m = token.match(/^(\d+)\s*,$/); // "3," token
      if (m) {
        const a = Math.max(1, parseInt(m[1], 10));
        for (let n = a; n <= pageCount; n++) out.add(n);
        continue;
      }

      // open-start from 1: "-5" or ",5"
      m = token.match(/^\-\s*(\d+)$/) || token.match(/^,(\d+)$/);
      if (m) {
        const b = Math.min(pageCount, parseInt(m[1], 10));
        for (let n = 1; n <= b; n++) out.add(n);
        continue;
      }

      // closed range: "a-b"
      m = token.match(/^(\d+)\s*\-\s*(\d+)$/);
      if (m) {
        let a = parseInt(m[1], 10), b = parseInt(m[2], 10);
        if (a > b) { const t = a; a = b; b = t; }
        for (let n = Math.max(1, a); n <= b; n++) {
          if (n > pageCount) { await logger(`Page ${n} beyond document page count (${pageCount}); ignoring remainder of range.`); break; }
          out.add(n);
        }
        continue;
      }

      // single page: "7"
      m = token.match(/^\d+$/);
      if (m) {
        const n = parseInt(token, 10);
        if (n <= 0) { await logger(`Ignoring non-positive page index: ${n}`); continue; }
        if (n > pageCount) { await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`); continue; }
        out.add(n);
        continue;
      }

      // fallback: "3-end" as a token
      m = token.match(/^(\d+)\s*\-\s*end$/i);
      if (m) {
        const a = Math.max(1, parseInt(m[1], 10));
        for (let n = a; n <= pageCount; n++) out.add(n);
        continue;
      }

      await logger(`Ignoring token in pageRanges: "${token}"`);
    }

    resolve(out);
  });
}

/** Compute final kept pages given mode */
function computeKeptPages(mode: string, selection: Set<number>, pageCount: number): number[] {
  const keep = new Set<number>();
  const m = String(mode || 'keep').toLowerCase();
  if (m === 'remove') {
    for (let i = 1; i <= pageCount; i++) if (!selection.has(i)) keep.add(i);
  } else { // default keep
    for (const n of selection) keep.add(n);
  }
  const sorted = Array.from(keep).sort((a, b) => a - b);
  return sorted;
}

export async function jobArrived(_s: Switch, _f: FlowElement, job: Job) {
  try {
    const name = job.getName();
    if (!name.toLowerCase().endsWith('.pdf')) return job.fail('Not a PDF job');

    const log = async (msg: string) => { try { await job.log(LogLevel.Info, msg); } catch {} };

    const pageRangesRaw = await pd(job, 'pageRanges');
    const pageMode = (await pd(job, 'pageMode')) || 'keep';

    if (!pageRangesRaw || !String(pageRangesRaw).trim())
      return job.fail('Missing required private data: pageRanges');

    // Read the input PDF
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const srcBytes = await fs.readFile(rwPath);

    const srcDoc = await PDFDocument.load(srcBytes);
    const pageCount = srcDoc.getPageCount();
    if (!pageCount) return job.fail('Source PDF has 0 pages');

    // Parse ranges (1-based)
    const selected = await parsePageRanges(pageRangesRaw, pageCount, log);
    const keptPages1Based = computeKeptPages(pageMode, selected, pageCount);

    if (keptPages1Based.length === 0)
      return job.fail(`No pages selected after applying pageMode="${pageMode}" and pageRanges.`);

    await log(`pageMode=${pageMode}; selected=${Array.from(selected).sort((a,b)=>a-b).join(',') || '(none)'}; keeping=${keptPages1Based.join(',')}`);

    // Create output by copying kept pages
    const outDoc = await PDFDocument.create();
    const indices0 = keptPages1Based.map(n => n - 1); // to 0-based

    const copied = await outDoc.copyPages(srcDoc, indices0);
    for (const p of copied) outDoc.addPage(p);

    // Save & send
    const outBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, outBytes);

    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e: any) {
    await job.fail(`Page range extractor error: ${e?.message || e}`);
  }
}
