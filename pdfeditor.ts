/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument } from 'pdf-lib';
import * as fs from 'fs/promises';

/**
 * Switch: PDF Page Range Extractor / Remover + Hard Cropping
 *
 * Private Data:
 *  - pageRanges: REQUIRED. (same formats as before)
 *  - pageMode:   OPTIONAL. "keep" | "remove"  (default: "keep")
 *  - crop:       OPTIONAL. Either CSS-like margins OR exact size:
 *      Margins (in inches; CSS shorthand):
 *        crop=0.25
 *        crop=0.25,0.5
 *        crop=0.25,0.5,0.75
 *        crop=0.25,0.5,0.25,0.5   // top,right,bottom,left
 *
 *      Exact size (in inches; width × height, viewer-oriented):
 *        crop=8.5x11
 *        crop=11 x 8.5
 *        crop=8.5×11
 *
 * Behavior:
 *  - Margin form: reduces page to the cropped rectangle (no scaling).
 *  - Size form: page becomes exactly W×H, centered; smaller => clip, larger => pad.
 *  - Resulting page size is recognized by downstream tools (MediaBox updated; content translated; boxes synced).
 */

/** Utility to fetch private data as string */
async function pd(job: Job, key: string): Promise<string> {
  return (await job.getPrivateData(key)) as string;
}

/** Normalize various dash characters to hyphen */
function normalizeDashes(s: string): string {
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
    const cleanedWhole = normalizeDashes(text).replace(/^\[\s*|\s*\]$/g, '');

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

    const parts = cleanedWhole.split(',');
    for (let token of parts) {
      token = token.trim().replace(/^\[|\]$/g, '');
      if (!token) continue;

      let m = token.match(/^(\d+)\s*\-\s*(?:end)?$/i);
      if (!m) m = token.match(/^(\d+)\s*,$/);
      if (m) {
        const a = Math.max(1, parseInt(m[1], 10));
        for (let n = a; n <= pageCount; n++) out.add(n);
        continue;
      }

      m = token.match(/^\-\s*(\d+)$/) || token.match(/^,(\d+)$/);
      if (m) {
        const b = Math.min(pageCount, parseInt(m[1], 10));
        for (let n = 1; n <= b; n++) out.add(n);
        continue;
      }

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

      m = token.match(/^\d+$/);
      if (m) {
        const n = parseInt(token, 10);
        if (n <= 0) { await logger(`Ignoring non-positive page index: ${n}`); continue; }
        if (n > pageCount) { await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`); continue; }
        out.add(n);
        continue;
      }

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
  } else {
    for (const n of selection) keep.add(n);
  }
  const sorted = Array.from(keep).sort((a, b) => a - b);
  return sorted;
}

/** ---------- Cropping helpers ---------- **/

/** Parse CSS-like crop margins (in inches) into {top,right,bottom,left} in points */
function parseCropMarginsInchesToPoints(raw: string) {
  if (!raw) return null;
  let s = String(raw).trim();
  if (!s) return null;

  s = s.replace(/^crop\s*=/i, '').trim();

  // JSON array form like [0.25,0.5,0.25,0.5]
  let nums: number[] | null = null;
  if (s.startsWith('[') && s.endsWith(']')) {
    try {
      const arr = JSON.parse(s);
      if (Array.isArray(arr)) nums = arr.map((v) => Number(v));
      else if (typeof arr === 'number') nums = [Number(arr)];
    } catch { /* fall through */ }
  }

  if (!nums) {
    const tokens = s.split(/[\s,]+/).filter(Boolean);
    if (tokens.length === 1 && /[x×]/i.test(tokens[0])) return null; // looks like size, not margins
    nums = tokens
      .map((t) => t.replace(/in$/i, ''))
      .map((t) => Number(t))
      .filter((n) => Number.isFinite(n));
  }

  if (!nums || nums.length === 0) return null;

  const toPts = (inches: number) => Math.max(0, inches * 72);

  let T: number, R: number, B: number, L: number;
  if (nums.length === 1) {
    T = R = B = L = toPts(nums[0]);
  } else if (nums.length === 2) {
    T = B = toPts(nums[0]);
    R = L = toPts(nums[1]);
  } else if (nums.length === 3) {
    T = toPts(nums[0]); R = L = toPts(nums[1]); B = toPts(nums[2]);
  } else {
    T = toPts(nums[0]); R = toPts(nums[1]); B = toPts(nums[2]); L = toPts(nums[3]);
  }
  return { top: T, right: R, bottom: B, left: L };
}

/** Parse size syntax "<w>x<h>" (inches) into points */
function parseCropSizeInchesToPoints(raw: string) {
  if (!raw) return null;
  let s = String(raw).trim();
  if (!s) return null;

  s = s.replace(/^crop\s*=/i, '').trim();
  const m = s.match(/^\s*([\d.]+)\s*(?:in)?\s*[x×]\s*([\d.]+)\s*(?:in)?\s*$/i);
  if (!m) return null;

  const wIn = Number(m[1]);
  const hIn = Number(m[2]);
  if (!Number.isFinite(wIn) || !Number.isFinite(hIn)) return null;

  const toPts = (inches: number) => Math.max(1, inches * 72);
  return { width: toPts(wIn), height: toPts(hIn) };
}

/** Determine whether crop string is size syntax or margins; return a discriminated union */
function parseCropSpec(raw: string):
  | { kind: 'margins'; top: number; right: number; bottom: number; left: number }
  | { kind: 'size'; width: number; height: number }
  | null {
  if (!raw || !String(raw).trim()) return null;

  const size = parseCropSizeInchesToPoints(raw);
  if (size) return { kind: 'size', ...size };

  const margins = parseCropMarginsInchesToPoints(raw);
  if (margins) return { kind: 'margins', ...margins };

  return null;
}

/** Map viewer-oriented TRBL to underlying page orientation based on rotation (0/90/180/270) */
function mapCropForRotation(crop: { top: number; right: number; bottom: number; left: number }, angle: number) {
  const a = ((Math.round(angle) % 360) + 360) % 360;
  if (a === 0) return { ...crop };
  if (a === 90)  return { top: crop.right,  right: crop.bottom, bottom: crop.left,  left: crop.top };
  if (a === 180) return { top: crop.bottom, right: crop.left,   bottom: crop.top,   left: crop.right };
  if (a === 270) return { top: crop.left,   right: crop.top,    bottom: crop.right, left: crop.bottom };
  return { ...crop };
}

/** Sync all common boxes to 0,0,w,h when available */
function syncBoxesTo(page: any, w: number, h: number) {
  try { page.setCropBox?.(0, 0, w, h); } catch {}
  try { page.setMediaBox?.(0, 0, w, h); } catch {}
  try { page.setBleedBox?.(0, 0, w, h); } catch {}
  try { page.setTrimBox?.(0, 0, w, h); } catch {}
  try { page.setArtBox?.(0, 0, w, h); } catch {}
}

/** Hard-crop by margins: setSize + translateContent so the page’s new size == cropped size */
function hardCropByMargins(page: any, cropPtsViewer: { top: number; right: number; bottom: number; left: number }) {
  const rotation = page.getRotation?.()?.angle ?? 0;
  const crop = mapCropForRotation(cropPtsViewer, rotation);

  const origW = page.getWidth(), origH = page.getHeight();
  const L = Math.max(0, crop.left);
  const R = Math.max(0, crop.right);
  const T = Math.max(0, crop.top);
  const B = Math.max(0, crop.bottom);

  const newW = Math.max(1, origW - (L + R));
  const newH = Math.max(1, origH - (T + B));

  // Resize page and translate content so (L,B) becomes new origin (0,0)
  page.setSize(newW, newH);
  if (typeof page.translateContent === 'function') page.translateContent(-L, -B);

  syncBoxesTo(page, newW, newH);
}

/** Hard size viewport: set page to W×H, centered; translate to keep content centered */
function hardSetExactSize(page: any, widthPtsViewer: number, heightPtsViewer: number) {
  const rotation = page.getRotation?.()?.angle ?? 0;
  let W = Math.max(1, widthPtsViewer);
  let H = Math.max(1, heightPtsViewer);
  if ([90,270].includes(((Math.round(rotation)%360)+360)%360)) { const t = W; W = H; H = t; }

  const origW = page.getWidth(), origH = page.getHeight();

  // Move existing content by the center difference (works for smaller or larger)
  const dx = (W - origW) / 2;
  const dy = (H - origH) / 2;

  page.setSize(W, H);
  if (typeof page.translateContent === 'function') page.translateContent(dx, dy);

  syncBoxesTo(page, W, H);
}

/** ---------- Main ---------- **/

export async function jobArrived(_s: Switch, _f: FlowElement, job: Job) {
  try {
    const name = job.getName();
    if (!name.toLowerCase().endsWith('.pdf')) return job.fail('Not a PDF job');

    const log = async (msg: string) => { try { await job.log(LogLevel.Info, msg); } catch {} };

    const pageRangesRaw = await pd(job, 'pageRanges');
    const pageMode = (await pd(job, 'pageMode')) || 'keep';
    const cropRaw = await pd(job, 'crop'); // OPTIONAL: margins or WxH

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

    // --- Optional hard cropping / sizing ---
    const cropSpec = parseCropSpec(cropRaw);

    if (cropSpec?.kind === 'margins') {
      await log(`Hard-cropping by margins (in): top=${(cropSpec.top/72).toFixed(3)}, right=${(cropSpec.right/72).toFixed(3)}, bottom=${(cropSpec.bottom/72).toFixed(3)}, left=${(cropSpec.left/72).toFixed(3)} (rotation-aware)`);
      for (const page of outDoc.getPages()) hardCropByMargins(page, cropSpec);
    } else if (cropSpec?.kind === 'size') {
      await log(`Setting exact page size (in): ${(cropSpec.width/72).toFixed(3)} x ${(cropSpec.height/72).toFixed(3)} (centered; no scaling)`);
      for (const page of outDoc.getPages()) hardSetExactSize(page, cropSpec.width, cropSpec.height);
    } else if (cropRaw && String(cropRaw).trim()) {
      await log(`Warning: Could not parse 'crop' PD value: "${cropRaw}". No cropping applied.`);
    }

    // Save & send
    const outBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, outBytes);

    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e: any) {
    await job.fail(`Page range extractor error: ${e?.message || e}`);
  }
}
