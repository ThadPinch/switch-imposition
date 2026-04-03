/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees } from 'pdf-lib';
import * as fs from 'fs/promises';

/**
 * Switch: PDF Page Range Extractor / Remover + Hard Cropping
 *
 * HOW TO USE IN SWITCH
 * 1. Add this script to a flow element.
 * 2. Configure the flow element properties shown in Switch:
 *    - Page Ranges
 *    - Page Mode
 *    - Crop / Size
 *    - Resize
 *    - Match Orientation
 * 3. Send in a PDF job.
 * 4. The script edits the PDF in place and forwards the resulting job.
 *
 * INPUT REQUIREMENTS
 * - The incoming job must be a PDF.
 * - "Page Ranges" is required.
 * - All other properties are optional.
 *
 * FLOW ELEMENT PROPERTIES
 * - pageranges / "Page Ranges"
 *   Required. Defines which pages are selected before Page Mode is applied.
 *
 *   Supported formats:
 *   - all
 *   - 1
 *   - 1,3,5
 *   - 2-6
 *   - 6-2
 *   - 3-end
 *   - 3-
 *   - -5
 *   - ,5
 *   - JSON number: 4
 *   - JSON array: [1,3,5]
 *   - JSON ranges: [[1,3],[5,8]]
 *   - JSON open-ended ranges: [[3,null],[null,5],["end",2],[4,"end"]]
 *
 *   Notes:
 *   - Page numbers are 1-based.
 *   - Invalid tokens are ignored and logged.
 *   - Pages past the document length are ignored and logged.
 *
 * - pagemode / "Page Mode"
 *   Optional. Controls what happens with the selected pages.
 *   - keep   = keep only the pages from Page Ranges
 *   - remove = remove the pages from Page Ranges and keep the rest
 *   Default: keep
 *
 * - crop / "Crop / Size"
 *   Optional. Two modes are supported:
 *
 *   1. Margin crop mode
 *      Crops by viewer-oriented margins in inches, without scaling content.
 *      Accepted shorthand:
 *      - 0.25                 => all sides
 *      - 0.25,0.5             => top/bottom, right/left
 *      - 0.25,0.5,0.75        => top, right/left, bottom
 *      - 0.25,0.5,0.25,0.5    => top, right, bottom, left
 *      - [0.25,0.5,0.25,0.5]  => JSON array form
 *
 *   2. Exact size mode
 *      Sets the final page size in inches, centered, without scaling content.
 *      If the content is larger than the new size, it is clipped.
 *      If the content is smaller than the new size, padding is added.
 *      Accepted examples:
 *      - 8.5x11
 *      - 11 x 8.5
 *      - 8.5×11
 *
 * - resize / "Resize"
 *   Optional. Resizes each kept page to the given final size in inches by
 *   scaling the page contents and annotations.
 *
 *   Accepted examples:
 *   - 8.5x11
 *   - 11 x 8.5
 *   - 8.5×11
 *
 *   Behavior:
 *   - The page width and height are changed to the requested size.
 *   - Page contents are scaled to fill the new width and height.
 *   - If the aspect ratio changes, scaling is non-uniform.
 *   - Rotation is respected, so the size is interpreted in viewer orientation.
 *
 * - matchorientation / "Match Orientation"
 *   Optional. Uses a size string only to infer portrait vs landscape.
 *   Examples:
 *   - 8.5x11   => portrait target
 *   - 11x8.5   => landscape target
 *
 *   Behavior:
 *   - Pages that already match the target orientation are left unchanged.
 *   - Pages that do not match are rotated 90 degrees.
 *   - Square pages are left unchanged.
 *   - This setting does not resize pages.
 *
 * PROCESSING ORDER
 * 1. Parse Page Ranges.
 * 2. Apply Page Mode to determine the kept pages.
 * 3. Copy only the kept pages into a new PDF.
 * 4. Apply Crop / Size, if provided.
 * 5. Apply Resize, if provided.
 * 6. Apply Match Orientation, if provided.
 * 7. Save and forward the modified PDF.
 *
 * OUTPUT BEHAVIOR
 * - The output is a PDF.
 * - The script updates the page boxes so downstream tools see the new size.
 * - For crop/size operations, content is translated but not scaled.
 * - For resize operations, content and annotations are scaled to the new size.
 *
 * EXAMPLES
 * - Keep only pages 1 through 4:
 *   Page Ranges = 1-4
 *   Page Mode   = keep
 *
 * - Remove the first 2 pages:
 *   Page Ranges = 1-2
 *   Page Mode   = remove
 *
 * - Keep pages 3 through end and crop 0.25 in off every side:
 *   Page Ranges = 3-end
 *   Page Mode   = keep
 *   Crop / Size = 0.25
 *
 * - Keep all pages but force final size to 8.5x11 portrait:
 *   Page Ranges        = all
 *   Crop / Size        = 8.5x11
 *   Match Orientation  = 8.5x11
 *
 * - Keep all pages and resize everything to 5.5x8.5:
 *   Page Ranges = all
 *   Resize      = 5.5x8.5
 */

function getScriptTimeout() {
  return 300;
}

function getSettingsDefinition() {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<settings>' +
      '<setting name="pageranges" displayName="Page Ranges" type="multiline" required="Yes">' +
        '<description>Required. Examples: all, 1,3,5-7, 3-end, [1,[3,5],[8,null]]</description>' +
      '</setting>' +
      '<setting name="pagemode" displayName="Page Mode" type="enum" required="Yes">' +
        '<choices><choice value="keep"/><choice value="remove"/></choices>' +
        '<description>Keep the selected pages, or remove them and keep the rest.</description>' +
      '</setting>' +
      '<setting name="crop" displayName="Crop / Size" type="string" required="No">' +
        '<description>Optional. Margins in inches or an exact size like 8.5x11.</description>' +
      '</setting>' +
      '<setting name="resize" displayName="Resize" type="string" required="No">' +
        '<description>Optional. Final size like 8.5x11; scales page contents and annotations.</description>' +
      '</setting>' +
      '<setting name="matchorientation" displayName="Match Orientation" type="string" required="No">' +
        '<description>Optional. Size string like 8.5x11 used only to infer portrait vs landscape.</description>' +
      '</setting>' +
    '</settings>'
  );
}

function getDefaultSettings() {
  return {
    pageranges: '',
    pagemode: 'keep',
    crop: '',
    resize: '',
    matchorientation: ''
  };
}

async function tryGetProp(f: FlowElement, name: string): Promise<string> {
  try {
    if (f && typeof f.getPropertyValue === 'function') {
      let v = f.getPropertyValue(name);
      if (v && typeof v.then === 'function') v = await v;
      if (v != null) return String(v);
    }
  } catch {}

  try {
    if (f && typeof f.getPropertyStringValue === 'function') {
      let v2 = f.getPropertyStringValue(name);
      if (v2 && typeof v2.then === 'function') v2 = await v2;
      if (v2 != null) return String(v2);
    }
  } catch {}

  return '';
}

async function getVal(f: FlowElement, names: string[]): Promise<string> {
  for (const name of names) {
    const v = await tryGetProp(f, name);
    if (v) return v;
  }
  return '';
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

/** Parse page ranges into a Set of 1-based page numbers */
function parsePageRanges(raw: string, pageCount: number, logger: (msg: string) => Promise<void>): Promise<Set<number>> {
  return new Promise(async (resolve) => {
    const out = new Set<number>();
    if (!raw || !String(raw).trim()) {
      resolve(out);
      return;
    }

    const text = String(raw).trim();

    if (/^all$/i.test(text)) {
      for (let i = 1; i <= pageCount; i++) out.add(i);
      resolve(out);
      return;
    }

    try {
      const j = JSON.parse(text);
      const addNum = async (n: number) => {
        if (!Number.isFinite(n) || n <= 0) {
          await logger(`Ignoring invalid page index: ${n}`);
          return;
        }
        if (n > pageCount) {
          await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`);
          return;
        }
        out.add(Math.floor(n));
      };

      if (typeof j === 'number') {
        await addNum(j);
        resolve(out);
        return;
      }

      if (Array.isArray(j)) {
        for (const item of j) {
          if (typeof item === 'number') {
            await addNum(item);
            continue;
          }

          if (Array.isArray(item) && (item.length === 2 || item.length === 1)) {
            const a0 = item[0];
            const b0 = item.length === 2 ? item[1] : item[0];

            let a = toIndexOrMarker(a0, pageCount);
            let b = toIndexOrMarker(b0, pageCount);

            if (a === 'END') a = pageCount;
            if (b === 'END') b = pageCount;

            if (a == null) a = 1;
            if (b == null) b = pageCount;

            if (!Number.isFinite(a as number) || !Number.isFinite(b as number)) {
              await logger(`Ignoring invalid range: ${JSON.stringify(item)}`);
              continue;
            }

            let A = Math.max(1, Math.floor(a as number));
            let B = Math.floor(b as number);
            if (A > B) {
              const t = A;
              A = B;
              B = t;
            }

            for (let n = A; n <= B; n++) {
              if (n > pageCount) {
                await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`);
                break;
              }
              out.add(n);
            }
            continue;
          }

          await logger(`Ignoring unrecognized JSON token in pageRanges: ${JSON.stringify(item)}`);
        }

        resolve(out);
        return;
      }
    } catch {}

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
        let a = parseInt(m[1], 10);
        let b = parseInt(m[2], 10);
        if (a > b) {
          const t = a;
          a = b;
          b = t;
        }
        for (let n = Math.max(1, a); n <= b; n++) {
          if (n > pageCount) {
            await logger(`Page ${n} beyond document page count (${pageCount}); ignoring remainder of range.`);
            break;
          }
          out.add(n);
        }
        continue;
      }

      m = token.match(/^\d+$/);
      if (m) {
        const n = parseInt(token, 10);
        if (n <= 0) {
          await logger(`Ignoring non-positive page index: ${n}`);
          continue;
        }
        if (n > pageCount) {
          await logger(`Page ${n} beyond document page count (${pageCount}); ignoring.`);
          continue;
        }
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
    for (let i = 1; i <= pageCount; i++) {
      if (!selection.has(i)) keep.add(i);
    }
  } else {
    for (const n of selection) keep.add(n);
  }
  return Array.from(keep).sort((a, b) => a - b);
}

function parseCropMarginsInchesToPoints(raw: string) {
  if (!raw) return null;
  let s = String(raw).trim();
  if (!s) return null;

  s = s.replace(/^crop\s*=/i, '').trim();

  let nums: number[] | null = null;
  if (s.startsWith('[') && s.endsWith(']')) {
    try {
      const arr = JSON.parse(s);
      if (Array.isArray(arr)) nums = arr.map((v) => Number(v));
      else if (typeof arr === 'number') nums = [Number(arr)];
    } catch {}
  }

  if (!nums) {
    const tokens = s.split(/[\s,]+/).filter(Boolean);
    if (tokens.length === 1 && /[x×]/i.test(tokens[0])) return null;
    nums = tokens
      .map((t) => t.replace(/in$/i, ''))
      .map((t) => Number(t))
      .filter((n) => Number.isFinite(n));
  }

  if (!nums || nums.length === 0) return null;

  const toPts = (inches: number) => Math.max(0, inches * 72);

  let T: number;
  let R: number;
  let B: number;
  let L: number;
  if (nums.length === 1) {
    T = R = B = L = toPts(nums[0]);
  } else if (nums.length === 2) {
    T = B = toPts(nums[0]);
    R = L = toPts(nums[1]);
  } else if (nums.length === 3) {
    T = toPts(nums[0]);
    R = L = toPts(nums[1]);
    B = toPts(nums[2]);
  } else {
    T = toPts(nums[0]);
    R = toPts(nums[1]);
    B = toPts(nums[2]);
    L = toPts(nums[3]);
  }

  return { top: T, right: R, bottom: B, left: L };
}

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

function parseOrientationTarget(raw: string): 'portrait' | 'landscape' | null {
  if (!raw) return null;
  let s = String(raw).trim();
  if (!s) return null;
  s = s.replace(/^matchOrientation\s*=/i, '').trim();

  const m = s.match(/^\s*([\d.]+)\s*(?:in)?\s*[x×]\s*([\d.]+)\s*(?:in)?\s*$/i);
  if (!m) return null;

  const w = Number(m[1]);
  const h = Number(m[2]);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return null;
  if (w === h) return null;
  return w > h ? 'landscape' : 'portrait';
}

function getDisplayedPageOrientation(page: any): 'portrait' | 'landscape' | 'square' {
  const rot = ((Math.round(page.getRotation?.()?.angle ?? 0) % 360) + 360) % 360;
  let w = page.getWidth();
  let h = page.getHeight();
  if (rot === 90 || rot === 270) {
    const t = w;
    w = h;
    h = t;
  }
  if (w === h) return 'square';
  return w > h ? 'landscape' : 'portrait';
}

function matchPagesOrientation(doc: any, target: 'portrait' | 'landscape') {
  let rotated = 0;
  let unchanged = 0;
  for (const page of doc.getPages()) {
    const current = getDisplayedPageOrientation(page);
    if (current === 'square' || current === target) {
      unchanged++;
      continue;
    }
    const rot = ((Math.round(page.getRotation?.()?.angle ?? 0) % 360) + 360) % 360;
    page.setRotation(degrees((rot + 90) % 360));
    rotated++;
  }
  return { rotated, unchanged };
}

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

function mapCropForRotation(crop: { top: number; right: number; bottom: number; left: number }, angle: number) {
  const a = ((Math.round(angle) % 360) + 360) % 360;
  if (a === 0) return { ...crop };
  if (a === 90) return { top: crop.right, right: crop.bottom, bottom: crop.left, left: crop.top };
  if (a === 180) return { top: crop.bottom, right: crop.left, bottom: crop.top, left: crop.right };
  if (a === 270) return { top: crop.left, right: crop.top, bottom: crop.right, left: crop.bottom };
  return { ...crop };
}

function syncBoxesTo(page: any, w: number, h: number) {
  try { page.setCropBox?.(0, 0, w, h); } catch {}
  try { page.setMediaBox?.(0, 0, w, h); } catch {}
  try { page.setBleedBox?.(0, 0, w, h); } catch {}
  try { page.setTrimBox?.(0, 0, w, h); } catch {}
  try { page.setArtBox?.(0, 0, w, h); } catch {}
}

function hardCropByMargins(page: any, cropPtsViewer: { top: number; right: number; bottom: number; left: number }) {
  const rotation = page.getRotation?.()?.angle ?? 0;
  const crop = mapCropForRotation(cropPtsViewer, rotation);

  const origW = page.getWidth();
  const origH = page.getHeight();
  const L = Math.max(0, crop.left);
  const R = Math.max(0, crop.right);
  const T = Math.max(0, crop.top);
  const B = Math.max(0, crop.bottom);

  const newW = Math.max(1, origW - (L + R));
  const newH = Math.max(1, origH - (T + B));

  page.setSize(newW, newH);
  if (typeof page.translateContent === 'function') page.translateContent(-L, -B);

  syncBoxesTo(page, newW, newH);
}

function hardSetExactSize(page: any, widthPtsViewer: number, heightPtsViewer: number) {
  const rotation = page.getRotation?.()?.angle ?? 0;
  let W = Math.max(1, widthPtsViewer);
  let H = Math.max(1, heightPtsViewer);
  if ([90, 270].includes(((Math.round(rotation) % 360) + 360) % 360)) {
    const t = W;
    W = H;
    H = t;
  }

  const origW = page.getWidth();
  const origH = page.getHeight();
  const dx = (W - origW) / 2;
  const dy = (H - origH) / 2;

  page.setSize(W, H);
  if (typeof page.translateContent === 'function') page.translateContent(dx, dy);

  syncBoxesTo(page, W, H);
}

function hardResizeTo(page: any, widthPtsViewer: number, heightPtsViewer: number) {
  const rotation = page.getRotation?.()?.angle ?? 0;
  let targetW = Math.max(1, widthPtsViewer);
  let targetH = Math.max(1, heightPtsViewer);
  if ([90, 270].includes(((Math.round(rotation) % 360) + 360) % 360)) {
    const t = targetW;
    targetW = targetH;
    targetH = t;
  }

  const origW = Math.max(1, page.getWidth());
  const origH = Math.max(1, page.getHeight());
  const sx = targetW / origW;
  const sy = targetH / origH;

  if (typeof page.scaleContent === 'function') page.scaleContent(sx, sy);
  if (typeof page.scaleAnnotations === 'function') page.scaleAnnotations(sx, sy);
  page.setSize(targetW, targetH);

  syncBoxesTo(page, targetW, targetH);
}

export async function jobArrived(_s: Switch, f: FlowElement, job: Job) {
  try {
    const name = job.getName();
    if (!String(name || '').toLowerCase().endsWith('.pdf')) return job.fail('Not a PDF job');

    const log = async (msg: string) => {
      try {
        await job.log(LogLevel.Info, msg);
      } catch {}
    };

    const pageRangesRaw = await getVal(f, ['pageranges', 'pageRanges']);
    const pageModeRaw = await getVal(f, ['pagemode', 'pageMode']);
    const cropRaw = await getVal(f, ['crop']);
    const resizeRaw = await getVal(f, ['resize']);
    const matchOrientationRaw = await getVal(f, ['matchorientation', 'matchOrientation']);
    const pageMode = (pageModeRaw || 'keep').trim().toLowerCase() || 'keep';

    if (!pageRangesRaw || !String(pageRangesRaw).trim()) {
      return job.fail('Missing required flow element property: Page Ranges');
    }

    const rwPath = await job.get(AccessLevel.ReadWrite);
    const srcBytes = await fs.readFile(rwPath);

    const srcDoc = await PDFDocument.load(srcBytes);
    const pageCount = srcDoc.getPageCount();
    if (!pageCount) return job.fail('Source PDF has 0 pages');

    const selected = await parsePageRanges(pageRangesRaw, pageCount, log);
    const keptPages1Based = computeKeptPages(pageMode, selected, pageCount);

    if (keptPages1Based.length === 0) {
      return job.fail(`No pages selected after applying pageMode="${pageMode}" and pageRanges.`);
    }

    await log(
      `pageMode=${pageMode}; selected=${Array.from(selected).sort((a, b) => a - b).join(',') || '(none)'}; keeping=${keptPages1Based.join(',')}`
    );

    const outDoc = await PDFDocument.create();
    const indices0 = keptPages1Based.map((n) => n - 1);
    const copied = await outDoc.copyPages(srcDoc, indices0);
    for (const p of copied) outDoc.addPage(p);

    const cropSpec = parseCropSpec(cropRaw);
    if (cropSpec?.kind === 'margins') {
      await log(
        `Hard-cropping by margins (in): top=${(cropSpec.top / 72).toFixed(3)}, right=${(cropSpec.right / 72).toFixed(3)}, bottom=${(cropSpec.bottom / 72).toFixed(3)}, left=${(cropSpec.left / 72).toFixed(3)} (rotation-aware)`
      );
      for (const page of outDoc.getPages()) hardCropByMargins(page, cropSpec);
    } else if (cropSpec?.kind === 'size') {
      await log(
        `Setting exact page size (in): ${(cropSpec.width / 72).toFixed(3)} x ${(cropSpec.height / 72).toFixed(3)} (centered; no scaling)`
      );
      for (const page of outDoc.getPages()) hardSetExactSize(page, cropSpec.width, cropSpec.height);
    } else if (cropRaw && String(cropRaw).trim()) {
      await log(`Warning: Could not parse Crop / Size property value: "${cropRaw}". No cropping applied.`);
    }

    const resizeSpec = parseCropSizeInchesToPoints(resizeRaw);
    if (resizeSpec) {
      await log(
        `Resizing pages to (in): ${(resizeSpec.width / 72).toFixed(3)} x ${(resizeSpec.height / 72).toFixed(3)} (content and annotations scaled)`
      );
      for (const page of outDoc.getPages()) hardResizeTo(page, resizeSpec.width, resizeSpec.height);
    } else if (resizeRaw && String(resizeRaw).trim()) {
      await log(`Warning: Could not parse Resize property value: "${resizeRaw}". Expected "8.5x11".`);
    }

    const targetOrientation = parseOrientationTarget(matchOrientationRaw);
    if (targetOrientation) {
      const { rotated, unchanged } = matchPagesOrientation(outDoc, targetOrientation);
      await log(`matchOrientation target=${targetOrientation}; rotated=${rotated}; unchanged=${unchanged}`);
    } else if (matchOrientationRaw && String(matchOrientationRaw).trim()) {
      await log(
        `Warning: Could not parse Match Orientation property value: "${matchOrientationRaw}". Expected "8.5x11".`
      );
    }

    const outBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, outBytes);

    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e: any) {
    await job.fail(`PDF editor error: ${e?.message || e}`);
  }
}
