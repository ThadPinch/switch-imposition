/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb, StandardFonts } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as http from 'http';

/* ---------- helpers ---------- */
const PT = 72, pt = (inch: number) => inch * PT;

type Layout = {
  sheetWIn: number;
  sheetHIn: number;
  cols: number;
  rows: number;
  cellWpt: number;
  cellHpt: number;
  gapHpt: number;
  gapVpt: number;
  offX: number;
  offY: number;
  waste: number; // empty cells
  orientation: 'portrait' | 'landscape';
};

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  const colsMax = Math.max(1, Math.floor((availW + gapH) / (cellW + gapH)));
  const rowsMax = Math.max(1, Math.floor((availH + gapV) / (cellH + gapV)));
  return { colsMax, rowsMax };
}

/** Draw crop marks for an individual placement */
function drawIndividualCrops(
  page: any,
  centerX: number,
  centerY: number,
  cutW: number,
  cutH: number,
  gapIn: number = 0.0625,
  lenIn: number = 0.125,
  strokePt: number = 0.5,
  isLeftEdge: boolean,
  isRightEdge: boolean,
  isBottomEdge: boolean,
  isTopEdge: boolean,
  gapHorizontal: number,
  gapVertical: number
) {
  const off = pt(gapIn);
  const perimeterLen = pt(lenIn);
  const maxInteriorLenH = Math.max(0, (gapHorizontal - off * 2) * 0.4);
  const maxInteriorLenV = Math.max(0, (gapVertical - off * 2) * 0.4);
  const interiorLenH = Math.min(pt(0.03125), maxInteriorLenH);
  const interiorLenV = Math.min(pt(0.03125), maxInteriorLenV);
  const k = rgb(0,0,0);

  const halfW = cutW / 2;
  const halfH = cutH / 2;
  const xL = centerX - halfW;
  const xR = centerX + halfW;
  const yB = centerY - halfH;
  const yT = centerY + halfH;

  const topLen = isTopEdge ? perimeterLen : interiorLenV;
  const leftLen = isLeftEdge ? perimeterLen : interiorLenH;
  page.drawLine({ start:{x:xL, y:yT + off}, end:{x:xL, y:yT + off + topLen}, thickness: strokePt, color: k });
  page.drawLine({ start:{x:xL - off - leftLen, y:yT}, end:{x:xL - off, y:yT}, thickness: strokePt, color: k });

  const rightLen = isRightEdge ? perimeterLen : interiorLenH;
  page.drawLine({ start:{x:xR, y:yT + off}, end:{x:xR, y:yT + off + topLen}, thickness: strokePt, color: k });
  page.drawLine({ start:{x:xR + off, y:yT}, end:{x:xR + off + rightLen, y:yT}, thickness: strokePt, color: k });

  const bottomLen = isBottomEdge ? perimeterLen : interiorLenV;
  page.drawLine({ start:{x:xL, y:yB - off}, end:{x:xL, y:yB - off - bottomLen}, thickness: strokePt, color: k });
  page.drawLine({ start:{x:xL - off - leftLen, y:yB}, end:{x:xL - off, y:yB}, thickness: strokePt, color: k });

  page.drawLine({ start:{x:xR, y:yB - off}, end:{x:xR, y:yB - off - bottomLen}, thickness: strokePt, color: k });
  page.drawLine({ start:{x:xR + off, y:yB}, end:{x:xR + off + rightLen, y:yB}, thickness: strokePt, color: k });
}

/** Draw a lavender overlay box with ID information (for cover) */
function drawLavenderOverlay(
  page: any,
  centerX: number,
  centerY: number,
  cutW: number,
  cutH: number,
  orderId: string,
  itemId: string,
  font: any,
  boldFont: any
) {
  const halfW = cutW / 2;
  const halfH = cutH / 2;
  const xL = centerX - halfW;
  const yB = centerY - halfH;

  const lavender = rgb(0.7, 0.5, 1);
  page.drawRectangle({
    x: xL,
    y: yB,
    width: cutW,
    height: cutH,
    color: lavender,
    opacity: 0.9
  });

  const white = rgb(1, 1, 1);
  const lineHeight = 14;

  const idText = `OrderID: ${orderId}`;
  const idSize = 12;
  const idWidth = boldFont.widthOfTextAtSize(idText, idSize);
  const idX = centerX - idWidth / 2;
  const idY = centerY + lineHeight / 2;

  page.drawText(idText, { x: idX, y: idY, size: idSize, font: boldFont, color: white });

  const itemText = `OrderItemID: ${itemId}`;
  const itemSize = 12;
  const itemWidth = font.widthOfTextAtSize(itemText, itemSize);
  const itemX = centerX - itemWidth / 2;
  const itemY = centerY - lineHeight / 2 - 4;

  page.drawText(itemText, { x: itemX, y: itemY, size: itemSize, font, color: white });
}

/** HTTP GET a PDF into a Uint8Array */
function httpGetBytes(url: string): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const req = http.get(url, (res) => {
      if (res.statusCode && res.statusCode >= 400) {
        reject(new Error(`HTTP ${res.statusCode} for ${url}`));
        return;
      }
      const chunks: Buffer[] = [];
      res.on('data', (d) => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(d)));
      res.on('end', () => resolve(new Uint8Array(Buffer.concat(chunks))));
    });
    req.on('error', reject);
  });
}

/** PLAN LAYOUT: uses fixed outer sheet margins of 0.125" and inter-cell gaps from payload; fits by CUT size */
function planLayout(
  sheetWIn: number,
  sheetHIn: number,
  orderItems: any[]
): Layout | null {
  const required = orderItems.length;

  const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
  const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
  if (!maxCutWIn || !maxCutHIn) return null;

  // Inter-cell gaps come from payload
  const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
  const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);

  // Fixed outer margins (sheet edges) = 0.125"
  const outerMarginHIn = 0.125;
  const outerMarginVIn = 0.125;

  const availWIn = sheetWIn - 2 * outerMarginHIn;
  const availHIn = sheetHIn - 2 * outerMarginVIn;

  const { colsMax, rowsMax } = gridFit(availWIn, availHIn, maxCutWIn, maxCutHIn, gapHIn, gapVIn);
  const maxPlacements = colsMax * rowsMax;
  if (maxPlacements < required) return null;

  let cols = Math.min(colsMax, required);
  let rowsNeeded = Math.ceil(required / cols);
  while (rowsNeeded > rowsMax && cols > 1) {
    cols -= 1;
    rowsNeeded = Math.ceil(required / cols);
  }
  if (rowsNeeded > rowsMax) return null;
  const rows = rowsNeeded;

  const cellWpt = pt(maxCutWIn);
  const cellHpt = pt(maxCutHIn);
  const gapHpt = pt(gapHIn);
  const gapVpt = pt(gapVIn);

  const arrWpt = cols * cellWpt + (cols - 1) * gapHpt;
  const arrHpt = rows * cellHpt + (rows - 1) * gapVpt;

  const sheetWpt = pt(sheetWIn);
  const sheetHpt = pt(sheetHIn);

  // Center the array inside fixed outer margins
  const offX = pt(outerMarginHIn) + (sheetWpt - pt(outerMarginHIn) * 2 - arrWpt) / 2;
  const offY = pt(outerMarginVIn) + (sheetHpt - pt(outerMarginVIn) * 2 - arrHpt) / 2;

  const waste = cols * rows - required;

  return {
    sheetWIn,
    sheetHIn,
    cols,
    rows,
    cellWpt,
    cellHpt,
    gapHpt,
    gapVpt,
    offX,
    offY,
    waste,
    orientation: sheetHIn >= sheetWIn ? 'portrait' : 'landscape'
  };
}

/** Decide 180° rotation for a given (row, col) based on item props */
function computeArtRotationDegrees(it: any, r: number, c: number): number {
  const mode = String(it.artRotation ?? 'None').trim().toLowerCase();
  const startRot = !!it.rotateFirstColumnOrRow;

  if (mode === 'rows') {
    const isRot = (r % 2 === 0) ? startRot : !startRot;
    return isRot ? 180 : 0;
  }
  if (mode === 'columns' || mode === 'cols' || mode === 'column') {
    const isRot = (c % 2 === 0) ? startRot : !startRot;
    return isRot ? 180 : 0;
  }
  return 0; // "None" or anything else
}

/** Adjust (x,y) so a 180° rotation keeps the artwork centered in its cell */
function adjustXYForRotation(x: number, y: number, width: number, height: number, deg: number) {
  const norm = ((deg % 360) + 360) % 360;
  if (norm === 180) {
    // pdf-lib rotates around the draw origin; translate by +W/+H to keep placement
    return { x: x + width, y: y + height };
  }
  return { x, y };
}

/** Create a cover page with artwork, crops, and overlay */
async function createCoverPage(
  outDoc: PDFDocument,
  layout: Layout,
  orderItems: any[],
  itemAssets: any[],
  perItemEmbeddedPages: Map<number, any[]>,
  placements: any[],
  pageIndex: number,
  font: any,
  boldFont: any
) {
  const sheetWpt = pt(layout.sheetWIn);
  const sheetHpt = pt(layout.sheetHIn);
  const page = outDoc.addPage([sheetWpt, sheetHpt]);

  // First draw crop marks
  for (const plc of placements) {
    const it = orderItems[plc.itemIdx];
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);

    const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

    const isLeftEdge = plc.c === 0;
    const isRightEdge = plc.c === layout.cols - 1;
    const isBottomEdge = plc.r === 0;
    const isTopEdge = plc.r === layout.rows - 1;

    drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i,
      0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, layout.gapHpt, layout.gapVpt);
  }

  // Then draw artwork from the specified page (with rotation logic)
  for (const plc of placements) {
    const asset = itemAssets[plc.itemIdx];
    if (pageIndex >= asset.pageCount) continue;

    const embeddedPages = perItemEmbeddedPages.get(asset.it.id as number)!;
    const it = asset.it;
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);
    const bleedWpt_i = pt((+it.bleedWidthInches || 0) || (+it.cutWidthInches || 0));
    const bleedHpt_i = pt((+it.bleedHeightInches || 0) || (+it.cutHeightInches || 0));
    const hasBleed_i = bleedWpt_i > cutWpt_i || bleedHpt_i > cutHpt_i;
    const placeW = hasBleed_i ? bleedWpt_i : cutWpt_i;
    const placeH = hasBleed_i ? bleedHpt_i : cutHpt_i;

    const ep = embeddedPages[Math.min(pageIndex, embeddedPages.length - 1)];

    const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;
    const x0 = cellCenterX - placeW / 2;
    const y0 = cellCenterY - placeH / 2;

    const rotDeg = computeArtRotationDegrees(it, plc.r, plc.c);
    const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rotDeg);

    page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(rotDeg) });
  }

  // Finally draw lavender overlays (not rotated)
  for (const plc of placements) {
    const it = orderItems[plc.itemIdx];
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);

    const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

    drawLavenderOverlay(
      page,
      cellCenterX,
      cellCenterY,
      cutWpt_i,
      cutHpt_i,
      String(it.orderId ?? ''),
      String(it.id ?? ''),
      font,
      boldFont
    );
  }
}

/* ---------- entry ---------- */
export async function jobArrived(_s: Switch, _f: FlowElement, job: Job) {
  try {
    const pd = async (k: string) => (await job.getPrivateData(k)) as string;
    let payloadRaw = await pd('payload');

    async function tryParse(text: string) {
      const cleaned = (text || '').replace(/^﻿/, '').trim();
      try { return JSON.parse(cleaned); } catch (e1:any) {
        if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
          const unwrapped = cleaned.slice(1, -1).replace(/\"/g, '"');
          try { return JSON.parse(unwrapped); } catch (e2:any) { /* fall through */ }
        }
        throw e1;
      }
    }

    let payload: any = null;

    if (payloadRaw) {
      try {
        payload = await tryParse(payloadRaw);
      } catch (e:any) {
        await job.log(LogLevel.Warning, `Failed to parse payload PD; will try reading asset. Error: ${e.message || e}`);
      }
    }

    if (!payload) {
      try {
        const inPath = await job.get(AccessLevel.ReadOnly);
        const buf = await fs.readFile(inPath);
        const asText = buf.toString('utf8');
        payload = await tryParse(asText);
      } catch (e:any) {
        return job.fail('Invalid JSON in payload private data and no parseable JSON from input asset');
      }
    }

    const orderItems: any[] = payload?.orderItems || [];
    if (!orderItems.length) return job.fail('No orderItems in payload');

    // ---- Diagnostics logging ----
    const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
    const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
    const maxBleedWIn = Math.max(...orderItems.map(it => (+it.bleedWidthInches || +it.cutWidthInches || 0)));
    const maxBleedHIn = Math.max(...orderItems.map(it => (+it.bleedHeightInches || +it.cutHeightInches || 0)));
    const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
    const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);
    
    // Get sheet dimensions from payload
    const impositionWidth = +(orderItems[0]?.impositionWidth || 19);
    const impositionHeight = +(orderItems[0]?.impositionHeight || 13);
    
    // NEW: Get explicit orientation from payload, or infer from dimensions
    const requestedOrientation = orderItems[0]?.impositionOrientation?.toLowerCase();
    let sheetWIn: number;
    let sheetHIn: number;
    let actualOrientation: string;
    
    if (requestedOrientation === 'portrait') {
      sheetWIn = Math.min(impositionWidth, impositionHeight);
      sheetHIn = Math.max(impositionWidth, impositionHeight);
      actualOrientation = 'portrait';
    } else if (requestedOrientation === 'landscape') {
      sheetWIn = Math.max(impositionWidth, impositionHeight);
      sheetHIn = Math.min(impositionWidth, impositionHeight);
      actualOrientation = 'landscape';
    } else {
      sheetWIn = impositionWidth;
      sheetHIn = impositionHeight;
      actualOrientation = sheetHIn > sheetWIn ? 'portrait' : 'landscape';
    }
    
    await job.log(LogLevel.Info, 
      `Sheet ${sheetWIn}x${sheetHIn} (${actualOrientation}${requestedOrientation ? ' - explicitly requested' : ' - inferred from dimensions'}); ` +
      `Cut ${maxCutWIn}x${maxCutHIn}; Bleed ${maxBleedWIn}x${maxBleedHIn}; ` +
      `Gaps H=${gapHIn} V=${gapVIn}; Items=${orderItems.length}`
    );
    await job.log(LogLevel.Info, `Outer sheet margins set to 0.125" (fixed).`);

    // Use the determined sheet dimensions directly - no orientation testing
    const layout = planLayout(sheetWIn, sheetHIn, orderItems);

    if (!layout) {
      return job.fail(`Items cannot fit on the specified ${sheetWIn}x${sheetHIn} sheet (${actualOrientation}) with current gaps and fixed 0.125" outer margins`);
    }

    await job.log(LogLevel.Info, `Impose ${layout.cols}x${layout.rows} on ${layout.sheetWIn}x${layout.sheetHIn} (${layout.orientation}). Empty cells: ${layout.waste}`);

    const sheetWpt = pt(layout.sheetWIn);
    const sheetHpt = pt(layout.sheetHIn);

    const baseUrl = 'http://10.1.0.79/api/switch/GetLocalArtwork/';

    // Load item assets (bytes + pageCount)
    const itemAssets = await Promise.all(orderItems.map(async (it) => {
      const url = `${baseUrl}${it.id}?pw=51ee6f3a3da5f642470202617cbcbd23`;
      let bytes: Uint8Array;
      try {
        bytes = await httpGetBytes(url);
      } catch (e:any) {
        try {
          const jobPath = await job.get(AccessLevel.ReadWrite);
          const jobDir = jobPath.substring(0, jobPath.lastIndexOf('/')+1);
          bytes = await fs.readFile(jobDir + (it.localArtworkPath || ''));
        } catch (e2) {
          throw new Error(`Failed to fetch art for item ${it.id}: ${e?.message || e}`);
        }
      }
      const srcDoc = await PDFDocument.load(bytes);
      const pageCount = srcDoc.getPageCount();
      return { it, bytes, pageCount };
    }));

    const maxPages = Math.max(...itemAssets.map(a => a.pageCount));

    const outDoc = await PDFDocument.create();

    const font = await outDoc.embedFont(StandardFonts.Helvetica);
    const boldFont = await outDoc.embedFont(StandardFonts.HelveticaBold);

    // --- BATCH EMBED PER ITEM: one call per item for all pages
    const perItemEmbeddedPages: Map<number, any[]> = new Map(); // itemId -> [embedded pages]

    async function ensureEmbeddedPagesForItem(itAsset: any) {
      const itId = itAsset.it.id as number;
      if (perItemEmbeddedPages.has(itId)) return;
      const idxs = Array.from({ length: itAsset.pageCount }, (_, i) => i);
      const embedded = await outDoc.embedPdf(itAsset.bytes, idxs);
      perItemEmbeddedPages.set(itId, embedded);
    }

    for (const asset of itemAssets) {
      await ensureEmbeddedPagesForItem(asset);
    }

    // placements in row-major order
    const placements: any[] = [];
    for (let r = 0; r < layout.rows; r++) {
      for (let c = 0; c < layout.cols; c++) {
        const idx = r * layout.cols + c;
        if (idx < orderItems.length) placements.push({ r, c, itemIdx: idx });
      }
    }

    // --- Cover pages based on inksBack ---
    const anyBackInks = orderItems.some(it => (+it.inksBack || 0) !== 0);
    const numCoverPages = anyBackInks ? 2 : 1;
    await job.log(LogLevel.Info, `Cover pages: ${numCoverPages} (inksBack ${anyBackInks ? 'non-zero detected' : 'all zero'})`);

    // Create cover pages (pageIndex 0, and 1 if needed)
    await createCoverPage(outDoc, layout, orderItems, itemAssets, perItemEmbeddedPages, placements, 0, font, boldFont);
    if (numCoverPages === 2) {
      await createCoverPage(outDoc, layout, orderItems, itemAssets, perItemEmbeddedPages, placements, 1, font, boldFont);
    }

    // Build imposed pages (normal production pages)
    for (let p = 0; p < maxPages; p++) {
      const page = outDoc.addPage([sheetWpt, sheetHpt]);

      // crop marks pass
      for (const plc of placements) {
        const it = orderItems[plc.itemIdx];
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);

        const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
        const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

        const isLeftEdge = plc.c === 0;
        const isRightEdge = plc.c === layout.cols - 1;
        const isBottomEdge = plc.r === 0;
        const isTopEdge = plc.r === layout.rows - 1;

        drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i,
          0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, layout.gapHpt, layout.gapVpt);
      }

      // artwork pass
      for (const plc of placements) {
        const asset = itemAssets[plc.itemIdx];
        if (p >= asset.pageCount) continue;

        const embeddedPages = perItemEmbeddedPages.get(asset.it.id as number)!;

        const it = asset.it;
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);
        const bleedWpt_i = pt((+it.bleedWidthInches || 0) || (+it.cutWidthInches || 0));
        const bleedHpt_i = pt((+it.bleedHeightInches || 0) || (+it.cutHeightInches || 0));
        const hasBleed_i = bleedWpt_i > cutWpt_i || bleedHpt_i > cutHpt_i;
        const placeW = hasBleed_i ? bleedWpt_i : cutWpt_i;
        const placeH = hasBleed_i ? bleedHpt_i : cutHpt_i;

        const ep = embeddedPages[Math.min(p, embeddedPages.length - 1)];

        const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
        const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;
        const x0 = cellCenterX - placeW / 2;
        const y0 = cellCenterY - placeH / 2;

        const rotDeg = computeArtRotationDegrees(it, plc.r, plc.c);
        const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rotDeg);

        page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(rotDeg) });
      }
    }

    // Save compactly and send
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const pdfBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, pdfBytes);

    const base = 'Batch-' + payload.batchId + '.pdf';
    if ((job as any).sendToSingle) await (job as any).sendToSingle(base);
    else job.sendTo(rwPath, 0, base);
  } catch (e:any) {
    await job.fail(`Batching impose error: ${e.message || e}`);
  }
}
