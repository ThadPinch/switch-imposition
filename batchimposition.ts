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

    // ---- Diagnostics logging (requested) ----
    const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
    const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
    const maxBleedWIn = Math.max(...orderItems.map(it => (+it.bleedWidthInches || +it.cutWidthInches || 0)));
    const maxBleedHIn = Math.max(...orderItems.map(it => (+it.bleedHeightInches || +it.cutHeightInches || 0)));
    const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
    const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);
    await job.log(LogLevel.Info, 
      `Sheet ${orderItems[0]?.impositionWidth || 19}x${orderItems[0]?.impositionHeight || 13}; ` +
      `Cut ${maxCutWIn}x${maxCutHIn}; Bleed ${maxBleedWIn}x${maxBleedHIn}; ` +
      `Gaps H=${gapHIn} V=${gapVIn}; Items=${orderItems.length}`
    );
    await job.log(LogLevel.Info, `Outer sheet margins set to 0.125" (fixed).`);

    // Sheet size from payload if present, else default 13x19
    const baseWIn = +(orderItems[0]?.impositionWidth || 19);
    const baseHIn = +(orderItems[0]?.impositionHeight || 13);

    // Try portrait (baseW x baseH) and landscape (swapped) and pick the best that fits
    const portrait = planLayout(baseWIn, baseHIn, orderItems);
    const landscape = planLayout(baseHIn, baseWIn, orderItems);

    let layout: Layout | null = null;
    if (portrait && !landscape) layout = portrait;
    else if (!portrait && landscape) layout = landscape;
    else if (portrait && landscape) {
      // Prefer fewer empty cells; tie-break on more columns (wider across)
      if (portrait.waste !== landscape.waste) layout = portrait.waste < landscape.waste ? portrait : landscape;
      else layout = landscape.cols > portrait.cols ? landscape : portrait;
    }

    if (!layout) return job.fail('Items cannot fit on the specified sheet in either orientation with current gaps and fixed 0.125" outer margins');

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
      // Single embedPdf call with the full page index array
      const embedded = await outDoc.embedPdf(itAsset.bytes, idxs);
      perItemEmbeddedPages.set(itId, embedded);
    }

    // placements in row-major order
    const placements: any[] = [];
    for (let r = 0; r < layout.rows; r++) {
      for (let c = 0; c < layout.cols; c++) {
        const idx = r * layout.cols + c;
        if (idx < orderItems.length) placements.push({ r, c, itemIdx: idx });
      }
    }

    // Build imposed pages
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

        await ensureEmbeddedPagesForItem(asset);
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
        const x = cellCenterX - placeW / 2;
        const y = cellCenterY - placeH / 2;

        page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(0) });
      }
    }

    // Create cover: duplicate first imposed page at front and overlay lavender boxes
    if (outDoc.getPageCount() > 0) {
      const [dup] = await outDoc.copyPages(outDoc, [0]); // duplicate first page
      outDoc.insertPage(0, dup);
      const cover = outDoc.getPage(0);

      for (const plc of placements) {
        const it = orderItems[plc.itemIdx];
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);

        const cellCenterX = layout.offX + plc.c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
        const cellCenterY = layout.offY + plc.r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

        drawLavenderOverlay(
          cover,
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
