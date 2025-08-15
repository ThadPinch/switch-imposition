/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb, StandardFonts } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as http from 'http';

/* ---------- helpers ---------- */
const PT = 72, pt = (inch: number) => inch * PT;

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  const colsMax = Math.max(1, Math.floor((availW + gapH) / (cellW + gapH)));
  const rowsMax = Math.max(1, Math.floor((availH + gapV) / (cellH + gapV)));
  return { colsMax, rowsMax };
}

/** Draw crop marks for an individual placement (same style as earlier script) */
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

/** Draw a lavender overlay box with ID information */
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

  // CHANGED: label + ensure we're using the item's orderId
  const idText = `OrderID: ${orderId}`;
  const idSize = 12;
  const idWidth = boldFont.widthOfTextAtSize(idText, idSize);
  const idX = centerX - idWidth / 2;
  const idY = centerY + lineHeight / 2;

  page.drawText(idText, {
    x: idX,
    y: idY,
    size: idSize,
    font: boldFont,
    color: white
  });

  const itemText = `OrderItemID: ${itemId}`;
  const itemSize = 12;
  const itemWidth = font.widthOfTextAtSize(itemText, itemSize);
  const itemX = centerX - itemWidth / 2;
  const itemY = centerY - lineHeight / 2 - 4;

  page.drawText(itemText, {
    x: itemX,
    y: itemY,
    size: itemSize,
    font: font,
    color: white
  });
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

/**
 * Build a reusable embedded page that carries its own CropBox.
 * MediaBox equals bleed (or cut if no bleed); CropBox = centered cut rectangle.
 */
async function buildPlacementXObject(
  outDoc: any,
  srcBytes: Uint8Array,
  pageIndex: number,
  cutWpt: number,
  cutHpt: number,
  bleedWpt: number,
  bleedHpt: number,
  hasBleed: boolean
) {
  const wrapper = await PDFDocument.create();
  const [srcEp] = await wrapper.embedPdf(srcBytes, [pageIndex]);

  const pageW = hasBleed ? bleedWpt : cutWpt;
  const pageH = hasBleed ? bleedHpt : cutHpt;
  const placementPage = wrapper.addPage([pageW, pageH]);

  const artX = (pageW - srcEp.width) / 2;
  const artY = (pageH - srcEp.height) / 2;
  placementPage.drawPage(srcEp, { x: artX, y: artY, rotate: degrees(0) });

  const cropX = hasBleed ? (bleedWpt - cutWpt) / 2 : 0;
  const cropY = hasBleed ? (bleedHpt - cutHpt) / 2 : 0;
  placementPage.setCropBox(cropX, cropY, cutWpt, cutHpt);

  const wrapperBytes = await wrapper.save();
  const [embeddedPlacement] = await outDoc.embedPdf(wrapperBytes, [0]);
  return embeddedPlacement;
}

/** Create a cover sheet page that duplicates the first imposed page and adds lavender overlays */
async function createCoverSheet(
  outDoc: any,
  firstPageBytes: Uint8Array,
  sheetWpt: number,
  sheetHpt: number,
  placements: any[],
  orderItems: any[],
  offX: number,
  offY: number,
  cellWpt: number,
  cellHpt: number,
  gapHpt: number,
  gapVpt: number,
  // CHANGED: removed global orderId param; we’ll use per-item orderId instead
  font: any,
  boldFont: any
) {
  const coverPage = outDoc.addPage([sheetWpt, sheetHpt]);

  const [bgPage] = await outDoc.embedPdf(firstPageBytes, [0]);
  coverPage.drawPage(bgPage, { x: 0, y: 0 });

  // CHANGED: pull orderId from each item
  for (const plc of placements) {
    const it = orderItems[plc.itemIdx];
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);

    const cellCenterX = offX + plc.c * (cellWpt + gapHpt) + cellWpt / 2;
    const cellCenterY = offY + plc.r * (cellHpt + gapVpt) + cellHpt / 2;

    drawLavenderOverlay(
      coverPage,
      cellCenterX,
      cellCenterY,
      cutWpt_i,
      cutHpt_i,
      String(it.orderId ?? ''),   // CHANGED
      String(it.id ?? ''),        // unchanged
      font,
      boldFont
    );
  }

  return coverPage;
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

    // CHANGED: remove broken top-level orderId read
    // const orderId = payload.orderId || '';  // <-- removed

    // Fixed press sheet size (inches)
    const sheetWIn = 13;
    const sheetHIn = 19;

    const sheetWpt = pt(sheetWIn);
    const sheetHpt = pt(sheetHIn);

    const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
    const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
    if (!maxCutWIn || !maxCutHIn) return job.fail('Invalid cut sizes in payload');

    const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
    const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);

    const gapHpt = pt(gapHIn);
    const gapVpt = pt(gapVIn);

    const availWIn = sheetWIn - 2 * gapHIn;
    const availHIn = sheetHIn - 2 * gapVIn;

    const { colsMax, rowsMax } = gridFit(availWIn, availHIn, maxCutWIn, maxCutHIn, gapHIn, gapVIn);
    if (colsMax < 1 || rowsMax < 1) return job.fail('Items cannot fit on 13x19 with current margins');

    let cols = Math.min(colsMax, orderItems.length);
    let rowsNeeded = Math.ceil(orderItems.length / cols);
    while (rowsNeeded > rowsMax && cols > 1) {
      cols -= 1;
      rowsNeeded = Math.ceil(orderItems.length / cols);
    }
    if (rowsNeeded > rowsMax) return job.fail(`Not enough space: need ${rowsNeeded} rows but only ${rowsMax} fit.`);
    const rows = rowsNeeded;

    const cellWpt = pt(maxCutWIn);
    const cellHpt = pt(maxCutHIn);

    const arrWpt = cols * cellWpt + (cols - 1) * gapHpt;
    const arrHpt = rows * cellHpt + (rows - 1) * gapVpt;

    const offX = pt(gapHIn) + (sheetWpt - pt(gapHIn) * 2 - arrWpt) / 2;
    const offY = pt(gapVIn) + (sheetHpt - pt(gapVIn) * 2 - arrHpt) / 2;

    const baseUrl = 'http://10.1.0.79/api/switch/GetLocalArtwork/';

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

    const embedCache: Map<number, Map<number, any>> = new Map();

    async function getEmbeddedFor(itAsset: any, pageIndex: number) {
      const itId = itAsset.it.id as number;
      if (!embedCache.has(itId)) embedCache.set(itId, new Map());
      const m = embedCache.get(itId)!;
      if (m.has(pageIndex)) return m.get(pageIndex);

      const cutWpt_i = pt(+itAsset.it.cutWidthInches || 0);
      const cutHpt_i = pt(+itAsset.it.cutHeightInches || 0);
      const bleedWpt_i = pt((+itAsset.it.bleedWidthInches || 0) || (+itAsset.it.cutWidthInches || 0));
      const bleedHpt_i = pt((+itAsset.it.bleedHeightInches || 0) || (+itAsset.it.cutHeightInches || 0));
      const hasBleed_i = bleedWpt_i > cutWpt_i || bleedHpt_i > cutHpt_i;

      const embedded = await buildPlacementXObject(
        outDoc,
        itAsset.bytes,
        Math.min(pageIndex, itAsset.pageCount - 1),
        cutWpt_i,
        cutHpt_i,
        bleedWpt_i,
        bleedHpt_i,
        hasBleed_i
      );
      m.set(pageIndex, embedded);
      return embedded;
    }

    const placements: any[] = [];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const idx = r * cols + c;
        if (idx < orderItems.length) placements.push({ r, c, itemIdx: idx });
      }
    }

    let firstPageBytes: Uint8Array | null = null;

    for (let p = 0; p < maxPages; p++) {
      const page = outDoc.addPage([sheetWpt, sheetHpt]);

      for (const plc of placements) {
        const it = orderItems[plc.itemIdx];
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);

        const cellCenterX = offX + plc.c * (cellWpt + gapHpt) + cellWpt / 2;
        const cellCenterY = offY + plc.r * (cellHpt + gapVpt) + cellHpt / 2;

        const isLeftEdge = plc.c === 0;
        const isRightEdge = plc.c === cols - 1;
        const isBottomEdge = plc.r === 0;
        const isTopEdge = plc.r === rows - 1;

        drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i,
          0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapHpt, gapVpt);
      }

      for (const plc of placements) {
        const asset = itemAssets[plc.itemIdx];
        if (p >= asset.pageCount) continue;

        const it = asset.it;
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);
        const bleedWpt_i = pt((+it.bleedWidthInches || 0) || (+it.cutWidthInches || 0));
        const bleedHpt_i = pt((+it.bleedHeightInches || 0) || (+it.cutHeightInches || 0));
        const hasBleed_i = bleedWpt_i > cutWpt_i || bleedHpt_i > cutHpt_i;
        const placeW = hasBleed_i ? bleedWpt_i : cutWpt_i;
        const placeH = hasBleed_i ? bleedHpt_i : cutHpt_i;

        const ep = await getEmbeddedFor(asset, p);

        const cellCenterX = offX + plc.c * (cellWpt + gapHpt) + cellWpt / 2;
        const cellCenterY = offY + plc.r * (cellHpt + gapVpt) + cellHpt / 2;
        const x = cellCenterX - placeW / 2;
        const y = cellCenterY - placeH / 2;

        page.drawPage(ep, { x, y, rotate: degrees(0) });
      }

      if (p === 0 && !firstPageBytes) {
        const tempDoc = await PDFDocument.create();
        const [copiedPage] = await tempDoc.copyPages(outDoc, [0]);
        tempDoc.addPage(copiedPage);
        firstPageBytes = await tempDoc.save();
      }
    }

    if (firstPageBytes) {
      const finalDoc = await PDFDocument.create();

      // CHANGED: call createCoverSheet without a global orderId
      await createCoverSheet(
        finalDoc,
        firstPageBytes,
        sheetWpt,
        sheetHpt,
        placements,
        orderItems,
        offX,
        offY,
        cellWpt,
        cellHpt,
        gapHpt,
        gapVpt,
        font,
        boldFont
      );

      const outDocBytes = await outDoc.save();
      const outDocForCopy = await PDFDocument.load(outDocBytes);
      const pagesToCopy = outDocForCopy.getPageCount();
      const copiedPages = await finalDoc.copyPages(outDocForCopy, Array.from({length: pagesToCopy}, (_, i) => i));
      copiedPages.forEach(page => finalDoc.addPage(page));

      const rwPath = await job.get(AccessLevel.ReadWrite);
      await fs.writeFile(rwPath, await finalDoc.save());

      const base = 'Batch-' + payload.batchId + '.pdf';
      if ((job as any).sendToSingle) await (job as any).sendToSingle(base);
      else job.sendTo(rwPath, 0, base);
    } else {
      const rwPath = await job.get(AccessLevel.ReadWrite);
      await fs.writeFile(rwPath, await outDoc.save());

      const base = 'Batch-' + payload.batchId + '.pdf';
      if ((job as any).sendToSingle) await (job as any).sendToSingle(base);
      else job.sendTo(rwPath, 0, base);
    }
  } catch (e:any) {
    await job.fail(`Batching impose error: ${e.message || e}`);
  }
}
