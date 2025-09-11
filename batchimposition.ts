/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb, StandardFonts } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as http from 'http';

// OPTIONAL: barcode (install with: npm i bwip-js)
let bwipjs: any = null;
try { bwipjs = require('bwip-js'); } catch { /* ok: fallback to text-only bug */ }

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
  waste: number;
  orientation: 'portrait' | 'landscape';
};

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  const colsMax = Math.max(1, Math.floor((availW + gapH) / (cellW + gapH)));
  const rowsMax = Math.max(1, Math.floor((availH + gapV) / (cellH + gapV)));
  return { colsMax, rowsMax };
}

/** Crops */
function drawIndividualCrops(
  page: any, centerX: number, centerY: number, cutW: number, cutH: number,
  gapIn: number = 0.0625, lenIn: number = 0.125, strokePt: number = 0.5,
  isLeftEdge: boolean, isRightEdge: boolean, isBottomEdge: boolean, isTopEdge: boolean,
  gapHorizontal: number, gapVertical: number
) {
  const off = pt(gapIn);
  const perimeterLen = pt(lenIn);
  const maxInteriorLenH = Math.max(0, (gapHorizontal - off * 2) * 0.4);
  const maxInteriorLenV = Math.max(0, (gapVertical - off * 2) * 0.4);
  const interiorLenH = Math.min(pt(0.03125), maxInteriorLenH);
  const interiorLenV = Math.min(pt(0.03125), maxInteriorLenV);
  const k = rgb(0,0,0);

  const halfW = cutW / 2, halfH = cutH / 2;
  const xL = centerX - halfW, xR = centerX + halfW;
  const yB = centerY - halfH, yT = centerY + halfH;

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

/** Lavender overlay (cover) */
function drawLavenderOverlay(
  page: any, centerX: number, centerY: number, cutW: number, cutH: number,
  orderId: string, itemId: string, font: any, boldFont: any
) {
  const halfW = cutW / 2, halfH = cutH / 2;
  const xL = centerX - halfW, yB = centerY - halfH;

  const lavender = rgb(0.7, 0.5, 1);
  page.drawRectangle({ x: xL, y: yB, width: cutW, height: cutH, color: lavender, opacity: 0.9 });

  const white = rgb(1, 1, 1);
  const idText = `OrderID: ${orderId}`;
  const itemText = `OrderItemID: ${itemId}`;
  const idSize = 12, itemSize = 12, lh = 14;
  const idW = boldFont.widthOfTextAtSize(idText, idSize);
  const itemW = font.widthOfTextAtSize(itemText, itemSize);

  page.drawText(idText,   { x: centerX - idW   / 2, y: centerY + lh / 2,     size: idSize,   font: boldFont, color: white });
  page.drawText(itemText, { x: centerX - itemW / 2, y: centerY - lh / 2 - 4, size: itemSize, font,          color: white });
}

/** HTTP GET a PDF */
function httpGetBytes(url: string): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const req = http.get(url, (res) => {
      if (res.statusCode && res.statusCode >= 400) { reject(new Error(`HTTP ${res.statusCode} for ${url}`)); return; }
      const chunks: Buffer[] = [];
      res.on('data', (d) => chunks.push(Buffer.isBuffer(d) ? d : Buffer.from(d)));
      res.on('end', () => resolve(new Uint8Array(Buffer.concat(chunks))));
    });
    req.on('error', reject);
  });
}

/** PLAN LAYOUT */
function planLayout(sheetWIn: number, sheetHIn: number, orderItems: any[]): Layout | null {
  const required = orderItems.length;

  const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
  const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
  if (!maxCutWIn || !maxCutHIn) return null;

  const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
  const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);

  const outerMarginHIn = 0.125, outerMarginVIn = 0.125;
  const availWIn = sheetWIn - 2 * outerMarginHIn;
  const availHIn = sheetHIn - 2 * outerMarginVIn;

  const { colsMax, rowsMax } = gridFit(availWIn, availHIn, maxCutWIn, maxCutHIn, gapHIn, gapVIn);
  const maxPlacements = colsMax * rowsMax;
  if (maxPlacements < required) return null;

  let cols = Math.min(colsMax, required);
  let rowsNeeded = Math.ceil(required / cols);
  while (rowsNeeded > rowsMax && cols > 1) { cols -= 1; rowsNeeded = Math.ceil(required / cols); }
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

  const offX = pt(outerMarginHIn) + (sheetWpt - pt(outerMarginHIn) * 2 - arrWpt) / 2;
  const offY = pt(outerMarginVIn) + (sheetHpt - pt(outerMarginVIn) * 2 - arrHpt) / 2;

  const waste = cols * rows - required;

  return { sheetWIn, sheetHIn, cols, rows, cellWpt, cellHpt, gapHpt, gapVpt, offX, offY, waste, orientation: sheetHIn >= sheetWIn ? 'portrait' : 'landscape' };
}

/** 0°/180° rotation rule */
function computeArtRotationDegrees(it: any, r: number, c: number): number {
  const mode = String(it.artRotation ?? 'None').trim().toLowerCase();
  const startRot = !!it.rotateFirstColumnOrRow;
  if (mode === 'rows')         return ((r % 2 === 0) ? startRot : !startRot) ? 180 : 0;
  if (mode.startsWith('col'))  return ((c % 2 === 0) ? startRot : !startRot) ? 180 : 0;
  return 0;
}

/** Keep placement centered when rotated 180° */
function adjustXYForRotation(x: number, y: number, width: number, height: number, deg: number) {
  const norm = ((deg % 360) + 360) % 360;
  if (norm === 180) return { x: x + width, y: y + height };
  return { x, y };
}

/** Convert desired sheet shift to pre-rotation delta (so +X/+Y always right/up on sheet) */
function preRotationShiftFor(deg: number, sxIn: number, syIn: number) {
  const sx = pt(sxIn || 0), sy = pt(syIn || 0);
  const norm = ((deg % 360) + 360) % 360;
  if (norm === 0)   return { sx,  sy };
  if (norm === 180) return { sx: -sx, sy: -sy };
  if (norm === 90)  return { sx: -sy, sy:  sx };
  if (norm === 270) return { sx:  sy, sy: -sx };
  return { sx, sy };
}

/** For duplex: flip column on every odd sheet page so backs align with fronts */
function effectiveCol(c: number, layout: Layout, flipPositionsThisPage: boolean) {
  return flipPositionsThisPage ? (layout.cols - 1 - c) : c;
}

/** For duplex: also invert rotation on odd pages (add 180°) */
function rotationForPage(it: any, rEff: number, cEff: number, flipPositionsThisPage: boolean) {
  let deg = computeArtRotationDegrees(it, rEff, cEff);
  if (flipPositionsThisPage) deg = (deg + 180) % 360;
  return deg;
}

/* ---------- Gutter Bug ---------- */

type BugSide = 'left'|'right'|'top'|'bottom';
const BUG_THICKNESS_IN = 0.125;

function availableGapForSide(
  side: BugSide, layout: Layout, r: number, c: number,
  xL: number, xR: number, yB: number, yT: number,
  sheetWpt: number, sheetHpt: number
) {
  if (side === 'left')   return c > 0 ? layout.gapHpt : xL;
  if (side === 'right')  return c < layout.cols-1 ? layout.gapHpt : (sheetWpt - xR);
  if (side === 'bottom') return r > 0 ? layout.gapVpt : yB;
  if (side === 'top')    return r < layout.rows-1 ? layout.gapVpt : (sheetHpt - yT);
  return 0;
}

function pickBugSide(
  pos: string, layout: Layout, r: number, c: number,
  xL: number, xR: number, yB: number, yT: number,
  sheetWpt: number, sheetHpt: number
): BugSide | null {
  const needPt = pt(BUG_THICKNESS_IN);
  const norm = String(pos || 'Inside').toLowerCase();

  const gaps: Record<BugSide, number> = {
    left:  availableGapForSide('left', layout, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt),
    right: availableGapForSide('right', layout, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt),
    bottom:availableGapForSide('bottom', layout, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt),
    top:   availableGapForSide('top', layout, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt),
  };
  const distances: Record<BugSide, number> = { left: xL, right: sheetWpt - xR, bottom: yB, top: sheetHpt - yT };

  function pickWithinAxis(axis: 'h'|'v', inside: boolean): BugSide | null {
    const candidates: BugSide[] = axis === 'h' ? ['left','right'] : ['bottom','top'];
    const viable = candidates.filter(s => gaps[s] >= needPt);
    if (!viable.length) return null;
    viable.sort((a,b) => inside ? (distances[b] - distances[a]) : (distances[a] - distances[b]));
    return viable[0] ?? null;
  }

  if (norm === 'left' || norm === 'right' || norm === 'top' || norm === 'bottom') {
    return gaps[norm as BugSide] >= needPt ? (norm as BugSide) : null;
  }

  const horizMax = Math.max(gaps.left, gaps.right);
  const vertMax  = Math.max(gaps.top, gaps.bottom);
  const horizOK = horizMax >= needPt;
  const vertOK  = vertMax  >= needPt;

  const inside = norm === 'inside';
  let pick: BugSide | null = null;

  if (horizOK && vertOK) {
    const axis = (horizMax >= vertMax) ? 'h' : 'v';
    pick = pickWithinAxis(axis, inside) ?? pickWithinAxis(axis === 'h' ? 'v' : 'h', inside);
  } else if (horizOK) pick = pickWithinAxis('h', inside);
  else if (vertOK)    pick = pickWithinAxis('v', inside);

  return pick;
}

/** Generate barcode PNG; when vertical=true, we make a *tall* barcode (bars vertical) */
async function makeBarcodePngBytes(text: string, vertical: boolean): Promise<Uint8Array | null> {
  if (!bwipjs) return null;
  try {
    const opts: any = {
      bcid: 'code128',
      text,
      scale: 2,
      height: 8,
      includetext: false,
      textxalign: 'center',
      backgroundcolor: 'FFFFFF'
    };
    if (vertical) opts.rotate = 'R';   // 90° right – barcode becomes tall
    const buf: Buffer = await bwipjs.toBuffer(opts);
    return new Uint8Array(buf);
  } catch { return null; }
}

/** Short text renderer for bug strip (auto-fit; supports vertical) */
function drawBugText(
  page: any, font: any, text: string,
  bx: number, by: number, bw: number, bh: number,
  vertical: boolean, anchorLeft: boolean
) {
  if (!text) return;
  const k = rgb(0,0,0);
  const maxThickness = vertical ? bw : bh;
  const maxLong = vertical ? bh : bw;

  let size = Math.min(6, maxThickness * 0.40);
  while (font.widthOfTextAtSize(text, size) > (maxLong - 2) && size > 3) size -= 0.25;

  if (!vertical) {
    const x = anchorLeft ? bx : (bx + (bw - font.widthOfTextAtSize(text, size)) / 2);
    const y = by + (bh - size) / 2;
    page.drawText(text, { x, y, size, font, color: k });
  } else {
    const x = bx + (bw - size) / 2;
    const y = by + bh - font.widthOfTextAtSize(text, size); // start near far end; flows “away”
    page.drawText(text, { x, y, size, font, color: k, rotate: degrees(90) });
  }
}

/** Draw short text inside a rectangular area. */
function drawBugTextInArea(
  page: any,
  font: any,
  text: string,
  x: number, y: number, w: number, h: number,
  vertical: boolean,
  anchorStart: boolean
) {
  if (!text) return;
  const k = rgb(0,0,0);
  const maxThickness = vertical ? w : h;
  const maxAlong     = vertical ? h : w;
  let size = Math.min(6, maxThickness * 0.40);
  while (font.widthOfTextAtSize(text, size) > (maxAlong - 2) && size > 3) size -= 0.25;

  if (!vertical) {
    const tx = anchorStart ? x : (x + (w - font.widthOfTextAtSize(text, size)) / 2);
    const ty = y + (h - size) / 2;
    page.drawText(text, { x: tx, y: ty, size, font, color: k });
  } else {
    const tx = x + (w - size) / 2;
    const ty = anchorStart ? y : (y + (h - font.widthOfTextAtSize(text, size)) / 2);
    page.drawText(text, { x: tx, y: ty, size, font, color: k, rotate: degrees(90) });
  }
}

/** Draw gutter bug (and optional barcode) – CENTERED ON THE ART EDGE */
async function drawGutterBug(
  page: any,
  outDoc: PDFDocument,
  it: any,
  layout: Layout,
  r: number,
  c: number,
  placeW: number,
  placeH: number,
  sheetWpt: number,
  sheetHpt: number,
  barcodeCache: Map<string, any>,
  font: any
) {
  if (!it.includeGutterBug) return;

  // Edges of the ARTWORK at the size it was placed (bleed if present)
  const cellCenterX = layout.offX + c * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
  const cellCenterY = layout.offY + r * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

  const xL = cellCenterX - placeW / 2;
  const xR = cellCenterX + placeW / 2;
  const yB = cellCenterY - placeH / 2;
  const yT = cellCenterY + placeH / 2;

  const side = pickBugSide(String(it.bugPosition), layout, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt);
  if (!side) return;

  const thickPt = pt(BUG_THICKNESS_IN);
  let bx = xL, by = yB, bw = placeW, bh = thickPt, vertical = false;
  if (side === 'top')    { bx = xL;           by = yT;           bw = placeW;   bh = thickPt; vertical = false; }
  if (side === 'bottom') { bx = xL;           by = yB - thickPt; bw = placeW;   bh = thickPt; vertical = false; }
  if (side === 'left')   { bx = xL - thickPt; by = yB;           bw = thickPt;  bh = placeH;  vertical = true;  }
  if (side === 'right')  { bx = xR;           by = yB;           bw = thickPt;  bh = placeH;  vertical = true;  }

  // White strip
  page.drawRectangle({ x: bx, y: by, width: bw, height: bh, color: rgb(1,1,1), opacity: 1 });

  const wantBarcode = !!it.includeAutoShipBarcodeInBug;
  const pathText = String(it.localArtworkPath || '');
  const PAD = 2; // pts gap at center

  if (!vertical) {
    const centerX = cellCenterX;
    const leftW  = Math.max(0, centerX - bx - PAD);
    const rightW = Math.max(0, bx + bw - (centerX + PAD));
    const rightX = centerX + PAD;

    if (wantBarcode) {
      const key = `${it.orderId}-${it.id}-H`;
      let img = barcodeCache.get(key);
      if (!img) {
        const bytes = await makeBarcodePngBytes(`${it.orderId}-${it.id}`, false);
        if (bytes) { img = await outDoc.embedPng(bytes); barcodeCache.set(key, img); }
      }
      if (img && leftW > 0) {
        const iw = img.width, ih = img.height;
        const scale = Math.min(leftW / iw, bh / ih);
        const w = iw * scale, h = ih * scale;
        const x = centerX - PAD - w;
        const y = by + (bh - h) / 2;
        page.drawImage(img, { x, y, width: w, height: h });
      }
    }

    drawBugTextInArea(page, font, pathText, rightX, by, rightW, bh, false, true);

  } else {
    const centerY = cellCenterY;
    const botH = Math.max(0, centerY - by - PAD);
    const topH = Math.max(0, by + bh - (centerY + PAD));
    const topY = centerY + PAD;

    if (wantBarcode) {
      const key = `${it.orderId}-${it.id}-V`;
      let img = barcodeCache.get(key);
      if (!img) {
        const bytes = await makeBarcodePngBytes(`${it.orderId}-${it.id}`, true);
        if (bytes) { img = await outDoc.embedPng(bytes); barcodeCache.set(key, img); }
      }
      if (img && botH > 0) {
        const iw = img.width, ih = img.height;
        const scale = Math.min(bw / iw, botH / ih);
        const w = iw * scale, h = ih * scale;
        const x = bx + (bw - w) / 2;
        const y = centerY - PAD - h;
        page.drawImage(img, { x, y, width: w, height: h });
      }
    }

    drawBugTextInArea(page, font, pathText, bx, topY, bw, topH, true, true);
  }
}


/** COVER PAGE */
async function createCoverPage(
  outDoc: PDFDocument, layout: Layout, orderItems: any[], itemAssets: any[],
  perItemEmbeddedPages: Map<number, any[]>, placements: any[], pageIndex: number,
  font: any, boldFont: any, flipPositionsThisPage: boolean
) {
  const sheetWpt = pt(layout.sheetWIn), sheetHpt = pt(layout.sheetHIn);
  const page = outDoc.addPage([sheetWpt, sheetHpt]);

  // Crops
  for (const plc of placements) {
    const it = orderItems[plc.itemIdx];
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);

    const cEff = effectiveCol(plc.c, layout, flipPositionsThisPage);
    const rEff = plc.r;

    const cellCenterX = layout.offX + cEff * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + rEff * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

    const isLeftEdge = cEff === 0, isRightEdge = cEff === layout.cols - 1;
    const isBottomEdge = rEff === 0, isTopEdge = rEff === layout.rows - 1;

    drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i, 0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, layout.gapHpt, layout.gapVpt);
  }

  // Artwork
  for (const plc of placements) {
    const asset = itemAssets[plc.itemIdx];
    if (pageIndex >= asset.pageCount) continue;
    const embeddedPages = perItemEmbeddedPages.get(asset.it.id as number)!;
    const it = asset.it;

    const cutWpt_i = pt(+it.cutWidthInches || 0), cutHpt_i = pt(+it.cutHeightInches || 0);
    const bleedWpt_i = pt((+it.bleedWidthInches || 0) || (+it.cutWidthInches || 0));
    const bleedHpt_i = pt((+it.bleedHeightInches || 0) || (+it.cutHeightInches || 0));
    const hasBleed_i = bleedWpt_i > cutWpt_i || bleedHpt_i > cutHpt_i;
    const placeW = hasBleed_i ? bleedWpt_i : cutWpt_i;
    const placeH = hasBleed_i ? bleedHpt_i : cutHpt_i;

    const ep = embeddedPages[Math.min(pageIndex, embeddedPages.length - 1)];

    const cEff = effectiveCol(plc.c, layout, flipPositionsThisPage);
    const rEff = plc.r;

    const cellCenterX = layout.offX + cEff * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + rEff * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

    const rotDeg = rotationForPage(it, rEff, cEff, flipPositionsThisPage);

    // Invert X shift for odd pages (pre-rotation)
    const inputShiftX = flipPositionsThisPage ? -(+it.imageShiftX || 0) : (+it.imageShiftX || 0);
    const { sx, sy } = preRotationShiftFor(rotDeg, inputShiftX, +it.imageShiftY);

    const x0 = cellCenterX - placeW / 2 + sx;
    const y0 = cellCenterY - placeH / 2 + sy;
    const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rotDeg);

    page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(rotDeg) });
  }

  // Overlays
  for (const plc of placements) {
    const it = orderItems[plc.itemIdx];
    const cutWpt_i = pt(+it.cutWidthInches || 0);
    const cutHpt_i = pt(+it.cutHeightInches || 0);

    const cEff = effectiveCol(plc.c, layout, flipPositionsThisPage);
    const rEff = plc.r;

    const cellCenterX = layout.offX + cEff * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
    const cellCenterY = layout.offY + rEff * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

    drawLavenderOverlay(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i, String(it.orderId ?? ''), String(it.id ?? ''), font, boldFont);
  }
}

function clamp(n: number, lo: number, hi: number) { return Math.max(lo, Math.min(hi, n)); }
function normDeg(d: number) { return ((d % 360) + 360) % 360; }
function deg2rad(d: number) { return normDeg(d) * Math.PI / 180; }


// Default size for in-art barcode (clamped to the cut area)
const IN_ART_BARCODE_DEFAULT_W_IN = 1.00;
const IN_ART_BARCODE_DEFAULT_H_IN = 0.25;

/**
 * Draw an in-art AutoShip barcode on the last production sheet.
 * X/Y are in inches from the CUT-BOX lower-left in *art space* (unrotated).
 * We rotate that intended position about the CUT center by rotDeg so the
 * barcode stays in the same place on the artwork no matter the placement rotation.
 */
async function drawInArtBarcodeOnLastSheet(
  page: any,
  outDoc: PDFDocument,
  it: any,
  isLastProductionSheet: boolean,
  rotDeg: number,
  x0: number, y0: number,             // pre-rotation artwork bottom-left (PLACED area)
  cutWpt: number, cutHpt: number,     // cut size in pts
  bleedWpt: number, bleedHpt: number, // bleed size in pts (or = cut)
  barcodeCache: Map<string, any>,
  font: any
) {
  if (!isLastProductionSheet) return;
  if (!it?.includeAutoShipBarcodeInArtOnLastSheet) return;

  // Cut-box offset inside placed area (if bleed is present, cut sits centered inside bleed)
  const cutOffX = Math.max(0, (bleedWpt - cutWpt) / 2);
  const cutOffY = Math.max(0, (bleedHpt - cutHpt) / 2);

  // User-specified offsets from CUT lower-left (in points)
  const inXptRaw = pt(+it.inArtBarcodeX || 0);
  const inYptRaw = pt(+it.inArtBarcodeY || 0);

  // Constrain negative inputs defensively
  const inXpt = Math.max(0, inXptRaw);
  const inYpt = Math.max(0, inYptRaw);

  // Caps so the barcode cannot exceed what remains from (inX,inY) to the cut edges (0° case).
  const capW = Math.max(2, cutWpt - inXpt);
  const capH = Math.max(2, cutHpt - inYpt);

  // Target size, clamped to caps above
  const targetW = Math.min(pt(IN_ART_BARCODE_DEFAULT_W_IN), capW);
  const targetH = Math.min(pt(IN_ART_BARCODE_DEFAULT_H_IN), capH);

  // Barcode image (horizontal); we will rotate with the art by rotDeg
  const cacheKey = `${it.orderId}-${it.id}-INART-H`;
  let img = barcodeCache.get(cacheKey);
  if (!img) {
    const bytes = await makeBarcodePngBytes(`${it.orderId}-${it.id}`, /*vertical=*/false);
    if (bytes) { img = await outDoc.embedPng(bytes); barcodeCache.set(cacheKey, img); }
  }

  // Compute size (w,h) with aspect respected; fall back to text if needed
  let w = targetW, h = targetH, useTextFallback = !img;

  if (img) {
    const iw = img.width, ih = img.height;
    const scale = Math.max(0.01, Math.min(targetW / iw, targetH / ih));
    w = iw * scale; h = ih * scale;
  }

  // CUT center in sheet coords
  const cutLLx = x0 + cutOffX;
  const cutLLy = y0 + cutOffY;
  const cutCx  = cutLLx + cutWpt / 2;
  const cutCy  = cutLLy + cutHpt / 2;

  // Intended *center* of the barcode in unrotated (art-space) coords:
  // Start at CUT LL, add user offsets, then add half the barcode size
  const centerX0 = cutLLx + inXpt + w / 2;
  const centerY0 = cutLLy + inYpt + h / 2;

  // Rotate that center about the CUT center by rotDeg
  const theta = deg2rad(rotDeg);
  const cosT = Math.cos(theta), sinT = Math.sin(theta);
  const vx = centerX0 - cutCx;
  const vy = centerY0 - cutCy;
  const centerX = cutCx + (vx * cosT - vy * sinT);
  const centerY = cutCy + (vx * sinT + vy * cosT);

  // Convert center -> bottom-left draw point for a rectangle rotated by rotDeg
  // bottom-left = center - R * (w/2, h/2)
  const halfWx = (w / 2) * cosT - (h / 2) * sinT;
  const halfWy = (w / 2) * sinT + (h / 2) * cosT;
  const drawX = centerX - halfWx;
  const drawY = centerY - halfWy;

  if (!useTextFallback) {
    page.drawImage(img, { x: drawX, y: drawY, width: w, height: h, rotate: degrees(normDeg(rotDeg)) });
    return;
  }

  // ---- Text fallback if bwip-js not present ----
  const text = `${it.orderId}-${it.id}`;
  // Fit text to target box width (unrotated width constraint in art-space); conservative fit
  let size = 8;
  const maxWForText = targetW;
  while (size > 5 && font.widthOfTextAtSize(text, size) > maxWForText) size -= 0.5;
  const tw = font.widthOfTextAtSize(text, size);
  const th = size;

  // Recompute using text box w/h
  const centerX0_txt = cutLLx + inXpt + tw / 2;
  const centerY0_txt = cutLLy + inYpt + th / 2;
  const vx2 = centerX0_txt - cutCx;
  const vy2 = centerY0_txt - cutCy;
  const centerX_txt = cutCx + (vx2 * cosT - vy2 * sinT);
  const centerY_txt = cutCy + (vx2 * sinT + vy2 * cosT);
  const halfWx2 = (tw / 2) * cosT - (th / 2) * sinT;
  const halfWy2 = (tw / 2) * sinT + (th / 2) * cosT;
  const drawX_txt = centerX_txt - halfWx2;
  const drawY_txt = centerY_txt - halfWy2;

  page.drawText(text, { x: drawX_txt, y: drawY_txt, size, font, color: rgb(0,0,0), rotate: degrees(normDeg(rotDeg)) });
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
      try { payload = await tryParse(payloadRaw); }
      catch (e:any) {
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

    // ---- Diagnostics ----
    const maxCutWIn = Math.max(...orderItems.map(it => +it.cutWidthInches || 0));
    const maxCutHIn = Math.max(...orderItems.map(it => +it.cutHeightInches || 0));
    const maxBleedWIn = Math.max(...orderItems.map(it => (+it.bleedWidthInches || +it.cutWidthInches || 0)));
    const maxBleedHIn = Math.max(...orderItems.map(it => (+it.bleedHeightInches || +it.cutHeightInches || 0)));
    const gapHIn = Math.max(...orderItems.map(it => +it.impositionMarginHorizontal || 0), 0);
    const gapVIn = Math.max(...orderItems.map(it => +it.impositionMarginVertical || 0), 0);

    const impositionWidth = +(orderItems[0]?.impositionWidth || 19);
    const impositionHeight = +(orderItems[0]?.impositionHeight || 13);

    const requestedOrientation = orderItems[0]?.impositionOrientation?.toLowerCase();
    let sheetWIn: number, sheetHIn: number, actualOrientation: string;
    if (requestedOrientation === 'portrait') {
      sheetWIn = Math.min(impositionWidth, impositionHeight);
      sheetHIn = Math.max(impositionWidth, impositionHeight);
      actualOrientation = 'portrait';
    } else if (requestedOrientation === 'landscape') {
      sheetWIn = Math.max(impositionWidth, impositionHeight);
      sheetHIn = Math.min(impositionWidth, impositionHeight);
      actualOrientation = 'landscape';
    } else {
      sheetWIn = impositionWidth; sheetHIn = impositionHeight;
      actualOrientation = sheetHIn > sheetWIn ? 'portrait' : 'landscape';
    }

    await job.log(LogLevel.Info,
      `Sheet ${sheetWIn}x${sheetHIn} (${actualOrientation}${requestedOrientation ? ' - explicit' : ' - inferred'}); ` +
      `Cut ${maxCutWIn}x${maxCutHIn}; Bleed ${maxBleedWIn}x${maxBleedHIn}; Gaps H=${gapHIn} V=${gapVIn}; Items=${orderItems.length}`
    );
    await job.log(LogLevel.Info, `Outer sheet margins set to 0.125".`);

    const layout = planLayout(sheetWIn, sheetHIn, orderItems);
    if (!layout) return job.fail(`Items cannot fit on ${sheetWIn}x${sheetHIn} (${actualOrientation}) with current gaps and 0.125" margins`);

    await job.log(LogLevel.Info, `Impose ${layout.cols}x${layout.rows} on ${layout.sheetWIn}x${layout.sheetHIn} (${layout.orientation}). Empty cells: ${layout.waste}`);

    const sheetWpt = pt(layout.sheetWIn), sheetHpt = pt(layout.sheetHIn);

    const baseUrl = 'http://10.1.0.79/api/switch/GetLocalArtwork/';

    // Load item assets
    const itemAssets = await Promise.all(orderItems.map(async (it) => {
      const url = `${baseUrl}${it.id}?pw=51ee6f3a3da5f642470202617cbcbd23`;
      let bytes: Uint8Array;
      try { bytes = await httpGetBytes(url); }
      catch (e:any) {
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

    // Embed all pages per item
    const perItemEmbeddedPages: Map<number, any[]> = new Map();
    async function ensureEmbeddedPagesForItem(itAsset: any) {
      const itId = itAsset.it.id as number;
      if (perItemEmbeddedPages.has(itId)) return;
      const idxs = Array.from({ length: itAsset.pageCount }, (_, i) => i);
      const embedded = await outDoc.embedPdf(itAsset.bytes, idxs);
      perItemEmbeddedPages.set(itId, embedded);
    }
    for (const asset of itemAssets) await ensureEmbeddedPagesForItem(asset);

    // placements
    const placements: any[] = [];
    for (let r = 0; r < layout.rows; r++) {
      for (let c = 0; c < layout.cols; c++) {
        const idx = r * layout.cols + c;
        if (idx < orderItems.length) placements.push({ r, c, itemIdx: idx });
      }
    }

    // Cover pages by includeCoverSheet + inksBack
    const includeCover = orderItems.some(it => !!it.includeCoverSheet);
    const anyBackInks = orderItems.some(it => (+it.inksBack || 0) !== 0);
    const numCoverPages = includeCover ? (anyBackInks ? 2 : 1) : 0;
    await job.log(LogLevel.Info, `Cover pages: ${numCoverPages} (${includeCover ? 'includeCoverSheet=true' : 'disabled'}; inksBack ${anyBackInks ? 'non-zero' : 'zero'})`);
    if (anyBackInks) {
      await job.log(LogLevel.Info, 'Back-side pages (odd indices) will mirror positions horizontally, invert rotations (add 180°), and invert X shift.');
    }

    if (numCoverPages >= 1)
      await createCoverPage(outDoc, layout, orderItems, itemAssets, perItemEmbeddedPages, placements, 0, font, boldFont, /*flipPositionsThisPage=*/false);
    if (numCoverPages === 2)
      await createCoverPage(outDoc, layout, orderItems, itemAssets, perItemEmbeddedPages, placements, 1, font, boldFont, /*flipPositionsThisPage=*/true);

    // Cache for per-item barcode images
    const barcodeCache: Map<string, any> = new Map();

    // Production pages
    for (let p = 0; p < maxPages; p++) {
      const page = outDoc.addPage([sheetWpt, sheetHpt]);
      const flipPositionsThisPage = anyBackInks && (p % 2 === 1);

      // Crops
      for (const plc of placements) {
        const it = orderItems[plc.itemIdx];
        const cutWpt_i = pt(+it.cutWidthInches || 0);
        const cutHpt_i = pt(+it.cutHeightInches || 0);

        const cEff = effectiveCol(plc.c, layout, flipPositionsThisPage);
        const rEff = plc.r;

        const cellCenterX = layout.offX + cEff * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
        const cellCenterY = layout.offY + rEff * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

        const isLeftEdge = cEff === 0, isRightEdge = cEff === layout.cols - 1;
        const isBottomEdge = rEff === 0, isTopEdge = rEff === layout.rows - 1;

        drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt_i, cutHpt_i, 0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, layout.gapHpt, layout.gapVpt);
      }

      // Artwork + Bugs
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

        const cEff = effectiveCol(plc.c, layout, flipPositionsThisPage);
        const rEff = plc.r;

        const cellCenterX = layout.offX + cEff * (layout.cellWpt + layout.gapHpt) + layout.cellWpt / 2;
        const cellCenterY = layout.offY + rEff * (layout.cellHpt + layout.gapVpt) + layout.cellHpt / 2;

        const rotDeg = rotationForPage(it, rEff, cEff, flipPositionsThisPage);

        // Invert X shift for odd pages (pre-rotation)
        const inputShiftX = flipPositionsThisPage ? -(+it.imageShiftX || 0) : (+it.imageShiftX || 0);
        const { sx, sy } = preRotationShiftFor(rotDeg, inputShiftX, +it.imageShiftY);

        const x0 = cellCenterX - placeW / 2 + sx;
        const y0 = cellCenterY - placeH / 2 + sy;
        const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rotDeg);

        page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(rotDeg) });

        await drawGutterBug(page, outDoc, it, layout, rEff, cEff, placeW, placeH, sheetWpt, sheetHpt, barcodeCache, font);

        // >>> ADD THIS: draw in-art barcode on LAST production sheet <<<
        const isLastProductionSheet = (p === maxPages - 1);
        await drawInArtBarcodeOnLastSheet(
          page,
          outDoc,
          it,
          isLastProductionSheet,
          rotDeg,
          x0, y0,
          cutWpt_i, cutHpt_i,
          bleedWpt_i, bleedHpt_i,
          barcodeCache,
          font
        );
      }

    }

    // Save & send
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const pdfBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, pdfBytes);

    const base = 'Batch-' + payload.batchId + '-' + payload.artworkUrlId + '.pdf';
    if ((job as any).sendToSingle) await (job as any).sendToSingle(base);
    else job.sendTo(rwPath, 0, base);
  } catch (e:any) {
    await job.fail(`Batching impose error: ${e.message || e}`);
  }
}
