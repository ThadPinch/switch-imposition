/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb, StandardFonts } from 'pdf-lib';
import * as fs from 'fs/promises';

/* ---------- optional barcode ---------- */
let bwipjs: any = null;
try {
  // eslint-disable-next-line @typescript-eslint/no-var-requires
  bwipjs = require('bwip-js');
} catch { /* ok: text-only bug fallback */ }

/* ---------- helpers ---------- */
const PT = 72, pt = (inch: number) => inch * PT;

type BindingEdge = 'Left' | 'Right' | 'Top' | 'Bottom';

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  const cols = Math.floor((availW + gapH) / (cellW + gapH));
  const rows = Math.floor((availH + gapV) / (cellH + gapV));
  return { cols, rows, up: cols * rows };
}

/** math utils (for rotation-stable in-art placement) */
function normDeg(d: number) { return ((d % 360) + 360) % 360; }
function deg2rad(d: number) { return normDeg(d) * Math.PI / 180; }

/** Crops (per placement) */
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

/** Teal overlay (cover) */
function drawTealOverlay(page:any, cx:number, cy:number, cutW:number, cutH:number, orderId:string, lineId:string, font:any, bold:any){
  const x = cx - cutW/2, y = cy - cutH/2;
  page.drawRectangle({ x, y, width:cutW, height:cutH, color: rgb(0,0.5,0.5), opacity:0.9 });
  const white = rgb(1,1,1), lh = 14, idS = 12;
  const t1 = `OrderID: ${orderId}`, w1 = bold.widthOfTextAtSize(t1, idS);
  page.drawText(t1, { x:cx - w1/2, y: cy + lh/2, size:idS, font:bold, color:white });
  const t2 = `OrderItemID: ${lineId}`, w2 = font.widthOfTextAtSize(t2, idS);
  page.drawText(t2, { x:cx - w2/2, y: cy - lh/2 - 4, size:idS, font, color:white });
}

/** NEW: 0°/180° art rotation rule (adds EvenPages) */
function computeArtRotationDegrees(
  mode: string, startRot: boolean, r: number, c: number, artPageIndex: number
): number {
  const m = String(mode ?? 'None').trim().toLowerCase();
  if (m === 'evenpages') return (artPageIndex % 2 === 1) ? 180 : 0; // 0-based: odd index = even-numbered page
  if (m === 'rows')      return ((r % 2 === 0) ? !!startRot : !startRot) ? 180 : 0;
  if (m === 'columns' || m === 'cols' || m === 'column')
                         return ((c % 2 === 0) ? !!startRot : !startRot) ? 180 : 0;
  return 0;
}

/** keep centered for 180° */
function adjustXYForRotation(x:number, y:number, w:number, h:number, deg:number){
  const n = ((deg%360)+360)%360;
  return n===180 ? { x:x+w, y:y+h } : { x, y };
}

/** pre-rotation shift so +X/+Y is always right/up on sheet */
function preRotationShiftFor(deg:number, sxIn:number, syIn:number){
  const sx = pt(sxIn||0), sy = pt(syIn||0);
  const n = ((deg%360)+360)%360;
  if (n===0)   return { sx, sy };
  if (n===180) return { sx:-sx, sy:-sy };
  if (n===90)  return { sx:-sy, sy: sx };
  if (n===270) return { sx: sy, sy:-sx };
  return { sx, sy };
}

/** NEW: effective position (binding edge) */
function getEffectivePosition(
  r:number, c:number, cols:number, rows:number, bindingEdge:BindingEdge, flip:boolean
): { rEff:number, cEff:number } {
  if (!flip) return { rEff:r, cEff:c };
  const e = String(bindingEdge||'Left').toLowerCase();
  if (e==='left' || e==='right') return { rEff:r, cEff: cols-1-c };   // mirror columns
  // top/bottom
  return { rEff: rows-1-r, cEff:c };                                   // mirror rows
}

/** NEW: back-page shift inversion (binding edge) */
function getShiftAdjustments(
  bindingEdge:BindingEdge, flip:boolean, sx:number, sy:number
): { shiftX:number, shiftY:number } {
  if (!flip) return { shiftX: sx||0, shiftY: sy||0 };
  const e = String(bindingEdge||'Left').toLowerCase();
  if (e==='left' || e==='right') return { shiftX: -(sx||0), shiftY: sy||0 };
  return { shiftX: sx||0, shiftY: -(sy||0) };
}

/** NEW: overall rotation for a cell on this sheet */
function rotationForPage(
  artRotationMode:string, rotateFirstCR:boolean,
  rEff:number, cEff:number,
  flip:boolean,
  artPageIndex:number
): number {
  let deg = computeArtRotationDegrees(artRotationMode, rotateFirstCR, rEff, cEff, artPageIndex);
  if (flip) deg = (deg + 180) % 360;   // always add 180° on backs
  return deg;
}

/* ---------- Gutter Bug ---------- */
type BugSide = 'left'|'right'|'top'|'bottom';
const BUG_THICKNESS_IN = 0.125;   // strip thickness
const BUG_STANDOFF_IN  = 0.125;   // NEW: desired gap between art edge and bug strip

function availableGapForSide(
  side:BugSide,
  cols:number, rows:number, gapHpt:number, gapVpt:number,
  r:number, c:number,
  xL:number, xR:number, yB:number, yT:number,
  sheetWpt:number, sheetHpt:number
){
  if (side==='left')   return c>0 ? gapHpt : xL;
  if (side==='right')  return c<cols-1 ? gapHpt : (sheetWpt - xR);
  if (side==='bottom') return r>0 ? gapVpt : yB;
  if (side==='top')    return r<rows-1 ? gapVpt : (sheetHpt - yT);
  return 0;
}

function pickBugSide(
  requested: string,
  cols:number, rows:number, gapHpt:number, gapVpt:number,
  r:number, c:number,
  xL:number, xR:number, yB:number, yT:number,
  sheetWpt:number, sheetHpt:number
): BugSide | null {
  const thickPt = pt(BUG_THICKNESS_IN);
  const wantStandOffPt = pt(BUG_STANDOFF_IN);
  const needFull = thickPt + wantStandOffPt;
  const needMin  = thickPt;

  const req = String(requested||'Inside').toLowerCase() as any;

  const gaps:Record<BugSide,number> = {
    left:  availableGapForSide('left',  cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    right: availableGapForSide('right', cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    top:   availableGapForSide('top',   cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    bottom:availableGapForSide('bottom',cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
  };

  const dist:Record<BugSide,number> = { left:xL, right:sheetWpt-xR, bottom:yB, top:sheetHpt-yT };

  function pickAxis(axis:'h'|'v', inside:boolean, needPt:number): BugSide | null {
    const sides = axis==='h' ? (['left','right'] as BugSide[]) : (['bottom','top'] as BugSide[]);
    const viable = sides.filter(s => gaps[s] >= needPt);
    if (!viable.length) return null;
    viable.sort((a,b)=> inside ? (dist[b]-dist[a]) : (dist[a]-dist[b]));
    return viable[0]??null;
  }

  if (req==='left'||req==='right'||req==='top'||req==='bottom') {
    return gaps[req] >= needMin ? req : null;
  }

  const horizMax = Math.max(gaps.left, gaps.right);
  const vertMax  = Math.max(gaps.top, gaps.bottom);
  const horizOK = horizMax >= needMin, vertOK = vertMax >= needMin;
  const inside = req==='inside';

  if (horizOK && vertOK)
    return pickAxis(horizMax>=vertMax?'h':'v', inside, needFull)
        ?? pickAxis(horizMax>=vertMax?'v':'h', inside, needFull)
        ?? pickAxis(horizMax>=vertMax?'h':'v', inside, needMin)
        ?? pickAxis(horizMax>=vertMax?'v':'h', inside, needMin);
  if (horizOK) return pickAxis('h', inside, needFull) ?? pickAxis('h', inside, needMin);
  if (vertOK)  return pickAxis('v', inside, needFull) ?? pickAxis('v', inside, needMin);
  return null;
}

async function makeBarcodePngBytes(text:string): Promise<Uint8Array|null>{
  if (!bwipjs) return null;
  try{
    const buf:Buffer = await bwipjs.toBuffer({
      bcid:'code128', text,
      scale:2, height:8,
      includetext:false, textxalign:'center',
      backgroundcolor:'FFFFFF'
    });
    return new Uint8Array(buf);
  }catch{ return null; }
}

/** text helper (fits strip; horizontal or vertical) */
function drawBugText(page:any, font:any, text:string, bx:number, by:number, bw:number, bh:number, vertical:boolean, leftJustX:number, baselineY:number){
  if (!text) return;
  const k = rgb(0,0,0);
  const TARGET_PT = 10;                // desired 10 pt
  const MIN_PT = 6;
  const THICKNESS_OVERSHOOT_PT = 2;    // allow a slight overshoot past strip thickness

  const maxThickness = vertical ? bw : bh;
  const maxAlong = vertical ? bh : bw;

  let size = Math.min(TARGET_PT, maxThickness + THICKNESS_OVERSHOOT_PT);
  size = Math.max(size, MIN_PT);

  const maxUsableAlong = Math.max(2, maxAlong - pt(0.02));
  while (font.widthOfTextAtSize(text, size) > maxUsableAlong && size > MIN_PT) size -= 0.5;
  if (size < MIN_PT) size = MIN_PT;

  const tw = font.widthOfTextAtSize(text, size);

  if (!vertical){
    const x = leftJustX + pt(0.04);
    const y = baselineY - size/2;
    const maxX = bx + bw - pt(0.02);
    page.drawText(text, { x: Math.min(x, maxX - tw), y, size, font, color:k });
  }else{
    const x = bx + (bw - size)/2;
    const y = baselineY + pt(0.04);
    page.drawText(text, { x, y, size, font, color:k, rotate:degrees(90) });
  }
}

/**
 * Draw gutter bug + barcode aligned to BLEED edge and centered along that edge,
 * now OFFSET from the art edge by BUG_STANDOFF_IN (default 1/8")
 */
async function drawGutterBug(
  page:any, outDoc:PDFDocument,
  sheetWpt:number, sheetHpt:number,
  cols:number, rows:number, gapHpt:number, gapVpt:number,
  r:number, c:number,
  centerX:number, centerY:number,
  placeW:number, placeH:number,
  options: { include:boolean, includeBarcode:boolean, bugPosition:string, localArtworkPath:string, orderId:string, lineId:string },
  barcodeCache: Map<string, any>,
  font:any
){
  if (!options.include) return;

  // BLEED (or cut) edges of the artwork footprint:
  const xL = centerX - placeW/2, xR = centerX + placeW/2;
  const yB = centerY - placeH/2, yT = centerY + placeH/2;

  const side = pickBugSide(options.bugPosition, cols, rows, gapHpt, gapVpt, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt);
  if (!side) return;

  const thick = pt(BUG_THICKNESS_IN);
  const wantStandOffPt = pt(BUG_STANDOFF_IN);
  // Actual available gap on the chosen side
  const availGap = availableGapForSide(side, cols, rows, gapHpt, gapVpt, r, c, xL, xR, yB, yT, sheetWpt, sheetHpt);
  // Use up to 1/8" standoff, capped so the strip still fits in the gutter
  const standOffPt = Math.min(wantStandOffPt, Math.max(0, availGap - thick));

  let bx=xL, by=yB, bw=placeW, bh=thick, vertical=false;

  if (side==='top')    { bx=xL; by=yT + standOffPt;           bw=placeW; bh=thick;   vertical=false; }
  if (side==='bottom') { bx=xL; by=yB - thick - standOffPt;   bw=placeW; bh=thick;   vertical=false; }
  if (side==='left')   { bx=xL - thick - standOffPt; by=yB;   bw=thick;  bh=placeH;  vertical=true;  }
  if (side==='right')  { bx=xR + standOffPt;        by=yB;    bw=thick;  bh=placeH;  vertical=true;  }

  // white strip
  page.drawRectangle({ x:bx, y:by, width:bw, height:bh, color:rgb(1,1,1), opacity:1 });

  const midX = bx + bw/2, midY = by + bh/2;
  const label = String(options.localArtworkPath||'');
  const wantBarcode = !!options.includeBarcode;

  // compute barcode image (cache)
  let barcodeImg:any=null;
  if (wantBarcode){
    const key = `${options.orderId}-${options.lineId}`;
    barcodeImg = barcodeCache.get(key);
    if (!barcodeImg){
      const bytes = await makeBarcodePngBytes(key);
      if (bytes){
        barcodeImg = await outDoc.embedPng(bytes);
        barcodeCache.set(key, barcodeImg);
      }
    }
  }

  if (!vertical){
    // horizontal: barcode LEFT half, text RIGHT half (left-justified from centerline)
    const halfW = bw/2 - pt(0.04);
    if (barcodeImg){
      const iw = barcodeImg.width, ih = barcodeImg.height;
      const maxW = Math.max(1, halfW), maxH = bh*0.85;
      const scale = Math.min(maxW/iw, maxH/ih);
      const w = iw*scale, h=ih*scale;
      const x = bx + (bw/2 - w)/2;
      const y = by + (bh - h)/2;
      page.drawImage(barcodeImg, { x, y, width:w, height:h });
    }
    drawBugText(page, font, label, bx, by, bw, bh, false, midX, midY);
  }else{
    // vertical: barcode BOTTOM half (rot 90), text TOP half (rot 90)
    const halfH = bh/2 - pt(0.04);
    if (barcodeImg){
      const iw = barcodeImg.width, ih = barcodeImg.height;
      const scaleByWidth = bw / ih;
      const finalH = iw * scaleByWidth;
      const maxH = halfH * 0.95;
      const scale = Math.min(scaleByWidth, maxH / finalH * scaleByWidth);
      const w = iw*scale, h = ih*scale;
      const finalWidth = h, finalHeight = w;
      const x = bx + (bw - finalWidth)/2;
      const y = by + (bh/2 - finalHeight)/2;
      page.drawImage(barcodeImg, { x, y, width:finalWidth, height:finalHeight, rotate:degrees(90) });
    }
    drawBugText(page, font, label, bx, by, bw, bh, true, 0, midY);
  }
}

/** NEW: COVER PAGE (binding-edge aware; no bugs) */
function createCoverPage(
  outDoc: PDFDocument, embeddedPages: any[], pageIndex: number,
  sheetSize:[number,number],
  cols:number, rows:number,
  offX:number, offY:number,
  cutWpt:number, cutHpt:number,
  placeW:number, placeH:number,
  gapHpt:number, gapVpt:number,
  orderId:string, lineId:string,
  font:any, boldFont:any,
  useRepeat:boolean,
  artRotationMode:string, rotateFirstCR:boolean,
  imageShiftXIn:number, imageShiftYIn:number,
  bindingEdge: BindingEdge,
  flipPositionsThisPage: boolean
){
  const page = outDoc.addPage(sheetSize);

  // crops
  for (let r=0;r<rows;r++){
    for (let c=0;c<cols;c++){
      const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipPositionsThisPage);
      const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;
      drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.25, 0.125, 0.5, cEff===0, cEff===cols-1, rEff===0, rEff===rows-1, gapHpt, gapVpt);
    }
  }

  // art
  if (pageIndex < embeddedPages.length){
    const ep0 = embeddedPages[pageIndex];

    const place = (r:number,c:number)=>{
      const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipPositionsThisPage);
      const rot = rotationForPage(artRotationMode, rotateFirstCR, rEff, cEff, flipPositionsThisPage, pageIndex);

      const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;

      const { shiftX, shiftY } = getShiftAdjustments(bindingEdge, flipPositionsThisPage, imageShiftXIn, imageShiftYIn);
      const { sx, sy } = preRotationShiftFor(rot, shiftX, shiftY);

      const x0 = cx - placeW/2 + sx, y0 = cy - placeH/2 + sy;
      const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rot);
      page.drawPage(ep0, { x, y, width:placeW, height:placeH, rotate:degrees(rot) });
    };

    if (useRepeat){
      for (let r=0;r<rows;r++) for (let c=0;c<cols;c++) place(r,c);
    } else {
      let placed = pageIndex;
      for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
        if (placed < embeddedPages.length){ place(r,c); placed++; }
      }
    }
  }

  // overlay (IDs)
  for (let r=0;r<rows;r++){
    for (let c=0;c<cols;c++){
      const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipPositionsThisPage);
      const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;
      drawTealOverlay(page, cx, cy, cutWpt, cutHpt, orderId, lineId, font, boldFont);
    }
  }
}

/* ---------- In-art AutoShip barcode (targeted page) ---------- */

// Default target size (clamped to remaining cut area from the given X/Y)
const IN_ART_BARCODE_DEFAULT_W_IN = 1.00;
const IN_ART_BARCODE_DEFAULT_H_IN = 0.25;

/**
 * Draw an in-art AutoShip barcode on a specific artwork page.
 * Caller passes `shouldDraw` for the current artwork page index.
 * Uses bwip-js if available; otherwise falls back to text.
 */
async function drawInArtBarcodeAtTargetPage(
  page: any,
  outDoc: PDFDocument,
  opts: { enabled:boolean, xIn:number, yIn:number, orderId:string, lineId:string },
  shouldDraw: boolean,
  rotDeg: number,
  x0: number, y0: number,             // pre-rotation placed bottom-left (bleed or cut)
  cutWpt: number, cutHpt: number,     // cut size (pts)
  bleedWpt: number, bleedHpt: number, // bleed size (pts)
  barcodeCache: Map<string, any>,
  font: any
) {
  if (!shouldDraw) return;
  if (!opts.enabled) return;

  // Cut-box lower-left inside the placed area
  const cutOffX = Math.max(0, (bleedWpt - cutWpt) / 2);
  const cutOffY = Math.max(0, (bleedHpt - cutHpt) / 2);

  const userX = Math.max(0, pt(opts.xIn || 0));
  const userY = Math.max(0, pt(opts.yIn || 0));

  // Target box (clamped so we never exceed the cut area in the 0° case)
  const capW = Math.max(2, cutWpt - userX);
  const capH = Math.max(2, cutHpt - userY);
  const targetW = Math.min(pt(IN_ART_BARCODE_DEFAULT_W_IN), capW);
  const targetH = Math.min(pt(IN_ART_BARCODE_DEFAULT_H_IN), capH);

  // Barcode image (horizontal); rotate with artwork by rotDeg
  const cacheKey = `${opts.orderId}-${opts.lineId}-INART-H`;
  let img = barcodeCache.get(cacheKey);
  if (!img) {
    const bytes = await makeBarcodePngBytes(`${opts.orderId}-${opts.lineId}`);
    if (bytes) { img = await outDoc.embedPng(bytes); barcodeCache.set(cacheKey, img); }
  }

  // Compute final (w,h); fall back to text if needed
  let w = targetW, h = targetH, useTextFallback = !img;
  if (img) {
    const iw = img.width, ih = img.height;
    const scale = Math.max(0.01, Math.min(targetW / iw, targetH / ih));
    w = iw * scale; h = ih * scale;
  }

  // Cut center
  const cutLLx = x0 + cutOffX;
  const cutLLy = y0 + cutOffY;
  const cutCx  = cutLLx + cutWpt / 2;
  const cutCy  = cutLLy + cutHpt / 2;

  // Intended center in unrotated art-space
  const centerX0 = cutLLx + userX + w / 2;
  const centerY0 = cutLLy + userY + h / 2;

  // Rotate around cut center by rotDeg
  const theta = deg2rad(rotDeg);
  const cosT = Math.cos(theta), sinT = Math.sin(theta);
  const vx = centerX0 - cutCx, vy = centerY0 - cutCy;
  const centerX = cutCx + (vx * cosT - vy * sinT);
  const centerY = cutCy + (vx * sinT + vy * cosT);

  // Convert center -> bottom-left for rotated rect
  const halfWx = (w / 2) * cosT - (h / 2) * sinT;
  const halfWy = (w / 2) * sinT + (h / 2) * cosT;
  const drawX = centerX - halfWx;
  const drawY = centerY - halfWy;

  if (!useTextFallback) {
    page.drawImage(img, { x: drawX, y: drawY, width: w, height: h, rotate: degrees(normDeg(rotDeg)) });
    return;
  }

  // Text fallback
  const text = `${opts.orderId}-${opts.lineId}`;
  let size = 8;
  const maxWForText = targetW;
  while (size > 5 && font.widthOfTextAtSize(text, size) > maxWForText) size -= 0.5;
  const tw = font.widthOfTextAtSize(text, size);
  const th = size;

  const centerX0_txt = cutLLx + userX + tw / 2;
  const centerY0_txt = cutLLy + userY + th / 2;
  const vx2 = centerX0_txt - cutCx, vy2 = centerY0_txt - cutCy;
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
    const name = job.getName();
    if (!name.toLowerCase().endsWith('.pdf')) return job.fail('Not a PDF job');

    const pd = async (k:string)=>(await job.getPrivateData(k)) as string;

    const cutW   = +await pd('cutWidthInches')  || 0;
    const cutH   = +await pd('cutHeightInches') || 0;

    // IDs
    const orderId = await pd('orderId') || '';
    const lineId  = await pd('lineId')  || '';

    // bleed (fallback to cut)
    const bleedW = (+await pd('bleedWidthInches'))  || cutW;
    const bleedH = (+await pd('bleedHeightInches')) || cutH;

    // gaps
    const gapH = +await pd('impositionMarginHorizontal') || 0;
    const gapV = +await pd('impositionMarginVertical')   || 0;

    const sheetW = +await pd('impositionWidth')  || 0;
    const sheetH = +await pd('impositionHeight') || 0;

    // imposition flags
    const booklet = (await pd('booklet')) === 'true' || (await pd('booklet')) === true;
    const impositionRepeat     = (await pd('impositionRepeat'))     === 'true' || (await pd('impositionRepeat'))     === true;
    const impositionCutAndStack= (await pd('impositionCutAndStack'))=== 'true' || (await pd('impositionCutAndStack'))=== true;

    // rotation + shift
    const artRotationMode = (await pd('artRotation')) || 'None';              // supports: None | Rows | Columns | EvenPages
    const rotateFirstCR   = ((await pd('rotateFirstColumnOrRow')) === 'true') || ((await pd('rotateFirstColumnOrRow')) === true);
    const imageShiftXIn   = +(await pd('imageShiftX')) || 0;
    const imageShiftYIn   = +(await pd('imageShiftY')) || 0;

    // binding edge
    const bindingEdge: BindingEdge = (await pd('bindingEdge') as BindingEdge) || 'Left';

    // cover & bug options
    const includeCoverSheet = ((await pd('includeCoverSheet')) === 'true') || ((await pd('includeCoverSheet')) === true);
    const includeGutterBug  = ((await pd('includeGutterBug'))  === 'true') || ((await pd('includeGutterBug'))  === true);
    const includeBarcode    = ((await pd('includeAutoShipBarcodeInBug')) === 'true') || ((await pd('includeAutoShipBarcodeInBug')) === true);
    const bugPosition       = (await pd('bugPosition')) || 'Inside';
    const localArtworkPath  = (await pd('localArtworkPath')) || '';

    // In-art barcode private data (page, enable, position)
    const includeInArtBarcode = ((await pd('includeAutoShipBarcodeInArtOnLastSheet')) === 'true') || ((await pd('includeAutoShipBarcodeInArtOnLastSheet')) === true);
    const inArtBarcodeXIn     = +(await pd('inArtBarcodeX')) || 0;
    const inArtBarcodeYIn     = +(await pd('inArtBarcodeY')) || 0;
    // which artwork page to place the in-art barcode on (1-based; 0 => last)
    const inArtBarcodePageRaw = +(await pd('inArtBarcodePage')) || 0;

    const inksBack = +((await pd('inksBack')) || 0);
    const numCoverPages = includeCoverSheet ? (inksBack===0 ? 1 : 2) : 0;

    const useRepeat = impositionRepeat || (!booklet && !impositionCutAndStack);
    if (!sheetW || !sheetH || !cutW || !cutH) return job.fail('Missing/invalid size parameters');

    // ---- plan (0.125" outer margins) ----
    const outerM = 0.125;
    const availW = sheetW - 2*outerM;
    const availH = sheetH - 2*outerM;
    const cols = Math.floor((availW + gapH) / (cutW + gapH));
    const rows = Math.floor((availH + gapV) / (cutH + gapV));
    if (!cols || !rows) return job.fail('Piece too large for sheet with current gaps and 0.125" outer margins');

    // points
    const cutWpt = pt(cutW),  cutHpt = pt(cutH);
    const bleedWpt = pt(bleedW), bleedHpt = pt(bleedH);
    const gapHpt = pt(gapH), gapVpt = pt(gapV);
    const sheetWpt = pt(sheetW), sheetHpt = pt(sheetH);
    const outerHpt = pt(outerM), outerVpt = pt(outerM);

    const arrWpt = cols*cutWpt + (cols-1)*gapHpt;
    const arrHpt = rows*cutHpt + (rows-1)*gapVpt;
    const offX = outerHpt + (sheetWpt - 2*outerHpt - arrWpt)/2;
    const offY = outerVpt + (sheetHpt - 2*outerVpt - arrHpt)/2;

    // open src
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const src = await fs.readFile(rwPath);

    const outDoc = await PDFDocument.create();
    const srcDoc = await PDFDocument.load(src);
    const pageCount = srcDoc.getPageCount();
    if (!pageCount) return job.fail('Source PDF has 0 pages');

    // resolve target artwork page (0-based)
    let targetArtIndex = (inArtBarcodePageRaw > 0) ? (inArtBarcodePageRaw - 1) : (pageCount - 1);
    if (targetArtIndex < 0 || targetArtIndex >= pageCount) {
      targetArtIndex = Math.max(0, Math.min(pageCount - 1, targetArtIndex));
    }

    const sheetSize:[number,number] = [sheetWpt, sheetHpt];
    const perSheet = cols*rows;
    const font = await outDoc.embedFont(StandardFonts.Helvetica);
    const bold = await outDoc.embedFont(StandardFonts.HelveticaBold);

    const hasBleed = (bleedWpt>cutWpt) || (bleedHpt>cutHpt);
    const placeW = hasBleed ? bleedWpt : cutWpt;
    const placeH = hasBleed ? bleedHpt : cutHpt;

    await job.log(
      LogLevel.Info,
      `Single impose: sheet=${sheetW}x${sheetH}", cut=${cutW}x${cutH}", bleed=${bleedW}x${bleedH}", cols=${cols}, rows=${rows}, gaps H=${gapH} V=${gapV}, coverPages=${numCoverPages}, bug=${includeGutterBug?'on':'off'} (${bugPosition}, barcode=${includeBarcode?'on':'off'}), shift=(${imageShiftXIn},${imageShiftYIn})in; bindingEdge=${bindingEdge}.`
    );
    if (inksBack !== 0) {
      const e = String(bindingEdge).toLowerCase();
      const axis = (e==='left'||e==='right') ? 'flip columns (mirror horizontally) & invert X-shift' : 'flip rows (mirror vertically) & invert Y-shift';
      await job.log(LogLevel.Info, `Back sheets: ${axis}; add 180° to rotations.`);
    }
    if (String(artRotationMode).toLowerCase()==='evenpages'){
      await job.log(LogLevel.Info, `artRotation=EvenPages: artwork pages 2,4,6,... will be rotated 180°.`);
    }

    // embed pages
    const idxs = Array.from({length:pageCount}, (_,i)=>i);
    const embedded = await outDoc.embedPdf(src, idxs);

    // track output sheet index (0-based) to detect odd sheets for duplex
    let outSheetIndex = 0;

    // ---- cover (no bugs; no in-art) ----
    if (numCoverPages>=1){
      createCoverPage(
        outDoc, embedded, 0, sheetSize, cols, rows, offX, offY,
        cutWpt, cutHpt, placeW, placeH, gapHpt, gapVpt,
        orderId, lineId, font, bold, useRepeat,
        artRotationMode, rotateFirstCR, imageShiftXIn, imageShiftYIn,
        bindingEdge, /*flipPositionsThisPage*/ false
      );
      outSheetIndex++;
    }
    if (numCoverPages===2){
      createCoverPage(
        outDoc, embedded, 1, sheetSize, cols, rows, offX, offY,
        cutWpt, cutHpt, placeW, placeH, gapHpt, gapVpt,
        orderId, lineId, font, bold, useRepeat,
        artRotationMode, rotateFirstCR, imageShiftXIn, imageShiftYIn,
        bindingEdge, /*flipPositionsThisPage*/ true
      );
      outSheetIndex++;
    }

    // ---- production ----
    const barcodeCache:Map<string,any> = new Map();

    const placeOne = async (
      page:any, ep:any, r:number, c:number,
      flipPositionsThisPage:boolean,
      artIdx:number                    // which artwork page is being placed
    )=>{
      // binding-edge aware effective position
      const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipPositionsThisPage);

      // rotation
      const rot = rotationForPage(artRotationMode, rotateFirstCR, rEff, cEff, flipPositionsThisPage, artIdx);

      // center at effective position
      const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;

      // shift (invert axis per binding edge on backs)
      const { shiftX, shiftY } = getShiftAdjustments(bindingEdge, flipPositionsThisPage, imageShiftXIn, imageShiftYIn);
      const { sx, sy } = preRotationShiftFor(rot, shiftX, shiftY);

      const x0 = cx - placeW/2 + sx, y0 = cy - placeH/2 + sy;
      const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rot);
      page.drawPage(ep, { x, y, width:placeW, height:placeH, rotate:degrees(rot) });

      // gutter bug (use effective r/c to evaluate edges & gaps)
      await drawGutterBug(
        page, outDoc, sheetWpt, sheetHpt,
        cols, rows, gapHpt, gapVpt,
        rEff, cEff, cx, cy, placeW, placeH,
        {
          include: includeGutterBug,
          includeBarcode,
          bugPosition,
          localArtworkPath,
          orderId, lineId
        },
        barcodeCache, font
      );

      // in-art barcode on the requested artwork page only
      const shouldDraw = (artIdx === targetArtIndex);
      await drawInArtBarcodeAtTargetPage(
        page, outDoc,
        { enabled: includeInArtBarcode, xIn: inArtBarcodeXIn, yIn: inArtBarcodeYIn, orderId, lineId },
        shouldDraw,
        rot,
        x0, y0,
        cutWpt, cutHpt,
        bleedWpt, bleedHpt,
        barcodeCache,
        font
      );
    };

    if (useRepeat){
      for (let p=0;p<embedded.length;p++){
        const page = outDoc.addPage(sheetSize);

        // crops (binding-edge aware)
        const flipThisPage = (inksBack !== 0) && (outSheetIndex % 2 === 1);
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipThisPage);
          const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
          const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;
          drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.25, 0.125, 0.5, cEff===0, cEff===cols-1, rEff===0, rEff===rows-1, gapHpt, gapVpt);
        }

        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++)
          await placeOne(page, embedded[p], r, c, flipThisPage, /*artIdx*/ p);

        outSheetIndex++;
      }
    } else if (impositionCutAndStack) {
      return job.fail('Cut and stack imposition not yet implemented');
    } else if (booklet) {
      return job.fail('Booklet imposition not yet implemented');
    } else {
      let placed = 0, per = cols*rows;
      while (placed < embedded.length || (placed % per) !== 0){
        const page = outDoc.addPage(sheetSize);

        // crops (binding-edge aware)
        const flipThisPage = (inksBack !== 0) && (outSheetIndex % 2 === 1);
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const { rEff, cEff } = getEffectivePosition(r, c, cols, rows, bindingEdge, flipThisPage);
          const cx = offX + cEff*(cutWpt+gapHpt) + cutWpt/2;
          const cy = offY + rEff*(cutHpt+gapVpt) + cutHpt/2;
          drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.25, 0.125, 0.5, cEff===0, cEff===cols-1, rEff===0, rEff===rows-1, gapHpt, gapVpt);
        }

        placed = placed - (placed % per);
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const idx = placed % embedded.length;
          await placeOne(page, embedded[idx], r, c, flipThisPage, /*artIdx*/ idx);
          placed++;
          if (placed>=embedded.length && (placed%per)===0) break;
        }

        outSheetIndex++;
      }
    }

    // save & send
    const bytes = await outDoc.save({ useObjectStreams:true });
    await fs.writeFile(rwPath, bytes);
    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e:any) {
    await job.fail(`Imposition error: ${e.message || e}`);
  }
}
