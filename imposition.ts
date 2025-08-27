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

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  const cols = Math.floor((availW + gapH) / (cellW + gapH));
  const rows = Math.floor((availH + gapV) / (cellH + gapV));
  return { cols, rows, up: cols * rows };
}

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

/** 0°/180° rotation choice */
function computeCellRotation(mode: string, startRot: boolean, r: number, c: number): number {
  const m = String(mode ?? 'None').trim().toLowerCase();
  if (m === 'rows')   return ((r % 2 === 0) ? !!startRot : !startRot) ? 180 : 0;
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

/* ---------- Gutter Bug ---------- */
type BugSide = 'left'|'right'|'top'|'bottom';
const BUG_THICKNESS_IN = 0.125;

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
  const need = pt(BUG_THICKNESS_IN);
  const req = String(requested||'Inside').toLowerCase() as any;

  const gaps:Record<BugSide,number> = {
    left:  availableGapForSide('left',  cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    right: availableGapForSide('right', cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    top:   availableGapForSide('top',   cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
    bottom:availableGapForSide('bottom',cols, rows, gapHpt, gapVpt, r,c,xL,xR,yB,yT,sheetWpt,sheetHpt),
  };

  const dist:Record<BugSide,number> = { left:xL, right:sheetWpt-xR, bottom:yB, top:sheetHpt-yT };

  function pickAxis(axis:'h'|'v', inside:boolean): BugSide | null {
    const sides = axis==='h' ? (['left','right'] as BugSide[]) : (['bottom','top'] as BugSide[]);
    const viable = sides.filter(s => gaps[s] >= need);
    if (!viable.length) return null;
    viable.sort((a,b)=> inside ? (dist[b]-dist[a]) : (dist[a]-dist[b]));
    return viable[0]??null;
  }

  if (req==='left'||req==='right'||req==='top'||req==='bottom') {
    return gaps[req] >= need ? req : null;
  }

  const horizMax = Math.max(gaps.left, gaps.right);
  const vertMax  = Math.max(gaps.top, gaps.bottom);
  const horizOK = horizMax >= need, vertOK = vertMax >= need;
  const inside = req==='inside';

  if (horizOK && vertOK)  return pickAxis(horizMax>=vertMax?'h':'v', inside) ?? pickAxis(horizMax>=vertMax?'v':'h', inside);
  if (horizOK)            return pickAxis('h', inside);
  if (vertOK)             return pickAxis('v', inside);
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
  let size = Math.min(6, (vertical?bw:bh) * 0.40);
  if (size < 3) size = 3;
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

/** Draw gutter bug + barcode aligned to BLEED edge and centered along that edge */
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
  let bx=xL, by=yB, bw=placeW, bh=thick, vertical=false;

  if (side==='top')    { bx=xL; by=yT;           bw=placeW; bh=thick;   vertical=false; }
  if (side==='bottom') { bx=xL; by=yB - thick;   bw=placeW; bh=thick;   vertical=false; }
  if (side==='left')   { bx=xL - thick; by=yB;   bw=thick;  bh=placeH;  vertical=true;  }
  if (side==='right')  { bx=xR;        by=yB;    bw=thick;  bh=placeH;  vertical=true;  }

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

/** COVER PAGE (no bugs) */
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
  imageShiftXIn:number, imageShiftYIn:number
){
  const page = outDoc.addPage(sheetSize);

  // crops
  for (let r=0;r<rows;r++){
    for (let c=0;c<cols;c++){
      const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
      drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.0625, 0.125, 0.5, c===0, c===cols-1, r===0, r===rows-1, gapHpt, gapVpt);
    }
  }

  // art
  if (pageIndex < embeddedPages.length){
    const ep0 = embeddedPages[pageIndex];
    const place = (r:number,c:number)=>{
      const rot = computeCellRotation(artRotationMode, rotateFirstCR, r, c);
      const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
      const { sx, sy } = preRotationShiftFor(rot, imageShiftXIn, imageShiftYIn);
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
      const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
      drawTealOverlay(page, cx, cy, cutWpt, cutHpt, orderId, lineId, font, boldFont);
    }
  }
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
    const artRotationMode = (await pd('artRotation')) || 'None';
    const rotateFirstCR   = ((await pd('rotateFirstColumnOrRow')) === 'true') || ((await pd('rotateFirstColumnOrRow')) === true);
    const imageShiftXIn   = +(await pd('imageShiftX')) || 0;
    const imageShiftYIn   = +(await pd('imageShiftY')) || 0;

    // cover & bug options
    const includeCoverSheet = ((await pd('includeCoverSheet')) === 'true') || ((await pd('includeCoverSheet')) === true);
    const includeGutterBug  = ((await pd('includeGutterBug'))  === 'true') || ((await pd('includeGutterBug'))  === true);
    const includeBarcode    = ((await pd('includeAutoShipBarcodeInBug')) === 'true') || ((await pd('includeAutoShipBarcodeInBug')) === true);
    const bugPosition       = (await pd('bugPosition')) || 'Inside';
    const localArtworkPath  = (await pd('localArtworkPath')) || '';

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

    const sheetSize:[number,number] = [sheetWpt, sheetHpt];
    const perSheet = cols*rows;
    const font = await outDoc.embedFont(StandardFonts.Helvetica);
    const bold = await outDoc.embedFont(StandardFonts.HelveticaBold);

    const hasBleed = (bleedWpt>cutWpt) || (bleedHpt>cutHpt);
    const placeW = hasBleed ? bleedWpt : cutWpt;
    const placeH = hasBleed ? bleedHpt : cutHpt;

    await job.log(LogLevel.Info,
      `Single impose (no slug): sheet=${sheetW}x${sheetH}", cut=${cutW}x${cutH}", bleed=${bleedW}x${bleedH}", cols=${cols}, rows=${rows}, gaps H=${gapH} V=${gapV}, coverPages=${numCoverPages}, bug=${includeGutterBug?'on':'off'} (${bugPosition}, barcode=${includeBarcode?'on':'off'}), shift=(${imageShiftXIn},${imageShiftYIn})in`
    );

    // embed pages
    const idxs = Array.from({length:pageCount}, (_,i)=>i);
    const embedded = await outDoc.embedPdf(src, idxs);

    // ---- cover (no bugs) ----
    if (numCoverPages>=1){
      createCoverPage(outDoc, embedded, 0, sheetSize, cols, rows, offX, offY, cutWpt, cutHpt, placeW, placeH, gapHpt, gapVpt, orderId, lineId, font, bold, useRepeat, artRotationMode, rotateFirstCR, imageShiftXIn, imageShiftYIn);
    }
    if (numCoverPages===2){
      createCoverPage(outDoc, embedded, 1, sheetSize, cols, rows, offX, offY, cutWpt, cutHpt, placeW, placeH, gapHpt, gapVpt, orderId, lineId, font, bold, useRepeat, artRotationMode, rotateFirstCR, imageShiftXIn, imageShiftYIn);
    }

    // ---- production ----
    const barcodeCache:Map<string,any> = new Map();

    const placeOne = async (page:any, ep:any, r:number, c:number)=>{
      const rot = computeCellRotation(artRotationMode, rotateFirstCR, r, c);
      const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
      const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
      const { sx, sy } = preRotationShiftFor(rot, imageShiftXIn, imageShiftYIn);
      const x0 = cx - placeW/2 + sx, y0 = cy - placeH/2 + sy;
      const { x, y } = adjustXYForRotation(x0, y0, placeW, placeH, rot);
      page.drawPage(ep, { x, y, width:placeW, height:placeH, rotate:degrees(rot) });

      await drawGutterBug(
        page, outDoc, sheetWpt, sheetHpt,
        cols, rows, gapHpt, gapVpt,
        r, c, cx, cy, placeW, placeH,
        {
          include: includeGutterBug,
          includeBarcode,
          bugPosition,
          localArtworkPath,
          orderId, lineId
        },
        barcodeCache, font
      );
    };

    if (useRepeat){
      for (let p=0;p<embedded.length;p++){
        const page = outDoc.addPage(sheetSize);

        // crops
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
          const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
          drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.0625, 0.125, 0.5, c===0, c===cols-1, r===0, r===rows-1, gapHpt, gapVpt);
        }

        const ep = embedded[p];
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++) await placeOne(page, ep, r, c);
      }
    } else if (impositionCutAndStack) {
      return job.fail('Cut and stack imposition not yet implemented');
    } else if (booklet) {
      return job.fail('Booklet imposition not yet implemented');
    } else {
      let placed = 0, per = cols*rows;
      while (placed < embedded.length || (placed % per) !== 0){
        const page = outDoc.addPage(sheetSize);

        // crops
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const cx = offX + c*(cutWpt+gapHpt) + cutWpt/2;
          const cy = offY + r*(cutHpt+gapVpt) + cutHpt/2;
          drawIndividualCrops(page, cx, cy, cutWpt, cutHpt, 0.0625, 0.125, 0.5, c===0, c===cols-1, r===0, r===rows-1, gapHpt, gapVpt);
        }

        placed = placed - (placed % per);
        for (let r=0;r<rows;r++) for (let c=0;c<cols;c++){
          const idx = placed % embedded.length;
          const ep = embedded[idx];
          await placeOne(page, ep, r, c);
          placed++;
          if (placed>=embedded.length && (placed%per)===0) break;
        }
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
