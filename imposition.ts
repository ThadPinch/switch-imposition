/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb } from 'pdf-lib';
import * as fs from 'fs/promises';

/* ---------- helpers ---------- */
const PT = 72, pt = (inch: number) => inch * PT;

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gap: number) {
  // Account for gaps between cells
  const cols = Math.floor((availW + gap) / (cellW + gap));
  const rows = Math.floor((availH + gap) / (cellH + gap));
  return { cols, rows, up: cols * rows };
}

/** Draw crop marks for an individual placement */
function drawIndividualCrops(
  page: any,
  centerX: number,      // center of the placement cell (in points)
  centerY: number,
  cutW: number,         // cut width in points
  cutH: number,         // cut height in points
  gapIn: number = 0.0625, // 1/16" offset from trim
  lenIn: number = 0.125,  // 0.125" tick length for perimeter marks
  strokePt: number = 0.5,   // slightly thicker for better visibility
  isLeftEdge: boolean,
  isRightEdge: boolean,
  isBottomEdge: boolean,
  isTopEdge: boolean,
  gapBetweenCards: number
) {
  const off = pt(gapIn);
  const perimeterLen = pt(lenIn);
  // Interior marks should stop before entering the bleed area
  // They extend only into the gap between cards, not into the bleed
  const maxInteriorLen = (gapBetweenCards - off * 2) * 0.4; // 40% of available gap space
  const interiorLen = Math.min(pt(0.03125), maxInteriorLen); // Max 1/32" or less
  const k = rgb(0,0,0);

  // Calculate corners of the cut area
  const halfW = cutW / 2;
  const halfH = cutH / 2;
  const xL = centerX - halfW;
  const xR = centerX + halfW;
  const yB = centerY - halfH;
  const yT = centerY + halfH;

  // Top-left corner
  const topLen = isTopEdge ? perimeterLen : interiorLen;
  const leftLen = isLeftEdge ? perimeterLen : interiorLen;
  page.drawLine({ start:{x:xL, y:yT + off}, end:{x:xL, y:yT + off + topLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xL - off - leftLen, y:yT}, end:{x:xL - off, y:yT}, thickness: strokePt, color: k }); // horizontal

  // Top-right corner
  const rightLen = isRightEdge ? perimeterLen : interiorLen;
  page.drawLine({ start:{x:xR, y:yT + off}, end:{x:xR, y:yT + off + topLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xR + off, y:yT}, end:{x:xR + off + rightLen, y:yT}, thickness: strokePt, color: k }); // horizontal

  // Bottom-left corner
  const bottomLen = isBottomEdge ? perimeterLen : interiorLen;
  page.drawLine({ start:{x:xL, y:yB - off}, end:{x:xL, y:yB - off - bottomLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xL - off - leftLen, y:yB}, end:{x:xL - off, y:yB}, thickness: strokePt, color: k }); // horizontal

  // Bottom-right corner
  page.drawLine({ start:{x:xR, y:yB - off}, end:{x:xR, y:yB - off - bottomLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xR + off, y:yB}, end:{x:xR + off + rightLen, y:yB}, thickness: strokePt, color: k }); // horizontal
}

/* ---------- entry ---------- */
export async function jobArrived(_s: Switch, _f: FlowElement, job: Job) {
  try {
    const name = job.getName();
    if (!name.toLowerCase().endsWith('.pdf')) return job.fail('Not a PDF job');

    // --- parameters ---
    const pd = async (k: string) => (await job.getPrivateData(k)) as string;

    const cutW   = +await pd('cutWidthInches')  || 0;
    const cutH   = +await pd('cutHeightInches') || 0;
    const bleedW = +await pd('bleedWidthInches')  || (cutW + 0.25);
    const bleedH = +await pd('bleedHeightInches') || (cutH + 0.25);
    const margin = +await pd('impositionMargin') || 0;
    const sheetW = +await pd('impositionWidth') || 0;
    const sheetH = +await pd('impositionHeight') || 0;

    if (!sheetW || !sheetH || !cutW || !cutH) return job.fail('Missing/invalid size parameters');

    // --- calculate grid fit with original orientation ---
    const gap = margin;
    
    // Calculate available space with full margins
    let availW = sheetW - 2*margin;
    let availH = sheetH - 2*margin;
    
    // Check if we need to reduce margins due to content overflow
    const minCols = Math.floor((availW + gap) / (cutW + gap));
    const minRows = Math.floor((availH + gap) / (cutH + gap));
    
    // If even one item doesn't fit with full margins, calculate margin reduction
    let marginReductionW = 0;
    let marginReductionH = 0;
    
    if (minCols < 1) {
      // Calculate how much we need to reduce horizontal margins
      const neededWidth = cutW;
      const currentAvail = sheetW - 2*margin;
      if (currentAvail < neededWidth) {
        marginReductionW = (neededWidth - currentAvail) / 2;
        availW = sheetW - 2*(margin - marginReductionW);
      }
    }
    
    if (minRows < 1) {
      // Calculate how much we need to reduce vertical margins
      const neededHeight = cutH;
      const currentAvail = sheetH - 2*margin;
      if (currentAvail < neededHeight) {
        marginReductionH = (neededHeight - currentAvail) / 2;
        availH = sheetH - 2*(margin - marginReductionH);
      }
    }
    
    // Recalculate grid fit with adjusted available space
    const cols = Math.floor((availW + gap) / (cutW + gap));
    const rows = Math.floor((availH + gap) / (cutH + gap));
    
    if (!cols || !rows) return job.fail('Piece too large for sheet even with reduced margins');

    // --- convert to points ---
    const cutWpt = pt(cutW);
    const cutHpt = pt(cutH);
    const bleedWpt = pt(bleedW);
    const bleedHpt = pt(bleedH);
    const gapPt = pt(gap);
    
    // Apply margin reductions
    const effectiveMarginLeft = pt(margin - marginReductionW);
    const effectiveMarginRight = pt(margin - marginReductionW);
    const effectiveMarginTop = pt(margin - marginReductionH);
    const effectiveMarginBottom = pt(margin - marginReductionH);
    
    const sheetWpt = pt(sheetW);
    const sheetHpt = pt(sheetH);
    
    // Array dimensions using cut size + gap
    const arrWpt = cols * cutWpt + (cols - 1) * gapPt;
    const arrHpt = rows * cutHpt + (rows - 1) * gapPt;
    
    // Starting position (bottom-left of array) with adjusted margins
    const offX = effectiveMarginLeft + (sheetWpt - effectiveMarginLeft - effectiveMarginRight - arrWpt) / 2;
    const offY = effectiveMarginBottom + (sheetHpt - effectiveMarginBottom - effectiveMarginTop - arrHpt) / 2;

    // --- open job, create imposed doc ---
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const src = await fs.readFile(rwPath);
    const outDoc = await PDFDocument.create();
    const srcPages = await outDoc.embedPdf(src);
    if (!srcPages.length) return job.fail('Source PDF has 0 pages');

    const sheetSize: [number, number] = [sheetWpt, sheetHpt];
    let placed = 0, perSheet = cols * rows;

    while (placed < srcPages.length || placed % perSheet !== 0) {
      const page = outDoc.addPage(sheetSize);

      // First, draw all crop marks
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          // Calculate cell center based on cut size + gap
          const cellCenterX = offX + c * (cutWpt + gapPt) + cutWpt / 2;
          const cellCenterY = offY + r * (cutHpt + gapPt) + cutHpt / 2;
          
          // Determine if this is an edge card (for different crop lengths)
          const isLeftEdge = c === 0;
          const isRightEdge = c === cols - 1;
          const isBottomEdge = r === 0;
          const isTopEdge = r === rows - 1;
          
          drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt, cutHpt, 
            0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapPt);
        }
      }

      // Then, place pages on top of the crop marks
      placed = placed - (placed % perSheet); // Reset to start of this sheet
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const ep = srcPages[placed % srcPages.length];
          
          // Calculate cell center based on cut size + gap
          const cellCenterX = offX + c * (cutWpt + gapPt) + cutWpt / 2;
          const cellCenterY = offY + r * (cutHpt + gapPt) + cutHpt / 2;
          
          // Center the artwork within the bleed box
          const artX = cellCenterX - bleedWpt / 2;
          const artY = cellCenterY - bleedHpt / 2;
          
          // Further center if artwork is smaller than bleed
          const dx = (bleedWpt - ep.width) / 2;
          const dy = (bleedHpt - ep.height) / 2;

          // Place the page
          page.drawPage(ep, { x: artX + dx, y: artY + dy, rotate: degrees(0) });
          
          placed++;
          if (placed >= srcPages.length && placed % perSheet === 0) break;
        }
      }
    }

    // write back & route
    await fs.writeFile(rwPath, await outDoc.save());
    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e:any) {
    await job.fail(`Imposition error: ${e.message || e}`);
  }
}
