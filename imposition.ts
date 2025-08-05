/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb } from 'pdf-lib';
import * as fs from 'fs/promises';

/* ---------- helpers ---------- */
const PT = 72, pt = (inch: number) => inch * PT;

function gridFit(availW: number, availH: number, cellW: number, cellH: number, gapH: number, gapV: number) {
  // Account for gaps between cells
  const cols = Math.floor((availW + gapH) / (cellW + gapH));
  const rows = Math.floor((availH + gapV) / (cellH + gapV));
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
  gapHorizontal: number,
  gapVertical: number
) {
  const off = pt(gapIn);
  const perimeterLen = pt(lenIn);
  // Interior marks should stop before entering the bleed area
  // They extend only into the gap between cards, not into the bleed
  const maxInteriorLenH = (gapHorizontal - off * 2) * 0.4; // 40% of available horizontal gap space
  const maxInteriorLenV = (gapVertical - off * 2) * 0.4; // 40% of available vertical gap space
  const interiorLenH = Math.min(pt(0.03125), maxInteriorLenH); // Max 1/32" or less
  const interiorLenV = Math.min(pt(0.03125), maxInteriorLenV); // Max 1/32" or less
  const k = rgb(0,0,0);

  // Calculate corners of the cut area
  const halfW = cutW / 2;
  const halfH = cutH / 2;
  const xL = centerX - halfW;
  const xR = centerX + halfW;
  const yB = centerY - halfH;
  const yT = centerY + halfH;

  // Top-left corner
  const topLen = isTopEdge ? perimeterLen : interiorLenV;
  const leftLen = isLeftEdge ? perimeterLen : interiorLenH;
  page.drawLine({ start:{x:xL, y:yT + off}, end:{x:xL, y:yT + off + topLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xL - off - leftLen, y:yT}, end:{x:xL - off, y:yT}, thickness: strokePt, color: k }); // horizontal

  // Top-right corner
  const rightLen = isRightEdge ? perimeterLen : interiorLenH;
  page.drawLine({ start:{x:xR, y:yT + off}, end:{x:xR, y:yT + off + topLen}, thickness: strokePt, color: k }); // vertical
  page.drawLine({ start:{x:xR + off, y:yT}, end:{x:xR + off + rightLen, y:yT}, thickness: strokePt, color: k }); // horizontal

  // Bottom-left corner
  const bottomLen = isBottomEdge ? perimeterLen : interiorLenV;
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
    const marginH = +await pd('impositionMarginHorizontal') || 0;
    const marginV = +await pd('impositionMarginVertical') || 0;
    const sheetW = +await pd('impositionWidth') || 0;
    const sheetH = +await pd('impositionHeight') || 0;
    
    // Imposition type parameters
    const booklet = (await pd('booklet')) === 'true' || (await pd('booklet')) === true;
    const impositionRepeat = (await pd('impositionRepeat')) === 'true' || (await pd('impositionRepeat')) === true;
    const impositionCutAndStack = (await pd('impositionCutAndStack')) === 'true' || (await pd('impositionCutAndStack')) === true;
    
    // Default to repeat if none specified
    const useRepeat = impositionRepeat || (!booklet && !impositionCutAndStack);

    if (!sheetW || !sheetH || !cutW || !cutH) return job.fail('Missing/invalid size parameters');

    // --- calculate grid fit with original orientation ---
    const gapH = marginH;
    const gapV = marginV;
    
    // Calculate available space with full margins
    let availW = sheetW - 2*marginH;
    let availH = sheetH - 2*marginV;
    
    // Check if we need to reduce margins due to content overflow
    const minCols = Math.floor((availW + gapH) / (cutW + gapH));
    const minRows = Math.floor((availH + gapV) / (cutH + gapV));
    
    // If even one item doesn't fit with full margins, calculate margin reduction
    let marginReductionW = 0;
    let marginReductionH = 0;
    
    if (minCols < 1) {
      // Calculate how much we need to reduce horizontal margins
      const neededWidth = cutW;
      const currentAvail = sheetW - 2*marginH;
      if (currentAvail < neededWidth) {
        marginReductionW = (neededWidth - currentAvail) / 2;
        availW = sheetW - 2*(marginH - marginReductionW);
      }
    }
    
    if (minRows < 1) {
      // Calculate how much we need to reduce vertical margins
      const neededHeight = cutH;
      const currentAvail = sheetH - 2*marginV;
      if (currentAvail < neededHeight) {
        marginReductionH = (neededHeight - currentAvail) / 2;
        availH = sheetH - 2*(marginV - marginReductionH);
      }
    }
    
    // Recalculate grid fit with adjusted available space
    const cols = Math.floor((availW + gapH) / (cutW + gapH));
    const rows = Math.floor((availH + gapV) / (cutH + gapV));
    
    if (!cols || !rows) return job.fail('Piece too large for sheet even with reduced margins');

    // --- convert to points ---
    const cutWpt = pt(cutW);
    const cutHpt = pt(cutH);
    const bleedWpt = pt(bleedW);
    const bleedHpt = pt(bleedH);
    const gapHpt = pt(gapH);
    const gapVpt = pt(gapV);
    
    // Apply margin reductions
    const effectiveMarginLeft = pt(marginH - marginReductionW);
    const effectiveMarginRight = pt(marginH - marginReductionW);
    const effectiveMarginTop = pt(marginV - marginReductionH);
    const effectiveMarginBottom = pt(marginV - marginReductionH);
    
    const sheetWpt = pt(sheetW);
    const sheetHpt = pt(sheetH);
    
    // Array dimensions using cut size + gaps
    const arrWpt = cols * cutWpt + (cols - 1) * gapHpt;
    const arrHpt = rows * cutHpt + (rows - 1) * gapVpt;
    
    // Starting position (bottom-left of array) with adjusted margins
    const offX = effectiveMarginLeft + (sheetWpt - effectiveMarginLeft - effectiveMarginRight - arrWpt) / 2;
    const offY = effectiveMarginBottom + (sheetHpt - effectiveMarginBottom - effectiveMarginTop - arrHpt) / 2;

    // --- open job, create imposed doc (embed pages EXPLICITLY) ---
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const src = await fs.readFile(rwPath);

    const outDoc = await PDFDocument.create();

    // Load original to count pages; embed each page by index to avoid "first page only".
    const srcDoc = await PDFDocument.load(src);
    const pageCount = srcDoc.getPageCount();

    const srcPages: any[] = [];
    for (let i = 0; i < pageCount; i++) {
      const [embedded] = await outDoc.embedPdf(src, [i]);
      srcPages.push(embedded);
    }

    await job.log(
      LogLevel.Info,
      `Repeat mode: source has ${pageCount} page(s); embedding ${srcPages.length}.`
    );

    if (!srcPages.length) return job.fail('Source PDF has 0 pages');

    const sheetSize: [number, number] = [sheetWpt, sheetHpt];
    let perSheet = cols * rows;

    // Handle repeat imposition (one page per sheet)
    if (useRepeat) {
      // For repeat imposition, create one sheet per source page
      for (let srcPageIdx = 0; srcPageIdx < srcPages.length; srcPageIdx++) {
        const page = outDoc.addPage(sheetSize);
        const ep = srcPages[srcPageIdx];

        // First, draw all crop marks
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            // Calculate cell center based on cut size + gaps
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            
            // Determine if this is an edge card (for different crop lengths)
            const isLeftEdge = c === 0;
            const isRightEdge = c === cols - 1;
            const isBottomEdge = r === 0;
            const isTopEdge = r === rows - 1;
            
            drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt, cutHpt, 
              0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapHpt, gapVpt);
          }
        }

        // Then, place the same page multiple times
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            // Calculate cell center based on cut size + gaps
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            
            // Center the artwork within the bleed box
            const artX = cellCenterX - bleedWpt / 2;
            const artY = cellCenterY - bleedHpt / 2;
            
            // Further center if artwork is smaller than bleed
            const dx = (bleedWpt - ep.width) / 2;
            const dy = (bleedHpt - ep.height) / 2;

            // Place the page
            page.drawPage(ep, { x: artX + dx, y: artY + dy, rotate: degrees(0) });
          }
        }
      }
    } else if (impositionCutAndStack) {
      // TODO: Implement cut and stack imposition
      return job.fail('Cut and stack imposition not yet implemented');
    } else if (booklet) {
      // TODO: Implement booklet imposition
      return job.fail('Booklet imposition not yet implemented');
    } else {
      // Original logic for single page repeating
      let placed = 0;
      while (placed < srcPages.length || placed % perSheet !== 0) {
        const page = outDoc.addPage(sheetSize);

        // First, draw all crop marks
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            // Calculate cell center based on cut size + gaps
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            
            // Determine if this is an edge card (for different crop lengths)
            const isLeftEdge = c === 0;
            const isRightEdge = c === cols - 1;
            const isBottomEdge = r === 0;
            const isTopEdge = r === rows - 1;
            
            drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt, cutHpt, 
              0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapHpt, gapVpt);
          }
        }

        // Then, place pages on top of the crop marks
        placed = placed - (placed % perSheet); // Reset to start of this sheet
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const ep = srcPages[placed % srcPages.length];
            
            // Calculate cell center based on cut size + gaps
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            
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
    }

    // write back & route
    await fs.writeFile(rwPath, await outDoc.save());
    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e:any) {
    await job.fail(`Imposition error: ${e.message || e}`);
  }
}
