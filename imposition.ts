/// <reference types="switch-scripting" />
// @ts-nocheck
import { PDFDocument, degrees, rgb, StandardFonts } from 'pdf-lib';
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

/** Draw a small slug line near a long edge (0.25" in) and center it along that edge. */
function drawSlugLine(page: any, text: string, sheetW: number, sheetH: number, font: any, marginIn: number = 0.25) {
  const m = pt(marginIn);
  let size = 7; // default; will downsize if needed
  const minSize = 4;
  const isLandscape = sheetW >= sheetH;
  const k = rgb(0,0,0);

  if (isLandscape) {
    // Horizontal along bottom edge
    let maxWidth = sheetW - 2 * m;
    let tw = font.widthOfTextAtSize(text, size);
    while (tw > maxWidth && size > minSize) { size -= 0.5; tw = font.widthOfTextAtSize(text, size); }
    const x = (sheetW - tw) / 2;
    const y = m; // 0.25" from bottom
    page.drawText(text, { x, y, size, font, color: k });
  } else {
    // Vertical along left edge
    let maxLen = sheetH - 2 * m;
    let tw = font.widthOfTextAtSize(text, size);
    while (tw > maxLen && size > minSize) { size -= 0.5; tw = font.widthOfTextAtSize(text, size); }
    const x = m; // 0.25" from left
    const y = (sheetH - tw) / 2; // center along vertical span
    page.drawText(text, { x, y, size, font, color: k, rotate: degrees(90) });
  }
}

/** Draw a teal overlay box with ID information */
function drawTealOverlay(
  page: any,
  centerX: number,
  centerY: number,
  cutW: number,
  cutH: number,
  orderId: string,
  lineId: string,
  font: any,
  boldFont: any
) {
  const halfW = cutW / 2;
  const halfH = cutH / 2;
  const xL = centerX - halfW;
  const yB = centerY - halfH;

  const teal = rgb(0, 0.5, 0.5);
  page.drawRectangle({
    x: xL,
    y: yB,
    width: cutW,
    height: cutH,
    color: teal,
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

  const itemText = `OrderItemID: ${lineId}`;
  const itemSize = 12;
  const itemWidth = font.widthOfTextAtSize(itemText, itemSize);
  const itemX = centerX - itemWidth / 2;
  const itemY = centerY - lineHeight / 2 - 4;
  page.drawText(itemText, { x: itemX, y: itemY, size: itemSize, font, color: white });
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

    // Slug info keys
    const orderId = await pd('orderId') || '';
    const lineId = await pd('lineId') || '';
    const pagesStr = await pd('pages') || '';

    // Bleed: if not explicitly provided, default to cut + 0.25"
    const bleedWRaw = +await pd('bleedWidthInches');
    const bleedHRaw = +await pd('bleedHeightInches');

    const bleedW = bleedWRaw > 0 ? bleedWRaw : (cutW + 0.25);
    const bleedH = bleedHRaw > 0 ? bleedHRaw : (cutH + 0.25);

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
      const neededWidth = cutW;
      const currentAvail = sheetW - 2*marginH;
      if (currentAvail < neededWidth) {
        marginReductionW = (neededWidth - currentAvail) / 2;
        availW = sheetW - 2*(marginH - marginReductionW);
      }
    }

    if (minRows < 1) {
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

    // --- open job, create imposed doc ---
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const src = await fs.readFile(rwPath);

    const outDoc = await PDFDocument.create();

    // Load original to count pages and to build embed list
    const srcDoc = await PDFDocument.load(src);
    const pageCount = srcDoc.getPageCount();
    if (!pageCount) return job.fail('Source PDF has 0 pages');

    const sheetSize: [number, number] = [sheetWpt, sheetHpt];
    const perSheet = cols * rows;

    // Prepare fonts and slug text once
    const font = await outDoc.embedFont(StandardFonts.Helvetica);
    const boldFont = await outDoc.embedFont(StandardFonts.HelveticaBold);
    const slugText = `OrderID: ${orderId} | ItemID: ${lineId} | No. Up: ${perSheet} | Pages: ${pagesStr} | CutHeight: ${cutH} | CutWidth: ${cutW}`;

    // Determine if artwork has bleed
    const hasBleed = (bleedWpt > cutWpt) || (bleedHpt > cutHpt);

    await job.log(
      LogLevel.Info,
      `Artwork: cut=${cutW}x${cutH}", bleed=${bleedW}x${bleedH}", hasBleed=${hasBleed}; cols=${cols}, rows=${rows}`
    );

    // --- BATCH EMBED ALL SOURCE PAGES ONCE (major size reduction) ---
    const pageIdxs = Array.from({ length: pageCount }, (_, i) => i);
    const embeddedPages = await outDoc.embedPdf(src, pageIdxs);

    // Size we draw per placement equals bleed (if present) otherwise cut
    const placeW = hasBleed ? bleedWpt : cutWpt;
    const placeH = hasBleed ? bleedHpt : cutHpt;

    // ---------- Build imposed pages ----------
    if (useRepeat) {
      // One sheet per source page, repeating the same page in every cell
      for (let srcPageIdx = 0; srcPageIdx < embeddedPages.length; srcPageIdx++) {
        const page = outDoc.addPage(sheetSize);
        drawSlugLine(page, slugText, sheetWpt, sheetHpt, font);
        const ep = embeddedPages[srcPageIdx];

        // Draw all crop marks
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const isLeftEdge = c === 0;
            const isRightEdge = c === cols - 1;
            const isBottomEdge = r === 0;
            const isTopEdge = r === rows - 1;
            drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt, cutHpt,
              0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapHpt, gapVpt);
          }
        }

        // Place page into every cell
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const x = cellCenterX - placeW / 2;
            const y = cellCenterY - placeH / 2;
            page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(0) });
          }
        }
      }
    } else if (impositionCutAndStack) {
      return job.fail('Cut and stack imposition not yet implemented');
    } else if (booklet) {
      return job.fail('Booklet imposition not yet implemented');
    } else {
      // Fill sheets sequentially with successive source pages
      let placed = 0;
      while (placed < embeddedPages.length || (placed % perSheet) !== 0) {
        const page = outDoc.addPage(sheetSize);
        drawSlugLine(page, slugText, sheetWpt, sheetHpt, font);

        // Draw crop marks grid
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const isLeftEdge = c === 0;
            const isRightEdge = c === cols - 1;
            const isBottomEdge = r === 0;
            const isTopEdge = r === rows - 1;
            drawIndividualCrops(page, cellCenterX, cellCenterY, cutWpt, cutHpt,
              0.0625, 0.125, 0.5, isLeftEdge, isRightEdge, isBottomEdge, isTopEdge, gapHpt, gapVpt);
          }
        }

        // Reset to start of this sheet
        placed = placed - (placed % perSheet);

        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const idx = placed % embeddedPages.length;
            const ep = embeddedPages[idx];

            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const x = cellCenterX - placeW / 2;
            const y = cellCenterY - placeH / 2;

            page.drawPage(ep, { x, y, width: placeW, height: placeH, rotate: degrees(0) });

            placed++;
            if (placed >= embeddedPages.length && (placed % perSheet) === 0) break;
          }
        }
      }
    }

    // ---------- Create a cover page by duplicating page 0 and overlaying info ----------
    if (outDoc.getPageCount() > 0) {
      const [dup] = await outDoc.copyPages(outDoc, [0]); // duplicate first imposed page
      outDoc.insertPage(0, dup);
      const cover = outDoc.getPage(0);

      // Optional: put the slug on the cover too (under overlays)
      drawSlugLine(cover, slugText, sheetWpt, sheetHpt, font);

      // Draw teal overlays on each position
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
          const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
          drawTealOverlay(cover, cellCenterX, cellCenterY, cutWpt, cutHpt, orderId, lineId, font, boldFont);
        }
      }
    }

    // Save compactly and route
    const pdfBytes = await outDoc.save({ useObjectStreams: true });
    await fs.writeFile(rwPath, pdfBytes);

    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e:any) {
    await job.fail(`Imposition error: ${e.message || e}`);
  }
}
