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
  const maxInteriorLenH = (gapHorizontal - off * 2) * 0.4;
  const maxInteriorLenV = (gapVertical - off * 2) * 0.4;
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

/** Draw a small slug line near a long edge (0.25\" in) and center it along that edge.
 *  Draw this BEFORE crops/art so it sits underneath artwork.
 */
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
  
  // Draw teal rectangle with 90% opacity (0.1 transparency)
  const teal = rgb(0, 0.5, 0.5); // Teal color
  page.drawRectangle({
    x: xL,
    y: yB,
    width: cutW,
    height: cutH,
    color: teal,
    opacity: 0.9
  });
  
  // Draw text in white
  const white = rgb(1, 1, 1);
  const lineHeight = 14;
  
  // ID text
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
  
  // OrderItemID text
  const itemText = `OrderItemID: ${lineId}`;
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

/**
 * Build a reusable embedded page that carries its own CropBox.
 * We wrap the source page inside a small one-page PDF whose MediaBox equals the bleed size
 * (or the cut size if no bleed). Then we set that wrapper page's CropBox to the cut rectangle
 * centered on the page. We embed this wrapper once per source page and reuse it for every placement.
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
  // Wrapper doc holds a single page with the desired boxes
  const wrapper = await PDFDocument.create();
  const [srcEp] = await wrapper.embedPdf(srcBytes, [pageIndex]);

  const pageW = hasBleed ? bleedWpt : cutWpt;
  const pageH = hasBleed ? bleedHpt : cutHpt;
  const placementPage = wrapper.addPage([pageW, pageH]);

  // Center original artwork on wrapper page (no scaling, matching the original code's behavior)
  const artX = (pageW - srcEp.width) / 2;
  const artY = (pageH - srcEp.height) / 2;
  placementPage.drawPage(srcEp, { x: artX, y: artY, rotate: degrees(0) });

  // Define CropBox: centered cut rectangle inside the bleed-sized page (or full page if no bleed)
  const cropX = hasBleed ? (bleedWpt - cutWpt) / 2 : 0;
  const cropY = hasBleed ? (bleedHpt - cutHpt) / 2 : 0;
  const cropW = cutWpt;
  const cropH = cutHpt;
  placementPage.setCropBox(cropX, cropY, cropW, cropH);

  // Embed wrapper's single page into the output doc and return that embedded page object
  const wrapperBytes = await wrapper.save();
  const [embeddedPlacement] = await outDoc.embedPdf(wrapperBytes, [0]);
  return embeddedPlacement;
}

/** Create a cover sheet page that duplicates the first imposed page and adds teal overlays */
async function createCoverSheet(
  outDoc: any,
  firstPageBytes: Uint8Array,
  sheetWpt: number,
  sheetHpt: number,
  cols: number,
  rows: number,
  offX: number,
  offY: number,
  cutWpt: number,
  cutHpt: number,
  gapHpt: number,
  gapVpt: number,
  orderId: string,
  lineId: string,
  font: any,
  boldFont: any,
  slugText: string
) {
  // Create a new page for the cover sheet
  const coverPage = outDoc.addPage([sheetWpt, sheetHpt]);
  
  // Draw slug line
  // drawSlugLine(coverPage, slugText, sheetWpt, sheetHpt, font);
  
  // Embed the first imposed page as background
  const [bgPage] = await outDoc.embedPdf(firstPageBytes, [0]);
  coverPage.drawPage(bgPage, { x: 0, y: 0 });
  
  // Draw teal overlays on each position
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
      const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
      
      drawTealOverlay(
        coverPage,
        cellCenterX,
        cellCenterY,
        cutWpt,
        cutHpt,
        orderId,
        lineId,
        font,
        boldFont
      );
    }
  }
  
  return coverPage;
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

    // --- open job, create imposed doc (embed pages EXPLICITLY) ---
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const src = await fs.readFile(rwPath);

    const outDoc = await PDFDocument.create();

    // Load original to count pages (we will NOT embed these directly). We'll build cropped wrappers instead.
    const srcDoc = await PDFDocument.load(src);
    const pageCount = srcDoc.getPageCount();

    if (!pageCount) return job.fail('Source PDF has 0 pages');

    const sheetSize: [number, number] = [sheetWpt, sheetHpt];
    let perSheet = cols * rows;

    // Prepare fonts and slug text once
    const font = await outDoc.embedFont(StandardFonts.Helvetica);
    const boldFont = await outDoc.embedFont(StandardFonts.HelveticaBold);
    const slugText = `OrderID: ${orderId} | ItemID: ${lineId} | No. Up: ${perSheet} | Pages: ${pagesStr} | CutHeight: ${cutH} | CutWidth: ${cutW}`;

    // Determine if artwork has bleed
    const hasBleed = (bleedWpt > cutWpt) || (bleedHpt > cutHpt);

    await job.log(
      LogLevel.Info,
      `Artwork: cut=${cutW}x${cutH}", bleed=${bleedW}x${bleedH}", hasBleed=${hasBleed}`
    );

    // Build ONE embedded placement page (with CropBox) per source page, then reuse it for all placements
    const placementEmbeds: any[] = [];
    for (let i = 0; i < pageCount; i++) {
      const embed = await buildPlacementXObject(outDoc, src, i, cutWpt, cutHpt, bleedWpt, bleedHpt, hasBleed);
      placementEmbeds.push(embed);
    }

    // Size we draw per placement equals the wrapper page size
    const placeW = hasBleed ? bleedWpt : cutWpt;
    const placeH = hasBleed ? bleedHpt : cutHpt;

    // Store the first imposed page for the cover sheet
    let firstImposedPageBytes: Uint8Array | null = null;
    let tempDoc: any = null;

    // Handle repeat imposition (one sheet per source page)
    if (useRepeat) {
      for (let srcPageIdx = 0; srcPageIdx < placementEmbeds.length; srcPageIdx++) {
        const page = outDoc.addPage(sheetSize);
        // Draw slug underneath everything
        drawSlugLine(page, slugText, sheetWpt, sheetHpt, font);
        const ep = placementEmbeds[srcPageIdx];

        // First, draw all crop marks
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

        // Place the same (cropped) page multiple times per sheet
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const x = cellCenterX - placeW / 2;
            const y = cellCenterY - placeH / 2;
            page.drawPage(ep, { x, y, rotate: degrees(0) });
          }
        }

        // Capture the first imposed page for the cover sheet
        if (srcPageIdx === 0 && !firstImposedPageBytes) {
          tempDoc = await PDFDocument.create();
          const [copiedPage] = await tempDoc.copyPages(outDoc, [0]);
          tempDoc.addPage(copiedPage);
          firstImposedPageBytes = await tempDoc.save();
        }
      }
    } else if (impositionCutAndStack) {
      return job.fail('Cut and stack imposition not yet implemented');
    } else if (booklet) {
      return job.fail('Booklet imposition not yet implemented');
    } else {
      // Original logic for multi-page repeating across sheets
      let placed = 0;
      let pageIndex = 0;
      while (placed < placementEmbeds.length || placed % perSheet !== 0) {
        const page = outDoc.addPage(sheetSize);
        // Draw slug underneath everything
        drawSlugLine(page, slugText, sheetWpt, sheetHpt, font);

        // First, draw all crop marks
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

        placed = placed - (placed % perSheet); // Reset to start of this sheet
        for (let r = 0; r < rows; r++) {
          for (let c = 0; c < cols; c++) {
            const ep = placementEmbeds[placed % placementEmbeds.length];

            const cellCenterX = offX + c * (cutWpt + gapHpt) + cutWpt / 2;
            const cellCenterY = offY + r * (cutHpt + gapVpt) + cutHpt / 2;
            const x = cellCenterX - placeW / 2;
            const y = cellCenterY - placeH / 2;
            page.drawPage(ep, { x, y, rotate: degrees(0) });

            placed++;
            if (placed >= placementEmbeds.length && placed % perSheet === 0) break;
          }
        }

        // Capture the first imposed page for the cover sheet
        if (pageIndex === 0 && !firstImposedPageBytes) {
          tempDoc = await PDFDocument.create();
          const [copiedPage] = await tempDoc.copyPages(outDoc, [pageIndex]);
          tempDoc.addPage(copiedPage);
          firstImposedPageBytes = await tempDoc.save();
        }
        pageIndex++;
      }
    }

    // Create and insert cover sheet at the beginning
    if (firstImposedPageBytes) {
      // Create a new document for the final output with cover sheet
      const finalDoc = await PDFDocument.create();
      
      // Create the cover sheet
      await createCoverSheet(
        finalDoc,
        firstImposedPageBytes,
        sheetWpt,
        sheetHpt,
        cols,
        rows,
        offX,
        offY,
        cutWpt,
        cutHpt,
        gapHpt,
        gapVpt,
        orderId,
        lineId,
        font,
        boldFont,
        slugText
      );
      
      // Copy all pages from the original output document
      const outDocBytes = await outDoc.save();
      const outDocForCopy = await PDFDocument.load(outDocBytes);
      const pagesToCopy = outDocForCopy.getPageCount();
      const copiedPages = await finalDoc.copyPages(outDocForCopy, Array.from({length: pagesToCopy}, (_, i) => i));
      copiedPages.forEach(page => finalDoc.addPage(page));
      
      // Save the final document with cover sheet
      await fs.writeFile(rwPath, await finalDoc.save());
    } else {
      // Fallback: save without cover sheet if something went wrong
      await fs.writeFile(rwPath, await outDoc.save());
    }

    // write back & route
    if ((job as any).sendToSingle) await (job as any).sendToSingle();
    else job.sendTo(rwPath, 0);
  } catch (e:any) {
    await job.fail(`Imposition error: ${e.message || e}`);
  }
}