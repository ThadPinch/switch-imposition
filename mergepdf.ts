/// <reference types="switch-scripting" />
// @ts-nocheck

import { PDFDocument } from 'pdf-lib';
import * as fs from 'fs/promises';
import * as path from 'path';

function isPdfFile(fileName: string): boolean {
  return /\.pdf$/i.test(fileName || '');
}

function sortByRelativePath(files: string[], rootDir: string): string[] {
  return files.sort((a, b) => {
    const ra = path.relative(rootDir, a);
    const rb = path.relative(rootDir, b);
    return ra.localeCompare(rb, undefined, { numeric: true, sensitivity: 'base' });
  });
}

async function collectPdfFilesRecursive(rootDir: string): Promise<string[]> {
  const out: string[] = [];

  async function walk(dir: string): Promise<void> {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    entries.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));

    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
      } else if (entry.isFile() && isPdfFile(entry.name)) {
        out.push(fullPath);
      }
    }
  }

  await walk(rootDir);
  return sortByRelativePath(out, rootDir);
}

function mergedOutputName(jobName: string): string {
  const base = path.parse(String(jobName || '').trim()).name || 'merged';
  return `${base}.pdf`;
}

async function resolveOutputPath(job: Job, fallbackDir: string, outputName: string): Promise<string> {
  try {
    if (typeof job.createPathWithName === 'function') {
      const p = await job.createPathWithName(outputName, false);
      if (p) return String(p);
    }
  } catch {
    // Fall back to writing directly in fallbackDir.
  }
  return path.join(fallbackDir, outputName);
}

export async function jobArrived(_s: Switch, _f: FlowElement, job: Job) {
  try {
    const rwPath = await job.get(AccessLevel.ReadWrite);
    const st = await fs.stat(rwPath);

    let inputPdfs: string[] = [];
    let rootForSort = rwPath;
    let outputDir = rwPath;

    if (st.isDirectory()) {
      inputPdfs = await collectPdfFilesRecursive(rwPath);
      rootForSort = rwPath;
      outputDir = rwPath;
    } else if (st.isFile()) {
      if (!isPdfFile(path.basename(rwPath))) {
        return job.fail(`Input job is not a PDF: ${rwPath}`);
      }
      inputPdfs = [rwPath];
      rootForSort = path.dirname(rwPath);
      outputDir = path.dirname(rwPath);
    } else {
      return job.fail('Input job path is neither a file nor a folder');
    }

    inputPdfs = sortByRelativePath(inputPdfs, rootForSort);

    if (!inputPdfs.length) {
      return job.fail('No PDF files found in the incoming job');
    }

    await job.log(LogLevel.Info, `Merging ${inputPdfs.length} PDF file(s) into one output PDF.`);

    const merged = await PDFDocument.create();

    for (const pdfPath of inputPdfs) {
      const srcBytes = await fs.readFile(pdfPath);
      const srcDoc = await PDFDocument.load(srcBytes);
      const pageCount = srcDoc.getPageCount();
      if (!pageCount) continue;

      const pageIndexes = Array.from({ length: pageCount }, (_, i) => i);
      const pages = await merged.copyPages(srcDoc, pageIndexes);
      for (const p of pages) merged.addPage(p);
    }

    if (merged.getPageCount() === 0) {
      return job.fail('All input PDFs were empty (0 pages); no merged output created');
    }

    const outName = mergedOutputName(job.getName());
    const outPath = await resolveOutputPath(job, outputDir, outName);
    const outBytes = await merged.save();
    await fs.writeFile(outPath, outBytes);

    await job.log(LogLevel.Info, `Merged PDF created: ${path.basename(outPath)} (${merged.getPageCount()} pages).`);
    await job.sendToSingle(outPath);
  } catch (e: any) {
    await job.fail(`Merge PDF failed: ${e?.message || e}`);
  }
}
