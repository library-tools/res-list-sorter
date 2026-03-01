import { jsPDF } from "jspdf";

type Audience = "adult" | "junior";

type Entry = {
  rawLines: string[];
  barcode: string;
  shelfmark: string;
  author: string;
  itemType: string;
  sequence: string;
  audience: Audience;
  originalIndex: number;
};

type ParseResult = {
  libraryName: string;
  reportDate: string;
  entries: Entry[];
};

type SortResult = {
  originalCount: number;
  adultCount: number;
  juniorCount: number;
  isValid: boolean;
  adultDoc: RenderDocument;
  juniorDoc: RenderDocument;
};

type FontStyle = "normal" | "bold" | "italic";
type PdfFontFamily = "helvetica" | "times" | "courier";

type RenderSegment = {
  text: string;
  style: FontStyle;
};

type RenderLine = {
  segments: RenderSegment[];
  rightText?: string;
  rightStyle?: FontStyle;
  continuation?: boolean;
};

type RenderBlock = {
  lines: RenderLine[];
  sectionHeading: RenderLine;
  hasHeading: boolean;
};

type RenderDocument = {
  title: string;
  blocks: RenderBlock[];
};

const sourceText = mustGet<HTMLTextAreaElement>("sourceText");
const pdfFont = mustGet<HTMLSelectElement>("pdfFont");
const pdfTextSize = mustGet<HTMLSelectElement>("pdfTextSize");
const pdfColumns = mustGet<HTMLSelectElement>("pdfColumns");
const saveSettingsButton = mustGet<HTMLButtonElement>("saveSettingsButton");
const sortButton = mustGet<HTMLButtonElement>("sortButton");
const downloadActions = mustGet<HTMLElement>("downloadActions");
const downloadAdultButton = mustGet<HTMLButtonElement>("downloadAdultButton");
const downloadJuniorButton = mustGet<HTMLButtonElement>("downloadJuniorButton");
const statusPanel = mustGet<HTMLElement>("statusPanel");

const BARCODE_PATTERN = "30120\\d+";
const entryStartRegex = new RegExp(`^\\s*(.*?)\\s+(${BARCODE_PATTERN})\\s*$`);
const reportLineColumns = 45;
const SETTINGS_STORAGE_KEY = "resListSettings";
const MAX_INPUT_BYTES = 1_048_576;
const MAX_RENDER_LINES = 12_000;
const MAX_PDF_PAGES = 500;

let latestResult: SortResult | null = null;
let lastSortedSource: string | null = null;
let saveSettingsResetTimer: number | null = null;

loadSavedSettings();

saveSettingsButton.addEventListener("click", () => {
  saveCurrentSettings();
  showSavedSettingsFeedback();
});

sourceText.addEventListener("input", () => {
  if (lastSortedSource === null) {
    return;
  }

  if (sourceText.value !== lastSortedSource) {
    resetDownloads();
    latestResult = null;
    lastSortedSource = null;
    hideStatus();
  }
});

pdfTextSize.addEventListener("change", () => {
  if (!downloadActions.classList.contains("hidden")) {
    resetDownloads();
    latestResult = null;
    lastSortedSource = null;
    hideStatus();
  }
});

pdfFont.addEventListener("change", () => {
  if (!downloadActions.classList.contains("hidden")) {
    resetDownloads();
    latestResult = null;
    lastSortedSource = null;
    hideStatus();
  }
});

pdfColumns.addEventListener("change", () => {
  if (!downloadActions.classList.contains("hidden")) {
    resetDownloads();
    latestResult = null;
    lastSortedSource = null;
    hideStatus();
  }
});

sortButton.addEventListener("click", () => {
  resetDownloads();
  latestResult = null;
  lastSortedSource = null;

  const source = sourceText.value;
  if (!source.trim()) {
    showError("Please paste a reservation list before sorting.");
    return;
  }

  const inputBytes = new TextEncoder().encode(source).length;
  if (inputBytes > MAX_INPUT_BYTES) {
    showError(
      `Input is too large (${formatBytes(inputBytes)}). Maximum allowed is ${formatBytes(MAX_INPUT_BYTES)}.`,
    );
    return;
  }

  try {
    const parsed = parseList(source);
    const sorted = buildSortedLists(parsed);
    latestResult = sorted;

    const statusLine =
      `Original items: ${sorted.originalCount} | ` +
      `Adult items: ${sorted.adultCount} | ` +
      `Junior items: ${sorted.juniorCount}`;

    if (!sorted.isValid) {
      showError(statusLine + "\nIntegrity check: FAIL\n\nDo not use output until this is resolved.");
      return;
    }

    showOk(statusLine);
    lastSortedSource = sourceText.value;

    downloadActions.classList.remove("hidden");
  } catch (error) {
    console.error(error);
    showError(mapSortErrorToUserMessage(error));
  }
});

downloadAdultButton.addEventListener("click", () => {
  if (!latestResult || !latestResult.isValid) {
    showError("There was an error. Please reload the page and try again.");
    return;
  }

  downloadPdf("adult-list.pdf", latestResult.adultDoc, getPdfTextSize(), getPdfColumns(), getPdfFontFamily());
});

downloadJuniorButton.addEventListener("click", () => {
  if (!latestResult || !latestResult.isValid) {
    showError("There was an error. Please reload the page and try again.");
    return;
  }

  downloadPdf("junior-list.pdf", latestResult.juniorDoc, getPdfTextSize(), getPdfColumns(), getPdfFontFamily());
});

function parseList(input: string): ParseResult {
  const lines = input.replace(/\r/g, "").split("\n");
  const firstEntryIndex = lines.findIndex((line) => entryStartRegex.test(line));

  if (firstEntryIndex < 0) {
    throw new Error("No item entries found. Check that the pasted text includes full reservation entries.");
  }

  const headerLines = lines.slice(0, firstEntryIndex);
  const { libraryName, reportDate } = extractHeaderMetadata(headerLines);
  const entries: Entry[] = [];
  let currentLines: string[] = [];

  const finalizeCurrent = (): void => {
    if (currentLines.length === 0) {
      return;
    }

    const entry = parseEntry(currentLines, entries.length);
    entries.push(entry);
    currentLines = [];
  };

  for (let i = firstEntryIndex; i < lines.length; i += 1) {
    const line = lines[i];
    if (entryStartRegex.test(line)) {
      finalizeCurrent();
      currentLines = [line];
      continue;
    }

    if (currentLines.length > 0) {
      currentLines.push(line);
    }
  }

  finalizeCurrent();

  const rawEntryCount = lines.filter((line) => entryStartRegex.test(line)).length;
  if (entries.length !== rawEntryCount) {
    throw new Error(
      `Parsing integrity check failed: ${rawEntryCount} entries detected in source but only ${entries.length} were successfully parsed. Do not use the output.`,
    );
  }

  if (entries.length === 0) {
    throw new Error("Entries could not be parsed.");
  }

  return { libraryName, reportDate, entries };
}

function extractHeaderMetadata(headerLines: string[]): { libraryName: string; reportDate: string } {
  let libraryName = "Unknown Library";
  let reportDate = "Unknown date";

  for (const line of headerLines) {
    const libraryMatch = line.match(/^\s*Items at\s+(.+?)\s*$/i);
    if (libraryMatch) {
      libraryName = libraryMatch[1].trim();
    }

    const dateMatch = line.match(/^\s*res_itm_noloan\s+(\d{2}\/\d{2}\/\d{2})\s*$/i);
    if (dateMatch) {
      reportDate = dateMatch[1];
    }
  }

  return { libraryName, reportDate };
}

function parseEntry(rawLines: string[], originalIndex: number): Entry {
  const firstLine = rawLines[0] ?? "";
  const match = firstLine.match(entryStartRegex);

  if (!match) {
    throw new Error(`Invalid entry start at item ${originalIndex + 1}.`);
  }

  const preBarcode = match[1].trim();
  const barcode = match[2];
  const shelfmark = preBarcode.split(/\s+/)[0] ?? "";
  const author = (rawLines[1] ?? "").trim();
  const itemType = findField(rawLines, "Item Type");
  const sequence = findField(rawLines, "Sequence");
  const audience = classifyAudience(itemType, sequence);

  return {
    rawLines,
    barcode,
    shelfmark,
    author,
    itemType,
    sequence,
    audience,
    originalIndex,
  };
}

function findField(lines: string[], fieldName: string): string {
  const regex = new RegExp(`^\\s*${escapeRegex(fieldName)}\\s*:\\s*(.*)$`, "i");
  for (const line of lines) {
    const match = line.match(regex);
    if (match) {
      return (match[1] ?? "").trim();
    }
  }
  return "";
}

function classifyAudience(itemType: string, sequence: string): Audience {
  const type = itemType.trim();

  if (/^adult\b/i.test(type)) {
    return "adult";
  }

  if (/^junior\b/i.test(type)) {
    return "junior";
  }

  if (/children/i.test(type) || /children/i.test(sequence)) {
    return "junior";
  }

  // Fallback: treat as adult unless junior markers are present.
  return "adult";
}

function buildSortedLists(parsed: ParseResult): SortResult {
  const allEntries = parsed.entries;
  const adultEntries = allEntries.filter((entry) => entry.audience === "adult").sort(compareEntries);
  const juniorEntries = allEntries.filter((entry) => entry.audience === "junior").sort(compareEntries);

  const originalCount = allEntries.length;
  const adultCount = adultEntries.length;
  const juniorCount = juniorEntries.length;

  const allIndices = new Set(allEntries.map((e) => e.originalIndex));
  const splitIndices = new Set([...adultEntries, ...juniorEntries].map((e) => e.originalIndex));
  const isValid =
    splitIndices.size === allIndices.size &&
    [...allIndices].every((i) => splitIndices.has(i));

  const adultDoc = buildDocument("Adult", parsed.libraryName, parsed.reportDate, adultEntries);
  const juniorDoc = buildDocument("Junior", parsed.libraryName, parsed.reportDate, juniorEntries);

  return {
    originalCount,
    adultCount,
    juniorCount,
    isValid,
    adultDoc,
    juniorDoc,
  };
}

function compareEntries(a: Entry, b: Entry): number {
  const typeBucketCompare = compareTypeBucket(a.itemType, b.itemType);
  if (typeBucketCompare !== 0) {
    return typeBucketCompare;
  }

  const itemTypeCompare = compareText(a.itemType, b.itemType);
  if (itemTypeCompare !== 0) {
    return itemTypeCompare;
  }

  const sequenceCompare = compareSequence(
    normalizeSequenceForSorting(a.itemType, a.sequence),
    normalizeSequenceForSorting(b.itemType, b.sequence),
  );
  if (sequenceCompare !== 0) {
    return sequenceCompare;
  }

  const shelfmarkCompare = compareText(a.shelfmark, b.shelfmark);
  if (shelfmarkCompare !== 0) {
    return shelfmarkCompare;
  }

  const authorCompare = compareText(a.author, b.author);
  if (authorCompare !== 0) {
    return authorCompare;
  }

  const barcodeCompare = compareText(a.barcode, b.barcode);
  if (barcodeCompare !== 0) {
    return barcodeCompare;
  }

  return a.originalIndex - b.originalIndex;
}

function compareText(a: string, b: string): number {
  return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
}

function compareTypeBucket(aType: string, bType: string): number {
  return itemTypeBucket(aType) - itemTypeBucket(bType);
}

function itemTypeBucket(itemType: string): number {
  // Explicit rule: any DVD item type is treated as media/other.
  if (/\bdvd\b/i.test(itemType)) {
    return 3;
  }

  if (/^\s*(adult|junior)\s+non[- ]fiction\b/i.test(itemType)) {
    return 0;
  }

  if (/^\s*(adult|junior)\s+fiction\b/i.test(itemType)) {
    return 1;
  }

  if (/graphic\s+fiction/i.test(itemType)) {
    return 2;
  }

  return 3;
}

function compareSequence(a: string, b: string): number {
  const aBlank = a.trim() === "";
  const bBlank = b.trim() === "";

  if (aBlank && bBlank) {
    return 0;
  }

  if (aBlank) {
    return -1;
  }

  if (bBlank) {
    return 1;
  }

  return compareText(a, b);
}

function normalizeSequenceForSorting(itemType: string, sequence: string): string {
  if (!/^adult fiction$/i.test(itemType.trim())) {
    return sequence;
  }

  const normalized = sequence.trim().toLowerCase();

  // Local shelf policy: Thriller is filed with Crime.
  if (normalized === "thriller") {
    return "Crime";
  }

  // Local shelf policy: these are filed with general fiction.
  if (
    normalized === "historical" ||
    normalized === "romance" ||
    normalized === "saga" ||
    normalized === "horror" ||
    normalized === "western"
  ) {
    return "";
  }

  return sequence;
}

function buildDocument(
  audienceLabel: string,
  libraryName: string,
  reportDate: string,
  entries: Entry[],
): RenderDocument {
  const title = `${audienceLabel} reservation list — ${libraryName} — ${reportDate}`;
  const blocks = buildRenderBlocks(entries);
  return { title, blocks };
}

function buildRenderBlocks(entries: Entry[]): RenderBlock[] {
  const blocks: RenderBlock[] = [];
  let lastGroupKey: string | null = null;
  let currentSectionHeading: RenderLine | null = null;

  for (const entry of entries) {
    const effectiveSequence = normalizeSequenceForSorting(entry.itemType, entry.sequence);
    const groupKey = `${entry.itemType}\u0000${effectiveSequence}`;

    const cleanedEntryLines = buildEntryRenderLines(entry.rawLines);

    if (groupKey !== lastGroupKey) {
      const sectionHeading = makeStyledLine(formatGroupHeading(entry.itemType, effectiveSequence), "bold");
      currentSectionHeading = sectionHeading;
      const headingLines: RenderLine[] = [sectionHeading, makePlainLine("")];

      blocks.push({
        lines: [...headingLines, ...cleanedEntryLines, makePlainLine("")],
        sectionHeading,
        hasHeading: true,
      });

      lastGroupKey = groupKey;
      continue;
    }

    if (!currentSectionHeading) {
      currentSectionHeading = makeStyledLine(formatGroupHeading(entry.itemType, effectiveSequence), "bold");
    }

    blocks.push({
      lines: [...cleanedEntryLines, makePlainLine("")],
      sectionHeading: currentSectionHeading,
      hasHeading: false,
    });
  }

  if (blocks.length > 0) {
    const lastBlockLines = blocks[blocks.length - 1].lines;
    while (lastBlockLines.length > 0 && isBlankLine(lastBlockLines[lastBlockLines.length - 1])) {
      lastBlockLines.pop();
    }
  }

  return blocks;
}

function formatGroupHeading(itemType: string, effectiveSequence: string): string {
  if (effectiveSequence.trim() === "") {
    return `${itemType} —`;
  }

  return `${itemType} — ${effectiveSequence}`;
}

function buildEntryRenderLines(rawLines: string[]): RenderLine[] {
  const nonBlank = rawLines.filter((line) => line.trim() !== "");
  if (nonBlank.length === 0) {
    return [];
  }

  const first = nonBlank[0].replace(/^\s+/, "");
  const firstMatch = first.match(new RegExp(`^(.*?)\\s+(${BARCODE_PATTERN})\\s*$`));
  const lines: RenderLine[] = [];

  if (firstMatch) {
    lines.push({
      segments: [{ text: firstMatch[1].trimEnd(), style: "normal" }],
      rightText: firstMatch[2],
      rightStyle: "normal",
    });
  } else {
    lines.push(makePlainLine(first));
  }

  const remaining = nonBlank.slice(1).map((line) => line.replace(/^\s+/, ""));
  const metadataStart = remaining.findIndex((line) => isMetadataLine(line));
  const titleIndex = metadataStart > 0 ? metadataStart - 1 : -1;

  for (let i = 0; i < remaining.length; i += 1) {
    const line = remaining[i];

    if (shouldOmitEntryLine(line)) {
      continue;
    }

    if (i === titleIndex) {
      lines.push(makeStyledLine(line, "italic"));
      continue;
    }

    lines.push(makePlainLine(line));
  }

  return lines;
}

function downloadPdf(
  fileName: string,
  docModel: RenderDocument,
  textSizePt: number,
  columnCount: 1 | 2,
  fontFamily: PdfFontFamily,
): void {
  const doc = new jsPDF({ unit: "pt", format: "a4" });
  doc.setFont(fontFamily, "normal");
  doc.setFontSize(textSizePt);

  const margin = 36;
  const textHeight = doc.getTextDimensions("Mg").h;
  const lineHeight = Math.ceil(textHeight * 1.25);
  const fitSlack = Math.floor(lineHeight * 0.6);
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const maxTextWidth = pageWidth - margin * 2;
  const charWidth = doc.getTextWidth("0");
  const continuationIndent = charWidth * 2;
  const reportLineWidth = reportLineColumns * charWidth;
  const columnGap = columnCount === 2 ? 24 : 0;
  const columnWidth = columnCount === 2 ? (maxTextWidth - columnGap) / 2 : maxTextWidth;
  const rightAlignedWidth = Math.min(columnWidth, reportLineWidth);

  const footerY = pageHeight - 16;
  const usableHeight = footerY - margin;

  const wrappedTitle = wrapRenderLine(doc, makeStyledLine(docModel.title, "bold"), maxTextWidth, maxTextWidth, fontFamily);
  const wrappedBlocks = docModel.blocks.map((block) => ({
    lines: block.lines.flatMap((line) => wrapRenderLine(doc, line, columnWidth, rightAlignedWidth, fontFamily)),
    sectionHeading: block.sectionHeading,
    hasHeading: block.hasHeading,
  }));

  const totalWrappedLines = wrappedBlocks.reduce((sum, block) => sum + block.lines.length, wrappedTitle.length);
  if (totalWrappedLines > MAX_RENDER_LINES) {
    showError(
      `Output is too large to render safely (${totalWrappedLines.toLocaleString()} lines). ` +
        `Maximum allowed is ${MAX_RENDER_LINES.toLocaleString()} lines. Please split the list into smaller batches.`,
    );
    return;
  }

  const columnStartXs =
    columnCount === 2 ? [margin, margin + columnWidth + columnGap] : [margin];
  let renderAborted = false;

  const abortRender = (message: string): void => {
    showError(message);
    renderAborted = true;
  };

  const startNewPage = (): { y: number; columnIndex: number; contentStartY: number } => {
    const contentStartY = drawPageTitle(doc, wrappedTitle, pageWidth / 2, margin, lineHeight, fontFamily) + lineHeight;
    return { y: contentStartY, columnIndex: 0, contentStartY };
  };

  let { y, columnIndex, contentStartY } = startNewPage();

  const drawContinuedSectionHeading = (sectionHeading: RenderLine): void => {
    const leftX = columnStartXs[columnIndex];
    const barcodeRightX = leftX + rightAlignedWidth;
    const continuedHeading = makeStyledLine(`${lineText(sectionHeading)} (cont.)`, "bold");
    const headingLines = [
      ...wrapRenderLine(doc, continuedHeading, columnWidth, rightAlignedWidth, fontFamily),
      makePlainLine(""),
    ];
    drawWrappedLines(doc, headingLines, leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily);
    y += headingLines.length * lineHeight;
  };

  const advanceColumnOrPage = (): boolean => {
    if (columnIndex < columnStartXs.length - 1) {
      columnIndex += 1;
      y = contentStartY;
      return false;
    }

    if (doc.getNumberOfPages() >= MAX_PDF_PAGES) {
      abortRender(
        `Output is too large to render safely (${MAX_PDF_PAGES.toLocaleString()} pages limit). ` +
          "Please split the list into smaller batches.",
      );
      return true;
    }

    doc.addPage();
    const next = startNewPage();
    y = next.y;
    columnIndex = next.columnIndex;
    contentStartY = next.contentStartY;
    return true;
  };

  const drawOneLine = (line: RenderLine): void => {
    const leftX = columnStartXs[columnIndex];
    const barcodeRightX = leftX + rightAlignedWidth;
    drawWrappedLines(doc, [line], leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily);
    y += lineHeight;
  };

  for (const block of wrappedBlocks) {
    let linesToDraw = block.lines;
    const trailingBlankCount = countTrailingBlankLines(linesToDraw);
    if (trailingBlankCount > 0) {
      const trimmedLines = linesToDraw.slice(0, linesToDraw.length - trailingBlankCount);
      const fullHeight = linesToDraw.length * lineHeight;
      const trimmedHeight = trimmedLines.length * lineHeight;
      const remaining = usableHeight - y;

      if (fullHeight > remaining + fitSlack && trimmedHeight <= remaining + fitSlack) {
        linesToDraw = trimmedLines;
      }
    }

    const blockHeight = linesToDraw.length * lineHeight;
    const remaining = usableHeight - y;

    if (blockHeight > usableHeight) {
      for (const line of linesToDraw) {
        if (y + lineHeight > usableHeight) {
          const movedToNewPage = advanceColumnOrPage();
          if (renderAborted) {
            return;
          }

          if (movedToNewPage && !block.hasHeading) {
            drawContinuedSectionHeading(block.sectionHeading);
          }
        }

        drawOneLine(line);
      }
    } else {
      if (blockHeight > remaining + fitSlack) {
        const movedToNewPage = advanceColumnOrPage();
        if (renderAborted) {
          return;
        }

        if (movedToNewPage && !block.hasHeading) {
          drawContinuedSectionHeading(block.sectionHeading);
        }
      }

      const leftX = columnStartXs[columnIndex];
      const barcodeRightX = leftX + rightAlignedWidth;
      drawWrappedLines(doc, linesToDraw, leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily);
      y += blockHeight;
    }
  }

  if (renderAborted) {
    return;
  }

  const pageCount = doc.getNumberOfPages();
  for (let page = 1; page <= pageCount; page += 1) {
    doc.setPage(page);
    doc.setFont(fontFamily, "normal");
    doc.setFontSize(textSizePt);
    doc.text(`Page ${page} of ${pageCount}`, pageWidth / 2, footerY, { align: "center" });
  }

  doc.save(fileName);
}

function drawPageTitle(
  doc: jsPDF,
  wrappedTitle: RenderLine[],
  pageCenterX: number,
  margin: number,
  lineHeight: number,
  fontFamily: PdfFontFamily,
): number {
  let y = margin;
  for (const line of wrappedTitle) {
    drawCenteredLine(doc, line, pageCenterX, y, fontFamily);
    y += lineHeight;
  }

  return y;
}

function getPdfTextSize(): number {
  const parsed = Number.parseInt(pdfTextSize.value, 10);
  if (!Number.isFinite(parsed)) {
    return 11;
  }

  return Math.min(14, Math.max(8, parsed));
}

function getPdfFontFamily(): PdfFontFamily {
  const value = pdfFont.value;
  if (value === "times" || value === "courier") {
    return value;
  }

  return "helvetica";
}

function getPdfColumns(): 1 | 2 {
  return pdfColumns.value === "2" ? 2 : 1;
}

function wrapRenderLine(
  doc: jsPDF,
  line: RenderLine,
  maxTextWidth: number,
  rightAlignedWidth: number,
  fontFamily: PdfFontFamily,
): RenderLine[] {
  const fullText = lineText(line);
  if (fullText === "") {
    return [line];
  }

  const style = line.segments[0]?.style ?? "normal";

  if (line.rightText) {
    const rightStyle = line.rightStyle ?? "normal";
    setFontStyle(doc, rightStyle, fontFamily);
    const barcodeWidth = doc.getTextWidth(line.rightText);
    setFontStyle(doc, "normal", fontFamily);
    const gapWidth = doc.getTextWidth("  ");
    const leftMaxWidth = Math.max(20, rightAlignedWidth - barcodeWidth - gapWidth);
    const wrappedLeft = doc.splitTextToSize(fullText, leftMaxWidth) as string[];

    if (wrappedLeft.length <= 1) {
      return [{ segments: [{ text: wrappedLeft[0] ?? fullText, style }], rightText: line.rightText, rightStyle }];
    }

    const prefix = wrappedLeft.slice(0, -1).map((text, index) => ({
      segments: [{ text, style }],
      continuation: index > 0,
    }));
    const finalLine: RenderLine = {
      segments: [{ text: wrappedLeft[wrappedLeft.length - 1], style }],
      rightText: line.rightText,
      rightStyle,
      continuation: true,
    };

    return [...prefix, finalLine];
  }

  const wrapped = doc.splitTextToSize(fullText, maxTextWidth) as string[];
  if (wrapped.length === 0) {
    return [line];
  }

  return wrapped
    .filter((text) => text !== "")
    .map((text, index) => ({
      segments: [{ text, style }],
      continuation: index > 0,
    }));
}

function drawWrappedLines(
  doc: jsPDF,
  lines: RenderLine[],
  margin: number,
  lineHeight: number,
  startY: number,
  barcodeRightX: number,
  continuationIndent: number,
  fontFamily: PdfFontFamily,
): void {
  let y = startY;

  for (const line of lines) {
    let x = margin + (line.continuation ? continuationIndent : 0);
    for (const segment of line.segments) {
      if (segment.text === "") {
        continue;
      }

      setFontStyle(doc, segment.style, fontFamily);
      doc.text(segment.text, x, y);
      x += doc.getTextWidth(segment.text);
    }

    if (line.rightText) {
      setFontStyle(doc, line.rightStyle ?? "normal", fontFamily);
      doc.text(line.rightText, barcodeRightX, y, { align: "right" });
    }

    y += lineHeight;
  }
}

function drawCenteredLine(doc: jsPDF, line: RenderLine, centerX: number, y: number, fontFamily: PdfFontFamily): void {
  const text = lineText(line);
  if (text === "") {
    return;
  }

  const style = line.segments[0]?.style ?? "normal";
  setFontStyle(doc, style, fontFamily);
  doc.text(text, centerX, y, { align: "center" });
}

function isMetadataLine(line: string): boolean {
  return /^(Item Type|Sequence|Reserved at)\s*:/i.test(line);
}

function shouldOmitEntryLine(line: string): boolean {
  return /^Sequence\s*:\s*$/i.test(line);
}

function makePlainLine(text: string): RenderLine {
  return { segments: [{ text, style: "normal" }] };
}

function makeStyledLine(text: string, style: FontStyle): RenderLine {
  return { segments: [{ text, style }] };
}

function lineText(line: RenderLine): string {
  return line.segments.map((segment) => segment.text).join("");
}

function isBlankLine(line: RenderLine): boolean {
  return lineText(line) === "" && !line.rightText;
}

function countTrailingBlankLines(lines: RenderLine[]): number {
  let count = 0;
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    if (!isBlankLine(lines[i])) {
      break;
    }

    count += 1;
  }

  return count;
}

function setFontStyle(doc: jsPDF, style: FontStyle, fontFamily: PdfFontFamily): void {
  doc.setFont(fontFamily, style === "italic" ? "italic" : style === "bold" ? "bold" : "normal");
}

function loadSavedSettings(): void {
  try {
    const raw = window.localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) {
      return;
    }

    const parsed = JSON.parse(raw) as { font?: string; textSize?: string; columns?: string };

    if (parsed.font && hasOption(pdfFont, parsed.font)) {
      pdfFont.value = parsed.font;
    }

    if (parsed.textSize && hasOption(pdfTextSize, parsed.textSize)) {
      pdfTextSize.value = parsed.textSize;
    }

    if (parsed.columns && hasOption(pdfColumns, parsed.columns)) {
      pdfColumns.value = parsed.columns;
    }
  } catch {
    return;
  }
}

function saveCurrentSettings(): void {
  const settings = {
    font: getPdfFontFamily(),
    textSize: String(getPdfTextSize()),
    columns: String(getPdfColumns()),
  };

  try {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch {
    return;
  }
}

function showSavedSettingsFeedback(): void {
  if (saveSettingsResetTimer !== null) {
    window.clearTimeout(saveSettingsResetTimer);
    saveSettingsResetTimer = null;
  }

  saveSettingsButton.textContent = "Saved";
  saveSettingsButton.disabled = true;

  saveSettingsResetTimer = window.setTimeout(() => {
    saveSettingsButton.textContent = "Save settings";
    saveSettingsButton.disabled = false;
    saveSettingsResetTimer = null;
  }, 1000);
}

function hasOption(select: HTMLSelectElement, value: string): boolean {
  return Array.from(select.options).some((option) => option.value === value);
}

function mapSortErrorToUserMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : "";

  if (
    message.includes("No item entries found") ||
    message.includes("Invalid entry start") ||
    message.includes("Entries could not be parsed")
  ) {
    return "Could not read this reservation list. Please check the pasted text and click Sort again.";
  }

  const integrityMatch = message.match(/Parsing integrity check failed:\s*(\d+)\s+entries detected in source but only\s+(\d+)/i);
  if (integrityMatch) {
    const sourceCount = integrityMatch[1];
    const processedCount = integrityMatch[2];
    return `Processing failed: the source list has ${sourceCount} entries, but only ${processedCount} made it into the new list. Please check the pasted text and try again.`;
  }

  return "There was an error while sorting. Please reload the page and try again.";
}

function showOk(message: string): void {
  statusPanel.textContent = message;
  statusPanel.classList.remove("hidden", "error");
}

function showError(message: string): void {
  statusPanel.textContent = message;
  statusPanel.classList.remove("hidden");
  statusPanel.classList.add("error");
}

function resetDownloads(): void {
  downloadActions.classList.add("hidden");
}

function hideStatus(): void {
  statusPanel.textContent = "";
  statusPanel.classList.add("hidden");
  statusPanel.classList.remove("error");
}

function mustGet<T extends HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element: ${id}`);
  }

  return element as T;
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) {
    return `${bytes} B`;
  }

  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
