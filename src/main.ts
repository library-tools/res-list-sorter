import JsBarcode from "jsbarcode";
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
  barcodeValue?: string;
  barcodeTight?: boolean;
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
const FAIRY_FOLK_MYT_SEQUENCE = "Children's Fairy /Folk/Myt";
const CHILDRENS_GRAPHIC_NOVELS_SEQUENCE = "Children's Graphic Novels";
const TEEN_GRAPHIC_NOVELS_SEQUENCE = "Teen Graphic Novels";
const BARCODE_IMAGE_TARGET_WIDTH_PT = 92;
const BARCODE_IMAGE_MIN_WIDTH_PT = 64;
const BARCODE_IMAGE_HEIGHT_FACTOR = 0.9;
const BARCODE_IMAGE_MIN_HEIGHT_PT = 12;
const BARCODE_IMAGE_MAX_HEIGHT_PT = 18;
const BARCODE_NORMAL_GAP_FACTOR = 2;
const BARCODE_TIGHT_GAP_FACTOR = 1;
const BARCODE_MIN_GAP_PT = 2;

let latestResult: SortResult | null = null;
let lastSortedSource: string | null = null;
let saveSettingsResetTimer: number | null = null;

loadSavedSettings();

saveSettingsButton.addEventListener("click", () => {
  if (saveCurrentSettings()) {
    showSavedSettingsFeedback();
  } else {
    showError("Could not save settings in this browser.");
  }
});

sourceText.addEventListener("input", () => {
  hideStatus();

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

    showOk(statusLine);
    lastSortedSource = sourceText.value;

    downloadActions.classList.remove("hidden");
  } catch (error) {
    console.error(error);
    showError(mapSortErrorToUserMessage(error));
  }
});

downloadAdultButton.addEventListener("click", () => {
  if (!latestResult) {
    showError("There was an error. Please reload the page and try again.");
    return;
  }

  try {
    downloadPdf("adult-list.pdf", latestResult.adultDoc, getPdfTextSize(), getPdfColumns(), getPdfFontFamily());
  } catch (error) {
    console.error(error);
    showError("There was an error while creating the PDF. Please try again.");
  }
});

downloadJuniorButton.addEventListener("click", () => {
  if (!latestResult) {
    showError("There was an error. Please reload the page and try again.");
    return;
  }

  try {
    downloadPdf("junior-list.pdf", latestResult.juniorDoc, getPdfTextSize(), getPdfColumns(), getPdfFontFamily());
  } catch (error) {
    console.error(error);
    showError("There was an error while creating the PDF. Please try again.");
  }
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

  if (entries.length === 0) {
    throw new Error("Entries could not be parsed.");
  }

  return { libraryName, reportDate, entries };
}

function extractHeaderMetadata(headerLines: string[]): { libraryName: string; reportDate: string } {
  let libraryName = "Unknown Library";
  let reportDate = "Unknown date";
  const reportTitle = "Reserved Items not on loan by site";

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

  if (reportDate === "Unknown date") {
    for (let i = 0; i < headerLines.length; i += 1) {
      const dateMatch = headerLines[i].match(/\b(\d{2}\/\d{2}\/\d{2})\b/);
      if (!dateMatch) {
        continue;
      }

      let j = i + 1;
      while (j < headerLines.length && headerLines[j].trim() === "") {
        j += 1;
      }

      if (j < headerLines.length && headerLines[j].trim().toLowerCase() === reportTitle.toLowerCase()) {
        reportDate = dateMatch[1];
        break;
      }
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

  const adultDoc = buildDocument("Adult", parsed.libraryName, parsed.reportDate, adultEntries);
  const juniorDoc = buildDocument("Junior", parsed.libraryName, parsed.reportDate, juniorEntries);

  return {
    originalCount,
    adultCount,
    juniorCount,
    adultDoc,
    juniorDoc,
  };
}

function compareEntries(a: Entry, b: Entry): number {
  const aSequenceForSort = normalizeSequenceForSorting(a.itemType, a.sequence);
  const bSequenceForSort = normalizeSequenceForSorting(b.itemType, b.sequence);

  const specialSequenceCompare = compareSpecialSequenceBucket(aSequenceForSort, bSequenceForSort);
  if (specialSequenceCompare !== 0) {
    return specialSequenceCompare;
  }

  if (isSpecialSequence(aSequenceForSort) && isSpecialSequence(bSequenceForSort)) {
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

  const typeBucketCompare = compareTypeBucket(a.itemType, b.itemType);
  if (typeBucketCompare !== 0) {
    return typeBucketCompare;
  }

  const itemTypeCompare = compareText(a.itemType, b.itemType);
  if (itemTypeCompare !== 0) {
    return itemTypeCompare;
  }

  const sequenceCompare = compareSequence(
    aSequenceForSort,
    bSequenceForSort,
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

function compareSpecialSequenceBucket(aSequence: string, bSequence: string): number {
  const aSpecialName = specialSequenceName(aSequence);
  const bSpecialName = specialSequenceName(bSequence);

  if (!aSpecialName && !bSpecialName) {
    return 0;
  }

  if (aSpecialName && !bSpecialName) {
    return 1;
  }

  if (!aSpecialName && bSpecialName) {
    return -1;
  }

  return compareText(aSpecialName!, bSpecialName!);
}

function isSpecialSequence(sequence: string): boolean {
  return specialSequenceName(sequence) !== null;
}

function specialSequenceName(sequence: string): string | null {
  const trimmed = sequence.trim();

  if (
    trimmed === FAIRY_FOLK_MYT_SEQUENCE ||
    trimmed === CHILDRENS_GRAPHIC_NOVELS_SEQUENCE ||
    trimmed === TEEN_GRAPHIC_NOVELS_SEQUENCE
  ) {
    return trimmed;
  }

  return null;
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
    const groupKey = renderGroupKey(entry.itemType, effectiveSequence);

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

    blocks.push({
      lines: [...cleanedEntryLines, makePlainLine("")],
      sectionHeading: currentSectionHeading!,
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
  const specialSequence = specialSequenceName(effectiveSequence);
  if (specialSequence) {
    return `${specialSequence} —`;
  }

  if (effectiveSequence.trim() === "") {
    return `${itemType} —`;
  }

  return `${itemType} — ${effectiveSequence}`;
}

function renderGroupKey(itemType: string, effectiveSequence: string): string {
  const specialSequence = specialSequenceName(effectiveSequence);
  if (specialSequence) {
    return `sequence\u0000${specialSequence}`;
  }

  return `${itemType}\u0000${effectiveSequence}`;
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
    const lineLabel = formatFirstRowLabel(firstMatch[1]);
    lines.push({
      segments: [{ text: lineLabel, style: "normal" }],
      rightText: firstMatch[2],
      rightStyle: "normal",
      barcodeValue: firstMatch[2],
    });
  } else {
    lines.push(makePlainLine(first));
  }

  const remaining = nonBlank.slice(1).map((line) => line.replace(/^\s+/, ""));
  const metadataStart = remaining.findIndex((line) => isMetadataLine(line));
  const contentBeforeMetadata = metadataStart >= 0 ? remaining.slice(0, metadataStart) : remaining;
  const hasAuthorLine = (rawLines[1] ?? "").trim() !== "";
  const authorLine = hasAuthorLine ? contentBeforeMetadata[0] ?? "" : "";
  const titleLines = hasAuthorLine ? contentBeforeMetadata.slice(1) : contentBeforeMetadata;
  const joinedTitle = titleLines.join(" ").trim();

  if (authorLine.trim() !== "") {
    lines.push(makePlainLine(authorLine));
  }

  if (joinedTitle !== "") {
    lines.push(makeStyledLine(joinedTitle, "italic"));
  }

  const metadataLines = metadataStart >= 0 ? remaining.slice(metadataStart) : [];
  for (const line of metadataLines) {
    if (shouldOmitEntryLine(line)) {
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
  const barcodeImageCache = new Map<string, string>();

  const margin = 36;
  const textHeight = doc.getTextDimensions("Mg").h;
  const lineHeight = Math.ceil(textHeight * 1.25);
  const fitSlack = Math.min(4, Math.floor(lineHeight * 0.2));
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
  const contentBottomY = footerY;

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

  const continuedSectionHeadingLines = (sectionHeading: RenderLine): RenderLine[] => {
    const leftX = columnStartXs[columnIndex];
    const continuedHeading = makeStyledLine(`${lineText(sectionHeading)} (cont.)`, "bold");
    const headingLines: RenderLine[] = [
      ...wrapRenderLine(doc, continuedHeading, columnWidth, rightAlignedWidth, fontFamily),
      makePlainLine(""),
    ];

    if (totalRenderHeight(headingLines, lineHeight) > contentBottomY - contentStartY) {
      abortRender("Section heading is too long to render safely. Please reduce heading length and try again.");
      return [];
    }

    return headingLines;
  };

  const drawContinuedSectionHeading = (sectionHeading: RenderLine): void => {
    const leftX = columnStartXs[columnIndex];
    const barcodeRightX = leftX + rightAlignedWidth;
    const headingLines = continuedSectionHeadingLines(sectionHeading);
    if (renderAborted || headingLines.length === 0) {
      return;
    }

    if (y + totalRenderHeight(headingLines, lineHeight) > contentBottomY) {
      advanceColumnOrPage();
      if (renderAborted) {
        return;
      }
    }

    drawWrappedLines(doc, headingLines, leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily, textHeight, barcodeImageCache);
    y += totalRenderHeight(headingLines, lineHeight);
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
    drawWrappedLines(doc, [line], leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily, textHeight, barcodeImageCache);
    y += renderLineHeight(line, lineHeight);
  };

  for (const block of wrappedBlocks) {
    let linesToDraw = block.lines;
    const trailingBlankCount = countTrailingBlankLines(linesToDraw);
    if (trailingBlankCount > 0) {
      const trimmedLines = linesToDraw.slice(0, linesToDraw.length - trailingBlankCount);
      const fullHeight = totalRenderHeight(linesToDraw, lineHeight);
      const trimmedHeight = totalRenderHeight(trimmedLines, lineHeight);
      const remaining = contentBottomY - y;

      if (fullHeight > remaining + fitSlack && trimmedHeight <= remaining + fitSlack) {
        linesToDraw = trimmedLines;
      }
    }

    const blockHeight = totalRenderHeight(linesToDraw, lineHeight);
    const remaining = contentBottomY - y;

    if (blockHeight > contentBottomY - contentStartY) {
      for (const line of linesToDraw) {
        if (y + renderLineHeight(line, lineHeight) > contentBottomY) {
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
      drawWrappedLines(doc, linesToDraw, leftX, lineHeight, y, barcodeRightX, continuationIndent, fontFamily, textHeight, barcodeImageCache);
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
    setFontStyle(doc, style, fontFamily);
    const normalGapWidth = Math.max(BARCODE_MIN_GAP_PT, doc.getTextWidth(" ") * BARCODE_NORMAL_GAP_FACTOR);
    const tightGapWidth = Math.max(BARCODE_MIN_GAP_PT, doc.getTextWidth(" ") * BARCODE_TIGHT_GAP_FACTOR);
    const barcodeImageWidth = line.barcodeValue ? barcodeImageWidthPt(rightAlignedWidth) : 0;

    if (line.barcodeValue) {
      const labelWidth = doc.getTextWidth(fullText);
      const continuationIndentWidth = Math.max(BARCODE_MIN_GAP_PT, doc.getTextWidth("0") * 2);
      const fullLineWidth = labelWidth + normalGapWidth + barcodeImageWidth + normalGapWidth + barcodeWidth;
      if (fullLineWidth <= rightAlignedWidth) {
        return [{ segments: [{ text: fullText, style }], rightText: line.rightText, rightStyle, barcodeValue: line.barcodeValue }];
      }

      const tightLineWidth = labelWidth + tightGapWidth + barcodeImageWidth + tightGapWidth + barcodeWidth;
      if (tightLineWidth <= rightAlignedWidth) {
        return [{
          segments: [{ text: fullText, style }],
          rightText: line.rightText,
          rightStyle,
          barcodeValue: line.barcodeValue,
          barcodeTight: true,
        }];
      }

      const firstLineWidth = labelWidth + tightGapWidth + barcodeImageWidth;
      if (firstLineWidth <= rightAlignedWidth) {
        return [
          { segments: [{ text: fullText, style }], barcodeValue: line.barcodeValue, barcodeTight: true },
          { segments: [{ text: "", style }], rightText: line.rightText, rightStyle, continuation: true },
        ];
      }

      const finalLabelWidth = Math.max(20, rightAlignedWidth - barcodeImageWidth - tightGapWidth - continuationIndentWidth);
      const wrappedLeft = doc.splitTextToSize(fullText, finalLabelWidth) as string[];

      if (wrappedLeft.length <= 1) {
        return [
          { segments: [{ text: wrappedLeft[0] ?? fullText, style }], barcodeValue: line.barcodeValue, barcodeTight: true },
          { segments: [{ text: "", style }], rightText: line.rightText, rightStyle, continuation: true },
        ];
      }

      const prefix = wrappedLeft.slice(0, -1).map((text, index) => ({
        segments: [{ text, style }],
        continuation: index > 0,
      }));
      const finalLabelLine: RenderLine = {
        segments: [{ text: wrappedLeft[wrappedLeft.length - 1], style }],
        barcodeValue: line.barcodeValue,
        continuation: true,
        barcodeTight: true,
      };

      return [
        ...prefix,
        finalLabelLine,
        { segments: [{ text: "", style }], rightText: line.rightText, rightStyle, continuation: true },
      ];
    }

    const leftMaxWidth = Math.max(20, rightAlignedWidth - barcodeWidth - tightGapWidth);
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
  textHeight: number,
  barcodeImageCache: Map<string, string>,
): void {
  let y = startY;

  for (const line of lines) {
    let x = margin + (line.continuation ? continuationIndent : 0);
    const currentLineHeight = renderLineHeight(line, lineHeight);
    for (const segment of line.segments) {
      if (segment.text === "") {
        continue;
      }

      setFontStyle(doc, segment.style, fontFamily);
      doc.text(segment.text, x, y);
      x += doc.getTextWidth(segment.text);
    }

    if (line.barcodeValue) {
      const normalGapWidth = Math.max(BARCODE_MIN_GAP_PT, doc.getTextWidth(" ") * BARCODE_NORMAL_GAP_FACTOR);
      const tightGapWidth = Math.max(BARCODE_MIN_GAP_PT, doc.getTextWidth(" ") * BARCODE_TIGHT_GAP_FACTOR);
      const gapWidth = line.barcodeTight ? tightGapWidth : normalGapWidth;
      const imageWidth = barcodeImageWidthPt(barcodeRightX - margin);
      const imageHeight = barcodeImageHeightPt(lineHeight);
      const rightTextWidth = doc.getTextWidth(line.rightText ?? "");
      const centeredX = x + Math.max(0, (barcodeRightX - rightTextWidth - x - imageWidth) / 2);
      const minImageX = x + gapWidth;
      const maxImageX = barcodeRightX - rightTextWidth - gapWidth - imageWidth;
      const imageX = Math.max(minImageX, Math.min(centeredX, maxImageX));
      const imageY = y - textHeight + Math.max(0, (currentLineHeight - imageHeight) / 2);
      const barcodeDataUrl = getBarcodeDataUrl(line.barcodeValue, imageWidth, imageHeight, barcodeImageCache);
      doc.addImage(barcodeDataUrl, "PNG", imageX, imageY, imageWidth, imageHeight);
    }

    if (line.rightText) {
      setFontStyle(doc, line.rightStyle ?? "normal", fontFamily);
      doc.text(line.rightText, barcodeRightX, y, { align: "right" });
    }

    y += currentLineHeight;
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

function renderLineHeight(line: RenderLine, baseLineHeight: number): number {
  if (!line.barcodeValue) {
    return baseLineHeight;
  }

  return Math.max(baseLineHeight, barcodeImageHeightPt(baseLineHeight) + 4);
}

function totalRenderHeight(lines: RenderLine[], baseLineHeight: number): number {
  return lines.reduce((sum, line) => sum + renderLineHeight(line, baseLineHeight), 0);
}

function barcodeImageWidthPt(availableWidth: number): number {
  return Math.max(BARCODE_IMAGE_MIN_WIDTH_PT, Math.min(BARCODE_IMAGE_TARGET_WIDTH_PT, availableWidth * 0.32));
}

function barcodeImageHeightPt(baseLineHeight: number): number {
  return Math.max(BARCODE_IMAGE_MIN_HEIGHT_PT, Math.min(BARCODE_IMAGE_MAX_HEIGHT_PT, baseLineHeight * BARCODE_IMAGE_HEIGHT_FACTOR));
}

function getBarcodeDataUrl(barcodeValue: string, widthPt: number, heightPt: number, cache: Map<string, string>): string {
  const cacheKey = `${barcodeValue}:${Math.round(widthPt)}:${Math.round(heightPt)}`;
  const cached = cache.get(cacheKey);
  if (cached) {
    return cached;
  }

  try {
    const canvas = document.createElement("canvas");
    const scale = 2;
    const widthPx = Math.max(160, Math.round(widthPt * scale));
    const heightPx = Math.max(52, Math.round(heightPt * scale));
    canvas.width = widthPx;
    canvas.height = heightPx;
    const moduleWidth = Math.max(1, Math.round(widthPt / 40));

    JsBarcode(canvas, barcodeValue, {
      format: "CODE128",
      displayValue: false,
      margin: 0,
      width: moduleWidth,
      height: Math.max(28, heightPx - 6),
    });

    const dataUrl = canvas.toDataURL("image/png");
    cache.set(cacheKey, dataUrl);
    return dataUrl;
  } catch (error) {
    throw new Error(
      `Barcode rendering failed for ${barcodeValue}: ${error instanceof Error ? error.message : "unknown error"}`,
    );
  }
}

function formatFirstRowLabel(preBarcode: string): string {
  const trimmed = preBarcode.trim();
  if (trimmed === "") {
    return trimmed;
  }

  const parts = trimmed.split(/\s+/);
  const shelfmark = parts[0] ?? "";
  const suffix = parts.slice(1).join(" ");
  return suffix ? `${shelfmark}/${suffix}` : shelfmark;
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

function saveCurrentSettings(): boolean {
  const settings = {
    font: getPdfFontFamily(),
    textSize: String(getPdfTextSize()),
    columns: String(getPdfColumns()),
  };

  try {
    window.localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
    return true;
  } catch {
    return false;
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
    return "Could not read this reservation list. Please check the pasted text and click Create PDFs again.";
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
