import type { Entry } from "./types";

// Branch-specific shelving choices, set from the page controls.
export type ShelvingOptions = {
  classicsSeparate: boolean;
  cultWithClassics: boolean;
  thrillersWith: "crime" | "fiction";
};

export const DEFAULT_SHELVING_OPTIONS: ShelvingOptions = {
  classicsSeparate: true,
  cultWithClassics: false,
  thrillersWith: "crime",
};

const FAIRY_FOLK_MYT_SEQUENCE = "Children's Fairy /Folk/Myt";
const CHILDRENS_GRAPHIC_NOVELS_SEQUENCE = "Children's Graphic Novels";
const TEEN_GRAPHIC_NOVELS_SEQUENCE = "Teen Graphic Novels";
const CLASSICS_SEQUENCE = "Classics";
const CLASSICS_AND_CULT_SEQUENCE = "Classics/Cult";

export function compareEntries(a: Entry, b: Entry, options: ShelvingOptions): number {
  const aSequenceForSort = effectiveSequence(a, options);
  const bSequenceForSort = effectiveSequence(b, options);

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

export function effectiveSequence(entry: Entry, options: ShelvingOptions): string {
  const { itemType, sequence, shelfSuffix } = entry;

  if (options.classicsSeparate) {
    const classicsSection = options.cultWithClassics ? CLASSICS_AND_CULT_SEQUENCE : CLASSICS_SEQUENCE;

    if (isClassicsSuffix(shelfSuffix)) {
      return classicsSection;
    }

    if (options.cultWithClassics && isCultSuffix(shelfSuffix)) {
      return classicsSection;
    }
  }

  if (!/^adult fiction$/i.test(itemType.trim())) {
    return sequence;
  }

  const normalized = sequence.trim().toLowerCase();

  if (normalized === "thriller") {
    return options.thrillersWith === "crime" ? "Crime" : "";
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

function isClassicsSuffix(shelfSuffix: string): boolean {
  return shelfSuffix.trim().toLowerCase() === CLASSICS_SEQUENCE.toLowerCase();
}

function isCultSuffix(shelfSuffix: string): boolean {
  return /\bcult\b/i.test(shelfSuffix);
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

export function specialSequenceName(sequence: string): string | null {
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
