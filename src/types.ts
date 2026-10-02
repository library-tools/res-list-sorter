export type Audience = "adult" | "junior";

export type Entry = {
  rawLines: string[];
  barcode: string;
  shelfmark: string;
  shelfSuffix: string;
  author: string;
  itemType: string;
  sequence: string;
  audience: Audience;
  originalIndex: number;
};
