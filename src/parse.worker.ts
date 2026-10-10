import type { Cell, DataRow, Field, FileEntry, Issue, ParsedFile } from "./lib/types";

declare function importScripts(...urls: string[]): void;
declare const XLSX: {
  read(data: ArrayBuffer, options: Record<string, unknown>): { SheetNames: string[]; Sheets: Record<string, unknown> };
  write(book: unknown, options: Record<string, unknown>): ArrayBuffer | Uint8Array;
  utils: {
    sheet_to_json(sheet: unknown, options: Record<string, unknown>): Cell[][];
    book_new(): unknown;
    aoa_to_sheet(rows: unknown[][]): unknown;
    book_append_sheet(book: unknown, sheet: unknown, name: string): void;
  };
};

const worker = self as unknown as {
  XLSX?: typeof XLSX;
  postMessage(message: unknown, transfer?: Transferable[]): void;
  onmessage: ((event: MessageEvent<
    | { action: "parse"; files: FileEntry[] }
    | { action: "export"; rows: DataRow[]; issues: Issue[]; fields: Field[] }
  >) => void) | null;
};
const SHEETJS_URL = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

worker.onmessage = async (event) => {
  try {
    if (!worker.XLSX) importScripts(SHEETJS_URL);
    if (event.data.action === "parse") {
      const files = event.data.files || [];
      const parsed: ParsedFile[] = [];
      for (let index = 0; index < files.length; index += 1) {
        const entry = files[index];
        try {
          const workbook = XLSX.read(await entry.file.arrayBuffer(), { type: "array", dense: true, cellDates: true });
          const sheets = workbook.SheetNames.map((name) => ({
            name,
            matrix: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: "", raw: true }),
          }));
          parsed.push({ path: entry.path, name: entry.file.name, size: entry.file.size, sheets });
        } catch (error) {
          worker.postMessage({ type: "file-error", path: entry.path, message: error instanceof Error ? error.message : "Could not read workbook" });
        }
        worker.postMessage({ type: "progress", done: index + 1, total: files.length });
      }
      worker.postMessage({ type: "complete", files: parsed });
      return;
    }

    const { rows, issues, fields } = event.data;
    const header = [...fields.map((field) => field.name), "Source file", "Sheet", "Excel row"];
    const data = rows.map((row) => [
      ...fields.map((field) => {
        const value = row.values[field.id] || "";
        if (field.type === "number" && value.trim() && Number.isFinite(Number(value))) return Number(value);
        return value;
      }),
      row.source,
      row.sheet,
      row.excelRow,
    ]);
    const byId = new Map(rows.map((row) => [row.id, row]));
    const issueRows = issues.map((issue) => {
      const row = byId.get(issue.rowId);
      return [row?.source || "", row?.sheet || "", row?.excelRow || "", fields.find((field) => field.id === issue.fieldId)?.name || "", issue.type, issue.message];
    });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([header, ...data]), "Data");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["Source file", "Sheet", "Excel row", "Field", "Issue", "Details"], ...issueRows]), "Issues");
    const result = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const bytes = result instanceof ArrayBuffer
      ? result
      : result.buffer.slice(result.byteOffset, result.byteOffset + result.byteLength) as ArrayBuffer;
    worker.postMessage({ type: "complete-export", bytes }, [bytes]);
  } catch (error) {
    worker.postMessage({ type: "fatal", message: error instanceof Error ? error.message : "Excel parser could not load" });
  }
};
