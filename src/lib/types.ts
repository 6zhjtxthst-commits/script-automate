export type Cell = string | number | boolean | Date | null | undefined;
export type FieldType = "text" | "number" | "date" | "email";

export interface Field {
  id: string;
  name: string;
  type: FieldType;
  required: boolean;
}

export interface SourceSheet {
  id: string;
  path: string;
  fileName: string;
  name: string;
  size: number;
  matrix: Cell[][];
  headerRow: number;
  included: boolean;
  revision: number;
}

export interface ParsedFile {
  path: string;
  name: string;
  size: number;
  sheets: Array<{ name: string; matrix: Cell[][] }>;
}

export interface Project {
  version: number;
  sheets: SourceSheet[];
  fields: Field[];
  mappings: Record<string, string>;
  overrides: Record<string, string>;
  edits: Record<string, string>;
  placeholders: string[];
  duplicateFields: string[];
  rowKeyField: string;
}

export interface DataRow {
  id: string;
  values: Record<string, string>;
  source: string;
  sheet: string;
  excelRow: number;
}

export interface Issue {
  rowId: string;
  fieldId: string;
  type: "missing" | "invalid" | "duplicate";
  message: string;
}

export interface FileEntry {
  path: string;
  file: File;
}

export interface WritableFileHandle {
  createWritable(): Promise<{ write(data: Blob | BufferSource): Promise<void>; close(): Promise<void> }>;
}

export interface LocalDirectoryHandle {
  kind: "directory";
  name: string;
  entries(): AsyncIterable<[string, LocalDirectoryHandle | LocalFileHandle]>;
  getFileHandle(name: string, options?: { create?: boolean }): Promise<WritableFileHandle>;
  queryPermission(options?: { mode: "read" | "readwrite" }): Promise<PermissionState>;
  requestPermission(options?: { mode: "read" | "readwrite" }): Promise<PermissionState>;
}

export interface LocalFileHandle {
  kind: "file";
  getFile(): Promise<File>;
}
