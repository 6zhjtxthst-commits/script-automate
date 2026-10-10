import type { Cell, DataRow, Field, Issue, ParsedFile, Project } from "./types";

export const PROJECT_VERSION = 1;

export function emptyProject(): Project {
  return {
    version: PROJECT_VERSION,
    sheets: [],
    fields: [],
    mappings: {},
    overrides: {},
    edits: {},
    placeholders: ["N/A", "NA", "-"],
    duplicateFields: [],
    rowKeyField: "",
  };
}

export function text(value: unknown): string {
  if (value == null) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  return String(value);
}

export function normalizeHeader(value: unknown): string {
  return text(value).trim().toLocaleLowerCase().replace(/[\s_\-]+/g, " ");
}

export function isMissing(value: unknown, placeholders: string[] = []): boolean {
  const normalized = text(value).trim().toLocaleLowerCase();
  return !normalized || placeholders.some((item) => normalized === text(item).trim().toLocaleLowerCase());
}

export function guessHeaderRow(matrix: Cell[][]): number {
  let bestRow = 0;
  let bestScore = -1;
  matrix.slice(0, 15).forEach((row, index) => {
    const filled = row.map(text).map((value) => value.trim()).filter(Boolean);
    const nextFilled = (matrix[index + 1] || []).map(text).filter((value) => value.trim()).length;
    const score = filled.length >= 2 ? filled.length * 2 + Math.min(nextFilled, filled.length) : -1;
    if (score > bestScore) {
      bestScore = score;
      bestRow = index;
    }
  });
  return bestRow;
}

export function sheetId(path: string, name: string): string {
  return `${path}\u001f${name}`;
}

export function columnKey(sheet: { id: string }, index: number): string {
  return `${sheet.id}\u001f${index}`;
}

function nextFieldId(fields: Field[]): string {
  let number = fields.length + 1;
  while (fields.some((field) => field.id === `f${number}`)) number += 1;
  return `f${number}`;
}

export function mergeFiles(project: Project, files: ParsedFile[]): Project {
  const incomingPaths = new Set(files.map((file) => file.path));
  const oldSheets = new Map(project.sheets.map((sheet) => [sheet.id, sheet]));
  const sheets = project.sheets.filter((sheet) => !incomingPaths.has(sheet.path));
  const fields = [...project.fields];
  const mappings = { ...project.mappings };

  files.forEach((file) => {
    file.sheets.forEach((sourceSheet) => {
      const id = sheetId(file.path, sourceSheet.name);
      const previous = oldSheets.get(id);
      const matrix = sourceSheet.matrix || [];
      const headerRow = previous && previous.headerRow < matrix.length
        ? previous.headerRow
        : guessHeaderRow(matrix);
      const sheet = {
        id,
        path: file.path,
        fileName: file.name,
        name: sourceSheet.name,
        size: file.size,
        matrix,
        headerRow,
        included: previous?.included ?? true,
        revision: (previous?.revision ?? 0) + 1,
      };
      sheets.push(sheet);

      (matrix[headerRow] || []).forEach((rawHeader) => {
        const name = text(rawHeader).trim();
        const normalized = normalizeHeader(name);
        if (!normalized || mappings[normalized]) return;
        const existing = fields.find((field) => normalizeHeader(field.name) === normalized);
        const field: Field = existing || { id: nextFieldId(fields), name, type: "text", required: false };
        if (!existing) fields.push(field);
        mappings[normalized] = field.id;
      });
    });
  });

  return { ...project, sheets, fields, mappings };
}

export function ensureFields(project: Project): Project {
  const fields = [...project.fields];
  const mappings = { ...project.mappings };
  project.sheets.forEach((sheet) => {
    (sheet.matrix[sheet.headerRow] || []).forEach((rawHeader) => {
      const name = text(rawHeader).trim();
      const normalized = normalizeHeader(name);
      if (!normalized || mappings[normalized]) return;
      const existing = fields.find((field) => normalizeHeader(field.name) === normalized);
      const field: Field = existing || { id: nextFieldId(fields), name, type: "text", required: false };
      if (!existing) fields.push(field);
      mappings[normalized] = field.id;
    });
  });
  return { ...project, fields, mappings };
}

export function buildRows(project: Project): DataRow[] {
  const fieldIds = new Set(project.fields.map((field) => field.id));
  const rows: DataRow[] = [];

  project.sheets.filter((sheet) => sheet.included).forEach((sheet) => {
    const headers = sheet.matrix[sheet.headerRow] || [];
    const seenKeys = new Map();
    for (let index = sheet.headerRow + 1; index < sheet.matrix.length; index += 1) {
      const source = sheet.matrix[index] || [];
      if (!source.some((cell) => text(cell).trim())) continue;
      const values: Record<string, string> = {};
      headers.forEach((header, column) => {
        const fieldId = project.overrides[columnKey(sheet, column)] || project.mappings[normalizeHeader(header)];
        if (!fieldIds.has(fieldId)) return;
        const value = text(source[column]);
        if (values[fieldId] == null || (!text(values[fieldId]).trim() && value.trim())) values[fieldId] = value;
      });

      const keyValue = project.rowKeyField ? text(values[project.rowKeyField]).trim() : "";
      const occurrence = keyValue ? (seenKeys.get(keyValue) || 0) + 1 : 0;
      if (keyValue) seenKeys.set(keyValue, occurrence);
      const id = keyValue
        ? `${sheet.id}\u001fkey:${keyValue}\u001f${occurrence}`
        : `${sheet.id}\u001frev:${sheet.revision}\u001frow:${index + 1}`;
      project.fields.forEach((field) => {
        const editKey = `${id}\u001f${field.id}`;
        if (Object.hasOwn(project.edits, editKey)) values[field.id] = project.edits[editKey];
      });
      rows.push({ id, values, source: sheet.path, sheet: sheet.name, excelRow: index + 1 });
    }
  });
  return rows;
}

export function cellEditKey(rowId: string, fieldId: string): string {
  return `${rowId}\u001f${fieldId}`;
}

export function valueIssue(value: unknown, field: Field, placeholders: string[]): string {
  if (isMissing(value, placeholders)) return field.required ? "Missing required value" : "";
  const content = text(value).trim();
  if (field.type === "number" && !Number.isFinite(Number(content))) return "Expected a number";
  if (field.type === "date" && Number.isNaN(Date.parse(content))) return "Expected a date";
  if (field.type === "email" && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(content)) return "Expected an email address";
  return "";
}

export function collectIssues(rows: DataRow[], project: Project): Issue[] {
  const issues: Issue[] = [];
  const duplicateGroups = new Map<string, DataRow[]>();
  const duplicateFields = project.duplicateFields.filter((id) => project.fields.some((field) => field.id === id));

  rows.forEach((row) => {
    project.fields.forEach((field) => {
      const message = valueIssue(row.values[field.id], field, project.placeholders);
      if (message) issues.push({ rowId: row.id, fieldId: field.id, type: message.startsWith("Missing") ? "missing" : "invalid", message });
    });
    if (duplicateFields.length) {
      const values = duplicateFields.map((id) => text(row.values[id]).trim().toLocaleLowerCase());
      if (values.every((value) => value && !isMissing(value, project.placeholders))) {
        const key = JSON.stringify(values);
        duplicateGroups.set(key, [...(duplicateGroups.get(key) || []), row]);
      }
    }
  });

  duplicateGroups.forEach((group) => {
    if (group.length < 2) return;
    group.forEach((row) => issues.push({ rowId: row.id, fieldId: "", type: "duplicate", message: `Possible duplicate (${group.length} rows)` }));
  });
  return issues;
}

export function validProject(value: unknown): value is Project {
  const candidate = value as Project | null;
  return candidate?.version === PROJECT_VERSION && Array.isArray(candidate.sheets) &&
    Array.isArray(candidate.fields) && typeof candidate.mappings === "object" &&
    typeof candidate.edits === "object";
}
