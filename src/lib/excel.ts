import type { DataRow, Field, FileEntry, Issue, LocalDirectoryHandle, ParsedFile } from "./types";

export function parseWorkbooks(
  files: FileEntry[],
  onProgress?: (done: number, total: number) => void,
  onFileError?: (path: string, message: string) => void,
): Promise<ParsedFile[]> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("../parse.worker.ts", import.meta.url), { type: "classic" });
    worker.onmessage = ({ data }: MessageEvent<{ type: string; done: number; total: number; path: string; message: string; files: ParsedFile[] }>) => {
      if (data.type === "progress") onProgress?.(data.done, data.total);
      if (data.type === "file-error") onFileError?.(data.path, data.message);
      if (data.type === "fatal") {
        worker.terminate();
        reject(new Error(data.message));
      }
      if (data.type === "complete") {
        worker.terminate();
        resolve(data.files);
      }
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "Excel worker failed"));
    };
    worker.postMessage({ action: "parse", files });
  });
}

export function workbookBytes(rows: DataRow[], issues: Issue[], fields: Field[]): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("../parse.worker.ts", import.meta.url), { type: "classic" });
    worker.onmessage = ({ data }: MessageEvent<{ type: string; bytes?: ArrayBuffer; message?: string }>) => {
      if (data.type === "complete-export" && data.bytes) {
        worker.terminate();
        resolve(data.bytes);
      }
      if (data.type === "fatal") {
        worker.terminate();
        reject(new Error(data.message || "Could not create workbook"));
      }
    };
    worker.onerror = (event) => {
      worker.terminate();
      reject(new Error(event.message || "Excel worker failed"));
    };
    worker.postMessage({ action: "export", rows, issues, fields });
  });
}

export function downloadBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

export async function readDirectory(handle: LocalDirectoryHandle, prefix = ""): Promise<FileEntry[]> {
  const files: FileEntry[] = [];
  for await (const [name, entry] of handle.entries()) {
    const relative = prefix ? `${prefix}/${name}` : name;
    if (entry.kind === "directory") files.push(...await readDirectory(entry, relative));
    if (entry.kind === "file" && /\.xlsx?$/i.test(name) && !name.startsWith("~$")) {
      files.push({ path: relative, file: await entry.getFile() });
    }
  }
  return files;
}
