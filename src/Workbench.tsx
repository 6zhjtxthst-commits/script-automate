import { useEffect, useMemo, useRef, useState, type ChangeEvent } from "react";
import {
  buildRows, cellEditKey, collectIssues, columnKey, emptyProject, ensureFields,
  mergeFiles, normalizeHeader, text, validProject,
} from "./lib/data";
import { downloadBlob, parseWorkbooks, readDirectory, workbookBytes } from "./lib/excel";
import { loadLocal, saveLocal } from "./lib/storage";
import type { DataRow, Field, FieldType, FileEntry, Issue, LocalDirectoryHandle, Project, SourceSheet } from "./lib/types";
import "../workbench.css";

type View = "sources" | "fields" | "review" | "export";
type ReviewFilter = "all" | "issues" | "missing" | "duplicates";
type UndoAction = Record<string, string | undefined>;

const viewMeta: Record<View, { title: string; description: string; number: string }> = {
  sources: { title: "Bring your workbooks together.", description: "Add files or choose a folder. Review sheets and header rows before mapping.", number: "01" },
  fields: { title: "Make the columns agree.", description: "Keep every source column, then decide which ones mean the same thing.", number: "02" },
  review: { title: "Fill the gaps.", description: "Correct values, review duplicates, and apply one answer to several rows.", number: "03" },
  export: { title: "Take the clean copy.", description: "Export a consolidated workbook with an issues sheet and source traceability.", number: "04" },
};

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function excelName(prefix: string, extension: string): string {
  const stamp = new Date().toISOString().slice(0, 16).replace("T", "-").replace(":", "");
  return `${prefix}-${stamp}.${extension}`;
}

function CellEditor({ value, changed, issue, label, onCommit }: {
  value: string;
  changed: boolean;
  issue?: Issue;
  label: string;
  onCommit: (value: string) => void;
}) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  return (
    <input
      className={`cell-input${issue ? " cell-input-issue" : ""}${changed ? " cell-input-changed" : ""}`}
      aria-label={label}
      aria-invalid={Boolean(issue)}
      title={issue?.message || (changed ? "Edited value" : "")}
      value={draft}
      onChange={(event) => setDraft(event.target.value)}
      onBlur={() => { if (draft !== value) onCommit(draft); }}
      onKeyDown={(event) => { if (event.key === "Enter") event.currentTarget.blur(); }}
    />
  );
}

function StatusBadge({ issues }: { issues: Issue[] }) {
  if (!issues.length) return <span className="status-badge status-good">Ready</span>;
  const duplicate = issues.some((issue) => issue.type === "duplicate");
  return <span className="status-badge status-warn">{duplicate ? "Review" : "Needs input"}</span>;
}

export default function Workbench() {
  const [project, setProject] = useState<Project>(emptyProject);
  const [hydrated, setHydrated] = useState(false);
  const [view, setView] = useState<View>("sources");
  const [folderHandle, setFolderHandle] = useState<LocalDirectoryHandle | null>(null);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState("");
  const [saveState, setSaveState] = useState("Opening workspace");
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [fileErrors, setFileErrors] = useState<string[]>([]);
  const [filter, setFilter] = useState<ReviewFilter>("all");
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(1);
  const [selectedRows, setSelectedRows] = useState<Set<string>>(new Set());
  const [bulkField, setBulkField] = useState("");
  const [bulkValue, setBulkValue] = useState("");
  const [undoStack, setUndoStack] = useState<UndoAction[]>([]);
  const [placeholderDraft, setPlaceholderDraft] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const folderInputRef = useRef<HTMLInputElement>(null);
  const projectInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    folderInputRef.current?.setAttribute("webkitdirectory", "");
    Promise.all([loadLocal<Project>("project"), loadLocal<LocalDirectoryHandle>("folder")])
      .then(([savedProject, savedFolder]) => {
        if (validProject(savedProject)) setProject(savedProject);
        if (savedFolder) setFolderHandle(savedFolder);
      })
      .catch(() => setNotice("Local autosave is unavailable in this browser session."))
      .finally(() => { setHydrated(true); setSaveState("Saved on this device"); });
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    setSaveState("Saving…");
    const timer = window.setTimeout(() => {
      saveLocal("project", project)
        .then(() => setSaveState("Saved on this device"))
        .catch(() => setSaveState("Autosave unavailable — export a project file"));
    }, 650);
    return () => window.clearTimeout(timer);
  }, [project, hydrated]);

  useEffect(() => setPlaceholderDraft(project.placeholders.join(", ")), [project.placeholders]);

  const rows = useMemo(() => buildRows(project), [project]);
  const issues = useMemo(() => collectIssues(rows, project), [rows, project]);
  const issueMap = useMemo(() => {
    const map = new Map<string, Issue[]>();
    issues.forEach((issue) => map.set(issue.rowId, [...(map.get(issue.rowId) || []), issue]));
    return map;
  }, [issues]);
  const fileCount = useMemo(() => new Set(project.sheets.map((sheet) => sheet.path)).size, [project.sheets]);
  const includedSheets = project.sheets.filter((sheet) => sheet.included).length;
  const missingCount = issues.filter((issue) => issue.type === "missing").length;
  const duplicateCount = issues.filter((issue) => issue.type === "duplicate").length;

  async function importEntries(entries: FileEntry[]) {
    const excelFiles = entries.filter((entry) => /\.xlsx?$/i.test(entry.file.name) && !entry.file.name.startsWith("~$"));
    if (!excelFiles.length) { setError("No .xlsx or .xls files were found."); return; }
    setBusy(true);
    setError("");
    setNotice("");
    setFileErrors([]);
    setProgress(`Reading 0 of ${excelFiles.length} files`);
    const errors: string[] = [];
    try {
      const files = await parseWorkbooks(
        excelFiles,
        (done, total) => setProgress(`Reading ${done} of ${total} files`),
        (path, message) => errors.push(`${path}: ${message}`),
      );
      if (!files.length) throw new Error("No workbook could be read. Check the file format and try again.");
      setProject((current) => mergeFiles(current, files));
      setFileErrors(errors);
      setSelectedRows(new Set());
      setPage(1);
      setView("sources");
      setNotice(`Imported ${files.length} ${files.length === 1 ? "file" : "files"}. Review the sheets and header rows below.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Import failed.");
    } finally {
      setBusy(false);
      setProgress("");
    }
  }

  function handleFileInput(event: ChangeEvent<HTMLInputElement>, fromFolder: boolean) {
    const entries = Array.from(event.target.files || []).map((file) => ({
      path: fromFolder ? (file as File & { webkitRelativePath?: string }).webkitRelativePath || file.name : file.name,
      file,
    }));
    event.target.value = "";
    void importEntries(entries);
  }

  async function selectFolder() {
    const browser = window as Window & { showDirectoryPicker?: (options: { mode: "readwrite" }) => Promise<LocalDirectoryHandle> };
    if (!browser.showDirectoryPicker) { folderInputRef.current?.click(); return; }
    try {
      const handle = await browser.showDirectoryPicker({ mode: "readwrite" });
      setFolderHandle(handle);
      try { await saveLocal("folder", handle); } catch { setNotice("Folder selected for this session. Your browser could not remember it."); }
      await importEntries(await readDirectory(handle, handle.name));
    } catch (caught) {
      if (caught instanceof DOMException && caught.name === "AbortError") return;
      setError("Could not open that folder. You can still choose files or use the folder upload fallback.");
      folderInputRef.current?.click();
    }
  }

  async function reopenFolder() {
    if (!folderHandle) return;
    try {
      const state = await folderHandle.queryPermission({ mode: "readwrite" });
      const permission = state === "granted" ? state : await folderHandle.requestPermission({ mode: "readwrite" });
      if (permission !== "granted") throw new Error("Folder access was not granted.");
      await importEntries(await readDirectory(folderHandle, folderHandle.name));
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Could not reopen the folder.");
    }
  }

  function updateSheet(id: string, patch: Partial<SourceSheet>) {
    setProject((current) => ensureFields({
      ...current,
      sheets: current.sheets.map((sheet) => sheet.id === id ? { ...sheet, ...patch } : sheet),
    }));
  }

  function updateField(id: string, patch: Partial<Field>) {
    setProject((current) => ({
      ...current,
      fields: current.fields.map((field) => field.id === id ? { ...field, ...patch } : field),
    }));
  }

  function addField() {
    setProject((current) => ({
      ...current,
      fields: [...current.fields, { id: crypto.randomUUID(), name: "New field", type: "text", required: false }],
    }));
  }

  function removeField(id: string) {
    setProject((current) => ({
      ...current,
      fields: current.fields.filter((field) => field.id !== id),
      mappings: Object.fromEntries(Object.entries(current.mappings).filter(([, target]) => target !== id)),
      overrides: Object.fromEntries(Object.entries(current.overrides).filter(([, target]) => target !== id)),
      duplicateFields: current.duplicateFields.filter((fieldId) => fieldId !== id),
      rowKeyField: current.rowKeyField === id ? "" : current.rowKeyField,
    }));
  }

  function setColumnMapping(sheet: SourceSheet, column: number, fieldId: string) {
    setProject((current) => ({ ...current, overrides: { ...current.overrides, [columnKey(sheet, column)]: fieldId } }));
  }

  function toggleDuplicateField(id: string) {
    setProject((current) => ({
      ...current,
      duplicateFields: current.duplicateFields.includes(id)
        ? current.duplicateFields.filter((item) => item !== id)
        : [...current.duplicateFields, id],
    }));
  }

  function applyEdits(changes: Record<string, string>) {
    const previous: UndoAction = {};
    Object.entries(changes).forEach(([key, value]) => {
      if (project.edits[key] !== value) previous[key] = project.edits[key];
    });
    if (!Object.keys(previous).length) return;
    setProject((current) => ({ ...current, edits: { ...current.edits, ...changes } }));
    setUndoStack((current) => [...current.slice(-19), previous]);
  }

  function undo() {
    const previous = undoStack.at(-1);
    if (!previous) return;
    setProject((current) => {
      const edits = { ...current.edits };
      Object.entries(previous).forEach(([key, value]) => {
        if (value === undefined) delete edits[key];
        else edits[key] = value;
      });
      return { ...current, edits };
    });
    setUndoStack((current) => current.slice(0, -1));
  }

  function toggleSelected(id: string) {
    setSelectedRows((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  }

  function applyBulkFill() {
    if (!bulkField || !selectedRows.size) return;
    const changes: Record<string, string> = {};
    selectedRows.forEach((id) => { changes[cellEditKey(id, bulkField)] = bulkValue; });
    applyEdits(changes);
    setSelectedRows(new Set());
    setBulkValue("");
    setNotice(`Updated ${Object.keys(changes).length} rows.`);
  }

  const visibleRows = useMemo(() => {
    const needle = query.trim().toLocaleLowerCase();
    return rows.filter((row) => {
      const rowIssues = issueMap.get(row.id) || [];
      if (filter === "issues" && !rowIssues.length) return false;
      if (filter === "missing" && !rowIssues.some((issue) => issue.type === "missing")) return false;
      if (filter === "duplicates" && !rowIssues.some((issue) => issue.type === "duplicate")) return false;
      return !needle || row.source.toLocaleLowerCase().includes(needle) ||
        Object.values(row.values).some((value) => text(value).toLocaleLowerCase().includes(needle));
    });
  }, [rows, issueMap, filter, query]);
  const pageSize = 25;
  const pageCount = Math.max(1, Math.ceil(visibleRows.length / pageSize));
  const currentPage = Math.min(page, pageCount);
  const pageRows = visibleRows.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  async function exportWorkbook(preferFolder: boolean) {
    if (!rows.length) return;
    setBusy(true);
    setError("");
    setProgress("Creating workbook…");
    try {
      const bytes = await workbookBytes(rows, issues, project.fields);
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const name = excelName("sheetwork-clean", "xlsx");
      if (preferFolder && folderHandle) {
        try {
          const state = await folderHandle.queryPermission({ mode: "readwrite" });
          const permission = state === "granted" ? state : await folderHandle.requestPermission({ mode: "readwrite" });
          if (permission !== "granted") throw new Error("Write permission was not granted");
          const file = await folderHandle.getFileHandle(name, { create: true });
          const writable = await file.createWritable();
          await writable.write(blob);
          await writable.close();
          setNotice(`Saved ${name} to ${folderHandle.name}.`);
          return;
        } catch {
          setNotice("Direct folder save was unavailable. The workbook was downloaded instead.");
        }
      }
      downloadBlob(blob, name);
      if (!preferFolder || !folderHandle) setNotice(`Downloaded ${name}.`);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Export failed.");
    } finally {
      setBusy(false);
      setProgress("");
    }
  }

  function saveProjectFile() {
    downloadBlob(new Blob([JSON.stringify(project)], { type: "application/json" }), excelName("sheetwork-project", "json"));
    setNotice("Project file downloaded. It includes imported data, mappings, and edits.");
  }

  async function openProjectFile(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;
    try {
      const candidate: unknown = JSON.parse(await file.text());
      if (!validProject(candidate)) throw new Error("This is not a compatible Sheetwork project file.");
      setProject(candidate);
      setView("sources");
      setSelectedRows(new Set());
      setUndoStack([]);
      setNotice(`Opened project from ${file.name}.`);
      setError("");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Could not open project file.");
    }
  }

  function startFresh() {
    if (!window.confirm("Start a new workspace? Export a project file first if you may need this work again.")) return;
    setProject(emptyProject());
    setSelectedRows(new Set());
    setUndoStack([]);
    setView("sources");
    setNotice("New workspace ready.");
  }

  function renderCell(row: DataRow, field: Field) {
    const key = cellEditKey(row.id, field.id);
    const issue = (issueMap.get(row.id) || []).find((item) => item.fieldId === field.id);
    return (
      <CellEditor
        key={key}
        value={row.values[field.id] || ""}
        changed={Object.hasOwn(project.edits, key)}
        issue={issue}
        label={`${field.name}, ${row.source}, row ${row.excelRow}`}
        onCommit={(value) => applyEdits({ [key]: value })}
      />
    );
  }

  if (!hydrated) return <div className="boot-screen">Opening Sheetwork…</div>;

  return (
    <div className="app-shell">
      <input ref={fileInputRef} className="sr-only" type="file" accept=".xlsx,.xls" multiple onChange={(event) => handleFileInput(event, false)} />
      <input ref={folderInputRef} className="sr-only" type="file" accept=".xlsx,.xls" multiple onChange={(event) => handleFileInput(event, true)} />
      <input ref={projectInputRef} className="sr-only" type="file" accept=".json" onChange={openProjectFile} />

      <header className="app-header">
        <div className="brand-lockup"><span className="brand-symbol" aria-hidden="true"><span /><span /><span /><span /></span><span>Sheetwork<span className="brand-dot">.</span></span></div>
        <div className="header-center">A local workspace for uneven Excel files</div>
        <div className="header-actions">
          <button className="text-button project-open" type="button" onClick={() => projectInputRef.current?.click()}>Open project</button>
          <button className="button button-dark header-export" type="button" disabled={!rows.length || busy} onClick={() => { setView("export"); }}>Export <span aria-hidden="true">↗</span></button>
        </div>
      </header>

      <div className="app-layout">
        <aside className="sidebar">
          <div className="sidebar-intro"><span className="sidebar-kicker">WORKSPACE</span><strong>From scattered sheets<br />to a single view.</strong></div>
          <nav className="step-nav" aria-label="Workflow">
            {(Object.keys(viewMeta) as View[]).map((step) => (
              <button key={step} className={`step-link${view === step ? " is-current" : ""}`} type="button" aria-current={view === step ? "step" : undefined} onClick={() => setView(step)}>
                <span className="step-number">{viewMeta[step].number}</span><span>{step === "sources" ? "Sources" : step === "fields" ? "Fields" : step === "review" ? "Review" : "Export"}</span>
                <span className="step-arrow" aria-hidden="true">↗</span>
              </button>
            ))}
          </nav>
          <div className="sidebar-bottom"><span className="local-indicator" /> Files stay in your browser<br /><small>{saveState}</small></div>
        </aside>

        <main className="main-area">
          <div className="page-heading">
            <div><div className="page-context">{viewMeta[view].number} / {view.toUpperCase()}</div><h1>{viewMeta[view].title}</h1><p>{viewMeta[view].description}</p></div>
            {folderHandle && <button className="folder-pill" type="button" disabled={busy} onClick={reopenFolder} title="Refresh files from the saved folder"><span aria-hidden="true">▣</span> {folderHandle.name} <span aria-hidden="true">↻</span></button>}
          </div>

          {(error || notice || progress) && <div className={`message-bar${error ? " message-error" : ""}`} role={error ? "alert" : "status"}>{error || progress || notice}</div>}
          {busy && <div className="busy-line" aria-hidden="true" />}

          <div className="stats-strip" aria-label="Workspace summary">
            <div><strong>{fileCount}</strong><span>Files</span></div>
            <div><strong>{includedSheets}</strong><span>Sheets</span></div>
            <div><strong>{rows.length.toLocaleString()}</strong><span>Rows</span></div>
            <div><strong>{missingCount.toLocaleString()}</strong><span>Missing</span></div>
            <div><strong>{duplicateCount.toLocaleString()}</strong><span>Duplicate flags</span></div>
          </div>

          {view === "sources" && <>
            <section className="import-panel" aria-labelledby="import-title">
              <div className="import-copy"><span className="section-overline">IMPORT</span><h2 id="import-title">Start with what you have.</h2><p>Add individual Excel files or scan a local folder. The originals are never changed.</p></div>
              <div className="import-actions"><button className="button button-dark" type="button" disabled={busy} onClick={() => fileInputRef.current?.click()}>Add workbooks <span aria-hidden="true">↗</span></button><button className="button button-outline" type="button" disabled={busy} onClick={selectFolder}>Choose folder <span aria-hidden="true">▣</span></button></div>
            </section>
            {fileErrors.length > 0 && <div className="inline-warning"><strong>Some files were skipped</strong><ul>{fileErrors.map((item) => <li key={item}>{item}</li>)}</ul></div>}
            <section className="content-section" aria-labelledby="sheet-list-title">
              <div className="section-heading"><div><h2 id="sheet-list-title">Sheets in this workspace</h2><p>Turn off summary sheets and correct the header row when needed.</p></div><span className="section-count">{project.sheets.length} sheets</span></div>
              {project.sheets.length ? <div className="sheet-list">{project.sheets.map((sheet) => {
                const header = sheet.matrix[sheet.headerRow] || [];
                const dataRows = sheet.matrix.slice(sheet.headerRow + 1).filter((row) => row.some((cell) => text(cell).trim())).length;
                return <article className={`sheet-item${sheet.included ? "" : " is-muted"}`} key={sheet.id}>
                  <label className="sheet-check"><input type="checkbox" checked={sheet.included} onChange={(event) => updateSheet(sheet.id, { included: event.target.checked })} /><span className="sr-only">Include {sheet.name}</span></label>
                  <div className="sheet-main"><div className="sheet-title-row"><strong>{sheet.name}</strong><span>{dataRows.toLocaleString()} rows</span></div><p title={sheet.path}>{sheet.path} <span>· {formatSize(sheet.size)}</span></p><div className="header-preview">{header.slice(0, 7).map((cell, index) => <span key={`${sheet.id}-${index}`}>{text(cell) || `Column ${index + 1}`}</span>)}{header.length > 7 && <span>+{header.length - 7} more</span>}</div></div>
                  <label className="header-row-control">Header row<input type="number" min="1" max={Math.max(1, sheet.matrix.length)} value={sheet.headerRow + 1} onChange={(event) => updateSheet(sheet.id, { headerRow: Math.max(0, Math.min(sheet.matrix.length - 1, Number(event.target.value) - 1 || 0)) })} /></label>
                </article>;
              })}</div> : <div className="empty-state"><span aria-hidden="true">▧</span><strong>No workbooks yet</strong><p>Choose files or a folder to see the sheets you can combine.</p></div>}
            </section>
            <div className="section-actions"><button className="button button-dark" type="button" disabled={!project.sheets.length} onClick={() => setView("fields")}>Continue to fields <span aria-hidden="true">→</span></button></div>
          </>}

          {view === "fields" && <>
            <section className="content-section"><div className="section-heading"><div><h2>Fields in the clean table</h2><p>Rename fields, set a type, and mark values that must be present.</p></div><button className="button button-outline" type="button" onClick={addField}>Add field +</button></div>
              {project.fields.length ? <div className="field-list"><div className="field-list-head"><span>Field name</span><span>Type</span><span>Required</span><span>Duplicate key</span><span /></div>{project.fields.map((field) => <div className="field-row" key={field.id}>
                <label><span className="mobile-label">Field name</span><input className="form-input" aria-label={`Name for ${field.name}`} value={field.name} onChange={(event) => updateField(field.id, { name: event.target.value })} /></label>
                <label><span className="mobile-label">Type</span><select className="form-input" aria-label={`Type for ${field.name}`} value={field.type} onChange={(event) => updateField(field.id, { type: event.target.value as FieldType })}><option value="text">Text</option><option value="number">Number</option><option value="date">Date</option><option value="email">Email</option></select></label>
                <label className="check-cell"><span className="mobile-label">Required</span><input type="checkbox" checked={field.required} onChange={(event) => updateField(field.id, { required: event.target.checked })} /><span className="sr-only">{field.name} required</span></label>
                <label className="check-cell"><span className="mobile-label">Duplicate key</span><input type="checkbox" checked={project.duplicateFields.includes(field.id)} onChange={() => toggleDuplicateField(field.id)} /><span className="sr-only">Use {field.name} to flag duplicates</span></label>
                <button className="icon-button" type="button" aria-label={`Remove ${field.name}`} title="Remove field" onClick={() => removeField(field.id)}>×</button>
              </div>)}</div> : <div className="empty-state"><span aria-hidden="true">◇</span><strong>No fields yet</strong><p>Import workbooks or add a field to define your clean table.</p></div>}
            </section>
            <section className="settings-grid">
              <label className="setting-card"><strong>Stable row ID</strong><span>Used to carry edits forward when a file is imported again.</span><select className="form-input" value={project.rowKeyField} onChange={(event) => setProject((current) => ({ ...current, rowKeyField: event.target.value }))}><option value="">None — new snapshot on reimport</option>{project.fields.map((field) => <option key={field.id} value={field.id}>{field.name}</option>)}</select></label>
              <label className="setting-card"><strong>Missing value markers</strong><span>Empty and whitespace cells are always missing. Add other markers separated by commas.</span><input className="form-input" value={placeholderDraft} onChange={(event) => setPlaceholderDraft(event.target.value)} onBlur={() => setProject((current) => ({ ...current, placeholders: placeholderDraft.split(",").map((item) => item.trim()).filter(Boolean) }))} /></label>
            </section>
            <section className="content-section"><div className="section-heading"><div><h2>Source column mapping</h2><p>Map different header names to one field. Choose Ignore only when a column is unnecessary.</p></div></div>
              {project.sheets.filter((sheet) => sheet.included).length ? <div className="mapping-list">{project.sheets.filter((sheet) => sheet.included).map((sheet) => <details className="mapping-sheet" key={sheet.id}><summary><strong>{sheet.name}</strong><span>{sheet.path}</span><span aria-hidden="true">⌄</span></summary><div className="mapping-rows">{(sheet.matrix[sheet.headerRow] || []).map((header, column) => {
                const mapped = project.overrides[columnKey(sheet, column)] ?? project.mappings[normalizeHeader(header)] ?? "__ignore__";
                return <label className="mapping-row" key={column}><span>{text(header) || `Column ${column + 1}`}</span><span aria-hidden="true">→</span><select className="form-input" value={mapped} onChange={(event) => setColumnMapping(sheet, column, event.target.value)}><option value="__ignore__">Ignore this column</option>{project.fields.map((field) => <option key={field.id} value={field.id}>{field.name}</option>)}</select></label>;
              })}</div></details>)}</div> : <div className="empty-state"><span aria-hidden="true">◇</span><strong>Nothing to map</strong><p>Include a sheet in Sources to see its columns here.</p></div>}
            </section>
            <div className="section-actions"><button className="button button-dark" type="button" disabled={!rows.length} onClick={() => setView("review")}>Continue to review <span aria-hidden="true">→</span></button></div>
          </>}

          {view === "review" && <>
            <section className="content-section review-section"><div className="section-heading"><div><h2>Review imported rows</h2><p>Edits save on this device. Changed cells show a small corner mark.</p></div><button className="button button-outline" type="button" disabled={!undoStack.length} onClick={undo}>Undo last edit ↶</button></div>
              <div className="review-toolbar"><div className="filter-group" role="group" aria-label="Filter rows">{(["all", "issues", "missing", "duplicates"] as ReviewFilter[]).map((item) => <button type="button" className={filter === item ? "is-active" : ""} key={item} onClick={() => { setFilter(item); setPage(1); }}>{item === "all" ? "All rows" : item === "issues" ? "All issues" : item === "missing" ? "Missing" : "Duplicates"}</button>)}</div><label className="search-box"><span className="sr-only">Search rows</span><span aria-hidden="true">⌕</span><input value={query} onChange={(event) => { setQuery(event.target.value); setPage(1); }} placeholder="Search values or files" /></label></div>
              {selectedRows.size > 0 && <div className="bulk-bar"><strong>{selectedRows.size} selected</strong><select className="form-input" aria-label="Field to bulk fill" value={bulkField} onChange={(event) => setBulkField(event.target.value)}><option value="">Choose field</option>{project.fields.map((field) => <option key={field.id} value={field.id}>{field.name}</option>)}</select><input className="form-input" aria-label="Value to apply" value={bulkValue} onChange={(event) => setBulkValue(event.target.value)} placeholder="Value for selected rows" /><button className="button button-dark" type="button" disabled={!bulkField} onClick={applyBulkFill}>Apply value</button><button className="text-button" type="button" onClick={() => setSelectedRows(new Set())}>Clear</button></div>}
              {pageRows.length ? <>
                <div className="table-scroll"><table className="data-table"><thead><tr><th className="select-column"><input type="checkbox" aria-label="Select rows on this page" checked={pageRows.every((row) => selectedRows.has(row.id))} onChange={(event) => setSelectedRows((current) => { const next = new Set(current); pageRows.forEach((row) => event.target.checked ? next.add(row.id) : next.delete(row.id)); return next; })} /></th><th>Status</th><th>Source</th>{project.fields.map((field) => <th key={field.id}>{field.name}{field.required && <span className="required-mark" title="Required"> *</span>}</th>)}</tr></thead><tbody>{pageRows.map((row) => <tr key={row.id}><td><input type="checkbox" aria-label={`Select row ${row.excelRow} in ${row.source}`} checked={selectedRows.has(row.id)} onChange={() => toggleSelected(row.id)} /></td><td><StatusBadge issues={issueMap.get(row.id) || []} /></td><td className="source-cell"><strong>{row.source.split("/").at(-1)}</strong><span>{row.sheet} · row {row.excelRow}</span></td>{project.fields.map((field) => <td key={field.id}>{renderCell(row, field)}</td>)}</tr>)}</tbody></table></div>
                <div className="mobile-records">{pageRows.map((row) => <article className="record-card" key={row.id}><div className="record-top"><label><input type="checkbox" checked={selectedRows.has(row.id)} onChange={() => toggleSelected(row.id)} /><span>{row.source.split("/").at(-1)} · row {row.excelRow}</span></label><StatusBadge issues={issueMap.get(row.id) || []} /></div><details><summary>View and edit fields</summary><div className="record-fields">{project.fields.map((field) => <label key={field.id}><span>{field.name}{field.required ? " *" : ""}</span>{renderCell(row, field)}</label>)}</div></details></article>)}</div>
                <div className="pagination"><span>Showing {(currentPage - 1) * pageSize + 1}–{Math.min(currentPage * pageSize, visibleRows.length)} of {visibleRows.length.toLocaleString()}</span><div><button type="button" disabled={currentPage <= 1} onClick={() => setPage(currentPage - 1)}>← Previous</button><span>{currentPage} / {pageCount}</span><button type="button" disabled={currentPage >= pageCount} onClick={() => setPage(currentPage + 1)}>Next →</button></div></div>
              </> : <div className="empty-state"><span aria-hidden="true">◫</span><strong>{rows.length ? "No rows match this view" : "No rows to review"}</strong><p>{rows.length ? "Try another filter or search." : "Import and include at least one sheet to begin."}</p></div>}
            </section>
            <div className="section-actions"><button className="button button-dark" type="button" disabled={!rows.length} onClick={() => setView("export")}>Continue to export <span aria-hidden="true">→</span></button></div>
          </>}

          {view === "export" && <>
            <section className="export-layout"><div className="export-primary"><span className="section-overline">READY WHEN YOU ARE</span><h2>{rows.length ? `${rows.length.toLocaleString()} rows in one workbook.` : "Add workbooks to begin."}</h2><p>The export has a Data sheet and an Issues sheet. It includes the source file, sheet, and original row number for every record.</p><div className="export-buttons"><button className="button button-dark" type="button" disabled={!rows.length || busy} onClick={() => void exportWorkbook(Boolean(folderHandle))}>{folderHandle ? `Save to ${folderHandle.name}` : "Download workbook"} <span aria-hidden="true">↗</span></button>{folderHandle && <button className="button button-outline" type="button" disabled={!rows.length || busy} onClick={() => void exportWorkbook(false)}>Download instead</button>}</div><div className="export-fine">Original workbooks stay untouched. A new timestamped file is created for every export.</div></div><div className="export-summary"><h3>Before you export</h3><div><span>Included sheets</span><strong>{includedSheets}</strong></div><div><span>Required values missing</span><strong>{missingCount}</strong></div><div><span>Possible duplicate rows</span><strong>{duplicateCount}</strong></div><div><span>Other validation issues</span><strong>{issues.filter((issue) => issue.type === "invalid").length}</strong></div><p>Open issues do not block export. They are listed in the Issues sheet.</p></div></section>
            <section className="content-section"><div className="section-heading"><div><h2>Open issues</h2><p>Review these now or export them for later.</p></div><button className="text-button" type="button" onClick={() => { setFilter("issues"); setPage(1); setView("review"); }}>Review all →</button></div>{issues.length ? <div className="issue-list">{issues.slice(0, 8).map((issue, index) => { const row = rows.find((item) => item.id === issue.rowId); return <div key={`${issue.rowId}-${issue.fieldId}-${index}`}><span className="issue-type">{issue.type}</span><strong>{issue.message}</strong><span>{row?.source.split("/").at(-1)} · {row?.sheet} · row {row?.excelRow}</span></div>; })}{issues.length > 8 && <p className="list-more">+{issues.length - 8} more issues in the export</p>}</div> : <div className="all-clear">No open issues in the current table.</div>}</section>
            <section className="project-tools"><div><h2>Keep your workspace</h2><p>Project files contain imported data, mappings, and edits. Use one to continue on another computer.</p></div><div><button className="button button-outline" type="button" onClick={saveProjectFile}>Save project file</button><button className="text-button" type="button" onClick={() => projectInputRef.current?.click()}>Open project file</button><button className="text-button danger-link" type="button" onClick={startFresh}>Start fresh</button></div></section>
          </>}
        </main>
      </div>

      <footer className="app-footer"><span>Sheetwork · Local Excel workspace</span><span>Files are processed in your browser. Folder access is always your choice.</span></footer>
    </div>
  );
}
