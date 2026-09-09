"use client";

import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type MatrixRow = Record<string, string> & { id: string };

const DEFAULT_COLUMNS = [
  "Nama",
  "Jenis Kegiatan",
  "Sub Kegiatan",
  "Status Kegiatan",
  "Kasubid",
  "Kabid",
  "Kasubid/Kabid Bidang lain",
  "Sekban",
  "Kaban",
  "Keterangan Tambahan",
  "Status",
  "Tanggal Selesai",
  "Perda/Perbup/Kepbup/Perkaban/KepKaban dll",
  "Catatan",
  "LINK",
];

const STATUS_OPTIONS = ["Direncanakan", "Berjalan", "Selesai", "Tertunda"];

function statusClass(status: string) {
  return `status ${status.toLowerCase().replace(/\s+/g, "-")}`;
}

function createEmptyRow(columns: string[]): MatrixRow {
  return columns.reduce<MatrixRow>((result, column) => ({ ...result, [column]: "" }), {
    id: crypto.randomUUID(),
  });
}

function detectHeaderRow(rows: unknown[][]) {
  return rows.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 3);
}

function activityPayload(row: MatrixRow) {
  const { id, ...data } = row;
  return { id, data };
}

function columnsFromRows(data: MatrixRow[]) {
  const available = new Set(data.flatMap((row) => Object.keys(row).filter((key) => key !== "id")));
  return [...DEFAULT_COLUMNS.filter((column) => available.has(column)), ...Array.from(available).filter((column) => !DEFAULT_COLUMNS.includes(column))];
}

export default function DashboardPage() {
  const [columns, setColumns] = useState(DEFAULT_COLUMNS);
  const [rows, setRows] = useState<MatrixRow[]>([]);
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("Semua status");
  const [isDialogOpen, setDialogOpen] = useState(false);
  const [editingRow, setEditingRow] = useState<MatrixRow | null>(null);
  const [draft, setDraft] = useState<MatrixRow>(() => createEmptyRow(DEFAULT_COLUMNS));
  const [notice, setNotice] = useState("");
  const [nameSort, setNameSort] = useState<"asc" | "desc">("asc");
  const fileInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    async function loadActivities() {
      try {
        const response = await fetch("/api/activities");
        if (!response.ok) throw new Error("Tidak dapat mengambil data.");
        const activities = (await response.json()) as { id: string; data: Record<string, string> }[];
        const remoteRows = activities.map((activity) => ({ id: activity.id, ...activity.data }));
        setRows(remoteRows);
        setColumns(columnsFromRows(remoteRows));
      } catch {
        setNotice("Database belum siap. Jalankan npm run db:setup, lalu muat ulang halaman.");
      }
    }
    void loadActivities();
  }, []);

  const filteredRows = useMemo(() => {
    const normalizedQuery = query.toLowerCase().trim();
    return rows.filter((row) => {
      const matchesQuery = !normalizedQuery || columns.some((column) => row[column]?.toLowerCase().includes(normalizedQuery));
      const matchesStatus = statusFilter === "Semua status" || row.Status === statusFilter;
      return matchesQuery && matchesStatus;
    }).sort((left, right) => (left.Nama || "").localeCompare(right.Nama || "", "id", { sensitivity: "base" }) * (nameSort === "asc" ? 1 : -1));
  }, [columns, nameSort, query, rows, statusFilter]);

  const usesMatriksLayout = columns.length === DEFAULT_COLUMNS.length && DEFAULT_COLUMNS.every((column, index) => columns[index] === column);

  function openNewDialog() {
    setEditingRow(null);
    setDraft(createEmptyRow(columns));
    setDialogOpen(true);
  }

  function openEditDialog(row: MatrixRow) {
    setEditingRow(row);
    setDraft({ ...row });
    setDialogOpen(true);
  }

  async function saveRow(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    try {
      const response = await fetch("/api/activities", {
        method: editingRow ? "PATCH" : "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(activityPayload(draft)),
      });
      if (!response.ok) throw new Error("Gagal menyimpan.");
      if (editingRow) setRows((current) => current.map((row) => (row.id === editingRow.id ? draft : row)));
      else setRows((current) => [draft, ...current]);
      setNotice(editingRow ? "Data kegiatan berhasil diperbarui." : "Data kegiatan berhasil ditambahkan.");
      setDialogOpen(false);
    } catch {
      setNotice("Data tidak tersimpan. Periksa koneksi database.");
    }
  }

  async function deleteRow(id: string) {
    if (!window.confirm("Hapus data kegiatan ini?")) return;
    try {
      const response = await fetch("/api/activities", { method: "DELETE", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id }) });
      if (!response.ok) throw new Error("Gagal menghapus.");
      setRows((current) => current.filter((row) => row.id !== id));
      setNotice("Data kegiatan telah dihapus.");
    } catch {
      setNotice("Data tidak dapat dihapus. Periksa koneksi database.");
    }
  }

  async function replaceRows(nextColumns: string[], nextRows: MatrixRow[], successMessage: string) {
    try {
      const response = await fetch("/api/activities", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ activities: nextRows.map(activityPayload) }) });
      if (!response.ok) throw new Error("Gagal mengganti data.");
      setColumns(nextColumns);
      setRows(nextRows);
      setStatusFilter("Semua status");
      setNotice(successMessage);
    } catch {
      setNotice("Data tidak tersimpan. Periksa koneksi database.");
    }
  }

  function handleImport(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (loadEvent) => {
      const workbook = XLSX.read(loadEvent.target?.result, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const values = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: "" });
      const headerIndex = detectHeaderRow(values);
      if (headerIndex < 0) {
        setNotice("Header tabel tidak ditemukan. Pastikan ada minimal tiga kolom pada Excel.");
        return;
      }
      const importedColumns = values[headerIndex]
        .map((value, index) => String(value).trim() || `Kolom ${index + 1}`)
        .filter((value, index, array) => value && array.indexOf(value) === index);
      const importedRows = values
        .slice(headerIndex + 1)
        .filter((row) => row.some((value) => String(value).trim()))
        .map((row) =>
          importedColumns.reduce<MatrixRow>((record, column, index) => ({
            ...record,
            [column]: String(row[index] ?? ""),
          }), { id: crypto.randomUUID() }),
        );
      await replaceRows(importedColumns, importedRows, `${importedRows.length} baris dari Excel berhasil diimpor.`);
    };
    reader.readAsArrayBuffer(file);
    event.target.value = "";
  }

  function exportExcel() {
    const worksheet = XLSX.utils.json_to_sheet(rows.map(({ id, ...row }) => row), { header: columns });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Matriks Kegiatan");
    XLSX.writeFile(workbook, "matriks-kegiatan.xlsx");
  }

  return (
    <main className="shell">
      <section className="content" id="dashboard">
        <header className="topbar">
          <div>
            <p className="eyebrow">MANAJEMEN PROGRAM</p>
            <h1>Matriks Pelaksanaan Kegiatan</h1>
            <p className="subtitle">Kelola rencana, pelaksanaan, dan tindak lanjut kegiatan dalam satu tempat.</p>
          </div>
          <button className="primary" onClick={openNewDialog}>+ Tambah kegiatan</button>
        </header>

        <section className="table-card" id="data-kegiatan">
          <div className="table-header">
            <div><h2>Data kegiatan</h2><p>{filteredRows.length} dari {rows.length} data ditampilkan</p></div>
            <div className="actions">
              <input ref={fileInput} className="sr-only" type="file" accept=".xlsx,.xls" onChange={handleImport} />
              <button className="secondary" onClick={() => fileInput.current?.click()}>⇧ Impor Excel</button>
              <button className="secondary" onClick={exportExcel}>⇩ Ekspor</button>
            </div>
          </div>

          {notice && <div className="notice" role="status">{notice}<button onClick={() => setNotice("")}>×</button></div>}

          <div className="filters">
            <label className="search"><span>⌕</span><input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Cari kegiatan, target, atau penanggung jawab..." /></label>
            <select value={statusFilter} onChange={(event) => setStatusFilter(event.target.value)}>
              <option>Semua status</option>
              {STATUS_OPTIONS.map((status) => <option key={status}>{status}</option>)}
            </select>
            <select value={nameSort} onChange={(event) => setNameSort(event.target.value as "asc" | "desc")} aria-label="Urutkan berdasarkan nama">
              <option value="asc">Nama: A–Z</option>
              <option value="desc">Nama: Z–A</option>
            </select>
          </div>

          <div className="table-scroll">
            <table>
              <thead>
                {usesMatriksLayout ? <>
                  <tr className="group-header">
                    <th rowSpan={2}>No.</th><th rowSpan={2}>Nama</th>
                    <th colSpan={3}>Jenis Kegiatan</th>
                    <th colSpan={6}>Keterangan / Progress</th>
                    <th rowSpan={2}>Status</th><th rowSpan={2}>Tanggal Selesai</th>
                    <th rowSpan={2}>Perda/Perbup/Kepbup/Perkaban/KepKaban dll</th><th rowSpan={2}>Catatan</th><th rowSpan={2}>Link</th>
                    <th className="sticky-action" rowSpan={2}>Aksi</th>
                  </tr>
                  <tr><th>Jenis Kegiatan</th><th>Sub Kegiatan</th><th>Status Kegiatan</th><th>Kasubid</th><th>Kabid</th><th>Kasubid/Kabid Bidang lain</th><th>Sekban</th><th>Kaban</th><th>Keterangan Tambahan</th></tr>
                </> : <tr><th>No.</th>{columns.map((column) => <th key={column}>{column}</th>)}<th className="sticky-action">Aksi</th></tr>}
              </thead>
              <tbody>
                {filteredRows.length ? filteredRows.map((row, index) => (
                  <tr key={row.id}>
                    <td className="row-number">{index + 1}</td>
                    {columns.map((column) => <td key={column}>{column === "Status" || column === "Status Kegiatan" ? <span className={statusClass(row[column] || "-")}>{row[column] || "-"}</span> : column === "Sub Kegiatan" ? (row[column]?.split("\n").filter(Boolean).length ? <ol className="point-list-view">{row[column].split("\n").filter(Boolean).map((point, i) => <li key={i}>{point}</li>)}</ol> : <span className="empty">—</span>) : column === "LINK" ? (row[column]?.split("\n").filter(Boolean).length ? <ol className="point-list-view">{row[column].split("\n").filter(Boolean).map((url, i) => <li key={i}><a className="link" href={url} target="_blank" rel="noreferrer">Buka link {row[column].split("\n").filter(Boolean).length > 1 ? i + 1 : ""}</a></li>)}</ol> : <span className="empty">—</span>) : row[column] || <span className="empty">—</span>}</td>)}
                    <td className="sticky-action"><button className="icon-button" aria-label="Ubah data" title="Ubah" onClick={() => openEditDialog(row)}>✎</button><button className="icon-button danger" aria-label="Hapus data" title="Hapus" onClick={() => deleteRow(row.id)}>⌫</button></td>
                  </tr>
                )) : <tr><td className="no-data" colSpan={columns.length + 2}>Tidak ada data yang sesuai. Tambahkan kegiatan atau ubah pencarian.</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
      </section>

      {isDialogOpen && (
        <div className="backdrop" role="presentation" onMouseDown={() => setDialogOpen(false)}>
          <form className="dialog" onSubmit={saveRow} onMouseDown={(event) => event.stopPropagation()}>
            <div className="dialog-header"><div><p className="eyebrow">DATA KEGIATAN</p><h2>{editingRow ? "Ubah kegiatan" : "Tambah kegiatan"}</h2></div><button type="button" className="close" onClick={() => setDialogOpen(false)}>×</button></div>
            <div className="form-grid">
              {columns.map((column) => (
                <label key={column} className={column === "Keterangan" || column === "Subkegiatan" ? "wide" : ""}>
                  <span>{column}</span>
                  {column === "Status" || column === "Status Kegiatan" ? (
                    <select value={draft[column] || ""} onChange={(event) => setDraft({ ...draft, [column]: event.target.value })}>
                      <option value="">Pilih status</option>{STATUS_OPTIONS.map((status) => <option key={status}>{status}</option>)}
                    </select>
                  ) : column === "Sub Kegiatan" || column === "LINK" ? (
                    <div className="point-list">
                      {(draft[column] ? draft[column].split("\n") : [""]).map((point, index, points) => (
                        <div className="point-row" key={index}>
                          <span className="point-index">{index + 1}.</span>
                          <input
                            type={column === "LINK" ? "url" : "text"}
                            value={point}
                            onChange={(event) => {
                              const next = [...points];
                              next[index] = event.target.value;
                              setDraft({ ...draft, [column]: next.join("\n") });
                            }}
                            placeholder={column === "LINK" ? `Link ${index + 1} (https://...)` : `Poin sub kegiatan ${index + 1}`}
                          />
                          {points.length > 1 && (
                            <button type="button" className="icon-button danger" title="Hapus poin" onClick={() => setDraft({ ...draft, [column]: points.filter((_, i) => i !== index).join("\n") })}>⌫</button>
                          )}
                        </div>
                      ))}
                      <button type="button" className="text-button" onClick={() => setDraft({ ...draft, [column]: [...(draft[column] ? draft[column].split("\n") : [""]), ""].join("\n") })}>+ Tambah {column === "LINK" ? "link" : "poin"}</button>
                    </div>
                  ) : ["Jenis Kegiatan", "Keterangan Tambahan", "Perda/Perbup/Kepbup/Perkaban/KepKaban dll", "Catatan"].includes(column) ? (
                    <textarea value={draft[column] || ""} onChange={(event) => setDraft({ ...draft, [column]: event.target.value })} placeholder={`Masukkan ${column.toLowerCase()}`} />
                  ) : (
                    <input value={draft[column] || ""} onChange={(event) => setDraft({ ...draft, [column]: event.target.value })} placeholder={`Masukkan ${column.toLowerCase()}`} />
                  )}
                </label>
              ))}
            </div>
            <div className="dialog-footer"><button type="button" className="secondary" onClick={() => setDialogOpen(false)}>Batal</button><button className="primary" type="submit">{editingRow ? "Simpan perubahan" : "Simpan kegiatan"}</button></div>
          </form>
        </div>
      )}
    </main>
  );
}