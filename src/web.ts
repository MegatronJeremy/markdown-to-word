import { exportToDocx } from "./exporter";

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const input = $<HTMLTextAreaElement>("md");
const status = $<HTMLElement>("status");
const nameEl = $<HTMLInputElement>("name");

async function convert() {
  const md = input.value;
  if (!md.trim()) {
    status.textContent = "Paste or drop some Markdown first.";
    return;
  }
  const base = (nameEl.value.trim() || "document").replace(/\.docx$/i, "").replace(/[^\w .-]+/g, "_");
  try {
    const bytes = await exportToDocx(md, { title: base });
    const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = base + ".docx";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    status.textContent = `Done: ${base}.docx (${Math.round(bytes.length / 1024)} KB).`;
  } catch (e) {
    status.textContent = "Conversion failed: " + (e instanceof Error ? e.message : String(e));
  }
}

async function loadFile(f: File) {
  input.value = await f.text();
  nameEl.value = f.name.replace(/\.(md|markdown|txt)$/i, "");
  status.textContent = `Loaded ${f.name} locally.`;
}

$("convert").addEventListener("click", convert);
$<HTMLInputElement>("file").addEventListener("change", (e) => {
  const f = (e.target as HTMLInputElement).files?.[0];
  if (f) void loadFile(f);
});
const drop = $("drop");
for (const ev of ["dragenter", "dragover"]) drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); });
for (const ev of ["dragleave", "drop"]) drop.addEventListener(ev, () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  const f = (e as DragEvent).dataTransfer?.files[0];
  if (f) void loadFile(f);
});
