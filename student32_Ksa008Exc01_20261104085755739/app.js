/* global JSZip, XLSX */
const shifts = [
  { code: "A", label: "A勤務", start: "09:00", end: "10:40" },
  { code: "B", label: "B勤務", start: "10:55", end: "12:35" },
  { code: "C", label: "C勤務", start: "13:20", end: "15:00" },
  { code: "D", label: "D勤務", start: "15:15", end: "16:55" },
  { code: "E", label: "E勤務", start: "17:10", end: "18:50" }
];
const state = { reservations: [], rooms: new Map(), results: [], allResults: [] };
const $ = (id) => document.getElementById(id);
const dateInput = $("dateInput");
dateInput.value = new Date().toISOString().slice(0, 10);

function cellValue(row, names) {
  const key = Object.keys(row).find((candidate) => names.some((name) => candidate.replace(/\s/g, "").includes(name)));
  return key ? String(row[key] ?? "").trim() : "";
}
function timeToMinutes(value) {
  const digits = String(value).replace(/[^0-9]/g, "").padStart(4, "0");
  return Number(digits.slice(0, 2)) * 60 + Number(digits.slice(2, 4));
}
function normalizedDate(value) {
  const text = String(value ?? "").trim().replace(/[年月]/g, "/").replace(/日/g, "");
  const match = text.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  return match ? `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}` : text;
}
async function readZip(file) {
  const zip = await JSZip.loadAsync(file);
  const rows = [];
  for (const entry of Object.values(zip.files)) {
    if (entry.dir || !/\.(xlsx|xls|csv)$/i.test(entry.name)) continue;
    const data = await entry.async("arraybuffer");
    if (/\.csv$/i.test(entry.name)) {
      rows.push(...XLSX.utils.sheet_to_json(XLSX.read(data, { type: "array", raw: false }), { defval: "" }));
      continue;
    }
    const book = XLSX.read(data, { type: "array", cellDates: true, raw: false });
    book.SheetNames.forEach((name) => rows.push(...XLSX.utils.sheet_to_json(book.Sheets[name], { defval: "" })));
  }
  return rows;
}
function parseRows(rows) {
  state.reservations = [];
  state.rooms = new Map();
  rows.forEach((row) => {
    const code = cellValue(row, ["教室コード", "教室ｺｰﾄﾞ", "room code"]);
    const name = cellValue(row, ["教室名", "room name"]) || code;
    const date = normalizedDate(cellValue(row, ["利用日", "日付", "date"]));
    const start = cellValue(row, ["開始時刻", "開始時間", "start"]);
    const end = cellValue(row, ["終了時刻", "終了時間", "end"]);
    if (!code && !name) return;
    if (!state.rooms.has(code || name)) state.rooms.set(code || name, { code, name, group: cellValue(row, ["教室グループ名", "グループ名", "group"]) || "-" });
    if (date && start && end) state.reservations.push({ code: code || name, date, start: timeToMinutes(start), end: timeToMinutes(end) });
  });
  return state.rooms.size;
}
function escapeHtml(value) { return value.replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char])); }
function availableRooms(date, startValue, endValue) {
  const start = timeToMinutes(startValue);
  const end = timeToMinutes(endValue);
  return [...state.rooms.values()]
    .filter((room) => !state.reservations.some((reservation) => reservation.code === (room.code || room.name) && reservation.date === date && reservation.start < end && reservation.end > start))
    .sort((left, right) => (left.code || left.name).localeCompare(right.code || right.name, "ja", { numeric: true }));
}
function renderResults() {
  const date = dateInput.value;
  const shift = shifts.find((item) => item.code === $("shiftSelect").value) || shifts[0];
  const startValue = $("startInput").value;
  const endValue = $("endInput").value;
  const rooms = availableRooms(date, startValue, endValue);
  state.results = rooms;
  state.allResults = shifts.map((item) => ({ ...item, rooms: availableRooms(date, item.start, item.end) }));
  $("results").hidden = false;
  $("resultCount").textContent = `${rooms.length} 教室`;
  $("resultMeta").textContent = `${date.replaceAll("-", "/")}  ${shift.label}（${startValue} - ${endValue}）の空き状況`;
  $("resultBody").innerHTML = rooms.map((room) => `<tr><td>${escapeHtml(room.code || "-")}</td><td><strong>${escapeHtml(room.name || "-")}</strong></td><td>${escapeHtml(room.group)}</td><td>${date.replaceAll("-", "/")}</td><td>${startValue} - ${endValue}</td><td class="available">空き</td></tr>`).join("");
  $("emptyState").hidden = rooms.length > 0;
  $("downloadButton").disabled = rooms.length === 0;
  $("pdfButton").disabled = rooms.length === 0;
  $("allCsvButton").disabled = state.rooms.size === 0;
  $("allPdfButton").disabled = state.rooms.size === 0;
}
function csvDownload(rows, filename) {
  const csv = "\uFEFF" + rows.map((row) => row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(",")).join("\n");
  const link = document.createElement("a");
  link.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
  link.download = filename;
  link.click();
  URL.revokeObjectURL(link.href);
}
$("zipFile").addEventListener("change", async (event) => {
  const file = event.target.files[0]; if (!file) return;
  $("fileName").textContent = file.name; $("dataStatus").textContent = "データを読み込んでいます…";
  try { const count = parseRows(await readZip(file)); $("dataStatus").textContent = `${count} 教室のデータを読み込みました。`; $("dataStatus").classList.add("ready"); $("searchButton").disabled = false; }
  catch (error) { $("dataStatus").textContent = "ZIPを読み込めませんでした。Excelファイルを確認してください。"; $("errorMessage").textContent = error.message; }
});
$("dropzone").addEventListener("dragover", (event) => { event.preventDefault(); $("dropzone").classList.add("is-dragging"); });
$("dropzone").addEventListener("dragleave", () => $("dropzone").classList.remove("is-dragging"));
$("dropzone").addEventListener("drop", (event) => { event.preventDefault(); $("dropzone").classList.remove("is-dragging"); const fileInput = $("zipFile"); fileInput.files = event.dataTransfer.files; fileInput.dispatchEvent(new Event("change")); });
$("searchButton").addEventListener("click", () => { const start = timeToMinutes($("startInput").value); const end = timeToMinutes($("endInput").value); $("errorMessage").textContent = start >= end ? "終了時刻は開始時刻より後にしてください。" : ""; if (start < end) renderResults(); });
$("shiftSelect").addEventListener("change", (event) => { const option = event.target.selectedOptions[0]; $("startInput").value = option.dataset.start; $("endInput").value = option.dataset.end; });
$("downloadButton").addEventListener("click", () => { const rows = [["教室コード", "教室名", "教室グループ", "利用日", "指定開始", "指定終了", "状態"], ...state.results.map((room) => [room.code, room.name, room.group, dateInput.value, $("startInput").value, $("endInput").value, "空き"])]; csvDownload(rows, `空き教室_${dateInput.value}_${$("shiftSelect").value}勤務.csv`); });
function openPrintPage(allShifts) {
  const date = dateInput.value;
  let title;
  let headers;
  let rows;
  if (allShifts) {
    title = "空き教室一覧（全勤務）";
    headers = ["教室コード", "教室名", "グループ", ...shifts.map((item) => `${item.label} ${item.start}-${item.end}`)];
    rows = [...state.rooms.values()].sort((left, right) => (left.code || left.name).localeCompare(right.code || right.name, "ja", { numeric: true })).map((room) => [room.code || "-", room.name || "-", room.group, ...state.allResults.map((item) => item.rooms.some((result) => (result.code || result.name) === (room.code || room.name)) ? "空き" : "予約あり")]);
  } else {
    const shift = shifts.find((item) => item.code === $("shiftSelect").value) || shifts[0];
    title = `空き教室一覧（${shift.label}）`;
    headers = ["教室コード", "教室名", "グループ", "利用日", "指定時間", "状態"];
    rows = state.results.map((room) => [room.code || "-", room.name || "-", room.group, date.replaceAll("-", "/"), `${$("startInput").value} - ${$("endInput").value}`, "空き"]);
  }
  localStorage.setItem("roomscout-print-data", JSON.stringify({ title, date, headers, rows, allShifts }));
  window.open("print.html", "_blank");
}
$("pdfButton").addEventListener("click", () => openPrintPage(false));
$("allCsvButton").addEventListener("click", () => { const rows = [["教室コード", "教室名", "グループ", ...shifts.map((item) => `${item.label}（${item.start}-${item.end}）`)]]; const roomMap = new Map(state.rooms); [...roomMap.values()].sort((left, right) => (left.code || left.name).localeCompare(right.code || right.name, "ja", { numeric: true })).forEach((room) => rows.push([room.code || "-", room.name || "-", room.group, ...state.allResults.map((item) => item.rooms.some((result) => (result.code || result.name) === (room.code || room.name)) ? "空き" : "予約あり")])); csvDownload(rows, `空き教室_全勤務_${dateInput.value}.csv`); });
$("allPdfButton").addEventListener("click", () => openPrintPage(true));