const data = JSON.parse(localStorage.getItem("roomscout-print-data") || "null");
const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));

if (data) {
  document.title = `${data.title} | ROOMSCOUT`;
  document.getElementById("printTitle").textContent = data.title;
  document.getElementById("printDate").textContent = String(data.date || "").replaceAll("-", "/");
  document.getElementById("printHead").innerHTML = `<tr>${data.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr>`;
  document.getElementById("printBody").innerHTML = data.rows.map((row) => `<tr>${row.map((cell, index) => {
    const isShiftCell = data.allShifts && index >= 3;
    const isAvailable = cell === "空き";
    const value = isShiftCell ? (isAvailable ? "○" : "×") : cell;
    const className = isShiftCell ? (isAvailable ? "availability available" : "availability reserved") : "";
    return `<td class="${className}">${escapeHtml(value)}</td>`;
  }).join("")}</tr>`).join("");
} else {
  document.getElementById("printTitle").textContent = "印刷するデータがありません";
  document.getElementById("printDate").textContent = "検索画面から結果を作成してください。";
}

document.getElementById("printButton").addEventListener("click", () => window.print());
document.getElementById("backButton").addEventListener("click", () => window.close());