let parsedData = null;

document.getElementById("csvFile").addEventListener("change", handleFile);
document.getElementById("downloadBtn").addEventListener("click", exportToExcel);

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  // Show file name
  document.getElementById("fileName").textContent = `Selected file: ${file.name}`;

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: function(results) {
      parsedData = results.data;
      document.getElementById("downloadBtn").disabled = false; // enable download
      alert("✅ File uploaded successfully. Now click 'Download Processed File'.");
    }
  });
}

function exportToExcel() {
  if (!parsedData) {
    alert("⚠️ No file uploaded yet.");
    return;
  }

  // Keep headers but remove "Reference"
  let headers = Object.keys(parsedData[0]).filter(h => h !== "Reference");

  const rows = parsedData.map(row => {
    let description = row["Description"] || "";
    let match = description.match(/\b\d{6}\b/);
    let reference = match ? match[0] : "999999";

    // Put reference under Matter no
    let matterNo = reference;
    return headers.map(h => h === "Matter no" ? matterNo : row[h] || "");
  });

  const worksheetData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(worksheetData);

  // Apply borders + red fill for rows containing 999999 in Matter no
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const matterIndex = headers.indexOf("Matter no");

  for (let R = 1; R <= range.e.r; ++R) {
    let matterCell = ws[XLSX.utils.encode_cell({r: R, c: matterIndex})];
    let isRed = matterCell && matterCell.v === "999999";

    for (let C = 0; C <= range.e.c; ++C) {
      let cellAddress = XLSX.utils.encode_cell({r: R, c: C});
      if (!ws[cellAddress]) ws[cellAddress] = { t: "s", v: "" };

      ws[cellAddress].s = {
        fill: isRed ? { fgColor: { rgb: "FF0000" } } : undefined, // 🔴 red background
        border: {
          top: { style: "thin", color: { rgb: "000000" } },
          bottom: { style: "thin", color: { rgb: "000000" } },
          left: { style: "thin", color: { rgb: "000000" } },
          right: { style: "thin", color: { rgb: "000000" } }
        }
      };
    }
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Transactions");

  // Trigger download
  XLSX.writeFile(wb, "processed_transactions.xlsx");

  // Refresh page after download
  setTimeout(() => location.reload(), 1000);
}
