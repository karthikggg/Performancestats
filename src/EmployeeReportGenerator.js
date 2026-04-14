import React, { useState, useRef } from "react";
import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import "bootstrap/dist/css/bootstrap.min.css";
 
const INITIAL_EMPLOYEE_DATA = [
  { id: 10411774, name: "Varsha Mahadevan" },
  { id: 10412819, name: "Saquib Tanweer" },
  { id: 10412814, name: "Jeff Rohit" },
  { id: 10412996, name: "Arjun Thirumalaikumar" },
  { id: 10415412, name: "Veera Sabarinathan" },
  { id: 10416036, name: "Bala Thirupathi Raaja" },
  { id: 10417367, name: "Shalini Subramanian" },
  { id: 10418289, name: "Sivasankari Arumugam" },
  { id: 10418626, name: "Prabhakaran Sekar" },
  { id: 10419645, name: "Mohammed Umar Mansoor" },
  { id: 10420136, name: "Sindhuja Prabakaran" },
  { id: 10420137, name: "Siddharthan Mayilsamy" },
  { id: 10420134, name: "Harishwar Nesamani" },
  { id: 10421099, name: "Anshuman Dey" },
  { id: 10421101, name: "Ajit Balaji" },
  { id: 10421103, name: "Kumaran Ramachandran" },
  { id: 10421093, name: "Sabariraj Iyyappan" },
  { id: 10421094, name: "Keerthana Ganesh" },
  { id: 10421096, name: "Sabarish Gupta Obilisetti" },
  { id: 10421097, name: "Rajamouli Ramaiyan" },
  { id: 10421667, name: "Thaanu Kumar M" },
  { id: 10422354, name: "Akash.M Murali" },
  { id: 10422783, name: "Yazhini Krishnamoorthy" },
  { id: 10422784, name: "Sabana Satik" },
  { id: 10423034, name: "Allan Augustine" },
  { id: 10423370, name: "Ganesa Murugan" },
  { id: 10425117, name: "Akshay Gopakumar" },
  { id: 10425125, name: "Sachin Rajesh" },
  { id: 10425121, name: "Sathish Kumar Sankaranagappan" },
  { id: 10425123, name: "Manju Kolli" },
  { id: 10425115, name: "Parthasarathy Letchumanan" },
  { id: 10425124, name: "Siva Ganesh Santhanam" },
  { id: 10425415, name: "Jenithson Thommai" },
  { id: 10425416, name: "Saranesh Duraisamy" },
  { id: 10426079, name: "Rex Fleming" },
  { id: 10426890, name: "GN Karthik" },
  { id: 10428929, name: "Mohammed Wihaj" },
  { id: 10429597, name: "Asmitha Gnanaprakash" },
  { id: 10428930, name: "Kaviprabha G" },
  { id: 10428931, name: "Gurumoorthy Vijayarangan" },
  { id: 10428932, name: "Rekha B" },
  { id: 10428935, name: "Krishna Chaitanya" },
  { id: 10428934, name: "Hariharan N" },
  { id: 10429979, name: "Karthikeyan Shankar" },
  { id: 10430834, name: "Aishwarya Rajamohan" },
  { id: 10430832, name: "Vinod Ram" },
  { id: 10430830, name: "Deepika Raghuraj" },
  { id: 10431147, name: "Divya Dharshini" },
  { id: 10431142, name: "Melwin Manoj" },
  { id: 10431141, name: "Anusree Anil" },
  { id: 10433153, name: "Ramprakash Rajan" },
  { id: 10433152, name: "Ayyapparaj Dhamodhaan" },
  { id: 10433154, name: "Harshaavardhan Subramani" },
  { id: 10432953, name: "Pooja Raghavendra" },
  { id: 10433441, name: "Bhuvan Balasubramanian" },
  { id: 10445740, name: "Swetha Mani" },
  { id: 10446572, name: "Vijay Kumar R" },
  { id: 10446964, name: "Arul Mani Joseph" },
  { id: 10446962, name: "Manoj Rajasekaran" },
  { id: 10446965, name: "Ali Mehran Kandrikar" },
  { id: 10446967, name: "Naveen Srinivasan" },
  { id: 10446966, name: "Kishore Ganesan" },
  { id: 10447158, name: "Prasanth Rajendran" },
  { id: 10447160, name: "Avi Sharma" },
  { id: 10447662, name: "Veeravisvavinayagam Kumaravelu" },
  { id: 10447398, name: "Epsi Surendran" },
  { id: 10447277, name: "Saran Kumar G" },
  { id: 10447157, name: "Karthik Govindasamy" },
  { id: 10447281, name: "Tharun Kumar V" },
  { id: 10447280, name: "Nitish Kumar" },
  { id: 10447397, name: "Divya Barani Karthikeyean" },
  { id: 10447156, name: "Vishwa Alagiri" },
  { id: 10447155, name: "Shantha Kumar Saravanan" },
  { id: 10447163, name: "Meenakshi Maragathavel" },
  { id: 10447396, name: "Durairaj Saravanakumar" },
  { id: 10447663, name: "Sai Kumar C" },
  { id: 10447166, name: "Priyadharshini Mohan" },
  { id: 10447162, name: "Vishnu Bose" },
  { id: 10447273, name: "Lakshmi Aishwarya Ratakondala" },
  { id: 10447276, name: "Shanmuga Priya. Ramesh" },
  { id: 10447165, name: "Priyea Dharshani B" },
  { id: 10447164, name: "Yuvaraj Selvam" },
  { id: 10447275, name: "Ashwin Kumar S" },
  { id: 10447167, name: "Janani Venkatesalu" },
  { id: 10447161, name: "Jayasree Mohanakrishnan" },
  { id: 10447334, name: "Kiranraj Ravichandran" },
  { id: 10447335, name: "Priyadharshini James" },
  { id: 10447336, name: "Moneshwar Devaraj" },
  { id: 10447337, name: "Shifhana Banu Usain" },
  { id: 10447338, name: "Goutham Sakthivel" },
  { id: 10447279, name: "Dilip Suresh" },
  { id: 10447665, name: "Kishore Sivalingam" },
  { id: 10448179, name: "Dhuruva Gowshik Ganesan" },
  { id: 10429329, name: "Aarthi Madhan" },
  { id: 10429332, name: "Ajay Dhandapani" },
  { id: 10429336, name: "Akash Sampath" },
  { id: 10429339, name: "Arun Sajeev" },
  { id: 10429340, name: "Aswini Haribabu" },
  { id: 10429343, name: "Augustina Albert Sagayaraj" },
  { id: 10429346, name: "Deepika Subramani" },
  { id: 10429349, name: "Dhanalakshmi Sundar" },
  { id: 10429354, name: "Harihara Ponnaiah" },
  { id: 10429259, name: "Kamaleeshwari Sasi Kapoor Singh" },
  { id: 10429360, name: "Nithish Thivya" },
  { id: 10429367, name: "Praveen Kumar Thanigaiarasu" },
  { id: 10429361, name: "Rajeshwari Rajagopal" },
  { id: 10429368, name: "Yuvasree Balasubramaniam" },
  { id: 10429384, name: "Saran T" },
  { id: 10448387, name: "Karthick Gurunathan" },
  { id: 10448384, name: "Nishanthini Umapathy" },
  { id: 10448382, name: "Rohit Subramani" },
  { id: 10448381, name: "Samyuktha Balakrishnaian" },
  { id: 10448390, name: "Vijayalakshmi Dhanabalan" },
  { id: 10448377, name: "Vignesh Murugan" },
  { id: 10448376, name: "Mariya Antony Britto" },
  { id: 10448380, name: "Sarath Kumar Ravikumar" },
  { id: 10448379, name: "Surbash Lakshmi Gandhan" },
  { id: 10448386, name: "Karthikeyan Panchavaranam" },
  { id: 10448388, name: "Dharsini Nethaji" },
  { id: 10448378, name: "Tamilarasi Balamurugan" },
  { id: 10448385, name: "Nadhiya Siva Subramanian" },
  { id: 10448393, name: "Sathish Kumar Venkatesan" },
  { id: 10449931, name: "Angu selvam Murugan" },
  { id: 10450247, name: "Ranjana Mohan" },
  { id: 10450249, name: "Pradeep Joel Xavier" },
  { id: 10450402, name: "Vedhasree Manivannan" },
  { id: 10451121, name: "Anitha Ananthan" },
  { id: 10451414, name: "Sarathirajan K" },
  { id: 10451358, name: "Siddhanth Ramesh" },
  { id: 10453089, name: "Divya Shree" },
  { id: 10453088, name: "Sneha Hari Doss" },
  { id: 10453090, name: "Manoj Thiruppathi" },
  { id: 10453092, name: "Sandhiya Kollapuri" },
  { id: 10453152, name: "Kirthika Jayaraman" },
  { id: 10457539, name: "Saranya Selvamani" },
  { id: 10466495, name: "Naveen Kumar Sankar" },
  { id: 10468964, name: "Hemavathy Rajendran" },
  { id: 10470269, name: "Amrutha Rajan" },
  { id: 10471150, name: "Nivedhaa Mohankumar" },
  { id: 10479182, name: "Anurag M" },
  { id: 10479183, name: "Uday Kumar" },
  { id: 10479181, name: "Sabari Ganesh K" },
  { id: 10480914, name: "Gowthami Jayashankar" },
  { id: 10481531, name: "Saran Kirthic" },
  { id: 10480915, name: "Bhavani Dhanabalan" },
  { id: 10480917, name: "Yugeshwaran Aroumougam" },
  { id: 10481530, name: "Sonia Selva Kumar" },
  { id: 10484450, name: "Mahalakshmi Nagaraj" },
  { id: 10480916, name: "Shayan Ahmed Viringipuram" },
  { id: 10488858, name: "Harini S K" },
  { id: 10508240, name: "Iswarya Jayabalan" },
  { id: 10470689, name: "Sudha Birendarkumar" },
  { id: 10470691, name: "Naveen Kumar Anandan" },
  { id: 10470693, name: "Priya Dharshini K" },
  { id: 10470993, name: "Ritesh Suresh" },
  { id: 10470976, name: "Deepika Sampath Kumar" },
  { id: 10470692, name: "Sruthi Mathivanan" },
  { id: 10471128, name: "Rangarajan Basker" },
  { id: 10471013, name: "Tarun Akash Pazhani S" },
  { id: 10470694, name: "Rojini.S Sathish Kumar" },
  { id: 10471007, name: "Akash N Natarajan C" },
  { id: 10470998, name: "Madhumitha.C Chandhiran.N" },
  { id: 10470997, name: "Najir Hussain Nashim Miyan" },
  { id: 10470679, name: "Logeshwari S Sundaramoorthy" },
  { id: 10514086, name: "Yashvanth Munusamy" },
  { id: 10514083, name: "Somalakshmi Dhanachezhiyan" },
  { id: 10514084, name: "Srilekha P" },
  { id: 10514076, name: "Pooja Gnanaprakasam" },
  { id: 10514077, name: "Dhanush Siva" },
  { id: 10514337, name: "Mohamed Jakeria" },
  { id: 10523035, name: "Lavanya Mahanti" },
  { id: 10524417, name: "Karthick Kumar" },
  { id: 10523034, name: "Krishnaraj Mohan" },
  { id: 10544112, name: "Keerthana J" },
  { id: 10544116, name: "Govarthan Mohan" },
  { id: 10544115, name: "Devakumar Y" },
  { id: 10544114, name: "Monisha Babu" },
  { id: 10544117, name: "Mahalakshmi G" },
  { id: 10544113, name: "Sathish E" },
  { id: 10550702, name: "Rajesh Ramesh" },
  { id: 10555847, name: "Shyam Sundar" },
  { id: 10562158, name: "Mukul Vyas Parameswar" },
  { id: 10562159, name: "Akshay V Kumar" },
  { id: 10548702, name: "Abarajithan Govindarajan" },
  { id: 10547527, name: "Abirami Panchatcharam" },
  { id: 10547580, name: "Akash P M" },
  { id: 10547528, name: "Archana Venkatesan" },
  { id: 10547546, name: "Dhivyaa Krishnakumar" },
  { id: 10548722, name: "Hemanth Kumar Anandan" },
  { id: 10547544, name: "Joshini N" },
  { id: 10547529, name: "Kavipriya Dhanasekaran" },
  { id: 10548703, name: "Mahalakshmi Madhan" },
  { id: 10547548, name: "Mahenthra Babu" },
  { id: 10547542, name: "Neelufur Begam" },
  { id: 10547547, name: "Sabarish Suresh" },
  { id: 10547577, name: "Srinidhi Venugopal" },
  { id: 10547550, name: "Sudharshan Kumaresan" },
  { id: 10547545, name: "Vikram Sudhakaran" },
];
 
function makeEntry(emp) {
  return { ...emp, status: "pending", message: "" };
}
 
function copyWorksheet(srcSheet, destSheet) {
  srcSheet.columns.forEach((col, i) => {
    if (col.width) destSheet.getColumn(i + 1).width = col.width;
  });
  srcSheet.eachRow({ includeEmpty: true }, (srcRow, rowNumber) => {
    const destRow = destSheet.getRow(rowNumber);
    destRow.height = srcRow.height;
    srcRow.eachCell({ includeEmpty: true }, (srcCell, colNumber) => {
      const destCell = destRow.getCell(colNumber);
      if (srcCell.value && typeof srcCell.value === "object" && srcCell.value.formula) {
        destCell.value = { ...srcCell.value };
      } else {
        destCell.value = srcCell.value;
      }
      destCell.style = JSON.parse(JSON.stringify(srcCell.style));
    });
    destRow.commit();
  });
  try {
    Object.keys(srcSheet._merges || {}).forEach((key) => destSheet.mergeCells(key));
  } catch (_) {}
}
 
const STATUS_META = {
  pending:    { bg: "#f1f3f5", color: "#6c757d", label: "Pending"     },
  processing: { bg: "#fff3cd", color: "#856404", label: "Processing…" },
  done:       { bg: "#d1e7dd", color: "#0f5132", label: "Done ✓"      },
  skipped:    { bg: "#e2e3e5", color: "#41464b", label: "Skipped"      },
  error:      { bg: "#f8d7da", color: "#842029", label: "Error"        },
};
// heydfdfdsdf
 
const LOG_COLOR = { info: "#d4d4d4", success: "#6fcf97", warn: "#f2c94c", error: "#eb5757" };
 
export default function EmployeeReportGenerator() {
  const [file, setFile]                   = useState(null);
  const [rowsToInsert, setRowsToInsert]   = useState([""]);
  const [selectedColor, setSelectedColor] = useState("#D9E1F2");
  const [isProcessing, setIsProcessing]   = useState(false);
 
  const [newId, setNewId]       = useState("");
  const [newName, setNewName]   = useState("");
  const [idStatus, setIdStatus] = useState(null);
 
  const [bulkInput, setBulkInput]   = useState("");
  const [bulkErrors, setBulkErrors] = useState([]);
 
  const [queue, setQueue]           = useState([]);
  const [progressLog, setProgressLog] = useState([]);
  const logRef = useRef(null);
 
  const addLog = (msg, type = "info") => {
    setProgressLog((prev) => {
      const next = [...prev, { msg, type }];
      setTimeout(() => {
        if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
      }, 40);
      return next;
    });
  };
 
  const setEntryStatus = (id, status, message = "") =>
    setQueue((prev) => prev.map((e) => e.id === id ? { ...e, status, message } : e));
 
  // ── Row helpers ──────────────────────────────────────────────────────────
  const handleRowChange = (i, v) => {
    const u = [...rowsToInsert]; u[i] = v; setRowsToInsert(u);
  };
 
  // ── Single-add ───────────────────────────────────────────────────────────
  const handleIdChange = (v) => {
    setNewId(v);
    const parsed = parseInt(v, 10);
    if (!isNaN(parsed)) {
      const match = INITIAL_EMPLOYEE_DATA.find((e) => e.id === parsed);
      if (match) { setNewName(match.name); setIdStatus("found"); }
      else        { setNewName("");        setIdStatus("not_found"); }
    } else { setNewName(""); setIdStatus(null); }
  };
 
  const handleAddSingle = () => {
    const parsed = parseInt(newId, 10);
    if (isNaN(parsed) || !newName.trim()) { alert("Enter a valid ID and name."); return; }
    if (queue.some((e) => e.id === parsed)) { alert(`ID ${parsed} already in queue.`); return; }
    setQueue((prev) => [...prev, makeEntry({ id: parsed, name: newName.trim() })]);
    setNewId(""); setNewName(""); setIdStatus(null);
  };
 
  // ── Bulk-add ─────────────────────────────────────────────────────────────
  const handleBulkAdd = () => {
    const lines = bulkInput.split("\n").map((l) => l.trim()).filter(Boolean);
    const errors = [], added = [];
    lines.forEach((line, idx) => {
      const parts = line.split(/[,\t]/).map((p) => p.trim());
      const id    = parseInt(parts[0], 10);
      const name  = parts.slice(1).join(" ").trim();
      if (parts.length < 2 || isNaN(id))   { errors.push(`Line ${idx + 1}: bad format — expected "ID, Name"`); return; }
      if (!name)                            { errors.push(`Line ${idx + 1}: name is empty`); return; }
      if (queue.some((e) => e.id === id) || added.some((e) => e.id === id)) {
        errors.push(`Line ${idx + 1}: ID ${id} duplicate — skipped`); return;
      }
      added.push(makeEntry({ id, name }));
    });
    setBulkErrors(errors);
    if (added.length) { setQueue((prev) => [...prev, ...added]); setBulkInput(""); }
  };
 
  // ── Process ───────────────────────────────────────────────────────────────
  const processFile = async () => {
    if (!file)          { alert("Upload an Excel file first.");       return; }
    if (!queue.length)  { alert("Add at least one user to the queue."); return; }
 
    setIsProcessing(true);
    setProgressLog([]);
    setQueue((prev) => prev.map((e) => ({ ...e, status: "pending", message: "" })));
 
    try {
      addLog("Reading workbook…");
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      addLog("Workbook loaded.", "success");
 
      if (rowsToInsert.some((r) => r.trim())) {
        addLog("Inserting rows above 'Utilization %'…");
        workbook.eachSheet((ws) => {
          let target = null;
          ws.eachRow((row, ri) => row.eachCell((c) => {
            if (c.value?.toString().toLowerCase() === "utilization %") target = ri;
          }));
          if (target !== null) {
            const maxCol = ws.columnCount;
            rowsToInsert.forEach((val, i) => {
              const at = target + i;
              ws.spliceRows(at, 0, []);
              const nr = ws.getRow(at);
              nr.getCell(1).value = val;
              for (let col = 1; col <= maxCol; col++) {
                const cell = nr.getCell(col);
                cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + selectedColor.replace("#", "") } };
                cell.font = { bold: true };
                cell.alignment = { horizontal: "center" };
                cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
              }
              nr.commit();
            });
          }
        });
        addLog("Rows inserted.", "success");
      }
 
      const refSheet = workbook.getWorksheet("REF");
      if (!refSheet) {
        addLog("REF sheet not found — cannot create user sheets.", "error");
      } else {
        const existingNames = new Set(workbook.worksheets.map((s) => s.name));
        addLog(`Creating sheets for ${queue.length} user(s)…`);
 
        for (const emp of queue) {
          setEntryStatus(emp.id, "processing");
          await new Promise((res) => setTimeout(res, 40));
 
          if (existingNames.has(emp.name)) {
            setEntryStatus(emp.id, "skipped", "sheet already exists");
            addLog(`  ⊘  ${emp.name} — skipped (sheet exists)`, "warn");
            continue;
          }
          try {
            const newSheet = workbook.addWorksheet(emp.name);
            copyWorksheet(refSheet, newSheet);
            newSheet.getCell("A1").value = emp.id;
            existingNames.add(emp.name);
            setEntryStatus(emp.id, "done", "sheet created");
            addLog(`  ✓  ${emp.name} (${emp.id}) — sheet created`, "success");
          } catch (err) {
            setEntryStatus(emp.id, "error", err.message);
            addLog(`  ✗  ${emp.name} — ${err.message}`, "error");
          }
          await new Promise((res) => setTimeout(res, 40));
        }
      }
 
      addLog("Writing output…");
      const out = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([out]), "Updated_File.xlsx");
      addLog("Downloaded successfully!", "success");
    } catch (err) {
      addLog(`Fatal: ${err.message}`, "error");
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };
 
  // ── Derived counts ────────────────────────────────────────────────────────
  const counts = queue.reduce((acc, e) => { acc[e.status] = (acc[e.status] || 0) + 1; return acc; }, {});
 
  const Badge = ({ status }) => {
    const m = STATUS_META[status];
    return (
      <span style={{ background: m.bg, color: m.color, borderRadius: 4,
        padding: "2px 8px", fontSize: 12, fontWeight: 500, whiteSpace: "nowrap" }}>
        {m.label}
      </span>
    );
  };
 
  return (
    <div className="container mt-4 mb-5" style={{ maxWidth: 860 }}>
      <div className="card shadow-sm p-4">
        <h4 className="text-center mb-4">Stats New Process Installer</h4>
 
        {/* File */}
        <div className="mb-3">
          <label className="form-label fw-semibold">Upload Excel file</label>
          <input type="file" className="form-control" accept=".xlsx"
            onChange={(e) => setFile(e.target.files[0])} />
        </div>
 
        {/* Rows */}
        <div className="mb-3">
          <label className="form-label fw-semibold">Rows to insert above "Utilization %"</label>
          {rowsToInsert.map((row, i) => (
            <div className="input-group mb-2" key={i}>
              <input type="text" className="form-control" value={row}
                onChange={(e) => handleRowChange(i, e.target.value)} placeholder="Row text" />
              <button className="btn btn-outline-danger btn-sm"
                onClick={() => setRowsToInsert(rowsToInsert.filter((_, idx) => idx !== i))}>
                Remove
              </button>
            </div>
          ))}
          {rowsToInsert.length < 10 && (
            <button className="btn btn-outline-primary btn-sm mt-1"
              onClick={() => setRowsToInsert([...rowsToInsert, ""])}>+ Add row</button>
          )}
        </div>
 
        {/* Colour */}
        <div className="mb-4 d-flex align-items-center gap-3">
          <label className="form-label fw-semibold mb-0">Row colour</label>
          <input type="color" className="form-control form-control-color" style={{ width: 48 }}
            value={selectedColor} onChange={(e) => setSelectedColor(e.target.value)} />
          <span style={{ fontSize: 13, color: "#6c757d" }}>{selectedColor}</span>
        </div>
 
        <hr />
        <h5 className="mb-3">Add new users to queue</h5>
 
        <div className="row g-3 mb-4">
          {/* Single */}
          <div className="col-md-5">
            <div className="card border p-3 h-100">
              <p className="fw-semibold mb-2" style={{ fontSize: 14 }}>Single user</p>
              <input type="text"
                className={`form-control form-control-sm mb-1 ${idStatus === "found" ? "is-valid" : idStatus === "not_found" ? "is-invalid" : ""}`}
                placeholder="Employee ID" value={newId}
                onChange={(e) => handleIdChange(e.target.value)} />
              {idStatus === "found"     && <div className="valid-feedback d-block mb-1" style={{ fontSize: 12 }}>Auto-filled from master list</div>}
              {idStatus === "not_found" && <div className="invalid-feedback d-block mb-1" style={{ fontSize: 12 }}>Not in master — enter name manually</div>}
              <input type="text" className="form-control form-control-sm mb-2"
                placeholder="Employee name" value={newName}
                onChange={(e) => setNewName(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && handleAddSingle()} />
              <button className="btn btn-primary btn-sm" onClick={handleAddSingle}>Add to queue</button>
            </div>
          </div>
 
          {/* Bulk */}
          <div className="col-md-7">
            <div className="card border p-3 h-100">
              <p className="fw-semibold mb-1" style={{ fontSize: 14 }}>
                Bulk add
                <span style={{ fontWeight: 400, color: "#6c757d", fontSize: 13 }}> — one per line: ID, Name</span>
              </p>
              <textarea className="form-control form-control-sm mb-2" rows={5}
                placeholder={"10512345, Ravi Kumar\n10512346, Anitha Devi\n10512347, Karthik R"}
                value={bulkInput} onChange={(e) => setBulkInput(e.target.value)} />
              {bulkErrors.length > 0 && (
                <ul className="mb-2" style={{ fontSize: 12, color: "#842029", paddingLeft: 18 }}>
                  {bulkErrors.map((err, i) => <li key={i}>{err}</li>)}
                </ul>
              )}
              <button className="btn btn-primary btn-sm" onClick={handleBulkAdd}>
                Add all to queue
              </button>
            </div>
          </div>
        </div>
 
        {/* Queue table */}
        {queue.length > 0 && (
          <div className="mb-4">
            <div className="d-flex align-items-center justify-content-between mb-2">
              <span className="fw-semibold" style={{ fontSize: 14 }}>
                Queue — {queue.length} user{queue.length !== 1 ? "s" : ""}
                {(isProcessing || (counts.done || counts.skipped || counts.error)) ? (
                  <span className="ms-2" style={{ fontSize: 12, color: "#6c757d" }}>
                    {counts.done || 0} done · {counts.skipped || 0} skipped · {counts.error || 0} error{(counts.error || 0) !== 1 ? "s" : ""}
                  </span>
                ) : null}
              </span>
              {!isProcessing && (
                <button className="btn btn-outline-danger btn-sm" onClick={() => setQueue([])}>
                  Clear all
                </button>
              )}
            </div>
 
            {/* Summary bar when processing */}
            {isProcessing && (
              <div className="mb-2">
                <div className="progress" style={{ height: 6, borderRadius: 4 }}>
                  <div className="progress-bar bg-success" style={{ width: `${((counts.done || 0) / queue.length) * 100}%`, transition: "width 0.3s" }} />
                  <div className="progress-bar bg-warning" style={{ width: `${((counts.processing || 0) / queue.length) * 100}%`, transition: "width 0.3s" }} />
                </div>
              </div>
            )}
 
            <div style={{ maxHeight: 280, overflowY: "auto", border: "1px solid #dee2e6", borderRadius: 6 }}>
              <table className="table table-sm mb-0" style={{ fontSize: 13 }}>
                <thead style={{ position: "sticky", top: 0, background: "#f8f9fa", zIndex: 1 }}>
                  <tr>
                    <th style={{ width: 32 }}>#</th>
                    <th style={{ width: 110 }}>ID</th>
                    <th>Name</th>
                    <th style={{ width: 120 }}>Status</th>
                    <th style={{ width: 36 }}></th>
                  </tr>
                </thead>
                <tbody>
                  {queue.map((emp, idx) => (
                    <tr key={emp.id}
                      style={{ background: emp.status === "processing" ? "#fffde7" : emp.status === "done" ? "#f0fff4" : emp.status === "error" ? "#fff5f5" : "" }}>
                      <td style={{ color: "#6c757d" }}>{idx + 1}</td>
                      <td style={{ fontFamily: "monospace" }}>{emp.id}</td>
                      <td>
                        {emp.status === "processing" && (
                          <span className="spinner-border spinner-border-sm me-1"
                            style={{ width: 11, height: 11, borderWidth: 2, verticalAlign: "middle" }} />
                        )}
                        {emp.name}
                        {emp.message && (
                          <span style={{ color: "#6c757d", fontSize: 11, marginLeft: 6 }}>({emp.message})</span>
                        )}
                      </td>
                      <td><Badge status={emp.status} /></td>
                      <td>
                        {!isProcessing && (
                          <button className="btn btn-sm" style={{ padding: "0 5px", fontSize: 13, color: "#adb5bd" }}
                            onClick={() => setQueue((prev) => prev.filter((e) => e.id !== emp.id))}>✕</button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
 
        {/* Log */}
        {progressLog.length > 0 && (
          <div className="mb-4">
            <p className="fw-semibold mb-1" style={{ fontSize: 14 }}>Processing log</p>
            <div ref={logRef} style={{
              background: "#1a1a2e", borderRadius: 6, padding: "10px 14px",
              fontFamily: "monospace", fontSize: 12, maxHeight: 190,
              overflowY: "auto", lineHeight: 1.8
            }}>
              {progressLog.map((entry, i) => (
                <div key={i} style={{ color: LOG_COLOR[entry.type] || LOG_COLOR.info }}>
                  {entry.msg}
                </div>
              ))}
            </div>
          </div>
        )}
 
        {/* Action button */}
        <button className="btn btn-success w-100 py-2" onClick={processFile}
          disabled={isProcessing || !queue.length || !file}>
          {isProcessing ? (
            <>
              <span className="spinner-border spinner-border-sm me-2" />
              Processing {queue.length} user{queue.length !== 1 ? "s" : ""}…
            </>
          ) : (
            `Process & Download  (${queue.length} user${queue.length !== 1 ? "s" : ""})`
          )}
        </button>
      </div>
    </div>
  );
}
 
