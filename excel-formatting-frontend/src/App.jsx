import { useState } from 'react'
import Papa from "papaparse"
import * as XLSX from "xlsx"

export default function App() {
  const [csvFile, setCsvFile] = useState(null)
  const [xlsxFile, setXlsxFile] = useState(null)
  const [csvHeaders, setCsvHeaders] = useState([])
  const [xlsxHeaders, setXlsxHeaders] = useState([])
  const [mapping, setMapping] = useState({})
  const [csvData, setCsvData] = useState([])

  const handleCsvUpload = (e) => {
    const file = e.target.files[0]
    setCsvFile(file)
    Papa.parse(file, {
      complete: (result) => {
        setCsvHeaders(result.data[0])
        setCsvData(result.data)
      }
    })
  }

  const handleXlsxUpload = (e) => {
    const file = e.target.files[0]
    setXlsxFile(file)
    const reader = new FileReader()
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "array" })
      const ws = workbook.Sheets[workbook.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
      setXlsxHeaders(rows[0])
    }
    reader.readAsArrayBuffer(file)
  }

  const handleGenerate = () => {
    const reader = new FileReader()
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: "array", cellStyles: true })
      const sheetName = workbook.SheetNames[0]
      const ws = workbook.Sheets[sheetName]
      const headerRow = XLSX.utils.sheet_to_json(ws, { header: 1 })[0]
      const csvRows = csvData.slice(1)
      const today = new Date()
      const day = String(today.getDate()).padStart(2, "0")
      const month = String(today.getMonth() + 1).padStart(2, "0")
      const year = today.getFullYear()
      const fileName = `Supplier Product (${day}.${month}.${year}).xlsx`

      csvRows.forEach((csvRow, rowIndex) => {
        headerRow.forEach((xlsxCol, colIndex) => {
          const mappedCsvCol = mapping[xlsxCol]
          if (!mappedCsvCol) return
          const csvIndex = csvHeaders.indexOf(mappedCsvCol)
          const value = csvRow[csvIndex] ?? ""
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex + 1, c: colIndex })
          ws[cellAddress] = { v: value, t: "s" }
        })
      })

      XLSX.writeFile(workbook, fileName, { cellStyles: true })
    }
    reader.readAsArrayBuffer(xlsxFile)
  }

  const mappedCount = Object.values(mapping).filter(Boolean).length

  return (
    <div style={styles.page}>
      <div style={styles.container}>

        {/* Header */}
        <div style={styles.header}>
          <div style={styles.headerIcon}>⇌</div>
          <div>
            <h1 style={styles.title}>Column Mapper</h1>
            <p style={styles.subtitle}>Map your CSV data into an Excel template — no copy-pasting</p>
          </div>
        </div>

        {/* Step 1 & 2 - Upload */}
        <div style={styles.card}>
          <p style={styles.stepLabel}>Step 1 — Upload files</p>
          <div style={styles.uploadRow}>

            <label style={styles.dropzone}>
              <input type="file" accept=".csv" onChange={handleCsvUpload} style={{ display: "none" }} />
              <span style={styles.dropzoneIcon}>📊</span>
              {csvFile
                ? <><span style={styles.fileName}>✓ {csvFile.name}</span><span style={styles.replaceText}>Click to replace</span></>
                : <><span style={styles.dropzoneLabel}>Upload CSV</span><span style={styles.dropzoneHint}>Your source data</span></>
              }
            </label>

            <span style={styles.arrow}>→</span>

            <label style={styles.dropzone}>
              <input type="file" accept=".xlsx" onChange={handleXlsxUpload} style={{ display: "none" }} />
              <span style={styles.dropzoneIcon}>📋</span>
              {xlsxFile
                ? <><span style={styles.fileName}>✓ {xlsxFile.name}</span><span style={styles.replaceText}>Click to replace</span></>
                : <><span style={styles.dropzoneLabel}>Upload Excel Template</span><span style={styles.dropzoneHint}>The file to fill in</span></>
              }
            </label>

          </div>
        </div>

        {/* Step 3 - Mapping */}
        {xlsxHeaders.length > 0 && csvHeaders.length > 0 && (
          <div style={styles.card}>
            <div style={styles.mappingHeader}>
              <p style={styles.stepLabel}>Step 2 — Map columns</p>
              <span style={styles.badge}>{mappedCount} / {xlsxHeaders.length} mapped</span>
            </div>
            <p style={styles.mappingHint}>For each Excel column, choose which CSV column to pull from.</p>

            <div style={styles.mappingTableHeader}>
              <span>Excel column</span>
              <span></span>
              <span>CSV column</span>
            </div>

            {xlsxHeaders.map((xlsxCol) => (
              <div key={xlsxCol} style={styles.mappingRow}>
                <span style={styles.colName}>{xlsxCol}</span>
                <span style={styles.rowArrow}>→</span>
                <select
                  style={styles.select}
                  onChange={(e) => setMapping((prev) => ({ ...prev, [xlsxCol]: e.target.value }))}
                >
                  <option value="">— skip —</option>
                  {csvHeaders.map((csvCol) => (
                    <option key={csvCol} value={csvCol}>{csvCol}</option>
                  ))}
                </select>
              </div>
            ))}
          </div>
        )}

        {/* Generate */}
        {mappedCount > 0 && (
          <button style={styles.button} onClick={handleGenerate}>
            Generate Excel →
          </button>
        )}

      </div>
    </div>
  )
}

const styles = {
  page: {
    minHeight: "100vh",
    background: "#f4f6f9",
    display: "flex",
    justifyContent: "center",
    padding: "48px 16px",
    fontFamily: "'Segoe UI', sans-serif",
  },
  container: {
    width: "100%",
    maxWidth: 680,
    display: "flex",
    flexDirection: "column",
    gap: 20,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: 16,
    marginBottom: 8,
  },
  headerIcon: {
    width: 48,
    height: 48,
    background: "#1a56db",
    borderRadius: 12,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 22,
    color: "#fff",
    flexShrink: 0,
  },
  title: {
    margin: 0,
    fontSize: 26,
    fontWeight: 700,
    color: "#111827",
  },
  subtitle: {
    margin: "4px 0 0",
    fontSize: 14,
    color: "#6b7280",
  },
  card: {
    background: "#fff",
    borderRadius: 12,
    padding: "24px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.08)",
    border: "1px solid #e5e7eb",
  },
  stepLabel: {
    margin: "0 0 16px",
    fontWeight: 600,
    fontSize: 13,
    color: "#1a56db",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  uploadRow: {
    display: "flex",
    gap: 12,
    alignItems: "center",
  },
  dropzone: {
    flex: 1,
    border: "2px dashed #d1d5db",
    borderRadius: 10,
    padding: "24px 16px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: 6,
    cursor: "pointer",
    transition: "border-color 0.2s",
    textAlign: "center",
  },
  dropzoneIcon: {
    fontSize: 28,
  },
  dropzoneLabel: {
    fontWeight: 600,
    fontSize: 14,
    color: "#111827",
  },
  dropzoneHint: {
    fontSize: 12,
    color: "#9ca3af",
  },
  fileName: {
    fontWeight: 600,
    fontSize: 13,
    color: "#1a56db",
  },
  replaceText: {
    fontSize: 11,
    color: "#9ca3af",
  },
  arrow: {
    fontSize: 20,
    color: "#9ca3af",
    flexShrink: 0,
  },
  mappingHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    marginBottom: 4,
  },
  badge: {
    background: "#eff6ff",
    color: "#1a56db",
    border: "1px solid #bfdbfe",
    borderRadius: 20,
    padding: "2px 12px",
    fontSize: 12,
    fontWeight: 600,
  },
  mappingHint: {
    margin: "0 0 16px",
    fontSize: 13,
    color: "#6b7280",
  },
  mappingTableHeader: {
    display: "grid",
    gridTemplateColumns: "1fr 32px 1fr",
    padding: "0 12px 8px",
    fontSize: 11,
    fontWeight: 700,
    color: "#9ca3af",
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  mappingRow: {
    display: "grid",
    gridTemplateColumns: "1fr 32px 1fr",
    alignItems: "center",
    padding: "10px 12px",
    borderRadius: 8,
    marginBottom: 6,
    background: "#f9fafb",
    border: "1px solid #f3f4f6",
  },
  colName: {
    fontSize: 13,
    fontWeight: 500,
    color: "#111827",
  },
  rowArrow: {
    color: "#d1d5db",
    textAlign: "center",
    fontSize: 16,
  },
  select: {
    width: "100%",
    padding: "7px 10px",
    borderRadius: 6,
    border: "1px solid #d1d5db",
    fontSize: 13,
    color: "#111827",
    background: "#fff",
    cursor: "pointer",
  },
  button: {
    padding: "14px",
    background: "#1a56db",
    color: "#fff",
    border: "none",
    borderRadius: 10,
    fontSize: 15,
    fontWeight: 700,
    cursor: "pointer",
    width: "100%",
  },
}