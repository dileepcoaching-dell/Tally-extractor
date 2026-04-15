# 📊 Tally Data Extractor

A Streamlit web app to extract data from **Tally Prime / Tally ERP 9** via its built-in XML/HTTP server and export it to a formatted **Excel workbook**.

---

## ✅ Features

| Feature | Details |
|---|---|
| **Company Selection** | Lists only companies currently open in Tally |
| **Profit & Loss** | Exports Income & Expense sheets with date range |
| **Balance Sheet** | Exports Assets & Liabilities as on a selected date |
| **Ledger Master** | Full ledger list with customisable columns |
| **Voucher Register** | All vouchers with party, amount, narration |
| **Excel Export** | Dark-themed, formatted multi-sheet Excel workbook |
| **Custom Ledger Columns** | Select and reorder exactly which ledger fields to export |

---

## 🚀 Setup & Run

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Enable Tally XML/HTTP Server

In **Tally Prime**:
> `F1: Help` → `Settings` → `Connectivity` → Enable **Tally Gateway Server** (port 9000)

In **Tally ERP 9**:
> `F12: Configure` → `Advanced Configuration` → Enable **Enable ODBC Server** or **HTTP Server**

### 3. Run the app

```bash
streamlit run tally_extractor.py
```

The app opens in your browser at `http://localhost:8501`

---

## 🔌 How Tally Communication Works

The app sends **XML POST requests** to `http://localhost:9000` — the same HTTP port Tally exposes for third-party integrations.

Example XML request to fetch companies:
```xml
<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>List of Companies</REPORTNAME>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>
```

---

## 📁 Output Excel Format

Each export creates a `.xlsx` file with separate sheets:

| Sheet Name | Contents |
|---|---|
| `Ledgers` | Ledger master with your selected columns |
| `P&L Income` | Income / Revenue groups & ledgers |
| `P&L Expenses` | Expense groups & ledgers |
| `BS Liabilities` | Capital, loans, creditors |
| `BS Assets` | Fixed assets, debtors, cash/bank |
| `Vouchers` | All voucher entries (optional) |

---

## 🛠️ Troubleshooting

| Issue | Fix |
|---|---|
| "Cannot reach Tally" | Check Tally is open and HTTP server is enabled |
| Empty company list | Open at least one company in Tally first |
| Empty data in sheets | Verify date range has transactions in Tally |
| Port error | Default is 9000; check Tally config for actual port |
