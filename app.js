// Field mappings
const SCHEMAS = {
  master: {
    emailField:   "Email Addresses",
    nameField:    "Full Name",
    addressFields:["Address Line 1", "Address Line 2", "Address Line 3"],
    countryField: "Country ID",
    zipField:     "Zip Code",
    stateField:   "State",
    customerField:"Customer #"
  },
  small: {
    emailField:   "Email Address",
    nameField:    "Name",
    addressFields:["Address 1", "Address 2", "Address 3"],
    countryField: "Country ID",
    zipField:     "Postal Code",
    stateField:   "State / Region"
  }
};

// Levenshtein distance & normalized similarity
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + cost
      );
    }
  }
  return dp[m][n];
}
function similarity(a, b) {
  if (!a.length && !b.length) return 1;
  const dist = levenshtein(a, b);
  return (Math.max(a.length, b.length) - dist) / Math.max(a.length, b.length);
}

// UI refs
const masterInput = document.getElementById("masterFile");
const smallInput  = document.getElementById("smallFile");
const runBtn      = document.getElementById("runBtn");
const thresholdEl = document.getElementById("threshold");
const statusEl    = document.getElementById("status");

function setStatus(msg) {
  statusEl.innerText = msg;
}

// Read XLSX into JS objects
function readXlsx(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      resolve(json);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Plain export via SheetJS (for non‑styled)
function exportPlain(data, filename) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, filename);
}

// Styled export via ExcelJS (highlights trigger cell)
async function exportStyledDuplicates(data, filename) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Duplicates");

  // Header row
  const headers = Object.keys(data[0]);
  ws.addRow(headers);

  // Data + styling
  data.forEach(row => {
    const vals = headers.map(h => row[h]);
    const excelRow = ws.addRow(vals);

    let colIdx = null;
    if (row._matchMethod === "email") {
      colIdx = headers.indexOf(SCHEMAS.small.emailField) + 1;
    } else if (row._matchMethod === "name") {
      colIdx = headers.indexOf(SCHEMAS.small.nameField) + 1;
    } else if (row._matchMethod === "address") {
      colIdx = headers.indexOf(SCHEMAS.small.addressFields[0]) + 1;
    }

    if (colIdx) {
      excelRow.getCell(colIdx).fill = {
        type: 'pattern',
        pattern:'solid',
        fgColor:{ argb:'FFFFFF00' }
      };
    }
  });

  // Download
  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a   = document.createElement('a');
  a.href    = url;
  a.download= filename;
  a.click();
  URL.revokeObjectURL(url);
}

// Concatenate address lines with commas
function concatAddress(row, fields) {
  return fields
    .map(f => (row[f] || "").trim())
    .filter(x => x)
    .join(", ");
}

// Extract leading house number (digits)
function leadingNumber(addr) {
  const m = addr.match(/^(\d+)/);
  return m ? m[1] : null;
}

// Normalize names: lowercase, replace & with ' and ', collapse spaces
function normalizeName(str) {
  return str
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Main dedupe flow
runBtn.addEventListener("click", async () => {
  if (!masterInput.files[0] || !smallInput.files[0]) {
    return alert("Please select both files.");
  }
  runBtn.disabled = true;
  setStatus("Loading master list…");
  const master = await readXlsx(masterInput.files[0]);
  setStatus("Loading smaller list…");
  const small  = await readXlsx(smallInput.files[0]);

  const threshold     = parseFloat(thresholdEl.value) || 0.85;
  const duplicates    = [];
  const nonDuplicates = [];

  // Build email→master map
  const emailMap = {};
  master.forEach(row => {
    (row[SCHEMAS.master.emailField] || "")
      .split(',')
      .map(e => e.trim().toLowerCase())
      .filter(e => e)
      .forEach(email => {
        (emailMap[email] ||= []).push(row);
      });
  });

  setStatus("Comparing records…");
  for (let i = 0; i < small.length; i++) {
    const s    = small[i];
    const name = (s[SCHEMAS.small.nameField] || "").trim();
    if (!name) continue;  // skip empty

    let isDup     = false;
    let matchInfo = {};

    // 1) Exact email
    const email = (s[SCHEMAS.small.emailField] || "").trim().toLowerCase();
    if (email && emailMap[email]) {
      isDup     = true;
      matchInfo = {method: "email", matched: emailMap[email][0]};
    }

    // 2) Fuzzy name
    if (!isDup) {
      let bestScore = 0, bestMatch = null;
      const tgt = normalizeName(name);
      for (const m of master) {
        const mRaw = (m[SCHEMAS.master.nameField] || "").trim();
        const mName = normalizeName(mRaw);
        const sc    = similarity(tgt, mName);
        if (sc > bestScore) {
          bestScore = sc;
          bestMatch = m;
        }
      }
      if (bestScore >= threshold) {
        isDup     = true;
        matchInfo = {method: "name", score: bestScore.toFixed(2), matched: bestMatch};
      }
    }

    // 3) Fuzzy address (after name)
    if (!isDup) {
      // require zip & state & country
      const zipSmall     = (s[SCHEMAS.small.zipField] || "").trim();
      const stateSmall   = (s[SCHEMAS.small.stateField] || "").trim();
      const countrySmall = (s[SCHEMAS.small.countryField] || "").trim();

      // filter candidates by postal/state/country
      const candidates = master.filter(m => {
        if (countrySmall && (m[SCHEMAS.master.countryField]||"").trim() !== countrySmall)
          return false;
        if (zipSmall && (m[SCHEMAS.master.zipField]||"").trim() !== zipSmall)
          return false;
        if (stateSmall && (m[SCHEMAS.master.stateField]||"").trim() !== stateSmall)
          return false;
        return true;
      });

      const addrSmall = concatAddress(s, SCHEMAS.small.addressFields).toLowerCase();
      const numSmall  = leadingNumber(addrSmall);

      let bestScore = 0, bestMatch = null;
      for (const m of candidates) {
        const addrMaster = concatAddress(m, SCHEMAS.master.addressFields).toLowerCase();
        const numMaster  = leadingNumber(addrMaster);
        if (!numSmall || !numMaster || numSmall !== numMaster) continue;
        const sc = similarity(addrSmall, addrMaster);
        if (sc > bestScore) {
          bestScore = sc;
          bestMatch = m;
        }
      }
      if (bestScore >= threshold) {
        isDup     = true;
        matchInfo = {method: "address", score: bestScore.toFixed(2), matched: bestMatch};
      }
    }

    // collect results
    if (isDup) {
      duplicates.push({
        [SCHEMAS.master.customerField]: matchInfo.matched[SCHEMAS.master.customerField] || "",
        ...s,
        _matchMethod:     matchInfo.method,
        _matchScore:      matchInfo.score || "",
        _matchedFullName: matchInfo.matched[SCHEMAS.master.nameField] || "",
        _matchedAddress:  matchInfo.method === "address"
                            ? concatAddress(matchInfo.matched, SCHEMAS.master.addressFields)
                            : ""
      });
    } else {
      nonDuplicates.push(s);
    }

    if ((i+1) % 50 === 0) {
      setStatus(`Processed ${i+1}/${small.length}…`);
      await new Promise(r => setTimeout(r, 10));
    }
  }

  setStatus("Writing output files…");
  if (duplicates.length) {
    await exportStyledDuplicates(duplicates,    "duplicates.xlsx");
  }
  if (nonDuplicates.length) {
    exportPlain(nonDuplicates, "non_duplicates.xlsx");
  }

  setStatus(`Done! Duplicates: ${duplicates.length}, Non‑duplicates: ${nonDuplicates.length}`);
  runBtn.disabled = false;
});
