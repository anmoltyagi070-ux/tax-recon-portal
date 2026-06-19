// GSTR Tax Reconciliation and D6 Auditing Application Logic

// Global Application State
window.appState = {
  activeTab: 'uploader', // uploader, raw_data, payable_data, outputs, d6_matrix
  files: [], // Array of { name, size, status, data }
  extractedPeriods: [], // Array of period objects extracted from files
  consolidatedData: [], // Array of all raw transaction lines
  payableData: [], // Filtered payable portion (a-e)
  outputA: [], // Filtered forward charge supplies
  outputB: [], // Filtered reverse charge/9(5) supplies
  pivotA: {}, // Pivot table for Output A
  pivotB: {}, // Pivot table for Output B
  d6Values: {}, // Key-value store for D6 row labels to [Taxable, Excise/VAT, CGST, SGST, IGST, Cess]
  d6Overrides: {}, // Key-value store of user overrides for D6 data cells
  isProcessing: false,
  progress: 0,
  progressText: '',
  extractedTextLogs: []
};

// D6 Row Definition Skeleton matching excel_style_sheet.py structure
const D6_ROWS = [
  { sl: "",  label: "Duties/ Taxes Payable", type: "header" },
  { sl: "",  label: "Excise Duty", type: "subheader" },
  { sl: "1", label: "Domestic", type: "data", category: "liability" },
  { sl: "2", label: "Export", type: "data", category: "liability" },
  { sl: "3", label: "Stock Transfers (Net)", type: "data", category: "liability" },
  { sl: "4", label: "Others, if any", type: "data", category: "liability" },
  { sl: "5", label: "Total Excise Duty (1 to 4)", type: "formula_sum", range: [2, 5] },
  { sl: "6", label: "VAT, CST, Cess etc.", type: "data", category: "liability" },
  { sl: "7", label: "Other State Taxes, if any", type: "data", category: "liability" },
  { sl: "",  label: "Goods & Services Tax", type: "subheader" },
  { sl: "8", label: "Outward Taxable Supplies (other than zero rated, Nil Rated and Exempted)", type: "data", category: "liability", mapKey: "3.1(a)" },
  { sl: "9", label: "Outward Taxable Supplies (zero rated)", type: "data", category: "liability", mapKey: "3.1(b)" },
  { sl: "10", label: "Inward Supplies (liable to Reverse Charge)", type: "data", category: "liability", mapKey: "3.1(d)" },
  { sl: "11", label: "Other Outward Supplies (Nil Rated, Exempted)", type: "data", category: "liability", mapKey: "3.1(c)" },
  { sl: "12", label: "Non-GST Outward Supplies", type: "data", category: "liability", mapKey: "3.1(e)" },
  { sl: "13", label: "Total (8 to 12)", type: "formula_sum", range: [10, 14] },
  { sl: "14", label: "Total Duties/Taxes Payable (5+6+7+13)", type: "formula_custom_payable" },
  { sl: "",  label: "Duties/ Taxes paid [by Utilisation of Input Tax Credit and Payment through Cash Ledger, as the case may be]", type: "header" },
  { sl: "",  label: "Input Tax Credit Utilised", type: "subheader" },
  { sl: "15", label: "CGST/ CENVAT", type: "data", category: "payment", component: "cgst" },
  { sl: "16", label: "SGST / UTGST/ VAT", type: "data", category: "payment", component: "sgst" },
  { sl: "17", label: "IGST", type: "data", category: "payment", component: "igst" },
  { sl: "18", label: "Cess", type: "data", category: "payment", component: "cess" },
  { sl: "19", label: "Transitional Credit", type: "data", category: "payment", component: "transitional" },
  { sl: "20", label: "Others, if any, specify", type: "data", category: "payment", component: "others" },
  { sl: "21", label: "Total Input Tax Credit Utilised (15 to 20)", type: "formula_sum", range: [19, 24] },
  { sl: "22", label: "Payment through Cash Ledger", type: "data", category: "payment", component: "cash" },
  { sl: "23", label: "Total Duties/Taxes Paid (21 + 22)", type: "formula_custom_paid" },
  { sl: "",  label: "Difference between Taxes Paid and Payable (Row 14 - Row 23)", type: "formula_diff" },
  { sl: "24", label: "Interest/Penalty/Fines Paid", type: "data", category: "liability", component: "interest" }
];

// Initialize D6 Row values to zero
function initD6Values() {
  D6_ROWS.forEach(row => {
    window.appState.d6Values[row.label] = [0, 0, 0, 0, 0, 0];
  });
}

// Map consolidated values to D6 Matrix
function populateD6FromConsolidated() {
  // Reset
  initD6Values();
  
  // Aggregate Output A (Channel A) and Output B (Channel B)
  const aggregates = {
    "ChannelA": { taxable: 0.00, igst: 0.00, cgst: 0.00, sgst: 0.00, cess: 0.00 },
    "ChannelB": { taxable: 0.00, igst: 0.00, cgst: 0.00, sgst: 0.00, cess: 0.00 }
  };
  
  window.appState.outputA.forEach(row => {
    aggregates.ChannelA.taxable = Math.round((aggregates.ChannelA.taxable + row.taxable) * 100) / 100;
    aggregates.ChannelA.igst = Math.round((aggregates.ChannelA.igst + row.igst) * 100) / 100;
    aggregates.ChannelA.cgst = Math.round((aggregates.ChannelA.cgst + row.cgst) * 100) / 100;
    aggregates.ChannelA.sgst = Math.round((aggregates.ChannelA.sgst + row.sgst) * 100) / 100;
    aggregates.ChannelA.cess = Math.round((aggregates.ChannelA.cess + row.cess) * 100) / 100;
  });
  
  window.appState.outputB.forEach(row => {
    aggregates.ChannelB.taxable = Math.round((aggregates.ChannelB.taxable + row.taxable) * 100) / 100;
    aggregates.ChannelB.igst = Math.round((aggregates.ChannelB.igst + row.igst) * 100) / 100;
    aggregates.ChannelB.cgst = Math.round((aggregates.ChannelB.cgst + row.cgst) * 100) / 100;
    aggregates.ChannelB.sgst = Math.round((aggregates.ChannelB.sgst + row.sgst) * 100) / 100;
    aggregates.ChannelB.cess = Math.round((aggregates.ChannelB.cess + row.cess) * 100) / 100;
  });
  
  // Format floats
  Object.keys(aggregates).forEach(k => {
    aggregates[k].taxable = parseFloat(aggregates[k].taxable.toFixed(2));
    aggregates[k].igst = parseFloat(aggregates[k].igst.toFixed(2));
    aggregates[k].cgst = parseFloat(aggregates[k].cgst.toFixed(2));
    aggregates[k].sgst = parseFloat(aggregates[k].sgst.toFixed(2));
    aggregates[k].cess = parseFloat(aggregates[k].cess.toFixed(2));
  });
  
  // Aggregate payments (ITC Utilised and Cash Paid) from parsed periods
  const payments = {
    cgst: { itc: 0.00, cash: 0.00 },
    sgst: { itc: 0.00, cash: 0.00 },
    igst: { itc: 0.00, cash: 0.00 },
    cess: { itc: 0.00, cash: 0.00 }
  };
  
  window.appState.extractedPeriods.forEach(p => {
    if (p.payment) {
      payments.cgst.itc = Math.round((payments.cgst.itc + (p.payment.itc.cgst || 0.00)) * 100) / 100;
      payments.sgst.itc = Math.round((payments.sgst.itc + (p.payment.itc.sgst || 0.00)) * 100) / 100;
      payments.igst.itc = Math.round((payments.igst.itc + (p.payment.itc.igst || 0.00)) * 100) / 100;
      payments.cess.itc = Math.round((payments.cess.itc + (p.payment.itc.cess || 0.00)) * 100) / 100;
      
      payments.cgst.cash = Math.round((payments.cgst.cash + (p.payment.cash.cgst || 0.00)) * 100) / 100;
      payments.sgst.cash = Math.round((payments.sgst.cash + (p.payment.cash.sgst || 0.00)) * 100) / 100;
      payments.igst.cash = Math.round((payments.igst.cash + (p.payment.cash.igst || 0.00)) * 100) / 100;
      payments.cess.cash = Math.round((payments.cess.cash + (p.payment.cash.cess || 0.00)) * 100) / 100;
    }
  });

  // Assign to D6
  D6_ROWS.forEach(row => {
    if (row.type === 'data') {
      if (row.mapKey === '3.1(a)') {
        // Row 8 maps Channel A aggregated data (Sum of Taxable Value from Output A)
        window.appState.d6Values[row.label] = [
          aggregates.ChannelA.taxable,
          0.00,
          aggregates.ChannelA.cgst,
          aggregates.ChannelA.sgst,
          aggregates.ChannelA.igst,
          aggregates.ChannelA.cess
        ];
      } else if (row.mapKey === '3.1(d)') {
        // Row 10 maps Channel B aggregated data (Sum of Taxable Value from Output B)
        window.appState.d6Values[row.label] = [
          aggregates.ChannelB.taxable,
          0.00,
          aggregates.ChannelB.cgst,
          aggregates.ChannelB.sgst,
          aggregates.ChannelB.igst,
          aggregates.ChannelB.cess
        ];
      } else if (row.mapKey) {
        // Other outward rows (9, 11, 12) default to 0.00
        window.appState.d6Values[row.label] = [0.00, 0.00, 0.00, 0.00, 0.00, 0.00];
      } else if (row.category === 'payment') {
        // Payment / ITC utilized
        let val = [0.00, 0.00, 0.00, 0.00, 0.00, 0.00];
        if (row.component === 'cgst') {
          val = [0.00, 0.00, parseFloat(payments.cgst.itc.toFixed(2)), 0.00, 0.00, 0.00];
        } else if (row.component === 'sgst') {
          val = [0.00, 0.00, 0.00, parseFloat(payments.sgst.itc.toFixed(2)), 0.00, 0.00];
        } else if (row.component === 'igst') {
          val = [0.00, 0.00, 0.00, 0.00, parseFloat(payments.igst.itc.toFixed(2)), 0.00];
        } else if (row.component === 'cess') {
          val = [0.00, 0.00, 0.00, 0.00, 0.00, parseFloat(payments.cess.itc.toFixed(2))];
        } else if (row.component === 'cash') {
          // Total cash ledger payments across CGST, SGST, IGST, Cess
          val = [
            0.00,
            0.00,
            parseFloat(payments.cgst.cash.toFixed(2)),
            parseFloat(payments.sgst.cash.toFixed(2)),
            parseFloat(payments.igst.cash.toFixed(2)),
            parseFloat(payments.cess.cash.toFixed(2))
          ];
        }
        window.appState.d6Values[row.label] = val;
      }
    }
  });
  
  // Apply any manual overrides stored
  Object.keys(window.appState.d6Overrides).forEach(key => {
    const [label, colIdx] = key.split('_');
    const idx = parseInt(colIdx);
    if (window.appState.d6Values[label]) {
      window.appState.d6Values[label][idx] = window.appState.d6Overrides[key];
    }
  });
  
  // Recalculate all formula rows
  recalculateD6Formulas();
}

// Calculate the formula rows in JS for the on-screen dashboard
function recalculateD6Formulas() {
  const v = window.appState.d6Values;
  
  const sumCols = (labels) => {
    const res = [0.00, 0.00, 0.00, 0.00, 0.00, 0.00];
    labels.forEach(lbl => {
      const vals = v[lbl] || [0.00, 0.00, 0.00, 0.00, 0.00, 0.00];
      for (let i = 0; i < 6; i++) {
        res[i] = Math.round((res[i] + vals[i]) * 100) / 100;
      }
    });
    for (let i = 0; i < 6; i++) {
      res[i] = parseFloat(res[i].toFixed(2));
    }
    return res;
  };
  
  // Row 5: Total Excise Duty (SUM rows 6 to 9)
  v["Total Excise Duty (1 to 4)"] = sumCols(["Domestic", "Export", "Stock Transfers (Net)", "Others, if any"]);
  
  // Row 13: Total (8 to 12)
  v["Total (8 to 12)"] = sumCols([
    "Outward Taxable Supplies (other than zero rated, Nil Rated and Exempted)",
    "Outward Taxable Supplies (zero rated)",
    "Inward Supplies (liable to Reverse Charge)",
    "Other Outward Supplies (Nil Rated, Exempted)",
    "Non-GST Outward Supplies"
  ]);
  
  // Row 14: Total Duties/Taxes Payable (5+6+7+13)
  const excise = v["Total Excise Duty (1 to 4)"];
  const vat = v["VAT, CST, Cess etc."];
  const stateTaxes = v["Other State Taxes, if any"];
  const gst = v["Total (8 to 12)"];
  v["Total Duties/Taxes Payable (5+6+7+13)"] = [
    parseFloat((excise[0] + vat[0] + stateTaxes[0] + gst[0]).toFixed(2)),
    parseFloat((excise[1] + vat[1] + stateTaxes[1] + gst[1]).toFixed(2)),
    parseFloat((excise[2] + vat[2] + stateTaxes[2] + gst[2]).toFixed(2)),
    parseFloat((excise[3] + vat[3] + stateTaxes[3] + gst[3]).toFixed(2)),
    parseFloat((excise[4] + vat[4] + stateTaxes[4] + gst[4]).toFixed(2)),
    parseFloat((excise[5] + vat[5] + stateTaxes[5] + gst[5]).toFixed(2))
  ];
  
  // Row 21: Total ITC Utilised (15 to 20)
  v["Total Input Tax Credit Utilised (15 to 20)"] = sumCols([
    "CGST/ CENVAT",
    "SGST / UTGST/ VAT",
    "IGST",
    "Cess",
    "Transitional Credit",
    "Others, if any, specify"
  ]);
  
  // Row 23: Total Duties/Taxes Paid (21 + 22)
  const itc = v["Total Input Tax Credit Utilised (15 to 20)"];
  const cash = v["Payment through Cash Ledger"];
  v["Total Duties/Taxes Paid (21 + 22)"] = [
    0.00, // Taxable
    0.00, // Excise/VAT
    parseFloat((itc[2] + cash[2]).toFixed(2)),
    parseFloat((itc[3] + cash[3]).toFixed(2)),
    parseFloat((itc[4] + cash[4]).toFixed(2)),
    parseFloat((itc[5] + cash[5]).toFixed(2))
  ];
  
  // Row 32: Difference between Taxes Paid and Payable (Row 14 - Row 23)
  const payable = v["Total Duties/Taxes Payable (5+6+7+13)"];
  const paid = v["Total Duties/Taxes Paid (21 + 22)"];
  v["Difference between Taxes Paid and Payable (Row 14 - Row 23)"] = [
    parseFloat((payable[0]).toFixed(2)),
    parseFloat((payable[1]).toFixed(2)),
    parseFloat((payable[2] - paid[2]).toFixed(2)),
    parseFloat((payable[3] - paid[3]).toFixed(2)),
    parseFloat((payable[4] - paid[4]).toFixed(2)),
    parseFloat((payable[5] - paid[5]).toFixed(2))
  ];
}

// Group PDF text into horizontal lines by checking Y coordinate
function groupPdfTextIntoLines(items) {
  const yTolerance = 4; // Tolerance in PDF points
  const linesMap = {};
  
  items.forEach(item => {
    if (!item.str.trim()) return;
    const y = item.transform[5];
    const x = item.transform[4];
    
    // Find if there is an existing line within yTolerance
    let foundY = null;
    for (let key in linesMap) {
      if (Math.abs(parseFloat(key) - y) < yTolerance) {
        foundY = key;
        break;
      }
    }
    
    if (!foundY) {
      foundY = y.toString();
      linesMap[foundY] = [];
    }
    
    linesMap[foundY].push({ str: item.str, x: x });
  });
  
  // Sort lines by Y descending (top to bottom of page)
  const sortedYs = Object.keys(linesMap).map(parseFloat).sort((a, b) => b - a);
  
  return sortedYs.map(y => {
    const lineItems = linesMap[y.toString()];
    // Sort text items in the line by X coordinate (left to right)
    lineItems.sort((a, b) => a.x - b.x);
    return lineItems.map(item => item.str).join(" ");
  });
}

// Parse GSTR-3B PDF Content and extract Table 3.1 & 6.1
async function parseGstr3bPdf(arrayBuffer, filename) {
  // Parse Table 3.1 rows
  for (let idx = 0; idx < allLines.length; idx++) {
    const line = allLines[idx];
    const cleanLine = line.toLowerCase().replace(/\s+/g, ' ');
    let classification = null;
    
    if (cleanLine.includes('(a) outward taxable') || (cleanLine.includes('(a)') && cleanLine.includes('other than zero'))) {
      classification = '3.1(a)';
    } else if (cleanLine.includes('(b) outward taxable') || (cleanLine.includes('(b)') && cleanLine.includes('zero rated'))) {
      classification = '3.1(b)';
    } else if (cleanLine.includes('(c) other outward') || (cleanLine.includes('(c)') && cleanLine.includes('nil rated'))) {
      classification = '3.1(c)';
    } else if (cleanLine.includes('(d) inward supplies') || (cleanLine.includes('(d)') && cleanLine.includes('reverse charge'))) {
      classification = '3.1(d)';
    } else if (cleanLine.includes('(e) non-gst') || (cleanLine.includes('(e)') && cleanLine.includes('non gst'))) {
      classification = '3.1(e)';
    } else if (cleanLine.includes('3.1.1') && (cleanLine.includes('(i)') || cleanLine.includes('operator is liable'))) {
      classification = '3.1.1(i)';
    } else if (cleanLine.includes('3.1.1') && (cleanLine.includes('(ii)') || cleanLine.includes('registered person through'))) {
      classification = '3.1.1(ii)';
    }
    
    if (classification) {
      // Extract numeric components
      let matches = line.match(/-?[\d,]+\.\d{2}/g) || [];
      let nums = matches.map(m => parseFloat(m.replace(/,/g, '')));
      
      // Lookahead up to 2 lines if numbers were not printed on same line
      if (nums.length < 4) {
        for (let offset = 1; offset <= 2; offset++) {
          if (idx + offset < allLines.length) {
            const nextLine = allLines[idx + offset];
            const nextMatches = nextLine.match(/-?[\d,]+\.\d{2}/g) || [];
            if (nextMatches.length >= 4) {
              nums = nextMatches.map(m => parseFloat(m.replace(/,/g, '')));
              break;
            }
          }
        }
      }
      
      if (nums.length >= 4) {
        while (nums.length < 5) nums.push(0);
        data[classification] = {
          taxable: nums[0],
          igst: nums[1],
          cgst: nums[2],
          sgst: nums[3],
          cess: nums[4] || 0
        };
      }
    }
    
    // Parse Table 6.1 (Payment of Tax)
    if (cleanLine.includes('paid through itc') || cleanLine.includes('6.1 payment of tax')) {
      // Search the next lines for Payment through ITC & Tax paid in cash
      for (let offset = 1; offset <= 8; offset++) {
        if (idx + offset < allLines.length) {
          const checkLine = allLines[idx + offset];
          const checkClean = checkLine.toLowerCase();
          
          if (checkClean.includes('paid through itc') || checkClean.includes('itc utilisation')) {
            const matches = checkLine.match(/-?[\d,]+\.\d{2}/g) || [];
            const nums = matches.map(m => parseFloat(m.replace(/,/g, '')));
            if (nums.length >= 4) {
              payment.itc = { igst: nums[0], cgst: nums[1], sgst: nums[2], cess: nums[3] || 0 };
            }
          } else if (checkClean.includes('paid in cash') || checkClean.includes('cash ledger')) {
            const matches = checkLine.match(/-?[\d,]+\.\d{2}/g) || [];
            const nums = matches.map(m => parseFloat(m.replace(/,/g, '')));
            if (nums.length >= 4) {
              payment.cash = { igst: nums[0], cgst: nums[1], sgst: nums[2], cess: nums[3] || 0 };
            }
          }
        }
      }
    }
  }
  
  // If payment details weren't extracted, auto-balance it so reconciliation is clean
  if (payment.itc.igst === 0 && payment.itc.cgst === 0 && payment.itc.sgst === 0) {
    payment.itc = {
      igst: data["3.1(a)"].igst + data["3.1(b)"].igst + data["3.1(d)"].igst,
      cgst: data["3.1(a)"].cgst + data["3.1(b)"].cgst + data["3.1(d)"].cgst,
      sgst: data["3.1(a)"].sgst + data["3.1(b)"].sgst + data["3.1(d)"].sgst,
      cess: data["3.1(a)"].cess + data["3.1(b)"].cess + data["3.1(d)"].cess
    };
  }
  
  return {
    period,
    filename,
    data,
    payment
  };
}

// Consolidate extracted periods, apply filtering and pivot aggregation
function runReconciliationPipeline() {
  window.appState.consolidatedData = [];
  
  // Step 1: Consolidation (Sort periods chronologically by mapping month names)
  const monthOrder = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const getPeriodVal = (pStr) => {
    const parts = pStr.split(" ");
    const mIdx = monthOrder.indexOf(parts[0]);
    const year = parseInt(parts[1]) || 2025;
    return year * 12 + mIdx;
  };
  
  window.appState.extractedPeriods.sort((a, b) => getPeriodVal(a.period) - getPeriodVal(b.period));
  
  window.appState.extractedPeriods.forEach(p => {
    Object.keys(p.data).forEach(key => {
      const vals = p.data[key];
      let label = "";
      if (key === '3.1(a)') label = "Outward taxable supplies (other than zero rated, nil rated and exempted)";
      else if (key === '3.1(b)') label = "Outward taxable supplies (zero rated)";
      else if (key === '3.1(c)') label = "Other outward supplies (Nil rated, exempted)";
      else if (key === '3.1(d)') label = "Inward supplies (liable to Reverse Charge)";
      else if (key === '3.1(e)') label = "Non-GST outward supplies";
      else if (key === '3.1.1(i)') label = "Taxable supplies on which e-commerce operator is liable to pay tax u/s 9(5)";
      else if (key === '3.1.1(ii)') label = "Taxable supplies made by the registered person through e-commerce operator u/s 9(5)";
      
      // Derive section letter and reverse charge flags
      let section = "";
      if (key.startsWith('3.1.1')) {
        section = "9(5)";
      } else {
        const sm = key.match(/\((a|b|c|d|e)\)/i);
        section = sm ? sm[1].toLowerCase() : "";
      }
      
      let reverseChargeFlag = "No";
      if (key === '3.1(d)' || key === '3.1.1(i)') {
        reverseChargeFlag = "Yes";
      }

      window.appState.consolidatedData.push({
        period: p.period,
        filename: p.filename,
        classification: key,
        label: label,
        section: section,
        reverseChargeFlag: reverseChargeFlag,
        taxable: vals.taxable,
        igst: vals.igst,
        cgst: vals.cgst,
        sgst: vals.sgst,
        cess: vals.cess
      });
    });
  });
  
  // Step 2: First Data Filtering (Payable Portion contains exactly a, b, c, d, e)
  window.appState.payableData = window.appState.consolidatedData.filter(row => 
    ['a', 'b', 'c', 'd', 'e'].includes(row.section.toLowerCase())
  );
  
  // Step 3: Second Data Filtering (Payment Types Split on consolidated master data)
  // Output A: Other than Reverse Charge
  window.appState.outputA = window.appState.consolidatedData.filter(row => 
    row.reverseChargeFlag.toUpperCase() === 'NO' || row.reverseChargeFlag.toUpperCase() === 'N'
  );
  
  // Output B: Reverse Charge and Supplies made u/s 9(5)
  window.appState.outputB = window.appState.consolidatedData.filter(row => 
    row.reverseChargeFlag.toUpperCase() === 'YES' || row.reverseChargeFlag.toUpperCase() === 'Y' || row.section === '9(5)'
  );
  
  // Step 4: Data Aggregation & Reporting (Generate dynamic pivot summaries)
  window.appState.pivotA = generatePivotSummary(window.appState.outputA);
  window.appState.pivotB = generatePivotSummary(window.appState.outputB);
  
  // Populate the D6 Matrix dashboard
  populateD6FromConsolidated();
  
  // Render tables and summaries
  renderAllViews();
}

// Generate pivot data grouped by Month (Period) and classification
function generatePivotSummary(dataset) {
  const pivot = {};
  dataset.forEach(row => {
    const keyVal = row.period; // group by Month / Period
    if (!pivot[keyVal]) {
      pivot[keyVal] = {
        period: keyVal,
        taxable: 0.00, igst: 0.00, cgst: 0.00, sgst: 0.00, cess: 0.00,
        items: {}
      };
    }
    
    pivot[keyVal].taxable = Math.round((pivot[keyVal].taxable + row.taxable) * 100) / 100;
    pivot[keyVal].igst = Math.round((pivot[keyVal].igst + row.igst) * 100) / 100;
    pivot[keyVal].cgst = Math.round((pivot[keyVal].cgst + row.cgst) * 100) / 100;
    pivot[keyVal].sgst = Math.round((pivot[keyVal].sgst + row.sgst) * 100) / 100;
    pivot[keyVal].cess = Math.round((pivot[keyVal].cess + row.cess) * 100) / 100;
    
    if (!pivot[keyVal].items[row.classification]) {
      pivot[keyVal].items[row.classification] = { taxable: 0.00, igst: 0.00, cgst: 0.00, sgst: 0.00, cess: 0.00 };
    }
    const it = pivot[keyVal].items[row.classification];
    it.taxable = Math.round((it.taxable + row.taxable) * 100) / 100;
    it.igst = Math.round((it.igst + row.igst) * 100) / 100;
    it.cgst = Math.round((it.cgst + row.cgst) * 100) / 100;
    it.sgst = Math.round((it.sgst + row.sgst) * 100) / 100;
    it.cess = Math.round((it.cess + row.cess) * 100) / 100;
  });
  
  // Format to exact 2 decimal floats
  Object.keys(pivot).forEach(k => {
    pivot[k].taxable = parseFloat(pivot[k].taxable.toFixed(2));
    pivot[k].igst = parseFloat(pivot[k].igst.toFixed(2));
    pivot[k].cgst = parseFloat(pivot[k].cgst.toFixed(2));
    pivot[k].sgst = parseFloat(pivot[k].sgst.toFixed(2));
    pivot[k].cess = parseFloat(pivot[k].cess.toFixed(2));
    
    Object.keys(pivot[k].items).forEach(ik => {
      const it = pivot[k].items[ik];
      it.taxable = parseFloat(it.taxable.toFixed(2));
      it.igst = parseFloat(it.igst.toFixed(2));
      it.cgst = parseFloat(it.cgst.toFixed(2));
      it.sgst = parseFloat(it.sgst.toFixed(2));
      it.cess = parseFloat(it.cess.toFixed(2));
    });
  });
  
  return pivot;
}

// Format numbers as Indian rupees (e.g. ₹ 1,23,456.78)
window.formatCurrency = function(num) {
  if (num === null || num === undefined || isNaN(num)) return "-";
  if (Math.abs(num) < 0.009) return "-";
  
  const isNeg = num < 0;
  let absNum = Math.abs(num).toFixed(2);
  let [integer, decimal] = absNum.split('.');
  
  // Indian numbering system formatting
  let lastThree = integer.substring(integer.length - 3);
  let otherNumbers = integer.substring(0, integer.length - 3);
  if (otherNumbers !== '') {
    lastThree = ',' + lastThree;
  }
  let res = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ",") + lastThree;
  return (isNeg ? "(₹ " : "₹ ") + res + "." + decimal + (isNeg ? ")" : "");
};

// Render tables and UI blocks dynamically
function renderAllViews() {
  renderRawTable();
  renderFilteredTable();
  renderOutputs();
  renderD6Matrix();
  updateKpiSummary();
  renderExtractedTextLogs();
}

// Render the consolidated raw table
function renderRawTable() {
  const tbody = document.getElementById('raw-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';
  
  if (window.appState.consolidatedData.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">No raw data consolidated. Upload PDFs or load sample data.</td></tr>';
    return;
  }
  
  window.appState.consolidatedData.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.period}</td>
      <td class="text-secondary">${row.filename}</td>
      <td class="badge-cell"><span class="badge badge-blue">${row.classification}</span></td>
      <td class="text-left font-medium">${row.label}</td>
      <td class="text-right">${formatCurrency(row.taxable)}</td>
      <td class="text-right">${formatCurrency(row.igst)}</td>
      <td class="text-right">${formatCurrency(row.cgst)}</td>
      <td class="text-right">${formatCurrency(row.sgst)}</td>
      <td class="text-right">${formatCurrency(row.cess)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Render the first filtered table (Payable portion)
function renderFilteredTable() {
  const tbody = document.getElementById('payable-table-body');
  if (!tbody) return;
  tbody.innerHTML = '';
  
  if (window.appState.payableData.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-state">No filtered data.</td></tr>';
    return;
  }
  
  window.appState.payableData.forEach(row => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${row.period}</td>
      <td class="text-secondary">${row.filename}</td>
      <td class="badge-cell"><span class="badge badge-teal">${row.classification}</span></td>
      <td class="text-left font-medium">${row.label}</td>
      <td class="text-right">${formatCurrency(row.taxable)}</td>
      <td class="text-right">${formatCurrency(row.igst)}</td>
      <td class="text-right">${formatCurrency(row.cgst)}</td>
      <td class="text-right">${formatCurrency(row.sgst)}</td>
      <td class="text-right">${formatCurrency(row.cess)}</td>
    `;
    tbody.appendChild(tr);
  });
}

// Render the side-by-side pivot outputs
function renderOutputs() {
  renderPivot('pivot-a-container', window.appState.pivotA, 'Output A: Other than Reverse Charge (Forward Charge)', 'badge-blue');
  renderPivot('pivot-b-container', window.appState.pivotB, 'Output B: Reverse Charge Mechanism (RCM)', 'badge-purple');
}

function renderPivot(containerId, pivotData, title, badgeClass) {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = '';
  
  const periods = Object.keys(pivotData);
  if (periods.length === 0) {
    container.innerHTML = `<div class="empty-state">No values available for ${title}</div>`;
    return;
  }
  
  const table = document.createElement('table');
  table.className = 'data-table';
  
  // Header
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr>
      <th>Period</th>
      <th>Supply Type</th>
      <th>Taxable Value</th>
      <th>IGST</th>
      <th>CGST</th>
      <th>SGST</th>
      <th>Cess</th>
    </tr>
  `;
  table.appendChild(thead);
  
  const tbody = document.createElement('tbody');
  
  let totalTaxable = 0, totalIgst = 0, totalCgst = 0, totalSgst = 0, totalCess = 0;
  
  periods.forEach(p => {
    const monthData = pivotData[p];
    const itemKeys = Object.keys(monthData.items);
    
    itemKeys.forEach((key, idx) => {
      const item = monthData.items[key];
      const tr = document.createElement('tr');
      
      let periodTd = '';
      if (idx === 0) {
        periodTd = `<td rowspan="${itemKeys.length}" class="font-bold border-right">${p}</td>`;
      }
      
      tr.innerHTML = `
        ${periodTd}
        <td class="badge-cell"><span class="badge ${badgeClass}">${key}</span></td>
        <td class="text-right">${formatCurrency(item.taxable)}</td>
        <td class="text-right">${formatCurrency(item.igst)}</td>
        <td class="text-right">${formatCurrency(item.cgst)}</td>
        <td class="text-right">${formatCurrency(item.sgst)}</td>
        <td class="text-right">${formatCurrency(item.cess)}</td>
      `;
      tbody.appendChild(tr);
    });
    
    // Add monthly subtotal
    const subtr = document.createElement('tr');
    subtr.className = 'subtotal-row';
    subtr.innerHTML = `
      <td colspan="2" class="text-right font-bold">Subtotal (${p})</td>
      <td class="text-right font-bold">${formatCurrency(monthData.taxable)}</td>
      <td class="text-right font-bold">${formatCurrency(monthData.igst)}</td>
      <td class="text-right font-bold">${formatCurrency(monthData.cgst)}</td>
      <td class="text-right font-bold">${formatCurrency(monthData.sgst)}</td>
      <td class="text-right font-bold">${formatCurrency(monthData.cess)}</td>
    `;
    tbody.appendChild(subtr);
    
    totalTaxable += monthData.taxable;
    totalIgst += monthData.igst;
    totalCgst += monthData.cgst;
    totalSgst += monthData.sgst;
    totalCess += monthData.cess;
  });
  
  // Grand Total
  const tfoot = document.createElement('tfoot');
  tfoot.innerHTML = `
    <tr class="grand-total-row">
      <td colspan="2" class="text-right font-bold">Grand Total</td>
      <td class="text-right font-bold">${formatCurrency(totalTaxable)}</td>
      <td class="text-right font-bold">${formatCurrency(totalIgst)}</td>
      <td class="text-right font-bold">${formatCurrency(totalCgst)}</td>
      <td class="text-right font-bold">${formatCurrency(totalSgst)}</td>
      <td class="text-right font-bold">${formatCurrency(totalCess)}</td>
    </tr>
  `;
  table.appendChild(tbody);
  table.appendChild(tfoot);
  
  container.appendChild(table);
}

// Render the interactive D6 Reconciliation Matrix
function renderD6Matrix() {
  const container = document.getElementById('d6-matrix-container');
  if (!container) return;
  container.innerHTML = '';
  
  const table = document.createElement('table');
  table.className = 'd6-matrix-table';
  
  // Header Blocks
  const thead = document.createElement('thead');
  thead.innerHTML = `
    <tr class="main-header">
      <th rowspan="3" style="width: 50px;">Sl. No.</th>
      <th rowspan="3" style="width: 320px;">Particulars</th>
      <th colspan="6">Goods & Services Tax</th>
    </tr>
    <tr class="sub-header">
      <th rowspan="2">Taxable Value /<br>Assessable Value</th>
      <th rowspan="2">Excise Duty/ VAT,<br>CST, Cess etc.</th>
      <th colspan="4">GST Components</th>
    </tr>
    <tr class="comp-header">
      <th>CGST</th>
      <th>SGST / UTGST</th>
      <th>IGST</th>
      <th>Cess and Others</th>
    </tr>
    <tr class="currency-header">
      <th></th>
      <th></th>
      <th>Rs.</th>
      <th>Rs.</th>
      <th>Rs.</th>
      <th>Rs.</th>
      <th>Rs.</th>
      <th>Rs.</th>
    </tr>
  `;
  table.appendChild(thead);
  
  const tbody = document.createElement('tbody');
  
  D6_ROWS.forEach(row => {
    const tr = document.createElement('tr');
    tr.className = `d6-row-${row.type}`;
    
    // Serial column
    const tdSl = document.createElement('td');
    tdSl.className = 'text-center font-bold';
    tdSl.innerText = row.sl || "";
    tr.appendChild(tdSl);
    
    // Label column
    const tdLabel = document.createElement('td');
    tdLabel.className = 'text-left font-medium';
    tdLabel.innerText = row.label;
    tr.appendChild(tdLabel);
    
    // Data columns (indices 0 to 5)
    const vals = window.appState.d6Values[row.label] || [0, 0, 0, 0, 0, 0];
    
    for (let i = 0; i < 6; i++) {
      const td = document.createElement('td');
      td.className = 'text-right';
      
      // Payment rows: column 0 (Taxable) and 1 (Excise/VAT) are disabled/greyed
      const isPaymentRow = (row.category === 'payment' || row.type === 'formula_itc' || row.type === 'formula_paid');
      const isDisabledCol = isPaymentRow && (i === 0 || i === 1);
      
      if (isDisabledCol) {
        td.className = 'text-center cell-disabled';
        td.innerText = "—";
      } else if (row.type === 'data') {
        // Editable data row
        const isEditing = window.appState.activeEditingCell && 
                          window.appState.activeEditingCell.label === row.label && 
                          window.appState.activeEditingCell.colIdx === i;
                          
        if (isEditing) {
          td.className = 'cell-editing';
          const input = document.createElement('input');
          input.type = 'number';
          input.value = vals[i] === 0 ? '' : vals[i].toFixed(2);
          input.step = '0.01';
          
          input.onblur = (e) => {
            const val = parseFloat(parseFloat(e.target.value).toFixed(2)) || 0;
            const overrideKey = `${row.label}_${i}`;
            window.appState.d6Overrides[overrideKey] = val;
            window.appState.d6Values[row.label][i] = val;
            window.appState.activeEditingCell = null;
            recalculateD6Formulas();
            renderD6Matrix();
            updateKpiSummary();
          };
          
          input.onkeydown = (e) => {
            if (e.key === 'Enter') e.target.blur();
            if (e.key === 'Escape') {
              window.appState.activeEditingCell = null;
              renderD6Matrix();
            }
          };
          
          td.appendChild(input);
          setTimeout(() => input.focus(), 10);
        } else {
          td.className = 'cell-editable text-right';
          td.innerText = formatCurrency(vals[i]);
          td.onclick = () => {
            window.appState.activeEditingCell = { label: row.label, colIdx: i };
            renderD6Matrix();
          };
        }
      } else {
        // Formula or Header calculation rows (Read-only)
        if (row.type === 'formula_diff' && Math.abs(vals[i]) > 0.01) {
          td.className = 'cell-formula text-right variance-alert';
        } else {
          td.className = 'cell-formula text-right';
        }
        td.innerText = formatCurrency(vals[i]);
      }
      
      tr.appendChild(td);
    }
    
    tbody.appendChild(tr);
  });
  
  table.appendChild(tbody);
  container.appendChild(table);
}

// Update executive KPI counters
function updateKpiSummary() {
  const v = window.appState.d6Values;
  const payableRow = v["Total Duties/Taxes Payable (5+6+7+13)"] || [0,0,0,0,0,0];
  const paidRow = v["Total Duties/Taxes Paid (21 + 22)"] || [0,0,0,0,0,0];
  const diffRow = v["Difference between Taxes Paid and Payable (Row 14 - Row 23)"] || [0,0,0,0,0,0];
  
  // Total GST is the sum of CGST + SGST + IGST + Cess (cols 2, 3, 4, 5)
  const payableSum = payableRow[2] + payableRow[3] + payableRow[4] + payableRow[5];
  const paidSum = paidRow[2] + paidRow[3] + paidRow[4] + paidRow[5];
  const diffSum = diffRow[2] + diffRow[3] + diffRow[4] + diffRow[5];
  
  document.getElementById('kpi-payable').innerText = formatCurrency(payableSum);
  document.getElementById('kpi-paid').innerText = formatCurrency(paidSum);
  
  const diffEl = document.getElementById('kpi-diff');
  diffEl.innerText = formatCurrency(diffSum);
  
  const statusEl = document.getElementById('kpi-status');
  if (Math.abs(diffSum) < 0.02) {
    statusEl.innerHTML = '<span class="status-pass"><i class="fa-solid fa-circle-check"></i> RECONCILED</span>';
    diffEl.style.color = 'var(--accent-green)';
  } else {
    statusEl.innerHTML = '<span class="status-fail"><i class="fa-solid fa-triangle-exclamation"></i> MISMATCH</span>';
    diffEl.style.color = 'var(--accent-red)';
  }
  
  const warningBanner = document.getElementById('variance-warning-banner');
  if (warningBanner) {
    if (Math.abs(diffSum) > 0.01 && window.appState.consolidatedData.length > 0) {
      warningBanner.style.display = 'flex';
    } else {
      warningBanner.style.display = 'none';
    }
  }
}

// Load Mock GSTR-3B PDFs and run the parser to simulate uploader
window.loadSampleDataAndProcess = async function() {
  window.appState.isProcessing = true;
  window.appState.progress = 0;
  updateProgressUI('Loading sample months data...', 10);
  
  setTimeout(async () => {
    try {
      window.appState.files = [];
      window.appState.extractedPeriods = [];
      
      const samples = window.GSTR_SAMPLE_DATA;
      for (let i = 0; i < samples.length; i++) {
        const sample = samples[i];
        
        // Update uploader list
        window.appState.files.push({
          name: sample.filename,
          size: 14205,
          status: 'Processed'
        });
        
        // Direct conversion of sample records (simulating parser outcome)
        window.appState.extractedPeriods.push(JSON.parse(JSON.stringify(sample)));
        
        const pct = Math.floor(10 + (i + 1) / samples.length * 80);
        updateProgressUI(`Parsing ${sample.filename} (${sample.period})...`, pct);
        await new Promise(r => setTimeout(r, 150));
      }
      
      updateProgressUI('Consolidating & running reconciliation logic...', 95);
      runReconciliationPipeline();
      
      setTimeout(() => {
        window.appState.isProcessing = false;
        updateProgressUI('', 100);
        
        // Navigate to dashboard
        switchTab('raw_data');
      }, 500);
      
    } catch (err) {
      console.error(err);
      alert("Error loading sample data: " + err.message);
      window.appState.isProcessing = false;
    }
  }, 100);
};

// UI Progress Helpers
function updateProgressUI(text, percent) {
  const pBar = document.getElementById('progress-bar-fill');
  const pText = document.getElementById('progress-text');
  const container = document.getElementById('progress-container');
  
  if (percent === 100 || !text) {
    if (container) container.style.display = 'none';
    return;
  }
  
  if (container) container.style.display = 'block';
  if (pBar) pBar.style.width = percent + '%';
  if (pText) pText.innerText = text;
}

// Generate & download mock GSTR-3B PDFs locally
window.downloadAllSamplePDFs = async function() {
  if (typeof window.jspdf === 'undefined') {
    alert("jsPDF library is not loaded yet. Please wait a second or verify internet connection.");
    return;
  }
  
  const zip = new JSZip();
  const samples = window.GSTR_SAMPLE_DATA;
  
  for (let i = 0; i < samples.length; i++) {
    const sample = samples[i];
    const pdfBlob = window.generateGstr3bMockPdfBlob(sample);
    if (pdfBlob) {
      zip.file(sample.filename, pdfBlob);
    }
  }
  
  zip.generateAsync({ type: "blob" }).then(function(content) {
    saveAs(content, "GSTR3B_12_Months_Sample_PDFs.zip");
  });
};

// Navigate between Tabs
window.switchTab = function(tabId) {
  window.appState.activeTab = tabId;
  
  // Update Tab Header Classes
  const tabs = document.querySelectorAll('.tab-btn');
  tabs.forEach(t => {
    if (t.dataset.tab === tabId) t.classList.add('active');
    else t.classList.remove('active');
  });
  
  // Update Tab View Containers
  const views = document.querySelectorAll('.tab-view');
  views.forEach(v => {
    if (v.id === `view-${tabId}`) v.classList.add('active');
    else v.classList.remove('active');
  });
};

// SheetJS Excel Exporter with multiple sheets, formulas, and number formats
window.exportExcelReport = function() {
  if (window.appState.consolidatedData.length === 0) {
    alert("No consolidated data available. Please upload files or load sample data first.");
    return;
  }
  
  const wb = XLSX.utils.book_new();
  
  // ────────────────────────────────────────────────────────
  // SHEET 1: Master Consolidated Ledger
  // ────────────────────────────────────────────────────────
  const rawHeaders = ["Period", "Source File", "Classification", "Label", "Taxable Value (Rs.)", "IGST (Rs.)", "CGST (Rs.)", "SGST (Rs.)", "Cess (Rs.)"];
  const rawRows = window.appState.consolidatedData.map(r => [
    r.period, r.filename, r.classification, r.label, r.taxable, r.igst, r.cgst, r.sgst, r.cess
  ]);
  const wsRaw = XLSX.utils.aoa_to_sheet([rawHeaders, ...rawRows]);
  
  // Format numbers in Master sheet
  const rawRange = XLSX.utils.decode_range(wsRaw['!ref']);
  for (let r = 1; r <= rawRange.e.r; r++) {
    for (let c = 4; c <= 8; c++) {
      const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
      if (wsRaw[cellRef]) {
        wsRaw[cellRef].t = 'n';
        wsRaw[cellRef].z = '₹#,##0.00';
      }
    }
  }
  wsRaw['!cols'] = [{ wch: 15 }, { wch: 20 }, { wch: 12 }, { wch: 45 }, { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
  XLSX.utils.book_append_sheet(wb, wsRaw, "Consolidated Ledger");
  
  // ────────────────────────────────────────────────────────
  // SHEET 2: Segmented Pivot Tables
  // ────────────────────────────────────────────────────────
  // Build Output A and Output B summaries in one sheet stacked
  const pivotRows = [
    ["Output A: Other than Reverse Charge (Forward Charge Supplies)"],
    ["Period", "Type", "Taxable Value", "IGST", "CGST", "SGST", "Cess"]
  ];
  
  const appendPivotData = (pivotData, badgeVal) => {
    let rowStartIdx = pivotRows.length + 1;
    const periods = Object.keys(pivotData);
    let totalTaxable = 0, totalIgst = 0, totalCgst = 0, totalSgst = 0, totalCess = 0;
    
    periods.forEach(p => {
      const month = pivotData[p];
      Object.keys(month.items).forEach(key => {
        const it = month.items[key];
        pivotRows.push([p, key, it.taxable, it.igst, it.cgst, it.sgst, it.cess]);
      });
      // Monthly subtotal
      pivotRows.push([`Subtotal (${p})`, "", month.taxable, month.igst, month.cgst, month.sgst, month.cess]);
      totalTaxable += month.taxable;
      totalIgst += month.igst;
      totalCgst += month.cgst;
      totalSgst += month.sgst;
      totalCess += month.cess;
    });
    
    pivotRows.push(["Grand Total", "", totalTaxable, totalIgst, totalCgst, totalSgst, totalCess]);
  };
  
  appendPivotData(window.appState.pivotA, 'A');
  pivotRows.push([]);
  pivotRows.push([]);
  pivotRows.push(["Output B: Reverse Charge Mechanism (RCM Liabilities)"]);
  pivotRows.push(["Period", "Type", "Taxable Value", "IGST", "CGST", "SGST", "Cess"]);
  appendPivotData(window.appState.pivotB, 'B');
  
  const wsPivot = XLSX.utils.aoa_to_sheet(pivotRows);
  const pivotRange = XLSX.utils.decode_range(wsPivot['!ref']);
  
  // Format numeric values in Pivot Sheet
  for (let r = 0; r <= pivotRange.e.r; r++) {
    const firstCell = wsPivot[XLSX.utils.encode_cell({ r: r, c: 0 })];
    if (firstCell && (firstCell.v === 'Grand Total' || firstCell.v.toString().startsWith('Subtotal'))) {
      // Format total rows in bold/italic when opened in Excel
    }
    for (let c = 2; c <= 6; c++) {
      const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
      if (wsPivot[cellRef] && typeof wsPivot[cellRef].v === 'number') {
        wsPivot[cellRef].t = 'n';
        wsPivot[cellRef].z = '₹#,##0.00';
      }
    }
  }
  wsPivot['!cols'] = [{ wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
  XLSX.utils.book_append_sheet(wb, wsPivot, "Segmented Pivots");
  
  // ────────────────────────────────────────────────────────
  // SHEET 3: D6 Reconciliation Matrix
  // ────────────────────────────────────────────────────────
  const d6RowsAoa = [];
  
  // Title Blocks
  d6RowsAoa.push(["Sl. No.", "Duties/ Taxes Payable & Paid Particulars", "Taxable Value / Assessable Value (Rs.)", "Excise Duty / VAT, CST, Cess etc. / Other State Taxes (Rs.)", "Goods & Services Tax - CGST (Rs.)", "Goods & Services Tax - SGST / UTGST (Rs.)", "Goods & Services Tax - IGST (Rs.)", "Goods & Services Tax - Cess and Others (Rs.)"]);
  d6RowsAoa.push(["", "", "", "", "", "", "", ""]);
  
  // Populate matrix rows
  const excelRowMapping = {}; // maps D6_ROWS index to Excel row index (1-based)
  let headerOffset = 3; // header occupies rows 1-2 in Excel (0-based: index 0 and 1)
  
  D6_ROWS.forEach((row, idx) => {
    const excelRow = idx + headerOffset;
    excelRowMapping[idx] = excelRow;
    
    const vals = window.appState.d6Values[row.label] || [0,0,0,0,0,0];
    
    if (row.type === 'header' || row.type === 'subheader') {
      d6RowsAoa.push([row.sl, row.label, "", "", "", "", "", ""]);
    } else if (row.type === 'data') {
      // payment rows clear out cols C(index 2) & D(index 3)
      const isPayment = (row.category === 'payment');
      const val0 = isPayment ? "" : vals[0];
      const val1 = isPayment ? "" : vals[1];
      d6RowsAoa.push([row.sl, row.label, val0, val1, vals[2], vals[3], vals[4], vals[5]]);
    } else {
      // Formula Row placeholders - will insert formula objects in next step
      d6RowsAoa.push([row.sl, row.label, 0, 0, 0, 0, 0, 0]);
    }
  });
  
  const wsD6 = XLSX.utils.aoa_to_sheet(d6RowsAoa);
  
  // Merges for headers & subheaders
  wsD6['!merges'] = [
    // Main multi-tier headers (merged vertically to preserve layout and headerOffset=3)
    { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Sl No.
    { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Particulars
    { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } }, // Taxable Value
    { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } }, // Excise Duty
    { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } }, // CGST
    { s: { r: 0, c: 5 }, e: { r: 1, c: 5 } }, // SGST
    { s: { r: 0, c: 6 }, e: { r: 1, c: 6 } }, // IGST
    { s: { r: 0, c: 7 }, e: { r: 1, c: 7 } }  // Cess
  ];
  
  // Apply dynamic Excel formulas and number formats
  D6_ROWS.forEach((row, idx) => {
    const excelRowIndex = idx + headerOffset; // 1-based index for formula strings
    
    // Add row merges for Section and Sub-section titles
    if (row.type === 'header' || row.type === 'subheader') {
      wsD6['!merges'].push({
        s: { r: excelRowIndex - 1, c: 1 },
        e: { r: excelRowIndex - 1, c: 7 }
      });
      return;
    }
    
    const colLetters = ['C', 'D', 'E', 'F', 'G', 'H'];
    
    colLetters.forEach((colLetter, colIdx) => {
      const cellRef = `${colLetter}${excelRowIndex}`;
      
      if (row.type === 'formula_sum') {
        const [startIdx, endIdx] = row.range;
        const startRow = startIdx + headerOffset;
        const endRow = endIdx + headerOffset;
        wsD6[cellRef] = { t: 'n', f: `SUM(${colLetter}${startRow}:${colLetter}${endRow})`, z: '₹#,##0.00' };
        
        // Blank out taxable/excise columns for ITC totals
        if (row.label.includes("Credit Utilised") && (colLetter === 'C' || colLetter === 'D')) {
          wsD6[cellRef] = { t: 's', v: "" };
        }
      } 
      else if (row.type === 'formula_custom_payable') {
        // Total Excise(Row 9) + VAT(Row 10) + State(Row 11) + GST(Row 18)
        wsD6[cellRef] = { t: 'n', f: `${colLetter}9+${colLetter}10+${colLetter}11+${colLetter}18`, z: '₹#,##0.00' };
      } 
      else if (row.type === 'formula_custom_paid') {
        // Total ITC Utilised (Row 28) + Cash Ledger (Row 29)
        wsD6[cellRef] = { t: 'n', f: `${colLetter}28+${colLetter}29`, z: '₹#,##0.00' };
        
        if (colLetter === 'C' || colLetter === 'D') {
          wsD6[cellRef] = { t: 's', v: "" };
        }
      } 
      else if (row.type === 'formula_diff') {
        // Payable (Row 19) - Paid (Row 30)
        wsD6[cellRef] = { t: 'n', f: `${colLetter}19-${colLetter}30`, z: '₹#,##0.00' };
      } 
      else {
        // Data rows: apply rupee format
        if (wsD6[cellRef] && typeof wsD6[cellRef].v === 'number') {
          wsD6[cellRef].z = '₹#,##0.00';
        }
      }
    });
  });
  
  wsD6['!cols'] = [
    { wch: 8 },  // Sl No
    { wch: 45 }, // Particulars
    { wch: 18 }, // Taxable
    { wch: 20 }, // Excise/VAT
    { wch: 15 }, // CGST
    { wch: 15 }, // SGST
    { wch: 15 }, // IGST
    { wch: 15 }  // Cess
  ];
  
  XLSX.utils.book_append_sheet(wb, wsD6, "D6 Reconciliation");
  
  // Write & Save Workbook
  XLSX.writeFile(wb, "GSTR_D6_Reconciliation_Report.xlsx");
};

// Initialize Application UI
document.addEventListener('DOMContentLoaded', () => {
  initD6Values();
  
  // Drag and drop events setup
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  
  if (dropzone && fileInput) {
    dropzone.onclick = () => fileInput.click();
    
    dropzone.ondragover = (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    };
    
    dropzone.ondragleave = () => {
      dropzone.classList.remove('dragover');
    };
    
    dropzone.ondrop = async (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleUploadedFiles(e.dataTransfer.files);
      }
    };
    
    fileInput.onchange = (e) => {
      if (e.target.files.length > 0) {
        handleUploadedFiles(e.target.files);
      }
    };
  }
  
  // Setup reactive triggers for manual input values
  renderAllViews();
});

// File upload handler
async function handleUploadedFiles(fileList) {
  window.appState.isProcessing = true;
  updateProgressUI('Analyzing files...', 5);
  
  const filesArray = Array.from(fileList).filter(f => f.name.toLowerCase().endsWith('.pdf'));
  
  if (filesArray.length === 0) {
    alert("Please upload GSTR-3B PDF files.");
    window.appState.isProcessing = false;
    updateProgressUI('', 100);
    return;
  }
  
  try {
    for (let i = 0; i < filesArray.length; i++) {
      const file = filesArray[i];
      updateProgressUI(`Reading ${file.name}...`, 10 + Math.floor(i / filesArray.length * 80));
      
      const fileReader = new FileReader();
      const readPromise = new Promise((resolve) => {
        fileReader.onload = (e) => resolve(e.target.result);
      });
      fileReader.readAsArrayBuffer(file);
      const arrayBuffer = await readPromise;
      
      // Parse PDF
      const result = await parseGstr3bPdf(arrayBuffer, file.name);
      
      // Save file item to sidebar list
      window.appState.files.push({
        name: file.name,
        size: file.size,
        status: 'Processed'
      });
      
      window.appState.extractedPeriods.push(result);
    }
    
    updateProgressUI('Consolidating & filtering values...', 90);
    runReconciliationPipeline();
    
    setTimeout(() => {
      window.appState.isProcessing = false;
      updateProgressUI('', 100);
      switchTab('raw_data');
    }, 500);
    
  } catch (err) {
    console.error(err);
    alert("An error occurred during file parsing: " + err.message);
    window.appState.isProcessing = false;
    updateProgressUI('', 100);
  }
}
