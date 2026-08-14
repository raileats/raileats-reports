import * as XLSX from 'xlsx-js-style';

// Helper date formatter: "Saturday, 1 August, 2026"
const formatDisplayDate = (dateStr: string): string => {
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString('en-GB', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
};

// Clean source classifier from order data
const getOrderSource = (row: any): string => {
  const channel = String(row['Source'] || row['Channel'] || row['Booking Channel'] || '').toUpperCase();
  const orderId = String(row['IRCTC Order ID'] || row['Order ID'] || '').toUpperCase();

  if (channel.includes('MMT') || channel.includes('MAKEMYTRIP')) return 'MakeMyTrip';
  if (channel.includes('APP') || channel.includes('REL_APP')) return 'REL_Food_App';
  if (channel.includes('WEB') || channel.includes('WEBSITE')) return 'RELFood_WEBSITE';
  if (channel.includes('IRCTC') || orderId.startsWith('IR') || !channel) return 'RELFood_IRCTC';
  return 'RELFood_IRCTC';
};

interface MetricStats {
  orders: number;
  deliveredOrders: number;
  cancelledOrders: number;
  meals: number;
  value: number;
  prepaidValue: number;
  discount: number;
  revenue: number;
  complaints: number;
  feedback: number;
  undelivered: number;
  outletsSet: Set<string>;
}

const createEmptyStats = (): MetricStats => ({
  orders: 0,
  deliveredOrders: 0,
  cancelledOrders: 0,
  meals: 0,
  value: 0,
  prepaidValue: 0,
  discount: 0,
  revenue: 0,
  complaints: 0,
  feedback: 0,
  undelivered: 0,
  outletsSet: new Set<string>(),
});

const normalizeOrderId = (value: any): string => {
  if (value === null || value === undefined || value === '') return '';
  return String(value).trim().replace(/\u00A0/g, '').replace(/\s+/g, '').replace(/\.0+$/, '').toUpperCase();
};

const cleanOutletId = (value: any): string => {
  if (value === null || value === undefined || String(value).trim() === '') return '';
  return String(value)
    .trim()
    .replace(/\u00A0/g, '')
    .replace(/\s+/g, '')
    .replace(/\.0+$/, '');
};

const getOutletIdFromRow = (row: any): string => {
  if (!row) return '';
  const candidates = [
    row['Outlet ID'],
    row['Outlet Id'],
    row['OutletId'],
    row['OutletID'],
    row['Aggregator Outlet ID'],
    row['Aggregator Outlet Id'],
    row['Aggregator OutletId'],
    row['Aggregator OutletID'],
    row['Outlet Code'],
    row['Outlet'],
  ];
  for (const value of candidates) {
    const id = cleanOutletId(value);
    if (id) return id;
  }
  return '';
};

const normalizeFeedbackType = (value: any): 'FEEDBACK' | 'COMPLAINT' | '' => {
  const type = String(value ?? '').trim().toUpperCase();
  if (type === 'FEEDBACK') return 'FEEDBACK';
  if (type === 'COMPLAINT' || type === 'COMPLAIN') return 'COMPLAINT';
  return '';
};

const feedbackCreatedAtKey = (value: any): string => {
  if (value === null || value === undefined || String(value).trim() === '') return '';
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return '';
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
  }
  const raw = String(value).trim();
  const numeric = Number(raw);
  if (Number.isFinite(numeric) && numeric > 30000 && numeric < 70000) {
    const d = new Date(Date.UTC(1899, 11, 30) + numeric * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  // Feedback CSV uses values like `2026-08-01 18:52:14 IST`.
  // Parse the leading ISO date explicitly because native Date parsing can
  // reject the trailing `IST` timezone text.
  const isoMatch = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }

  const datePart = raw.split(/[T ]/)[0];
  let m = datePart.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) {
    const first = Number(m[1]);
    const second = Number(m[2]);
    const year = Number(m[3]);
    const month = first <= 12 ? first : second;
    const day = first <= 12 ? second : first;
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) {
      return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    }
  }
  const d = new Date(raw);
  if (isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

export const generateMainReportWorkbook = (masterData: any[], feedbackData: any[] = []) => {
  if (!masterData || masterData.length === 0) {
    alert('Koi data uplabdh nahi hai!');
    return;
  }

  // 1. Group order records by Delivery/Booking Date + Source.
  // Feedback/Complaint counts are kept in a separate map because their FTD
  // date must come from Feedback upload's `Created At`.
  const dateGroups: Record<string, Record<string, any[]>> = {};
  const sourcesList = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];
  const orderSourceMap = new Map<string, string>();
  const feedbackByDateSource: Record<string, Record<string, { complaint: number; feedback: number }>> = {};

  masterData.forEach((row) => {
    const rawDate = row['Delivery Date'] || row['Booking Date'] || 'Unknown Date';
    const dateKey = String(rawDate).split(' ')[0].split('T')[0];
    const src = getOrderSource(row);

    if (!dateGroups[dateKey]) {
      dateGroups[dateKey] = {
        RELFood_IRCTC: [],
        RELFood_WEBSITE: [],
        REL_Food_App: [],
        MakeMyTrip: [],
      };
    }
    if (!dateGroups[dateKey][src]) dateGroups[dateKey][src] = [];
    dateGroups[dateKey][src].push(row);

    const orderId = normalizeOrderId(row['IRCTC Order ID'] || row['Order ID'] || row['Order Id']);
    if (orderId) orderSourceMap.set(orderId, src);
  });

  feedbackData.forEach((row) => {
    const createdAt = row['Created At'] ?? row['CreatedAt'] ?? row['created_at'];
    const dateKey = feedbackCreatedAtKey(createdAt);
    const type = normalizeFeedbackType(row['Type'] ?? row['Feedback Type'] ?? row['FeedbackType']);
    const orderId = normalizeOrderId(row['Order ID'] ?? row['Order Id'] ?? row['OrderID'] ?? row['IRCTC Order ID']);
    if (!dateKey || !type) return;

    const src = orderSourceMap.get(orderId) || 'RELFood_IRCTC';
    if (!feedbackByDateSource[dateKey]) feedbackByDateSource[dateKey] = {};
    if (!feedbackByDateSource[dateKey][src]) feedbackByDateSource[dateKey][src] = { complaint: 0, feedback: 0 };
    if (type === 'COMPLAINT') feedbackByDateSource[dateKey][src].complaint += 1;
    if (type === 'FEEDBACK') feedbackByDateSource[dateKey][src].feedback += 1;

    if (!dateGroups[dateKey]) {
      dateGroups[dateKey] = { RELFood_IRCTC: [], RELFood_WEBSITE: [], REL_Food_App: [], MakeMyTrip: [] };
    }
  });

  // Sort dates chronologically.
  const sortedDates = Object.keys(dateGroups).sort((a, b) => a.localeCompare(b));

  // Matrix construction for Excel
  const excelRows: any[][] = [];
  const merges: XLSX.Range[] = [];

  // MTD Trackers across consecutive dates in month
  const mtdCumulativeBySource: Record<string, MetricStats> = {};
  sourcesList.forEach((s) => (mtdCumulativeBySource[s] = createEmptyStats()));
  let mtdGrandTotal = createEmptyStats();

  let currentRowIdx = 0;

  sortedDates.forEach((dateKey) => {
    const displayDate = formatDisplayDate(dateKey);
    const daySources = dateGroups[dateKey];

    // Day FTD Stats
    const dayStatsBySource: Record<string, MetricStats> = {};
    const dayTotalStats = createEmptyStats();

    sourcesList.forEach((s) => {
      const rows = daySources[s] || [];
      const stats = createEmptyStats();

      rows.forEach((r) => {
        const isDelivered = r['Final Status'] === 'Delivered';
        const isUndelivered = r['Final Status'] === 'Not Delivered' || String(r['IRCTC Status'] || '').toUpperCase().includes('UNDELIVERED');
        const sellingPrice = parseFloat(r['Final Selling Price'] || 0) || 0;
        const discount = parseFloat(r['Final Total Discount'] || 0) || 0;
        const rfComm = parseFloat(r['Final RF Commission'] || 0) || 0;
        const prepaid = parseFloat(r['PPD'] || 0) || 0;
        const mealCount = parseInt(r['Meals'] || '1', 10) || 1;
        const outletId = String(r['Outlet ID'] || '').trim();

        stats.orders += 1;
        if (isDelivered) stats.deliveredOrders += 1;
        if (r['Final Status'] === 'Cancelled') stats.cancelledOrders += 1;
        if (isUndelivered) stats.undelivered += 1;

        stats.meals += mealCount;
        stats.value += sellingPrice;
        stats.prepaidValue += prepaid;
        stats.discount += discount;
        stats.revenue += rfComm;

        // Feedback/Complaint is NOT inferred from Rating/Remarks.
        // It is added below from Feedback.Created At for this report date.
        if (isDelivered && outletId) stats.outletsSet.add(outletId);
      });

      // Feedback/Complaint FTD is grouped by Feedback.Created At.
      const feedbackDay = feedbackByDateSource[dateKey] || {};
      const feedbackForSource = feedbackDay[s];
      if (feedbackForSource) {
        stats.feedback += feedbackForSource.feedback;
        stats.complaints += feedbackForSource.complaint;
      }

      dayStatsBySource[s] = stats;

      // Accumulate Day Total
      dayTotalStats.orders += stats.orders;
      dayTotalStats.deliveredOrders += stats.deliveredOrders;
      dayTotalStats.cancelledOrders += stats.cancelledOrders;
      dayTotalStats.meals += stats.meals;
      dayTotalStats.value += stats.value;
      dayTotalStats.prepaidValue += stats.prepaidValue;
      dayTotalStats.discount += stats.discount;
      dayTotalStats.revenue += stats.revenue;
      dayTotalStats.complaints += stats.complaints;
      dayTotalStats.feedback += stats.feedback;
      dayTotalStats.undelivered += stats.undelivered;
      stats.outletsSet.forEach((o) => dayTotalStats.outletsSet.add(o));

      // Update MTD Accumulators
      const mtd = mtdCumulativeBySource[s];
      mtd.orders += stats.orders;
      mtd.deliveredOrders += stats.deliveredOrders;
      mtd.meals += stats.meals;
      mtd.value += stats.value;
      mtd.prepaidValue += stats.prepaidValue;
      mtd.discount += stats.discount;
      mtd.revenue += stats.revenue;
      mtd.complaints += stats.complaints;
      mtd.feedback += stats.feedback;
      mtd.undelivered += stats.undelivered;
    });

    // Update MTD Grand Total
    mtdGrandTotal.orders += dayTotalStats.orders;
    mtdGrandTotal.deliveredOrders += dayTotalStats.deliveredOrders;
    mtdGrandTotal.meals += dayTotalStats.meals;
    mtdGrandTotal.value += dayTotalStats.value;
    mtdGrandTotal.prepaidValue += dayTotalStats.prepaidValue;
    mtdGrandTotal.discount += dayTotalStats.discount;
    mtdGrandTotal.revenue += dayTotalStats.revenue;
    mtdGrandTotal.complaints += dayTotalStats.complaints;
    mtdGrandTotal.feedback += dayTotalStats.feedback;
    mtdGrandTotal.undelivered += dayTotalStats.undelivered;

    // --- ROW 1: RED FULL BANNER DATE HEADER ---
    const dateRow = new Array(39).fill('');
    dateRow[0] = displayDate;
    excelRows.push(dateRow);
    merges.push({ s: { r: currentRowIdx, c: 0 }, e: { r: currentRowIdx, c: 38 } });
    currentRowIdx++;

    // --- ROW 2: GROUP HEADERS ---
    const groupHeaderRow = [
      'Source',
      'ORDERS', '', '', '', '',
      'MEALS', '', '', '', '',
      'VALUE', '', '',
      'PREPAID', '', '', '',
      'DISCOUNT', '', '', '',
      'REVENUE', '', '', '',
      'Complaints', '', '', '',
      'Feedback', '', '', '',
      'IRCTC Undelivered', '', '', '',
      'Outlets',
    ];
    excelRows.push(groupHeaderRow);

    // Group Header Merges
    merges.push({ s: { r: currentRowIdx, c: 1 }, e: { r: currentRowIdx, c: 5 } }); // ORDERS
    merges.push({ s: { r: currentRowIdx, c: 6 }, e: { r: currentRowIdx, c: 10 } }); // MEALS
    merges.push({ s: { r: currentRowIdx, c: 11 }, e: { r: currentRowIdx, c: 13 } }); // VALUE
    merges.push({ s: { r: currentRowIdx, c: 14 }, e: { r: currentRowIdx, c: 17 } }); // PREPAID
    merges.push({ s: { r: currentRowIdx, c: 18 }, e: { r: currentRowIdx, c: 21 } }); // DISCOUNT
    merges.push({ s: { r: currentRowIdx, c: 22 }, e: { r: currentRowIdx, c: 25 } }); // REVENUE
    merges.push({ s: { r: currentRowIdx, c: 26 }, e: { r: currentRowIdx, c: 29 } }); // Complaints
    merges.push({ s: { r: currentRowIdx, c: 30 }, e: { r: currentRowIdx, c: 33 } }); // Feedback
    merges.push({ s: { r: currentRowIdx, c: 34 }, e: { r: currentRowIdx, c: 37 } }); // IRCTC Undelivered
    currentRowIdx++;

    // --- ROW 3: SUB-COLUMN HEADERS ---
    const subColRow = [
      '',
      // ORDERS (5)
      'FTD', 'MTD', 'LMTD', 'ASP', 'Del%',
      // MEALS (5)
      'FTD', 'MTD', 'LMTD', 'ASP', 'MPO',
      // VALUE (3)
      'FTD', 'MTD', 'LMTD',
      // PREPAID (4)
      'FTD', 'MTD', 'LMTD', '%',
      // DISCOUNT (4)
      'FTD', 'MTD', 'LMTD', '%',
      // REVENUE (4)
      'FTD', 'MTD', 'LMTD', '%',
      // Complaints (4)
      'FTD', 'MTD', 'LMTD', '%',
      // Feedback (4)
      'FTD', 'MTD', 'LMTD', '%',
      // IRCTC Undelivered (4)
      'FTD', 'MTD', 'LMTD', '%',
      // Outlets
      '',
    ];
    excelRows.push(subColRow);
    currentRowIdx++;

    // Helper for rendering row data
    const formatRowData = (label: string, ftd: MetricStats, mtd: MetricStats, outletsCount: number | string) => {
      const orderAsp = ftd.orders > 0 ? Math.round(ftd.value / ftd.orders) : 0;
      const delPct = ftd.orders > 0 ? `${((ftd.deliveredOrders / ftd.orders) * 100).toFixed(2)}%` : '0.00%';
      const mealAsp = ftd.meals > 0 ? Math.round(ftd.value / ftd.meals) : 0;
      const mpo = ftd.orders > 0 ? (ftd.meals / ftd.orders).toFixed(2) : '0.00';
      const prepaidPct = ftd.value > 0 ? `${((ftd.prepaidValue / ftd.value) * 100).toFixed(2)}%` : '0.00%';
      const discountPct = ftd.value > 0 ? `${((ftd.discount / ftd.value) * 100).toFixed(2)}%` : '0.00%';
      const revenuePct = ftd.value > 0 ? `${((ftd.revenue / ftd.value) * 100).toFixed(2)}%` : '0.00%';
      const complaintPct = ftd.deliveredOrders > 0 ? `${((ftd.complaints / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
      const feedbackPct = ftd.deliveredOrders > 0 ? `${((ftd.feedback / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
      const undeliveredPct = ftd.orders > 0 ? `${((ftd.undelivered / ftd.orders) * 100).toFixed(2)}%` : '0.00%';

      return [
        label,
        // ORDERS
        ftd.orders, mtd.orders, 0, orderAsp, delPct,
        // MEALS
        ftd.meals, mtd.meals, 0, mealAsp, mpo,
        // VALUE
        Math.round(ftd.value), Math.round(mtd.value), 0,
        // PREPAID
        Math.round(ftd.prepaidValue), Math.round(mtd.prepaidValue), 0, prepaidPct,
        // DISCOUNT
        Math.round(ftd.discount), Math.round(mtd.discount), 0, discountPct,
        // REVENUE
        Math.round(ftd.revenue), Math.round(mtd.revenue), 0, revenuePct,
        // Complaints
        ftd.complaints, mtd.complaints, 0, complaintPct,
        // Feedback
        ftd.feedback, mtd.feedback, 0, feedbackPct,
        // IRCTC Undelivered
        ftd.undelivered, mtd.undelivered, 0, undeliveredPct,
        // Outlets
        outletsCount,
      ];
    };

    // --- ROW 4: TOTAL ROW ---
    excelRows.push(formatRowData('Total', dayTotalStats, mtdGrandTotal, dayTotalStats.outletsSet.size));
    currentRowIdx++;

    // --- ROWS 5-8: SOURCE ROWS ---
    sourcesList.forEach((src) => {
      excelRows.push(
        formatRowData(
          src,
          dayStatsBySource[src] || createEmptyStats(),
          mtdCumulativeBySource[src] || createEmptyStats(),
          dayStatsBySource[src]?.outletsSet.size || 0
        )
      );
      currentRowIdx++;
    });

    // Merge Outlets count across all data rows for this date block
    merges.push({ s: { r: currentRowIdx - 5, c: 38 }, e: { r: currentRowIdx - 1, c: 38 } });

    // Empty separator row between dates
    excelRows.push(new Array(39).fill(''));
    currentRowIdx++;
  });

  // Create Sheet
  const ws = XLSX.utils.aoa_to_sheet(excelRows);
  ws['!merges'] = merges;

  // Match the dashboard's Excel-style color formatting in the downloaded XLSX.
  const fill = (rgb: string) => ({ patternType: 'solid', fgColor: { rgb } });
  const border = {
    top: { style: 'thin', color: { rgb: '222222' } },
    bottom: { style: 'thin', color: { rgb: '222222' } },
    left: { style: 'thin', color: { rgb: '222222' } },
    right: { style: 'thin', color: { rgb: '222222' } },
  };
  const groupFills: Record<number, string> = {
    0: '000000', 1: '72B9D2', 6: 'B5D37B', 11: 'F3B37E', 14: '8DC5DE',
    18: 'D99397', 22: '999999', 26: '5B91C8', 30: '87C34F', 34: '555555', 38: 'F2A900'
  };
  const subFills: Record<number, string> = {
    0: '000000', 1: 'B9DCE8', 6: 'D2E7BF', 11: 'F8D8C0', 14: 'C4E0EC',
    18: 'EFC7CA', 22: 'CFCFCF', 26: 'C4D9EF', 30: 'D4E8C5', 34: '666666', 38: 'F2A900'
  };
  const dataFills = ['CCE5EE','CCE5EE','CCE5EE','CCE5EE','CCE5EE','D9E8C8','D9E8C8','D9E8C8','D9E8C8','D9E8C8','F8DFCA','F8DFCA','F8DFCA','D6EAF2','D6EAF2','D6EAF2','D6EAF2','F0D6D9','F0D6D9','F0D6D9','F0D6D9','D0D0D0','D0D0D0','D0D0D0','D0D0D0','C8DDF0','C8DDF0','C8DDF0','C8DDF0','D9EACB','D9EACB','D9EACB','D9EACB','666666','666666','666666','666666'];
  const rowCount = excelRows.length;
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < 39; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      const cell = ws[ref];
      if (!cell) continue;
      const isDate = excelRows[r][0] && c === 0 && merges.some(m => m.s.r === r && m.s.c === 0 && m.e.c === 38);
      const isGroup = r >= 0 && excelRows[r][c] && r > 0 && excelRows[r - 1]?.[0] && merges.some(m => m.s.r === r && m.s.c === c);
      const style: any = { alignment: { horizontal: 'center', vertical: 'center', wrapText: false }, border };
      if (isDate) { style.fill = fill('FF0000'); style.font = { bold: true, color: { rgb: 'FFFFFF' }, sz: 14 }; }
      else if (r > 0 && c in groupFills && [1,6,11,14,18,22,26,30,34,38].includes(c)) {
        style.fill = fill(groupFills[c]); style.font = { bold: true, color: { rgb: c === 0 || c === 34 ? 'FFFFFF' : '000000' }, sz: 10 };
      } else if (r > 0 && r < rowCount && excelRows[r][c] && r > 0 && Object.values(subFills).length) {
        // Identify the sub-header rows by their known labels.
        const labels = ['FTD','MTD','LMTD','ASP','Del%','MPO','%'];
        if (labels.includes(String(excelRows[r][c]))) {
          let start = c; while (start > 0 && !subFills[start]) start--;
          style.fill = fill(subFills[start] || 'DCE7EC'); style.font = { bold: true, color: { rgb: start === 34 ? 'FFFFFF' : '000000' }, sz: 9 };
        } else if (c === 0) {
          style.fill = fill('000000'); style.font = { bold: true, color: { rgb: 'FFFFFF' }, sz: 9 };
        } else {
          style.fill = fill(dataFills[c - 1] || 'EDF4F7'); style.font = { sz: 9, bold: r % 7 === 3 };
        }
      } else if (c === 0) {
        style.fill = fill(r % 7 === 3 ? '000000' : 'EF0000'); style.font = { bold: true, color: { rgb: 'FFFFFF' }, sz: 9 };
      } else {
        style.fill = fill(dataFills[c - 1] || 'EDF4F7'); style.font = { sz: 9, bold: r % 7 === 3 };
      }
      if (c === 38) { style.fill = fill('FFF200'); style.font = { bold: true, color: { rgb: '111111' }, sz: 10 }; }
      ws[ref].s = style;
    }
  }
  ws['!rows'] = excelRows.map((_, r) => ({ hpt: r === 0 ? 24 : 18 }));

  // Set Auto Column Widths
  ws['!cols'] = [
    { wch: 18 }, // Source
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, // ORDERS
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, // MEALS
    { wch: 10 }, { wch: 11 }, { wch: 8 }, // VALUE
    { wch: 10 }, { wch: 11 }, { wch: 8 }, { wch: 9 }, // PREPAID
    { wch: 9 }, { wch: 9 }, { wch: 8 }, { wch: 9 }, // DISCOUNT
    { wch: 9 }, { wch: 10 }, { wch: 8 }, { wch: 9 }, // REVENUE
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, // Complaints
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, // Feedback
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, // IRCTC Undelivered
    { wch: 10 }, // Outlets
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Main Report Date Wise');
  XLSX.writeFile(wb, `Main_Report_Date_Wise_${new Date().toISOString().slice(0, 10)}.xlsx`);
};
