import * as XLSX from 'xlsx';

export interface ReportDayBlock {
  dateLabel: string;
  rawDate: string;
  outletsCount: number;
  dayTotal: RowMetrics;
  mtdTotal: RowMetrics;
  bySource: Record<string, { ftd: RowMetrics; mtd: RowMetrics }>;
}

export interface RowMetrics {
  orders: number;
  deliveredOrders: number;
  meals: number;
  value: number;
  prepaidValue: number;
  discount: number;
  revenue: number;
  complaints: number;
  feedback: number;
  undelivered: number;
}

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

const createEmptyMetrics = (): RowMetrics => ({
  orders: 0,
  deliveredOrders: 0,
  meals: 0,
  value: 0,
  prepaidValue: 0,
  discount: 0,
  revenue: 0,
  complaints: 0,
  feedback: 0,
  undelivered: 0,
});

const getSource = (row: any): string => {
  const channel = String(row['Source'] || row['Channel'] || row['Booking Channel'] || '').toUpperCase();
  const orderId = String(row['IRCTC Order ID'] || row['Order ID'] || '').toUpperCase();

  if (channel.includes('MMT') || channel.includes('MAKEMYTRIP')) return 'MakeMyTrip';
  if (channel.includes('APP') || channel.includes('REL_APP')) return 'REL_Food_App';
  if (channel.includes('WEB') || channel.includes('WEBSITE')) return 'RELFood_WEBSITE';
  return 'RELFood_IRCTC';
};

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

// 1. Core Data Transformer: Raw Master Orders -> Date-wise Aggregated Matrix
export const processMainReportData = (masterData: any[]): ReportDayBlock[] => {
  if (!masterData || masterData.length === 0) return [];

  // Group by Date
  const dateMap: Record<string, any[]> = {};
  masterData.forEach((row) => {
    const rawDate = row['Delivery Date'] || row['Booking Date'] || 'Unknown Date';
    const dateKey = String(rawDate).split(' ')[0].split('T')[0];
    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(row);
  });

  const sortedDates = Object.keys(dateMap).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

  // Running MTD Accumulators
  const mtdCumulativeBySource: Record<string, RowMetrics> = {};
  SOURCES.forEach((s) => (mtdCumulativeBySource[s] = createEmptyMetrics()));
  const mtdGrandTotal = createEmptyMetrics();

  return sortedDates.map((dateKey) => {
    const rows = dateMap[dateKey];
    const outletsSet = new Set<string>();

    const dayStatsBySource: Record<string, RowMetrics> = {};
    SOURCES.forEach((s) => (dayStatsBySource[s] = createEmptyMetrics()));
    const dayTotal = createEmptyMetrics();

    rows.forEach((r) => {
      const src = getSource(r);
      const isDelivered = r['Final Status'] === 'Delivered';
      const isUndelivered = r['Final Status'] === 'Not Delivered' || String(r['IRCTC Status'] || '').toUpperCase().includes('UNDELIVERED');
      const sellingPrice = parseFloat(r['Final Selling Price'] || r['Selling Price'] || 0) || 0;
      const discount = parseFloat(r['Final Total Discount'] || r['Total Discount'] || 0) || 0;
      const rfComm = parseFloat(r['Final RF Commission'] || r['RF Comm'] || 0) || 0;
      const prepaid = parseFloat(r['PPD'] || r['Prepaid'] || 0) || 0;
      const mealCount = parseInt(r['Meals'] || '1', 10) || 1;
      const outletId = String(r['Outlet ID'] || '').trim();

      const st = dayStatsBySource[src] || dayStatsBySource['RELFood_IRCTC'];
      st.orders += 1;
      if (isDelivered) st.deliveredOrders += 1;
      if (isUndelivered) st.undelivered += 1;
      st.meals += mealCount;
      st.value += sellingPrice;
      st.prepaidValue += prepaid;
      st.discount += discount;
      st.revenue += rfComm;

      if (r['Rating'] && parseFloat(r['Rating']) > 0) st.feedback += 1;
      if (r['Remarks'] && String(r['Remarks']).toLowerCase().includes('complaint')) st.complaints += 1;
      if (outletId) outletsSet.add(outletId);
    });

    // Accumulate to Day Total & MTD
    const bySourcePayload: Record<string, { ftd: RowMetrics; mtd: RowMetrics }> = {};

    SOURCES.forEach((s) => {
      const st = dayStatsBySource[s];
      dayTotal.orders += st.orders;
      dayTotal.deliveredOrders += st.deliveredOrders;
      dayTotal.meals += st.meals;
      dayTotal.value += st.value;
      dayTotal.prepaidValue += st.prepaidValue;
      dayTotal.discount += st.discount;
      dayTotal.revenue += st.revenue;
      dayTotal.complaints += st.complaints;
      dayTotal.feedback += st.feedback;
      dayTotal.undelivered += st.undelivered;

      const m = mtdCumulativeBySource[s];
      m.orders += st.orders;
      m.deliveredOrders += st.deliveredOrders;
      m.meals += st.meals;
      m.value += st.value;
      m.prepaidValue += st.prepaidValue;
      m.discount += st.discount;
      m.revenue += st.revenue;
      m.complaints += st.complaints;
      m.feedback += st.feedback;
      m.undelivered += st.undelivered;

      bySourcePayload[s] = {
        ftd: { ...st },
        mtd: { ...m },
      };
    });

    mtdGrandTotal.orders += dayTotal.orders;
    mtdGrandTotal.deliveredOrders += dayTotal.deliveredOrders;
    mtdGrandTotal.meals += dayTotal.meals;
    mtdGrandTotal.value += dayTotal.value;
    mtdGrandTotal.prepaidValue += dayTotal.prepaidValue;
    mtdGrandTotal.discount += dayTotal.discount;
    mtdGrandTotal.revenue += dayTotal.revenue;
    mtdGrandTotal.complaints += dayTotal.complaints;
    mtdGrandTotal.feedback += dayTotal.feedback;
    mtdGrandTotal.undelivered += dayTotal.undelivered;

    return {
      dateLabel: formatDisplayDate(dateKey),
      rawDate: dateKey,
      outletsCount: outletsSet.size,
      dayTotal: { ...dayTotal },
      mtdTotal: { ...mtdGrandTotal },
      bySource: bySourcePayload,
    };
  });
};

// 2. Export Date-Wise Formatted Excel (Matches Screenshot Exactly)
export const exportMainReportToExcel = (masterData: any[], fileName?: string) => {
  const blocks = processMainReportData(masterData);
  if (blocks.length === 0) return;

  const excelRows: any[][] = [];
  const merges: XLSX.Range[] = [];
  let rIdx = 0;

  const buildRow = (label: string, ftd: RowMetrics, mtd: RowMetrics, outlets: string | number = '') => {
    const orderAsp = ftd.orders > 0 ? Math.round(ftd.value / ftd.orders) : 0;
    const delPct = ftd.orders > 0 ? `${((ftd.deliveredOrders / ftd.orders) * 100).toFixed(0)}%` : '0%';
    const mealAsp = ftd.meals > 0 ? Math.round(ftd.value / ftd.meals) : 0;
    const mpo = ftd.orders > 0 ? (ftd.meals / ftd.orders).toFixed(2) : '0.00';
    const prepaidPct = ftd.value > 0 ? `${((ftd.prepaidValue / ftd.value) * 100).toFixed(2)}%` : '0.00%';
    const discountPct = ftd.value > 0 ? `${((ftd.discount / ftd.value) * 100).toFixed(2)}%` : '0.00%';
    const revenuePct = ftd.value > 0 ? `${((ftd.revenue / ftd.value) * 100).toFixed(1)}%` : '0.0%';
    const complaintPct = ftd.deliveredOrders > 0 ? `${((ftd.complaints / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
    const feedbackPct = ftd.deliveredOrders > 0 ? `${((ftd.feedback / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
    const undeliveredPct = ftd.orders > 0 ? `${((ftd.undelivered / ftd.orders) * 100).toFixed(2)}%` : '0.00%';

    return [
      label,
      ftd.orders, mtd.orders, 0, orderAsp, delPct,
      ftd.meals, mtd.meals, 0, mealAsp, mpo,
      Math.round(ftd.value), Math.round(mtd.value), 0,
      Math.round(ftd.prepaidValue), Math.round(mtd.prepaidValue), 0, prepaidPct,
      Math.round(ftd.discount), Math.round(mtd.discount), 0, discountPct,
      Math.round(ftd.revenue), Math.round(mtd.revenue), 0, revenuePct,
      ftd.complaints, mtd.complaints, 0, complaintPct,
      ftd.feedback, mtd.feedback, 0, feedbackPct,
      ftd.undelivered, mtd.undelivered, 0, undeliveredPct,
      outlets,
    ];
  };

  blocks.forEach((blk) => {
    // 1. Red Date Banner Row
    const dateRow = new Array(39).fill('');
    dateRow[0] = blk.dateLabel;
    excelRows.push(dateRow);
    merges.push({ s: { r: rIdx, c: 0 }, e: { r: rIdx, c: 38 } });
    rIdx++;

    // 2. Group Headers
    excelRows.push([
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
    ]);
    merges.push({ s: { r: rIdx, c: 1 }, e: { r: rIdx, c: 5 } });
    merges.push({ s: { r: rIdx, c: 6 }, e: { r: rIdx, c: 10 } });
    merges.push({ s: { r: rIdx, c: 11 }, e: { r: rIdx, c: 13 } });
    merges.push({ s: { r: rIdx, c: 14 }, e: { r: rIdx, c: 17 } });
    merges.push({ s: { r: rIdx, c: 18 }, e: { r: rIdx, c: 21 } });
    merges.push({ s: { r: rIdx, c: 22 }, e: { r: rIdx, c: 25 } });
    merges.push({ s: { r: rIdx, c: 26 }, e: { r: rIdx, c: 29 } });
    merges.push({ s: { r: rIdx, c: 30 }, e: { r: rIdx, c: 33 } });
    merges.push({ s: { r: rIdx, c: 34 }, e: { r: rIdx, c: 37 } });
    rIdx++;

    // 3. Sub-headers
    excelRows.push([
      '',
      'FTD', 'MTD', 'LMTD', 'ASP', 'Del%',
      'FTD', 'MTD', 'LMTD', 'ASP', 'MPO',
      'FTD', 'MTD', 'LMTD',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      '',
    ]);
    rIdx++;

    // 4. Total Row
    excelRows.push(buildRow('Total', blk.dayTotal, blk.mtdTotal, blk.outletsCount));
    rIdx++;

    // 5. Source Rows
    SOURCES.forEach((s) => {
      const srcData = blk.bySource[s];
      excelRows.push(buildRow(s, srcData.ftd, srcData.mtd, ''));
      rIdx++;
    });

    // Merge Outlets cell vertically across all rows of this date block
    merges.push({ s: { r: rIdx - 5, c: 38 }, e: { r: rIdx - 1, c: 38 } });

    // Spacer
    excelRows.push(new Array(39).fill(''));
    rIdx++;
  });

  const ws = XLSX.utils.aoa_to_sheet(excelRows);
  ws['!merges'] = merges;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Main Report');
  XLSX.writeFile(wb, fileName || `Main_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
};
