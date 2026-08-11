import * as XLSX from 'xlsx';

// Helper to sanitize float values
const toNum = (val: any): number => {
  const num = parseFloat(val);
  return isNaN(num) ? 0 : num;
};

export const generateVendorRDSWorkbook = (
  masterData: any[],
  penaltySummary: Record<string, number> = {},
  currentMonthRecords: any[] = [],
  outletsMasterInfo: Record<string, any> = {}
) => {
  if (!masterData || masterData.length === 0) {
    if (typeof window !== 'undefined') {
      alert('Kripya pehle master data process karein!');
    }
    return;
  }

  // 1. Group master rows by cleaned Outlet ID
  const outletGroups: Record<string, any[]> = {};
  masterData.forEach((row) => {
    const rawOutlet = row['Outlet ID'] || row['OutletId'] || row['Outlet Id'] || '';
    const outletId = String(rawOutlet).trim().replace(/\.0$/, '');
    if (!outletId) return;

    if (!outletGroups[outletId]) {
      outletGroups[outletId] = [];
    }
    outletGroups[outletId].push(row);
  });

  // Current month map for fast lookup
  const currentMonthMap: Record<string, any> = {};
  currentMonthRecords.forEach((cm) => {
    const cId = String(cm.outletId || '').trim().replace(/\.0$/, '');
    if (cId) currentMonthMap[cId] = cm;
  });

  // 2. Compute RDS Summary per Outlet
  const summaryRows = Object.keys(outletGroups).map((outletId) => {
    const rows = outletGroups[outletId];
    const firstRow = rows[0] || {};
    const outletInfo = outletsMasterInfo[outletId] || {};
    const curMonthInfo = currentMonthMap[outletId] || {};

    const deliveredOrders = rows.filter((r) => r['Final Status'] === 'Delivered');
    const totalDelivered = deliveredOrders.length;
    const totalCancelled = rows.filter((r) => r['Final Status'] === 'Cancelled').length;
    const totalNotDelivered = rows.filter((r) => r['Final Status'] === 'Not Delivered').length;

    // Financial calculations on delivered orders
    const totalVendorPrice = deliveredOrders.reduce((sum, r) => sum + toNum(r['Final Vendor Price']), 0);
    const totalBasePrice = deliveredOrders.reduce((sum, r) => sum + toNum(r['Final Base Price']), 0);
    const totalSellingPrice = deliveredOrders.reduce((sum, r) => sum + toNum(r['Final Selling Price']), 0);
    const totalVendorDiscount = deliveredOrders.reduce((sum, r) => sum + toNum(r['Final Vendor Discount']), 0);
    const totalPpd = deliveredOrders.reduce((sum, r) => sum + toNum(r['PPD']), 0);
    const totalCod = deliveredOrders.reduce((sum, r) => sum + toNum(r['COD']), 0);
    const totalMeals = deliveredOrders.reduce((sum, r) => sum + (parseInt(r['Meals'] || '1', 10) || 1), 0);

    // Vendor RDS Settlement Rules:
    // Net Vendor Base Share = Vendor Price - 50% Vendor Discount Share
    const netVendorShare = Number((totalVendorPrice - totalVendorDiscount).toFixed(2));
    
    // Penalties / Deductions
    const totalPenalty = penaltySummary[outletId] || 0;

    // Current Month / Previous Balance adjustments
    const previousBalance = toNum(curMonthInfo.previousBalance);
    const paidToVendor = toNum(curMonthInfo.paidToVendors);
    const receivedFromVendor = toNum(curMonthInfo.receivedFromVendor);
    const creditNote = toNum(curMonthInfo.creditNoteToVendor);

    // Net Payable = (Net Vendor Share - COD collected by vendor) - Penalties + Previous Balance - Paid + Received + CreditNote
    const netPayable = Number(
      (netVendorShare - totalCod - totalPenalty + previousBalance - paidToVendor + receivedFromVendor + creditNote).toFixed(2)
    );

    return {
      'Outlet ID': outletId,
      'Vendor Name': firstRow['Vendor Name'] || curMonthInfo.vendorName || outletInfo.outletName || '',
      'Station Code': firstRow['Station Code'] || curMonthInfo.stationCode || outletInfo.station || '',
      'State': outletInfo.state || firstRow['State'] || '',
      'GST No': outletInfo.gst || firstRow['GST No'] || '',
      'IRCTC Status': outletInfo.irctcStatus || firstRow['Outlet IRCTC Status'] || 'ACTIVE',
      'Total Orders': rows.length,
      'Delivered Orders': totalDelivered,
      'Cancelled Orders': totalCancelled,
      'Not Delivered Orders': totalNotDelivered,
      'Total Meals': totalMeals,
      'Total Selling Price (₹)': Number(totalSellingPrice.toFixed(2)),
      'Gross Vendor Price (₹)': Number(totalVendorPrice.toFixed(2)),
      'Vendor Discount (50%) (₹)': Number(totalVendorDiscount.toFixed(2)),
      'Net Vendor Share (₹)': netVendorShare,
      'PPD Orders Total (₹)': Number(totalPpd.toFixed(2)),
      'COD Collected by Vendor (₹)': Number(totalCod.toFixed(2)),
      'Total Penalty & Deductions (₹)': Number(totalPenalty.toFixed(2)),
      'Previous Balance (₹)': previousBalance,
      'Paid to Vendor (₹)': paidToVendor,
      'Payment Received (₹)': receivedFromVendor,
      'Credit Note (₹)': creditNote,
      'Net Balance Payable (₹)': netPayable,
    };
  });

  // 3. Create Workbook
  const workbook = XLSX.utils.book_new();

  // Tab 1: Vendor RDS Summary
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Vendor RDS Summary');

  // Tab 2: All Delivered Orders Detail
  const deliveredOrdersData = masterData.filter((r) => r['Final Status'] === 'Delivered');
  const deliveredSheet = XLSX.utils.json_to_sheet(deliveredOrdersData);
  XLSX.utils.book_append_sheet(workbook, deliveredSheet, 'Delivered Orders');

  // Tab 3: Complete Master Records
  const allMasterSheet = XLSX.utils.json_to_sheet(masterData);
  XLSX.utils.book_append_sheet(workbook, allMasterSheet, 'All Master Data');

  // Trigger Download
  const dateStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `RELFOOD_VENDOR_RDS_SETTLEMENT_${dateStr}.xlsx`);
};

export default generateVendorRDSWorkbook;
