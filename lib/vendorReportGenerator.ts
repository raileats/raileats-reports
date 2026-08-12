import * as XLSX from 'xlsx';

export const generateVendorReportWorkbook = (
  masterData: any[],
  outletsMasterInfo: Record<string, any> = {},
  penaltySummary: Record<string, number> = {}
) => {
  if (!masterData || masterData.length === 0) {
    alert('No data available to generate Vendor Report.');
    return;
  }

  // 1. Grouping and aggregation logic
  const outletMap: Record<string, any> = {};

  masterData.forEach((row) => {
    const outletId = String(row['Outlet ID'] || row['OutletId'] || '').trim();
    if (!outletId) return;

    const vendorName = String(row['Vendor Name'] || '').trim();
    const stationCode = String(row['Station Code'] || '').trim();
    const outletInfo = outletsMasterInfo[outletId] || {};

    const state = outletInfo.state || row['State'] || '';
    const gstNo = outletInfo.gst || row['GST No'] || '';
    const irctcStatus = outletInfo.irctcStatus || row['Outlet IRCTC Status'] || 'Active';

    if (!outletMap[outletId]) {
      outletMap[outletId] = {
        outletId,
        vendorName,
        stationCode,
        state,
        gstNo,
        irctcStatus,
        totalOrders: 0,
        deliveredOrders: 0,
        cancelledOrders: 0,
        undeliveredOrders: 0,
        totalMeals: 0,
        totalSellingPrice: 0,
        totalBasePrice: 0,
        totalVendorPrice: 0,
        totalRFCommission: 0,
        totalIRCTCCommission: 0,
        totalVendorDiscount: 0,
        totalRFDiscount: 0,
        totalDiscount: 0,
        totalGST: 0,
        totalDeliveryCharge: 0,
        penaltyAmount: penaltySummary[outletId] || 0,
      };
    }

    const isDelivered = row['Final Status'] === 'Delivered';
    const isCancelled = row['Final Status'] === 'Cancelled';
    const isUndelivered =
      row['Final Status'] === 'Not Delivered' ||
      String(row['IRCTC Status'] || '').toUpperCase().includes('UNDELIVERED');

    const sellingPrice = parseFloat(row['Final Selling Price'] || row['Selling Price'] || 0) || 0;
    const basePrice = parseFloat(row['Final Base Price'] || row['Base Price'] || 0) || 0;
    const vendorPrice = parseFloat(row['Final Vendor Price'] || row['Vendor Price'] || 0) || 0;
    const rfComm = parseFloat(row['Final RF Commission'] || row['RF Comm'] || 0) || 0;
    const irctcComm = parseFloat(row['Final IRCTC Commission'] || row['IRCTC Comm'] || 0) || 0;
    const vendorDisc = parseFloat(row['Final Vendor Discount'] || 0) || 0;
    const rfDisc = parseFloat(row['Final RF Discount'] || 0) || 0;
    const totalDisc = parseFloat(row['Final Total Discount'] || row['Total Discount'] || 0) || 0;
    const gst = parseFloat(row['Final GST'] || row['GST'] || 0) || 0;
    const deliveryCharge = parseFloat(row['Delivery Charges'] || row['Delivery Charge'] || 0) || 0;
    const meals = parseInt(row['Meals'] || '1', 10) || 1;

    outletMap[outletId].totalOrders += 1;
    if (isDelivered) outletMap[outletId].deliveredOrders += 1;
    if (isCancelled) outletMap[outletId].cancelledOrders += 1;
    if (isUndelivered) outletMap[outletId].undeliveredOrders += 1;

    outletMap[outletId].totalMeals += meals;
    outletMap[outletId].totalSellingPrice += sellingPrice;
    outletMap[outletId].totalBasePrice += basePrice;
    outletMap[outletId].totalVendorPrice += vendorPrice;
    outletMap[outletId].totalRFCommission += rfComm;
    outletMap[outletId].totalIRCTCCommission += irctcComm;
    outletMap[outletId].totalVendorDiscount += vendorDisc;
    outletMap[outletId].totalRFDiscount += rfDisc;
    outletMap[outletId].totalDiscount += totalDisc;
    outletMap[outletId].totalGST += gst;
    outletMap[outletId].totalDeliveryCharge += deliveryCharge;
  });

  // 2. Map Row Data with Net Payment Calculation
  const reportRows = Object.values(outletMap).map((o: any) => {
    // Formula: Vendor Price - Vendor Discount - Penalty Deductions
    const netPayment = Number(
      (o.totalVendorPrice - o.totalVendorDiscount - o.penaltyAmount).toFixed(2)
    );

    return {
      'Outlet ID': o.outletId,
      'Vendor Name': o.vendorName,
      'Station Code': o.stationCode,
      'State': o.state,
      'GST Number': o.gstNo,
      'IRCTC Status': o.irctcStatus,
      'Total Orders': o.totalOrders,
      'Delivered Orders': o.deliveredOrders,
      'Cancelled Orders': o.cancelledOrders,
      'Undelivered Orders': o.undeliveredOrders,
      'Meal Count': o.totalMeals,
      'Total Selling Price (₹)': Number(o.totalSellingPrice.toFixed(2)),
      'Total Base Price (₹)': Number(o.totalBasePrice.toFixed(2)),
      'Final Vendor Payout (₹)': Number(o.totalVendorPrice.toFixed(2)),
      'Vendor Discount (₹)': Number(o.totalVendorDiscount.toFixed(2)),
      'RF Discount (₹)': Number(o.totalRFDiscount.toFixed(2)),
      'Total Discount (₹)': Number(o.totalDiscount.toFixed(2)),
      'RF Commission (₹)': Number(o.totalRFCommission.toFixed(2)),
      'IRCTC Commission (₹)': Number(o.totalIRCTCCommission.toFixed(2)),
      'GST 5% (₹)': Number(o.totalGST.toFixed(2)),
      'Delivery Charges (₹)': Number(o.totalDeliveryCharge.toFixed(2)),
      'Penalty / Deductions (₹)': Number(o.penaltyAmount.toFixed(2)),
      'Net Payment (₹)': netPayment,
    };
  });

  // 3. Totals Aggregation
  const totals = reportRows.reduce(
    (acc, row) => {
      acc.totalOrders += row['Total Orders'];
      acc.deliveredOrders += row['Delivered Orders'];
      acc.cancelledOrders += row['Cancelled Orders'];
      acc.undeliveredOrders += row['Undelivered Orders'];
      acc.totalMeals += row['Meal Count'];
      acc.totalSellingPrice += row['Total Selling Price (₹)'];
      acc.totalBasePrice += row['Total Base Price (₹)'];
      acc.totalVendorPrice += row['Final Vendor Payout (₹)'];
      acc.totalVendorDiscount += row['Vendor Discount (₹)'];
      acc.totalRFDiscount += row['RF Discount (₹)'];
      acc.totalDiscount += row['Total Discount (₹)'];
      acc.totalRFCommission += row['RF Commission (₹)'];
      acc.totalIRCTCCommission += row['IRCTC Commission (₹)'];
      acc.totalGST += row['GST 5% (₹)'];
      acc.totalDeliveryCharge += row['Delivery Charges (₹)'];
      acc.totalPenalty += row['Penalty / Deductions (₹)'];
      acc.totalNetPayment += row['Net Payment (₹)'];
      return acc;
    },
    {
      totalOrders: 0,
      deliveredOrders: 0,
      cancelledOrders: 0,
      undeliveredOrders: 0,
      totalMeals: 0,
      totalSellingPrice: 0,
      totalBasePrice: 0,
      totalVendorPrice: 0,
      totalVendorDiscount: 0,
      totalRFDiscount: 0,
      totalDiscount: 0,
      totalRFCommission: 0,
      totalIRCTCCommission: 0,
      totalGST: 0,
      totalDeliveryCharge: 0,
      totalPenalty: 0,
      totalNetPayment: 0,
    }
  );

  const grandTotalRow = {
    'Outlet ID': 'GRAND TOTAL',
    'Vendor Name': '',
    'Station Code': '',
    'State': '',
    'GST Number': '',
    'IRCTC Status': '',
    'Total Orders': totals.totalOrders,
    'Delivered Orders': totals.deliveredOrders,
    'Cancelled Orders': totals.cancelledOrders,
    'Undelivered Orders': totals.undeliveredOrders,
    'Meal Count': totals.totalMeals,
    'Total Selling Price (₹)': Number(totals.totalSellingPrice.toFixed(2)),
    'Total Base Price (₹)': Number(totals.totalBasePrice.toFixed(2)),
    'Final Vendor Payout (₹)': Number(totals.totalVendorPrice.toFixed(2)),
    'Vendor Discount (₹)': Number(totals.totalVendorDiscount.toFixed(2)),
    'RF Discount (₹)': Number(totals.totalRFDiscount.toFixed(2)),
    'Total Discount (₹)': Number(totals.totalDiscount.toFixed(2)),
    'RF Commission (₹)': Number(totals.totalRFCommission.toFixed(2)),
    'IRCTC Commission (₹)': Number(totals.totalIRCTCCommission.toFixed(2)),
    'GST 5% (₹)': Number(totals.totalGST.toFixed(2)),
    'Delivery Charges (₹)': Number(totals.totalDeliveryCharge.toFixed(2)),
    'Penalty / Deductions (₹)': Number(totals.totalPenalty.toFixed(2)),
    'Net Payment (₹)': Number(totals.totalNetPayment.toFixed(2)),
  };

  reportRows.push(grandTotalRow);

  // 4. Create Sheet
  const worksheet = XLSX.utils.json_to_sheet(reportRows);

  // Auto-fit Column Widths
  const colWidths = [
    { wch: 14 }, // Outlet ID
    { wch: 28 }, // Vendor Name
    { wch: 14 }, // Station Code
    { wch: 16 }, // State
    { wch: 18 }, // GST Number
    { wch: 14 }, // IRCTC Status
    { wch: 14 }, // Total Orders
    { wch: 16 }, // Delivered Orders
    { wch: 16 }, // Cancelled Orders
    { wch: 18 }, // Undelivered Orders
    { wch: 12 }, // Meal Count
    { wch: 22 }, // Total Selling Price
    { wch: 20 }, // Total Base Price
    { wch: 22 }, // Final Vendor Payout
    { wch: 20 }, // Vendor Discount
    { wch: 18 }, // RF Discount
    { wch: 18 }, // Total Discount
    { wch: 20 }, // RF Commission
    { wch: 22 }, // IRCTC Commission
    { wch: 14 }, // GST 5%
    { wch: 20 }, // Delivery Charges
    { wch: 24 }, // Penalty / Deductions
    { wch: 20 }, // Net Payment
  ];
  worksheet['!cols'] = colWidths;

  // 5. Build Workbook & Trigger Export
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor Summary Report');

  XLSX.writeFile(
    workbook,
    `RELFOOD_VENDOR_REPORT_${new Date().toISOString().slice(0, 10)}.xlsx`
  );
};
