import * as XLSX from 'xlsx';

export interface MasterOrderRow {
  'IRCTC Order ID': string;
  'RF Order ID'?: string;
  'Outlet ID': string;
  'Vendor Name': string;
  'Station Code': string;
  'State'?: string;
  'GST No'?: string;
  'Outlet IRCTC Status'?: string;
  'Train No'?: string;
  'Booking Date'?: string;
  'Delivery Date'?: string;
  'Payment Type'?: string;
  'RF Status'?: string;
  'IRCTC Status'?: string;
  'Final Status'?: string;
  'Final Vendor Price': number;
  'Final Base Price': number;
  'Final Total Commission': number;
  'Final IRCTC Commission': number;
  'Final RF Commission': number;
  'Final Total Discount': number;
  'Final Vendor Discount': number;
  'Final RF Discount': number;
  'Discounted Base Price': number;
  'Final GST': number;
  'Delivery Charges': number;
  'Final Selling Price': number;
  'Final Order Total': number;
  'PPD': number;
  'COD': number;
  'Meals': number;
  'Margin %': number;
  'Orders Count': number;
  [key: string]: any;
}

export interface CurrentMonthRow {
  outletId: string;
  vendorName?: string;
  stationCode?: string;
  previousBalance?: number;
  paidToVendors?: number;
  receivedFromVendor?: number;
  creditNoteToVendor?: number;
}

export interface OutletMasterInfo {
  outletId?: string;
  state?: string;
  gst?: string;
  irctcStatus?: string;
  outletName?: string;
  station?: string;
}

// Dynamic Invoice Generator helper (e.g. RF26-27-AUG0001)
export const getInvoicePrefix = (sampleDateStr?: string): string => {
  let dateObj = new Date();
  if (sampleDateStr) {
    const parsed = new Date(sampleDateStr);
    if (!isNaN(parsed.getTime())) {
      dateObj = parsed;
    }
  }

  const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const monthStr = monthNames[dateObj.getMonth()];
  const fullYear = dateObj.getFullYear();
  const month = dateObj.getMonth() + 1;

  let startYearShort = fullYear % 100;
  let endYearShort = (fullYear + 1) % 100;

  if (month < 4) {
    startYearShort = (fullYear - 1) % 100;
    endYearShort = fullYear % 100;
  }

  const fyStr = `${String(startYearShort).padStart(2, '0')}-${String(endYearShort).padStart(2, '0')}`;
  return `RF${fyStr}-${monthStr}`;
};

/**
 * 41 Columns Vendor RDS Aggregation Engine
 */
export const generateVendorRdsData = (
  masterOrders: MasterOrderRow[],
  penaltySummary: Record<string, number> = {},
  currentMonthList: CurrentMonthRow[] = [],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {}
): any[] => {
  // 1. Group Master Orders Outlet ID wise
  const outletOrderMap: Record<string, MasterOrderRow[]> = {};
  let sampleDate = '';

  masterOrders.forEach((row) => {
    const outletId = String(row['Outlet ID'] || '').trim().replace(/\.0$/, '');
    if (!outletId) return;

    if (!sampleDate && (row['Delivery Date'] || row['Booking Date'])) {
      sampleDate = row['Delivery Date'] || row['Booking Date'] || '';
    }

    if (!outletOrderMap[outletId]) {
      outletOrderMap[outletId] = [];
    }
    outletOrderMap[outletId].push(row);
  });

  const curMonthMap: Record<string, CurrentMonthRow> = {};
  currentMonthList.forEach((c) => {
    const oid = String(c.outletId || '').trim().replace(/\.0$/, '');
    if (oid) curMonthMap[oid] = c;
  });

  // 2. Union of All Outlets (Taaki 0-order outlets with previous balance miss na ho)
  const allOutletIds = new Set<string>([
    ...Object.keys(outletOrderMap),
    ...Object.keys(curMonthMap),
  ]);

  if (allOutletIds.size === 0) return [];

  const invoicePrefix = getInvoicePrefix(sampleDate);
  const aggregatedList: any[] = [];

  allOutletIds.forEach((outletId) => {
    const orders = outletOrderMap[outletId] || [];
    const outletInfo = outletsMasterInfo[outletId] || {};
    const curMonthInfo = curMonthMap[outletId] || {};
    const firstRow = orders[0] || {};

    const vendorName = String(firstRow['Vendor Name'] || curMonthInfo.vendorName || outletInfo.outletName || '').trim();
    const stationCode = String(firstRow['Station Code'] || curMonthInfo.stationCode || outletInfo.station || '').trim();
    const vendorWithStation = `${stationCode} ${vendorName}`.trim();
    const state = String(outletInfo.state || firstRow['State'] || '').trim();
    const gstNo = String(outletInfo.gst || firstRow['GST No'] || '').trim();

    let finalVendorPriceSum = 0;
    let finalBasePriceSum = 0;
    let finalRFCommissionSum = 0;
    let finalIRCTCCommissionSum = 0;
    let finalGSTSum = 0;
    let finalOrderTotalSum = 0;
    let finalTotalCommissionSum = 0;
    let deliveryChargesSum = 0;
    let finalSellingPriceSum = 0;
    let finalTotalDiscountSum = 0;
    let finalVendorDiscountSum = 0;
    let finalRFDiscountSum = 0;
    let ppdSum = 0;
    let codSum = 0;
    let ordersCountSum = 0;
    let mealsSum = 0;
    let discountedBasePriceSum = 0;
    let marginSumDelivered = 0;
    let deliveredCount = 0;

    // Financial calculations Delivered orders par sum honge
    orders.forEach((ord) => {
      const status = String(ord['Final Status'] || '').trim().toLowerCase();
      if (status === 'delivered') {
        finalVendorPriceSum += Number(ord['Final Vendor Price'] || 0);
        finalBasePriceSum += Number(ord['Final Base Price'] || 0);
        finalRFCommissionSum += Number(ord['Final RF Commission'] || 0);
        finalIRCTCCommissionSum += Number(ord['Final IRCTC Commission'] || 0);
        finalGSTSum += Number(ord['Final GST'] || 0);
        finalOrderTotalSum += Number(ord['Final Order Total'] || 0);
        finalTotalCommissionSum += Number(ord['Final Total Commission'] || 0);
        deliveryChargesSum += Number(ord['Delivery Charges'] || 0);
        finalSellingPriceSum += Number(ord['Final Selling Price'] || 0);
        finalTotalDiscountSum += Number(ord['Final Total Discount'] || 0);
        finalVendorDiscountSum += Number(ord['Final Vendor Discount'] || 0);
        finalRFDiscountSum += Number(ord['Final RF Discount'] || 0);
        ppdSum += Number(ord['PPD'] || 0);
        codSum += Number(ord['COD'] || 0);
        ordersCountSum += Number(ord['Orders Count'] || 1);
        mealsSum += Number(ord['Meals'] || 0);
        discountedBasePriceSum += Number(ord['Discounted Base Price'] || 0);
        marginSumDelivered += Number(ord['Margin %'] || 0);
        deliveredCount += 1;
      }
    });

    const finalVendorPrice = Number(finalVendorPriceSum.toFixed(2));
    const finalBasePrice = Number(finalBasePriceSum.toFixed(2));
    const finalRFCommission = Number(finalRFCommissionSum.toFixed(2));
    const finalIRCTCCommission = Number(finalIRCTCCommissionSum.toFixed(2));
    const finalGST = Number(finalGSTSum.toFixed(2));
    const finalOrderTotal = Number(finalOrderTotalSum.toFixed(2));
    const finalTotalCommission = Number(finalTotalCommissionSum.toFixed(2));
    const deliveryCharges = Number(deliveryChargesSum.toFixed(2));
    const finalSellingPrice = Number(finalSellingPriceSum.toFixed(2));
    const finalTotalDiscount = Number(finalTotalDiscountSum.toFixed(2));
    const finalVendorDiscount = Number(finalVendorDiscountSum.toFixed(2));
    const finalRFDiscount = Number(finalRFDiscountSum.toFixed(2));
    const ppd = Number(ppdSum.toFixed(2));
    const cod = Number(codSum.toFixed(2));
    const discountedBasePrice = Number(discountedBasePriceSum.toFixed(2));

    const avgMarginPct = deliveredCount > 0 ? Number((marginSumDelivered / deliveredCount).toFixed(2)) : 0;
    const penalty = Number((penaltySummary[outletId] || 0).toFixed(2));
    const paidToVendors = Number((curMonthInfo.paidToVendors || 0).toFixed(2));
    const previousBalance = Number((curMonthInfo.previousBalance || 0).toFixed(2));
    const paymentReceivedFromVendor = Number((curMonthInfo.receivedFromVendor || 0).toFixed(2));
    const creditNoteToVendor = Number((curMonthInfo.creditNoteToVendor || 0).toFixed(2));

    const grossCommission = Number((finalTotalCommission + penalty).toFixed(2));
    const isAndhra = state.toUpperCase().includes('ANDHRA');
    const igst = isAndhra ? 0 : Number((grossCommission * 0.18).toFixed(2));
    const cgst = isAndhra ? Number((grossCommission * 0.09).toFixed(2)) : 0;
    const sgst = isAndhra ? Number((grossCommission * 0.09).toFixed(2)) : 0;
    const totalTaxes = Number((igst + cgst + sgst).toFixed(2));

    const totalThisMonth = Number((grossCommission + totalTaxes).toFixed(2));
    const add = Number((finalGST + totalThisMonth + paidToVendors + deliveryCharges + previousBalance).toFixed(2));
    const less = Number((paymentReceivedFromVendor + creditNoteToVendor + finalRFDiscount + ppd).toFixed(2));
    const netPayment = Number((add - less).toFixed(2));

    const asPerReverse = Number(
      (
        finalVendorPrice +
        paymentReceivedFromVendor +
        creditNoteToVendor -
        penalty -
        totalTaxes -
        paidToVendors -
        previousBalance -
        finalVendorDiscount -
        cod
      ).toFixed(2)
    );

    const diff = Number((netPayment + asPerReverse).toFixed(2));

    aggregatedList.push({
      'Outlet ID': outletId,
      'Vendor Name': vendorName,
      'Station Code': stationCode,
      'GST Number': gstNo,
      'Vendor with Station Code': vendorWithStation,
      _sortKey: vendorWithStation.toLowerCase(),

      'Final Vendor Price': finalVendorPrice,
      'Final Base Price': finalBasePrice,
      'Final RF Commission': finalRFCommission,
      'Final IRCTC Commission': finalIRCTCCommission,
      'Final GST': finalGST,
      'Final Order Total': finalOrderTotal,
      'Final Total Commission': finalTotalCommission,
      'Penalty': penalty,
      'Gross Commission': grossCommission,
      'IGST': igst,
      'CGST': cgst,
      'SGST': sgst,
      'IGST+CGST+SGST': totalTaxes,
      'Total This Month': totalThisMonth,
      'Paid to Vendors By Relfood': paidToVendors,
      'Delivery Charges': deliveryCharges,
      'Previouse Balance': previousBalance,
      'Final Selling Price': finalSellingPrice,
      'Final Total Discount': finalTotalDiscount,
      'Final Vendor Discount': finalVendorDiscount,
      'Payment Received from Vendor to Relfood': paymentReceivedFromVendor,
      'Credit Note to Vendor by Relfood': creditNoteToVendor,
      'Final RF Discount': finalRFDiscount,
      'PPD': ppd,
      'COD': cod,
      'ADD': add,
      'Less': less,
      'Net Payment': netPayment,
      'As per Reverse': asPerReverse,
      'Diff': diff,
      'Orders Count': ordersCountSum,
      'Meals': mealsSum,
      'State': state,
      'Discounted Base Price': discountedBasePrice,
      'Margin %': avgMarginPct,
    });
  });

  // 3. Alphabetical Sort by Station + Vendor
  aggregatedList.sort((a, b) => a._sortKey.localeCompare(b._sortKey));

  // 4. Exact 41 Columns Ordered Output
  return aggregatedList.map((row, index) => {
    const invoiceSeq = String(index + 1).padStart(4, '0');
    return {
      'Outlet ID': row['Outlet ID'],
      'Vendor Name': row['Vendor Name'],
      'Station Code': row['Station Code'],
      'GST Number': row['GST Number'],
      'Vendor with Station Code': row['Vendor with Station Code'],
      'Invoice Number': `${invoicePrefix}${invoiceSeq}`,
      'Final Vendor Price': row['Final Vendor Price'],
      'Final Base Price': row['Final Base Price'],
      'Final RF Commission': row['Final RF Commission'],
      'Final IRCTC Commission': row['Final IRCTC Commission'],
      'Final GST': row['Final GST'],
      'Final Order Total': row['Final Order Total'],
      'Final Total Commission': row['Final Total Commission'],
      'Penalty': row['Penalty'],
      'Gross Commission': row['Gross Commission'],
      'IGST': row['IGST'],
      'CGST': row['CGST'],
      'SGST': row['SGST'],
      'IGST+CGST+SGST': row['IGST+CGST+SGST'],
      'Total This Month': row['Total This Month'],
      'Paid to Vendors By Relfood': row['Paid to Vendors By Relfood'],
      'Delivery Charges': row['Delivery Charges'],
      'Previouse Balance': row['Previouse Balance'],
      'Final Selling Price': row['Final Selling Price'],
      'Final Total Discount': row['Final Total Discount'],
      'Final Vendor Discount': row['Final Vendor Discount'],
      'Payment Received from Vendor to Relfood': row['Payment Received from Vendor to Relfood'],
      'Credit Note to Vendor by Relfood': row['Credit Note to Vendor by Relfood'],
      'Final RF Discount': row['Final RF Discount'],
      'PPD': row['PPD'],
      'COD': row['COD'],
      'ADD': row['ADD'],
      'Less': row['Less'],
      'Net Payment': row['Net Payment'],
      'As per Reverse': row['As per Reverse'],
      'Diff': row['Diff'],
      'Orders Count': row['Orders Count'],
      'Meals': row['Meals'],
      'State': row['State'],
      'Discounted Base Price': row['Discounted Base Price'],
      'Margin %': row['Margin %'],
    };
  });
};

/**
 * Direct Export matching app/page.tsx signature
 */
export const generateVendorRDSWorkbook = (
  masterOrders: MasterOrderRow[],
  penaltySummary: Record<string, number> = {},
  currentMonthRecords: CurrentMonthRow[] = [],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  fileNamePrefix: string = 'VENDOR_RDS_REPORT'
) => {
  const rdsRows = generateVendorRdsData(masterOrders, penaltySummary, currentMonthRecords, outletsMasterInfo);
  if (rdsRows.length === 0) {
    alert('Koi data available nahi hai export karne ke liye.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rdsRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor RDS');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`);
};
