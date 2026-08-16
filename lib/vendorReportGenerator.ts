import * as XLSX from 'xlsx';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

export interface VendorReportRow {
  'Aggregator Outlet ID': string;
  'Station Code': string;
  'Rank': number;
  'Station Name': string;
  'Vendor Name': string;
  'Vendor Price': number;
  'Net Payment': number;
  'Final Base Price': number;
  'Final Total Commission': number;
  'Final IRCTC Comm': number;
  'Final RF Commission': number;
  'Final GST': number;
  'Final Discount': number;
  'Final Vendor Discount': number;
  'Final RF Discount': number;
  'Delivery Charges': number;
  'Final Selling Price': number;
  'Final Order Total': number;
  'Discounted Base Price': number;
  'PPD': number;
  'COD': number;
  'Meals': number;
  'Check': string;
  'Count of Delivered Orders': number;
  'Count of Not_Delivered As per IRCTC Status': number;
  'Not_Delivered %': string;
  'Prepaid %': string;
  'Vendor Payment Type': string;
  'Discount Applied': string;
}

/**
 * Generates Outlet-level Vendor Summary Report matching exact screenshot format
 */
export const generateVendorWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  penaltySummary: Record<string, number> = {},
  currentMonthRecords: any[] = []
): VendorReportRow[] => {
  const outletOrdersMap: Record<string, MasterOrderRow[]> = {};

  // Group all master orders by Outlet ID
  masterOrders.forEach((row) => {
    const outletId = String(row['Outlet ID'] || '').trim().replace(/\.0$/, '');
    if (!outletId) return;

    if (!outletOrdersMap[outletId]) {
      outletOrdersMap[outletId] = [];
    }
    outletOrdersMap[outletId].push(row);
  });

  // Also consider outlets present in Outlet Master even if zero orders
  Object.keys(outletsMasterInfo).forEach((outId) => {
    const cleanId = String(outId).trim().replace(/\.0$/, '');
    if (cleanId && !outletOrdersMap[cleanId]) {
      outletOrdersMap[cleanId] = [];
    }
  });

  const currentMonthMap: Record<string, any> = {};
  (currentMonthRecords || []).forEach((c: any) => {
    const oid = String(c?.outletId || c?.['Outlet Id'] || c?.['Outlet ID'] || '').trim().replace(/\.0$/, '');
    if (oid) currentMonthMap[oid] = c;
  });

  const rawOutletList: any[] = [];

  Object.entries(outletOrdersMap).forEach(([outletId, orders]) => {
    let vendorPriceSum = 0;
    let basePriceSum = 0;
    let totalCommSum = 0;
    let irctcCommSum = 0;
    let rfCommSum = 0;
    let gstSum = 0;
    let totalDiscountSum = 0;
    let vendorDiscountSum = 0;
    let rfDiscountSum = 0;
    let deliveryChargesSum = 0;
    let sellingPriceSum = 0;
    let orderTotalSum = 0;
    let discountedBaseSum = 0;
    let ppdSum = 0;
    let codSum = 0;
    let mealsSum = 0;
    let deliveredOrdersCount = 0;
    let notDeliveredOrdersCount = 0;

    let vendorName = '';
    let stationCode = '';
    let stationName = '';
    let stationRank: number | '' = '';

    // Metadata from Outlet Master if available
    const outMaster = outletsMasterInfo[outletId];
    if (outMaster) {
      vendorName = outMaster.outletName || '';
      // Station Code remains the short code from Outlet Master.
      stationCode = (outMaster as any).stationCode || outMaster.station || '';
      // IMPORTANT: Station Name comes from Outlet Master `Station Name`
      // (column O in the uploaded Outlet Master report), never from Station Code.
      stationName =
        (outMaster as any).stationName ||
        (outMaster as any)['Station Name'] ||
        '';
      const rawStationRank = (outMaster as any).stationRank ?? (outMaster as any)['Station Rank'];
      if (rawStationRank !== undefined && rawStationRank !== null && String(rawStationRank).trim() !== '') {
        const parsedRank = Number(rawStationRank);
        stationRank = Number.isFinite(parsedRank) ? parsedRank : String(rawStationRank).trim() as any;
      }
    }

    orders.forEach((ord) => {
      if (!vendorName && ord['Vendor Name']) vendorName = ord['Vendor Name'];
      if (!stationCode && ord['Station Code']) stationCode = ord['Station Code'];
      // Only use order-level station text as a last-resort display name.
      if (!stationName) {
        stationName =
          String((ord as any)['Station Name'] || '').trim() ||
          String((ord as any)['Delivery Station Name'] || '').trim() ||
          stationCode;
      }

      const finalStatus = String(ord['Final Status'] || '').trim().toLowerCase();
      const irctcStatus = String(ord['IRCTC Status'] || '').trim().toLowerCase();
      const rfStatus = String(ord['RF Status'] || '').trim().toLowerCase();

      // Count Not Delivered based on status
      if (
        finalStatus === 'not delivered' ||
        finalStatus === 'cancelled' ||
        irctcStatus.includes('undelivered') ||
        irctcStatus.includes('cancel') ||
        rfStatus.includes('undelivered') ||
        rfStatus.includes('cancel')
      ) {
        notDeliveredOrdersCount += 1;
      }

      if (finalStatus === 'delivered') {
        vendorPriceSum += Number(ord['Final Vendor Price'] || 0);
        basePriceSum += Number(ord['Final Base Price'] || 0);
        totalCommSum += Number(ord['Final Total Commission'] || 0);
        irctcCommSum += Number(ord['Final IRCTC Commission'] || 0);
        rfCommSum += Number(ord['Final RF Commission'] || 0);
        gstSum += Number(ord['Final GST'] || 0);
        totalDiscountSum += Number(ord['Final Total Discount'] || 0);
        vendorDiscountSum += Number(ord['Final Vendor Discount'] || 0);
        rfDiscountSum += Number(ord['Final RF Discount'] || 0);
        deliveryChargesSum += Number(ord['Delivery Charges'] || 0);
        sellingPriceSum += Number(ord['Final Selling Price'] || 0);
        orderTotalSum += Number(ord['Final Order Total'] || 0);
        discountedBaseSum += Number(ord['Discounted Base Price'] || 0);
        ppdSum += Number(ord['PPD'] || 0);
        codSum += Number(ord['COD'] || 0);
        mealsSum += Number(ord['Meals'] || 0);
        deliveredOrdersCount += Number(ord['Orders Count'] || 1);
      }
    });

    const vendorPrice = Number(vendorPriceSum.toFixed(2));
    const vendorDiscount = Number(vendorDiscountSum.toFixed(2));
    const penaltyAmount =
      typeof penaltySummary === 'object' && penaltySummary !== null
        ? Number(penaltySummary[outletId] || 0)
        : 0;
    const netPayment = Number((vendorPrice - vendorDiscount - penaltyAmount).toFixed(2));
    const finalBasePrice = Number(basePriceSum.toFixed(2));
    const totalCommission = Number(totalCommSum.toFixed(2));
    const irctcComm = Number(irctcCommSum.toFixed(2));
    const rfComm = Number(rfCommSum.toFixed(2));
    const gst = Number(gstSum.toFixed(2));
    const totalDiscount = Number(totalDiscountSum.toFixed(2));
    const rfDiscount = Number(rfDiscountSum.toFixed(2));
    const deliveryCharges = Number(deliveryChargesSum.toFixed(2));
    const sellingPrice = Number(sellingPriceSum.toFixed(2));
    const orderTotal = Number(orderTotalSum.toFixed(2));
    const discountedBase = Number(discountedBaseSum.toFixed(2));
    const ppd = Number(ppdSum.toFixed(2));
    const cod = Number(codSum.toFixed(2));

    // Check %: (Total Commission / Vendor Price) %
    const checkPct = vendorPrice > 0 ? `${((totalCommission / vendorPrice) * 100).toFixed(2)}%` : '#DIV/0!';

    // Not Delivered %: (Not_Delivered / Delivered Orders) %
    const notDeliveredPct =
      deliveredOrdersCount > 0
        ? `${((notDeliveredOrdersCount / deliveredOrdersCount) * 100).toFixed(2)}%`
        : notDeliveredOrdersCount > 0
        ? '100.00%'
        : '#DIV/0!';

    // Prepaid %: (PPD / Selling Price) %
    const prepaidPct = sellingPrice > 0 ? `${((ppd / sellingPrice) * 100).toFixed(2)}%` : '#DIV/0!';

    rawOutletList.push({
      outletId,
      stationCode: stationCode || '',
      stationRank,
      stationName: stationName || stationCode || '',
      vendorName: vendorName || `Outlet ${outletId}`,
      vendorPrice,
      netPayment,
      finalBasePrice,
      totalCommission,
      irctcComm,
      rfComm,
      gst,
      totalDiscount,
      vendorDiscount,
      rfDiscount,
      deliveryCharges,
      sellingPrice,
      orderTotal,
      discountedBase,
      ppd,
      cod,
      meals: mealsSum,
      checkPct,
      deliveredOrdersCount,
      notDeliveredOrdersCount,
      notDeliveredPct,
      prepaidPct,
      vendorPaymentType: String(currentMonthMap[outletId]?.vendorPaymentType || currentMonthMap[outletId]?.['Vendor Payment Type'] || '').trim(),
      discountApplied: String(currentMonthMap[outletId]?.discountApplied || currentMonthMap[outletId]?.['Discount Applied'] || '').trim(),
    });
  });

  // BUSINESS ORDER:
  // 1) Station Rank ascending, keeping all outlets of the same station together.
  // 2) Within the same Station Rank, Final Base Price highest -> lowest.
  rawOutletList.sort((a, b) => {
    const rankA = typeof a.stationRank === 'number' ? a.stationRank : Number.POSITIVE_INFINITY;
    const rankB = typeof b.stationRank === 'number' ? b.stationRank : Number.POSITIVE_INFINITY;
    if (rankA !== rankB) return rankA - rankB;
    const baseDiff = b.finalBasePrice - a.finalBasePrice;
    if (baseDiff !== 0) return baseDiff;
    return String(a.outletId).localeCompare(String(b.outletId), undefined, { numeric: true });
  });

  // Map into exact columns matching the report structure.
  // `Rank` is the Station Rank, not a global outlet row number.
  return rawOutletList.map((row) => ({
    'Aggregator Outlet ID': row.outletId,
    'Station Code': row.stationCode,
    'Rank': row.stationRank,
    'Station Name': row.stationName,
    'Vendor Name': row.vendorName,
    'Vendor Price': row.vendorPrice,
    'Net Payment': row.netPayment,
    'Final Base Price': row.finalBasePrice,
    'Final Total Commission': row.totalCommission,
    'Final IRCTC Comm': row.irctcComm,
    'Final RF Commission': row.rfComm,
    'Final GST': row.gst,
    'Final Discount': row.totalDiscount,
    'Final Vendor Discount': row.vendorDiscount,
    'Final RF Discount': row.rfDiscount,
    'Delivery Charges': row.deliveryCharges,
    'Final Selling Price': row.sellingPrice,
    'Final Order Total': row.orderTotal,
    'Discounted Base Price': row.discountedBase,
    'PPD': row.ppd,
    'COD': row.cod,
    'Meals': row.meals,
    'Check': row.checkPct,
    'Count of Delivered Orders': row.deliveredOrdersCount,
    'Count of Not_Delivered As per IRCTC Status': row.notDeliveredOrdersCount,
    'Not_Delivered %': row.notDeliveredPct,
    'Prepaid %': row.prepaidPct,
    'Vendor Payment Type': row.vendorPaymentType,
    'Discount Applied': row.discountApplied,
  }));
};

/**
 * Direct Vendor Report Excel Workbook Generator
 */
export const generateVendorReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  penaltySummary: Record<string, number> | string = {},
  currentMonthRecords: any[] = [],
  fileNamePrefix: string = 'VENDOR_OUTLET_WISE_REPORT'
) => {
  const penaltyMap: Record<string, number> =
    typeof penaltySummary === 'object' && penaltySummary !== null ? penaltySummary : {};
  const prefix: string =
    typeof penaltySummary === 'string' ? penaltySummary : fileNamePrefix || 'VENDOR_OUTLET_WISE_REPORT';

  const vendorData = generateVendorWiseData(masterOrders, outletsMasterInfo, penaltyMap, currentMonthRecords);
  if (vendorData.length === 0) {
    alert('Vendor report export karne ke liye koi data available nahi hai.');
    return;
  }

  const columns = Object.keys(vendorData[0]) as (keyof VendorReportRow)[];
  const moneyColumns = new Set<keyof VendorReportRow>([
    'Vendor Price', 'Net Payment', 'Final Base Price', 'Final Total Commission',
    'Final IRCTC Comm', 'Final RF Commission', 'Final GST', 'Final Discount',
    'Final Vendor Discount', 'Final RF Discount', 'Delivery Charges',
    'Final Selling Price', 'Final Order Total', 'Discounted Base Price', 'PPD', 'COD'
  ]);
  const totalColumns = new Set<keyof VendorReportRow>([
    ...moneyColumns,
    'Meals',
    'Count of Delivered Orders',
    'Count of Not_Delivered As per IRCTC Status'
  ]);

  const totals: Record<string, any> = {};
  columns.forEach((column) => {
    totals[column as string] = '';
  });
  totals['Aggregator Outlet ID'] = 'TOTAL';
  totals['Rank'] = '';
  totals['Station Code'] = '';
  totals['Station Name'] = '';
  totals['Vendor Name'] = '';

  totalColumns.forEach((column) => {
    totals[column as string] = Number(
      vendorData.reduce((sum, row) => sum + (Number((row as any)[column]) || 0), 0).toFixed(2)
    );
  });

  const totalVendorPrice = Number(totals['Vendor Price'] || 0);
  const totalBasePrice = Number(totals['Final Base Price'] || 0);
  const totalSellingPrice = Number(totals['Final Selling Price'] || 0);
  const totalDelivered = Number(totals['Count of Delivered Orders'] || 0);
  const totalNotDelivered = Number(totals['Count of Not_Delivered As per IRCTC Status'] || 0);
  totals['Check'] =
    totalVendorPrice > 0
      ? `${((Number(totals['Final Total Commission'] || 0) / totalVendorPrice) * 100).toFixed(2)}%`
      : '#DIV/0!';
  totals['Not_Delivered %'] =
    totalDelivered > 0
      ? `${((totalNotDelivered / totalDelivered) * 100).toFixed(2)}%`
      : totalNotDelivered > 0
        ? '100.00%'
        : '#DIV/0!';
  totals['Prepaid %'] =
    totalSellingPrice > 0
      ? `${((Number(totals['PPD'] || 0) / totalSellingPrice) * 100).toFixed(2)}%`
      : '#DIV/0!';
  totals['Vendor Payment Type'] = '';
  totals['Discount Applied'] = '';

  const headerRow = columns.map((column) => column === 'Rank' ? 'Station Rank' : String(column));
  const totalRow = columns.map((column) => totals[column as string] ?? '');
  const bodyRows = vendorData.map((row) => columns.map((column) => (row as any)[column]));

  // Excel layout: TOTAL row first, then header, then data.
  const worksheet = XLSX.utils.aoa_to_sheet([totalRow, headerRow, ...bodyRows]);

  worksheet['!cols'] = columns.map((column) => {
    if (column === 'Vendor Name') return { wch: 38 };
    if (column === 'Station Name') return { wch: 28 };
    if (column === 'Aggregator Outlet ID') return { wch: 18 };
    if (column === 'Vendor Payment Type') return { wch: 20 };
    if (column === 'Discount Applied') return { wch: 18 };
    return { wch: 16 };
  });

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor Report');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${prefix}_${todayStr}.xlsx`);
};
