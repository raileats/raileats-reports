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
}

/**
 * Generates Outlet-level Vendor Summary Report matching exact screenshot format
 */
export const generateVendorWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  penaltySummary: Record<string, number> = {}
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

    // Metadata from Outlet Master if available
    const outMaster = outletsMasterInfo[outletId];
    if (outMaster) {
      vendorName = outMaster.outletName || '';
      stationCode = outMaster.station || '';
      stationName = outMaster.station || '';
    }

    orders.forEach((ord) => {
      if (!vendorName && ord['Vendor Name']) vendorName = ord['Vendor Name'];
      if (!stationCode && ord['Station Code']) stationCode = ord['Station Code'];
      if (!stationName) stationName = stationCode;

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
    });
  });

  // Sort by Final Base Price (Highest to Lowest)
  rawOutletList.sort((a, b) => b.finalBasePrice - a.finalBasePrice);

  // Map into exact columns matching image structure
  return rawOutletList.map((row, idx) => ({
    'Aggregator Outlet ID': row.outletId,
    'Station Code': row.stationCode,
    'Rank': idx + 1,
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
  }));
};

/**
 * Direct Vendor Report Excel Workbook Generator
 */
export const generateVendorReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  penaltySummary: Record<string, number> | string = {},
  fileNamePrefix: string = 'VENDOR_OUTLET_WISE_REPORT'
) => {
  const penaltyMap: Record<string, number> =
    typeof penaltySummary === 'object' && penaltySummary !== null ? penaltySummary : {};
  const prefix: string =
    typeof penaltySummary === 'string' ? penaltySummary : fileNamePrefix || 'VENDOR_OUTLET_WISE_REPORT';

  const vendorData = generateVendorWiseData(masterOrders, outletsMasterInfo, penaltyMap);
  if (vendorData.length === 0) {
    alert('Vendor report export karne ke liye koi data available nahi hai.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(vendorData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor Report');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${prefix}_${todayStr}.xlsx`);
};
