import * as XLSX from 'xlsx';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

export interface StationReportRow {
  'Station Code': string;
  'Rank': number;
  'Station Name': string;
  'Vendor Price': number;
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
  'Not Delivered Order': number;
  'Not Delivered %': string;
  'PPD % of Final Selling Price': string;
  'Feedback Good': number;
  'Feedback Bad': number;
  'Count of Delivered Outlets': number;
  'Total Station Vendors': number;
}

// ---------------------------------------------------------------------------
// 1. Flexible Key Finder (Case-insensitive, spaces/underscores ignore karta hai)
// ---------------------------------------------------------------------------
const getAnyValue = (row: any, searchKeys: string[]): any => {
  if (!row || typeof row !== 'object') return null;
  const rowKeys = Object.keys(row);
  for (const sk of searchKeys) {
    const cleanSk = sk.toLowerCase().replace(/[^a-z0-9]/g, '');
    const foundKey = rowKeys.find(
      (rk) => rk.toLowerCase().replace(/[^a-z0-9]/g, '') === cleanSk
    );
    if (
      foundKey &&
      row[foundKey] !== undefined &&
      row[foundKey] !== null &&
      String(row[foundKey]).trim() !== ''
    ) {
      return row[foundKey];
    }
  }
  return null;
};

// ---------------------------------------------------------------------------
// 2. Station Key Cleaner & Normalizer (E.g. "NDLS - New Delhi" -> "NDLS")
// ---------------------------------------------------------------------------
const normalizeStationKey = (val: any): string => {
  if (!val) return '';
  let str = String(val).trim().toUpperCase();
  
  // Agar "NDLS - NEW DELHI" ya "NDLS (NEW DELHI)" format ho toh code nikaal lo
  if (str.includes('-')) {
    str = str.split('-')[0].trim();
  } else if (str.includes('(')) {
    str = str.split('(')[0].trim();
  }
  return str.replace(/[^A-Z0-9]/g, '');
};

// ---------------------------------------------------------------------------
// 3. Feedback / Complaint Valid Entry Checker
// ---------------------------------------------------------------------------
const hasValidEntry = (val: any): boolean => {
  if (val === null || val === undefined) return false;
  const str = String(val).trim().toLowerCase();
  // Khali, '0', 'nan', 'null', 'nil', '-' ko ignore karega
  if (str === '' || str === '0' || str === 'nan' || str === 'null' || str === 'nil' || str === '-' || str === 'none') {
    return false;
  }
  return true;
};

// ---------------------------------------------------------------------------
// Main Station Wise Generator Function
// ---------------------------------------------------------------------------
export const generateStationWiseData = (
  masterOrders: any[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {}
): StationReportRow[] => {
  // Master Outlets Vendor Mapping
  const totalStationVendorsMap: Record<string, Set<string>> = {};
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const rawSt = out?.station || out?.stationCode || out?.stn_code || out?.deliveryStation || out?.stationName || '';
    const cleanSt = normalizeStationKey(rawSt);
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (cleanSt && outId) {
      if (!totalStationVendorsMap[cleanSt]) totalStationVendorsMap[cleanSt] = new Set();
      totalStationVendorsMap[cleanSt].add(outId);
    }
  });

  // Aggregation Map
  const stationMap: Record<string, any> = {};

  (masterOrders || []).forEach((row: any) => {
    // Match Delivery Station (From IRCTC Raw) OR Station Code (From Master Sheet)
    const rawStation = getAnyValue(row, [
      'Delivery Station',
      'DeliveryStation',
      'Station Code',
      'StationCode',
      'Delivery Station Code',
      'Delivery Station Name',
      'Station Name',
      'Station',
      'Stn Code',
      'stn_code'
    ]);

    const stationKey = normalizeStationKey(rawStation) || 'UNKNOWN';
    const displayStationName = String(
      getAnyValue(row, ['Station Name', 'Delivery Station Name', 'Delivery Station', 'Station']) || rawStation || stationKey
    ).trim();

    if (!stationMap[stationKey]) {
      stationMap[stationKey] = {
        stationCode: stationKey,
        stationName: displayStationName,
        vendorPrice: 0,
        finalBasePrice: 0,
        totalComm: 0,
        irctcComm: 0,
        rfComm: 0,
        gst: 0,
        totalDiscount: 0,
        vendorDiscount: 0,
        rfDiscount: 0,
        deliveryCharges: 0,
        sellingPrice: 0,
        orderTotal: 0,
        discountedBase: 0,
        ppd: 0,
        cod: 0,
        meals: 0,
        deliveredOrdersCount: 0,
        notDeliveredOrdersCount: 0,
        feedbackGoodCount: 0,
        feedbackBadCount: 0,
        deliveredOutlets: new Set<string>(),
      };
    }

    const st = stationMap[stationKey];

    // Status Values Check
    const finalStatus = String(getAnyValue(row, ['Final Status', 'Delivery Status', 'Status', 'Order Status']) || '').trim().toLowerCase();
    const irctcStatus = String(getAnyValue(row, ['IRCTC Status', 'IRCTC Delivery Status']) || '').trim().toLowerCase();
    const rfStatus = String(getAnyValue(row, ['RF Status']) || '').trim().toLowerCase();
    const outletId = String(getAnyValue(row, ['Outlet ID', 'OutletId', 'outlet_id']) || '').trim();

    // -------------------------------------------------------------
    // FEEDBACK & COMPLAINT COUNT LOGIC (Delivery Station Wise Sum)
    // -------------------------------------------------------------
    const feedbackVal = getAnyValue(row, [
      'Feedback',
      'Customer Feedback',
      'Feedback Good',
      'Feedback Comment',
      'Rating',
      'Review'
    ]);

    const complaintVal = getAnyValue(row, [
      'Complaint',
      'Feedback Bad',
      'Complaint Comment',
      'Comment',
      'Issue',
      'Customer Complaint'
    ]);

    if (hasValidEntry(feedbackVal)) {
      // Agar direct count number diya ho (jaise 1, 2) ya text comment ho
      const num = Number(feedbackVal);
      st.feedbackGoodCount += !isNaN(num) && num > 0 ? num : 1;
    }

    if (hasValidEntry(complaintVal)) {
      const num = Number(complaintVal);
      st.feedbackBadCount += !isNaN(num) && num > 0 ? num : 1;
    }

    // -------------------------------------------------------------
    // DELIVERY & NOT-DELIVERY LOGIC
    // -------------------------------------------------------------
    const isDelivered = finalStatus === 'delivered' || finalStatus === 'success';
    const isNotDelivered =
      finalStatus === 'not delivered' ||
      finalStatus === 'cancelled' ||
      finalStatus === 'undelivered' ||
      irctcStatus.includes('undelivered') ||
      irctcStatus.includes('cancel') ||
      rfStatus.includes('undelivered') ||
      rfStatus.includes('cancel');

    if (isNotDelivered) {
      st.notDeliveredOrdersCount += 1;
    }

    if (isDelivered) {
      st.vendorPrice += Number(getAnyValue(row, ['Final Vendor Price', 'Vendor Price', 'vendor_price']) || 0);
      st.finalBasePrice += Number(getAnyValue(row, ['Final Base Price', 'Total Base Price', 'Base Price', 'base_price']) || 0);
      st.totalComm += Number(getAnyValue(row, ['Final Total Commission', 'Total Commission', 'commission']) || 0);
      st.irctcComm += Number(getAnyValue(row, ['Final IRCTC Commission', 'IRCTC Comm', 'Vendor Comm']) || 0);
      st.rfComm += Number(getAnyValue(row, ['Final RF Commission', 'RF Commission', 'rf_comm']) || 0);
      st.gst += Number(getAnyValue(row, ['Final GST', 'Total Gst', 'GST', 'gst']) || 0);
      st.totalDiscount += Number(getAnyValue(row, ['Final Total Discount', 'Discount', 'discount']) || 0);
      st.vendorDiscount += Number(getAnyValue(row, ['Final Vendor Discount', 'Vendor Discount']) || 0);
      st.rfDiscount += Number(getAnyValue(row, ['Final RF Discount', 'RF Discount']) || 0);
      st.deliveryCharges += Number(getAnyValue(row, ['Delivery Charges', 'Delivery Charge']) || 0);
      st.sellingPrice += Number(getAnyValue(row, ['Final Selling Price', 'Selling Price']) || 0);
      st.orderTotal += Number(getAnyValue(row, ['Final Order Total', 'Order Total']) || 0);
      st.discountedBase += Number(getAnyValue(row, ['Discounted Base Price']) || 0);
      st.ppd += Number(getAnyValue(row, ['PPD', 'ppd']) || 0);
      st.cod += Number(getAnyValue(row, ['COD', 'cod']) || 0);
      st.meals += Number(getAnyValue(row, ['Meals', 'Meal Count']) || 1);
      st.deliveredOrdersCount += Number(getAnyValue(row, ['Orders Count']) || 1);

      if (outletId) {
        st.deliveredOutlets.add(outletId);
      }
    }
  });

  // Sort descending by Final Base Price
  const sortedStations = Object.values(stationMap).sort(
    (a: any, b: any) => b.finalBasePrice - a.finalBasePrice
  );

  // Return Formatted Rows
  return sortedStations.map((st: any, idx: number): StationReportRow => {
    const vPrice = Number(st.vendorPrice.toFixed(2));
    const bPrice = Number(st.finalBasePrice.toFixed(2));
    const tComm = Number(st.totalComm.toFixed(2));
    const sPrice = Number(st.sellingPrice.toFixed(2));
    const ppdVal = Number(st.ppd.toFixed(2));

    const checkPct = bPrice > 0 ? `${((tComm / bPrice) * 100).toFixed(2)}%` : '0.00%';
    const notDeliveredPct =
      st.deliveredOrdersCount > 0
        ? `${((st.notDeliveredOrdersCount / st.deliveredOrdersCount) * 100).toFixed(2)}%`
        : st.notDeliveredOrdersCount > 0
        ? '100.00%'
        : '0.00%';
    const ppdPct = sPrice > 0 ? `${((ppdVal / sPrice) * 100).toFixed(2)}%` : '0.00%';

    const totalVendors = totalStationVendorsMap[st.stationCode]
      ? totalStationVendorsMap[st.stationCode].size
      : Math.max(st.deliveredOutlets.size, 1);

    return {
      'Station Code': st.stationCode,
      'Rank': idx + 1,
      'Station Name': st.stationName,
      'Vendor Price': vPrice,
      'Final Base Price': bPrice,
      'Final Total Commission': tComm,
      'Final IRCTC Comm': Number(st.irctcComm.toFixed(2)),
      'Final RF Commission': Number(st.rfComm.toFixed(2)),
      'Final GST': Number(st.gst.toFixed(2)),
      'Final Discount': Number(st.totalDiscount.toFixed(2)),
      'Final Vendor Discount': Number(st.vendorDiscount.toFixed(2)),
      'Final RF Discount': Number(st.rfDiscount.toFixed(2)),
      'Delivery Charges': Number(st.deliveryCharges.toFixed(2)),
      'Final Selling Price': sPrice,
      'Final Order Total': Number(st.orderTotal.toFixed(2)),
      'Discounted Base Price': Number(st.discountedBase.toFixed(2)),
      'PPD': ppdVal,
      'COD': Number(st.cod.toFixed(2)),
      'Meals': st.meals,
      'Check': checkPct,
      'Count of Delivered Orders': st.deliveredOrdersCount,
      'Not Delivered Order': st.notDeliveredOrdersCount,
      'Not Delivered %': notDeliveredPct,
      'PPD % of Final Selling Price': ppdPct,
      'Feedback Good': st.feedbackGoodCount,
      'Feedback Bad': st.feedbackBadCount,
      'Count of Delivered Outlets': st.deliveredOutlets.size,
      'Total Station Vendors': totalVendors,
    };
  });
};

/**
 * Direct Excel Export Workbook Generator
 */
export const generateStationReportWorkbook = (
  masterOrders: any[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  fileNamePrefix: string = 'STATION_WISE_REPORT'
) => {
  const stationData = generateStationWiseData(masterOrders, outletsMasterInfo);
  if (stationData.length === 0) {
    alert('Station report export karne ke liye koi data available nahi hai.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(stationData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Station Wise Summary');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`);
};
