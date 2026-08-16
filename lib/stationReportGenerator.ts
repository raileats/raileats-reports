import * as XLSX from 'xlsx-js-style';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

export interface StationReportRow {
  'Station Code': string;
  'Rank': number;
  'Station Name': string;
  'Station Rank': number | string;
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
// 1. Universal Value Getter (Key name case/spaces ignore karega)
// ---------------------------------------------------------------------------
const getVal = (row: any, searchKeys: string[]): any => {
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
// 2. Station Matcher & Normalizer (Handles "NDLS", "NDLS - New Delhi", etc.)
// ---------------------------------------------------------------------------
const getCleanStationCode = (val: any): string => {
  if (!val) return '';
  let str = String(val).trim().toUpperCase();
  if (str.includes('-')) str = str.split('-')[0].trim();
  if (str.includes('(')) str = str.split('(')[0].trim();
  if (str.includes('/')) str = str.split('/')[0].trim();
  return str.replace(/[^A-Z0-9]/g, '');
};

// ---------------------------------------------------------------------------
// 3. Station Wise Data Generator (Accepts masterOrders + raw irctcOrders)
// ---------------------------------------------------------------------------
export const generateStationWiseData = (
  masterOrders: any[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = [] // <--- IRCTC RAW FILE ARRAY YAHAN PASS KAREIN
): StationReportRow[] => {

  // A. Total Vendor mapping
  const totalStationVendorsMap: Record<string, Set<string>> = {};
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const rawSt = out?.station || out?.stationCode || out?.stn_code || out?.deliveryStation || out?.stationName || '';
    const cleanSt = getCleanStationCode(rawSt);
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (cleanSt && outId) {
      if (!totalStationVendorsMap[cleanSt]) totalStationVendorsMap[cleanSt] = new Set();
      totalStationVendorsMap[cleanSt].add(outId);
    }
  });

  // Station Rank is maintained in Outlet Master at Outlet-ID level.
  // Convert it to a Station Code -> Rank map so station reports can use one
  // stable station-level rank even when multiple outlets belong to a station.
  const stationRankMap: Record<string, number> = {};
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const cleanSt = getCleanStationCode(out?.station || out?.stationCode || out?.stn_code || out?.deliveryStation || out?.stationName || '');
    const rank = Number(out?.stationRank ?? out?.['Station Rank']);
    if (cleanSt && Number.isFinite(rank)) {
      stationRankMap[cleanSt] = Number.isFinite(stationRankMap[cleanSt])
        ? Math.min(stationRankMap[cleanSt], rank)
        : rank;
    }
  });

  // B. STEP 1: Feedback Good/Bad map.
  // Primary source = raw IRCTC file.
  // Fallback = master rows, which also carry Delivery Station + Feedback Type.
  // This prevents zero counts when the raw IRCTC array is not persisted.
  const irctcFeedbackMap: Record<string, { good: number; bad: number }> = {};

  const feedbackSource =
    Array.isArray(irctcOrders) && irctcOrders.length > 0
      ? irctcOrders
      : (masterOrders || []);

  feedbackSource.forEach((row: any) => {
    const rawStation = getVal(row, [
      'Delivery Station',
      'DeliveryStation',
      'Delivery Station Name',
      'Station Code',
      'Station'
    ]);
    const stationCode = getCleanStationCode(rawStation);
    if (!stationCode) return;

    if (!irctcFeedbackMap[stationCode]) {
      irctcFeedbackMap[stationCode] = { good: 0, bad: 0 };
    }

    // EXACT requirement: only AD / Feedback Type is counted.
    // FEEDBACK -> Feedback Good
    // COMPLAIN -> Feedback Bad
    // No fallback to Comments / Remarks / other columns.
    const typeVal = String(
      getVal(row, ['Feedback Type', 'FeedbackType', 'FEEDBACK TYPE']) || ''
    ).trim().toUpperCase();

    if (typeVal === 'FEEDBACK') {
      irctcFeedbackMap[stationCode].good += 1;
    } else if (typeVal === 'COMPLAIN') {
      irctcFeedbackMap[stationCode].bad += 1;
    }
  });

  // C. STEP 2: Build the COMPLETE station list from Outlet Master first.
  // Business/financial numbers are intentionally still calculated ONLY from
  // delivered Master Data rows below. This keeps zero-business stations visible.
  const stationMap: Record<string, any> = {};

  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const rawStationCode = out?.stationCode || out?.station || out?.stn_code || out?.station_code || out?.deliveryStation || '';
    const stationCode = getCleanStationCode(rawStationCode);
    if (!stationCode) return;

    const masterStationName = String(out?.stationName || '').trim();
    if (!stationMap[stationCode]) {
      stationMap[stationCode] = {
        stationCode,
        stationName: masterStationName || stationCode,
        stationRank: stationRankMap[stationCode] ?? '',
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
        deliveredOutlets: new Set<string>(),
      };
    } else if (masterStationName && stationMap[stationCode].stationName === stationCode) {
      stationMap[stationCode].stationName = masterStationName;
    }
  });

  // Delivered Master Data is the ONLY source for station business.
  (masterOrders || []).forEach((row: any) => {
    const rawStation = getVal(row, [
      'Station Code',
      'StationCode',
      'Delivery Station',
      'DeliveryStation',
      'Station Name',
      'Station'
    ]);

    const stationCode = getCleanStationCode(rawStation);
    if (!stationCode) return;

    const masterRowStationName = String(
      getVal(row, ['Station Name', 'Delivery Station Name']) || ''
    ).trim();

    // Outlet Master remains authoritative for Station Name. Master Data is
    // only used for the station-code-wise delivered business calculation.
    if (!stationMap[stationCode]) {
      stationMap[stationCode] = {
        stationCode,
        stationName: masterRowStationName || stationCode,
        stationRank: stationRankMap[stationCode] ?? '',
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
        deliveredOutlets: new Set<string>(),
      };
    }

    const st = stationMap[stationCode];
    if ((!st.stationName || st.stationName === stationCode) && masterRowStationName) {
      st.stationName = masterRowStationName;
    }

    const finalStatus = String(getVal(row, ['Final Status', 'Delivery Status', 'Status']) || '').trim().toLowerCase();
    const irctcStatus = String(getVal(row, ['IRCTC Status', 'Order Status']) || '').trim().toLowerCase();
    const rfStatus = String(getVal(row, ['RF Status']) || '').trim().toLowerCase();
    const outletId = String(getVal(row, ['Outlet ID', 'OutletId', 'outlet_id']) || '').trim();

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
      st.vendorPrice += Number(getVal(row, ['Final Vendor Price', 'Vendor Price']) || 0);
      st.finalBasePrice += Number(getVal(row, ['Final Base Price', 'Base Price']) || 0);
      st.totalComm += Number(getVal(row, ['Final Total Commission', 'Total Commission']) || 0);
      st.irctcComm += Number(getVal(row, ['Final IRCTC Commission', 'IRCTC Comm']) || 0);
      st.rfComm += Number(getVal(row, ['Final RF Commission', 'RF Commission']) || 0);
      st.gst += Number(getVal(row, ['Final GST', 'GST']) || 0);
      st.totalDiscount += Number(getVal(row, ['Final Total Discount', 'Discount']) || 0);
      st.vendorDiscount += Number(getVal(row, ['Final Vendor Discount', 'Vendor Discount']) || 0);
      st.rfDiscount += Number(getVal(row, ['Final RF Discount', 'RF Discount']) || 0);
      st.deliveryCharges += Number(getVal(row, ['Delivery Charges', 'Delivery Charge']) || 0);
      st.sellingPrice += Number(getVal(row, ['Final Selling Price', 'Selling Price']) || 0);
      st.orderTotal += Number(getVal(row, ['Final Order Total', 'Order Total']) || 0);
      st.discountedBase += Number(getVal(row, ['Discounted Base Price']) || 0);
      st.ppd += Number(getVal(row, ['PPD', 'ppd']) || 0);
      st.cod += Number(getVal(row, ['COD', 'cod']) || 0);
      st.meals += Number(getVal(row, ['Meals', 'Meal Count']) || 1);
      st.deliveredOrdersCount += Number(getVal(row, ['Orders Count']) || 1);

      if (outletId) {
        st.deliveredOutlets.add(outletId);
      }
    }
  });

  // D. Business Rank = rank by delivered Final Base Price (highest first).
  // The visible list itself remains ordered by Station Rank from Outlet Master.
  const businessRankMap: Record<string, number> = {};
  const businessRankedStations = Object.values(stationMap).sort((a: any, b: any) => {
    const bp = Number(b.finalBasePrice || 0) - Number(a.finalBasePrice || 0);
    if (bp !== 0) return bp;
    const ar = Number.isFinite(Number(a.stationRank)) ? Number(a.stationRank) : Number.MAX_SAFE_INTEGER;
    const br = Number.isFinite(Number(b.stationRank)) ? Number(b.stationRank) : Number.MAX_SAFE_INTEGER;
    return ar - br;
  });
  businessRankedStations.forEach((st: any, index: number) => {
    businessRankMap[st.stationCode] = index + 1;
  });

  // E. Station list order comes from Outlet Master Station Rank.
  const sortedStations = Object.values(stationMap).sort((a: any, b: any) => {
    const ar = Number.isFinite(Number(a.stationRank)) ? Number(a.stationRank) : Number.MAX_SAFE_INTEGER;
    const br = Number.isFinite(Number(b.stationRank)) ? Number(b.stationRank) : Number.MAX_SAFE_INTEGER;
    if (ar !== br) return ar - br;
    return String(a.stationCode).localeCompare(String(b.stationCode));
  });

  // F. Final Mapping with Feedback Data.
  return sortedStations.map((st: any): StationReportRow => {
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
      : 0;

    const irctcFeedback = irctcFeedbackMap[st.stationCode] || { good: 0, bad: 0 };

    return {
      'Station Code': st.stationCode,
      'Rank': businessRankMap[st.stationCode] ?? sortedStations.length,
      'Station Name': st.stationName || st.stationCode,
      'Station Rank': st.stationRank ?? '',
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
      'Feedback Good': irctcFeedback.good,
      'Feedback Bad': irctcFeedback.bad,
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
  irctcOrders: any[] = [],
  fileNamePrefix: string = 'STATION_WISE_REPORT'
) => {
  const stationData = generateStationWiseData(masterOrders, outletsMasterInfo, irctcOrders);
  if (stationData.length === 0) {
    alert('Station report export karne ke liye koi data available nahi hai.');
    return;
  }

  const headers = Object.keys(stationData[0] || {});
  const worksheet = XLSX.utils.json_to_sheet(stationData);

  // Excel layout: Row 1 = Total, Row 2 = Header, Row 3+ = Station List.
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  for (let r = range.e.r; r >= 1; r--) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r, c });
      const to = XLSX.utils.encode_cell({ r: r + 1, c });
      if (worksheet[from]) worksheet[to] = worksheet[from];
      else delete worksheet[to];
    }
  }

  headers.forEach((header, c) => {
    worksheet[XLSX.utils.encode_cell({ r: 1, c })] = { t: 's', v: header };
  });

  const numericColumns = new Set([
    'Vendor Price','Final Base Price','Final Total Commission','Final IRCTC Comm',
    'Final RF Commission','Final GST','Final Discount','Final Vendor Discount',
    'Final RF Discount','Delivery Charges','Final Selling Price','Final Order Total',
    'Discounted Base Price','PPD','COD','Meals','Count of Delivered Orders',
    'Not Delivered Order','Feedback Good','Feedback Bad','Count of Delivered Outlets',
    'Total Station Vendors'
  ]);

  const totalRow: Record<string, any> = {};
  headers.forEach((header, index) => {
    if (index === 0) totalRow[header] = 'TOTAL';
    else if (numericColumns.has(header)) {
      totalRow[header] = Number(stationData.reduce((sum, row) => sum + (Number(row[header]) || 0), 0).toFixed(2));
    } else totalRow[header] = '';
  });
  headers.forEach((header, c) => {
    const cell = XLSX.utils.encode_cell({ r: 0, c });
    worksheet[cell] = { t: typeof totalRow[header] === 'number' ? 'n' : 's', v: totalRow[header] };
  });

  worksheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: range.e.r + 1, c: range.e.c },
  });
  worksheet['!autofilter'] = { ref: XLSX.utils.encode_range({
    s: { r: 1, c: 0 }, e: { r: range.e.r + 1, c: range.e.c }
  }) };
  worksheet['!freeze'] = { xSplit: 4, ySplit: 2 };

  const darkGreen = 'FF166534';
  const darkBrown = 'FF78350F';
  const red = 'FFDC2626';
  const headerStyle = { font: { bold: true, color: { rgb: 'FFFFFFFF' } }, fill: { fgColor: { rgb: 'FF0F172A' } }, alignment: { horizontal: 'center' } };
  const totalStyle = { font: { bold: true, color: { rgb: 'FF0F172A' } }, fill: { fgColor: { rgb: 'FFE2E8F0' } }, alignment: { horizontal: 'center' } };

  headers.forEach((_, c) => {
    const totalCell = worksheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (totalCell) totalCell.s = totalStyle;
    const headerCell = worksheet[XLSX.utils.encode_cell({ r: 1, c })];
    if (headerCell) headerCell.s = headerStyle;
  });

  stationData.forEach((row, rowIndex) => {
    const excelRow = rowIndex + 2;
    const baseBlank = Number(row['Final Base Price'] || 0) <= 0;
    const stationRank = Number(row['Station Rank']);
    const businessRank = Number(row['Rank']);
    const rowColor = baseBlank ? red : (Number.isFinite(stationRank) && Number.isFinite(businessRank) && businessRank <= stationRank ? darkGreen : darkBrown);
    headers.forEach((_, c) => {
      const cell = worksheet[XLSX.utils.encode_cell({ r: excelRow, c })];
      if (cell) cell.s = { font: { color: { rgb: rowColor } } };
    });
  });

  worksheet['!cols'] = headers.map((header) => ({
    wch: header === 'Station Name' ? 30 : header === 'Station Code' ? 14 : header === 'Station Rank' || header === 'Rank' ? 12 : 18
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Station Wise Summary');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`, { cellStyles: true } as any);
};
