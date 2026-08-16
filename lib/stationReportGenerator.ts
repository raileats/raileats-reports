import * as XLSX from 'xlsx';
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
  irctcOrders: any[] = []
): StationReportRow[] => {
  // IMPORTANT SOURCE RULES:
  // 1) Station list / Station Name / Station Rank come from Outlet Master.
  // 2) Business figures come ONLY from delivered Master Data rows, grouped by Station Code.
  // 3) Stations with no delivered business must still appear with zero business.
  const stationMap: Record<string, any> = {};
  const totalStationVendorsMap: Record<string, Set<string>> = {};

  const ensureStation = (stationCode: string, fallbackName = '') => {
    if (!stationCode) return null;
    if (!stationMap[stationCode]) {
      stationMap[stationCode] = {
        stationCode,
        stationName: fallbackName || stationCode,
        stationRank: '',
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
    return stationMap[stationCode];
  };

  // Build the COMPLETE station list from Outlet Master first.
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const rawCode = out?.stationCode || out?.station || out?.stn_code || out?.deliveryStation || '';
    const stationCode = getCleanStationCode(rawCode);
    if (!stationCode) return;

    const stationName = String(
      out?.stationName || out?.['Station Name'] || out?.station_name || ''
    ).trim() || stationCode;
    const stationRankRaw = out?.stationRank ?? out?.['Station Rank'] ?? out?.rank ?? '';
    const stationRank = Number(stationRankRaw);
    const st = ensureStation(stationCode, stationName);
    if (!st) return;

    // Prefer the actual Outlet Master station name, never the short station code.
    if (stationName && stationName !== stationCode) st.stationName = stationName;

    if (Number.isFinite(stationRank)) {
      st.stationRank = Number.isFinite(Number(st.stationRank))
        ? Math.min(Number(st.stationRank), stationRank)
        : stationRank;
    }

    const outletId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (outletId) {
      if (!totalStationVendorsMap[stationCode]) totalStationVendorsMap[stationCode] = new Set();
      totalStationVendorsMap[stationCode].add(outletId);
    }
  });

  // Feedback map: retain the existing IRCTC/master behaviour.
  const irctcFeedbackMap: Record<string, { good: number; bad: number }> = {};
  const feedbackSource = Array.isArray(irctcOrders) && irctcOrders.length > 0 ? irctcOrders : (masterOrders || []);
  feedbackSource.forEach((row: any) => {
    const rawStation = getVal(row, ['Delivery Station', 'DeliveryStation', 'Delivery Station Name', 'Station Code', 'Station']);
    const stationCode = getCleanStationCode(rawStation);
    if (!stationCode) return;
    if (!irctcFeedbackMap[stationCode]) irctcFeedbackMap[stationCode] = { good: 0, bad: 0 };
    const typeVal = String(getVal(row, ['Feedback Type', 'FeedbackType', 'FEEDBACK TYPE']) || '').trim().toUpperCase();
    if (typeVal === 'FEEDBACK') irctcFeedbackMap[stationCode].good += 1;
    if (typeVal === 'COMPLAIN') irctcFeedbackMap[stationCode].bad += 1;
  });

  // Master Data is the ONLY business source. Every delivered row is grouped by Station Code.
  (masterOrders || []).forEach((row: any) => {
    const rawStation = getVal(row, ['Station Code', 'StationCode', 'Delivery Station', 'DeliveryStation', 'Station']);
    const stationCode = getCleanStationCode(rawStation);
    if (!stationCode) return;

    // If a code exists in Master Data but not Outlet Master, keep it as a fallback station.
    const rowStationName = String(getVal(row, ['Station Name', 'Delivery Station Name']) || '').trim();
    const st = ensureStation(stationCode, rowStationName || stationCode)!;

    const finalStatus = String(getVal(row, ['Final Status', 'Delivery Status', 'Status']) || '').trim().toLowerCase();
    const irctcStatus = String(getVal(row, ['IRCTC Status', 'Order Status']) || '').trim().toLowerCase();
    const rfStatus = String(getVal(row, ['RF Status']) || '').trim().toLowerCase();
    const outletId = String(getVal(row, ['Outlet ID', 'OutletId', 'outlet_id', 'Aggregator Outlet ID']) || '').trim();

    const isDelivered = finalStatus === 'delivered' || finalStatus === 'success';
    const isNotDelivered =
      finalStatus === 'not delivered' || finalStatus === 'cancelled' || finalStatus === 'undelivered' ||
      irctcStatus.includes('undelivered') || irctcStatus.includes('cancel') ||
      rfStatus.includes('undelivered') || rfStatus.includes('cancel');

    if (isNotDelivered) st.notDeliveredOrdersCount += 1;
    if (!isDelivered) return;

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
    if (outletId) st.deliveredOutlets.add(outletId);
  });

  // Business Rank = delivered Final Base Price descending. Zero-business stations stay at the bottom.
  const sortedStations = Object.values(stationMap).sort((a: any, b: any) => {
    const bp = Number(b.finalBasePrice || 0) - Number(a.finalBasePrice || 0);
    if (bp !== 0) return bp;
    const ar = Number.isFinite(Number(a.stationRank)) ? Number(a.stationRank) : Number.MAX_SAFE_INTEGER;
    const br = Number.isFinite(Number(b.stationRank)) ? Number(b.stationRank) : Number.MAX_SAFE_INTEGER;
    if (ar !== br) return ar - br;
    return String(a.stationCode).localeCompare(String(b.stationCode), undefined, { numeric: true });
  });

  return sortedStations.map((st: any, idx: number): StationReportRow => {
    const vPrice = Number(st.vendorPrice.toFixed(2));
    const bPrice = Number(st.finalBasePrice.toFixed(2));
    const tComm = Number(st.totalComm.toFixed(2));
    const sPrice = Number(st.sellingPrice.toFixed(2));
    const ppdVal = Number(st.ppd.toFixed(2));
    const checkPct = bPrice > 0 ? `${((tComm / bPrice) * 100).toFixed(2)}%` : '0.00%';
    const notDeliveredPct = st.deliveredOrdersCount > 0
      ? `${((st.notDeliveredOrdersCount / st.deliveredOrdersCount) * 100).toFixed(2)}%`
      : st.notDeliveredOrdersCount > 0 ? '100.00%' : '0.00%';
    const ppdPct = sPrice > 0 ? `${((ppdVal / sPrice) * 100).toFixed(2)}%` : '0.00%';
    const totalVendors = totalStationVendorsMap[st.stationCode]?.size || st.deliveredOutlets.size || 0;
    const irctcFeedback = irctcFeedbackMap[st.stationCode] || { good: 0, bad: 0 };

    return {
      'Station Code': st.stationCode,
      'Rank': idx + 1,
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
 * Station Report Excel: TOTAL row on top, filterable header below it,
 * complete Outlet-Master station list, and delivered Master-Data business.
 */
export const generateStationReportWorkbook = (
  masterOrders: any[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = [],
  fileNamePrefix: string = 'STATION_WISE_REPORT'
) => {
  const stationData = generateStationWiseData(masterOrders, outletsMasterInfo, irctcOrders);
  if (!stationData.length) {
    alert('Station report export karne ke liye Outlet Master mein koi Station Code available nahi hai.');
    return;
  }

  const columns = Object.keys(stationData[0]);
  const numericColumns = new Set(columns.filter((c) => !['Station Code','Station Name','Check','Not Delivered %','PPD % of Final Selling Price'].includes(c)));
  const total: Record<string, any> = {};
  columns.forEach((c) => { total[c] = ''; });
  total['Station Code'] = 'TOTAL';
  numericColumns.forEach((c) => {
    total[c] = stationData.reduce((sum, r: any) => sum + (Number(r[c]) || 0), 0);
  });
  total['Station Name'] = `${stationData.length} Stations`;
  total['Rank'] = '';
  total['Station Rank'] = '';
  total['Check'] = '';
  total['Not Delivered %'] = '';
  total['PPD % of Final Selling Price'] = '';

  const ws = XLSX.utils.aoa_to_sheet([
    columns.map((c) => total[c] ?? ''),
    columns,
    ...stationData.map((r: any) => columns.map((c) => r[c] ?? '')),
  ]);

  const lastRow = stationData.length + 2;
  const lastCol = XLSX.utils.encode_col(columns.length - 1);
  ws['!autofilter'] = { ref: `A2:${lastCol}${lastRow}` };
  ws['!freeze'] = { xSplit: 0, ySplit: 2 };
  ws['!cols'] = columns.map((c) => ({
    wch: c === 'Station Code' ? 12 : c === 'Station Name' ? 34 : c === 'Rank' || c === 'Station Rank' ? 12 : 16
  }));

  const styleCell = (addr: string, style: any) => {
    if (ws[addr]) ws[addr].s = style;
  };
  const headerStyle = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '1E293B' } }, alignment: { horizontal: 'center', vertical: 'center' } };
  const totalStyle = { font: { bold: true, color: { rgb: '000000' } }, fill: { fgColor: { rgb: 'CBD5E1' } }, alignment: { horizontal: 'center', vertical: 'center' } };
  for (let c = 0; c < columns.length; c++) {
    styleCell(`${XLSX.utils.encode_col(c)}1`, totalStyle);
    styleCell(`${XLSX.utils.encode_col(c)}2`, headerStyle);
  }

  // Row text colour: zero business -> red; otherwise business rank vs station rank.
  stationData.forEach((r: any, i: number) => {
    const excelRow = i + 3;
    const base = Number(r['Final Base Price'] || 0);
    const businessRank = Number(r['Rank']);
    const stationRank = Number(r['Station Rank']);
    const rgb = base === 0 ? 'DC2626' : (Number.isFinite(stationRank) && businessRank <= stationRank ? '166534' : '78350F');
    for (let c = 0; c < columns.length; c++) {
      styleCell(`${XLSX.utils.encode_col(c)}${excelRow}`, { font: { color: { rgb } } });
    }
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Station Report');
  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `${fileNamePrefix}_${todayStr}.xlsx`);
};
