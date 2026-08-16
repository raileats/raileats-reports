import * as XLSX from 'xlsx';
import { OutletMasterInfo } from './vendorRdsGenerator';

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

const getVal = (row: any, searchKeys: string[]): any => {
  if (!row || typeof row !== 'object') return null;
  const rowKeys = Object.keys(row);
  for (const sk of searchKeys) {
    const cleanSk = sk.toLowerCase().replace(/[^a-z0-9]/g, '');
    const foundKey = rowKeys.find((rk) => rk.toLowerCase().replace(/[^a-z0-9]/g, '') === cleanSk);
    if (foundKey && row[foundKey] !== undefined && row[foundKey] !== null && String(row[foundKey]).trim() !== '') {
      return row[foundKey];
    }
  }
  return null;
};

const getCleanStationCode = (val: any): string => {
  if (val === null || val === undefined) return '';
  let str = String(val).trim().toUpperCase();
  if (!str) return '';
  // Do not destroy normal station codes. Only remove descriptive suffixes.
  if (str.includes(' - ')) str = str.split(' - ')[0].trim();
  if (str.includes('(')) str = str.split('(')[0].trim();
  return str.replace(/[^A-Z0-9]/g, '');
};

const toNumber = (value: any): number => {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
};

const masterStationFields = (out: any) => {
  const rawCode = out?.stationCode ?? out?.['Station Code'] ?? out?.station ?? out?.Station ?? out?.stn_code ?? out?.deliveryStation ?? '';
  const stationCode = getCleanStationCode(rawCode);
  const stationName = String(out?.stationName ?? out?.['Station Name'] ?? out?.station_name ?? '').trim();
  const rankRaw = out?.stationRank ?? out?.['Station Rank'] ?? out?.rank ?? '';
  const rankNum = Number(rankRaw);
  return { stationCode, stationName, stationRank: Number.isFinite(rankNum) ? rankNum : '' };
};

export const generateStationWiseData = (
  masterOrders: any[] = [],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = []
): StationReportRow[] => {
  const stationMap: Record<string, any> = {};
  const stationVendors: Record<string, Set<string>> = {};

  const ensureStation = (code: string, name = '', rank: number | string = '') => {
    if (!code) return null;
    if (!stationMap[code]) {
      stationMap[code] = {
        stationCode: code,
        stationName: name || code,
        stationRank: rank,
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
    return stationMap[code];
  };

  // 1) COMPLETE station list comes ONLY from Outlet Master.
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const { stationCode, stationName, stationRank } = masterStationFields(out);
    if (!stationCode) return;
    const st = ensureStation(stationCode, stationName || stationCode, stationRank);
    if (!st) return;
    if (stationName) st.stationName = stationName;
    if (Number.isFinite(Number(stationRank))) {
      st.stationRank = Number.isFinite(Number(st.stationRank))
        ? Math.min(Number(st.stationRank), Number(stationRank))
        : Number(stationRank);
    }
    const outletId = String(out?.outletId ?? out?.['Outlet ID'] ?? out?.['Aggregator Outlet ID'] ?? '').trim();
    if (outletId) {
      if (!stationVendors[stationCode]) stationVendors[stationCode] = new Set<string>();
      stationVendors[stationCode].add(outletId);
    }
  });

  // 2) Feedback Good/Bad: preserve existing IRCTC behaviour.
  const feedbackMap: Record<string, { good: number; bad: number }> = {};
  const feedbackSource = irctcOrders.length ? irctcOrders : masterOrders;
  feedbackSource.forEach((row: any) => {
    const code = getCleanStationCode(getVal(row, ['Delivery Station', 'DeliveryStation', 'Station Code', 'Station']));
    if (!code || !stationMap[code]) return;
    if (!feedbackMap[code]) feedbackMap[code] = { good: 0, bad: 0 };
    const type = String(getVal(row, ['Feedback Type', 'FeedbackType', 'FEEDBACK TYPE']) || '').trim().toUpperCase();
    if (type === 'FEEDBACK') feedbackMap[code].good += 1;
    if (type === 'COMPLAIN') feedbackMap[code].bad += 1;
  });

  // 3) BUSINESS comes ONLY from delivered Master Data, grouped by Station Code.
  masterOrders.forEach((row: any) => {
    const code = getCleanStationCode(getVal(row, ['Station Code', 'StationCode', 'Delivery Station', 'DeliveryStation', 'Station']));
    if (!code || !stationMap[code]) return;
    const st = stationMap[code];

    const finalStatus = String(getVal(row, ['Final Status']) || '').trim().toLowerCase();
    const irctcStatus = String(getVal(row, ['IRCTC Status', 'Order Status']) || '').trim().toLowerCase();
    const rfStatus = String(getVal(row, ['RF Status']) || '').trim().toLowerCase();
    const isNotDelivered = finalStatus === 'not delivered' || finalStatus === 'cancelled' || finalStatus === 'undelivered' || irctcStatus.includes('undelivered') || irctcStatus.includes('cancel') || rfStatus.includes('undelivered') || rfStatus.includes('cancel');
    const isDelivered = finalStatus === 'delivered' || finalStatus === 'success';

    if (isNotDelivered) st.notDeliveredOrdersCount += 1;
    if (!isDelivered) return;

    st.vendorPrice += toNumber(getVal(row, ['Final Vendor Price', 'Vendor Price']));
    st.finalBasePrice += toNumber(getVal(row, ['Final Base Price', 'Base Price']));
    st.totalComm += toNumber(getVal(row, ['Final Total Commission', 'Total Commission']));
    st.irctcComm += toNumber(getVal(row, ['Final IRCTC Commission', 'Final IRCTC Comm', 'IRCTC Comm']));
    st.rfComm += toNumber(getVal(row, ['Final RF Commission', 'RF Commission']));
    st.gst += toNumber(getVal(row, ['Final GST', 'GST']));
    st.totalDiscount += toNumber(getVal(row, ['Final Total Discount', 'Final Discount', 'Discount']));
    st.vendorDiscount += toNumber(getVal(row, ['Final Vendor Discount', 'Vendor Discount']));
    st.rfDiscount += toNumber(getVal(row, ['Final RF Discount', 'RF Discount']));
    st.deliveryCharges += toNumber(getVal(row, ['Delivery Charges', 'Delivery Charge']));
    st.sellingPrice += toNumber(getVal(row, ['Final Selling Price', 'Selling Price']));
    st.orderTotal += toNumber(getVal(row, ['Final Order Total', 'Order Total']));
    st.discountedBase += toNumber(getVal(row, ['Discounted Base Price']));
    st.ppd += toNumber(getVal(row, ['PPD', 'ppd']));
    st.cod += toNumber(getVal(row, ['COD', 'cod']));
    st.meals += toNumber(getVal(row, ['Meals', 'Meal Count'])) || 1;
    st.deliveredOrdersCount += toNumber(getVal(row, ['Orders Count'])) || 1;

    const outletId = String(getVal(row, ['Outlet ID', 'OutletId', 'Aggregator Outlet ID']) || '').trim();
    if (outletId) st.deliveredOutlets.add(outletId);
  });

  // 4) Station order = Station Rank. Business Rank is a separate rank based on delivered Final Base Price.
  const stationRows = Object.values(stationMap).sort((a: any, b: any) => {
    const ar = Number.isFinite(Number(a.stationRank)) ? Number(a.stationRank) : Number.MAX_SAFE_INTEGER;
    const br = Number.isFinite(Number(b.stationRank)) ? Number(b.stationRank) : Number.MAX_SAFE_INTEGER;
    if (ar !== br) return ar - br;
    return String(a.stationCode).localeCompare(String(b.stationCode), undefined, { numeric: true });
  });

  const businessRankOrder = [...stationRows].sort((a: any, b: any) => {
    const bp = Number(b.finalBasePrice || 0) - Number(a.finalBasePrice || 0);
    if (bp !== 0) return bp;
    return String(a.stationCode).localeCompare(String(b.stationCode), undefined, { numeric: true });
  });
  const businessRankMap: Record<string, number> = {};
  businessRankOrder.forEach((st: any, index) => { businessRankMap[st.stationCode] = index + 1; });

  return stationRows.map((st: any): StationReportRow => {
    const vPrice = Number(st.vendorPrice.toFixed(2));
    const bPrice = Number(st.finalBasePrice.toFixed(2));
    const tComm = Number(st.totalComm.toFixed(2));
    const sPrice = Number(st.sellingPrice.toFixed(2));
    const ppdVal = Number(st.ppd.toFixed(2));
    const checkPct = bPrice > 0 ? `${((tComm / bPrice) * 100).toFixed(2)}%` : '0.00%';
    const notDeliveredPct = st.deliveredOrdersCount > 0 ? `${((st.notDeliveredOrdersCount / st.deliveredOrdersCount) * 100).toFixed(2)}%` : st.notDeliveredOrdersCount > 0 ? '100.00%' : '0.00%';
    const ppdPct = sPrice > 0 ? `${((ppdVal / sPrice) * 100).toFixed(2)}%` : '0.00%';
    const totalVendors = stationVendors[st.stationCode]?.size || 0;
    const feedback = feedbackMap[st.stationCode] || { good: 0, bad: 0 };

    return {
      'Station Code': st.stationCode,
      'Rank': businessRankMap[st.stationCode] || 0,
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
      'Meals': Number(st.meals.toFixed(2)),
      'Check': checkPct,
      'Count of Delivered Orders': st.deliveredOrdersCount,
      'Not Delivered Order': st.notDeliveredOrdersCount,
      'Not Delivered %': notDeliveredPct,
      'PPD % of Final Selling Price': ppdPct,
      'Feedback Good': feedback.good,
      'Feedback Bad': feedback.bad,
      'Count of Delivered Outlets': st.deliveredOutlets.size,
      'Total Station Vendors': totalVendors,
    };
  });
};

export const generateStationReportWorkbook = (
  masterOrders: any[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = [],
  fileNamePrefix = 'STATION_WISE_REPORT'
) => {
  const stationData = generateStationWiseData(masterOrders, outletsMasterInfo, irctcOrders);
  if (!stationData.length) {
    alert('Station Report ke liye Outlet Master mein Station Code available nahi hai.');
    return;
  }

  const columns = Object.keys(stationData[0]);
  const total: Record<string, any> = {};
  columns.forEach((c) => { total[c] = ''; });
  total['Station Code'] = 'TOTAL';
  total['Station Name'] = `${stationData.length} Stations`;
  const nonNumeric = new Set(['Station Code','Station Name','Check','Not Delivered %','PPD % of Final Selling Price']);
  columns.forEach((c) => {
    if (!nonNumeric.has(c)) total[c] = stationData.reduce((sum, r: any) => sum + toNumber(r[c]), 0);
  });
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
    wch: c === 'Station Code' ? 12 : c === 'Station Name' ? 30 : c === 'Rank' || c === 'Station Rank' ? 12 : 16,
  }));

  // SheetJS community builds may not render styles in every viewer, but keep styles on cells for supported Excel viewers.
  const setStyle = (addr: string, style: any) => { if (ws[addr]) ws[addr].s = style; };
  const headerStyle = { font: { bold: true, color: { rgb: '1F2937' } }, fill: { fgColor: { rgb: 'E2E8F0' } }, alignment: { horizontal: 'center', vertical: 'center' } };
  const totalStyle = { font: { bold: true, color: { rgb: '111827' } }, fill: { fgColor: { rgb: 'CBD5E1' } }, alignment: { horizontal: 'center', vertical: 'center' } };
  columns.forEach((_, c) => {
    setStyle(`${XLSX.utils.encode_col(c)}1`, totalStyle);
    setStyle(`${XLSX.utils.encode_col(c)}2`, headerStyle);
  });
  stationData.forEach((r: any, i) => {
    const rowNum = i + 3;
    const base = toNumber(r['Final Base Price']);
    const businessRank = toNumber(r['Rank']);
    const stationRank = toNumber(r['Station Rank']);
    const rgb = base === 0 ? 'DC2626' : (Number.isFinite(stationRank) && businessRank <= stationRank ? '166534' : '78350F');
    columns.forEach((_, c) => setStyle(`${XLSX.utils.encode_col(c)}${rowNum}`, { font: { color: { rgb } } }));
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Station Report');
  XLSX.writeFile(wb, `${fileNamePrefix}_${new Date().toISOString().slice(0,10)}.xlsx`);
};
