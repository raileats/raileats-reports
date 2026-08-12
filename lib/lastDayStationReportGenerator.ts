import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// Strip all spaces & convert to standard string (e.g., 2460600201 or 2460600201.0 -> "2460600201")
const cleanOrderId = (val: any): string => {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  if (str.endsWith('.0')) {
    str = str.slice(0, -2);
  }
  return str.replace(/[^a-zA-Z0-9]/g, '');
};

const parseSafeNumber = (val: any, fallback = 0): number => {
  if (val === null || val === undefined || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const sanitized = String(val).replace(/[^0-9.-]/g, '').trim();
  const num = parseFloat(sanitized);
  return isNaN(num) ? fallback : num;
};

// Universal Date Normalizer (DD/MM/YYYY)
const normalizeToDeliveryDate = (dateVal: any): string => {
  if (!dateVal) return '';

  if (typeof dateVal === 'number' || (!isNaN(Number(dateVal)) && !String(dateVal).includes('-') && !String(dateVal).includes('/'))) {
    const serial = Number(dateVal);
    if (serial > 30000 && serial < 60000) {
      const utcDays = serial - 25569;
      const dateObj = new Date(utcDays * 86400 * 1000);
      return `${dateObj.getUTCDate()}/${dateObj.getUTCMonth() + 1}/${dateObj.getUTCFullYear()}`;
    }
  }

  const str = String(dateVal).trim();

  if (str.includes('-')) {
    const parts = str.split(' ')[0].split('T')[0].split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return `${parseInt(parts[2], 10)}/${parseInt(parts[1], 10)}/${parts[0]}`;
      } else {
        return `${parseInt(parts[0], 10)}/${parseInt(parts[1], 10)}/${parts[2]}`;
      }
    }
  }

  if (str.includes('/')) {
    const parts = str.split(' ')[0].split('/');
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        return `${parseInt(parts[0], 10)}/${parseInt(parts[1], 10)}/${parts[2]}`;
      }
    }
  }

  return str;
};

const dateToTimestamp = (dateStr: string): number => {
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    return new Date(year, month - 1, day).getTime();
  }
  return 0;
};

// Robust Key/Value Finder
const getValue = (row: any, keys: string[]): any => {
  if (!row || typeof row !== 'object') return null;
  const rowKeys = Object.keys(row);
  for (const k of keys) {
    const matched = rowKeys.find(
      (rk) => rk.trim().toLowerCase().replace(/[^a-z0-9]/g, '') === k.toLowerCase().replace(/[^a-z0-9]/g, '')
    );
    if (matched && row[matched] !== undefined && row[matched] !== null && String(row[matched]).trim() !== '') {
      return row[matched];
    }
  }
  return null;
};

// ---------------------------------------------------------------------------
// Main Generator
// ---------------------------------------------------------------------------

export const generateLastDayStationReportWorkbook = (
  masterData: any[],
  outletsMasterInfo: Record<string, any> = {},
  feedbackData: any[] = []
) => {
  if (!masterData || masterData.length === 0) {
    alert('Master Data khali hai! Pehle reports process karein.');
    return;
  }

  // 1. Find Max / Last Delivery Date from Master Data
  let maxTimestamp = -1;
  let lastDeliveryDateStr = '';

  masterData.forEach((row) => {
    const rawDate = getValue(row, ['Delivery Date', 'DeliveryDate', 'Booking Date', 'BookingDate', 'Order Date', 'Date']);
    const dStr = normalizeToDeliveryDate(rawDate);
    if (dStr) {
      const ts = dateToTimestamp(dStr);
      if (ts > maxTimestamp) {
        maxTimestamp = ts;
        lastDeliveryDateStr = dStr;
      }
    }
  });

  if (!lastDeliveryDateStr) {
    alert('Koi valid Delivery Date nahi mili!');
    return;
  }

  // 2. Filter Master rows for Last Day
  const lastDayRows = masterData.filter((row) => {
    const rawDate = getValue(row, ['Delivery Date', 'DeliveryDate', 'Booking Date', 'BookingDate', 'Order Date', 'Date']);
    return normalizeToDeliveryDate(rawDate) === lastDeliveryDateStr;
  });

  // 3. Create Order ID -> Station Code Mapping from Master / IRCTC Data
  const orderToStationMap: Record<string, string> = {};
  const stationTotalOutletsMap: Record<string, Set<string>> = {};

  // Master Outlets
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const stCode = String(out?.station || out?.stationCode || out?.stn_code || '').trim().toUpperCase();
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (stCode && outId) {
      if (!stationTotalOutletsMap[stCode]) stationTotalOutletsMap[stCode] = new Set<string>();
      stationTotalOutletsMap[stCode].add(outId);
    }
  });

  // Master Data se Order ID -> Station Code map
  masterData.forEach((r) => {
    const rawOrderId = getValue(r, ['Order ID', 'OrderId', 'IRCTC Order ID', 'Booking ID', 'Order No', 'OrderNumber']);
    const stCode = String(getValue(r, ['Station Code', 'StationCode', 'Station', 'stn_code']) || '').trim().toUpperCase();

    const cleanOrd = cleanOrderId(rawOrderId);
    if (cleanOrd && stCode) {
      orderToStationMap[cleanOrd] = stCode;
    }
  });

  // 4. Feedback & Complaint Logic (Strictly Order ID based -> Master Station Code)
  const feedbackStationMap: Record<string, { goodCount: number; badCount: number }> = {};

  if (Array.isArray(feedbackData) && feedbackData.length > 0) {
    feedbackData.forEach((fb) => {
      // Order ID from Feedback File (Column D)
      const rawOrderId = getValue(fb, ['Order ID', 'OrderId', 'Order Id', 'IRCTC Order ID', 'Booking ID', 'Order No', 'OrderNumber']);
      const cleanOrd = cleanOrderId(rawOrderId);

      if (!cleanOrd) return;

      // Match Order ID in Master/IRCTC to get Station Code
      const stCode = orderToStationMap[cleanOrd];
      if (!stCode) return; // Agar ye order Master me nahi mila toh skip

      if (!feedbackStationMap[stCode]) {
        feedbackStationMap[stCode] = { goodCount: 0, badCount: 0 };
      }

      // Check Type (Column L: feedback vs complaint)
      const typeVal = String(
        getValue(fb, ['Type', 'type', 'Feedback Type', 'feedback_type', 'Nature', 'Category', 'Feedback', 'Status']) || ''
      ).trim().toLowerCase();

      if (typeVal.includes('complaint') || typeVal.includes('complain') || typeVal.includes('bad') || typeVal.includes('issue')) {
        feedbackStationMap[stCode].badCount += 1;
      } else if (typeVal.includes('feedback') || typeVal.includes('good') || typeVal.includes('pos')) {
        feedbackStationMap[stCode].goodCount += 1;
      }
    });
  }

  // 5. Group Station Level Data for Last Day
  const stationMap: Record<
    string,
    {
      stationCode: string;
      stationName: string;
      vendorPrice: number;
      finalBasePrice: number;
      finalTotalComm: number;
      finalIrctcComm: number;
      finalRfComm: number;
      finalGst: number;
      finalDiscount: number;
      finalVendorDiscount: number;
      finalRfDiscount: number;
      deliveryCharges: number;
      finalSellingPrice: number;
      finalOrderTotal: number;
      discountedBasePrice: number;
      ppd: number;
      cod: number;
      meals: number;
      deliveredOrdersCount: number;
      notDeliveredOrdersCount: number;
      feedbackGoodCount: number;
      feedbackBadCount: number;
      deliveredOutlets: Set<string>;
      stationAllOutlets: Set<string>;
    }
  > = {};

  lastDayRows.forEach((r) => {
    const stationCode = String(getValue(r, ['Station Code', 'StationCode', 'Station', 'stn_code']) || 'UNKNOWN').trim().toUpperCase();
    const stationName = String(getValue(r, ['Station Name', 'StationName', 'Station']) || stationCode).trim().toUpperCase();
    const outletId = String(getValue(r, ['Outlet ID', 'OutletId', 'outlet_id']) || '').trim();

    const finalStatus = String(getValue(r, ['Final Status', 'final_status', 'Status', 'status']) || '').trim().toLowerCase();
    const irctcStatus = String(getValue(r, ['IRCTC Status', 'irctc_status']) || '').trim().toLowerCase();
    const rfStatus = String(getValue(r, ['RF Status', 'rf_status']) || '').trim().toLowerCase();

    const isDelivered = finalStatus === 'delivered' || finalStatus === 'success';
    const isNotDelivered =
      finalStatus === 'not delivered' ||
      finalStatus === 'cancelled' ||
      finalStatus === 'undelivered' ||
      irctcStatus.includes('undelivered') ||
      irctcStatus.includes('cancel') ||
      rfStatus.includes('undelivered') ||
      rfStatus.includes('cancel');

    if (!stationMap[stationCode]) {
      const fbCounts = feedbackStationMap[stationCode] || { goodCount: 0, badCount: 0 };
      stationMap[stationCode] = {
        stationCode,
        stationName,
        vendorPrice: 0,
        finalBasePrice: 0,
        finalTotalComm: 0,
        finalIrctcComm: 0,
        finalRfComm: 0,
        finalGst: 0,
        finalDiscount: 0,
        finalVendorDiscount: 0,
        finalRfDiscount: 0,
        deliveryCharges: 0,
        finalSellingPrice: 0,
        finalOrderTotal: 0,
        discountedBasePrice: 0,
        ppd: 0,
        cod: 0,
        meals: 0,
        deliveredOrdersCount: 0,
        notDeliveredOrdersCount: 0,
        feedbackGoodCount: fbCounts.goodCount,
        feedbackBadCount: fbCounts.badCount,
        deliveredOutlets: new Set<string>(),
        stationAllOutlets: stationTotalOutletsMap[stationCode]
          ? new Set<string>(stationTotalOutletsMap[stationCode])
          : new Set<string>(),
      };
    }

    const st = stationMap[stationCode];
    if (stationName && st.stationName === st.stationCode) {
      st.stationName = stationName;
    }

    if (isNotDelivered) {
      st.notDeliveredOrdersCount += 1;
    }

    if (isDelivered) {
      st.vendorPrice += parseSafeNumber(getValue(r, ['Final Vendor Price', 'Vendor Price', 'vendor_price']));
      st.finalBasePrice += parseSafeNumber(getValue(r, ['Final Base Price', 'Base Price', 'base_price']));
      st.finalTotalComm += parseSafeNumber(getValue(r, ['Final Total Commission', 'Total Commission', 'commission']));
      st.finalIrctcComm += parseSafeNumber(getValue(r, ['Final IRCTC Commission', 'IRCTC Comm', 'irctc_comm']));
      st.finalRfComm += parseSafeNumber(getValue(r, ['Final RF Commission', 'RF Commission', 'rf_comm']));
      st.finalGst += parseSafeNumber(getValue(r, ['Final GST', 'GST', 'gst']));
      st.finalDiscount += parseSafeNumber(getValue(r, ['Final Total Discount', 'Discount', 'total_discount']));
      st.finalVendorDiscount += parseSafeNumber(getValue(r, ['Final Vendor Discount', 'Vendor Discount']));
      st.finalRfDiscount += parseSafeNumber(getValue(r, ['Final RF Discount', 'RF Discount']));
      st.deliveryCharges += parseSafeNumber(getValue(r, ['Delivery Charges', 'delivery_charges']));
      st.finalSellingPrice += parseSafeNumber(getValue(r, ['Final Selling Price', 'Selling Price', 'selling_price']));
      st.finalOrderTotal += parseSafeNumber(getValue(r, ['Final Order Total', 'Order Total', 'order_total']));
      st.discountedBasePrice += parseSafeNumber(getValue(r, ['Discounted Base Price', 'discounted_base_price']));
      st.ppd += parseSafeNumber(getValue(r, ['PPD', 'ppd']));
      st.cod += parseSafeNumber(getValue(r, ['COD', 'cod']));
      st.meals += parseInt(String(getValue(r, ['Meals', 'meals']) || '1'), 10) || 1;
      st.deliveredOrdersCount += 1;

      if (outletId) {
        st.deliveredOutlets.add(outletId);
        st.stationAllOutlets.add(outletId);
      }
    }
  });

  // Sort descending by Final Base Price
  const sortedStations = Object.values(stationMap).sort(
    (a, b) => b.finalBasePrice - a.finalBasePrice
  );

  // Grand Totals Calculation
  let sumVendorPrice = 0;
  let sumFinalBasePrice = 0;
  let sumFinalTotalComm = 0;
  let sumFinalIrctcComm = 0;
  let sumFinalRfComm = 0;
  let sumFinalGst = 0;
  let sumFinalDiscount = 0;
  let sumFinalVendorDiscount = 0;
  let sumFinalRfDiscount = 0;
  let sumDeliveryCharges = 0;
  let sumFinalSellingPrice = 0;
  let sumFinalOrderTotal = 0;
  let sumDiscountedBasePrice = 0;
  let sumPpd = 0;
  let sumCod = 0;
  let sumMeals = 0;
  let sumDeliveredOrders = 0;
  let sumNotDeliveredOrders = 0;
  let sumGoodFeedback = 0;
  let sumBadFeedback = 0;
  let sumDeliveredOutlets = 0;
  let sumStationVendors = 0;

  sortedStations.forEach((st) => {
    sumVendorPrice += st.vendorPrice;
    sumFinalBasePrice += st.finalBasePrice;
    sumFinalTotalComm += st.finalTotalComm;
    sumFinalIrctcComm += st.finalIrctcComm;
    sumFinalRfComm += st.finalRfComm;
    sumFinalGst += st.finalGst;
    sumFinalDiscount += st.finalDiscount;
    sumFinalVendorDiscount += st.finalVendorDiscount;
    sumFinalRfDiscount += st.finalRfDiscount;
    sumDeliveryCharges += st.deliveryCharges;
    sumFinalSellingPrice += st.finalSellingPrice;
    sumFinalOrderTotal += st.finalOrderTotal;
    sumDiscountedBasePrice += st.discountedBasePrice;
    sumPpd += st.ppd;
    sumCod += st.cod;
    sumMeals += st.meals;
    sumDeliveredOrders += st.deliveredOrdersCount;
    sumNotDeliveredOrders += st.notDeliveredOrdersCount;
    sumGoodFeedback += st.feedbackGoodCount;
    sumBadFeedback += st.feedbackBadCount;
    sumDeliveredOutlets += st.deliveredOutlets.size;
    sumStationVendors += st.stationAllOutlets.size;
  });

  const overallCheckPct =
    sumFinalBasePrice > 0
      ? `${((sumFinalTotalComm / sumFinalBasePrice) * 100).toFixed(2)}%`
      : '0.00%';

  const overallNotDeliveredPct =
    sumDeliveredOrders > 0
      ? `${((sumNotDeliveredOrders / sumDeliveredOrders) * 100).toFixed(2)}%`
      : sumNotDeliveredOrders > 0
      ? '100.00%'
      : '0.00%';

  const overallPpdPct =
    sumFinalSellingPrice > 0
      ? `${((sumPpd / sumFinalSellingPrice) * 100).toFixed(2)}%`
      : '0.00%';

  // Summary Top Row
  const topSummaryRow = [
    '', '', '', '',
    Number(sumVendorPrice.toFixed(2)),
    Number(sumFinalBasePrice.toFixed(2)),
    Number(sumFinalTotalComm.toFixed(2)),
    Number(sumFinalIrctcComm.toFixed(2)),
    Number(sumFinalRfComm.toFixed(2)),
    Number(sumFinalGst.toFixed(2)),
    Number(sumFinalDiscount.toFixed(2)),
    Number(sumFinalVendorDiscount.toFixed(2)),
    Number(sumFinalRfDiscount.toFixed(2)),
    Number(sumDeliveryCharges.toFixed(2)),
    Number(sumFinalSellingPrice.toFixed(2)),
    Number(sumFinalOrderTotal.toFixed(2)),
    Number(sumDiscountedBasePrice.toFixed(2)),
    Number(sumPpd.toFixed(2)),
    Number(sumCod.toFixed(2)),
    sumMeals,
    overallCheckPct,
    sumDeliveredOrders,
    sumNotDeliveredOrders,
    overallNotDeliveredPct,
    overallPpdPct,
    sumGoodFeedback,
    sumBadFeedback,
    sumDeliveredOutlets,
    sumStationVendors,
  ];

  // Headers
  const headers = [
    'Station Code',
    'Rank',
    'Delivery Date',
    'Station Name',
    'Vendor Price',
    'Final Base Price',
    'Final Total Commission',
    'Final IRCTC Comm',
    'Final RF Commission',
    'Final GST',
    'Final Discount',
    'Final Vendor Discount',
    'Final RF Discount',
    'Delivery Charges',
    'Final Selling Price',
    'Final Order Total',
    'Discounted Base Price',
    'PPD',
    'COD',
    'Meals',
    'Check',
    'Count of Delivered Orders',
    'Not Delivered Order',
    'Not Delivered %',
    'PPD % of Final Selling Price',
    'Feedback Good',
    'Feedback Bad',
    'Count of Delivered Outlets',
    'Total Station Vendors',
  ];

  // Rows Data
  const rowsData: any[][] = sortedStations.map((st, index) => {
    const vPrice = Number(st.vendorPrice.toFixed(2));
    const bPrice = Number(st.finalBasePrice.toFixed(2));
    const tComm = Number(st.finalTotalComm.toFixed(2));
    const sPrice = Number(st.finalSellingPrice.toFixed(2));
    const ppdVal = Number(st.ppd.toFixed(2));

    const checkPct = bPrice > 0 ? `${((tComm / bPrice) * 100).toFixed(2)}%` : '#DIV/0!';
    const notDeliveredPct =
      st.deliveredOrdersCount > 0
        ? `${((st.notDeliveredOrdersCount / st.deliveredOrdersCount) * 100).toFixed(2)}%`
        : st.notDeliveredOrdersCount > 0
        ? '100.00%'
        : '#DIV/0!';
    const ppdPct = sPrice > 0 ? `${((ppdVal / sPrice) * 100).toFixed(2)}%` : '#DIV/0!';

    return [
      st.stationCode,
      index + 1,
      lastDeliveryDateStr,
      st.stationName,
      vPrice,
      bPrice,
      tComm,
      Number(st.finalIrctcComm.toFixed(2)),
      Number(st.finalRfComm.toFixed(2)),
      Number(st.finalGst.toFixed(2)),
      Number(st.finalDiscount.toFixed(2)),
      Number(st.finalVendorDiscount.toFixed(2)),
      Number(st.finalRfDiscount.toFixed(2)),
      Number(st.deliveryCharges.toFixed(2)),
      sPrice,
      Number(st.finalOrderTotal.toFixed(2)),
      Number(st.discountedBasePrice.toFixed(2)),
      ppdVal,
      Number(st.cod.toFixed(2)),
      st.meals,
      checkPct,
      st.deliveredOrdersCount,
      st.notDeliveredOrdersCount,
      notDeliveredPct,
      ppdPct,
      st.feedbackGoodCount,
      st.feedbackBadCount,
      st.deliveredOutlets.size,
      st.stationAllOutlets.size,
    ];
  });

  const fullSheetData = [topSummaryRow, headers, ...rowsData];
  const worksheet = XLSX.utils.aoa_to_sheet(fullSheetData);

  worksheet['!cols'] = [
    { wch: 14 }, { wch: 8 },  { wch: 14 }, { wch: 24 },
    { wch: 14 }, { wch: 16 }, { wch: 22 }, { wch: 16 },
    { wch: 18 }, { wch: 12 }, { wch: 14 }, { wch: 19 },
    { wch: 16 }, { wch: 15 }, { wch: 18 }, { wch: 16 },
    { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 10 },
    { wch: 12 }, { wch: 24 }, { wch: 20 }, { wch: 16 },
    { wch: 26 }, { wch: 16 }, { wch: 16 }, { wch: 26 },
    { wch: 22 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Last Day Station Report');

  XLSX.writeFile(
    workbook,
    `LAST_DAY_STATION_REPORT_${lastDeliveryDateStr.replace(/\//g, '-')}.xlsx`
  );
};
