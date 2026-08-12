import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// 1. AUTO-CLEANERS & NORMALIZERS
// ---------------------------------------------------------------------------

/**
 * Har tarah ki gandi Order ID ko standard pure number/clean string me badalta hai.
 * Handles: 12345.0, " 12345 ", "ORD-12345", 12345n, "IRCTC_12345"
 */
const autoCleanOrderId = (val: any): string => {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  
  // Agar scientific notation ya .0 float string ho (e.g. "12345.0")
  if (str.endsWith('.0')) {
    str = str.slice(0, -2);
  }
  
  // Strip out everything except alphanumeric
  const clean = str.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

  // Agar IRCTC / ORD prefix ho toh core digits/letters extract karna
  const digitsOnly = clean.replace(/\D/g, '');
  if (digitsOnly.length >= 4) {
    return digitsOnly; // Core numeric ID return karega
  }

  return clean;
};

/**
 * Outlet IDs clean helper
 */
const autoCleanOutletId = (val: any): string => {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/[^a-zA-Z0-9]/g, '');
};

/**
 * Safely parse any dirty number (e.g., "1,200.50", " ₹500 ", null)
 */
const parseSafeNumber = (val: any, fallback = 0): number => {
  if (val === null || val === undefined || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const sanitized = String(val).replace(/[^0-9.-]/g, '').trim();
  const num = parseFloat(sanitized);
  return isNaN(num) ? fallback : num;
};

/**
 * Universal date normalizer (Excel Serial, YYYY-MM-DD, DD/MM/YYYY, ISO string)
 */
const normalizeToDeliveryDate = (dateVal: any): string => {
  if (!dateVal) return '';

  // Excel Serial date handler
  if (typeof dateVal === 'number' || (!isNaN(Number(dateVal)) && !String(dateVal).includes('-') && !String(dateVal).includes('/'))) {
    const serial = Number(dateVal);
    if (serial > 30000 && serial < 60000) {
      const utcDays = serial - 25569;
      const dateObj = new Date(utcDays * 86400 * 1000);
      return `${dateObj.getUTCDate()}/${dateObj.getUTCMonth() + 1}/${dateObj.getUTCFullYear()}`;
    }
  }

  const str = String(dateVal).trim();

  // YYYY-MM-DD
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

  // DD/MM/YYYY
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

/**
 * Smart Value Finder: Row me se keywords ke base par value dhoondta hai
 */
const findValueByKeywords = (row: any, keywords: string[]): any => {
  if (!row || typeof row !== 'object') return null;
  const keys = Object.keys(row);
  for (const kw of keywords) {
    const matchedKey = keys.find((k) => k.trim().toLowerCase().replace(/[^a-z0-9]/g, '') === kw.toLowerCase().replace(/[^a-z0-9]/g, ''));
    if (matchedKey && row[matchedKey] !== undefined && row[matchedKey] !== null && String(row[matchedKey]).trim() !== '') {
      return row[matchedKey];
    }
  }
  return null;
};

// ---------------------------------------------------------------------------
// 2. MAIN REPORT GENERATOR
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

  // 1. Find Maximum / Last Delivery Date
  let maxTimestamp = -1;
  let lastDeliveryDateStr = '';

  masterData.forEach((row) => {
    const rawDate = findValueByKeywords(row, ['Delivery Date', 'DeliveryDate', 'Booking Date', 'BookingDate', 'Order Date', 'Date']);
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
    const rawDate = findValueByKeywords(row, ['Delivery Date', 'DeliveryDate', 'Booking Date', 'BookingDate', 'Order Date', 'Date']);
    return normalizeToDeliveryDate(rawDate) === lastDeliveryDateStr;
  });

  if (lastDayRows.length === 0) {
    alert(`Last date (${lastDeliveryDateStr}) par koi records nahi mile.`);
    return;
  }

  // 3. Build Mapping Lookup Tables
  const orderToStationMap: Record<string, string> = {};
  const outletToStationMap: Record<string, string> = {};
  const stationTotalOutletsMap: Record<string, Set<string>> = {};

  // Master Outlets
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const stCode = String(out?.station || out?.stationCode || out?.stn_code || '').trim().toUpperCase();
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (stCode && outId) {
      outletToStationMap[autoCleanOutletId(outId)] = stCode;
      if (!stationTotalOutletsMap[stCode]) stationTotalOutletsMap[stCode] = new Set<string>();
      stationTotalOutletsMap[stCode].add(outId);
    }
  });

  // Master Data records se Order ID -> Station Code Mapping
  masterData.forEach((r) => {
    const rawOrderId = findValueByKeywords(r, ['Order ID', 'OrderId', 'IRCTC Order ID', 'Booking ID', 'Order No', 'OrderNumber']);
    const stCode = String(findValueByKeywords(r, ['Station Code', 'StationCode', 'Station', 'stn_code']) || '').trim().toUpperCase();
    const rawOutletId = findValueByKeywords(r, ['Outlet ID', 'OutletId', 'Restaurant ID', 'restaurant_id']);

    const cleanOrd = autoCleanOrderId(rawOrderId);
    const cleanOut = autoCleanOutletId(rawOutletId);

    if (cleanOrd && stCode) {
      orderToStationMap[cleanOrd] = stCode;
    }
    if (cleanOut && stCode && !outletToStationMap[cleanOut]) {
      outletToStationMap[cleanOut] = stCode;
    }
  });

  // 4. Feedback & Complaint Auto-Match Engine
  const feedbackStationMap: Record<string, { goodCount: number; badCount: number }> = {};

  if (Array.isArray(feedbackData) && feedbackData.length > 0) {
    feedbackData.forEach((fb) => {
      const rawOrderId = findValueByKeywords(fb, ['Order ID', 'OrderId', 'IRCTC Order ID', 'Booking ID', 'Order No', 'OrderNumber', 'IRCTC_ID']);
      const rawOutletId = findValueByKeywords(fb, ['Outlet ID', 'OutletId', 'Restaurant ID', 'outlet_id']);
      const rawStation = findValueByKeywords(fb, ['Station Code', 'StationCode', 'Station', 'STATION CODE', 'stn_code']);

      const cleanOrd = autoCleanOrderId(rawOrderId);
      const cleanOut = autoCleanOutletId(rawOutletId);

      // Robust Station Code Resolution
      let stCode = '';
      if (cleanOrd && orderToStationMap[cleanOrd]) {
        stCode = orderToStationMap[cleanOrd];
      } else if (cleanOrd) {
        // Fallback: Agar exact match na mile to check partial sub-string
        const matchedKey = Object.keys(orderToStationMap).find(k => k.includes(cleanOrd) || cleanOrd.includes(k));
        if (matchedKey) stCode = orderToStationMap[matchedKey];
      }

      if (!stCode && cleanOut && outletToStationMap[cleanOut]) {
        stCode = outletToStationMap[cleanOut];
      }

      if (!stCode && rawStation) {
        stCode = String(rawStation).trim().toUpperCase();
      }

      if (!stCode) return; // Agar mapping kisi bhi tarike se na mile

      if (!feedbackStationMap[stCode]) {
        feedbackStationMap[stCode] = { goodCount: 0, badCount: 0 };
      }

      // Check Type / Status / Feedback Description
      const typeVal = String(
        findValueByKeywords(fb, ['Type', 'type', 'Feedback Type', 'feedback_type', 'Nature', 'Category', 'Feedback', 'Status', 'Remark', 'Complaint Type']) || ''
      ).trim().toLowerCase();

      const rating = parseSafeNumber(findValueByKeywords(fb, ['Rating', 'rating', 'Stars', 'star_rating']), 0);

      const isComplaint =
        typeVal.includes('complaint') ||
        typeVal.includes('comp') ||
        typeVal.includes('bad') ||
        typeVal.includes('poor') ||
        typeVal.includes('negative') ||
        typeVal.includes('issue') ||
        (rating > 0 && rating <= 3);

      const isFeedback =
        typeVal.includes('feedback') ||
        typeVal.includes('good') ||
        typeVal.includes('pos') ||
        typeVal.includes('sat') ||
        rating >= 4;

      if (isComplaint) {
        feedbackStationMap[stCode].badCount += 1;
      } else if (isFeedback) {
        feedbackStationMap[stCode].goodCount += 1;
      }
    });
  }

  // 5. Aggregate Station Level Data
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
    const stationCode = String(findValueByKeywords(r, ['Station Code', 'StationCode', 'Station', 'stn_code']) || 'UNKNOWN').trim().toUpperCase();
    const stationName = String(findValueByKeywords(r, ['Station Name', 'StationName', 'Station']) || stationCode).trim().toUpperCase();
    const outletId = String(findValueByKeywords(r, ['Outlet ID', 'OutletId', 'outlet_id']) || '').trim();

    const finalStatus = String(findValueByKeywords(r, ['Final Status', 'final_status', 'Status', 'status']) || '').trim().toLowerCase();
    const irctcStatus = String(findValueByKeywords(r, ['IRCTC Status', 'irctc_status']) || '').trim().toLowerCase();
    const rfStatus = String(findValueByKeywords(r, ['RF Status', 'rf_status']) || '').trim().toLowerCase();

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
      st.vendorPrice += parseSafeNumber(findValueByKeywords(r, ['Final Vendor Price', 'Vendor Price', 'vendor_price']));
      st.finalBasePrice += parseSafeNumber(findValueByKeywords(r, ['Final Base Price', 'Base Price', 'base_price']));
      st.finalTotalComm += parseSafeNumber(findValueByKeywords(r, ['Final Total Commission', 'Total Commission', 'commission']));
      st.finalIrctcComm += parseSafeNumber(findValueByKeywords(r, ['Final IRCTC Commission', 'IRCTC Comm', 'irctc_comm']));
      st.finalRfComm += parseSafeNumber(findValueByKeywords(r, ['Final RF Commission', 'RF Commission', 'rf_comm']));
      st.finalGst += parseSafeNumber(findValueByKeywords(r, ['Final GST', 'GST', 'gst']));
      st.finalDiscount += parseSafeNumber(findValueByKeywords(r, ['Final Total Discount', 'Discount', 'total_discount']));
      st.finalVendorDiscount += parseSafeNumber(findValueByKeywords(r, ['Final Vendor Discount', 'Vendor Discount']));
      st.finalRfDiscount += parseSafeNumber(findValueByKeywords(r, ['Final RF Discount', 'RF Discount']));
      st.deliveryCharges += parseSafeNumber(findValueByKeywords(r, ['Delivery Charges', 'delivery_charges']));
      st.finalSellingPrice += parseSafeNumber(findValueByKeywords(r, ['Final Selling Price', 'Selling Price', 'selling_price']));
      st.finalOrderTotal += parseSafeNumber(findValueByKeywords(r, ['Final Order Total', 'Order Total', 'order_total']));
      st.discountedBasePrice += parseSafeNumber(findValueByKeywords(r, ['Discounted Base Price', 'discounted_base_price']));
      st.ppd += parseSafeNumber(findValueByKeywords(r, ['PPD', 'ppd']));
      st.cod += parseSafeNumber(findValueByKeywords(r, ['COD', 'cod']));
      st.meals += parseInt(String(findValueByKeywords(r, ['Meals', 'meals']) || '1'), 10) || 1;
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

  // Summary Row (Top)
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
