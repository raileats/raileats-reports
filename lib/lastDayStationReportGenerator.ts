import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Safely parse numbers from dirty strings, floats, and comma formatted numbers
 */
const parseSafeNumber = (val: any, fallback = 0): number => {
  if (val === null || val === undefined || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const sanitized = String(val).replace(/,/g, '').trim();
  const num = parseFloat(sanitized);
  return isNaN(num) ? fallback : num;
};

/**
 * Universal date normalizer supporting Excel serial numbers, ISO strings, YYYY-MM-DD, DD/MM/YYYY
 */
const normalizeToDeliveryDate = (dateVal: any): string => {
  if (!dateVal) return '';

  // Handle Excel Serial Number (e.g. 45280)
  if (typeof dateVal === 'number' || (!isNaN(Number(dateVal)) && !String(dateVal).includes('-') && !String(dateVal).includes('/'))) {
    const serial = Number(dateVal);
    if (serial > 30000 && serial < 60000) {
      const utcDays = serial - 25569;
      const dateObj = new Date(utcDays * 86400 * 1000);
      return `${dateObj.getUTCDate()}/${dateObj.getUTCMonth() + 1}/${dateObj.getUTCFullYear()}`;
    }
  }

  const str = String(dateVal).trim();

  // YYYY-MM-DD or DD-MM-YYYY
  if (str.includes('-')) {
    const parts = str.split(' ')[0].split('T')[0].split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD
        return `${parseInt(parts[2], 10)}/${parseInt(parts[1], 10)}/${parts[0]}`;
      } else {
        // DD-MM-YYYY
        return `${parseInt(parts[0], 10)}/${parseInt(parts[1], 10)}/${parts[2]}`;
      }
    }
  }

  // DD/MM/YYYY or MM/DD/YYYY
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

/**
 * Convert normalized "D/M/YYYY" to timestamp for accurate max date comparison
 */
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

// ---------------------------------------------------------------------------
// Main Generator Function
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

  // 1. Find the Maximum/Last Delivery Date across all records
  let maxTimestamp = -1;
  let lastDeliveryDateStr = '';

  masterData.forEach((row) => {
    const dStr = normalizeToDeliveryDate(
      row['Delivery Date'] || row['delivery_date'] || row['Booking Date'] || row['booking_date'] || ''
    );
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

  // 2. Filter rows belonging ONLY to the last delivery date
  const lastDayRows = masterData.filter((row) => {
    const dStr = normalizeToDeliveryDate(
      row['Delivery Date'] || row['delivery_date'] || row['Booking Date'] || row['booking_date'] || ''
    );
    return dStr === lastDeliveryDateStr;
  });

  if (lastDayRows.length === 0) {
    alert(`Last date (${lastDeliveryDateStr}) par koi records nahi mile.`);
    return;
  }

  // Station Code to Outlet mapping helper from master outlets
  const outletToStationMap: Record<string, string> = {};
  const stationTotalOutletsMap: Record<string, Set<string>> = {};

  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const stCode = String(out?.station || out?.stationCode || '').trim().toUpperCase();
    const outId = String(out?.outletId || out?.restaurantId || '').trim();
    if (stCode && outId) {
      outletToStationMap[outId] = stCode;
      if (!stationTotalOutletsMap[stCode]) {
        stationTotalOutletsMap[stCode] = new Set<string>();
      }
      stationTotalOutletsMap[stCode].add(outId);
    }
  });

  masterData.forEach((r) => {
    const outId = String(r['Outlet ID'] || r['outlet_id'] || '').trim();
    const stCode = String(r['Station Code'] || r['station_code'] || '').trim().toUpperCase();
    if (outId && stCode && !outletToStationMap[outId]) {
      outletToStationMap[outId] = stCode;
    }
  });

  // 3. Process Feedback Data (Station Code Wise on Last Delivery Date)
  const feedbackStationMap: Record<string, { goodCount: number; badCount: number }> = {};

  if (Array.isArray(feedbackData) && feedbackData.length > 0) {
    feedbackData.forEach((fb) => {
      const fbDateStr = normalizeToDeliveryDate(
        fb['Delivery Date'] ||
        fb['Delivery date'] ||
        fb['Order Date'] ||
        fb['Feedback Date'] ||
        fb['Date'] ||
        ''
      );

      // Match last day date if date exists in feedback row
      if (fbDateStr && fbDateStr !== lastDeliveryDateStr) {
        return;
      }

      let stCode = String(
        fb['Station Code'] ||
        fb['Station'] ||
        fb['STATION CODE'] ||
        fb['stn_code'] ||
        ''
      ).trim().toUpperCase();

      const fbOutletId = String(fb['Outlet ID'] || fb['outlet_id'] || fb['Restaurant ID'] || '').trim();
      if (!stCode && fbOutletId && outletToStationMap[fbOutletId]) {
        stCode = outletToStationMap[fbOutletId];
      }

      if (!stCode) return;

      if (!feedbackStationMap[stCode]) {
        feedbackStationMap[stCode] = { goodCount: 0, badCount: 0 };
      }

      const rating = parseSafeNumber(fb['Rating'] || fb['rating'] || fb['Stars'], 0);
      const feedbackType = String(
        fb['Feedback Type'] ||
        fb['Feedback'] ||
        fb['Status'] ||
        fb['Customer Feedback'] ||
        fb['Remark'] ||
        ''
      ).trim().toLowerCase();

      const isGood =
        rating >= 4 ||
        feedbackType.includes('good') ||
        feedbackType.includes('positive') ||
        feedbackType.includes('excellent') ||
        feedbackType.includes('satisfied');

      const isBad =
        (rating > 0 && rating <= 3) ||
        feedbackType.includes('bad') ||
        feedbackType.includes('negative') ||
        feedbackType.includes('poor') ||
        feedbackType.includes('complaint') ||
        feedbackType.includes('dissatisfied');

      if (isGood) {
        feedbackStationMap[stCode].goodCount += 1;
      } else if (isBad) {
        feedbackStationMap[stCode].badCount += 1;
      }
    });
  }

  // 4. Group Master Data by Station Code
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
    const stationCode = String(r['Station Code'] || r['station_code'] || 'UNKNOWN').trim().toUpperCase();
    const stationName = String(r['Station Name'] || r['station_name'] || r['Station'] || stationCode).trim().toUpperCase();
    const outletId = String(r['Outlet ID'] || r['outlet_id'] || '').trim();

    const finalStatus = String(r['Final Status'] || r['status'] || '').trim().toLowerCase();
    const irctcStatus = String(r['IRCTC Status'] || '').trim().toLowerCase();
    const rfStatus = String(r['RF Status'] || '').trim().toLowerCase();

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
      st.vendorPrice += parseSafeNumber(r['Final Vendor Price'] || r['Vendor Price']);
      st.finalBasePrice += parseSafeNumber(r['Final Base Price'] || r['Base Price']);
      st.finalTotalComm += parseSafeNumber(r['Final Total Commission'] || r['Total Commission']);
      st.finalIrctcComm += parseSafeNumber(r['Final IRCTC Commission'] || r['IRCTC Comm']);
      st.finalRfComm += parseSafeNumber(r['Final RF Commission'] || r['RF Commission']);
      st.finalGst += parseSafeNumber(r['Final GST'] || r['GST']);
      st.finalDiscount += parseSafeNumber(r['Final Total Discount'] || r['Discount']);
      st.finalVendorDiscount += parseSafeNumber(r['Final Vendor Discount'] || r['Vendor Discount']);
      st.finalRfDiscount += parseSafeNumber(r['Final RF Discount'] || r['RF Discount']);
      st.deliveryCharges += parseSafeNumber(r['Delivery Charges']);
      st.finalSellingPrice += parseSafeNumber(r['Final Selling Price'] || r['Selling Price']);
      st.finalOrderTotal += parseSafeNumber(r['Final Order Total'] || r['Order Total']);
      st.discountedBasePrice += parseSafeNumber(r['Discounted Base Price']);
      st.ppd += parseSafeNumber(r['PPD']);
      st.cod += parseSafeNumber(r['COD']);
      st.meals += parseInt(String(r['Meals'] || '1'), 10) || 1;
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

  // Row 1: Top Summary Row
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

  // Row 2: Headers
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

  // Build rows data
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
