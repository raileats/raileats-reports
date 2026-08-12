import * as XLSX from 'xlsx';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const cleanId = (val: any): string => {
  if (val === null || val === undefined) return '';
  return String(val).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
};

const parseSafeNumber = (val: any, fallback = 0): number => {
  if (val === null || val === undefined || val === '') return fallback;
  if (typeof val === 'number') return isNaN(val) ? fallback : val;
  const sanitized = String(val).replace(/,/g, '').trim();
  const num = parseFloat(sanitized);
  return isNaN(num) ? fallback : num;
};

// Universal date normalizer (Excel Serial, YYYY-MM-DD, DD/MM/YYYY)
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

  console.log(`[Station Report] Total Master Rows: ${masterData.length}, Feedback Rows: ${feedbackData?.length || 0}`);

  // 1. Find Maximum/Last Delivery Date
  let maxTimestamp = -1;
  let lastDeliveryDateStr = '';

  masterData.forEach((row) => {
    const rawDate =
      row['Delivery Date'] ||
      row['delivery_date'] ||
      row['Booking Date'] ||
      row['booking_date'] ||
      row['Order Date'] ||
      '';
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

  console.log(`[Station Report] Selected Last Delivery Date: ${lastDeliveryDateStr}`);

  // 2. Filter Master rows belonging ONLY to the last delivery date
  const lastDayRows = masterData.filter((row) => {
    const rawDate =
      row['Delivery Date'] ||
      row['delivery_date'] ||
      row['Booking Date'] ||
      row['booking_date'] ||
      row['Order Date'] ||
      '';
    const dStr = normalizeToDeliveryDate(rawDate);
    return dStr === lastDeliveryDateStr;
  });

  if (lastDayRows.length === 0) {
    alert(`Last date (${lastDeliveryDateStr}) par koi records nahi mile.`);
    return;
  }

  // 3. Mapping Maps:
  // - Clean Order ID -> Station Code
  // - Clean Outlet ID -> Station Code
  const orderToStationMap: Record<string, string> = {};
  const outletToStationMap: Record<string, string> = {};
  const stationTotalOutletsMap: Record<string, Set<string>> = {};

  // Master Outlets
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const stCode = String(out?.station || out?.stationCode || out?.stn_code || '').trim().toUpperCase();
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (stCode && outId) {
      outletToStationMap[cleanId(outId)] = stCode;
      if (!stationTotalOutletsMap[stCode]) {
        stationTotalOutletsMap[stCode] = new Set<string>();
      }
      stationTotalOutletsMap[stCode].add(outId);
    }
  });

  // Master Data Map (All rows mapping)
  masterData.forEach((r) => {
    const rawOrderId =
      r['Order ID'] ||
      r['Order Id'] ||
      r['order_id'] ||
      r['IRCTC Order ID'] ||
      r['Booking ID'] ||
      r['Order No'] ||
      r['order_no'] ||
      '';
    const stCode = String(
      r['Station Code'] ||
      r['station_code'] ||
      r['Station'] ||
      r['station'] ||
      ''
    ).trim().toUpperCase();
    const rawOutletId = String(r['Outlet ID'] || r['outlet_id'] || r['Restaurant ID'] || '').trim();

    if (rawOrderId && stCode) {
      orderToStationMap[cleanId(rawOrderId)] = stCode;
    }
    if (rawOutletId && stCode && !outletToStationMap[cleanId(rawOutletId)]) {
      outletToStationMap[cleanId(rawOutletId)] = stCode;
    }
  });

  // 4. Process Feedback / Complaint Table
  const feedbackStationMap: Record<string, { goodCount: number; badCount: number }> = {};
  let matchedFeedbacks = 0;

  if (Array.isArray(feedbackData) && feedbackData.length > 0) {
    feedbackData.forEach((fb) => {
      // Pick Order ID with all variations
      const rawFbOrderId =
        fb['Order ID'] ||
        fb['Order Id'] ||
        fb['order_id'] ||
        fb['Booking ID'] ||
        fb['booking_id'] ||
        fb['IRCTC Order ID'] ||
        fb['Order No'] ||
        fb['order_no'] ||
        fb['Order Number'] ||
        '';

      const rawFbOutletId =
        fb['Outlet ID'] ||
        fb['outlet_id'] ||
        fb['Restaurant ID'] ||
        fb['restaurant_id'] ||
        '';

      // Pick Station Code:
      // Priority 1: Match Order ID to Master
      // Priority 2: Match Direct Station Code in row
      // Priority 3: Match Outlet ID to Master
      let stCode = '';
      const cOrderId = cleanId(rawFbOrderId);
      const cOutletId = cleanId(rawFbOutletId);

      if (cOrderId && orderToStationMap[cOrderId]) {
        stCode = orderToStationMap[cOrderId];
      } else if (fb['Station Code'] || fb['station_code'] || fb['Station'] || fb['STATION CODE'] || fb['stn_code']) {
        stCode = String(fb['Station Code'] || fb['station_code'] || fb['Station'] || fb['STATION CODE'] || fb['stn_code']).trim().toUpperCase();
      } else if (cOutletId && outletToStationMap[cOutletId]) {
        stCode = outletToStationMap[cOutletId];
      }

      if (!stCode) return; // Unmapped row skip

      if (!feedbackStationMap[stCode]) {
        feedbackStationMap[stCode] = { goodCount: 0, badCount: 0 };
      }

      // Read Type / Feedback Text / Complaint columns
      const typeVal = String(
        fb['Type'] ||
        fb['type'] ||
        fb['TYPE'] ||
        fb['Feedback Type'] ||
        fb['feedback_type'] ||
        fb['Nature'] ||
        fb['Category'] ||
        fb['Feedback'] ||
        fb['feedback'] ||
        fb['Remark'] ||
        fb['Status'] ||
        ''
      ).trim().toLowerCase();

      const rating = parseSafeNumber(fb['Rating'] || fb['rating'] || fb['Stars'] || fb['star_rating'], 0);

      const isComplaint =
        typeVal.includes('complaint') ||
        typeVal.includes('comp') ||
        typeVal.includes('bad') ||
        typeVal.includes('negative') ||
        typeVal.includes('poor') ||
        typeVal.includes('issue') ||
        (rating > 0 && rating <= 3);

      const isGood =
        typeVal.includes('feedback') ||
        typeVal.includes('feed') ||
        typeVal.includes('good') ||
        typeVal.includes('positive') ||
        typeVal.includes('satisf') ||
        rating >= 4;

      if (isComplaint) {
        feedbackStationMap[stCode].badCount += 1;
        matchedFeedbacks++;
      } else if (isGood) {
        feedbackStationMap[stCode].goodCount += 1;
        matchedFeedbacks++;
      }
    });
  }

  console.log(`[Station Report] Total Matched Feedbacks/Complaints: ${matchedFeedbacks}`, feedbackStationMap);

  // 5. Group Last Day Rows by Station Code
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

    const finalStatus = String(r['Final Status'] || r['final_status'] || r['status'] || '').trim().toLowerCase();
    const irctcStatus = String(r['IRCTC Status'] || r['irctc_status'] || '').trim().toLowerCase();
    const rfStatus = String(r['RF Status'] || r['rf_status'] || '').trim().toLowerCase();

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

  // Summary Row (Top Row)
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

  // Headers Row
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
