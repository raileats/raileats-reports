import * as XLSX from 'xlsx';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

export interface DateWiseReportRow {
  'Delivery Date': string;
  'Total Orders': number;
  'Delivered': number;
  'Cancelled': number;
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

/**
 * Robust Date Normalizer Helper
 * Converts Excel Serial, DD/MM/YYYY, YYYY-MM-DD, ISO Dates to standardized "D/M/YYYY"
 */
export const normalizeToDayMonthYear = (rawDate: any): string | null => {
  if (rawDate === undefined || rawDate === null || rawDate === '') return null;

  const valid = (year: number, month: number, day: number): Date | null => {
    const d = new Date(year, month - 1, day);
    return d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day ? d : null;
  };

  // Excel serial.
  if (typeof rawDate === 'number' || (!isNaN(Number(rawDate)) && !String(rawDate).includes('-') && !String(rawDate).includes('/'))) {
    const num = Number(rawDate);
    if (num > 20000 && num < 60000) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(num) * 86400000);
      // XLSX may have converted 08/01/2026 -> Jan 8 2026, etc.
      if (d.getFullYear() === 2026 && d.getUTCDate() === 8 && d.getUTCMonth() >= 0 && d.getUTCMonth() <= 9) {
        return `${d.getUTCMonth() + 1}/8/2026`
          .replace(/^(\d+)\/8\/2026$/, (_, m) => `${m}/8/2026`);
      }
      return `${d.getUTCDate()}/${d.getUTCMonth() + 1}/${d.getUTCFullYear()}`;
    }
  }

  // XLSX Date corruption repair for the August-2026 source: Jan 8 -> Aug 1,
  // Feb 8 -> Aug 2, ... Oct 8 -> Aug 10.
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    const year = rawDate.getFullYear();
    const monthIndex = rawDate.getMonth();
    const day = rawDate.getDate();
    if (year === 2026 && day === 8 && monthIndex >= 0 && monthIndex <= 9) {
      return `${monthIndex + 1}/8/2026`;
    }
    return `${day}/${monthIndex + 1}/${year}`;
  }

  const str = String(rawDate).trim();
  const datePart = str.split(/[T ]/)[0];

  // ISO YYYY-MM-DD / YYYY/MM/DD.
  const ymd = datePart.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (ymd) {
    const year = Number(ymd[1]);
    const month = Number(ymd[2]);
    const day = Number(ymd[3]);
    return valid(year, month, day) ? `${day}/${month}/${year}` : null;
  }

  // IMPORTANT: Master Data/source convention is MM/DD/YYYY.
  // Example: 08/01/2026 = 1 August 2026. Never let JS Date decide this.
  const slash = datePart.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (slash) {
    const first = Number(slash[1]);
    const second = Number(slash[2]);
    const year = Number(slash[3]);

    if (first >= 1 && first <= 12) {
      const month = first;
      const day = second;
      const d = valid(year, month, day);
      return d ? `${day}/${month}/${year}` : null;
    }

    // Only use DD/MM fallback when the first part cannot be a month.
    if (second >= 1 && second <= 12) {
      const day = first;
      const month = second;
      const d = valid(year, month, day);
      return d ? `${day}/${month}/${year}` : null;
    }
    return null;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return `${parsed.getDate()}/${parsed.getMonth() + 1}/${parsed.getFullYear()}`;
  }

  return null;
};

/**
 * Generates Date Wise Summary Data
 */
export const generateDateWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = []
): DateWiseReportRow[] => {
  const dateOrdersMap: Record<string, {
    delivered: MasterOrderRow[];
    totalOrders: number;
    cancelledCount: number;
    notDeliveredCount: number;
    deliveredOutlets: Set<string>;
    feedbackGood: number;
    feedbackBad: number;
    stationCodes: Set<string>;
  }> = {};
  const allUniqueDates = new Set<string>();


  masterOrders.forEach((row) => {
    // Check multiple potential date field keys
    const rawDate =
      row['Delivery Date'] ||
      (row as any)['Delivery date'] ||
      (row as any)['DELIVERY DATE'] ||
      (row as any)['Order Date'] ||
      (row as any)['Date'];

    const normalizedDate = normalizeToDayMonthYear(rawDate);
    if (!normalizedDate) return;

    allUniqueDates.add(normalizedDate);


    if (!dateOrdersMap[normalizedDate]) {
      dateOrdersMap[normalizedDate] = {
        delivered: [],
        totalOrders: 0,
        cancelledCount: 0,
        notDeliveredCount: 0,
        deliveredOutlets: new Set<string>(),
        feedbackGood: 0,
        feedbackBad: 0,
        stationCodes: new Set<string>(),
      };
    }

    const finalStatus = String(row['Final Status'] || '').trim().toLowerCase();
    const irctcStatus = String(row['IRCTC Status'] || '').trim().toLowerCase();
    const rfStatus = String(row['RF Status'] || '').trim().toLowerCase();

    const outletId = String(
      (row as any)['Outlet ID'] || (row as any)['OutletId'] || (row as any)['outlet_id'] || ''
    ).trim();
    const rawStation = String(
      (row as any)['Station Code'] ||
      (row as any)['Delivery Station'] ||
      (row as any)['DeliveryStation'] ||
      (row as any)['Station Name'] ||
      (row as any)['Station'] ||
      ''
    ).trim();
    const cleanStation = rawStation
      .toUpperCase()
      .split('-')[0]
      .split('(')[0]
      .split('/')[0]
      .trim()
      .replace(/[^A-Z0-9]/g, '');

    if (cleanStation) dateOrdersMap[normalizedDate].stationCodes.add(cleanStation);

    dateOrdersMap[normalizedDate].totalOrders += 1;

    // Check Not Delivered / Cancelled status
    if (finalStatus === 'cancelled') {
      dateOrdersMap[normalizedDate].cancelledCount += 1;
    }

    if (
      finalStatus === 'not delivered' ||
      finalStatus === 'cancelled' ||
      finalStatus === 'undelivered' ||
      irctcStatus.includes('undelivered') ||
      irctcStatus.includes('cancel') ||
      rfStatus.includes('undelivered') ||
      rfStatus.includes('cancel')
    ) {
      dateOrdersMap[normalizedDate].notDeliveredCount += 1;
    }

    // Match delivered status (handles 'delivered', 'DELIVERED', 'delivered ', 'success')
    if (finalStatus === 'delivered' || finalStatus === 'success' || (!finalStatus && irctcStatus.includes('deliver'))) {
      dateOrdersMap[normalizedDate].delivered.push(row);
      if (outletId) dateOrdersMap[normalizedDate].deliveredOutlets.add(outletId);
    }
  });

  // Feedback uses the exact Station Report rule:
  // only Feedback Type = FEEDBACK / COMPLAIN is counted.
  const feedbackSource = Array.isArray(irctcOrders) && irctcOrders.length > 0 ? irctcOrders : masterOrders;
  feedbackSource.forEach((row: any) => {
    const rawDate =
      row['Delivery Date'] ||
      row['Delivery date'] ||
      row['DELIVERY DATE'] ||
      row['Order Date'] ||
      row['Date'];
    const normalizedDate = normalizeToDayMonthYear(rawDate);
    if (!normalizedDate || !dateOrdersMap[normalizedDate]) return;

    const typeVal = String(
      row['Feedback Type'] || row['FeedbackType'] || row['FEEDBACK TYPE'] || ''
    ).trim().toUpperCase();

    if (typeVal === 'FEEDBACK') {
      dateOrdersMap[normalizedDate].feedbackGood += 1;
    } else if (typeVal === 'COMPLAIN') {
      dateOrdersMap[normalizedDate].feedbackBad += 1;
    }
  });

  // IMPORTANT: Do NOT manufacture calendar days.
  // Date Wise must contain ONLY dates actually present in Master Data.
  const sortedDateList = Array.from(allUniqueDates).sort((a, b) => {
    const toTs = (key: string) => {
      const [day, month, year] = key.split('/').map(Number);
      return new Date(year, month - 1, day).getTime();
    };
    return toTs(a) - toTs(b);
  });

  const reportRows: DateWiseReportRow[] = [];

  sortedDateList.forEach((dateKey) => {
    const bucket = dateOrdersMap[dateKey] || {
      delivered: [],
      totalOrders: 0,
      cancelledCount: 0,
      notDeliveredCount: 0,
      deliveredOutlets: new Set<string>(),
      feedbackGood: 0,
      feedbackBad: 0,
      stationCodes: new Set<string>(),
    };
    const deliveredOrders = bucket.delivered;

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

    deliveredOrders.forEach((ord) => {
      vendorPriceSum += Number(ord['Final Vendor Price'] || (ord as any)['Vendor Price'] || 0);
      basePriceSum += Number(ord['Final Base Price'] || (ord as any)['Base Price'] || 0);
      totalCommSum += Number(ord['Final Total Commission'] || (ord as any)['Total Commission'] || 0);
      irctcCommSum += Number(ord['Final IRCTC Commission'] || (ord as any)['IRCTC Comm'] || 0);
      rfCommSum += Number(ord['Final RF Commission'] || (ord as any)['RF Commission'] || 0);
      gstSum += Number(ord['Final GST'] || (ord as any)['GST'] || 0);
      totalDiscountSum += Number(ord['Final Total Discount'] || (ord as any)['Total Discount'] || (ord as any)['Discount'] || 0);
      vendorDiscountSum += Number(ord['Final Vendor Discount'] || (ord as any)['Vendor Discount'] || 0);
      rfDiscountSum += Number(ord['Final RF Discount'] || (ord as any)['RF Discount'] || 0);
      deliveryChargesSum += Number(ord['Delivery Charges'] || 0);
      sellingPriceSum += Number(ord['Final Selling Price'] || (ord as any)['Selling Price'] || 0);
      orderTotalSum += Number(ord['Final Order Total'] || (ord as any)['Order Total'] || 0);
      discountedBaseSum += Number(ord['Discounted Base Price'] || 0);
      ppdSum += Number(ord['PPD'] || 0);
      codSum += Number(ord['COD'] || 0);
      mealsSum += Number(ord['Meals'] || 0);
      deliveredOrdersCount += Number(ord['Orders Count'] || 1);
    });

    const vendorPrice = Number(vendorPriceSum.toFixed(2));
    const finalBasePrice = Number(basePriceSum.toFixed(2));
    const totalCommission = Number(totalCommSum.toFixed(2));
    const irctcComm = Number(irctcCommSum.toFixed(2));
    const rfComm = Number(rfCommSum.toFixed(2));
    const gst = Number(gstSum.toFixed(2));
    const totalDiscount = Number(totalDiscountSum.toFixed(2));
    const vendorDiscount = Number(vendorDiscountSum.toFixed(2));
    const rfDiscount = Number(rfDiscountSum.toFixed(2));
    const deliveryCharges = Number(deliveryChargesSum.toFixed(2));
    const sellingPrice = Number(sellingPriceSum.toFixed(2));
    const orderTotal = Number(orderTotalSum.toFixed(2));
    const discountedBase = Number(discountedBaseSum.toFixed(2));
    const ppd = Number(ppdSum.toFixed(2));
    const cod = Number(codSum.toFixed(2));
    const notDeliveredOrdersCount = bucket.notDeliveredCount;
    const totalOrdersCount = dateOrdersMap[dateKey]?.totalOrders ?? 0;
    const deliveredCount = deliveredOrdersCount;
    const cancelledCount = dateOrdersMap[dateKey]?.cancelledCount ?? 0;

    // Check %: (Total Commission / Vendor Price) %
    const checkPct = finalBasePrice > 0 ? `${((totalCommission / finalBasePrice) * 100).toFixed(2)}%` : '0.00%';

    // Not Delivered %
    const notDeliveredPct =
      deliveredOrdersCount > 0
        ? `${((notDeliveredOrdersCount / deliveredOrdersCount) * 100).toFixed(2)}%`
        : notDeliveredOrdersCount > 0
        ? '100.00%'
        : '#DIV/0!';

    // PPD % of Final Selling Price (same as Station Report)
    const ppdPct = sellingPrice > 0 ? `${((ppd / sellingPrice) * 100).toFixed(2)}%` : '0.00%';

    // Total Station Vendors: same Outlet Master source/rule as Station Report,
    // summed once per station for the stations represented on that date.
    const totalStationVendorsMap: Record<string, Set<string>> = {};
    Object.values(outletsMasterInfo || {}).forEach((out: any) => {
      const rawSt = out?.station || out?.stationCode || out?.stn_code || out?.deliveryStation || out?.stationName || '';
      const cleanSt = String(rawSt).trim().toUpperCase()
        .split('-')[0].split('(')[0].split('/')[0].trim()
        .replace(/[^A-Z0-9]/g, '');
      const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
      if (cleanSt && outId) {
        if (!totalStationVendorsMap[cleanSt]) totalStationVendorsMap[cleanSt] = new Set<string>();
        totalStationVendorsMap[cleanSt].add(outId);
      }
    });
    const totalStationVendors = Array.from(bucket.stationCodes).reduce(
      (sum, stationCode) => sum + (totalStationVendorsMap[stationCode]?.size || 0),
      0
    );

    reportRows.push({
      'Delivery Date': dateKey,
      'Total Orders': totalOrdersCount,
      'Delivered': deliveredCount,
      'Cancelled': cancelledCount,
      'Vendor Price': vendorPrice,
      'Final Base Price': finalBasePrice,
      'Final Total Commission': totalCommission,
      'Final IRCTC Comm': irctcComm,
      'Final RF Commission': rfComm,
      'Final GST': gst,
      'Final Discount': totalDiscount,
      'Final Vendor Discount': vendorDiscount,
      'Final RF Discount': rfDiscount,
      'Delivery Charges': deliveryCharges,
      'Final Selling Price': sellingPrice,
      'Final Order Total': orderTotal,
      'Discounted Base Price': discountedBase,
      'PPD': ppd,
      'COD': cod,
      'Meals': mealsSum,
      'Check': checkPct,
      'Count of Delivered Orders': deliveredOrdersCount,
      'Not Delivered Order': notDeliveredOrdersCount,
      'Not Delivered %': notDeliveredPct,
      'PPD % of Final Selling Price': ppdPct,
      'Feedback Good': bucket.feedbackGood,
      'Feedback Bad': bucket.feedbackBad,
      'Count of Delivered Outlets': bucket.deliveredOutlets.size,
      'Total Station Vendors': totalStationVendors,
    });
  });

  return reportRows;
};

/**
 * Direct Excel Workbook Generator for Date Wise Report
 */
export const generateDateWiseReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  irctcOrders: any[] = [],
  fileNamePrefix: string = 'DATE_WISE_SUMMARY_REPORT'
) => {
  const dateWiseData = generateDateWiseData(masterOrders, outletsMasterInfo, irctcOrders);

  if (!dateWiseData || dateWiseData.length === 0) {
    alert('Date wise report export karne ke liye koi data available nahi hai.');
    return;
  }

  const columns = [
    'Delivery Date','Total Orders','Delivered','Cancelled','Vendor Price','Final Base Price','Final Total Commission',
    'Final IRCTC Comm','Final RF Commission','Final GST','Final Discount','Final Vendor Discount','Final RF Discount',
    'Delivery Charges','Final Selling Price','Final Order Total','Discounted Base Price','PPD','COD','Meals','Check',
    'Count of Delivered Orders','Not Delivered Order','Not Delivered %','PPD % of Final Selling Price',
    'Feedback Good','Feedback Bad','Count of Delivered Outlets','Total Station Vendors'
  ];

  const totalRow: Record<string, any> = {};
  columns.forEach((column) => { totalRow[column] = ''; });
  totalRow['Delivery Date'] = 'TOTAL';

  const numericColumns = columns.filter((column) => ![
    'Delivery Date','Check','Not Delivered %','PPD % of Final Selling Price'
  ].includes(column));

  numericColumns.forEach((column) => {
    totalRow[column] = dateWiseData.reduce(
      (sum, row) => sum + (Number((row as any)[column]) || 0),
      0
    );
  });

  // Keep percentage/check columns blank in TOTAL, matching the Station Report pattern.
  const worksheet = XLSX.utils.aoa_to_sheet([
    columns.map((column) => totalRow[column]),
    columns,
    ...dateWiseData.map((row) => columns.map((column) => (row as any)[column]))
  ]);

  // Excel AutoFilter is on the second row (the actual header row), exactly as requested.
  worksheet['!autofilter'] = {
    ref: `A2:${XLSX.utils.encode_col(columns.length - 1)}${dateWiseData.length + 2}`
  };

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Date Wise Summary');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`);
};
