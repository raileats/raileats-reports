import * as XLSX from 'xlsx';

export interface FeedbackReportRow {
  'Outlet Id': string;
  'Outlet Name': string;
  'Station Code': string;
  'Old Count': number;
  'Old Ratings': number;
  'Old Sum': number;
  Complaint: number;
  Feedback: number;
  'Current Count': number;
  'Current Rating': number;
  'Total Count': number;
  'Total Rating Sum': number;
  'Total Rating': number;
}

const clean = (value: any): string => {
  if (value === null || value === undefined) return '';
  return String(value).trim().replace(/\.0$/, '');
};

const getValue = (row: any, keys: string[]): any => {
  if (!row || typeof row !== 'object') return '';
  const rowKeys = Object.keys(row);
  for (const key of keys) {
    const normalized = key.toLowerCase().replace(/[^a-z0-9]/g, '');
    const found = rowKeys.find(
      (rk) => rk.toLowerCase().replace(/[^a-z0-9]/g, '') === normalized
    );
    if (
      found &&
      row[found] !== null &&
      row[found] !== undefined &&
      String(row[found]).trim() !== ''
    ) {
      return row[found];
    }
  }
  return '';
};

const toNumber = (value: any): number => {
  const n = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(n) ? n : 0;
};

const normalizeOrderId = (value: any): string => {
  if (value === null || value === undefined || value === '') return '';
  let s = String(value).trim().replace(/\u00A0/g, '').replace(/\s+/g, '');
  s = s.replace(/\.0+$/, '');
  if (/^[+-]?\d+(?:\.\d+)?e[+-]?\d+$/i.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) s = String(Math.trunc(n));
  }
  return s.toUpperCase();
};

const normalizeType = (value: any): 'FEEDBACK' | 'COMPLAINT' | '' => {
  const type = clean(value).toUpperCase();
  if (type === 'FEEDBACK') return 'FEEDBACK';
  if (type === 'COMPLAIN' || type === 'COMPLAINT') return 'COMPLAINT';
  return '';
};

/**
 * Creates the final Outlet-wise Feedback + Rating report.
 *
 * Current data:
 *   Feedback file -> Order ID + Type
 *   IRCTC file    -> Order Id -> Outlet Id / Outlet Name / Delivery Station
 *
 * Historical data:
 *   Old Feedback / Old Ratings Excel -> Outlet Id + Old Count + Old Ratings + Old Sum
 *
 * Current Rating:
 *   Feedback * 5 / (Feedback + Complaint)
 *
 * Total Rating:
 *   (Old Sum + Feedback * 5) /
 *   (Old Count + Feedback + Complaint)
 *
 * Complaint adds to the count but contributes zero rating points.
 * Final outlet rows are the UNION of old outlets and current outlets.
 */
export const generateFeedbackReportData = (
  feedbackData: any[] = [],
  irctcData: any[] = [],
  oldRatingsData: any[] = []
): FeedbackReportRow[] => {
  // 1. Current Feedback/Complaint count by Order ID.
  const orderTypeCounts: Record<string, { complaint: number; feedback: number }> = {};

  for (const row of feedbackData) {
    const orderId = normalizeOrderId(
      getValue(row, ['Order ID', 'Order Id', 'OrderID', 'IRCTC Order ID'])
    );
    const type = normalizeType(getValue(row, ['Type', 'Feedback Type', 'FeedbackType']));
    if (!orderId || !type) continue;

    if (!orderTypeCounts[orderId]) {
      orderTypeCounts[orderId] = { complaint: 0, feedback: 0 };
    }

    if (type === 'COMPLAINT') orderTypeCounts[orderId].complaint += 1;
    if (type === 'FEEDBACK') orderTypeCounts[orderId].feedback += 1;
  }

  // 2. IRCTC Order ID -> outlet information.
  const orderToOutlet: Record<string, any> = {};
  const outletToIrctc: Record<string, any> = {};

  for (const row of irctcData) {
    const orderId = normalizeOrderId(
      getValue(row, ['Order Id', 'Order ID', 'OrderID'])
    );
    const outletId = clean(
      getValue(row, ['Outlet Id', 'Outlet ID', 'OutletId'])
    );

    if (orderId && !orderToOutlet[orderId]) orderToOutlet[orderId] = row;
    if (outletId && !outletToIrctc[outletId]) outletToIrctc[outletId] = row;
  }

  // 3. Aggregate current data by Outlet ID.
  const outletMap: Record<string, FeedbackReportRow> = {};

  for (const [orderId, counts] of Object.entries(orderTypeCounts)) {
    const irctc = orderToOutlet[orderId] || {};
    const feedbackRow = feedbackData.find((row) =>
      normalizeOrderId(
        getValue(row, ['Order ID', 'Order Id', 'OrderID', 'IRCTC Order ID'])
      ) === orderId
    );

    const outletId = clean(
      getValue(irctc, ['Outlet Id', 'Outlet ID', 'OutletId']) ||
      getValue(feedbackRow, ['Outlet ID', 'Outlet Id', 'OutletId'])
    );
    if (!outletId) continue;

    const fallback = outletToIrctc[outletId] || {};
    const outletName = clean(
      getValue(irctc, ['Outlet Name', 'OutletName']) ||
      getValue(feedbackRow, ['Outlet', 'Outlet Name']) ||
      getValue(fallback, ['Outlet Name', 'OutletName'])
    );
    const stationCode = clean(
      getValue(irctc, ['Delivery Station', 'Station Code', 'StationCode']) ||
      getValue(fallback, ['Delivery Station', 'Station Code'])
    ).toUpperCase();

    if (!outletMap[outletId]) {
      outletMap[outletId] = {
        'Outlet Id': outletId,
        'Outlet Name': outletName,
        'Station Code': stationCode,
        'Old Count': 0,
        'Old Ratings': 0,
        'Old Sum': 0,
        Complaint: 0,
        Feedback: 0,
        'Current Count': 0,
        'Current Rating': 0,
        'Total Count': 0,
        'Total Rating Sum': 0,
        'Total Rating': 0,
      };
    }

    if (!outletMap[outletId]['Outlet Name']) outletMap[outletId]['Outlet Name'] = outletName;
    if (!outletMap[outletId]['Station Code']) outletMap[outletId]['Station Code'] = stationCode;

    outletMap[outletId].Complaint += counts.complaint;
    outletMap[outletId].Feedback += counts.feedback;
  }

  // 4. Old Ratings / Old Feedback data by Outlet ID.
  for (const row of oldRatingsData) {
    const outletId = clean(
      getValue(row, ['Outlet Id', 'Outlet ID', 'OutletId', 'Aggregator Outlet ID'])
    );
    if (!outletId) continue;

    const oldCount = toNumber(getValue(row, ['Old Count', 'OldCount', 'Count']));
    const oldRatings = toNumber(getValue(row, ['Old Ratings', 'Old Rating', 'Rating']));
    const rawOldSum = getValue(row, ['Old Sum', 'OldSum']);
    const oldSum = rawOldSum !== ''
      ? toNumber(rawOldSum)
      : Number((oldCount * oldRatings).toFixed(2));

    if (!outletMap[outletId]) {
      outletMap[outletId] = {
        'Outlet Id': outletId,
        'Outlet Name': clean(getValue(row, ['Outlet Name', 'Outlet'])),
        'Station Code': clean(getValue(row, ['Station Code', 'Station', 'Delivery Station'])).toUpperCase(),
        'Old Count': oldCount,
        'Old Ratings': oldRatings,
        'Old Sum': oldSum,
        Complaint: 0,
        Feedback: 0,
        'Current Count': 0,
        'Current Rating': 0,
        'Total Count': 0,
        'Total Rating Sum': 0,
        'Total Rating': 0,
      };
    } else {
      if (!outletMap[outletId]['Outlet Name']) {
        outletMap[outletId]['Outlet Name'] = clean(getValue(row, ['Outlet Name', 'Outlet']));
      }
      if (!outletMap[outletId]['Station Code']) {
        outletMap[outletId]['Station Code'] =
          clean(getValue(row, ['Station Code', 'Station', 'Delivery Station'])).toUpperCase();
      }
      outletMap[outletId]['Old Count'] = oldCount;
      outletMap[outletId]['Old Ratings'] = oldRatings;
      outletMap[outletId]['Old Sum'] = oldSum;
    }
  }

  // 5. Calculate current and total ratings.
  return Object.values(outletMap)
    .map((row) => {
      const currentCount = row.Feedback + row.Complaint;
      const currentRating = currentCount > 0
        ? Number(((row.Feedback * 5) / currentCount).toFixed(2))
        : 0;

      const totalCount = row['Old Count'] + currentCount;
      const totalRatingSum = Number(
        (row['Old Sum'] + row.Feedback * 5).toFixed(2)
      );
      const totalRating = totalCount > 0
        ? Number((totalRatingSum / totalCount).toFixed(2))
        : 0;

      return {
        ...row,
        'Current Count': currentCount,
        'Current Rating': currentRating,
        'Total Count': totalCount,
        'Total Rating Sum': totalRatingSum,
        'Total Rating': totalRating,
      };
    })
    .sort((a, b) =>
      a['Outlet Id'].localeCompare(b['Outlet Id'], undefined, { numeric: true })
    );
};

export const generateFeedbackReportWorkbook = (
  feedbackData: any[] = [],
  irctcData: any[] = [],
  oldRatingsData: any[] = [],
  fileNamePrefix = 'FEEDBACK_REPORT'
) => {
  const rows = generateFeedbackReportData(feedbackData, irctcData, oldRatingsData);

  if (!rows.length) {
    alert('Feedback Report ke liye current ya Old Feedback data available nahi hai.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 16 }, { wch: 42 }, { wch: 16 },
    { wch: 12 }, { wch: 13 }, { wch: 13 },
    { wch: 12 }, { wch: 12 }, { wch: 14 },
    { wch: 14 }, { wch: 13 }, { wch: 16 }, { wch: 14 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Feedback Report');

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${today}.xlsx`);
};
