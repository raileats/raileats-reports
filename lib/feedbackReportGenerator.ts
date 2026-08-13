import * as XLSX from 'xlsx';

export interface FeedbackReportRow {
  'Outlet ID': string;
  'Outlet Name': string;
  'Station Code': string;
  'Complaint': number;
  'Feedback': number;
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
    if (found && row[found] !== null && row[found] !== undefined && String(row[found]).trim() !== '') {
      return row[found];
    }
  }
  return '';
};

const normalizeType = (value: any): 'FEEDBACK' | 'COMPLAINT' | '' => {
  const type = clean(value).toUpperCase();
  if (type === 'FEEDBACK') return 'FEEDBACK';
  if (type === 'COMPLAIN' || type === 'COMPLAINT') return 'COMPLAINT';
  return '';
};

/**
 * Creates Outlet-wise Feedback Report.
 *
 * Source 1: Feedback report
 *   Order ID + Type
 *   FEEDBACK / feedback -> Feedback count
 *   COMPLAIN / COMPLAINT / complaint -> Complaint count
 *
 * Source 2: IRCTC report
 *   Order Id -> Outlet Id, Outlet Name, Delivery Station
 *
 * Join key: Order ID (Feedback) = Order Id (IRCTC)
 * Then aggregate the matched orders by Outlet ID + Outlet Name + Station Code.
 */
export const generateFeedbackReportData = (
  feedbackData: any[] = [],
  irctcData: any[] = []
): FeedbackReportRow[] => {
  // Step 1: Count Feedback/Complaint by Order ID.
  const orderTypeCounts: Record<string, { complaint: number; feedback: number }> = {};

  for (const row of feedbackData || []) {
    const orderId = clean(getValue(row, ['Order ID', 'Order Id', 'OrderID']));
    if (!orderId) continue;

    const type = normalizeType(getValue(row, ['Type']));
    if (!type) continue;

    if (!orderTypeCounts[orderId]) {
      orderTypeCounts[orderId] = { complaint: 0, feedback: 0 };
    }

    if (type === 'COMPLAINT') orderTypeCounts[orderId].complaint += 1;
    if (type === 'FEEDBACK') orderTypeCounts[orderId].feedback += 1;
  }

  // Step 2: Build Order ID -> IRCTC outlet/station mapping.
  // Duplicate IRCTC rows for the same Order Id are ignored after the first
  // usable mapping so one order does not get counted multiple times.
  const orderToOutlet: Record<string, {
    outletId: string;
    outletName: string;
    stationCode: string;
  }> = {};

  for (const row of irctcData || []) {
    const orderId = clean(getValue(row, ['Order Id', 'Order ID', 'OrderID']));
    if (!orderId || orderToOutlet[orderId]) continue;

    const outletId = clean(getValue(row, ['Outlet Id', 'Outlet ID', 'OutletId']));
    const outletName = clean(getValue(row, ['Outlet Name', 'OutletName']));
    const stationCode = clean(getValue(row, ['Delivery Station', 'DeliveryStation', 'Station Code', 'StationCode'])).toUpperCase();

    if (!outletId) continue;

    orderToOutlet[orderId] = {
      outletId,
      outletName,
      stationCode,
    };
  }

  // Step 3: Join and aggregate outlet-wise.
  const outletMap: Record<string, FeedbackReportRow> = {};

  Object.entries(orderTypeCounts).forEach(([orderId, counts]) => {
    const irctc = orderToOutlet[orderId];
    if (!irctc) return;

    const key = `${irctc.outletId}|||${irctc.stationCode}`;

    if (!outletMap[key]) {
      outletMap[key] = {
        'Outlet ID': irctc.outletId,
        'Outlet Name': irctc.outletName,
        'Station Code': irctc.stationCode,
        'Complaint': 0,
        'Feedback': 0,
      };
    }

    outletMap[key]['Complaint'] += counts.complaint;
    outletMap[key]['Feedback'] += counts.feedback;
  });

  return Object.values(outletMap).sort((a, b) => {
    const stationCompare = a['Station Code'].localeCompare(b['Station Code']);
    if (stationCompare !== 0) return stationCompare;
    const nameCompare = a['Outlet Name'].localeCompare(b['Outlet Name']);
    if (nameCompare !== 0) return nameCompare;
    return a['Outlet ID'].localeCompare(b['Outlet ID']);
  });
};

export const generateFeedbackReportWorkbook = (
  feedbackData: any[] = [],
  irctcData: any[] = [],
  fileNamePrefix = 'FEEDBACK_REPORT'
) => {
  const rows = generateFeedbackReportData(feedbackData, irctcData);

  if (!rows.length) {
    alert('Feedback Report ke liye matching Order ID data nahi mila. Feedback aur IRCTC reports check karein.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  worksheet['!cols'] = [
    { wch: 16 },
    { wch: 42 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Feedback Report');

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${today}.xlsx`);
};
