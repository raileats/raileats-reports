import * as XLSX from 'xlsx';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

// Normalize date to DD-MM-YYYY format for consistent column headers
const normalizeDateStr = (rawDate: string): string => {
  if (!rawDate) return '';
  const d = String(rawDate).trim();
  
  // If already DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = d.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (dmyMatch) {
    const day = dmyMatch[1].padStart(2, '0');
    const month = dmyMatch[2].padStart(2, '0');
    const year = dmyMatch[3];
    return `${day}-${month}-${year}`;
  }

  // If YYYY-MM-DD
  const ymdMatch = d.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (ymdMatch) {
    const year = ymdMatch[1];
    const month = ymdMatch[2].padStart(2, '0');
    const day = ymdMatch[3].padStart(2, '0');
    return `${day}-${month}-${year}`;
  }

  return d;
};

// Date parser for chronological sorting
const parseDateForSort = (dStr: string): number => {
  const parts = dStr.split('-');
  if (parts.length === 3) {
    return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0])).getTime();
  }
  return 0;
};

export const generateVendorDateWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {}
) => {
  const uniqueDatesSet = new Set<string>();
  const outletDataMap: Record<
    string,
    {
      outletId: string;
      vendorName: string;
      stnCode: string;
      dateCounts: Record<string, number>;
      totalDelivered: number;
    }
  > = {};

  // 1. Group & Aggregate
  masterOrders.forEach((ord) => {
    const outletId = String(ord['Outlet ID'] || '').trim().replace(/\.0$/, '');
    if (!outletId) return;

    if (!outletDataMap[outletId]) {
      const outMaster = outletsMasterInfo[outletId];
      outletDataMap[outletId] = {
        outletId,
        vendorName: outMaster?.outletName || ord['Vendor Name'] || `Outlet ${outletId}`,
        stnCode: outMaster?.station || ord['Station Code'] || '',
        dateCounts: {},
        totalDelivered: 0,
      };
    }

    const finalStatus = String(ord['Final Status'] || '').trim().toLowerCase();
    
    // Only count delivered orders
    if (finalStatus === 'delivered') {
      const rawDate = ord['Delivery Date'] || ord['Booking Date'] || '';
      const dateKey = normalizeDateStr(rawDate);

      if (dateKey) {
        uniqueDatesSet.add(dateKey);
        outletDataMap[outletId].dateCounts[dateKey] =
          (outletDataMap[outletId].dateCounts[dateKey] || 0) + (Number(ord['Orders Count']) || 1);
        outletDataMap[outletId].totalDelivered += Number(ord['Orders Count']) || 1;
      }
    }
  });

  // 2. Chronologically sort date columns
  const sortedDateColumns = Array.from(uniqueDatesSet).sort(
    (a, b) => parseDateForSort(a) - parseDateForSort(b)
  );

  // 3. Format rows matching image layout
  const rows = Object.values(outletDataMap).map((item) => {
    const rowObj: Record<string, any> = {
      'Row Labels': Number(item.outletId) || item.outletId,
      'Name': item.vendorName,
      'STN Code': item.stnCode,
    };

    sortedDateColumns.forEach((dt) => {
      // If 0, keep blank as shown in pivot screenshot
      rowObj[dt] = item.dateCounts[dt] ? item.dateCounts[dt] : '';
    });

    return rowObj;
  });

  // Sort by Outlet ID (Row Labels) ascending
  rows.sort((a, b) => Number(a['Row Labels']) - Number(b['Row Labels']));

  return { rows, dateColumns: sortedDateColumns };
};

/**
 * Direct Vendor Date Wise Report Excel Generator
 */
export const generateVendorDateWiseReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  fileNamePrefix: string = 'VENDOR_REPORT_DATE_WISE'
) => {
  const { rows } = generateVendorDateWiseData(masterOrders, outletsMasterInfo);
  if (rows.length === 0) {
    alert('Date-wise report export karne ke liye koi data available nahi hai.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor Date Wise');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`);
};
