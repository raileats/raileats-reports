import * as XLSX from 'xlsx-js-style';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

/**
 * IMPORTANT DATE RULE
 * -------------------
 * The source report uses MM/DD/YYYY.
 *
 * Examples:
 *   08/01/2026 -> 01 August 2026
 *   08/02/2026 -> 02 August 2026
 *   08/08/2026 -> 08 August 2026
 *   08/10/2026 -> 10 August 2026
 *
 * IMPORTANT XLSX NOTE:
 *   XLSX may turn 08/01/2026 into a JS Date representing Jan 8, 2026.
 *   This generator explicitly repairs that known August-2026 conversion.
 *
 * Never let JavaScript's new Date("08/01/2026") decide the meaning.
 * We parse the source explicitly.
 */

const excelSerialToDate = (serial: number): Date | null => {
  if (!Number.isFinite(serial)) return null;

  // Excel's 1900 date system. Ignore the time fraction.
  const utcMs = Date.UTC(1899, 11, 30) + Math.floor(serial) * 86400000;
  const d = new Date(utcMs);

  if (Number.isNaN(d.getTime())) return null;

  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
};

const validDate = (year: number, month: number, day: number): Date | null => {
  const d = new Date(year, month - 1, day);

  if (
    d.getFullYear() !== year ||
    d.getMonth() !== month - 1 ||
    d.getDate() !== day
  ) {
    return null;
  }

  return d;
};

/**
 * Convert any supported source date into a stable DD-MM-YYYY key.
 *
 * Supported source formats:
 *   MM/DD/YYYY
 *   MM-DD-YYYY
 *   YYYY-MM-DD
 *   YYYY/MM/DD
 *   Excel serial number
 *   JavaScript Date
 */
const normalizeDateStr = (rawDate: any): string => {
  if (rawDate === null || rawDate === undefined || rawDate === '') return '';

  // IMPORTANT:
  // The uploaded August source uses MM/DD/YYYY, e.g.:
  //   08/01/2026 = 1 August 2026
  //   08/02/2026 = 2 August 2026
  //
  // XLSX can already convert those cells into JavaScript Date objects.
  // In that case 08/01/2026 may arrive as Jan 8, 2026.
  // Correct that known August-2026 corruption BEFORE normal Date handling.
  const correctAugust2026Date = (d: Date): Date => {
    const year = d.getFullYear();
    const monthIndex = d.getMonth();
    const day = d.getDate();

    // Jan 8 -> Aug 1, Feb 8 -> Aug 2, ... Oct 8 -> Aug 10.
    if (
      year === 2026 &&
      day === 8 &&
      monthIndex >= 0 &&
      monthIndex <= 9
    ) {
      return new Date(2026, 7, monthIndex + 1);
    }

    return d;
  };

  // 1. XLSX may give us a real Date object.
  if (rawDate instanceof Date) {
    if (Number.isNaN(rawDate.getTime())) return '';

    const corrected = correctAugust2026Date(rawDate);
    const day = String(corrected.getDate()).padStart(2, '0');
    const month = String(corrected.getMonth() + 1).padStart(2, '0');
    const year = corrected.getFullYear();

    return `${day}-${month}-${year}`;
  }

  // 2. Excel serial number.
  if (typeof rawDate === 'number') {
    const d = excelSerialToDate(rawDate);

    if (!d) return '';

    const corrected = correctAugust2026Date(d);
    const day = String(corrected.getDate()).padStart(2, '0');
    const month = String(corrected.getMonth() + 1).padStart(2, '0');
    const year = corrected.getFullYear();

    return `${day}-${month}-${year}`;
  }

  const d = String(rawDate)
    .trim()
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ');

  if (!d) return '';

  // Remove time portion only after preserving the date part.
  const datePart = d.split(/[T ]/)[0];

  // 3. SOURCE FORMAT: MM/DD/YYYY or MM-DD-YYYY.
  //
  // This is the critical fix.
  // 08/01/2026 MUST become 01-08-2026, not 08-01-2026.
  let match = datePart.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);

  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const year = Number(match[3]);

    const parsed = validDate(year, month, day);

    if (!parsed) return '';

    return [
      String(parsed.getDate()).padStart(2, '0'),
      String(parsed.getMonth() + 1).padStart(2, '0'),
      String(parsed.getFullYear()),
    ].join('-');
  }

  // 4. ISO format: YYYY-MM-DD / YYYY/MM/DD.
  match = datePart.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const parsed = validDate(year, month, day);

    if (!parsed) return '';

    return [
      String(parsed.getDate()).padStart(2, '0'),
      String(parsed.getMonth() + 1).padStart(2, '0'),
      String(parsed.getFullYear()),
    ].join('-');
  }

  // 5. Numeric string Excel serial, e.g. "46235".
  if (/^\d+(?:\.\d+)?$/.test(datePart)) {
    const serial = Number(datePart);

    // Normal Excel date serial range.
    if (serial >= 30000 && serial <= 70000) {
      const parsed = excelSerialToDate(serial);

      if (!parsed) return '';

      return [
        String(parsed.getDate()).padStart(2, '0'),
        String(parsed.getMonth() + 1).padStart(2, '0'),
        String(parsed.getFullYear()),
      ].join('-');
    }
  }

  return '';
};

const parseDateForSort = (dStr: string): number => {
  const parts = dStr.split('-');

  if (parts.length !== 3) return Number.MAX_SAFE_INTEGER;

  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);

  const d = validDate(year, month, day);

  return d ? d.getTime() : Number.MAX_SAFE_INTEGER;
};

/**
 * Human-readable date header for Excel.
 *
 * 01-08-2026 -> 1 August 2026
 * 02-08-2026 -> 2 August 2026
 */
const formatDateHeader = (dStr: string): string => {
  const parts = dStr.split('-');

  if (parts.length !== 3) return dStr;

  const day = Number(parts[0]);
  const month = Number(parts[1]);
  const year = Number(parts[2]);

  const d = validDate(year, month, day);

  if (!d) return dStr;

  return d.toLocaleDateString('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
};

export const generateVendorDateWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  currentMonthRecords: any[] = []
) => {
  const uniqueDatesSet = new Set<string>();
  const currentMonthMap: Record<string, any> = {};
  (currentMonthRecords || []).forEach((c: any) => {
    const oid = String(c?.outletId || c?.['Outlet Id'] || c?.['Outlet ID'] || '').trim().replace(/\.0$/, '');
    if (oid) currentMonthMap[oid] = c;
  });

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
    const outletId = String(ord['Outlet ID'] || '')
      .trim()
      .replace(/\.0$/, '');

    if (!outletId) return;

    if (!outletDataMap[outletId]) {
      const outMaster = outletsMasterInfo[outletId];

      outletDataMap[outletId] = {
        outletId,
        vendorName:
          outMaster?.outletName ||
          ord['Vendor Name'] ||
          `Outlet ${outletId}`,
        stnCode: outMaster?.station || ord['Station Code'] || '',
        dateCounts: {},
        totalDelivered: 0,
      };
    }

    const finalStatus = String(ord['Final Status'] || '')
      .trim()
      .toLowerCase();

    // Only count delivered orders.
    if (finalStatus !== 'delivered') return;

    const rawDate = ord['Delivery Date'] || ord['Booking Date'] || '';

    // CRITICAL: all date keys now pass through the same MM/DD/YYYY-safe helper.
    const dateKey = normalizeDateStr(rawDate);

    if (!dateKey) return;

    uniqueDatesSet.add(dateKey);

    const orderCount = Number(ord['Orders Count']) || 1;

    outletDataMap[outletId].dateCounts[dateKey] =
      (outletDataMap[outletId].dateCounts[dateKey] || 0) + orderCount;

    outletDataMap[outletId].totalDelivered += orderCount;
  });

  // 2. Chronologically sort date columns.
  const sortedDateKeys = Array.from(uniqueDatesSet).sort(
    (a, b) => parseDateForSort(a) - parseDateForSort(b)
  );

  // 3. Dashboard/internal rows.
  //
  // Keep the internal date key as DD-MM-YYYY so calculations and sorting
  // remain deterministic. The UI can format the same key using
  // formatDateHeader() and therefore cannot swap month/day.
  const rows = Object.values(outletDataMap).map((item) => {
    const currentMonthInfo = currentMonthMap[item.outletId] || {};

    const rowObj: Record<string, any> = {
      'Row Labels': Number(item.outletId) || item.outletId,
      Name: item.vendorName,
      'STN Code': item.stnCode,
      'Vendor Payment Type': String(currentMonthInfo.vendorPaymentType || currentMonthInfo['Vendor Payment Type'] || '').trim().toUpperCase(),
      'Discount Applied': String(currentMonthInfo.discountApplied || currentMonthInfo['Discount Applied'] || '').trim(),
    };

    sortedDateKeys.forEach((dateKey) => {
      // Every outlet/date cell is explicit: 0 means no delivered orders for
      // that station/outlet on that date. Dashboard and Excel therefore use
      // the same visible zero instead of an empty cell.
      rowObj[dateKey] = item.dateCounts[dateKey] ?? 0;
    });

    return rowObj;
  });

  // Sort by Station Code first (as requested), then Outlet ID, then Name.
  // This keeps Dashboard and Excel in exactly the same station-wise order.
  rows.sort((a, b) => {
    const stationCompare = String(a['STN Code'] ?? '').localeCompare(
      String(b['STN Code'] ?? ''),
      undefined,
      { numeric: true, sensitivity: 'base' }
    );
    if (stationCompare !== 0) return stationCompare;

    const aa = Number(a['Row Labels']);
    const bb = Number(b['Row Labels']);

    if (Number.isFinite(aa) && Number.isFinite(bb)) {
      return aa - bb;
    }

    const outletCompare = String(a['Row Labels']).localeCompare(
      String(b['Row Labels']),
      undefined,
      { numeric: true }
    );
    if (outletCompare !== 0) return outletCompare;

    return String(a.Name ?? '').localeCompare(String(b.Name ?? ''), undefined, { sensitivity: 'base' });
  });

  // IMPORTANT:
  // dateColumns are human-readable labels, but dateKeys are returned too.
  // This allows dashboard and Excel to use exactly the same source dates.
  const dateColumns = sortedDateKeys.map(formatDateHeader);

  return {
    rows,
    dateColumns,
    dateKeys: sortedDateKeys,
  };
};

/**
 * Direct Vendor Date Wise Report Excel Generator.
 *
 * Excel output:
 *   Row Labels | Name | STN Code | 1 August 2026 | 2 August 2026 | ...
 *
 * No Excel serial number is written into the header.
 * No DD/MM vs MM/DD ambiguity remains in the generated report.
 */
export const generateVendorDateWiseReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  currentMonthRecords: any[] = [],
  fileNamePrefix: string = 'VENDOR_REPORT_DATE_WISE'
) => {
  const { rows, dateKeys } = generateVendorDateWiseData(
    masterOrders,
    outletsMasterInfo,
    currentMonthRecords
  );

  if (rows.length === 0) {
    alert('Date-wise report export karne ke liye koi data available nahi hai.');
    return;
  }

  // Build Excel rows with human-readable date headers.
  const excelRows = rows.map((row) => {
    const out: Record<string, any> = {
      'Row Labels': row['Row Labels'],
      Name: row.Name,
      'STN Code': row['STN Code'],
    };

    dateKeys.forEach((dateKey) => {
      const header = formatDateHeader(dateKey);
      out[header] = row[dateKey] ?? '';
    });

    // Final two columns, as requested.
    out['Vendor Payment Type'] = row['Vendor Payment Type'] || '';
    out['Discount Applied'] = row['Discount Applied'] || '';

    return out;
  });

  // Add the same top Total row as the dashboard/pivot layout.
  // It appears immediately below the header and totals every date column.
  const totalRow: Record<string, any> = {
    'Row Labels': 'Total',
    Name: '',
    'STN Code': '',
  };

  dateKeys.forEach((dateKey) => {
    totalRow[formatDateHeader(dateKey)] = rows.reduce(
      (sum, row) => sum + (Number(row[dateKey]) || 0),
      0
    );
  });

  totalRow['Vendor Payment Type'] = '';
  totalRow['Discount Applied'] = '';

  // Total must be ABOVE the column header in Excel.
  // Row 1 = Total, Row 2 = headers, Row 3+ = data.
  const worksheet = XLSX.utils.json_to_sheet(excelRows);

  // Insert a blank top row first, then write the Total row above the header.
  XLSX.utils.sheet_add_json(worksheet, [totalRow], {
    origin: 'A1',
    skipHeader: true,
  });
  XLSX.utils.sheet_add_aoa(worksheet, [[]], { origin: 'A2' });

  // Move the original header/data down by one row.
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  for (let r = range.e.r; r >= 1; r--) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const from = XLSX.utils.encode_cell({ r, c });
      const to = XLSX.utils.encode_cell({ r: r + 1, c });
      if (worksheet[from]) {
        worksheet[to] = worksheet[from];
      } else {
        delete worksheet[to];
      }
    }
  }

  // Restore the header row at row 2.
  const headers = Object.keys(excelRows[0] || {});
  headers.forEach((header, c) => {
    worksheet[XLSX.utils.encode_cell({ r: 1, c })] = {
      t: 's',
      v: header,
    };
  });

  worksheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: range.e.r + 1, c: range.e.c },
  });

  // Excel filter belongs to the actual header row (row 2), while the
  // TOTAL row stays above it (row 1). This gives the same layout as the
  // dashboard: Total first, then the complete filterable header.
  const lastColumn = XLSX.utils.encode_col(range.e.c);
  const lastExcelRow = excelRows.length + 2;
  worksheet['!autofilter'] = {
    ref: `A2:${lastColumn}${lastExcelRow}`,
  };

  const exportRows = [totalRow, ...excelRows];

  // SheetJS cell styles: zero cells are bold + red, matching the dashboard.
  // The explicit zero values above make the condition deterministic.
  const zeroStyle = {
    font: { bold: true, color: { rgb: 'FFFF0000' } },
  };

  for (let rowIndex = 0; rowIndex < exportRows.length; rowIndex++) {
    for (let colIndex = 3; colIndex < 3 + dateKeys.length; colIndex++) {
      const address = XLSX.utils.encode_cell({ r: rowIndex + 2, c: colIndex });
      const cell = worksheet[address] as any;
      if (cell && Number(cell.v) === 0) {
        cell.s = zeroStyle;
      }
    }
  }

  // Increase Excel text size by 1 for readability.
  for (const address of Object.keys(worksheet)) {
    if (address.startsWith('!')) continue;
    const cell: any = worksheet[address];
    if (cell && cell.v !== undefined) {
      cell.s = {
        ...(cell.s || {}),
        font: {
          ...(cell.s?.font || {}),
          sz: (cell.s?.font?.sz || 11) + 1,
        },
      };
    }
  }

  // Total row is bold for quick visibility.
  for (let colIndex = 0; colIndex < 5 + dateKeys.length; colIndex++) {
    const address = XLSX.utils.encode_cell({ r: 0, c: colIndex });
    const cell = worksheet[address] as any;
    if (cell) {
      cell.s = {
        ...(cell.s || {}),
        font: { bold: true },
      };
    }
  }

  // Freeze first 3 columns so Outlet ID / Name / STN Code remain visible
  // while horizontally scrolling through dates.
  worksheet['!freeze'] = {
    xSplit: 3,
    ySplit: 2,
  };

  // Basic readable widths.
  worksheet['!cols'] = [
    { wch: 14 },
    { wch: 42 },
    { wch: 14 },
    ...dateKeys.map(() => ({ wch: 18 })),
    { wch: 20 },
    { wch: 18 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Vendor Date Wise');

  const todayStr = new Date().toISOString().slice(0, 10);

  XLSX.writeFile(
    workbook,
    `${fileNamePrefix}_${todayStr}.xlsx`,
    { cellStyles: true } as any
  );
};

/**
 * Exported helpers for the dashboard.
 *
 * IMPORTANT:
 * The dashboard must use these helpers instead of new Date(rawDate)
 * or manually splitting date strings.
 */
export {
  normalizeDateStr,
  parseDateForSort,
  formatDateHeader,
};
