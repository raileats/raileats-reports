'use client';

// UI BUILD: DAY/NIGHT THEME + ALL-REPORT HORIZONTAL SCROLL + 3 FROZEN COLUMNS + CLICK ROW HIGHLIGHT + EXCEL-ALIGNED DASHBOARD

import React, { useState, useEffect, useMemo, useRef } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { generateMainReportWorkbook } from '@/lib/mainReportGenerator';
import { generateVendorRDSWorkbook, generateVendorRdsData } from '@/lib/vendorRdsGenerator';
import { generateStationReportWorkbook, generateStationWiseData } from '@/lib/stationReportGenerator';
import { generateVendorReportWorkbook, generateVendorWiseData } from '@/lib/vendorReportGenerator';
import { generateDateWiseReportWorkbook } from '@/lib/dateWiseReportGenerator';
import { generateVendorDateWiseReportWorkbook, generateVendorDateWiseData } from '@/lib/vendorDateWiseReportGenerator';
import { generateLastDayStationReportWorkbook, generateLastDayStationWiseData } from '@/lib/lastDayStationReportGenerator';
import MainReportMatrix from '@/components/MainReportMatrix';

// --- Native IndexedDB Storage Engine ---
const DB_NAME = 'RelFoodMasterDB';
const STORE_NAME = 'master_reports';

const openDB = (): Promise<IDBDatabase> => {
  return new Promise((resolve, reject) => {
    if (typeof window === 'undefined') return;
    const req = indexedDB.open(DB_NAME, 3);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
};

const saveToDB = async (key: string, data: any): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    store.put(data, key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
};

const loadFromDB = async (key: string): Promise<any> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const store = tx.objectStore(STORE_NAME);
    const req = store.get(key);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
};

const clearDB = async (): Promise<void> => {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    const store = tx.objectStore(STORE_NAME);
    store.clear();
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
};

// Master Final Status Calculation Rule
const computeFinalStatus = (rfStatusRaw: string, irctcStatusRaw: string): string => {
  const rf = (rfStatusRaw || '').trim().toUpperCase();
  const irctc = (irctcStatusRaw || '').trim().toUpperCase();

  if (rf.includes('UNDELIVERED') || irctc.includes('UNDELIVERED')) return 'Not Delivered';
  if (rf.includes('CANCEL') || irctc.includes('CANCEL')) return 'Cancelled';
  if (
    rf.includes('DELIVER') ||
    rf.includes('CONFIRM') ||
    rf.includes('PLACED') ||
    irctc.includes('DELIVER') ||
    irctc.includes('CONFIRM') ||
    irctc.includes('PLACED') ||
    irctc.includes('PENDING')
  ) {
    return 'Delivered';
  }
  return 'Delivered';
};

const cleanOutletId = (val: any): string => {
  if (val === null || val === undefined || String(val).trim() === '') return '';
  return String(val)
    .trim()
    .replace(/\u00A0/g, '')
    .replace(/\s+/g, '')
    .replace(/\.0+$/, '');
};

const getOutletIdFromRow = (row: any): string => {
  if (!row) return '';
  const candidates = [
    row['Outlet ID'],
    row['Outlet Id'],
    row['OutletId'],
    row['OutletID'],
    row['Aggregator Outlet ID'],
    row['Aggregator Outlet Id'],
    row['Aggregator OutletId'],
    row['Aggregator OutletID'],
    row['Outlet Code'],
    row['Outlet'],
  ];
  for (const value of candidates) {
    const id = cleanOutletId(value);
    if (id) return id;
  }
  return '';
};

const getOutletMasterStationName = (row: any): string => {
  if (!row) return '';
  // Outlet Master keeps the full Station Name under column O.
  const named = String(row['Station Name'] ?? '').trim();
  if (named) return named;
  const keys = Object.keys(row);
  const columnOKey = keys[14];
  return columnOKey ? String(row[columnOKey] ?? '').trim() : '';
};

// ---------------------------------------------------------------------------
// Feedback Report Engine
// Feedback source: Order ID + Type (Feedback / Complaint)
// IRCTC source: Order Id -> Outlet Id / Outlet Name / Delivery Station
// Counts are accumulated order-wise first, then aggregated Outlet-wise.
// ---------------------------------------------------------------------------
const cleanOrderId = (val: any): string => {
  if (val === null || val === undefined || val === '') return '';

  // Excel/CSV can represent the same Order ID in slightly different forms
  // (spaces, .0, scientific notation, hidden characters). Normalize all of
  // them before joining Feedback <-> IRCTC.
  let s = String(val).trim();
  s = s.replace(/\u00A0/g, ' ');
  s = s.replace(/\.0+$/, '');
  s = s.replace(/\s+/g, '');

  // Convert a numeric/scientific representation safely when possible.
  if (/^[+-]?\d+(?:\.\d+)?e[+-]?\d+$/i.test(s)) {
    const n = Number(s);
    if (Number.isFinite(n)) s = String(Math.trunc(n));
  }

  return s.toUpperCase();
};

const cleanTextKey = (val: any): string => {
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/\u00A0/g, ' ')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
};

const normalizeFeedbackType = (val: any): 'Feedback' | 'Complaint' | '' => {
  const type = String(val ?? '').trim().toUpperCase();
  if (type === 'FEEDBACK') return 'Feedback';
  if (type === 'COMPLAINT' || type === 'COMPLAIN') return 'Complaint';
  return '';
};

interface FeedbackOrderCount {
  complaint: number;
  feedback: number;
}

interface FeedbackReportRow {
  'Outlet Id': string;
  'Outlet Name': string;
  'Station Code': string;
  Complaint: number;
  Feedback: number;
  'Current Count': number;
  'Current Rating': number;
  'Current Sum': number;
  'Old Count': number;
  'Old Ratings': number;
  'Old Sum': number;
  'Total Count': number;
  'Total Rating Sum': number;
  'Total Rating': number;
}

const buildFeedbackReport = (
  feedbackRows: any[] = [],
  irctcRows: any[] = [],
  masterRows: any[] = [],
  outletsMasterInfo: Record<string, any> = {},
  oldRatingsRows: any[] = []
): FeedbackReportRow[] => {
  // STEP 1: Count every Feedback-file row by Order ID + Type.
  // Never overwrite duplicate Order IDs: every Feedback/Complaint row counts.
  const orderCounts: Record<string, FeedbackOrderCount> = {};

  feedbackRows.forEach((row) => {
    const orderId = cleanOrderId(
      row['Order ID'] ?? row['Order Id'] ?? row['OrderID'] ?? row['IRCTC Order ID']
    );
    const type = normalizeFeedbackType(
      row['Type'] ?? row['Feedback Type'] ?? row['FeedbackType']
    );
    if (!orderId || !type) return;

    if (!orderCounts[orderId]) {
      orderCounts[orderId] = { complaint: 0, feedback: 0 };
    }

    if (type === 'Complaint') orderCounts[orderId].complaint += 1;
    if (type === 'Feedback') orderCounts[orderId].feedback += 1;
  });

  // STEP 2: Build IRCTC lookup maps.
  // Order ID is the ONLY join key for current Feedback rows.
  // Final Outlet ID always comes from IRCTC.
  const irctcOrderMap = new Map<string, any>();
  const irctcOutletMap = new Map<string, any>();

  irctcRows.forEach((row) => {
    const orderId = cleanOrderId(
      row['Order Id'] ?? row['Order ID'] ?? row['OrderID']
    );
    const outletId = cleanOutletId(
      row['Outlet Id'] ?? row['Outlet ID'] ?? row['OutletId']
    );

    if (orderId && !irctcOrderMap.has(orderId)) irctcOrderMap.set(orderId, row);
    if (outletId && !irctcOutletMap.has(outletId)) irctcOutletMap.set(outletId, row);
  });

  // STEP 3: Outlet Master fallback for descriptive fields.
  const masterOutletMap = new Map<string, any>();
  Object.values(outletsMasterInfo || {}).forEach((row: any) => {
    const outletId = cleanOutletId(row?.outletId ?? row?.['Outlet Id'] ?? row?.['Outlet ID']);
    if (outletId) masterOutletMap.set(outletId, row);
  });

  // STEP 4: Build a direct Feedback Order ID -> first Feedback row map.
  //
  // IMPORTANT PERFORMANCE FIX:
  // Never use feedbackRows.find() inside the orderCounts loop.
  // With ~10,000 Feedback rows that becomes O(N²) and can perform
  // 100+ million comparisons. A Map makes this O(1) per order.
  const feedbackRowMap = new Map<string, any>();
  feedbackRows.forEach((row) => {
    const orderId = cleanOrderId(
      row['Order ID'] ?? row['Order Id'] ?? row['OrderID'] ?? row['IRCTC Order ID']
    );
    if (orderId && !feedbackRowMap.has(orderId)) {
      feedbackRowMap.set(orderId, row);
    }
  });

  // STEP 5: Aggregate by Outlet ID.
  const outletMap: Record<string, FeedbackReportRow> = {};

  Object.entries(orderCounts).forEach(([orderId, counts]) => {
    const feedbackRow = feedbackRowMap.get(orderId);
    const irctc = irctcOrderMap.get(orderId);

    // IMPORTANT:
    // Final Feedback Report ka Outlet ID SIRF IRCTC report se liya jayega.
    // Feedback file ka Outlet ID kabhi use nahi hoga.
    //
    // Primary join:
    // Feedback Order ID -> IRCTC Order Id -> IRCTC Outlet Id
    if (!irctc) return;

    const outletId = cleanOutletId(
      irctc['Outlet Id'] ?? irctc['Outlet ID'] ?? irctc['OutletId']
    );

    if (!outletId) return;

    const irctcOutlet = irctc;
    const masterOutlet = masterOutletMap.get(outletId) || {};

    const outletName = String(
      irctcOutlet['Outlet Name'] ??
      feedbackRow?.['Outlet'] ??
      masterOutlet['outletName'] ??
      masterOutlet['Outlet Name'] ??
      ''
    ).trim();

    const stationCode = String(
      irctcOutlet['Delivery Station'] ??
      irctcOutlet['Station Code'] ??
      masterOutlet['station'] ??
      masterOutlet['Station'] ??
      ''
    ).trim();

    if (!outletMap[outletId]) {
      outletMap[outletId] = {
        'Outlet Id': outletId,
        'Outlet Name': outletName,
        'Station Code': stationCode,
        Complaint: 0,
        Feedback: 0,
        'Current Count': 0,
        'Current Rating': 0,
        'Current Sum': 0,
        'Old Count': 0,
        'Old Ratings': 0,
        'Old Sum': 0,
        'Total Count': 0,
        'Total Rating Sum': 0,
        'Total Rating': 0,
      };
    } else {
      if (!outletMap[outletId]['Outlet Name'] && outletName) {
        outletMap[outletId]['Outlet Name'] = outletName;
      }
      if (!outletMap[outletId]['Station Code'] && stationCode) {
        outletMap[outletId]['Station Code'] = stationCode;
      }
    }

    outletMap[outletId].Complaint += counts.complaint;
    outletMap[outletId].Feedback += counts.feedback;
  });

  // STEP 6: If Feedback file is not available in IndexedDB (for example the
  // user installed the new page after an older merge), recover the report from
  // the IRCTC Feedback Type data. This prevents the report from becoming blank
  // solely because the optional Feedback file was not persisted by the old build.
  if (Object.keys(outletMap).length === 0 && irctcRows.length > 0) {
    irctcRows.forEach((row) => {
      const type = normalizeFeedbackType(
        row['Feedback Type'] ?? row['FeedbackType'] ?? row['Type']
      );
      if (!type) return;

      const outletId = cleanOutletId(
        row['Outlet Id'] ?? row['Outlet ID'] ?? row['OutletId']
      );
      if (!outletId) return;

      if (!outletMap[outletId]) {
        outletMap[outletId] = {
          'Outlet Id': outletId,
          'Outlet Name': String(row['Outlet Name'] || '').trim(),
          'Station Code': String(row['Delivery Station'] || row['Station Code'] || '').trim(),
          Complaint: 0,
          Feedback: 0,
          'Current Count': 0,
          'Current Rating': 0,
          'Current Sum': 0,
          'Old Count': 0,
          'Old Ratings': 0,
          'Old Sum': 0,
          'Total Count': 0,
          'Total Rating Sum': 0,
          'Total Rating': 0,
        };
      }

      if (type === 'Complaint') outletMap[outletId].Complaint += 1;
      if (type === 'Feedback') outletMap[outletId].Feedback += 1;
    });
  }


  // STEP 7: Merge historical Old Feedback / Old Ratings by Outlet ID.
  // Old data is used only when the Outlet ID already exists in the current
  // IRCTC-matched report. Old-only outlets are NOT added.
  const oldMap: Record<string, {
    outletId: string;
    outletName: string;
    stationCode: string;
    oldCount: number;
    oldRatings: number;
    oldSum: number;
  }> = {};

  oldRatingsRows.forEach((row: any) => {
    const outletId = cleanOutletId(
      row['Outlet Id'] ??
      row['Outlet ID'] ??
      row['OutletId'] ??
      row['Aggregator Outlet ID']
    );
    if (!outletId) return;

    const num = (v: any) => {
      const n = Number(String(v ?? '').replace(/,/g, '').trim());
      return Number.isFinite(n) ? n : 0;
    };

    const oldCount = num(row['Old Count'] ?? row['OldCount'] ?? row['Count']);
    const oldRatings = num(row['Old Ratings'] ?? row['Old Rating'] ?? row['Rating']);
    const rawOldSum = row['Old Sum'] ?? row['OldSum'];
    const oldSum = rawOldSum !== undefined && rawOldSum !== ''
      ? num(rawOldSum)
      : Number((oldCount * oldRatings).toFixed(2));

    oldMap[outletId] = {
      outletId,
      outletName: String(row['Outlet Name'] ?? row['Outlet'] ?? '').trim(),
      stationCode: String(
        row['Station Code'] ?? row['Station'] ?? row['Delivery Station'] ?? ''
      ).trim().toUpperCase(),
      oldCount,
      oldRatings,
      oldSum,
    };
  });

  // IMPORTANT:
  // Old Ratings / Old Feedback is the FINAL OUTLET MASTER for this report.
  // Every Outlet ID present here will appear in the final report.
  // Current Feedback/IRCTC data only fills Complaint and Feedback counts
  // when a current match exists.
  Object.values(oldMap).forEach((old) => {
    const current = outletMap[old.outletId];
    if (!current) return;

    if (!current['Outlet Name'] && old.outletName) {
      current['Outlet Name'] = old.outletName;
    }
    if (!current['Station Code'] && old.stationCode) {
      current['Station Code'] = old.stationCode;
    }

    current['Old Count'] = old.oldCount;
    current['Old Ratings'] = old.oldRatings;
    current['Old Sum'] = old.oldSum;
  });

  // STEP 8: Build the FINAL outlet list.
  //
  // IMPORTANT BUSINESS RULE:
  // The final Feedback Report's Outlet ID list MUST come from the
  // Old Feedback / Old Ratings file.
  //
  // Current Feedback + IRCTC data only supplies the current
  // Complaint / Feedback counts for those Outlet IDs.
  //
  // Therefore:
  //   Old Ratings Outlet IDs = FINAL BASE ROWS
  //   Current matched counts = attached to those rows
  //
  // Current-only / IRCTC-only outlets are NOT added to this report.
  const finalOutletMap: Record<string, FeedbackReportRow> = {};

  if (Object.keys(oldMap).length > 0) {
    Object.values(oldMap).forEach((old) => {
      const current = outletMap[old.outletId];

      finalOutletMap[old.outletId] = {
        'Outlet Id': old.outletId,
        'Outlet Name': old.outletName || current?.['Outlet Name'] || '',
        'Station Code': old.stationCode || current?.['Station Code'] || '',
        Complaint: current?.Complaint || 0,
        Feedback: current?.Feedback || 0,
        'Current Count': 0,
        'Current Rating': 0,
        'Current Sum': 0,
        'Old Count': old.oldCount,
        'Old Ratings': old.oldRatings,
        'Old Sum': old.oldSum,
        'Total Count': 0,
        'Total Rating Sum': 0,
        'Total Rating': 0,
      };
    });
  } else {
    // If Old Feedback / Old Ratings has not been uploaded, preserve the
    // existing current-report behavior rather than returning an empty report.
    Object.entries(outletMap).forEach(([outletId, row]) => {
      finalOutletMap[outletId] = row;
    });
  }

  // STEP 9: Rating calculations.
  //
  // Current Rating = Feedback * 5 / (Feedback + Complaint)
  //
  // Current Sum = Current Count * Current Rating
  //
  // Till Date Ratings =
  //   (Old Sum + Current Sum)
  //   / (Old Count + Current Count)
  //
  // Complaint contributes to count but gives 0 rating points.
  return Object.values(finalOutletMap)
    .map((row) => {
      const currentCount = row.Complaint + row.Feedback;

      // Current Rating = Feedback * 5 / (Feedback + Complaint)
      const currentRating = currentCount > 0
        ? Number(((row.Feedback * 5) / currentCount).toFixed(2))
        : 0;

      // Current Sum = Current Count * Current Rating
      // Uses the displayed/rounded Current Rating exactly as requested.
      const currentSum = Number(
        (currentCount * currentRating).toFixed(2)
      );

      // Total Count = Old Count + Current Count
      const totalCount = row['Old Count'] + currentCount;

      // Total Rating Sum = Old Sum + Current Sum
      const totalRatingSum = Number(
        (row['Old Sum'] + currentSum).toFixed(2)
      );

      // Till Date Ratings = (Old Sum + Current Sum)
      //                    / (Old Count + Current Count)
      const tillDateRatings = totalCount > 0
        ? Number((totalRatingSum / totalCount).toFixed(2))
        : 0;

      return {
        ...row,
        'Current Count': currentCount,
        'Current Rating': currentRating,
        'Current Sum': currentSum,
        'Total Count': totalCount,
        'Total Rating Sum': totalRatingSum,
        'Total Rating': tillDateRatings,
      };
    })
    .sort((a, b) => {
      const outletCompare = a['Outlet Id'].localeCompare(
        b['Outlet Id'],
        undefined,
        { numeric: true }
      );
      if (outletCompare !== 0) return outletCompare;
      return a['Outlet Name'].localeCompare(b['Outlet Name']);
    });

};

const generateFeedbackReportWorkbook = (rows: FeedbackReportRow[]) => {
  if (!rows.length) {
    alert('Feedback Report ke liye data available nahi hai. Current Feedback + IRCTC ya Old Feedback file check karein.');
    return;
  }

  const exportRows = rows.map((r) => ({
    'Outlet Id': r['Outlet Id'],
    'Outlet Name': r['Outlet Name'],
    'Station Code': r['Station Code'],
    'Old Count': r['Old Count'],
    'Old Ratings': r['Old Ratings'],
    'Old Sum': r['Old Sum'],
    Complaint: r.Complaint,
    Feedback: r.Feedback,
    'Current Count': r['Current Count'],
    'Current Rating': r['Current Rating'],
    'Current Sum': r['Current Sum'],
    'Total Count': r['Total Count'],
    'Total Rating Sum': r['Total Rating Sum'],
    'Till Date Ratings': r['Total Rating'],
  }));

  const ws = XLSX.utils.json_to_sheet(exportRows);
  ws['!cols'] = [
    { wch: 16 }, { wch: 42 }, { wch: 16 },
    { wch: 12 }, { wch: 13 }, { wch: 13 },
    { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 13 }, { wch: 16 }, { wch: 14 },
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Feedback Report');
  XLSX.writeFile(wb, `FEEDBACK_REPORT_${new Date().toISOString().slice(0, 10)}.xlsx`);
};

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

interface MetricStats {
  orders: number;
  deliveredOrders: number;
  meals: number;
  value: number;
  prepaidValue: number;
  discount: number;
  revenue: number;
  complaints: number;
  feedback: number;
  undelivered: number;
  outletsSet: Set<string>;
}

const createEmptyStats = (): MetricStats => ({
  orders: 0,
  deliveredOrders: 0,
  meals: 0,
  value: 0,
  prepaidValue: 0,
  discount: 0,
  revenue: 0,
  complaints: 0,
  feedback: 0,
  undelivered: 0,
  outletsSet: new Set<string>(),
});

const getSourceChannel = (row: any): string => {
  const channel = String(row['Source'] || row['Channel'] || row['Booking Channel'] || '').toUpperCase();
  if (channel.includes('MMT') || channel.includes('MAKEMYTRIP')) return 'MakeMyTrip';
  if (channel.includes('APP') || channel.includes('REL_APP')) return 'REL_Food_App';
  if (channel.includes('WEB') || channel.includes('WEBSITE')) return 'RELFood_WEBSITE';
  return 'RELFood_IRCTC';
};

// --- Date Engine: DD/MM/YYYY-safe + Excel serial + ISO/Date support ---
const parseReportDate = (dateVal: any): Date | null => {
  if (dateVal === null || dateVal === undefined || dateVal === '') return null;

  // IMPORTANT:
  // The actual report is the 1-Aug-2026 to 10-Aug-2026 report.
  // In the uploaded Excel, these dates can arrive through XLSX as Date/serial
  // values that have already been interpreted as:
  //   Jan 8 2026 -> source 08/01/2026 -> 1 Aug 2026
  //   Feb 8 2026 -> source 08/02/2026 -> 2 Aug 2026
  //   ...
  //   Oct 8 2026 -> source 08/10/2026 -> 10 Aug 2026
  //
  // This happens because the source date is MM/DD/YYYY but an earlier parser
  // interpreted it as DD/MM/YYYY. We correct that exact 2026 August pattern
  // BEFORE any normal Date handling.

  const correctAugustCorruptedDate = (year: number, monthIndex: number, day: number): Date | null => {
    // Jan 8 through Oct 8 -> Aug 1 through Aug 10.
    if (
      year === 2026 &&
      day === 8 &&
      monthIndex >= 0 &&
      monthIndex <= 9
    ) {
      return new Date(2026, 7, monthIndex + 1);
    }
    return null;
  };

  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return null;

    const year = dateVal.getFullYear();
    const monthIndex = dateVal.getMonth();
    const day = dateVal.getDate();

    const corrected = correctAugustCorruptedDate(year, monthIndex, day);
    if (corrected) return corrected;

    return new Date(year, monthIndex, day);
  }

  const raw = String(dateVal).trim();
  if (!raw) return null;

  // Excel serial date.
  const numeric = Number(raw);
  if (Number.isFinite(numeric) && numeric > 30000 && numeric < 70000) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    const d = new Date(excelEpoch + numeric * 86400000);

    const year = d.getUTCFullYear();
    const monthIndex = d.getUTCMonth();
    const day = d.getUTCDate();

    const corrected = correctAugustCorruptedDate(year, monthIndex, day);
    if (corrected) return corrected;

    return new Date(year, monthIndex, day);
  }

  // SOURCE FORMAT IS MM/DD/YYYY.
  // Examples:
  //   08/01/2026 = 1 August 2026
  //   08/02/2026 = 2 August 2026
  //   ...
  //   08/10/2026 = 10 August 2026
  const datePart = raw.split(/[T ]/)[0];

  let match = datePart.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const year = Number(match[3]);

    const d = new Date(year, month - 1, day);
    if (
      d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day
    ) {
      return d;
    }
    return null;
  }

  // YYYY-MM-DD / YYYY/MM/DD.
  match = datePart.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const d = new Date(year, month - 1, day);
    if (
      d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day
    ) {
      return d;
    }
    return null;
  }

  const fallback = new Date(raw);
  if (isNaN(fallback.getTime())) return null;

  const fallbackYear = fallback.getFullYear();
  const fallbackMonth = fallback.getMonth();
  const fallbackDay = fallback.getDate();

  const corrected = correctAugustCorruptedDate(
    fallbackYear,
    fallbackMonth,
    fallbackDay
  );
  if (corrected) return corrected;

  return new Date(fallbackYear, fallbackMonth, fallbackDay);
};

// SOURCE FILE DATE CONVENTION: MM/DD/YYYY.
// VENDOR DATE WISE CURRENT REPORT: 1-Aug-2026 through 10-Aug-2026.
// SOURCE FILE DATE CONVENTION: MM/DD/YYYY (not DD/MM/YYYY).
const reportDateKey = (dateVal: any): string => {
  const d = parseReportDate(dateVal);
  if (!d) return 'UNKNOWN';

  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
};

const formatFullDisplayDate = (dateVal: any): string => {
  const d = parseReportDate(dateVal);
  if (!d) return String(dateVal || 'Unknown Date');

  return d.toLocaleDateString('en-GB', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
};

// Feedback Date Engine: Main Report Feedback/Complaint FTD must use
// the Feedback upload's exact `Created At` date, not Delivery Date/Rating/Remarks.
// Supports Date objects, Excel serials, ISO timestamps and MM/DD/YYYY timestamps.
const parseFeedbackCreatedAt = (dateVal: any): Date | null => {
  if (dateVal === null || dateVal === undefined || String(dateVal).trim() === '') return null;

  if (dateVal instanceof Date) {
    return isNaN(dateVal.getTime()) ? null : new Date(dateVal.getTime());
  }

  const raw = String(dateVal).trim();
  const numeric = Number(raw);
  if (Number.isFinite(numeric) && numeric > 30000 && numeric < 70000) {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + numeric * 86400000);
  }

  // Feedback CSV uses values like:
  // 2026-08-01 18:52:14 IST
  // JavaScript's Date parser can reject the trailing `IST`, so parse the
  // leading YYYY-MM-DD explicitly before falling back to native parsing.
  const isoMatch = raw.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) {
      return d;
    }
  }

  const iso = new Date(raw);
  const datePart = raw.split(/[T ]/)[0];
  const slash = datePart.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (slash) {
    const first = Number(slash[1]);
    const second = Number(slash[2]);
    const year = Number(slash[3]);
    // Feedback export uses the same MM/DD/YYYY convention as the current source files.
    // If the first part cannot be a month, safely fall back to DD/MM/YYYY.
    const month = first <= 12 ? first : second;
    const day = first <= 12 ? second : first;
    const d = new Date(year, month - 1, day);
    if (d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day) return d;
  }

  return isNaN(iso.getTime()) ? null : iso;
};

const feedbackCreatedAtKey = (dateVal: any): string => {
  const d = parseFeedbackCreatedAt(dateVal);
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

type ReportType =
  | 'MASTER'
  | 'MAIN_REPORT'
  | 'VENDOR_RDS'
  | 'STATION_REPORT'
  | 'VENDOR_REPORT'
  | 'DATE_WISE'
  | 'VENDOR_DATE_WISE'
  | 'LAST_DAY_STATION'
  | 'OUTLETS_MASTER'
  | 'PENALTIES'
  | 'FEEDBACK_REPORT';

// EXACT column order used by lib/vendorRdsGenerator.ts / Vendor RDS Excel.
// The first three columns are intentionally fixed in the on-screen table.
const VENDOR_REPORT_COLUMNS = [
  'Aggregator Outlet ID', 'Station Code', 'Rank', 'Station Name', 'Vendor Name',
  'Vendor Price', 'Final Base Price', 'Final Total Commission', 'Final IRCTC Comm',
  'Final RF Commission', 'Final GST', 'Final Discount', 'Final Vendor Discount',
  'Final RF Discount', 'Delivery Charges', 'Final Selling Price', 'Final Order Total',
  'Discounted Base Price', 'PPD', 'COD', 'Meals', 'Check',
  'Count of Delivered Orders', 'Count of Not_Delivered As per IRCTC Status',
  'Not_Delivered %', 'Prepaid %', 'Vendor Payment Type', 'Discount Applied',
  'Previouse Balance', 'Paid to Vendors By Relfood',
  'Payment Received from Vendor to Relfood', 'Net Payment'
] as const;

const VENDOR_RDS_COLUMNS = [
  'Outlet ID',
  'Vendor Name',
  'Station Code',
  'GST Number',
  'Vendor with Station Code',
  'Invoice Number',
  'Final Vendor Price',
  'Final Base Price',
  'Final RF Commission',
  'Final IRCTC Commission',
  'Final GST',
  'Final Order Total',
  'Final Total Commission',
  'Penalty',
  'Gross Commission',
  'IGST',
  'CGST',
  'SGST',
  'IGST+CGST+SGST',
  'Total This Month',
  'Paid to Vendors By Relfood',
  'Delivery Charges',
  'Previouse Balance',
  'Final Selling Price',
  'Final Total Discount',
  'Final Vendor Discount',
  'Payment Received from Vendor to Relfood',
  'Credit Note to Vendor by Relfood',
  'Final RF Discount',
  'PPD',
  'COD',
  'ADD',
  'Less',
  'Net Payment',
  'As per Reverse',
  'Diff',
  'Orders Count',
  'Meals',
  'State',
  'Discounted Base Price',
  'Margin %',
  'Vendor Payment Type',
  'Discount Applied',
] as const;

const VENDOR_RDS_MONEY_COLUMNS = new Set([
  'Final Vendor Price', 'Final Base Price', 'Final RF Commission',
  'Final IRCTC Commission', 'Final GST', 'Final Order Total',
  'Final Total Commission', 'Penalty', 'Gross Commission', 'IGST',
  'CGST', 'SGST', 'IGST+CGST+SGST', 'Total This Month',
  'Paid to Vendors By Relfood', 'Delivery Charges', 'Previouse Balance',
  'Final Selling Price', 'Final Total Discount', 'Final Vendor Discount',
  'Payment Received from Vendor to Relfood', 'Credit Note to Vendor by Relfood',
  'Final RF Discount', 'PPD', 'COD', 'ADD', 'Less', 'Net Payment',
  'As per Reverse', 'Diff', 'Discounted Base Price', 'Margin %',
]);

const VENDOR_RDS_FIXED_WIDTHS = {
  outletId: 100,
  vendorName: 350,
  stationCode: 140,
};

const formatVendorRdsCell = (key: string, value: any): string => {
  if (value === null || value === undefined || value === '') return '-';
  if (VENDOR_RDS_MONEY_COLUMNS.has(key)) {
    const n = Number(value);
    return Number.isFinite(n) ? n.toFixed(2) : String(value);
  }
  return String(value);
};

// Excel-style multi-select filter used by the Vendor Report dashboard header.
function VendorHeaderFilter({ column, options, selected, position, onToggle, onSelectAll, onClose }: {
  column: string;
  options: { key: string; label: string }[];
  selected?: string[];
  position: { top: number; left: number; width: number };
  onToggle: (key: string) => void;
  onSelectAll: () => void;
  onClose: () => void;
}) {
  const menuRef = useRef<HTMLDivElement | null>(null);
  const [optionSearch, setOptionSearch] = useState('');
  useEffect(() => {
    const onMouseDown = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) onClose();
    };
    const onKeyDown = (event: KeyboardEvent) => { if (event.key === 'Escape') onClose(); };
    document.addEventListener('mousedown', onMouseDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onMouseDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [onClose]);
  const filteredOptions = options.filter((option) => option.label.toLowerCase().includes(optionSearch.toLowerCase()));
  const allSelected = selected === undefined;
  const selectedSet = new Set(selected || []);
  return (
    <div ref={menuRef} className="fixed z-[9999] rounded-lg border border-slate-300 bg-white text-slate-800 shadow-2xl"
      style={{ top: position.top, left: Math.max(8, Math.min(position.left, window.innerWidth - Math.max(300, position.width) - 8)), width: Math.max(300, position.width) }}
      onMouseDown={(e) => e.stopPropagation()}>
      <div className="flex items-center justify-between border-b border-slate-200 px-3 py-2">
        <div className="max-w-[230px] truncate text-sm font-bold">Filter: {column}</div>
        <button type="button" onClick={onClose} className="rounded px-2 py-1 text-slate-500 hover:bg-slate-100">✕</button>
      </div>
      <div className="border-b border-slate-200 p-2">
        <input autoFocus value={optionSearch} onChange={(e) => setOptionSearch(e.target.value)} placeholder="Search values..."
          className="w-full rounded border border-slate-300 px-2 py-1.5 text-sm outline-none focus:border-indigo-500" />
      </div>
      <div className="max-h-[330px] overflow-y-auto p-2">
        <label className="flex cursor-pointer items-center gap-2 rounded px-2 py-1.5 font-semibold hover:bg-slate-100">
          <input type="checkbox" checked={allSelected} onChange={onSelectAll} /><span>(Select All)</span>
        </label>
        <div className="my-1 border-t border-slate-200" />
        {filteredOptions.length === 0 ? <div className="px-2 py-3 text-sm text-slate-500">No values found</div> : filteredOptions.map((option) => (
          <label key={option.key} className="flex cursor-pointer items-center gap-2 rounded px-2 py-1.5 hover:bg-slate-100">
            <input type="checkbox" checked={allSelected || selectedSet.has(option.key)} onChange={() => onToggle(option.key)} />
            <span className="truncate" title={option.label}>{option.label}</span>
          </label>
        ))}
      </div>
      <div className="flex items-center justify-between border-t border-slate-200 px-3 py-2 text-xs text-slate-500">
        <span>{allSelected ? 'All values' : `${selected?.length || 0} selected`}</span>
        <button type="button" onClick={onClose} className="rounded bg-indigo-600 px-3 py-1.5 font-semibold text-white hover:bg-indigo-700">OK</button>
      </div>
    </div>
  );
}

export default function Page() {
  const [data, setData] = useState<any[]>([]);
  // Raw IRCTC report is kept separately so Station Report can count
  // Delivery Station + Feedback Type directly from the source file.
  const [irctcRawData, setIrctcRawData] = useState<any[]>([]);
  const [feedbackRawData, setFeedbackRawData] = useState<any[]>([]);
  // Historical outlet-wise rating/count source (uploaded as Old Feedback / Old Ratings Excel).
  const [oldRatingsRawData, setOldRatingsRawData] = useState<any[]>([]);
  const [penaltySummary, setPenaltySummary] = useState<Record<string, number>>({});
  const [penaltyRawRecords, setPenaltyRawRecords] = useState<any[]>([]);
  const [currentMonthRecords, setCurrentMonthRecords] = useState<any[]>([]);
  const [outletsMasterInfo, setOutletsMasterInfo] = useState<Record<string, any>>({});
  const [isLoaded, setIsLoaded] = useState<boolean>(false);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedReport, setSelectedReport] = useState<ReportType>('MAIN_REPORT');
  const [vendorColumnFilters, setVendorColumnFilters] = useState<Record<string, string[] | undefined>>({});
  const [openVendorFilter, setOpenVendorFilter] = useState<string | null>(null);
  const [vendorFilterPosition, setVendorFilterPosition] = useState({ top: 0, left: 0, width: 320 });
  const [themeMode, setThemeMode] = useState<'day' | 'night'>('day');

  useEffect(() => {
    const saved = window.localStorage.getItem('relfood-theme');
    if (saved === 'night' || saved === 'day') setThemeMode(saved);
  }, []);

  const toggleTheme = () => {
    setThemeMode((prev) => {
      const next = prev === 'day' ? 'night' : 'day';
      window.localStorage.setItem('relfood-theme', next);
      return next;
    });
  };

  const handleTableRowClick = (event: React.MouseEvent<HTMLElement>) => {
    const target = event.target as HTMLElement | null;
    const row = target?.closest('tbody tr') as HTMLTableRowElement | null;
    if (!row) return;

    document.querySelectorAll('.portal-row-selected').forEach((el) => {
      el.classList.remove('portal-row-selected');
    });
    row.classList.add('portal-row-selected');
  };

  // Upload States
  const [rfFile, setRfFile] = useState<File | null>(null);
  const [irctcFile, setIrctcFile] = useState<File | null>(null);
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [oldFeedbackFile, setOldFeedbackFile] = useState<File | null>(null);
  const [penaltyFile, setPenaltyFile] = useState<File | null>(null);
  const [currentMonthFile, setCurrentMonthFile] = useState<File | null>(null);
  const [outletsFile, setOutletsFile] = useState<File | null>(null);

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');

  useEffect(() => {
    const fetchStoredData = async () => {
      try {
        const storedMaster = await loadFromDB('CURRENT_MASTER_DATA');
        const storedIrctc = await loadFromDB('CURRENT_IRCTC_DATA');
        const storedFeedback = await loadFromDB('CURRENT_FEEDBACK_DATA');
        const storedOldRatings = await loadFromDB('CURRENT_OLD_RATINGS_DATA');
        const storedPenalty = await loadFromDB('OUTLET_PENALTY_DATA');
        const storedCurrentMonth = await loadFromDB('CURRENT_MONTH_DATA');
        const storedOutletsInfo = await loadFromDB('OUTLET_MASTER_INFO');

        if (Array.isArray(storedMaster) && storedMaster.length > 0) setData(storedMaster);
        if (Array.isArray(storedIrctc) && storedIrctc.length > 0) setIrctcRawData(storedIrctc);
        if (Array.isArray(storedFeedback) && storedFeedback.length > 0) setFeedbackRawData(storedFeedback);
        if (Array.isArray(storedOldRatings) && storedOldRatings.length > 0) setOldRatingsRawData(storedOldRatings);
        if (storedPenalty && typeof storedPenalty === 'object') {
          setPenaltySummary(storedPenalty.outletTotals || {});
          setPenaltyRawRecords(storedPenalty.records || []);
        }
        if (Array.isArray(storedCurrentMonth) && storedCurrentMonth.length > 0) {
          setCurrentMonthRecords(storedCurrentMonth);
        }
        if (storedOutletsInfo && typeof storedOutletsInfo === 'object') {
          setOutletsMasterInfo(storedOutletsInfo);
        }
      } catch (err) {
        console.error('Failed to load DB:', err);
      } finally {
        setIsLoaded(true);
      }
    };
    fetchStoredData();
  }, []);

  const handleClearRecords = async () => {
    if (confirm('Kya aap sach me saara stored master data delete karna chahte hain?')) {
      await clearDB();
      setData([]);
      setIrctcRawData([]);
      setFeedbackRawData([]);
      setOldRatingsRawData([]);
      setPenaltySummary({});
      setPenaltyRawRecords([]);
      setCurrentMonthRecords([]);
      setOutletsMasterInfo({});
    }
  };

  const parseAnyFile = async (file: File): Promise<any[]> => {
    const arrayBuffer = await file.arrayBuffer();
    const uint = new Uint8Array(arrayBuffer.slice(0, 4));
    const isZip = uint[0] === 0x50 && uint[1] === 0x4B;

    if (isZip || file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
      try {
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
      } catch (err) {
        console.warn('XLSX fallback:', err);
      }
    }

    const text = new TextDecoder('utf-8').decode(arrayBuffer);
    if (text.includes('<table') || text.includes('<TABLE')) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(text, 'text/html');
      const table = doc.querySelector('table');
      if (table) {
        const rows = Array.from(table.querySelectorAll('tr'));
        if (rows.length >= 2) {
          const headers = Array.from(rows[0].querySelectorAll('th, td')).map((c: any) => c.innerText.trim());
          const parsedData: any[] = [];
          for (let i = 1; i < rows.length; i++) {
            const cells = Array.from(rows[i].querySelectorAll('td')).map((c: any) => c.innerText.trim());
            if (cells.length === headers.length) {
              const rowObj: any = {};
              headers.forEach((h, idx) => {
                rowObj[h] = cells[idx];
              });
              parsedData.push(rowObj);
            }
          }
          if (parsedData.length > 0) return parsedData;
        }
      }
    }

    return new Promise((resolve) => {
      Papa.parse(text, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => resolve(results.data),
        error: () => resolve([]),
      });
    });
  };

  const handleProcessAndMerge = async () => {
    // If the six core reports are already stored, Old Feedback/Ratings or the
    // current Feedback CSV can be uploaded alone. This avoids forcing the user
    // to upload all six reports again just to refresh the Feedback Report.
    const hasStoredCoreData = data.length > 0 && irctcRawData.length > 0;
    if ((!rfFile || !irctcFile) && (oldFeedbackFile || feedbackFile) && hasStoredCoreData) {
      try {
        setIsProcessing(true);

        if (feedbackFile) {
          setStatusText('Updating Current Feedback Data...');
          const freshFeedback = await parseAnyFile(feedbackFile);
          await saveToDB('CURRENT_FEEDBACK_DATA', freshFeedback);
          setFeedbackRawData(freshFeedback);
        }

        if (oldFeedbackFile) {
          setStatusText('Updating Old Feedback / Old Ratings...');
          const freshOldRatings = await parseAnyFile(oldFeedbackFile);
          await saveToDB('CURRENT_OLD_RATINGS_DATA', freshOldRatings);
          setOldRatingsRawData(freshOldRatings);
        }

        setStatusText('Feedback Report updated successfully.');
        setFeedbackFile(null);
        setOldFeedbackFile(null);
        setIsProcessing(false);
        setIsModalOpen(false);
        return;
      } catch (error: any) {
        console.error(error);
        alert('Error updating Feedback data: ' + error.message);
        setIsProcessing(false);
        return;
      }
    }

    if (!rfFile || !irctcFile) {
      alert('Kripya RF Report aur IRCTC Report upload karein. Ya agar existing 6 reports already stored hain, to sirf Feedback / Old Feedback file upload kar sakte hain.');
      return;
    }

    try {
      setIsProcessing(true);

      setStatusText('Reading RF Report...');
      const rfData = await parseAnyFile(rfFile);

      setStatusText('Reading IRCTC Report...');
      const irctcData = await parseAnyFile(irctcFile);
      // IMPORTANT: preserve the complete raw IRCTC source for Station Report feedback counts.
      await saveToDB('CURRENT_IRCTC_DATA', irctcData);
      setIrctcRawData(irctcData);

      let feedbackData: any[] = feedbackRawData;
      if (feedbackFile) {
        setStatusText('Reading Feedback Report...');
        feedbackData = await parseAnyFile(feedbackFile);
        await saveToDB('CURRENT_FEEDBACK_DATA', feedbackData);
        setFeedbackRawData(feedbackData);
      }

      // Historical rating/count file is optional. If it is not selected in a
      // later merge, keep the previously stored Old Feedback data.
      let oldRatingsData: any[] = oldRatingsRawData;
      if (oldFeedbackFile) {
        // Old Ratings is only ~1k outlet rows; parsing/storage is lightweight.
        // The expensive part was the Feedback Report calculation, not this file.
        setStatusText('Reading Old Feedback / Old Ratings...');
        oldRatingsData = await parseAnyFile(oldFeedbackFile);
        await saveToDB('CURRENT_OLD_RATINGS_DATA', oldRatingsData);
        setOldRatingsRawData(oldRatingsData);
      }

      const penaltyOutletMap: Record<string, number> = {};
      const penaltyRawFilteredList: any[] = [];
      if (penaltyFile) {
        setStatusText('Processing Penalty Report...');
        const penaltyData = await parseAnyFile(penaltyFile);
        const targetModes = ['PENALTY', 'COUPON', 'COMPLAINT_REFUND', 'PARTIAL_DELIVERY'];

        penaltyData.forEach((row) => {
          const mode = String(row['Transaction Mode'] || row['TransactionMode'] || '').trim().toUpperCase();
          if (targetModes.includes(mode)) {
            const outletId = cleanOutletId(row['Outlet Id'] || row['Outlet ID'] || row['OutletId'] || '');
            const amount = parseFloat(row['Amount'] || row['Total Payable'] || 0) || 0;

            if (outletId) {
              penaltyOutletMap[outletId] = Number(((penaltyOutletMap[outletId] || 0) + amount).toFixed(2));
            }

            penaltyRawFilteredList.push({
              outletId,
              orderId: String(row['Order Id'] || row['Order ID'] || '').trim(),
              mode,
              amount,
              vendorName: row['Vendor Name'] || '',
              date: row['Date'] || '',
              remarks: row['Remarks'] || '',
            });
          }
        });

        await saveToDB('OUTLET_PENALTY_DATA', {
          outletTotals: penaltyOutletMap,
          records: penaltyRawFilteredList,
        });
        setPenaltySummary(penaltyOutletMap);
        setPenaltyRawRecords(penaltyRawFilteredList);
      }

      let currentMonthParsedList: any[] = [];
      if (currentMonthFile) {
        setStatusText('Processing Current Month Data...');
        const curMonthData = await parseAnyFile(currentMonthFile);

        currentMonthParsedList = curMonthData.map((row) => ({
          outletId: cleanOutletId(row['Outlet Id'] || row['Outlet ID'] || row['OutletId'] || ''),
          vendorName: String(row['Vendor Name'] || '').trim(),
          stationCode: String(row['Station Code'] || '').trim(),
          previousBalance: parseFloat(row['Previouse Balance'] || row['Previous Balance'] || 0) || 0,
          paidToVendors: parseFloat(row['Paid to Vendors By Relfood'] || 0) || 0,
          receivedFromVendor: parseFloat(row['Payment Received from Vendor to Relfood'] || 0) || 0,
          creditNoteToVendor: parseFloat(row['Credit Note to Vendor by Relfood'] || 0) || 0,
          vendorPaymentType: String(row['Vendor Payment Type'] || '').trim().toUpperCase(),
          discountApplied: String(row['Discount Applied'] || '').trim(),
        }));

        await saveToDB('CURRENT_MONTH_DATA', currentMonthParsedList);
        setCurrentMonthRecords(currentMonthParsedList);
      }

      const outletsMap: Record<string, any> = {};
      if (outletsFile) {
        setStatusText('Processing Outlets Report...');
        const outletsRawData = await parseAnyFile(outletsFile);

        outletsRawData.forEach((row) => {
          const outletId = cleanOutletId(
            row['Aggregator Outlet ID'] ||
            row['Aggregator Outlet Id'] ||
            row['Outlet Id'] ||
            row['Outlet ID'] ||
            ''
          );

          if (outletId && !outletsMap[outletId]) {
            let gstVal = '';
            for (const k of Object.keys(row)) {
              if (k.toUpperCase().includes('GST')) {
                gstVal = String(row[k] || '').trim();
                if (gstVal) break;
              }
            }

            outletsMap[outletId] = {
              outletId,
              state: String(row['State'] || row['STATE'] || '').trim(),
              gst: gstVal,
              irctcStatus: String(row['Status'] || row['STATUS'] || '').trim(),
              outletName: String(row['Outlet Name'] || '').trim(),
              // Short station code used by existing reports.
              station: String(row['Station'] || row['Station Code'] || '').trim(),
              // Full station name must come from Outlet Master `Station Name`
              // (column O), never from the short Station Code.
              stationName: getOutletMasterStationName(row),
              stationCode: String(row['Station Code'] || row['Station'] || '').trim(),
              stationRank: (() => {
                const rawRank = row['Station Rank'];
                if (rawRank === undefined || rawRank === null || String(rawRank).trim() === '') return '';
                const n = Number(rawRank);
                return Number.isFinite(n) ? n : String(rawRank).trim();
              })(),
            };
          }
        });

        await saveToDB('OUTLET_MASTER_INFO', outletsMap);
        setOutletsMasterInfo(outletsMap);
      }

      setStatusText('Processing 19 Master Calculation Rules...');

      const irctcMap = new Map();
      irctcData.forEach((row) => {
        const orderId = String(row['Order Id'] || row['Order ID'] || '').trim().replace(/\.0$/, '');
        if (orderId) irctcMap.set(orderId, row);
      });

      // Feedback Report needs counts, not a last-row Map.
      // Keep the raw Feedback rows separately and count every row by Order ID + Type.
      const fbCountMap = new Map<string, FeedbackOrderCount>();
      feedbackData.forEach((row) => {
        const orderId = cleanOrderId(row['Order ID'] || row['Order Id'] || row['OrderID']);
        const type = normalizeFeedbackType(row['Type'] || row['Feedback Type'] || row['FeedbackType']);
        if (!orderId || !type) return;
        const current = fbCountMap.get(orderId) || { complaint: 0, feedback: 0 };
        if (type === 'Complaint') current.complaint += 1;
        if (type === 'Feedback') current.feedback += 1;
        fbCountMap.set(orderId, current);
      });

      const masterRows = rfData.map((rf) => {
        const orderId = String(rf['IRCTC OrderId'] || rf['Order Id'] || '').trim().replace(/\.0$/, '');
        const irctc = irctcMap.get(orderId) || {};
        const fbCounts = fbCountMap.get(orderId) || { complaint: 0, feedback: 0 };
        const outletId = getOutletIdFromRow(rf) || getOutletIdFromRow(irctc) || '';
        const outletInfo = outletsMap[outletId] || outletsMasterInfo[outletId] || {};

        const rfRawStatus = rf['Order Status'] || '';
        const irctcRawStatus = irctc['Order Status'] || '';
        const finalStatus = computeFinalStatus(rfRawStatus, irctcRawStatus);
        const finalVendorPrice = parseFloat(rf['Vendor Price'] || 0) || 0;

        const rfSellingAmount = parseFloat(rf['Selling Amount'] || 0) || 0;
        const rfDiscount = parseFloat(rf['Discount'] || 0) || 0;
        const rfGst = parseFloat(rf['GST'] || 0) || 0;
        const irctcDeliveryCharge = parseFloat(irctc['Delivery Charge'] || 0) || 0;

        const finalBasePrice = Number((rfSellingAmount + rfDiscount - rfGst - irctcDeliveryCharge).toFixed(2));
        const finalTotalCommission = Number((finalBasePrice - finalVendorPrice).toFixed(2));
        const finalIRCTCComm = Number((finalBasePrice * 0.15).toFixed(2));
        const finalRFCommission = Number((finalTotalCommission - finalIRCTCComm).toFixed(2));
        const finalTotalDiscount = Number(rfDiscount.toFixed(2));
        const finalVendorDiscount = Number((finalTotalDiscount * 0.5).toFixed(2));
        const finalRFDiscount = Number((finalTotalDiscount * 0.5).toFixed(2));
        const discountedBasePrice = Number((finalBasePrice - finalTotalDiscount).toFixed(2));
        const finalGST = Number((discountedBasePrice * 0.05).toFixed(2));
        const deliveryCharges = Number(irctcDeliveryCharge.toFixed(2));
        const finalSellingPrice = Number((discountedBasePrice + finalGST + deliveryCharges).toFixed(2));
        const finalOrderTotal = Number((finalBasePrice + finalGST + deliveryCharges).toFixed(2));

        const paymentType = String(rf['Payment Type'] || irctc['Transaction Type'] || '').trim().toUpperCase();
        const isPrepaid = paymentType.includes('PRE_PAID') || paymentType.includes('PREPAID') || paymentType.includes('ONLINE');
        const ppd = isPrepaid ? finalSellingPrice : 0;
        const isCOD = paymentType.includes('CASH') || paymentType.includes('COD');
        const cod = isCOD ? finalSellingPrice : 0;

        const meals = parseInt(irctc['Meal Count'] || '1', 10) || 1;
        const marginPct = finalBasePrice > 0 ? Number((((finalBasePrice - finalVendorPrice) / finalBasePrice) * 100).toFixed(2)) : 0;
        const ordersCount = finalStatus === 'Delivered' ? 1 : 0;

        return {
          'IRCTC Order ID': orderId,
          'RF Order ID': rf['Relfood OrderId'] || '',
          'Outlet ID': outletId,
          'Vendor Name': rf['Vendor Name'] || irctc['Vendor Name'] || '',
          'Station Code': rf['Station Code'] || irctc['Delivery Station'] || '',
          'Station Rank': outletInfo.stationRank ?? '',
          // Preserve the exact IRCTC feedback source inside master data.
          // This makes Station Report independent of IndexedDB raw-data state.
          'Delivery Station': irctc['Delivery Station'] || '',
          'Feedback Type': irctc['Feedback Type'] || '',
          'State': outletInfo.state || '',
          'GST No': outletInfo.gst || '',
          'Outlet IRCTC Status': outletInfo.irctcStatus || '',
          'Train No': rf['Train Number'] || irctc['Train No.'] || '',
          'Booking Date': rf['Booking Date'] || irctc['Date of Booking'] || '',
          'Delivery Date': rf['Delivery Date'] || irctc['Delivery Date'] || '',
          'Payment Type': paymentType,
          'RF Status': rfRawStatus,
          'IRCTC Status': irctcRawStatus,
          'Final Status': finalStatus,
          'Final Vendor Price': finalVendorPrice,
          'Final Base Price': finalBasePrice,
          'Final Total Commission': finalTotalCommission,
          'Final IRCTC Commission': finalIRCTCComm,
          'Final RF Commission': finalRFCommission,
          'Final Total Discount': finalTotalDiscount,
          'Final Vendor Discount': finalVendorDiscount,
          'Final RF Discount': finalRFDiscount,
          'Discounted Base Price': discountedBasePrice,
          'Final GST': finalGST,
          'Delivery Charges': deliveryCharges,
          'Final Selling Price': finalSellingPrice,
          'Final Order Total': finalOrderTotal,
          'PPD': ppd,
          'COD': cod,
          'Meals': meals,
          'Margin %': marginPct,
          'Orders Count': ordersCount,
          'Rating': '',
          'Remarks': irctc['Comments'] || '',
          'Feedback Complaint Count': fbCounts.complaint,
          'Feedback Count': fbCounts.feedback,
        };
      });

      await saveToDB('CURRENT_MASTER_DATA', masterRows);
      setData(masterRows);

      setRfFile(null);
      setIrctcFile(null);
      setFeedbackFile(null);
      setOldFeedbackFile(null);
      setPenaltyFile(null);
      setCurrentMonthFile(null);
      setOutletsFile(null);

      setIsProcessing(false);
      setIsModalOpen(false);
    } catch (error: any) {
      console.error(error);
      alert('Error processing reports: ' + error.message);
      setIsProcessing(false);
    }
  };

  // --- Main Report Matrix Engine ---
  const mainReportBlocks = useMemo(() => {
    if (!data || data.length === 0) return [];

    // Orders still drive the operational metrics/date blocks.
    // Feedback/Complaint metrics are intentionally driven separately from
    // feedbackRawData using the Feedback upload's `Created At` date.
    const dateMap: Record<string, any[]> = {};
    const orderSourceMap = new Map<string, string>();

    data.forEach((row) => {
      const rawDate = row['Delivery Date'] || row['Booking Date'] || 'Unknown Date';
      const dateKey = reportDateKey(rawDate);
      if (!dateMap[dateKey]) dateMap[dateKey] = [];
      dateMap[dateKey].push(row);

      const orderId = cleanOrderId(
        row['IRCTC Order ID'] ?? row['Order ID'] ?? row['Order Id']
      );
      if (orderId) orderSourceMap.set(orderId, getSourceChannel(row));
    });

    // Build Feedback/Complaint counts from EVERY row in the uploaded
    // Feedback file, grouped by Created At date + source.
    const feedbackByDateSource: Record<string, Record<string, { complaint: number; feedback: number }>> = {};

    feedbackRawData.forEach((row) => {
      const createdAt = row['Created At'] ?? row['CreatedAt'] ?? row['created_at'];
      const dateKey = feedbackCreatedAtKey(createdAt);
      const orderId = cleanOrderId(
        row['Order ID'] ?? row['Order Id'] ?? row['OrderID'] ?? row['IRCTC Order ID']
      );
      const type = normalizeFeedbackType(
        row['Type'] ?? row['Feedback Type'] ?? row['FeedbackType']
      );
      if (!dateKey || !type) return;

      // Match the feedback row to its master order only for source/channel.
      // The DATE always comes from Feedback.Created At.
      const src = orderSourceMap.get(orderId) || 'RELFood_IRCTC';
      if (!feedbackByDateSource[dateKey]) feedbackByDateSource[dateKey] = {};
      if (!feedbackByDateSource[dateKey][src]) {
        feedbackByDateSource[dateKey][src] = { complaint: 0, feedback: 0 };
      }
      if (type === 'Complaint') feedbackByDateSource[dateKey][src].complaint += 1;
      if (type === 'Feedback') feedbackByDateSource[dateKey][src].feedback += 1;

      // If Feedback has a Created At date with no order on that date,
      // still show the date block so the feedback is never lost.
      if (!dateMap[dateKey]) dateMap[dateKey] = [];
    });

    const sortedDates = Object.keys(dateMap).sort((a, b) => a.localeCompare(b));

    const mtdBySource: Record<string, MetricStats> = {};
    SOURCES.forEach((s) => (mtdBySource[s] = createEmptyStats()));
    const mtdGrandTotal = createEmptyStats();

    return sortedDates.map((dateKey) => {
      const rows = dateMap[dateKey];
      const dayStats: Record<string, MetricStats> = {};
      SOURCES.forEach((s) => (dayStats[s] = createEmptyStats()));
      const dayTotal = createEmptyStats();

      rows.forEach((r) => {
        const src = getSourceChannel(r);
        const isDelivered = r['Final Status'] === 'Delivered';
        const isUndelivered =
          r['Final Status'] === 'Not Delivered' ||
          String(r['IRCTC Status'] || '').toUpperCase().includes('UNDELIVERED');
        const sellingPrice = parseFloat(r['Final Selling Price'] || 0) || 0;
        const discount = parseFloat(r['Final Total Discount'] || 0) || 0;
        const rfComm = parseFloat(r['Final RF Commission'] || 0) || 0;
        const prepaid = parseFloat(r['PPD'] || 0) || 0;
        const mealCount = parseInt(r['Meals'] || '1', 10) || 1;
        const outletId = getOutletIdFromRow(r);

        const sStat = dayStats[src] || dayStats['RELFood_IRCTC'];
        sStat.orders += 1;
        if (isDelivered) sStat.deliveredOrders += 1;
        if (isUndelivered) sStat.undelivered += 1;
        sStat.meals += mealCount;
        sStat.value += sellingPrice;
        sStat.prepaidValue += prepaid;
        sStat.discount += discount;
        sStat.revenue += rfComm;
        // Active restaurants = unique outlets that actually delivered an order on this date.
        if (isDelivered && outletId) sStat.outletsSet.add(outletId);
      });

      // IMPORTANT: Feedback/Complaint FTD comes ONLY from Feedback.Created At.
      // Do not infer it from Rating or Remarks on the master order.
      const feedbackDay = feedbackByDateSource[dateKey] || {};
      SOURCES.forEach((s) => {
        const fb = feedbackDay[s];
        if (!fb) return;
        dayStats[s].complaints += fb.complaint;
        dayStats[s].feedback += fb.feedback;
      });

      SOURCES.forEach((s) => {
        const st = dayStats[s];
        dayTotal.orders += st.orders;
        dayTotal.deliveredOrders += st.deliveredOrders;
        dayTotal.meals += st.meals;
        dayTotal.value += st.value;
        dayTotal.prepaidValue += st.prepaidValue;
        dayTotal.discount += st.discount;
        dayTotal.revenue += st.revenue;
        dayTotal.complaints += st.complaints;
        dayTotal.feedback += st.feedback;
        dayTotal.undelivered += st.undelivered;
        st.outletsSet.forEach((o) => dayTotal.outletsSet.add(o));

        const m = mtdBySource[s];
        m.orders += st.orders;
        m.deliveredOrders += st.deliveredOrders;
        m.meals += st.meals;
        m.value += st.value;
        m.prepaidValue += st.prepaidValue;
        m.discount += st.discount;
        m.revenue += st.revenue;
        m.complaints += st.complaints;
        m.feedback += st.feedback;
        m.undelivered += st.undelivered;
      });

      mtdGrandTotal.orders += dayTotal.orders;
      mtdGrandTotal.deliveredOrders += dayTotal.deliveredOrders;
      mtdGrandTotal.meals += dayTotal.meals;
      mtdGrandTotal.value += dayTotal.value;
      mtdGrandTotal.prepaidValue += dayTotal.prepaidValue;
      mtdGrandTotal.discount += dayTotal.discount;
      mtdGrandTotal.revenue += dayTotal.revenue;
      mtdGrandTotal.complaints += dayTotal.complaints;
      mtdGrandTotal.feedback += dayTotal.feedback;
      mtdGrandTotal.undelivered += dayTotal.undelivered;

      return {
        dateLabel: dateKey === 'UNKNOWN' ? 'Unknown Date' : formatFullDisplayDate(dateKey),
        rawDate: dateKey,
        dayTotal: { ...dayTotal },
        dayStats: JSON.parse(JSON.stringify(dayStats)),
        mtdTotal: { ...mtdGrandTotal },
        mtdBySource: JSON.parse(JSON.stringify(mtdBySource)),
        outletsCount: dayTotal.outletsSet.size,
      };
    });
  }, [data, feedbackRawData]);

  // --- Aggregate Views ---
  const stationSummary = useMemo(() => {
    return generateStationWiseData(data, outletsMasterInfo, irctcRawData);
  }, [data, outletsMasterInfo, irctcRawData]);

  // LAST_DAY_STATION dashboard MUST use the exact same data engine/order as its Excel export.
  // This prevents the dashboard from showing the generic Station Report data.
  const lastDayStationSummary = useMemo(() => {
    const report = generateLastDayStationWiseData(data, outletsMasterInfo, irctcRawData);
    if (!report) return [];

    return report.rowsData.map((values: any[]) => {
      const row: Record<string, any> = {};
      report.headers.forEach((header: string, index: number) => {
        row[header] = values[index];
      });
      return row;
    });
  }, [data, outletsMasterInfo, irctcRawData]);

  const vendorSummary = useMemo(() => {
    return generateVendorWiseData(data, outletsMasterInfo, penaltySummary, currentMonthRecords);
  }, [data, outletsMasterInfo, penaltySummary, currentMonthRecords]);

  // Total row stays above the header and is based on the complete report,
  // not on the current search filter.
  const vendorReportTotals = useMemo(() => {
    const numericColumns = [
      'Vendor Price', 'Final Base Price', 'Final Total Commission',
      'Final IRCTC Comm', 'Final RF Commission', 'Final GST', 'Final Discount',
      'Final Vendor Discount', 'Final RF Discount', 'Delivery Charges',
      'Final Selling Price', 'Final Order Total', 'Discounted Base Price',
      'PPD', 'COD', 'Meals', 'Count of Delivered Orders',
      'Count of Not_Delivered As per IRCTC Status',
      'Previouse Balance', 'Paid to Vendors By Relfood', 'Payment Received from Vendor to Relfood', 'Net Payment',
    ];
    const totals: Record<string, any> = {};
    numericColumns.forEach((column) => {
      totals[column] = Number(
        vendorSummary.reduce((sum: number, row: any) => sum + (Number(row[column]) || 0), 0).toFixed(2)
      );
    });
    const vendorPrice = Number(totals['Vendor Price'] || 0);
    const sellingPrice = Number(totals['Final Selling Price'] || 0);
    const delivered = Number(totals['Count of Delivered Orders'] || 0);
    const notDelivered = Number(totals['Count of Not_Delivered As per IRCTC Status'] || 0);

    // Business outlets for the payment/discount summary: only outlets with
    // actual business in both Vendor Price and Final Base Price. This matches
    // the requested Excel summary (e.g. PPD 85 / COD 221 and No 99 / Yes 207)
    // and excludes rows where the calculated Base Price exists but Vendor
    // business itself is zero.
    const businessOutlets = vendorSummary.filter((row: any) =>
      Number(row['Vendor Price'] || 0) > 0 && Number(row['Final Base Price'] || 0) > 0
    );
    const ppdBusinessCount = businessOutlets.filter(
      (row: any) => String(row['Vendor Payment Type'] || '').trim().toUpperCase() === 'PPD'
    ).length;
    const codBusinessCount = businessOutlets.filter(
      (row: any) => String(row['Vendor Payment Type'] || '').trim().toUpperCase() === 'COD'
    ).length;
    const discountNoBusinessCount = businessOutlets.filter(
      (row: any) => String(row['Discount Applied'] || '').trim().toUpperCase() === 'NO'
    ).length;
    const discountYesBusinessCount = businessOutlets.filter(
      (row: any) => String(row['Discount Applied'] || '').trim().toUpperCase() === 'YES'
    ).length;

    totals['Vendor Payment Type'] = `PPD ${ppdBusinessCount} COD ${codBusinessCount}`;
    totals['Discount Applied'] = `No ${discountNoBusinessCount} Yes ${discountYesBusinessCount}`;
    totals['Check'] = vendorPrice > 0
      ? `${((Number(totals['Final Total Commission'] || 0) / vendorPrice) * 100).toFixed(2)}%`
      : '#DIV/0!';
    totals['Not_Delivered %'] = delivered > 0
      ? `${((notDelivered / delivered) * 100).toFixed(2)}%`
      : notDelivered > 0 ? '100.00%' : '#DIV/0!';
    totals['Prepaid %'] = sellingPrice > 0
      ? `${((Number(totals['PPD'] || 0) / sellingPrice) * 100).toFixed(2)}%`
      : '#DIV/0!';
    return totals;
  }, [vendorSummary]);

  const vendorFilterOptions = useMemo(() => {
    const result: Record<string, { key: string; label: string }[]> = {};
    VENDOR_REPORT_COLUMNS.forEach((column) => {
      const seen = new Map<string, string>();
      vendorSummary.forEach((row: any) => {
        const raw = row[column];
        const key = raw === null || raw === undefined || String(raw).trim() === '' ? '__BLANK__' : String(raw);
        if (!seen.has(key)) {
          const label = key === '__BLANK__' ? '(Blanks)' : (typeof raw === 'number' ? raw.toLocaleString('en-IN', { maximumFractionDigits: 2 }) : key);
          seen.set(key, label);
        }
      });
      result[column] = Array.from(seen.entries()).map(([key, label]) => ({ key, label })).sort((a, b) => {
        if (a.key === '__BLANK__') return 1;
        if (b.key === '__BLANK__') return -1;
        return a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: 'base' });
      });
    });
    return result;
  }, [vendorSummary]);

  const filteredVendorSummary = useMemo(() => {
    const q = searchTerm.trim().toLowerCase();
    return vendorSummary.filter((row: any) => {
      const matchesSearch = !q ||
        String(row['Vendor Name'] || '').toLowerCase().includes(q) ||
        String(row['Aggregator Outlet ID'] || '').includes(searchTerm) ||
        String(row['Station Code'] || '').toLowerCase().includes(q) ||
        String(row['Station Name'] || '').toLowerCase().includes(q);
      if (!matchesSearch) return false;
      return VENDOR_REPORT_COLUMNS.every((column) => {
        const selected = vendorColumnFilters[column];
        if (selected === undefined) return true;
        const raw = row[column];
        const key = raw === null || raw === undefined || String(raw).trim() === '' ? '__BLANK__' : String(raw);
        return selected.includes(key);
      });
    });
  }, [vendorSummary, searchTerm, vendorColumnFilters]);

  const openVendorColumnFilter = (column: string, event: React.MouseEvent<HTMLButtonElement>) => {
    const rect = event.currentTarget.getBoundingClientRect();
    setVendorFilterPosition({ top: rect.bottom + 4, left: rect.left, width: Math.max(300, Math.min(380, rect.width + 180)) });
    setOpenVendorFilter(column);
  };

  const toggleVendorFilterValue = (column: string, key: string) => {
    setVendorColumnFilters((previous) => {
      const allKeys = (vendorFilterOptions[column] || []).map((option) => option.key);
      const current = previous[column];
      if (current === undefined) return { ...previous, [column]: allKeys.filter((value) => value !== key) };
      const next = current.includes(key) ? current.filter((value) => value !== key) : [...current, key];
      if (next.length === allKeys.length) {
        const copy = { ...previous };
        delete copy[column];
        return copy;
      }
      return { ...previous, [column]: next };
    });
  };

  const selectAllVendorFilterValues = (column: string) => {
    setVendorColumnFilters((previous) => {
      const copy = { ...previous };
      delete copy[column];
      return copy;
    });
  };

  const dateSummary = useMemo(() => {
    const map: Record<string, any> = {};

    data.forEach((r) => {
      const dt = r['Delivery Date'] || r['Booking Date'] || 'N/A';
      const key = reportDateKey(dt);
      const formatted = formatFullDisplayDate(dt);

      if (!map[key]) {
        map[key] = {
          date: formatted,
          rawDate: key,
          totalOrders: 0,
          delivered: 0,
          cancelled: 0,
          sellingPrice: 0,
          vendorPrice: 0,
          rfComm: 0,
        };
      }

      map[key].totalOrders += 1;
      if (r['Final Status'] === 'Delivered') map[key].delivered += 1;
      if (r['Final Status'] === 'Cancelled') map[key].cancelled += 1;
      map[key].sellingPrice += Number(r['Final Selling Price'] || 0) || 0;
      map[key].vendorPrice += Number(r['Final Vendor Price'] || 0) || 0;
      map[key].rfComm += Number(r['Final RF Commission'] || 0) || 0;
    });

    return Object.values(map).sort((a: any, b: any) => {
      if (a.rawDate === 'UNKNOWN') return 1;
      if (b.rawDate === 'UNKNOWN') return -1;
      return a.rawDate.localeCompare(b.rawDate);
    });
  }, [data]);

  // Vendor Date Wise dashboard MUST use the exact same data engine/order as
  // the Vendor Date Wise Excel export. This keeps:
  // Row Labels | Name | STN Code | 1 August 2026 | 2 August 2026 | ...
  // identical in Dashboard and Excel.
  const vendorDateWiseSummary = useMemo(() => {
    return generateVendorDateWiseData(data, outletsMasterInfo, currentMonthRecords);
  }, [data, outletsMasterInfo]);

  // Vendor Date Wise total row: sum the delivered-order counts for every
  // outlet for each date, matching the Excel-aligned date-wise report.
  // The total row is always shown at the top, independent of search filtering.
  const vendorDateWiseTotals = useMemo(() => {
    const totals: Record<string, number> = {};

    vendorDateWiseSummary.dateKeys.forEach((dateKey: string) => {
      totals[dateKey] = vendorDateWiseSummary.rows.reduce((sum: number, row: any) => {
        const value = Number(row[dateKey] ?? 0);
        return sum + (Number.isFinite(value) ? value : 0);
      }, 0);
    });

    return totals;
  }, [vendorDateWiseSummary]);

  // --- Outlet-wise Feedback / Complaint Report ---
  //
  // PERFORMANCE FIX:
  // Feedback Report can contain 1,000+ historical outlets and is not needed
  // while the user is viewing Master/Station/Vendor/etc. reports.
  // Calculate it only when the Feedback Report is actually selected.
  // This also prevents Old Ratings upload from triggering a heavy calculation
  // during every processing/render cycle.
  const feedbackReportRows = useMemo(() => {
    if (selectedReport !== 'FEEDBACK_REPORT') return [];

    return buildFeedbackReport(
      feedbackRawData,
      irctcRawData,
      data,
      outletsMasterInfo,
      oldRatingsRawData
    );
  }, [
    selectedReport,
    feedbackRawData,
    irctcRawData,
    outletsMasterInfo,
    oldRatingsRawData,
    data,
  ]);

  // Vendor RDS screen must use the EXACT same 41-column aggregation engine
  // as the Vendor RDS Excel export. This prevents the webpage from showing
  // the old 9-column Vendor Summary while Excel contains 41 columns.
  const vendorRdsRows = useMemo(() => {
    if (selectedReport !== 'VENDOR_RDS') return [];

    return generateVendorRdsData(
      data,
      penaltySummary,
      currentMonthRecords,
      outletsMasterInfo
    );
  }, [
    selectedReport,
    data,
    penaltySummary,
    currentMonthRecords,
    outletsMasterInfo,
  ]);

  // --- Exports ---
  const exportMasterExcel = () => {
    if (!data.length) return alert('No Data available!');
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Master Data');
    XLSX.writeFile(wb, `RELFOOD_MASTER_DATA_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportCurrentExcel = () => {
    switch (selectedReport) {
      case 'MASTER':
        exportMasterExcel();
        break;
      case 'MAIN_REPORT':
        generateMainReportWorkbook(data, feedbackRawData);
        break;
      case 'VENDOR_RDS':
        generateVendorRDSWorkbook(data, penaltySummary, currentMonthRecords, outletsMasterInfo);
        break;
      case 'STATION_REPORT':
        generateStationReportWorkbook(data, outletsMasterInfo, irctcRawData);
        break;
      case 'VENDOR_REPORT':
        generateVendorReportWorkbook(data, outletsMasterInfo, penaltySummary, currentMonthRecords);
        break;
      case 'DATE_WISE':
        generateDateWiseReportWorkbook(data);
        break;
      case 'VENDOR_DATE_WISE':
        generateVendorDateWiseReportWorkbook(data, outletsMasterInfo, currentMonthRecords);
        break;
      case 'LAST_DAY_STATION':
        generateLastDayStationReportWorkbook(data, outletsMasterInfo, irctcRawData);
        break;
      case 'OUTLETS_MASTER': {
        const ws = XLSX.utils.json_to_sheet(Object.values(outletsMasterInfo));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Outlets Master');
        XLSX.writeFile(wb, `OUTLETS_MASTER_${new Date().toISOString().slice(0, 10)}.xlsx`);
        break;
      }
      case 'FEEDBACK_REPORT':
        generateFeedbackReportWorkbook(feedbackReportRows);
        break;
      case 'PENALTIES': {
        const ws = XLSX.utils.json_to_sheet(penaltyRawRecords);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Penalties');
        XLSX.writeFile(wb, `PENALTY_REPORT_${new Date().toISOString().slice(0, 10)}.xlsx`);
        break;
      }
    }
  };

  const exportCurrentPDF = () => {
    if (!data.length && Object.keys(outletsMasterInfo).length === 0) {
      return alert('No Data available to generate PDF!');
    }

    // MAIN REPORT PDF must export the SAME date-matrix report that is visible
    // in the dashboard. It must never fall through to Master Data rows.
    if (selectedReport === 'MAIN_REPORT') {
      const doc = new jsPDF('landscape', 'pt', 'a4');
      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();
      const marginX = 18;
      const tableWidth = pageWidth - marginX * 2;

      doc.setFillColor(15, 23, 42);
      doc.rect(0, 0, pageWidth, 38, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(12);
      doc.text('RELFOOD ENTERPRISE PORTAL - MAIN REPORT', marginX, 24);

      const groups: { title: string; count: number; color: [number, number, number] }[] = [
        { title: 'ORDERS', count: 5, color: [114, 185, 210] },
        { title: 'MEALS', count: 5, color: [181, 211, 123] },
        { title: 'VALUE', count: 3, color: [243, 179, 126] },
        { title: 'PREPAID', count: 4, color: [141, 197, 222] },
        { title: 'DISCOUNT', count: 4, color: [217, 147, 151] },
        { title: 'REVENUE', count: 4, color: [153, 153, 153] },
        { title: 'Complaints', count: 4, color: [91, 145, 200] },
        { title: 'Feedback', count: 4, color: [135, 195, 79] },
        { title: 'IRCTC Undelivered', count: 4, color: [85, 85, 85] },
        { title: 'Outlets', count: 1, color: [242, 169, 0] },
      ];
      const subs = [
        ['FTD','MTD','LMTD','ASP','Del%'],
        ['FTD','MTD','LMTD','ASP','MPO'],
        ['FTD','MTD','LMTD'],
        ['FTD','MTD','LMTD','%'],
        ['FTD','MTD','LMTD','%'],
        ['FTD','MTD','LMTD','%'],
        ['FTD','MTD','LMTD','%'],
        ['FTD','MTD','LMTD','%'],
        ['FTD','MTD','LMTD','%'],
        [''],
      ];
      const rowsFor = (ftd: any, mtd: any, outletCount = 0) => {
        const num = (v: any) => Number(v || 0);
        const pct0 = (a: any,b: any,d=0) => b > 0 ? `${((num(a)/num(b))*100).toFixed(d)}%` : `0.${'0'.repeat(d)}%`;
        const asp = num(ftd.orders) ? Math.round(num(ftd.value)/num(ftd.orders)) : 0;
        const mealAsp = num(ftd.meals) ? Math.round(num(ftd.value)/num(ftd.meals)) : 0;
        const mpo = num(ftd.orders) ? (num(ftd.meals)/num(ftd.orders)).toFixed(2) : '0.00';
        return [
          ftd.orders, mtd.orders, 0, asp, pct0(ftd.deliveredOrders, ftd.orders, 0),
          ftd.meals, mtd.meals, 0, mealAsp, mpo,
          Math.round(num(ftd.value)), Math.round(num(mtd.value)), 0,
          Math.round(num(ftd.prepaidValue)), Math.round(num(mtd.prepaidValue)), 0, pct0(ftd.prepaidValue, ftd.value, 2),
          Math.round(num(ftd.discount)), Math.round(num(mtd.discount)), 0, pct0(ftd.discount, ftd.value, 2),
          Math.round(num(ftd.revenue)), Math.round(num(mtd.revenue)), 0, pct0(ftd.revenue, ftd.value, 1),
          ftd.complaints, mtd.complaints, 0, pct0(ftd.complaints, ftd.deliveredOrders, 2),
          ftd.feedback, mtd.feedback, 0, pct0(ftd.feedback, ftd.deliveredOrders, 2),
          ftd.undelivered, mtd.undelivered, 0, pct0(ftd.undelivered, ftd.orders, 2),
          outletCount,
        ];
      };

      const sourceNames = ['RELFood_IRCTC','RELFood_WEBSITE','REL_Food_App','MakeMyTrip'];
      let firstBlock = true;
      mainReportBlocks.forEach((blk: any) => {
        if (!firstBlock) doc.addPage('a4', 'landscape');
        firstBlock = false;
        const dateY = 52;
        doc.setFillColor(255, 0, 0);
        doc.rect(marginX, dateY, tableWidth, 22, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFont('helvetica', 'bold');
        doc.setFontSize(11);
        doc.text(blk.dateLabel, pageWidth / 2, dateY + 15, { align: 'center' });

        const head1 = ['Source', ...groups.map(g => g.title)];
        const head2 = ['', ...groups.flatMap((g,i) => subs[i])];
        const body = [
          ['Total', ...rowsFor(blk.dayTotal, blk.mtdTotal, Number(blk.outletsCount || 0))],
          ...sourceNames.map((src) => [''+src, ...rowsFor(blk.dayStats?.[src] || {}, blk.mtdBySource?.[src] || {}, 0)]),
        ];
        const bodyRows = body.map(r => r.map(v => String(v)));
        const colStyles: Record<number, any> = { 0: { cellWidth: 70, halign: 'left' } };
        for (let i=1;i<39;i++) colStyles[i] = { cellWidth: (tableWidth-70)/38, halign: 'center' };

        autoTable(doc, {
          head: [head1, head2],
          body: bodyRows,
          startY: dateY + 25,
          margin: { left: marginX, right: marginX },
          tableWidth,
          theme: 'grid',
          styles: { fontSize: 5.5, cellPadding: 1.2, lineColor: [34,34,34], lineWidth: 0.35, textColor: [0,0,0], overflow: 'hidden', halign: 'center', valign: 'middle' },
          headStyles: { fontSize: 5.5, fontStyle: 'bold', textColor: [0,0,0], fillColor: [220,231,236], lineColor: [34,34,34], lineWidth: 0.35 },
          columnStyles: colStyles,
          didParseCell: (hook) => {
            if (hook.section === 'head' && hook.row.index === 0) {
              if (hook.column.index === 0) hook.cell.styles.fillColor = [0,0,0], hook.cell.styles.textColor = [255,255,255];
              else {
                let cursor = 1;
                for (const g of groups) {
                  if (hook.column.index >= cursor && hook.column.index < cursor + g.count) {
                    hook.cell.styles.fillColor = g.color;
                    hook.cell.styles.textColor = g.title === 'IRCTC Undelivered' ? [255,255,255] : [0,0,0];
                    break;
                  }
                  cursor += g.count;
                }
              }
            }
            if (hook.section === 'head' && hook.row.index === 1) {
              const idx = hook.column.index;
              if (idx === 0) hook.cell.styles.fillColor = [0,0,0], hook.cell.styles.textColor = [255,255,255];
              else if (idx >= 1 && idx <= 5) hook.cell.styles.fillColor = [185,220,232];
              else if (idx <= 10) hook.cell.styles.fillColor = [210,231,191];
              else if (idx <= 13) hook.cell.styles.fillColor = [248,216,192];
              else if (idx <= 17) hook.cell.styles.fillColor = [196,224,236];
              else if (idx <= 21) hook.cell.styles.fillColor = [239,199,202];
              else if (idx <= 25) hook.cell.styles.fillColor = [207,207,207];
              else if (idx <= 29) hook.cell.styles.fillColor = [196,217,239];
              else if (idx <= 33) hook.cell.styles.fillColor = [212,232,197];
              else if (idx <= 37) hook.cell.styles.fillColor = [102,102,102], hook.cell.styles.textColor = [255,255,255];
              else hook.cell.styles.fillColor = [242,255,0];
            }
            if (hook.section === 'body') {
              const idx = hook.column.index;
              if (idx === 0) hook.cell.styles.fillColor = hook.row.index === 0 ? [0,0,0] : [239,0,0], hook.cell.styles.textColor = [255,255,255], hook.cell.styles.fontStyle = 'bold';
              else if (idx >= 1 && idx <= 5) hook.cell.styles.fillColor = [204,229,238];
              else if (idx <= 10) hook.cell.styles.fillColor = [217,232,200];
              else if (idx <= 13) hook.cell.styles.fillColor = [248,223,202];
              else if (idx <= 17) hook.cell.styles.fillColor = [214,234,242];
              else if (idx <= 21) hook.cell.styles.fillColor = [240,214,217];
              else if (idx <= 25) hook.cell.styles.fillColor = [208,208,208];
              else if (idx <= 29) hook.cell.styles.fillColor = [200,221,240];
              else if (idx <= 33) hook.cell.styles.fillColor = [217,234,203];
              else if (idx <= 37) hook.cell.styles.fillColor = [102,102,102], hook.cell.styles.textColor = [255,255,255];
              else hook.cell.styles.fillColor = [255,242,0], hook.cell.styles.fontStyle = 'bold';
            }
          },
        });
      });
      doc.save(`RELFOOD_MAIN_REPORT_${new Date().toISOString().slice(0,10)}.pdf`);
      return;
    }

    const doc = new jsPDF('landscape', 'pt', 'a4');
    const today = new Date().toLocaleDateString('en-IN');

    doc.setFillColor(15, 23, 42);
    doc.rect(0, 0, doc.internal.pageSize.getWidth(), 50, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text(`RELFOOD RDS PORTAL - ${selectedReport.replace(/_/g, ' ')}`, 30, 30);
    doc.setFontSize(9);
    doc.setFont('helvetica', 'normal');
    doc.text(`Generated: ${today} | Total Orders: ${data.length}`, doc.internal.pageSize.getWidth() - 200, 30);

    let head: string[][] = [];
    let body: any[][] = [];

    if (selectedReport === 'MASTER') {
      head = [['Order ID', 'Outlet ID', 'Vendor', 'Station', 'State', 'Status', 'Vendor ₹', 'Base ₹', 'GST ₹', 'RF Comm ₹', 'Selling ₹', 'Margin%']];
      body = data.map((r) => [
        r['IRCTC Order ID'], r['Outlet ID'], String(r['Vendor Name']).substring(0, 18), r['Station Code'], r['State'], r['Final Status'],
        `₹${r['Final Vendor Price']}`, `₹${r['Final Base Price']}`, `₹${r['Final GST']}`, `₹${r['Final RF Commission']}`, `₹${r['Final Selling Price']}`, `${r['Margin %']}%`,
      ]);
    } else if (selectedReport === 'STATION_REPORT' || selectedReport === 'LAST_DAY_STATION') {
      head = [['Station Code','Rank','Delivery Date','Station Rank','Station Name','Vendor Price','Final Base Price','Final Total Commission','Final IRCTC Comm','Final RF Commission','Final GST','Final Discount','Final Vendor Discount','Final RF Discount','Delivery Charges','Final Selling Price','Final Order Total','Discounted Base Price','PPD','COD','Meals','Check','Count of Delivered Orders','Not Delivered Order','Not Delivered %','PPD % of Final Selling Price','Feedback Good','Feedback Bad','Count of Delivered Outlets','Total Station Vendors']];
      const stationRows: any[] = selectedReport === 'LAST_DAY_STATION' ? lastDayStationSummary : stationSummary;
      body = stationRows.map((s: any) => [s['Station Code'],s['Rank'],s['Delivery Date'] || '',s['Station Rank'] ?? '',s['Station Name'] || '-',`₹${Number(s['Vendor Price'] || 0).toFixed(2)}`,`₹${Number(s['Final Base Price'] || 0).toFixed(2)}`,`₹${Number(s['Final Total Commission'] || 0).toFixed(2)}`,`₹${Number(s['Final IRCTC Comm'] || 0).toFixed(2)}`,`₹${Number(s['Final RF Commission'] || 0).toFixed(2)}`,`₹${Number(s['Final GST'] || 0).toFixed(2)}`,`₹${Number(s['Final Discount'] || 0).toFixed(2)}`,`₹${Number(s['Final Vendor Discount'] || 0).toFixed(2)}`,`₹${Number(s['Final RF Discount'] || 0).toFixed(2)}`,`₹${Number(s['Delivery Charges'] || 0).toFixed(2)}`,`₹${Number(s['Final Selling Price'] || 0).toFixed(2)}`,`₹${Number(s['Final Order Total'] || 0).toFixed(2)}`,`₹${Number(s['Discounted Base Price'] || 0).toFixed(2)}`,`₹${Number(s['PPD'] || 0).toFixed(2)}`,`₹${Number(s['COD'] || 0).toFixed(2)}`,`₹${Number(s['Meals'] || 0).toFixed(2)}`,s['Check'] || '-',s['Count of Delivered Orders'] || 0,s['Not Delivered Order'] || 0,s['Not Delivered %'] || '0.00%',s['PPD % of Final Selling Price'] || '0.00%',s['Feedback Good'] || 0,s['Feedback Bad'] || 0,s['Count of Delivered Outlets'] || 0,s['Total Station Vendors'] || 0]);
    } else if (selectedReport === 'VENDOR_REPORT') {
      head = [VENDOR_REPORT_COLUMNS.map((column) => column === 'Rank' ? 'Station Rank' : column === 'Aggregator Outlet ID' ? 'Outlet ID' : column)];
      body = vendorSummary.map((v) => VENDOR_REPORT_COLUMNS.map((column) => {
        const value = v[column];
        if (typeof value === 'number') return `₹${Number(value).toFixed(2)}`;
        return value ?? '-';
      }));
    } else if (selectedReport === 'VENDOR_RDS') {
      head = [VENDOR_RDS_COLUMNS.map((column) => column === 'Outlet ID' ? 'Outlet ID' : column)];
      body = vendorRdsRows.map((r: any) => VENDOR_RDS_COLUMNS.map((column) => {
        const value = r[column];
        if (typeof value === 'number') return Number(value).toFixed(2);
        return value ?? '-';
      }));
    } else if (selectedReport === 'DATE_WISE' || selectedReport === 'VENDOR_DATE_WISE') {
      head = [['Date', 'Total Orders', 'Delivered', 'Cancelled', 'Selling Amount', 'Vendor Price', 'RF Commission']];
      body = dateSummary.map((d) => [d.date,d.totalOrders,d.delivered,d.cancelled,`₹${d.sellingPrice.toFixed(2)}`,`₹${d.vendorPrice.toFixed(2)}`,`₹${d.rfComm.toFixed(2)}`]);
    } else if (selectedReport === 'OUTLETS_MASTER') {
      head = [['Outlet ID', 'Outlet Name', 'Station', 'State', 'GST Number', 'IRCTC Status']];
      body = Object.values(outletsMasterInfo).map((o) => [o.outletId,o.outletName||'-',o.station||'-',o.state||'-',o.gst||'-',o.irctcStatus||'-']);
    } else if (selectedReport === 'FEEDBACK_REPORT') {
      head = [['Outlet ID','Outlet Name','Station Code','Old Count','Old Ratings','Old Sum','Complaint','Feedback','Current Count','Current Rating','Current Sum','Total Count','Total Rating Sum','Till Date Ratings']];
      body = feedbackReportRows.map((r) => [r['Outlet Id'],String(r['Outlet Name']||'-').substring(0,28),r['Station Code']||'-',r['Old Count'],r['Old Ratings'],r['Old Sum'],r.Complaint,r.Feedback,r['Current Count'],r['Current Rating'],r['Current Sum'],r['Total Count'],r['Total Rating Sum'],r['Total Rating']]);
    } else if (selectedReport === 'PENALTIES') {
      head = [['Outlet ID','Order ID','Transaction Mode','Vendor Name','Date','Amount (₹)','Remarks']];
      body = penaltyRawRecords.map((p) => [p.outletId,p.orderId||'-',p.mode,p.vendorName||'-',p.date||'-',`₹${p.amount.toFixed(2)}`,p.remarks||'-']);
    }

    autoTable(doc, {
      head, body,
      startY: 65,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 4, textColor: [30,41,59] },
      headStyles: { fillColor: [37,99,235], textColor: [255,255,255], fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [248,250,252] },
      margin: { top: 60, bottom: 30, left: 20, right: 20 },
    });
    doc.save(`RELFOOD_${selectedReport}_${new Date().toISOString().slice(0,10)}.pdf`);
  };

  const renderMainReportRow = (label: string, ftd: MetricStats, mtd: MetricStats, isTotal: boolean = false) => {
    const orderAsp = ftd.orders > 0 ? Math.round(ftd.value / ftd.orders) : 0;
    const delPct = ftd.orders > 0 ? `${((ftd.deliveredOrders / ftd.orders) * 100).toFixed(0)}%` : '0%';
    const mealAsp = ftd.meals > 0 ? Math.round(ftd.value / ftd.meals) : 0;
    const mpo = ftd.orders > 0 ? (ftd.meals / ftd.orders).toFixed(2) : '0.00';
    const prepaidPct = ftd.value > 0 ? `${((ftd.prepaidValue / ftd.value) * 100).toFixed(2)}%` : '0.00%';
    const discountPct = ftd.value > 0 ? `${((ftd.discount / ftd.value) * 100).toFixed(2)}%` : '0.00%';
    const revenuePct = ftd.value > 0 ? `${((ftd.revenue / ftd.value) * 100).toFixed(1)}%` : '0.0%';
    const complaintPct = ftd.deliveredOrders > 0 ? `${((ftd.complaints / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
    const feedbackPct = ftd.deliveredOrders > 0 ? `${((ftd.feedback / ftd.deliveredOrders) * 100).toFixed(2)}%` : '0.00%';
    const undeliveredPct = ftd.orders > 0 ? `${((ftd.undelivered / ftd.orders) * 100).toFixed(2)}%` : '0.00%';

    return (
      <tr className={`portal-data-row border-b border-gray-300 ${isTotal ? 'font-bold bg-white text-black' : 'bg-white text-gray-800'}`}>
        <td className={`p-1.5 border border-gray-400 text-[11px] text-center whitespace-nowrap min-w-[130px] sticky left-0 z-10 ${isTotal ? 'bg-[#990000] text-white font-bold' : 'bg-red-600 text-white font-semibold'}`}>
          {label}
        </td>

        {/* ORDERS */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-medium min-w-[50px]">{ftd.orders}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-medium min-w-[50px]">{mtd.orders}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{orderAsp}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{delPct}</td>

        {/* MEALS */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-medium min-w-[50px]">{ftd.meals}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-medium min-w-[50px]">{mtd.meals}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{mealAsp}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{mpo}</td>

        {/* VALUE */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-semibold text-gray-900 min-w-[65px]">{Math.round(ftd.value)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-semibold text-gray-900 min-w-[65px]">{Math.round(mtd.value)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>

        {/* PREPAID */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[60px]">{Math.round(ftd.prepaidValue)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[60px]">{Math.round(mtd.prepaidValue)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[55px]">{prepaidPct}</td>

        {/* DISCOUNT */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[55px]">{Math.round(ftd.discount)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[55px]">{Math.round(mtd.discount)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[55px]">{discountPct}</td>

        {/* REVENUE */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-bold text-emerald-700 min-w-[60px]">{Math.round(ftd.revenue)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center font-bold text-emerald-700 min-w-[60px]">{Math.round(mtd.revenue)}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[50px]">{revenuePct}</td>

        {/* Complaints */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{ftd.complaints}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{mtd.complaints}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[50px]">{complaintPct}</td>

        {/* Feedback */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{ftd.feedback}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[45px]">{mtd.feedback}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[50px]">{feedbackPct}</td>

        {/* IRCTC Undelivered */}
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-rose-600 font-bold min-w-[45px]">{ftd.undelivered}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-rose-600 font-bold min-w-[45px]">{mtd.undelivered}</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center text-gray-400 min-w-[45px]">0</td>
        <td className="p-1.5 border border-gray-300 text-[11px] text-center min-w-[50px]">{undeliveredPct}</td>
      </tr>
    );
  };

  if (!isLoaded) {
    return (
      <div className="portal-clean min-h-screen bg-white flex items-center justify-center text-slate-600 text-sm">
        <div className="flex items-center gap-2">
          <div className="animate-spin w-4 h-4 border-2 border-indigo-500 border-t-transparent rounded-full" />
          Loading saved master records &amp; database...
        </div>
      </div>
    );
  }

  const penaltyOutletCount = Object.keys(penaltySummary).length;
  const outletsInfoCount = Object.keys(outletsMasterInfo).length;

  return (
    <div className={`portal-clean ${themeMode === 'night' ? 'portal-night' : 'portal-day'} min-h-screen p-3 md:p-6 font-sans`} onClick={handleTableRowClick}>
      {/* Top Header */}
      <header className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 pb-4 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-700 flex items-center justify-center text-white font-bold text-xl shadow-lg shadow-emerald-950">
            📊
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h1 className="text-lg md:text-xl font-black tracking-wide text-slate-900">
                RELFOOD ENTERPRISE PORTAL
              </h1>
              <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-emerald-500/20 text-emerald-400 border border-emerald-500/40">
                19 RULES ACTIVE
              </span>
              {outletsInfoCount > 0 && (
                <span className="px-2 py-0.5 rounded-md text-[10px] font-semibold bg-amber-500/20 text-amber-400 border border-amber-500/30">
                  {outletsInfoCount} OUTLETS
                </span>
              )}
              {penaltyOutletCount > 0 && (
                <span className="px-2 py-0.5 rounded-md text-[10px] font-semibold bg-rose-500/20 text-rose-400 border border-rose-500/30">
                  {penaltyOutletCount} PENALTIES
                </span>
              )}
            </div>
            <p className="text-xs text-slate-400 mt-0.5">
              Multi-Report Aggregation, Universal XLS/PDF Engine &amp; Real-time Calculations
            </p>
          </div>
        </div>

        {/* Global Action Controls */}
        <div className="flex items-center gap-2 flex-wrap w-full lg:w-auto">
          {data.length > 0 && (
            <>
              {/* Report Switcher Dropdown */}
              <div className="relative flex items-center">
                <span className="text-xs font-semibold text-slate-400 mr-2 hidden sm:inline">View:</span>
                <select
                  value={selectedReport}
                  onChange={(e) => setSelectedReport(e.target.value as ReportType)}
                  className="bg-slate-900 border border-indigo-500/50 text-indigo-300 font-semibold text-xs rounded-xl px-3 py-2 outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer"
                >
                  <option value="MAIN_REPORT">📊 Main Report (Date Matrix Layout)</option>
                  <option value="MASTER">📁 Master Data (All 19 Columns)</option>
                  <option value="VENDOR_RDS">📋 Vendor RDS Summary</option>
                  <option value="STATION_REPORT">🚉 Station Report</option>
                  <option value="VENDOR_REPORT">🏪 Vendor Report</option>
                  <option value="DATE_WISE">📈 Date Wise Summary</option>
                  <option value="VENDOR_DATE_WISE">📅 Vendor Date Wise</option>
                  <option value="LAST_DAY_STATION">🚉 Last Day Station</option>
                  <option value="OUTLETS_MASTER">🏢 Outlets Master (GST/State)</option>
                  <option value="PENALTIES">⚠️ Penalty &amp; Deductions</option>
                  <option value="FEEDBACK_REPORT">💬 Feedback Report</option>
                </select>
              </div>

              {/* Theme Toggle */}
              <button
                type="button"
                onClick={toggleTheme}
                className="theme-toggle px-3.5 py-2 rounded-xl text-xs font-bold transition flex items-center gap-1.5 border"
                title="Switch between Day and Night theme"
              >
                {themeMode === 'day' ? '🌙 Night' : '☀️ Day'}
              </button>

              {/* Download Buttons */}
              <button
                onClick={exportCurrentExcel}
                className="px-3.5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white transition shadow-lg shadow-emerald-950 flex items-center gap-1.5"
              >
                <span>📥</span> Excel (.xlsx)
              </button>

              <button
                onClick={exportCurrentPDF}
                className="px-3.5 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-xs font-bold text-white transition shadow-lg shadow-rose-950 flex items-center gap-1.5"
              >
                <span>📄</span> Download PDF
              </button>

              <button
                onClick={handleClearRecords}
                className="px-3 py-2 rounded-xl bg-slate-800 hover:bg-rose-900/50 border border-slate-700 text-xs font-semibold text-rose-400 transition"
                title="Clear Database"
              >
                🗑️
              </button>
            </>
          )}

          <button
            onClick={() => setIsModalOpen(true)}
            className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-bold text-white flex items-center gap-2 transition shadow-lg shadow-indigo-950"
          >
            <span>☁️</span> Upload 7 Reports
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="mt-4">
        {data.length === 0 ? (
          <div className="flex flex-col items-center justify-center p-16 rounded-2xl border border-dashed border-slate-800 bg-slate-900/30 text-center">
            <div className="text-5xl mb-4">📂</div>
            <h3 className="text-lg font-bold text-slate-300 mb-1">No Data Stored in Portal</h3>
            <p className="text-xs text-slate-500 max-w-md mb-6">
              Upload RF Report, IRCTC Report, Feedback, Old Feedback/Ratings, Penalty, Current Month &amp; Outlets Master. Calculations will be permanently preserved.
            </p>
            <button
              onClick={() => setIsModalOpen(true)}
              className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-bold text-white transition shadow-lg shadow-indigo-950"
            >
              Start Upload &amp; Processing
            </button>
          </div>
        ) : (
          <div>
            {/* Filter / Search Bar */}
            <div className="flex justify-between items-center mb-3 flex-wrap gap-2">
              <div className="flex items-center gap-3">
                <span className="text-xs font-bold px-3 py-1 bg-indigo-950/80 border border-indigo-700/60 text-indigo-300 rounded-lg">
                  Viewing: {selectedReport.replace(/_/g, ' ')}
                </span>
                <span className="text-xs text-slate-400">
                  Total Records: <strong className="text-emerald-400 font-bold">{data.length}</strong>
                </span>
              </div>
              <input
                type="text"
                placeholder="Search records..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="px-3.5 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 focus:outline-none focus:border-indigo-500 w-72"
              />
            </div>

            {/* MAIN REPORT: DATE-WISE EXCEL-STYLE MATRIX — FITS ENTIRE VIEW */}
            {selectedReport === 'MAIN_REPORT' && (
              <MainReportMatrix blocks={mainReportBlocks} searchTerm={searchTerm} />
            )}

            {/* ALL OTHER REPORT TABLES (WITH FULL HORIZONTAL SCROLL) */}
            {selectedReport !== 'MAIN_REPORT' && (
              <div 
                className="report-scroll w-full overflow-auto rounded-xl border shadow-sm max-h-[75vh]"
                style={{ 
                  WebkitOverflowScrolling: 'touch',
                  scrollbarWidth: 'auto',
                  scrollbarColor: '#cbd5e1 #f8fafc'
                }}
              >
                {/* 1. MASTER VIEW */}
                {selectedReport === 'MASTER' && (
                  <table className="portal-report-table portal-table-master w-full min-w-[1400px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Order ID</th>
                        <th className="p-3 font-semibold">Outlet ID</th>
                        <th className="p-3 font-semibold">Vendor Name</th>
                        <th className="p-3 font-semibold">Station</th>
                        <th className="p-3 font-semibold">State</th>
                        <th className="p-3 font-semibold">Status</th>
                        <th className="p-3 font-semibold text-right">Vendor Price</th>
                        <th className="p-3 font-semibold text-right">Base Price</th>
                        <th className="p-3 font-semibold text-right">GST (5%)</th>
                        <th className="p-3 font-semibold text-right">RF Comm</th>
                        <th className="p-3 font-semibold text-right">Selling Price</th>
                        <th className="p-3 font-semibold text-right">Margin %</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {data
                        .filter(
                          (r) =>
                            String(r['IRCTC Order ID'] || '').includes(searchTerm) ||
                            String(r['Outlet ID'] || '').includes(searchTerm) ||
                            String(r['Vendor Name'] || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                            String(r['Station Code'] || '').toLowerCase().includes(searchTerm.toLowerCase())
                        )
                        .slice(0, 100)
                        .map((row, i) => (
                          <tr key={i} className="portal-data-row hover:bg-slate-50">
                            <td className="p-3 font-medium text-white sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row['IRCTC Order ID']}</td>
                            <td className="p-3 text-slate-400">{row['Outlet ID']}</td>
                            <td className="p-3 font-medium">{row['Vendor Name']}</td>
                            <td className="p-3 text-cyan-300 font-mono">{row['Station Code']}</td>
                            <td className="p-3 text-amber-300/90">{row['State'] || '-'}</td>
                            <td className="p-3">
                              <span
                                className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  row['Final Status'] === 'Delivered'
                                    ? 'bg-emerald-950 text-emerald-400 border border-emerald-800'
                                    : row['Final Status'] === 'Cancelled'
                                    ? 'bg-rose-950 text-rose-400 border border-rose-800'
                                    : 'bg-amber-950 text-amber-400 border border-amber-800'
                                }`}
                              >
                                {row['Final Status']}
                              </span>
                            </td>
                            <td className="p-3 text-right">₹{row['Final Vendor Price']}</td>
                            <td className="p-3 text-right text-slate-200">₹{row['Final Base Price']}</td>
                            <td className="p-3 text-right text-cyan-400">₹{row['Final GST']}</td>
                            <td className="p-3 text-right text-emerald-400 font-bold">₹{row['Final RF Commission']}</td>
                            <td className="p-3 text-right text-amber-400 font-bold">₹{row['Final Selling Price']}</td>
                            <td className="p-3 text-right font-bold text-teal-400">{row['Margin %']}%</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 2. STATION / LAST DAY STATION VIEW */}
                {(selectedReport === 'STATION_REPORT' || selectedReport === 'LAST_DAY_STATION') && (() => {
                  const isLastDay = selectedReport === 'LAST_DAY_STATION';

                  // LAST_DAY_STATION columns are intentionally identical to the Excel generator.
                  const columns = isLastDay
                    ? [
                        'Station Code','Rank','Delivery Date','Station Rank','Station Name','Vendor Price','Final Base Price','Final Total Commission','Final IRCTC Comm','Final RF Commission','Final GST','Final Discount','Final Vendor Discount','Final RF Discount','Delivery Charges','Final Selling Price','Final Order Total','Discounted Base Price','PPD','COD','Meals','Check','Count of Delivered Orders','Not Delivered Order','Not Delivered %','PPD % of Final Selling Price','Feedback Good','Feedback Bad','Count of Delivered Outlets','Total Station Vendors'
                      ]
                    : [
                        'Station Code','Rank','Station Name','Station Rank','Vendor Price','Final Base Price','Final Total Commission','Final IRCTC Comm','Final RF Commission','Final GST','Final Discount','Final Vendor Discount','Final RF Discount','Delivery Charges','Final Selling Price','Final Order Total','Discounted Base Price','PPD','COD','Meals','Check','Count of Delivered Orders','Not Delivered Order','Not Delivered %','PPD % of Final Selling Price','Feedback Good','Feedback Bad','Count of Delivered Outlets','Total Station Vendors'
                      ];

                  const rows = (isLastDay ? lastDayStationSummary : stationSummary)
                    .filter((s: any) => String(s['Station Code'] || '').toLowerCase().includes(searchTerm.toLowerCase()));

                  return (
                    <table className="portal-report-table portal-table-station w-full min-w-[3000px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                      <thead className="sticky top-0 z-10 text-slate-700">
                        <tr>
                          {columns.map((col) => <th key={col} className="p-3 font-semibold text-center">{col}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((row: any, i: number) => (
                          <tr key={`${row['Station Code']}-${i}`} className="portal-data-row">
                            {columns.map((column: string, j: number) => {
                              const value = row[column];
                              return (
                                <td
                                  key={column}
                                  className={`p-3 ${j === 0 ? 'font-bold' : ''} ${j >= (isLastDay ? 4 : 3) && typeof value === 'number' ? 'text-right' : ''}`}
                                >
                                  {typeof value === 'number'
                                    ? value.toLocaleString('en-IN', { maximumFractionDigits: 2 })
                                    : (value ?? '-')}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  );
                })()}

                {/* 3A. VENDOR REPORT VIEW - STATION RANK + BUSINESS ORDER */}
                {selectedReport === 'VENDOR_REPORT' && (
                  <table className="portal-report-table portal-table-vendor w-full min-w-[2500px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 z-10 text-slate-700">
                      {/* TOTAL IS INTENTIONALLY ABOVE THE COLUMN HEADER */}
                      <tr className="bg-slate-300 font-extrabold text-slate-950">
                        {VENDOR_REPORT_COLUMNS.map((column: string, j: number) => {
                          const value = j === 0 ? 'TOTAL' : vendorReportTotals[column];
                          return (
                          <th key={`vendor-total-${j}`} className={`p-3 border-b border-slate-400 text-center ${j >= 5 && typeof value === 'number' ? 'text-right' : ''}`}>
                            {typeof value === 'number'
                              ? value.toLocaleString('en-IN', { maximumFractionDigits: 2 })
                              : (value ?? '')}
                          </th>
                          );
                        })}
                      </tr>
                      <tr className="bg-slate-900 text-slate-300">
                        {VENDOR_REPORT_COLUMNS.map((col: string) => {
                          const isFiltered = vendorColumnFilters[col] !== undefined;
                          const selectedCount = vendorColumnFilters[col]?.length ?? (vendorFilterOptions[col]?.length ?? 0);
                          return (
                            <th key={col} className="p-2 font-semibold text-center relative">
                              <div className="flex items-center justify-center gap-1">
                                <span>{col === 'Rank' ? 'Station Rank' : col}</span>
                                <button type="button" title={`Filter ${col}`} aria-label={`Filter ${col}`} onClick={(event) => openVendorColumnFilter(col, event)}
                                  className={`inline-flex h-6 w-6 shrink-0 items-center justify-center rounded ${isFiltered ? 'bg-amber-400 text-slate-900' : 'bg-white/10 text-slate-300'} hover:bg-white/25`}>
                                  <span className="text-[11px]">▼</span>
                                </button>
                              </div>
                              {isFiltered && <div className="mt-0.5 text-[9px] text-amber-300">{selectedCount} selected</div>}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredVendorSummary.map((row: any, i: number) => (
                          <tr key={`${row['Aggregator Outlet ID']}-${i}`} className="portal-data-row">
                            {VENDOR_REPORT_COLUMNS.map((column: string, j: number) => {
                              const value = row[column];
                              return (
                              <td key={j} className={`p-3 ${j < 5 ? 'font-medium' : ''} ${j >= 5 && typeof value === 'number' ? 'text-right' : ''}`}>
                                {typeof value === 'number'
                                  ? value.toLocaleString('en-IN', { maximumFractionDigits: 2 })
                                  : (value ?? '-')}
                              </td>
                              );
                            })}
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}
                {selectedReport === 'VENDOR_REPORT' && openVendorFilter && (
                  <VendorHeaderFilter
                    column={openVendorFilter}
                    options={vendorFilterOptions[openVendorFilter] || []}
                    selected={vendorColumnFilters[openVendorFilter]}
                    position={vendorFilterPosition}
                    onToggle={(key) => toggleVendorFilterValue(openVendorFilter, key)}
                    onSelectAll={() => selectAllVendorFilterValues(openVendorFilter)}
                    onClose={() => setOpenVendorFilter(null)}
                  />
                )}

                {/* 3B. VENDOR RDS VIEW - EXACT 41 COLUMNS AS EXCEL */}
                {selectedReport === 'VENDOR_RDS' && (
                  <table className="portal-report-table portal-table-rds w-full min-w-[5200px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 text-slate-400">
                      <tr>
                        {VENDOR_RDS_COLUMNS.map((column, index) => {
                          const sticky = index === 0
                            ? 'sticky left-0 z-40'
                            : index === 1
                              ? 'sticky left-[100px] z-40'
                              : index === 2
                                ? 'sticky left-[450px] z-40'
                                : '';
                          const width = index === 0
                            ? 'w-[100px] min-w-[100px] max-w-[100px]'
                            : index === 1
                              ? 'w-[350px] min-w-[350px] max-w-[350px]'
                              : index === 2
                                ? 'w-[140px] min-w-[140px] max-w-[140px]'
                                : '';
                          return (
                            <th
                              key={column}
                              className={`p-3 font-semibold border-b border-slate-800 bg-slate-900 ${sticky} ${width} ${index >= 6 ? 'text-right' : ''}`}
                            >
                              {column}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {vendorRdsRows
                        .filter((row: any) => {
                          const q = searchTerm.toLowerCase();
                          return (
                            String(row['Outlet ID'] ?? '').toLowerCase().includes(q) ||
                            String(row['Vendor Name'] ?? '').toLowerCase().includes(q) ||
                            String(row['Station Code'] ?? '').toLowerCase().includes(q) ||
                            String(row['GST Number'] ?? '').toLowerCase().includes(q) ||
                            String(row['Invoice Number'] ?? '').toLowerCase().includes(q)
                          );
                        })
                        .map((row: any) => (
                          <tr key={`${row['Outlet ID']}-${row['Invoice Number']}`} className="portal-data-row">
                            {VENDOR_RDS_COLUMNS.map((column, index) => {
                              const sticky = index === 0
                                ? 'sticky left-0 z-30'
                                : index === 1
                                  ? 'sticky left-[100px] z-30'
                                  : index === 2
                                    ? 'sticky left-[450px] z-30'
                                    : '';
                              const width = index === 0
                                ? 'w-[100px] min-w-[100px] max-w-[100px]'
                                : index === 1
                                  ? 'w-[350px] min-w-[350px] max-w-[350px]'
                                  : index === 2
                                    ? 'w-[140px] min-w-[140px] max-w-[140px]'
                                    : '';
                              const align = index >= 6 ? 'text-right' : index === 0 ? 'font-mono' : '';
                              const value = formatVendorRdsCell(column, row[column]);
                              const color = index === 0
                                ? 'font-bold text-indigo-300'
                                : index === 1
                                  ? 'font-medium text-white'
                                  : index === 2
                                    ? 'text-cyan-300 font-mono'
                                    : '';
                              return (
                                <td
                                  key={column}
                                  className={`p-3 bg-slate-900/95 ${sticky} ${width} ${align} ${color}`}
                                >
                                  {value}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 4. DATE WISE VIEW */}
                {selectedReport === 'DATE_WISE' && (
                  <table className="portal-report-table portal-table-date w-full min-w-[1100px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Date</th>
                        <th className="p-3 font-semibold text-center">Total Orders</th>
                        <th className="p-3 font-semibold text-center text-emerald-400">Delivered</th>
                        <th className="p-3 font-semibold text-center text-rose-400">Cancelled</th>
                        <th className="p-3 font-semibold text-right">Total Selling Amount</th>
                        <th className="p-3 font-semibold text-right">Vendor Price</th>
                        <th className="p-3 font-semibold text-right text-emerald-400">RF Commission</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {dateSummary
                        .filter((d) => d.date.toLowerCase().includes(searchTerm.toLowerCase()))
                        .map((row, i) => (
                          <tr key={i} className="portal-data-row hover:bg-slate-50">
                            <td className="p-3 font-bold text-white sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row.date}</td>
                            <td className="p-3 text-center">{row.totalOrders}</td>
                            <td className="p-3 text-center text-emerald-400 font-bold">{row.delivered}</td>
                            <td className="p-3 text-center text-rose-400">{row.cancelled}</td>
                            <td className="p-3 text-right font-bold text-amber-400">₹{row.sellingPrice.toFixed(2)}</td>
                            <td className="p-3 text-right">₹{row.vendorPrice.toFixed(2)}</td>
                            <td className="p-3 text-right font-bold text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 4B. VENDOR DATE WISE VIEW - EXACT EXCEL DATA LAYOUT */}
                {selectedReport === 'VENDOR_DATE_WISE' && (
                  <table className="portal-report-table portal-table-date w-full min-w-max text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-30 shadow-[2px_0_5px_rgba(0,0,0,0.5)] min-w-[110px] w-[110px]">
                          Row Labels
                        </th>
                        <th className="p-3 font-semibold sticky left-[110px] bg-slate-900 z-30 shadow-[2px_0_5px_rgba(0,0,0,0.35)] min-w-[310px] w-[310px]">
                          Name
                        </th>
                        <th className="p-3 font-semibold sticky left-[420px] bg-slate-900 z-30 shadow-[2px_0_5px_rgba(0,0,0,0.35)] min-w-[120px] w-[120px]">
                          STN Code
                        </th>
                        {vendorDateWiseSummary.dateColumns.map((dateLabel: string, index: number) => (
                          <th key={`${dateLabel}-${index}`} className="p-3 font-semibold text-center min-w-[110px]">
                            {dateLabel}
                          </th>
                        ))}
                        <th className="p-3 font-semibold text-center min-w-[150px]">Vendor Payment Type</th>
                        <th className="p-3 font-semibold text-center min-w-[140px]">Discount Applied</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {/* Total row — always visible at the top, just like the Excel report.
                          Deliberately dark, extra-bold, and two text sizes larger than normal rows. */}
                      <tr className="portal-data-row bg-slate-800/90 text-base font-extrabold text-slate-950">
                        <td className="p-3 text-base font-extrabold text-slate-950 sticky left-0 bg-slate-800 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.4)] min-w-[110px] w-[110px]">
                          Total
                        </td>
                        <td className="p-3 text-base font-extrabold text-slate-950 sticky left-[110px] bg-slate-800 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.25)] min-w-[310px] w-[310px]">
                          
                        </td>
                        <td className="p-3 text-base font-extrabold text-slate-950 sticky left-[420px] bg-slate-800 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.25)] min-w-[120px] w-[120px]">
                          
                        </td>
                        {vendorDateWiseSummary.dateKeys.map((dateKey: string, dateIndex: number) => {
                          const totalValue = Number(vendorDateWiseTotals[dateKey] ?? 0);
                          const totalIsZero = totalValue === 0;
                          return (
                            <td
                              key={`total-${dateKey}-${dateIndex}`}
                              className={`p-3 text-center min-w-[110px] text-base font-extrabold ${
                                totalIsZero ? 'text-red-700' : 'text-slate-950'
                              }`}
                            >
                              {totalValue}
                            </td>
                          );
                        })}
                        <td className="p-3 text-center min-w-[150px]"></td>
                        <td className="p-3 text-center min-w-[140px]"></td>
                      </tr>

                      {vendorDateWiseSummary.rows
                        .filter((row: any) => {
                          const q = searchTerm.trim().toLowerCase();
                          if (!q) return true;
                          return (
                            String(row['Row Labels'] ?? '').toLowerCase().includes(q) ||
                            String(row.Name ?? '').toLowerCase().includes(q) ||
                            String(row['STN Code'] ?? '').toLowerCase().includes(q)
                          );
                        })
                        .map((row: any, rowIndex: number) => (
                          <tr key={`${row['Row Labels']}-${rowIndex}`} className="portal-data-row hover:bg-slate-50">
                            <td className="p-3 font-bold text-indigo-300 font-mono sticky left-0 bg-slate-900/95 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.4)] min-w-[110px] w-[110px]">
                              {row['Row Labels']}
                            </td>
                            <td className="p-3 font-medium text-white sticky left-[110px] bg-slate-900/95 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.25)] min-w-[310px] w-[310px] max-w-[310px] truncate">
                              {row.Name || '-'}
                            </td>
                            <td className="p-3 text-cyan-300 font-mono sticky left-[420px] bg-slate-900/95 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.25)] min-w-[120px] w-[120px]">
                              {row['STN Code'] || '-'}
                            </td>
                            {vendorDateWiseSummary.dateKeys.map((dateKey: string, dateIndex: number) => {
                              const value = row[dateKey];
                              const isZero = value === 0 || value === '0' || value === '' || value === undefined || value === null;
                              return (
                                <td
                                  key={`${dateKey}-${dateIndex}`}
                                  className={`p-3 text-center min-w-[110px] ${
                                    isZero ? 'font-bold text-red-500' : 'font-medium'
                                  }`}
                                >
                                  {isZero ? 0 : value}
                                </td>
                              );
                            })}
                            <td className="p-3 text-center min-w-[150px]">{row['Vendor Payment Type'] || '-'}</td>
                            <td className="p-3 text-center min-w-[140px]">{row['Discount Applied'] || '-'}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 5. OUTLETS MASTER VIEW */}
                {selectedReport === 'OUTLETS_MASTER' && (
                  <table className="portal-report-table portal-table-outlets w-full min-w-[1000px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Outlet ID</th>
                        <th className="p-3 font-semibold">Outlet Name</th>
                        <th className="p-3 font-semibold">Station</th>
                        <th className="p-3 font-semibold">State</th>
                        <th className="p-3 font-semibold">GST Number</th>
                        <th className="p-3 font-semibold">IRCTC Status</th>
                        <th className="p-3 font-semibold">Station Rank</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {Object.values(outletsMasterInfo)
                        .filter((o) => o.outletId.includes(searchTerm) || o.outletName?.toLowerCase().includes(searchTerm.toLowerCase()))
                        .map((row, i) => (
                          <tr key={i} className="portal-data-row hover:bg-slate-50">
                            <td className="p-3 font-bold text-indigo-300 font-mono sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row.outletId}</td>
                            <td className="p-3 font-medium text-white">{row.outletName || '-'}</td>
                            <td className="p-3 text-cyan-300 font-mono">{row.station || '-'}</td>
                            <td className="p-3 text-amber-300">{row.state || '-'}</td>
                            <td className="p-3 font-mono text-slate-400">{row.gst || '-'}</td>
                            <td className="p-3">
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-800 text-slate-300">
                                {row.irctcStatus || 'Active'}
                              </span>
                            </td>
                            <td className="p-3 text-violet-300 font-bold">{row.stationRank || '-'}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 6. OUTLET-WISE FEEDBACK + RATING REPORT */}
                {selectedReport === 'FEEDBACK_REPORT' && (
                  <table className="portal-report-table portal-table-feedback w-full min-w-[1900px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 z-40 bg-slate-900 w-[100px] min-w-[100px] max-w-[100px] shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Outlet Id</th>
                        <th className="p-3 font-semibold sticky left-[100px] z-40 bg-slate-900 w-[300px] min-w-[300px] max-w-[300px]">Outlet Name</th>
                        <th className="p-3 font-semibold sticky left-[400px] z-40 bg-slate-900 w-[140px] min-w-[140px] max-w-[140px]">Station Code</th>
                        <th className="p-3 font-semibold text-center">Old Count</th>
                        <th className="p-3 font-semibold text-center">Old Ratings</th>
                        <th className="p-3 font-semibold text-center">Old Sum</th>
                        <th className="p-3 font-semibold text-center text-rose-400">Complaint</th>
                        <th className="p-3 font-semibold text-center text-emerald-400">Feedback</th>
                        <th className="p-3 font-semibold text-center">Current Count</th>
                        <th className="p-3 font-semibold text-center text-amber-400">Current Rating</th>
                        <th className="p-3 font-semibold text-center text-violet-400">Current Sum</th>
                        <th className="p-3 font-semibold text-center">Total Count</th>
                        <th className="p-3 font-semibold text-center">Total Rating Sum</th>
                        <th className="p-3 font-semibold text-center text-cyan-400">Till Date Ratings</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {feedbackReportRows
                        .filter((r) =>
                          r['Outlet Id'].toLowerCase().includes(searchTerm.toLowerCase()) ||
                          r['Outlet Name'].toLowerCase().includes(searchTerm.toLowerCase()) ||
                          r['Station Code'].toLowerCase().includes(searchTerm.toLowerCase())
                        )
                        .map((row) => (
                          <tr key={row['Outlet Id']} className="portal-data-row">
                            <td className="p-3 font-bold text-indigo-300 font-mono sticky left-0 z-30 bg-slate-900/95 w-[100px] min-w-[100px] max-w-[100px] shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row['Outlet Id']}</td>
                            <td className="p-3 font-medium text-white sticky left-[100px] z-30 bg-slate-900/95 w-[300px] min-w-[300px] max-w-[300px]">{row['Outlet Name'] || '-'}</td>
                            <td className="p-3 text-cyan-300 font-mono sticky left-[400px] z-30 bg-slate-900/95 w-[140px] min-w-[140px] max-w-[140px]">{row['Station Code'] || '-'}</td>
                            <td className="p-3 text-center">{row['Old Count']}</td>
                            <td className="p-3 text-center">{Number(row['Old Ratings'] || 0).toFixed(2)}</td>
                            <td className="p-3 text-center">{Number(row['Old Sum'] || 0).toFixed(2)}</td>
                            <td className="p-3 text-center font-bold text-rose-400">{row.Complaint}</td>
                            <td className="p-3 text-center font-bold text-emerald-400">{row.Feedback}</td>
                            <td className="p-3 text-center font-bold">{row['Current Count']}</td>
                            <td className="p-3 text-center font-bold text-amber-400">{Number(row['Current Rating'] || 0).toFixed(2)}</td>
                            <td className="p-3 text-center font-bold text-violet-400">{Number(row['Current Sum'] || 0).toFixed(2)}</td>
                            <td className="p-3 text-center font-bold">{row['Total Count']}</td>
                            <td className="p-3 text-center">{Number(row['Total Rating Sum'] || 0).toFixed(2)}</td>
                            <td className="p-3 text-center font-black text-cyan-400">{Number(row['Total Rating'] || 0).toFixed(2)}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 7. PENALTIES VIEW */}
                {selectedReport === 'PENALTIES' && (
                  <table className="portal-report-table portal-table-penalties w-full min-w-[1100px] text-left border-separate border-spacing-0 text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Outlet ID</th>
                        <th className="p-3 font-semibold">Order ID</th>
                        <th className="p-3 font-semibold">Transaction Mode</th>
                        <th className="p-3 font-semibold">Vendor Name</th>
                        <th className="p-3 font-semibold">Date</th>
                        <th className="p-3 font-semibold text-right">Amount (₹)</th>
                        <th className="p-3 font-semibold">Remarks</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {penaltyRawRecords
                        .filter((p) => p.outletId.includes(searchTerm) || p.orderId.includes(searchTerm))
                        .map((row, i) => (
                          <tr key={i} className="portal-data-row hover:bg-slate-50">
                            <td className="p-3 font-bold text-rose-300 font-mono sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row.outletId}</td>
                            <td className="p-3 text-white font-mono">{row.orderId || '-'}</td>
                            <td className="p-3">
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-rose-950 text-rose-400 border border-rose-800">
                                {row.mode}
                              </span>
                            </td>
                            <td className="p-3">{row.vendorName || '-'}</td>
                            <td className="p-3 text-slate-400">{row.date || '-'}</td>
                            <td className="p-3 text-right font-bold text-rose-400">₹{row.amount.toFixed(2)}</td>
                            <td className="p-3 text-slate-400 max-w-xs truncate">{row.remarks || '-'}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            <p className="text-[11px] text-slate-500 mt-2">
              * Laptop Touchpad par 2 ungliyon (Two fingers) ya horizontal scrollbar se right swipe karke poora data dekhein.
            </p>
          </div>
        )}
      </main>

      <style jsx global>{`
        /* Clean white portal theme + Excel-like horizontal browsing */
        .portal-clean {
          background: #ffffff !important;
          color: #0f172a !important;
        }

        .portal-clean [class*="bg-slate-950"],
        .portal-clean [class*="bg-slate-900"] {
          background-color: #ffffff !important;
        }

        .portal-clean [class*="bg-slate-800"] {
          background-color: #f8fafc !important;
        }

        .portal-clean [class*="border-slate-800"],
        .portal-clean [class*="border-slate-700"],
        .portal-clean [class*="border-slate-600"] {
          border-color: #dbe3ee !important;
        }

        .portal-clean [class*="text-slate-100"],
        .portal-clean [class*="text-slate-200"],
        .portal-clean [class*="text-slate-300"] {
          color: #334155 !important;
        }

        .portal-clean [class*="text-slate-400"] {
          color: #64748b !important;
        }

        .portal-clean [class*="text-slate-500"] {
          color: #94a3b8 !important;
        }

        .portal-report-table tbody td[class*="text-white"] {
          color: #0f172a !important;
        }

        .portal-clean input,
        .portal-clean select {
          background: #ffffff !important;
          color: #0f172a !important;
          border-color: #cbd5e1 !important;
        }

        .portal-clean main {
          min-width: 0;
        }

        .portal-report-table {
          --portal-c1: 140px;
          --portal-c2: 220px;
          --portal-c3: 160px;
          --portal-sticky-bg: #ffffff;
          border-color: #dbe3ee !important;
        }

        .portal-table-main { --portal-c1: 140px; --portal-c2: 60px; --portal-c3: 60px; }
        .portal-table-master { --portal-c1: 150px; --portal-c2: 120px; --portal-c3: 240px; }
        .portal-table-station { --portal-c1: 150px; --portal-c2: 180px; --portal-c3: 120px; }
        .portal-table-vendor { --portal-c1: 350px; --portal-c2: 120px; --portal-c3: 180px; }
        .portal-table-rds { --portal-c1: 100px; --portal-c2: 350px; --portal-c3: 140px; }
        .portal-table-date { --portal-c1: 230px; --portal-c2: 130px; --portal-c3: 130px; }
        .portal-table-outlets { --portal-c1: 130px; --portal-c2: 300px; --portal-c3: 160px; }
        .portal-table-feedback { --portal-c1: 100px; --portal-c2: 300px; --portal-c3: 140px; }
        .portal-table-penalties { --portal-c1: 130px; --portal-c2: 170px; --portal-c3: 180px; }

        .portal-report-table tbody tr:nth-child(odd) td {
          background: #ffffff !important;
        }

        .portal-report-table tbody tr:nth-child(even) td {
          background: #f8fafc !important;
        }

        .portal-report-table tbody tr:hover td {
          background: #eef6ff !important;
        }

        .portal-report-table tbody tr.portal-row-selected td {
          background: #bfdbfe !important;
          color: #0f172a !important;
          box-shadow: inset 0 1px 0 #93c5fd, inset 0 -1px 0 #93c5fd;
        }

        .portal-report-table thead tr:last-child > th:nth-child(1),
        .portal-report-table tbody tr > td:nth-child(1) {
          position: sticky !important;
          left: 0 !important;
          width: var(--portal-c1) !important;
          min-width: var(--portal-c1) !important;
          max-width: var(--portal-c1) !important;
          z-index: 30;
          box-sizing: border-box;
          background: var(--portal-sticky-bg) !important;
          box-shadow: 2px 0 5px rgba(15, 23, 42, 0.10);
        }

        .portal-report-table thead tr:last-child > th:nth-child(2),
        .portal-report-table tbody tr > td:nth-child(2) {
          position: sticky !important;
          left: var(--portal-c1) !important;
          width: var(--portal-c2) !important;
          min-width: var(--portal-c2) !important;
          max-width: var(--portal-c2) !important;
          z-index: 30;
          box-sizing: border-box;
          background: var(--portal-sticky-bg) !important;
        }

        .portal-report-table thead tr:last-child > th:nth-child(3),
        .portal-report-table tbody tr > td:nth-child(3) {
          position: sticky !important;
          left: calc(var(--portal-c1) + var(--portal-c2)) !important;
          width: var(--portal-c3) !important;
          min-width: var(--portal-c3) !important;
          max-width: var(--portal-c3) !important;
          z-index: 30;
          box-sizing: border-box;
          background: var(--portal-sticky-bg) !important;
          box-shadow: 2px 0 5px rgba(15, 23, 42, 0.10);
        }

        .portal-report-table tbody tr:nth-child(even) > td:nth-child(1),
        .portal-report-table tbody tr:nth-child(even) > td:nth-child(2),
        .portal-report-table tbody tr:nth-child(even) > td:nth-child(3) {
          background: #f8fafc !important;
        }

        .portal-report-table tbody tr.portal-row-selected > td:nth-child(1),
        .portal-report-table tbody tr.portal-row-selected > td:nth-child(2),
        .portal-report-table tbody tr.portal-row-selected > td:nth-child(3) {
          background: #bfdbfe !important;
        }

        .portal-report-table thead tr:last-child > th:nth-child(1),
        .portal-report-table thead tr:last-child > th:nth-child(2),
        .portal-report-table thead tr:last-child > th:nth-child(3) {
          z-index: 40 !important;
          background: #f1f5f9 !important;
          color: #334155 !important;
          box-shadow: 2px 0 5px rgba(15, 23, 42, 0.10);
        }

        .portal-report-table td,
        .portal-report-table th {
          border-color: #dbe3ee !important;
        }

        .portal-report-table tbody tr {
          cursor: pointer;
          transition: background-color 120ms ease;
        }

        /* Main matrix has grouped headers; keep the Source/FTD/MTD data columns fixed. */
        .portal-table-main thead tr:last-child > th:nth-child(1),
        .portal-table-main thead tr:last-child > th:nth-child(2),
        .portal-table-main thead tr:last-child > th:nth-child(3) {
          background: #f1f5f9 !important;
        }

        /* Always provide a real horizontal scrollbar for every report. */
        .portal-clean .overflow-x-auto {
          overflow-x: auto !important;
          overscroll-behavior-x: contain;
          scrollbar-color: #94a3b8 #f1f5f9;
          scrollbar-width: auto;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar {
          height: 12px;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar-track {
          background: #f1f5f9;
          border-radius: 999px;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar-thumb {
          background: #94a3b8;
          border-radius: 999px;
          border: 3px solid #f1f5f9;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar-thumb:hover {
          background: #64748b;
        }
      `}</style>

      {/* Final UI overrides: clean white theme + 3 frozen columns + row selection */}
      <style jsx global>{`
        /* =========================
           CLEAN WHITE THEME
           ========================= */
        .portal-clean,
        .portal-clean main,
        .portal-clean header,
        .portal-clean section {
          background: #ffffff !important;
          color: #0f172a !important;
        }

        /* Replace the dark report shell and dark utility panels. */
        .portal-clean [class*="bg-slate-950"],
        .portal-clean [class*="bg-slate-900"],
        .portal-clean [class*="bg-slate-800"],
        .portal-clean [class*="bg-slate-700"],
        .portal-clean [class*="bg-indigo-950"],
        .portal-clean [class*="bg-black"] {
          background-color: #ffffff !important;
        }

        .portal-clean [class*="text-slate-100"],
        .portal-clean [class*="text-slate-200"],
        .portal-clean [class*="text-slate-300"],
        .portal-clean [class*="text-slate-400"],
        .portal-clean [class*="text-slate-500"] {
          color: #475569 !important;
        }

        .portal-clean [class*="border-slate-800"],
        .portal-clean [class*="border-slate-700"],
        .portal-clean [class*="border-slate-600"] {
          border-color: #d7dee8 !important;
        }

        .portal-clean input,
        .portal-clean select {
          background: #ffffff !important;
          color: #0f172a !important;
          border-color: #cbd5e1 !important;
        }

        /* Modal also gets a clean white/light overlay. */
        .portal-clean [class*="bg-black/80"] {
          background-color: rgba(15, 23, 42, 0.22) !important;
        }

        /* =========================
           ALL REPORTS: HORIZONTAL SCROLL
           ========================= */
        .portal-clean .overflow-x-auto {
          overflow-x: auto !important;
          overflow-y: auto;
          -webkit-overflow-scrolling: touch;
          scrollbar-width: auto;
          scrollbar-color: #94a3b8 #eef2f7;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar {
          width: 10px;
          height: 12px;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar-track {
          background: #eef2f7;
        }

        .portal-clean .overflow-x-auto::-webkit-scrollbar-thumb {
          background: #94a3b8;
          border-radius: 8px;
          border: 3px solid #eef2f7;
        }

        /* =========================
           EVERY REPORT: FIRST 3 COLUMNS FROZEN
           ========================= */
        .portal-report-table {
          --freeze-c1: var(--portal-c1, 140px);
          --freeze-c2: var(--portal-c2, 220px);
          --freeze-c3: var(--portal-c3, 160px);
          --freeze-bg: #ffffff;
        }

        .portal-report-table thead tr:last-child > th:nth-child(1),
        .portal-report-table tbody > tr > td:nth-child(1) {
          position: sticky !important;
          left: 0 !important;
          width: var(--freeze-c1) !important;
          min-width: var(--freeze-c1) !important;
          max-width: var(--freeze-c1) !important;
          z-index: 31 !important;
          box-sizing: border-box !important;
          background: var(--freeze-bg) !important;
        }

        .portal-report-table thead tr:last-child > th:nth-child(2),
        .portal-report-table tbody > tr > td:nth-child(2) {
          position: sticky !important;
          left: var(--freeze-c1) !important;
          width: var(--freeze-c2) !important;
          min-width: var(--freeze-c2) !important;
          max-width: var(--freeze-c2) !important;
          z-index: 31 !important;
          box-sizing: border-box !important;
          background: var(--freeze-bg) !important;
        }

        .portal-report-table thead tr:last-child > th:nth-child(3),
        .portal-report-table tbody > tr > td:nth-child(3) {
          position: sticky !important;
          left: calc(var(--freeze-c1) + var(--freeze-c2)) !important;
          width: var(--freeze-c3) !important;
          min-width: var(--freeze-c3) !important;
          max-width: var(--freeze-c3) !important;
          z-index: 31 !important;
          box-sizing: border-box !important;
          background: var(--freeze-bg) !important;
          box-shadow: 3px 0 7px rgba(15, 23, 42, 0.12) !important;
        }

        .portal-report-table thead tr:last-child > th:nth-child(1),
        .portal-report-table thead tr:last-child > th:nth-child(2),
        .portal-report-table thead tr:last-child > th:nth-child(3) {
          z-index: 45 !important;
          background: #f1f5f9 !important;
          color: #334155 !important;
        }

        /* Main matrix has multiple header rows; freeze the first 3 visible
           metric columns in the bottom header row and all body cells. */
        .portal-table-main thead tr:last-child > th:nth-child(1),
        .portal-table-main thead tr:last-child > th:nth-child(2),
        .portal-table-main thead tr:last-child > th:nth-child(3) {
          z-index: 45 !important;
        }

        /* =========================
           EXCEL-LIKE ZEBRA ROWS + CLICK HIGHLIGHT
           ========================= */
        .portal-report-table tbody > tr {
          cursor: pointer !important;
          transition: background-color 120ms ease, box-shadow 120ms ease;
        }

        .portal-report-table tbody > tr:nth-child(odd) > td {
          background: #ffffff !important;
        }

        .portal-report-table tbody > tr:nth-child(even) > td {
          background: #f7f9fc !important;
        }

        .portal-report-table tbody > tr:hover > td {
          background: #eaf3ff !important;
        }

        /* Clicked row: this rule is intentionally LAST and !important so
           sticky cells and old Tailwind backgrounds cannot hide selection. */
        .portal-report-table tbody > tr.portal-row-selected > td,
        .portal-report-table tbody > tr.portal-row-selected > td:nth-child(1),
        .portal-report-table tbody > tr.portal-row-selected > td:nth-child(2),
        .portal-report-table tbody > tr.portal-row-selected > td:nth-child(3) {
          background: #bfdbfe !important;
          color: #0f172a !important;
          box-shadow: inset 0 2px 0 #60a5fa, inset 0 -2px 0 #60a5fa !important;
        }

        .portal-report-table tbody > tr.portal-row-selected {
          outline: 2px solid #3b82f6;
          outline-offset: -2px;
        }

        /* Keep report text readable on white cells. */
        .portal-report-table tbody td[class*="text-white"] {
          color: #0f172a !important;
        }

        .portal-report-table tbody td[class*="text-slate-300"],
        .portal-report-table tbody td[class*="text-slate-400"] {
          color: #475569 !important;
        }
      `}</style>

      {/* Final hard UI contract: scroll, 3 frozen columns, theme, selection */}
      <style jsx global>{`
        .portal-clean { min-width: 0 !important; }
        .report-scroll {
          display: block !important;
          width: 100% !important;
          max-width: 100% !important;
          overflow-x: scroll !important;
          overflow-y: auto !important;
          overscroll-behavior: contain;
          -webkit-overflow-scrolling: touch;
          scrollbar-width: auto !important;
        }
        .report-scroll::-webkit-scrollbar { width: 11px; height: 13px; }
        .report-scroll::-webkit-scrollbar-track { background: #e2e8f0; }
        .report-scroll::-webkit-scrollbar-thumb { background: #64748b; border-radius: 8px; border: 3px solid #e2e8f0; }
        .portal-report-table { width: max-content !important; table-layout: fixed !important; }
        .portal-report-table tbody tr { cursor: pointer !important; }

        .portal-day { background: #ffffff !important; color: #0f172a !important; }
        .portal-day .theme-toggle { background:#0f172a !important; color:#ffffff !important; border-color:#0f172a !important; }

        .portal-night { background:#07111f !important; color:#e5e7eb !important; }
        .portal-night header, .portal-night section, .portal-night main { background:#07111f !important; color:#e5e7eb !important; }
        .portal-night [class*="bg-white"] { background:#0f172a !important; }
        .portal-night [class*="bg-slate-50"] { background:#111827 !important; }
        .portal-night [class*="bg-slate-900"] { background:#0f172a !important; }
        .portal-night [class*="text-slate-900"] { color:#f8fafc !important; }
        .portal-night [class*="text-slate-800"] { color:#e5e7eb !important; }
        .portal-night [class*="text-slate-700"] { color:#cbd5e1 !important; }
        .portal-night [class*="text-slate-600"] { color:#94a3b8 !important; }
        .portal-night [class*="text-slate-500"], .portal-night [class*="text-slate-400"] { color:#94a3b8 !important; }
        .portal-night [class*="border-slate"], .portal-night [class*="border-gray"] { border-color:#334155 !important; }
        .portal-night .theme-toggle { background:#f8fafc !important; color:#0f172a !important; border-color:#e2e8f0 !important; }
        .portal-night .portal-report-table tbody tr:nth-child(odd) td { background:#0f172a !important; color:#e5e7eb !important; }
        .portal-night .portal-report-table tbody tr:nth-child(even) td { background:#111827 !important; color:#e5e7eb !important; }
        .portal-night .portal-report-table tbody tr:hover td { background:#172554 !important; }
        .portal-night .portal-report-table thead th { background:#1e293b !important; color:#e2e8f0 !important; }
        .portal-night .portal-report-table tbody tr.portal-row-selected td { background:#1d4ed8 !important; color:#ffffff !important; }
        .portal-night .portal-report-table tbody tr.portal-row-selected td:nth-child(1),
        .portal-night .portal-report-table tbody tr.portal-row-selected td:nth-child(2),
        .portal-night .portal-report-table tbody tr.portal-row-selected td:nth-child(3) { background:#1d4ed8 !important; color:#ffffff !important; }

        /* DAY sticky cells: all first 3 stay visible during horizontal scroll. */
        .portal-day .portal-report-table tbody tr:nth-child(odd) td:nth-child(-n+3) { background:#ffffff !important; }
        .portal-day .portal-report-table tbody tr:nth-child(even) td:nth-child(-n+3) { background:#f7f9fc !important; }
        .portal-day .portal-report-table tbody tr:hover td:nth-child(-n+3) { background:#eaf3ff !important; }
        .portal-day .portal-report-table tbody tr.portal-row-selected td:nth-child(-n+3) { background:#bfdbfe !important; }
        .portal-day .portal-report-table thead th:nth-child(-n+3) { background:#f1f5f9 !important; color:#334155 !important; }

        /* Keep sticky offsets tied to the actual first-three-column widths. */
        .portal-report-table thead th:nth-child(1), .portal-report-table tbody td:nth-child(1) { left:0 !important; position:sticky !important; z-index:31 !important; }
        .portal-report-table thead th:nth-child(2), .portal-report-table tbody td:nth-child(2) { left:var(--portal-c1) !important; position:sticky !important; z-index:31 !important; }
        .portal-report-table thead th:nth-child(3), .portal-report-table tbody td:nth-child(3) { left:calc(var(--portal-c1) + var(--portal-c2)) !important; position:sticky !important; z-index:31 !important; }
        .portal-report-table thead th:nth-child(-n+3) { z-index:45 !important; position:sticky !important; top:0 !important; }
      `}</style>

      {/* Upload Modal (7 Files) */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <div className="w-full max-w-xl rounded-2xl bg-white p-6 border border-slate-700 shadow-2xl text-slate-800">
            <h2 className="text-xl font-bold mb-2 flex items-center gap-2">
              <span>📊</span> Upload Reports (7 Files System)
            </h2>
            <p className="text-xs text-slate-400 mb-4">
              Files select karein. Sabhi calculated metrics aur summaries permanently IndexedDB me save ho jayengi.
            </p>

            <div className="space-y-3 max-h-[60vh] overflow-y-auto pr-1">
              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-emerald-400 block mb-1">
                  1. RF Report (.xls / .html / .csv) *
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx,.html"
                  onChange={(e) => setRfFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-emerald-600 file:text-white hover:file:bg-emerald-500 cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-emerald-400 block mb-1">
                  2. IRCTC Report (.csv / .xls / .xlsx) *
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={(e) => setIrctcFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-emerald-600 file:text-white hover:file:bg-emerald-500 cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-slate-300 block mb-1">
                  3. Feedback Report (Optional)
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={(e) => setFeedbackFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-slate-700 file:text-slate-200 hover:file:bg-slate-600 cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-violet-700/60">
                <label className="text-xs font-semibold text-violet-400 block mb-1">
                  4. Old Feedback / Old Ratings (Historical Outlet-wise Rating &amp; Count) (Optional)
                </label>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={(e) => setOldFeedbackFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-violet-700 file:text-white hover:file:bg-violet-600 cursor-pointer"
                />
                <p className="text-[10px] text-slate-500 mt-1">
                  Columns: Outlet Id, Outlet Name, Station Code, Old Count, Old Ratings, Old Sum
                </p>
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-rose-400 block mb-1">
                  5. Penalty &amp; Deduction Report (Optional)
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={(e) => setPenaltyFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-rose-700 file:text-white hover:file:bg-rose-600 cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-cyan-400 block mb-1">
                  6. Current Month Data (Previous Balance, Paid by Relfood) (Optional)
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={(e) => setCurrentMonthFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-cyan-700 file:text-white hover:file:bg-cyan-600 cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-amber-400 block mb-1">
                  7. Outlets Master (GST, State &amp; IRCTC Status) (Optional)
                </label>
                <input
                  type="file"
                  accept=".csv,.xls,.xlsx"
                  onChange={(e) => setOutletsFile(e.target.files?.[0] || null)}
                  className="w-full text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-amber-700 file:text-white hover:file:bg-amber-600 cursor-pointer"
                />
              </div>
            </div>

            {isProcessing && (
              <div className="mt-4 p-3 bg-indigo-950/60 border border-indigo-700/60 rounded-xl flex items-center gap-3">
                <div className="animate-spin w-4 h-4 border-2 border-indigo-400 border-t-transparent rounded-full" />
                <span className="text-xs font-medium text-indigo-300">{statusText}</span>
              </div>
            )}

            <div className="mt-6 flex justify-end gap-3">
              <button
                onClick={() => setIsModalOpen(false)}
                disabled={isProcessing}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-300 transition"
              >
                Cancel
              </button>
              <button
                onClick={handleProcessAndMerge}
                disabled={isProcessing}
                className="px-5 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 text-xs font-bold text-white transition flex items-center gap-2 shadow-lg shadow-indigo-950"
              >
                {isProcessing ? 'Processing...' : 'Merge & Process All'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
