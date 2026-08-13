'use client';

import React, { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { generateMainReportWorkbook } from '@/lib/mainReportGenerator';
import { generateVendorRDSWorkbook } from '@/lib/vendorRdsGenerator';
import { generateStationReportWorkbook } from '@/lib/stationReportGenerator';
import { generateVendorReportWorkbook } from '@/lib/vendorReportGenerator';
import { generateDateWiseReportWorkbook } from '@/lib/dateWiseReportGenerator';
import { generateVendorDateWiseReportWorkbook } from '@/lib/vendorDateWiseReportGenerator';
import { generateLastDayStationReportWorkbook } from '@/lib/lastDayStationReportGenerator';

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
  if (!val && val !== 0) return '';
  return String(val).trim().replace(/\.0$/, '');
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

// --- Date Fix: Excel Serial Number & Regular Dates ---
const formatFullDisplayDate = (dateVal: any): string => {
  if (!dateVal) return 'Unknown Date';
  
  const num = parseFloat(String(dateVal));
  let d: Date;

  // Check if it's an Excel serial date number like 46061.0001
  if (!isNaN(num) && num > 30000 && num < 70000) {
    const utcDays = Math.floor(num - 25569);
    const utcValue = utcDays * 86400;
    d = new Date(utcValue * 1000);
  } else {
    d = new Date(String(dateVal));
  }

  if (isNaN(d.getTime())) return String(dateVal);

  return d.toLocaleDateString('en-GB', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
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
  | 'PENALTIES';

export default function Page() {
  const [data, setData] = useState<any[]>([]);
  const [penaltySummary, setPenaltySummary] = useState<Record<string, number>>({});
  const [penaltyRawRecords, setPenaltyRawRecords] = useState<any[]>([]);
  const [currentMonthRecords, setCurrentMonthRecords] = useState<any[]>([]);
  const [outletsMasterInfo, setOutletsMasterInfo] = useState<Record<string, any>>({});
  // Raw IRCTC data is kept separately because Station Report feedback counts
  // must be calculated from the complete IRCTC report (Delivery Station + Feedback Type).
  const [irctcRawData, setIrctcRawData] = useState<any[]>([]);
  const [isLoaded, setIsLoaded] = useState<boolean>(false);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedReport, setSelectedReport] = useState<ReportType>('MAIN_REPORT');

  // Upload States
  const [rfFile, setRfFile] = useState<File | null>(null);
  const [irctcFile, setIrctcFile] = useState<File | null>(null);
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [penaltyFile, setPenaltyFile] = useState<File | null>(null);
  const [currentMonthFile, setCurrentMonthFile] = useState<File | null>(null);
  const [outletsFile, setOutletsFile] = useState<File | null>(null);

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');

  useEffect(() => {
    const fetchStoredData = async () => {
      try {
        const storedMaster = await loadFromDB('CURRENT_MASTER_DATA');
        const storedPenalty = await loadFromDB('OUTLET_PENALTY_DATA');
        const storedCurrentMonth = await loadFromDB('CURRENT_MONTH_DATA');
        const storedOutletsInfo = await loadFromDB('OUTLET_MASTER_INFO');
        const storedIrctcData = await loadFromDB('CURRENT_IRCTC_DATA');

        if (Array.isArray(storedMaster) && storedMaster.length > 0) setData(storedMaster);
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
        if (Array.isArray(storedIrctcData) && storedIrctcData.length > 0) {
          setIrctcRawData(storedIrctcData);
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
      setPenaltySummary({});
      setPenaltyRawRecords([]);
      setCurrentMonthRecords([]);
      setOutletsMasterInfo({});
      setIrctcRawData([]);
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
    if (!rfFile || !irctcFile) {
      alert('Kripya RF Report aur IRCTC Report zaroor upload karein!');
      return;
    }

    try {
      setIsProcessing(true);

      setStatusText('Reading RF Report...');
      const rfData = await parseAnyFile(rfFile);

      setStatusText('Reading IRCTC Report...');
      const irctcData = await parseAnyFile(irctcFile);

      let feedbackData: any[] = [];
      if (feedbackFile) {
        setStatusText('Reading Feedback Report...');
        feedbackData = await parseAnyFile(feedbackFile);
      }

      // IMPORTANT: Keep the COMPLETE IRCTC raw report. Station Report feedback
      // counts are based on IRCTC Delivery Station + Feedback Type, not on the
      // separately uploaded Feedback report and not only on matched master rows.
      await saveToDB('CURRENT_IRCTC_DATA', irctcData);
      setIrctcRawData(irctcData);

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
              station: String(row['Station'] || row['Station Name'] || '').trim(),
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

      const fbMap = new Map();
      feedbackData.forEach((row) => {
        const orderId = String(row['Order ID'] || row['Feedback Id'] || row['Order Id'] || '').trim().replace(/\.0$/, '');
        if (orderId) fbMap.set(orderId, row);
      });

      const masterRows = rfData.map((rf) => {
        const orderId = String(rf['IRCTC OrderId'] || rf['Order Id'] || '').trim().replace(/\.0$/, '');
        const irctc = irctcMap.get(orderId) || {};
        const fb = fbMap.get(orderId) || {};
        const outletId = cleanOutletId(rf['OutletId'] || rf['Outlet ID'] || irctc['Outlet Id'] || '');
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
          'Rating': fb['Rating'] || '',
          'Remarks': fb['Remarks'] || irctc['Comments'] || '',
        };
      });

      await saveToDB('CURRENT_MASTER_DATA', masterRows);
      setData(masterRows);

      setRfFile(null);
      setIrctcFile(null);
      setFeedbackFile(null);
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

    const dateMap: Record<string, any[]> = {};
    data.forEach((row) => {
      const rawDate = row['Delivery Date'] || row['Booking Date'] || 'Unknown Date';
      const dateKey = String(rawDate).split(' ')[0].split('T')[0];
      if (!dateMap[dateKey]) dateMap[dateKey] = [];
      dateMap[dateKey].push(row);
    });

    const sortedDates = Object.keys(dateMap).sort((a, b) => {
      const numA = parseFloat(a);
      const numB = parseFloat(b);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return new Date(a).getTime() - new Date(b).getTime();
    });

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
        const outletId = String(r['Outlet ID'] || '').trim();

        const sStat = dayStats[src] || dayStats['RELFood_IRCTC'];
        sStat.orders += 1;
        if (isDelivered) sStat.deliveredOrders += 1;
        if (isUndelivered) sStat.undelivered += 1;
        sStat.meals += mealCount;
        sStat.value += sellingPrice;
        sStat.prepaidValue += prepaid;
        sStat.discount += discount;
        sStat.revenue += rfComm;
        if (r['Rating'] && parseFloat(r['Rating']) > 0) sStat.feedback += 1;
        if (r['Remarks'] && String(r['Remarks']).toLowerCase().includes('complaint')) sStat.complaints += 1;
        if (outletId) sStat.outletsSet.add(outletId);
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
        dateLabel: formatFullDisplayDate(dateKey),
        rawDate: dateKey,
        dayTotal: { ...dayTotal },
        dayStats: JSON.parse(JSON.stringify(dayStats)),
        mtdTotal: { ...mtdGrandTotal },
        mtdBySource: JSON.parse(JSON.stringify(mtdBySource)),
        outletsCount: dayTotal.outletsSet.size,
      };
    });
  }, [data]);

  // --- Aggregate Views ---
  const stationSummary = useMemo(() => {
    // IMPORTANT: Feedback counts come ONLY from the raw IRCTC report:
    // Delivery Station + Feedback Type.
    // FEEDBACK => Feedback Good
    // COMPLAIN => Feedback Bad
    const feedbackMap: Record<string, { good: number; bad: number }> = {};

    (irctcRawData || []).forEach((r: any) => {
      const station = String(
        r['Delivery Station'] ?? r['DeliveryStation'] ?? ''
      ).trim().toUpperCase();
      const type = String(
        r['Feedback Type'] ?? r['FeedbackType'] ?? ''
      ).trim().toUpperCase();

      if (!station) return;
      if (!feedbackMap[station]) feedbackMap[station] = { good: 0, bad: 0 };

      if (type === 'FEEDBACK') feedbackMap[station].good += 1;
      if (type === 'COMPLAIN') feedbackMap[station].bad += 1;
    });

    const map: Record<string, any> = {};
    data.forEach((r) => {
      const stn = String(r['Station Code'] || 'UNKNOWN').trim().toUpperCase();
      if (!map[stn]) {
        map[stn] = {
          station: stn,
          state: r['State'] || '',
          totalOrders: 0,
          delivered: 0,
          cancelled: 0,
          sellingPrice: 0,
          vendorPrice: 0,
          rfComm: 0,
          gst: 0,
          feedbackGood: feedbackMap[stn]?.good || 0,
          feedbackBad: feedbackMap[stn]?.bad || 0,
        };
      }
      map[stn].totalOrders += 1;
      if (r['Final Status'] === 'Delivered') map[stn].delivered += 1;
      if (r['Final Status'] === 'Cancelled') map[stn].cancelled += 1;
      map[stn].sellingPrice += Number(r['Final Selling Price'] || 0);
      map[stn].vendorPrice += Number(r['Final Vendor Price'] || 0);
      map[stn].rfComm += Number(r['Final RF Commission'] || 0);
      map[stn].gst += Number(r['Final GST'] || 0);
    });
    return Object.values(map);
  }, [data, irctcRawData]);

  const vendorSummary = useMemo(() => {
    const map: Record<string, any> = {};
    data.forEach((r) => {
      const v = r['Vendor Name'] || 'UNKNOWN';
      if (!map[v]) {
        map[v] = {
          vendor: v,
          outletId: r['Outlet ID'],
          state: r['State'] || '',
          totalOrders: 0,
          delivered: 0,
          sellingPrice: 0,
          vendorPrice: 0,
          rfComm: 0,
          penalty: penaltySummary[r['Outlet ID']] || 0,
        };
      }
      map[v].totalOrders += 1;
      if (r['Final Status'] === 'Delivered') map[v].delivered += 1;
      map[v].sellingPrice += r['Final Selling Price'] || 0;
      map[v].vendorPrice += r['Final Vendor Price'] || 0;
      map[v].rfComm += r['Final RF Commission'] || 0;
    });
    return Object.values(map);
  }, [data, penaltySummary]);

  const dateSummary = useMemo(() => {
    const map: Record<string, any> = {};
    data.forEach((r) => {
      const dt = r['Delivery Date'] || r['Booking Date'] || 'N/A';
      const formatted = formatFullDisplayDate(dt);
      if (!map[formatted]) {
        map[formatted] = {
          date: formatted,
          totalOrders: 0,
          delivered: 0,
          cancelled: 0,
          sellingPrice: 0,
          vendorPrice: 0,
          rfComm: 0,
        };
      }
      map[formatted].totalOrders += 1;
      if (r['Final Status'] === 'Delivered') map[formatted].delivered += 1;
      if (r['Final Status'] === 'Cancelled') map[formatted].cancelled += 1;
      map[formatted].sellingPrice += r['Final Selling Price'] || 0;
      map[formatted].vendorPrice += r['Final Vendor Price'] || 0;
      map[formatted].rfComm += r['Final RF Commission'] || 0;
    });
    return Object.values(map);
  }, [data]);

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
        generateMainReportWorkbook(data);
        break;
      case 'VENDOR_RDS':
        generateVendorRDSWorkbook(data, penaltySummary, currentMonthRecords, outletsMasterInfo);
        break;
      case 'STATION_REPORT':
        generateStationReportWorkbook(data, outletsMasterInfo, irctcRawData);
        break;
      case 'VENDOR_REPORT':
        generateVendorReportWorkbook(data, outletsMasterInfo, penaltySummary);
        break;
      case 'DATE_WISE':
        generateDateWiseReportWorkbook(data);
        break;
      case 'VENDOR_DATE_WISE':
        generateVendorDateWiseReportWorkbook(data);
        break;
      case 'LAST_DAY_STATION':
        generateLastDayStationReportWorkbook(data, outletsMasterInfo);
        break;
      case 'OUTLETS_MASTER': {
        const ws = XLSX.utils.json_to_sheet(Object.values(outletsMasterInfo));
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Outlets Master');
        XLSX.writeFile(wb, `OUTLETS_MASTER_${new Date().toISOString().slice(0, 10)}.xlsx`);
        break;
      }
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

    if (selectedReport === 'MASTER' || selectedReport === 'MAIN_REPORT') {
      head = [['Order ID', 'Outlet ID', 'Vendor', 'Station', 'State', 'Status', 'Vendor ₹', 'Base ₹', 'GST ₹', 'RF Comm ₹', 'Selling ₹', 'Margin%']];
      body = data.map((r) => [
        r['IRCTC Order ID'],
        r['Outlet ID'],
        String(r['Vendor Name']).substring(0, 18),
        r['Station Code'],
        r['State'],
        r['Final Status'],
        `₹${r['Final Vendor Price']}`,
        `₹${r['Final Base Price']}`,
        `₹${r['Final GST']}`,
        `₹${r['Final RF Commission']}`,
        `₹${r['Final Selling Price']}`,
        `${r['Margin %']}%`,
      ]);
    } else if (selectedReport === 'STATION_REPORT' || selectedReport === 'LAST_DAY_STATION') {
      head = [['Station Code', 'State', 'Total Orders', 'Delivered', 'Cancelled', 'Selling Amount', 'Vendor Payout', 'RF Comm', 'GST (5%)', 'Feedback Good', 'Feedback Bad']];
      body = stationSummary.map((s) => [
        s.station,
        s.state || '-',
        s.totalOrders,
        s.delivered,
        s.cancelled,
        `₹${s.sellingPrice.toFixed(2)}`,
        `₹${s.vendorPrice.toFixed(2)}`,
        `₹${s.rfComm.toFixed(2)}`,
        `₹${s.gst.toFixed(2)}`,
        s.feedbackGood || 0,
        s.feedbackBad || 0,
      ]);
    } else if (selectedReport === 'VENDOR_REPORT' || selectedReport === 'VENDOR_RDS') {
      head = [['Vendor Name', 'Outlet ID', 'State', 'Total Orders', 'Delivered', 'Selling Amount', 'Vendor Payout', 'RF Commission', 'Penalty Total']];
      body = vendorSummary.map((v) => [
        String(v.vendor).substring(0, 22),
        v.outletId,
        v.state || '-',
        v.totalOrders,
        v.delivered,
        `₹${v.sellingPrice.toFixed(2)}`,
        `₹${v.vendorPrice.toFixed(2)}`,
        `₹${v.rfComm.toFixed(2)}`,
        `₹${v.penalty.toFixed(2)}`,
      ]);
    } else if (selectedReport === 'DATE_WISE' || selectedReport === 'VENDOR_DATE_WISE') {
      head = [['Date', 'Total Orders', 'Delivered', 'Cancelled', 'Selling Amount', 'Vendor Price', 'RF Commission']];
      body = dateSummary.map((d) => [
        d.date,
        d.totalOrders,
        d.delivered,
        d.cancelled,
        `₹${d.sellingPrice.toFixed(2)}`,
        `₹${d.vendorPrice.toFixed(2)}`,
        `₹${d.rfComm.toFixed(2)}`,
      ]);
    } else if (selectedReport === 'OUTLETS_MASTER') {
      head = [['Outlet ID', 'Outlet Name', 'Station', 'State', 'GST Number', 'IRCTC Status']];
      body = Object.values(outletsMasterInfo).map((o) => [
        o.outletId,
        o.outletName || '-',
        o.station || '-',
        o.state || '-',
        o.gst || '-',
        o.irctcStatus || '-',
      ]);
    } else if (selectedReport === 'PENALTIES') {
      head = [['Outlet ID', 'Order ID', 'Transaction Mode', 'Vendor Name', 'Date', 'Amount (₹)', 'Remarks']];
      body = penaltyRawRecords.map((p) => [
        p.outletId,
        p.orderId || '-',
        p.mode,
        p.vendorName || '-',
        p.date || '-',
        `₹${p.amount.toFixed(2)}`,
        p.remarks || '-',
      ]);
    }

    autoTable(doc, {
      head: head,
      body: body,
      startY: 65,
      theme: 'grid',
      styles: { fontSize: 8, cellPadding: 4, textColor: [30, 41, 59] },
      headStyles: { fillColor: [37, 99, 235], textColor: [255, 255, 255], fontStyle: 'bold' },
      alternateRowStyles: { fillColor: [248, 250, 252] },
      margin: { top: 60, bottom: 30, left: 20, right: 20 },
    });

    doc.save(`RELFOOD_${selectedReport}_${new Date().toISOString().slice(0, 10)}.pdf`);
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
      <tr className={`border-b border-gray-300 ${isTotal ? 'font-bold bg-white text-black' : 'bg-white text-gray-800'}`}>
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
      <div className="min-h-screen bg-slate-950 flex items-center justify-center text-slate-400 text-sm">
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
    <div className="min-h-screen bg-slate-950 text-slate-100 p-3 md:p-6 font-sans">
      {/* Top Header */}
      <header className="flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4 pb-4 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="w-11 h-11 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-700 flex items-center justify-center text-white font-bold text-xl shadow-lg shadow-emerald-950">
            📊
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h1 className="text-lg md:text-xl font-black tracking-wide text-transparent bg-clip-text bg-gradient-to-r from-white via-slate-200 to-slate-400">
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
                </select>
              </div>

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
            <span>☁️</span> Upload 6 Reports
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
              Upload RF Report, IRCTC Report, Feedback, Penalty, Current Month &amp; Outlets Master. Calculations will be permanently preserved.
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

            {/* MAIN REPORT: DATE-WISE MULTI-DAY MATRIX WITH FULL HORIZONTAL SCROLL */}
            {selectedReport === 'MAIN_REPORT' && (
              <div className="space-y-6 max-h-[75vh] overflow-y-auto pr-1">
                {mainReportBlocks
                  .filter((blk) => blk.dateLabel.toLowerCase().includes(searchTerm.toLowerCase()))
                  .map((blk, bIdx) => (
                    <div 
                      key={bIdx} 
                      className="w-full overflow-x-auto rounded-xl border border-gray-500 bg-white shadow-2xl"
                      style={{ 
                        WebkitOverflowScrolling: 'touch',
                        scrollbarWidth: 'auto',
                        scrollbarColor: '#475569 #0f172a'
                      }}
                    >
                      <table className="w-full min-w-[2100px] border-collapse text-[11px] whitespace-nowrap">
                        <thead>
                          {/* Banner 1: Red Date Header */}
                          <tr>
                            <th colSpan={38} className="bg-red-600 text-white font-bold py-2.5 text-center text-xs tracking-wider">
                              {blk.dateLabel}
                            </th>
                            <th className="bg-red-600 text-white text-[10px] text-center px-1 font-bold min-w-[60px]">
                              Outlets
                            </th>
                          </tr>

                          {/* Banner 2: Group Categories */}
                          <tr className="text-white font-bold text-center text-[10px]">
                            <th className="bg-black text-white p-2 border border-gray-400 sticky left-0 z-20 min-w-[140px] shadow-[2px_0_5px_rgba(0,0,0,0.4)]">
                              Source
                            </th>
                            <th colSpan={5} className="bg-[#5da0dc] border border-gray-300 text-white py-1">ORDERS</th>
                            <th colSpan={5} className="bg-[#78b778] border border-gray-300 text-white py-1">MEALS</th>
                            <th colSpan={3} className="bg-[#f2a879] border border-gray-300 text-white py-1">VALUE</th>
                            <th colSpan={4} className="bg-[#7db4db] border border-gray-300 text-white py-1">PREPAID</th>
                            <th colSpan={4} className="bg-[#e5989b] border border-gray-300 text-white py-1">DISCOUNT</th>
                            <th colSpan={4} className="bg-[#83b0df] border border-gray-300 text-white py-1">REVENUE</th>
                            <th colSpan={4} className="bg-[#7ea8db] border border-gray-300 text-white py-1">Complaints</th>
                            <th colSpan={4} className="bg-[#9ec899] border border-gray-300 text-white py-1">Feedback</th>
                            <th colSpan={4} className="bg-[#444444] border border-gray-300 text-white py-1">IRCTC Undelivered</th>
                            <th rowSpan={2} className="bg-[#f0c808] text-black font-extrabold border border-gray-400 text-center text-base min-w-[60px] align-middle">
                              {blk.outletsCount}
                            </th>
                          </tr>

                          {/* Banner 3: Metric Sub-Headers */}
                          <tr className="text-[10px] text-center font-bold bg-gray-100 text-gray-800">
                            <th className="border border-gray-300 p-1 sticky left-0 z-20 bg-gray-200 min-w-[140px] shadow-[2px_0_5px_rgba(0,0,0,0.3)]"></th>
                            {/* Orders */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">ASP</th><th className="border border-gray-300 px-2 py-1">Del%</th>
                            {/* Meals */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">ASP</th><th className="border border-gray-300 px-2 py-1">MPO</th>
                            {/* Value */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th>
                            {/* Prepaid */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                            {/* Discount */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                            {/* Revenue */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                            {/* Complaints */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                            {/* Feedback */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                            {/* Undelivered */}
                            <th className="border border-gray-300 px-2 py-1">FTD</th><th className="border border-gray-300 px-2 py-1">MTD</th><th className="border border-gray-300 px-2 py-1">LMTD</th><th className="border border-gray-300 px-2 py-1">%</th>
                          </tr>
                        </thead>

                        <tbody>
                          {/* Total Row */}
                          {renderMainReportRow('Total', blk.dayTotal, blk.mtdTotal, true)}

                          {/* Channel Source Rows */}
                          {SOURCES.map((src) => (
                            <React.Fragment key={src}>
                              {renderMainReportRow(src, blk.dayStats[src], blk.mtdBySource[src], false)}
                            </React.Fragment>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ))}
              </div>
            )}

            {/* ALL OTHER REPORT TABLES (WITH FULL HORIZONTAL SCROLL) */}
            {selectedReport !== 'MAIN_REPORT' && (
              <div 
                className="w-full overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/80 shadow-2xl max-h-[75vh]"
                style={{ 
                  WebkitOverflowScrolling: 'touch',
                  scrollbarWidth: 'auto',
                  scrollbarColor: '#475569 #0f172a'
                }}
              >
                {/* 1. MASTER VIEW */}
                {selectedReport === 'MASTER' && (
                  <table className="w-full min-w-[1400px] text-left border-collapse text-xs whitespace-nowrap">
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
                          <tr key={i} className="hover:bg-slate-800/40">
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
                {(selectedReport === 'STATION_REPORT' || selectedReport === 'LAST_DAY_STATION') && (
                  <table className="w-full min-w-[1200px] text-left border-collapse text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Station Code</th>
                        <th className="p-3 font-semibold">State</th>
                        <th className="p-3 font-semibold text-center">Total Orders</th>
                        <th className="p-3 font-semibold text-center text-emerald-400">Delivered</th>
                        <th className="p-3 font-semibold text-center text-rose-400">Cancelled</th>
                        <th className="p-3 font-semibold text-right">Total Selling Amount</th>
                        <th className="p-3 font-semibold text-right">Vendor Payout</th>
                        <th className="p-3 font-semibold text-right text-emerald-400">RF Commission</th>
                        <th className="p-3 font-semibold text-right text-cyan-400">GST</th>
                        <th className="p-3 font-semibold text-center text-emerald-400">Feedback Good</th>
                        <th className="p-3 font-semibold text-center text-rose-400">Feedback Bad</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {stationSummary
                        .filter((s) => s.station.toLowerCase().includes(searchTerm.toLowerCase()))
                        .map((row, i) => (
                          <tr key={i} className="hover:bg-slate-800/40">
                            <td className="p-3 font-bold text-white font-mono sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row.station}</td>
                            <td className="p-3 text-amber-300">{row.state || '-'}</td>
                            <td className="p-3 text-center font-semibold">{row.totalOrders}</td>
                            <td className="p-3 text-center text-emerald-400 font-bold">{row.delivered}</td>
                            <td className="p-3 text-center text-rose-400">{row.cancelled}</td>
                            <td className="p-3 text-right font-bold text-amber-400">₹{row.sellingPrice.toFixed(2)}</td>
                            <td className="p-3 text-right">₹{row.vendorPrice.toFixed(2)}</td>
                            <td className="p-3 text-right font-bold text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                            <td className="p-3 text-right text-cyan-400">₹{row.gst.toFixed(2)}</td>
                            <td className="p-3 text-center font-bold text-emerald-400">{row.feedbackGood || 0}</td>
                            <td className="p-3 text-center font-bold text-rose-400">{row.feedbackBad || 0}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 3. VENDOR / VENDOR RDS VIEW */}
                {(selectedReport === 'VENDOR_REPORT' || selectedReport === 'VENDOR_RDS') && (
                  <table className="w-full min-w-[1300px] text-left border-collapse text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Vendor Name</th>
                        <th className="p-3 font-semibold">Outlet ID</th>
                        <th className="p-3 font-semibold">State</th>
                        <th className="p-3 font-semibold text-center">Total Orders</th>
                        <th className="p-3 font-semibold text-center text-emerald-400">Delivered</th>
                        <th className="p-3 font-semibold text-right">Total Selling Amount</th>
                        <th className="p-3 font-semibold text-right">Vendor Payout</th>
                        <th className="p-3 font-semibold text-right text-emerald-400">RF Commission</th>
                        <th className="p-3 font-semibold text-right text-rose-400">Penalty Total</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {vendorSummary
                        .filter((v) => v.vendor.toLowerCase().includes(searchTerm.toLowerCase()))
                        .map((row, i) => (
                          <tr key={i} className="hover:bg-slate-800/40">
                            <td className="p-3 font-bold text-white sticky left-0 bg-slate-900/95 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">{row.vendor}</td>
                            <td className="p-3 text-slate-400 font-mono">{row.outletId}</td>
                            <td className="p-3 text-amber-300">{row.state || '-'}</td>
                            <td className="p-3 text-center">{row.totalOrders}</td>
                            <td className="p-3 text-center text-emerald-400 font-bold">{row.delivered}</td>
                            <td className="p-3 text-right font-bold text-amber-400">₹{row.sellingPrice.toFixed(2)}</td>
                            <td className="p-3 text-right">₹{row.vendorPrice.toFixed(2)}</td>
                            <td className="p-3 text-right font-bold text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                            <td className="p-3 text-right text-rose-400 font-semibold">₹{row.penalty.toFixed(2)}</td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 4. DATE WISE VIEW */}
                {(selectedReport === 'DATE_WISE' || selectedReport === 'VENDOR_DATE_WISE') && (
                  <table className="w-full min-w-[1100px] text-left border-collapse text-xs whitespace-nowrap">
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
                          <tr key={i} className="hover:bg-slate-800/40">
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

                {/* 5. OUTLETS MASTER VIEW */}
                {selectedReport === 'OUTLETS_MASTER' && (
                  <table className="w-full min-w-[1000px] text-left border-collapse text-xs whitespace-nowrap">
                    <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                      <tr>
                        <th className="p-3 font-semibold sticky left-0 bg-slate-900 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.5)]">Outlet ID</th>
                        <th className="p-3 font-semibold">Outlet Name</th>
                        <th className="p-3 font-semibold">Station</th>
                        <th className="p-3 font-semibold">State</th>
                        <th className="p-3 font-semibold">GST Number</th>
                        <th className="p-3 font-semibold">IRCTC Status</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 text-slate-300">
                      {Object.values(outletsMasterInfo)
                        .filter((o) => o.outletId.includes(searchTerm) || o.outletName?.toLowerCase().includes(searchTerm.toLowerCase()))
                        .map((row, i) => (
                          <tr key={i} className="hover:bg-slate-800/40">
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
                          </tr>
                        ))}
                    </tbody>
                  </table>
                )}

                {/* 6. PENALTIES VIEW */}
                {selectedReport === 'PENALTIES' && (
                  <table className="w-full min-w-[1100px] text-left border-collapse text-xs whitespace-nowrap">
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
                          <tr key={i} className="hover:bg-slate-800/40">
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

      {/* Upload Modal (6 Files) */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <div className="w-full max-w-xl rounded-2xl bg-slate-900 p-6 border border-slate-700 shadow-2xl text-white">
            <h2 className="text-xl font-bold mb-2 flex items-center gap-2">
              <span>📊</span> Upload Reports (6 Files System)
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

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-rose-400 block mb-1">
                  4. Penalty &amp; Deduction Report (Optional)
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
                  5. Current Month Data (Previous Balance, Paid by Relfood) (Optional)
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
                  6. Outlets Master (GST, State &amp; IRCTC Status) (Optional)
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
