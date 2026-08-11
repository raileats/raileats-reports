'use client';

import React, { useState, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { generateVendorRDSWorkbook } from '@/lib/vendorRdsGenerator';
import { generateStationReportWorkbook } from '@/lib/stationReportGenerator';
import { generateVendorReportWorkbook } from '@/lib/vendorReportGenerator';

// --- Native IndexedDB Storage Engine (Safe across Refresh & Large Data) ---
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

// Master Final Status Calculation Rule (51 Rules Mapping)
const computeFinalStatus = (rfStatusRaw: string, irctcStatusRaw: string): string => {
  const rf = (rfStatusRaw || '').trim().toUpperCase();
  const irctc = (irctcStatusRaw || '').trim().toUpperCase();

  if (rf.includes('UNDELIVERED') || irctc.includes('UNDELIVERED')) {
    return 'Not Delivered';
  }
  if (rf.includes('CANCEL') || irctc.includes('CANCEL')) {
    return 'Cancelled';
  }
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

// Universal helper to normalize Outlet Id (e.g. "10943.0" -> "10943")
const cleanOutletId = (val: any): string => {
  if (!val && val !== 0) return '';
  return String(val).trim().replace(/\.0$/, '');
};

export default function Page() {
  const [data, setData] = useState<any[]>([]);
  const [penaltySummary, setPenaltySummary] = useState<Record<string, number>>({});
  const [currentMonthRecords, setCurrentMonthRecords] = useState<any[]>([]);
  const [outletsMasterInfo, setOutletsMasterInfo] = useState<Record<string, any>>({});
  const [isLoaded, setIsLoaded] = useState<boolean>(false);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');

  // 6 Upload States
  const [rfFile, setRfFile] = useState<File | null>(null);
  const [irctcFile, setIrctcFile] = useState<File | null>(null);
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [penaltyFile, setPenaltyFile] = useState<File | null>(null);
  const [currentMonthFile, setCurrentMonthFile] = useState<File | null>(null);
  const [outletsFile, setOutletsFile] = useState<File | null>(null);

  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');

  // 1. Load saved master records on mount
  useEffect(() => {
    const fetchStoredData = async () => {
      try {
        const storedMaster = await loadFromDB('CURRENT_MASTER_DATA');
        const storedPenalty = await loadFromDB('OUTLET_PENALTY_DATA');
        const storedCurrentMonth = await loadFromDB('CURRENT_MONTH_DATA');
        const storedOutletsInfo = await loadFromDB('OUTLET_MASTER_INFO');

        if (Array.isArray(storedMaster) && storedMaster.length > 0) {
          setData(storedMaster);
        }
        if (storedPenalty && typeof storedPenalty === 'object') {
          setPenaltySummary(storedPenalty.outletTotals || {});
        }
        if (Array.isArray(storedCurrentMonth) && storedCurrentMonth.length > 0) {
          setCurrentMonthRecords(storedCurrentMonth);
        }
        if (storedOutletsInfo && typeof storedOutletsInfo === 'object') {
          setOutletsMasterInfo(storedOutletsInfo);
        }
      } catch (err) {
        console.error('Failed to load stored records from IndexedDB:', err);
      } finally {
        setIsLoaded(true);
      }
    };
    fetchStoredData();
  }, []);

  // 2. Clear All Records manually
  const handleClearRecords = async () => {
    if (confirm('Kya aap sach me saara stored data delete karna chahte hain?')) {
      await clearDB();
      setData([]);
      setPenaltySummary({});
      setCurrentMonthRecords([]);
      setOutletsMasterInfo({});
    }
  };

  // Universal File Parser (Handles CSV, XLS, XLSX, HTML binary/text)
  const parseAnyFile = async (file: File): Promise<any[]> => {
    const arrayBuffer = await file.arrayBuffer();

    // Check for ZIP/XLSX header (PK\x03\x04)
    const uint = new Uint8Array(arrayBuffer.slice(0, 4));
    const isZip = uint[0] === 0x50 && uint[1] === 0x4B;

    if (isZip || file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
      try {
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = wb.SheetNames[0];
        return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
      } catch (err) {
        console.warn('XLSX parse fallback:', err);
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
          const headers = Array.from(rows[0].querySelectorAll('th, td')).map((c: any) =>
            c.innerText.trim()
          );
          const parsedData: any[] = [];
          for (let i = 1; i < rows.length; i++) {
            const cells = Array.from(rows[i].querySelectorAll('td')).map((c: any) =>
              c.innerText.trim()
            );
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

  // 3. Merge & Process Engine (6 Files Pipeline)
  const handleProcessAndMerge = async () => {
    if (!rfFile || !irctcFile) {
      alert('Kripya RF Report aur IRCTC Report zaroor upload karein!');
      return;
    }

    try {
      setIsProcessing(true);

      // (A) Process RF Report
      setStatusText('Reading RF Report...');
      const rfData = await parseAnyFile(rfFile);

      // (B) Process IRCTC Report
      setStatusText('Reading IRCTC Report...');
      const irctcData = await parseAnyFile(irctcFile);

      // (C) Process Feedback Report (Optional)
      let feedbackData: any[] = [];
      if (feedbackFile) {
        setStatusText('Reading Feedback Report...');
        feedbackData = await parseAnyFile(feedbackFile);
      }

      // (D) Process Penalty Report (Optional) -> Outlet Id wise Sum
      const penaltyOutletMap: Record<string, number> = {};
      const penaltyRawFilteredList: any[] = [];
      if (penaltyFile) {
        setStatusText('Processing Penalty Report (PENALTY, COUPON, COMPLAINT_REFUND, PARTIAL_DELIVERY)...');
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
      }

      // (E) Process Current Month Data (Optional)
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

      // (F) Process Outlets Report (Optional) -> Keep Top (First) Row on Duplicate
      const outletsMap: Record<string, any> = {};
      if (outletsFile) {
        setStatusText('Processing Outlets Report (GST, State & IRCTC Status)...');
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

      // Lookup Maps
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

      // Master Calculation across RF Records
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

          // 19 Master Calculated Columns
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

          // Feedback / Remarks
          'Rating': fb['Rating'] || '',
          'Remarks': fb['Remarks'] || irctc['Comments'] || '',
        };
      });

      // Save safely into IndexedDB
      await saveToDB('CURRENT_MASTER_DATA', masterRows);
      setData(masterRows);

      // Reset modal inputs
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

  const handleExportExcel = () => {
    if (data.length === 0) return alert('No data to export!');

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Data');

    XLSX.writeFile(workbook, `RELFOOD_MASTER_DATA_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleExportVendorRDS = () => {
    if (data.length === 0) return alert('Pehle Master Reports upload & process karein!');
    generateVendorRDSWorkbook(data, penaltySummary, currentMonthRecords, outletsMasterInfo);
  };

  const handleExportStationReport = () => {
    if (data.length === 0) return alert('Pehle Master Reports upload & process karein!');
    generateStationReportWorkbook(data, outletsMasterInfo);
  };

  const handleExportVendorReport = () => {
    if (data.length === 0) return alert('Pehle Master Reports upload & process karein!');
    generateVendorReportWorkbook(data, outletsMasterInfo);
  };

  if (!isLoaded) {
    return (
      <div className="min-h-screen bg-slate-950 flex items-center justify-center text-slate-400 text-sm">
        Loading saved master records &amp; database...
      </div>
    );
  }

  const penaltyOutletCount = Object.keys(penaltySummary).length;
  const outletsInfoCount = Object.keys(outletsMasterInfo).length;

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6 font-sans">
      {/* Top Header */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-400 font-bold text-xl">
            📊
          </div>
          <div>
            <div className="flex items-center gap-2 flex-wrap">
              <h1 className="text-xl font-bold tracking-wide">RELFOOD MASTER &amp; RDS PORTAL</h1>
              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                19 RULES ENGINE
              </span>
              {outletsInfoCount > 0 && (
                <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-amber-500/20 text-amber-400 border border-amber-500/30">
                  {outletsInfoCount} OUTLETS MASTER
                </span>
              )}
              {penaltyOutletCount > 0 && (
                <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-rose-500/20 text-rose-400 border border-rose-500/30">
                  {penaltyOutletCount} PENALTIES STORED
                </span>
              )}
              {currentMonthRecords.length > 0 && (
                <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-cyan-500/20 text-cyan-400 border border-cyan-500/30">
                  {currentMonthRecords.length} CURRENT MONTH STORED
                </span>
              )}
            </div>
            <p className="text-xs text-slate-400">Order Calculations, Outlet Master (GST/State), Penalties, Station &amp; Vendor Reports</p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {(data.length > 0 || penaltyOutletCount > 0 || currentMonthRecords.length > 0 || outletsInfoCount > 0) && (
            <>
              <button
                onClick={handleClearRecords}
                className="px-3.5 py-2 rounded-xl bg-rose-950/60 hover:bg-rose-900 border border-rose-700/60 text-xs font-semibold text-rose-300 transition"
              >
                🗑️ Clear All Stored Data
              </button>
              {data.length > 0 && (
                <>
                  <button
                    onClick={handleExportExcel}
                    className="px-3 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white transition shadow-lg shadow-emerald-950 flex items-center gap-1.5"
                  >
                    📥 Master Data (.xlsx)
                  </button>
                  <button
                    onClick={handleExportVendorRDS}
                    className="px-3 py-2 rounded-xl bg-violet-600 hover:bg-violet-500 text-xs font-bold text-white transition shadow-lg shadow-violet-950 flex items-center gap-1.5"
                  >
                    📋 Vendor RDS (.xlsx)
                  </button>
                  <button
                    onClick={handleExportStationReport}
                    className="px-3 py-2 rounded-xl bg-amber-600 hover:bg-amber-500 text-xs font-bold text-white transition shadow-lg shadow-amber-950 flex items-center gap-1.5"
                  >
                    🚉 Station Report (.xlsx)
                  </button>
                  <button
                    onClick={handleExportVendorReport}
                    className="px-3 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-xs font-bold text-white transition shadow-lg shadow-rose-950 flex items-center gap-1.5"
                  >
                    🏪 Vendor Report (.xlsx)
                  </button>
                </>
              )}
            </>
          )}
          <button
            onClick={() => setIsModalOpen(true)}
            className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white flex items-center gap-2 transition shadow-lg shadow-indigo-950"
          >
            <span>☁️</span> Upload 6 Reports
          </button>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="mt-6">
        {data.length === 0 ? (
          <div className="flex flex-col items-center justify-center p-16 rounded-2xl border border-dashed border-slate-800 bg-slate-900/30 text-center">
            <div className="text-5xl mb-4">📂</div>
            <h3 className="text-lg font-bold text-slate-300 mb-1">No Data Stored in Portal</h3>
            <p className="text-xs text-slate-500 max-w-md mb-6">
              Upload RF Report, IRCTC Report, Feedback, Penalty, Current Month &amp; Outlets Master. Calculations &amp; Outlet summaries will be permanently preserved.
            </p>
            <button
              onClick={() => setIsModalOpen(true)}
              className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-bold text-white transition"
            >
              Start Upload &amp; Processing
            </button>
          </div>
        ) : (
          <div>
            {/* Search & Records Count Bar */}
            <div className="flex justify-between items-center mb-4 flex-wrap gap-2">
              <div className="flex items-center gap-3 flex-wrap">
                <span className="text-xs text-slate-400">
                  Total Master Orders: <strong className="text-emerald-400 font-bold">{data.length}</strong>
                </span>
                {outletsInfoCount > 0 && (
                  <span className="text-xs text-amber-400">
                    • Outlets Master: <strong>{outletsInfoCount}</strong>
                  </span>
                )}
                {penaltyOutletCount > 0 && (
                  <span className="text-xs text-rose-400">
                    • Outlets with Penalties: <strong>{penaltyOutletCount}</strong>
                  </span>
                )}
                {currentMonthRecords.length > 0 && (
                  <span className="text-xs text-cyan-400">
                    • Current Month Outlets: <strong>{currentMonthRecords.length}</strong>
                  </span>
                )}
              </div>
              <input
                type="text"
                placeholder="Search by Order ID, Vendor, Station, Outlet ID, State..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="px-3 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 focus:outline-none focus:border-indigo-500 w-72"
              />
            </div>

            {/* Master Data Table */}
            <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/60 shadow-xl max-h-[70vh]">
              <table className="w-full text-left border-collapse text-xs whitespace-nowrap">
                <thead className="sticky top-0 bg-slate-900 z-10 border-b border-slate-800 text-slate-400">
                  <tr>
                    <th className="p-3 font-semibold">Order ID</th>
                    <th className="p-3 font-semibold">Outlet ID</th>
                    <th className="p-3 font-semibold">Vendor Name</th>
                    <th className="p-3 font-semibold">State</th>
                    <th className="p-3 font-semibold">GST No</th>
                    <th className="p-3 font-semibold">Final Status</th>
                    <th className="p-3 font-semibold text-right">Vendor Price (₹)</th>
                    <th className="p-3 font-semibold text-right">Base Price (₹)</th>
                    <th className="p-3 font-semibold text-right">Disc. Base (₹)</th>
                    <th className="p-3 font-semibold text-right">Final GST (5%) (₹)</th>
                    <th className="p-3 font-semibold text-right">IRCTC Comm (₹)</th>
                    <th className="p-3 font-semibold text-right">RF Comm (₹)</th>
                    <th className="p-3 font-semibold text-right">Selling Price (₹)</th>
                    <th className="p-3 font-semibold text-right">PPD (₹)</th>
                    <th className="p-3 font-semibold text-right">COD (₹)</th>
                    <th className="p-3 font-semibold text-right">Margin %</th>
                    <th className="p-3 font-semibold text-center">Count</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60 text-slate-300">
                  {data
                    .filter(
                      (r) =>
                        String(r['IRCTC Order ID'] || '').includes(searchTerm) ||
                        String(r['Outlet ID'] || '').includes(searchTerm) ||
                        String(r['Vendor Name'] || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                        String(r['State'] || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                        String(r['Station Code'] || '').toLowerCase().includes(searchTerm.toLowerCase())
                    )
                    .slice(0, 100)
                    .map((row, i) => (
                      <tr key={i} className="hover:bg-slate-800/40">
                        <td className="p-3 font-medium text-white">{row['IRCTC Order ID']}</td>
                        <td className="p-3 text-slate-400">{row['Outlet ID']}</td>
                        <td className="p-3 font-medium">{row['Vendor Name']}</td>
                        <td className="p-3 text-amber-300/90">{row['State'] || '-'}</td>
                        <td className="p-3 text-slate-400 font-mono text-[11px]">{row['GST No'] || '-'}</td>
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
                        <td className="p-3 text-right font-medium text-slate-200">₹{row['Final Base Price']}</td>
                        <td className="p-3 text-right text-indigo-300">₹{row['Discounted Base Price']}</td>
                        <td className="p-3 text-right text-cyan-400 font-medium">₹{row['Final GST']}</td>
                        <td className="p-3 text-right text-indigo-400">₹{row['Final IRCTC Commission']}</td>
                        <td className="p-3 text-right text-emerald-400 font-bold">₹{row['Final RF Commission']}</td>
                        <td className="p-3 text-right text-amber-400 font-medium">₹{row['Final Selling Price']}</td>
                        <td className="p-3 text-right text-blue-400">₹{row['PPD']}</td>
                        <td className="p-3 text-right text-purple-400">₹{row['COD']}</td>
                        <td className="p-3 text-right font-bold text-teal-400">{row['Margin %']}%</td>
                        <td className="p-3 text-center">{row['Orders Count']}</td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>
            <p className="text-[11px] text-slate-500 mt-2">* Showing first 100 preview rows. Click &quot;Master Data (.xlsx)&quot; to export all records.</p>
          </div>
        )}
      </main>

      {/* Upload Modal (All 6 Reports) */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <div className="w-full max-w-xl rounded-2xl bg-slate-900 p-6 border border-slate-700 shadow-2xl text-white">
            <h2 className="text-xl font-bold mb-3 flex items-center gap-2">
              <span>📊</span> Upload Reports (6 Files System)
            </h2>
            <p className="text-xs text-slate-400 mb-4">
              Files select karein. Sabhi calculated metrics aur summaries permanently save ho jayengi.
            </p>

            <div className="space-y-3 max-h-[60vh] overflow-y-auto pr-1">
              {/* 1. RF Report */}
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

              {/* 2. IRCTC Report */}
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

              {/* 3. Feedback Report */}
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

              {/* 4. Penalty Report */}
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

              {/* 5. Current Month Data */}
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

              {/* 6. Outlets Master Report */}
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

            {/* Processing State text */}
            {isProcessing && (
              <div className="mt-4 p-3 bg-indigo-950/60 border border-indigo-700/60 rounded-xl flex items-center gap-3">
                <div className="animate-spin w-4 h-4 border-2 border-indigo-400 border-t-transparent rounded-full" />
                <span className="text-xs font-medium text-indigo-300">{statusText}</span>
              </div>
            )}

            {/* Actions */}
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
