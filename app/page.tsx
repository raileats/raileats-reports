'use client';

import React, { useState, useEffect } from 'react';
import Papa from 'papaparse';

// Master Final Status Calculation Rule (51 Rules Mapping) - Normal internal helper
const computeFinalStatus = (rfStatusRaw: string, irctcStatusRaw: string): string => {
  const rf = (rfStatusRaw || '').trim().toUpperCase();
  const irctc = (irctcStatusRaw || '').trim().toUpperCase();

  // 1. Undelivered Checks
  if (rf.includes('UNDELIVERED') || irctc.includes('UNDELIVERED')) {
    return 'Not Delivered';
  }

  // 2. Cancelled Checks
  if (rf.includes('CANCEL') || irctc.includes('CANCEL')) {
    return 'Cancelled';
  }

  // 3. Delivered / Confirmed / Placed / Pending Checks
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

export default function Page() {
  const [data, setData] = useState<any[]>([]);
  const [isLoaded, setIsLoaded] = useState<boolean>(false);
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');

  // Upload State
  const [rfFile, setRfFile] = useState<File | null>(null);
  const [irctcFile, setIrctcFile] = useState<File | null>(null);
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');

  // 1. Load saved master records on mount
  useEffect(() => {
    try {
      const saved = localStorage.getItem('RELFOOD_MASTER_DATA');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setData(parsed);
        }
      }
    } catch (err) {
      console.error('Failed to load saved records:', err);
    } finally {
      setIsLoaded(true);
    }
  }, []);

  // 2. Save master records to storage
  const saveRecordsToStorage = (records: any[]) => {
    setData(records);
    try {
      localStorage.setItem('RELFOOD_MASTER_DATA', JSON.stringify(records));
    } catch (err) {
      console.warn('Storage limit warning:', err);
    }
  };

  // 3. Clear Old Records
  const handleClearRecords = () => {
    if (confirm('Kya aap sach me purana data delete karna chahte hain?')) {
      setData([]);
      localStorage.removeItem('RELFOOD_MASTER_DATA');
    }
  };

  // Helper: Read CSV
  const parseCSV = (file: File): Promise<any[]> => {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => resolve(results.data),
        error: (err) => reject(err),
      });
    });
  };

  // Helper: Read RF Report (HTML / XLS)
  const parseRFReport = async (file: File): Promise<any[]> => {
    const text = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return [];

    const rows = Array.from(table.querySelectorAll('tr'));
    if (rows.length < 2) return [];

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
    return parsedData;
  };

  // 4. Merge Engine with 19 Master Calculated Columns
  const handleProcessAndMerge = async () => {
    if (!rfFile || !irctcFile) {
      alert('Kripya RF Report aur IRCTC Report zaroor upload karein!');
      return;
    }

    try {
      setIsProcessing(true);
      setStatusText('Reading RF Report...');
      const rfData = await parseRFReport(rfFile);

      setStatusText('Reading IRCTC Report...');
      const irctcData = await parseCSV(irctcFile);

      let feedbackData: any[] = [];
      if (feedbackFile) {
        setStatusText('Reading Feedback Report...');
        feedbackData = await parseCSV(feedbackFile);
      }

      setStatusText('Processing 19 Master Calculation Rules...');

      // Lookup Map for IRCTC Report by Order Id
      const irctcMap = new Map();
      irctcData.forEach((row) => {
        const orderId = String(row['Order Id'] || '').trim().replace(/\.0$/, '');
        if (orderId) irctcMap.set(orderId, row);
      });

      // Lookup Map for Feedback Report by Order ID
      const fbMap = new Map();
      feedbackData.forEach((row) => {
        const orderId = String(row['Order ID'] || row['Feedback Id'] || '').trim().replace(/\.0$/, '');
        if (orderId) fbMap.set(orderId, row);
      });

      // Master Calculation across RF Records
      const masterRows = rfData.map((rf) => {
        const orderId = String(rf['IRCTC OrderId'] || '').trim().replace(/\.0$/, '');
        const irctc = irctcMap.get(orderId) || {};
        const fb = fbMap.get(orderId) || {};

        // Raw statuses
        const rfRawStatus = rf['Order Status'] || '';
        const irctcRawStatus = irctc['Order Status'] || '';

        // (1) Final Status (51 Rules)
        const finalStatus = computeFinalStatus(rfRawStatus, irctcRawStatus);

        // (2) Final Vendor Price
        const finalVendorPrice = parseFloat(rf['Vendor Price'] || 0) || 0;

        // Raw financial inputs
        const rfSellingAmount = parseFloat(rf['Selling Amount'] || 0) || 0;
        const rfDiscount = parseFloat(rf['Discount'] || 0) || 0;
        const rfGst = parseFloat(rf['GST'] || 0) || 0;
        const irctcDeliveryCharge = parseFloat(irctc['Delivery Charge'] || 0) || 0;

        // (3) Final Base Price = RF Selling Amount + Discount - GST - IRCTC Delivery Charges
        const finalBasePrice = Number((rfSellingAmount + rfDiscount - rfGst - irctcDeliveryCharge).toFixed(2));

        // (4) Final Total Commission = Final Base Price - Final Vendor Price
        const finalTotalCommission = Number((finalBasePrice - finalVendorPrice).toFixed(2));

        // (5) Final IRCTC Comm = 15% of Final Base Price
        const finalIRCTCComm = Number((finalBasePrice * 0.15).toFixed(2));

        // (6) Final RF Commission = Final Total Commission - Final IRCTC Commission
        const finalRFCommission = Number((finalTotalCommission - finalIRCTCComm).toFixed(2));

        // (7) Final GST = 5% of Final Base Price
        const finalGST = Number((finalBasePrice * 0.05).toFixed(2));

        // (8) Final Total Discount = Same from RF Report Discount
        const finalTotalDiscount = Number(rfDiscount.toFixed(2));

        // (9) Final Vendor Discount = 50% of Final Total Discount
        const finalVendorDiscount = Number((finalTotalDiscount * 0.5).toFixed(2));

        // (10) Final RF Discount = 50% of Final Total Discount
        const finalRFDiscount = Number((finalTotalDiscount * 0.5).toFixed(2));

        // (11) Delivery Charges = Same from IRCTC report
        const deliveryCharges = Number(irctcDeliveryCharge.toFixed(2));

        // (12) Final Selling Price = Final Base Price + Final GST + Delivery Charges - Final Total Discount
        const finalSellingPrice = Number((finalBasePrice + finalGST + deliveryCharges - finalTotalDiscount).toFixed(2));

        // (13) Final Order Total = Final Base Price + Final GST + Delivery Charges
        const finalOrderTotal = Number((finalBasePrice + finalGST + deliveryCharges).toFixed(2));

        // (14) Discounted Base Price = Final Base Price - Final Total Discount
        const discountedBasePrice = Number((finalBasePrice - finalTotalDiscount).toFixed(2));

        // Payment Type
        const paymentType = String(rf['Payment Type'] || irctc['Transaction Type'] || '').trim().toUpperCase();

        // (15) PPD = Final Selling Price if PRE_PAID else 0
        const isPrepaid = paymentType.includes('PRE_PAID') || paymentType.includes('PREPAID') || paymentType.includes('ONLINE');
        const ppd = isPrepaid ? finalSellingPrice : 0;

        // (16) COD = Final Selling Price if CASH_ON_DELIVERY else 0
        const isCOD = paymentType.includes('CASH') || paymentType.includes('COD');
        const cod = isCOD ? finalSellingPrice : 0;

        // (17) Meals = Same from IRCTC
        const meals = parseInt(irctc['Meal Count'] || '1', 10) || 1;

        // (18) Check Margin % = ((Final Base Price - Vendor Price) / Final Base Price) * 100
        const marginPct = finalBasePrice > 0 ? Number((((finalBasePrice - finalVendorPrice) / finalBasePrice) * 100).toFixed(2)) : 0;

        // (19) Orders Count = 1 if Delivered else 0
        const ordersCount = finalStatus === 'Delivered' ? 1 : 0;

        return {
          'IRCTC Order ID': orderId,
          'RF Order ID': rf['Relfood OrderId'] || '',
          'Outlet ID': rf['OutletId'] || irctc['Outlet Id'] || '',
          'Vendor Name': rf['Vendor Name'] || irctc['Vendor Name'] || '',
          'Station Code': rf['Station Code'] || irctc['Delivery Station'] || '',
          'Train No': rf['Train Number'] || irctc['Train No.'] || '',
          'Booking Date': rf['Booking Date'] || irctc['Date of Booking'] || '',
          'Delivery Date': rf['Delivery Date'] || irctc['Delivery Date'] || '',
          'Delivery Time': rf['Delivery Time'] || '',
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
          'Final GST': finalGST,
          'Final Total Discount': finalTotalDiscount,
          'Final Vendor Discount': finalVendorDiscount,
          'Final RF Discount': finalRFDiscount,
          'Delivery Charges': deliveryCharges,
          'Final Selling Price': finalSellingPrice,
          'Final Order Total': finalOrderTotal,
          'Discounted Base Price': discountedBasePrice,
          'PPD': ppd,
          'COD': cod,
          'Meals': meals,
          'Margin %': marginPct,
          'Orders Count': ordersCount,

          // Feedback Fields
          'Rating': fb['Rating'] || '',
          'Remarks': fb['Remarks'] || irctc['Comments'] || '',
        };
      });

      saveRecordsToStorage(masterRows);
      setIsProcessing(false);
      setIsModalOpen(false);
    } catch (error: any) {
      console.error(error);
      alert('Error parsing files: ' + error.message);
      setIsProcessing(false);
    }
  };

  // 5. Export Master Data to Excel / CSV
  const handleExportExcel = () => {
    if (data.length === 0) return alert('No data to export!');
    const csv = Papa.unparse(data);
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `RELFOOD_MASTER_DATA_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  if (!isLoaded) return null;

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6 font-sans">
      {/* Top Header */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-400 font-bold text-xl">
            📊
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-xl font-bold tracking-wide">RELFOOD MASTER DATA</h1>
              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                19 RULES ENGINE
              </span>
            </div>
            <p className="text-xs text-slate-400">Order-level Calculations &amp; Persistent Storage</p>
          </div>
        </div>

        <div className="flex items-center gap-2 flex-wrap">
          {data.length > 0 && (
            <>
              <button
                onClick={handleClearRecords}
                className="px-3.5 py-2 rounded-xl bg-rose-950/60 hover:bg-rose-900 border border-rose-700/60 text-xs font-semibold text-rose-300 transition"
              >
                🗑️ Clear Old Records
              </button>
              <button
                onClick={handleExportExcel}
                className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white transition shadow-lg shadow-emerald-950 flex items-center gap-1.5"
              >
                📥 Download Master Data (Excel/CSV)
              </button>
            </>
          )}
          <button
            onClick={() => setIsModalOpen(true)}
            className="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white flex items-center gap-2 transition shadow-lg shadow-indigo-950"
          >
            <span>☁️</span> Upload 3 Raw Reports
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
              Upload RF Report, IRCTC Report &amp; Feedback Report. The 19 Master Calculated Rules will automatically process and stay stored locally.
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
            {/* Search & Info */}
            <div className="flex justify-between items-center mb-4 flex-wrap gap-2">
              <span className="text-xs text-slate-400">
                Total Stored Master Records: <strong className="text-emerald-400 font-bold">{data.length}</strong> (Safe across refresh)
              </span>
              <input
                type="text"
                placeholder="Search by Order ID, Vendor, Station..."
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
                    <th className="p-3 font-semibold">Station</th>
                    <th className="p-3 font-semibold">Final Status</th>
                    <th className="p-3 font-semibold text-right">Vendor Price (₹)</th>
                    <th className="p-3 font-semibold text-right">Base Price (₹)</th>
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
                        r['IRCTC Order ID'].includes(searchTerm) ||
                        r['Vendor Name'].toLowerCase().includes(searchTerm.toLowerCase()) ||
                        r['Station Code'].toLowerCase().includes(searchTerm.toLowerCase())
                    )
                    .slice(0, 100)
                    .map((row, i) => (
                      <tr key={i} className="hover:bg-slate-800/40">
                        <td className="p-3 font-medium text-white">{row['IRCTC Order ID']}</td>
                        <td className="p-3 text-slate-400">{row['Outlet ID']}</td>
                        <td className="p-3 font-medium">{row['Vendor Name']}</td>
                        <td className="p-3">{row['Station Code']}</td>
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
            <p className="text-[11px] text-slate-500 mt-2">* Showing first 100 preview rows. Click &quot;Download Master Data&quot; to export all records.</p>
          </div>
        )}
      </main>

      {/* Upload Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <div className="w-full max-w-xl rounded-2xl bg-slate-900 p-6 border border-slate-700 shadow-2xl text-white">
            <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
              <span>📊</span> Upload 3 Raw Reports
            </h2>
            <p className="text-xs text-slate-400 mb-6">
              Teeno files select karein. System inko merge karke 19 Master Calculated columns generate karega aur local storage me store karega.
            </p>

            <div className="space-y-4">
              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-emerald-400 block mb-1">
                  1. RF Report (.xls / HTML) *
                </label>
                <input
                  type="file"
                  accept=".xls,.xlsx,.html,.csv"
                  onChange={(e) => setRfFile(e.target.files?.[0] || null)}
                  className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-emerald-600 file:text-white cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-blue-400 block mb-1">
                  2. IRCTC Report (.csv) *
                </label>
                <input
                  type="file"
                  accept=".csv"
                  onChange={(e) => setIrctcFile(e.target.files?.[0] || null)}
                  className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-blue-600 file:text-white cursor-pointer"
                />
              </div>

              <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
                <label className="text-xs font-semibold text-amber-400 block mb-1">
                  3. Feedback Report (.csv) (Optional)
                </label>
                <input
                  type="file"
                  accept=".csv"
                  onChange={(e) => setFeedbackFile(e.target.files?.[0] || null)}
                  className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-amber-600 file:text-white cursor-pointer"
                />
              </div>
            </div>

            {isProcessing && (
              <div className="mt-4 p-3 rounded-lg bg-indigo-950/60 border border-indigo-700/50 text-indigo-300 text-xs animate-pulse text-center">
                ⏳ {statusText}
              </div>
            )}

            <div className="mt-6 flex justify-end gap-3">
              <button
                onClick={() => setIsModalOpen(false)}
                className="px-4 py-2 text-xs rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300"
              >
                Cancel
              </button>
              <button
                onClick={handleProcessAndMerge}
                disabled={isProcessing || !rfFile || !irctcFile}
                className="px-5 py-2 text-xs font-semibold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white disabled:opacity-50 shadow-lg"
              >
                {isProcessing ? 'Processing...' : 'Process & Store Master Data'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
