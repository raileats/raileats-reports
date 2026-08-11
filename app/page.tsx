import React, { useState } from 'react';
import Papa from 'papaparse';

export default function FileUploadModal({ isOpen, onClose, onDataProcessed }) {
  const [rfFile, setRfFile] = useState(null);
  const [irctcFile, setIrctcFile] = useState(null);
  const [feedbackFile, setFeedbackFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusText, setStatusText] = useState('');

  if (!isOpen) return null;

  // Helper to read CSV
  const parseCSV = (file) => {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => resolve(results.data),
        error: (err) => reject(err),
      });
    });
  };

  // Helper to read RF Report (HTML / XLS)
  const parseRFReport = async (file) => {
    const text = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return [];

    const rows = Array.from(table.querySelectorAll('tr'));
    if (rows.length < 2) return [];

    const headers = Array.from(rows[0].querySelectorAll('th, td')).map((c) =>
      c.innerText.trim()
    );

    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const cells = Array.from(rows[i].querySelectorAll('td')).map((c) =>
        c.innerText.trim()
      );
      if (cells.length === headers.length) {
        const rowObj = {};
        headers.forEach((h, idx) => {
          rowObj[h] = cells[idx];
        });
        data.push(rowObj);
      }
    }
    return data;
  };

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

      let feedbackData = [];
      if (feedbackFile) {
        setStatusText('Reading Feedback Report...');
        feedbackData = await parseCSV(feedbackFile);
      }

      setStatusText('Merging & Calculating AC to BJ formulas...');

      // Lookup map for RF Report
      const rfMap = new Map();
      rfData.forEach((row) => {
        const orderId = String(row['IRCTC OrderId'] || '').trim();
        if (orderId) rfMap.set(orderId, row);
      });

      // Lookup map for Feedback Report
      const fbMap = new Map();
      feedbackData.forEach((row) => {
        const orderId = String(row['Order ID'] || '').trim();
        if (orderId) fbMap.set(orderId, row);
      });

      // Merge base on IRCTC / Master list
      const masterRows = irctcData.map((irctc) => {
        const orderId = String(irctc['Order Id'] || '').trim();
        const rf = rfMap.get(orderId) || {};
        const fb = fbMap.get(orderId) || {};

        // Calculations & Column mapping
        const basePrice = parseFloat(irctc['Total Base Price'] || rf['Total Base Price'] || 0);
        const gst = parseFloat(irctc['Total Gst'] || rf['GST'] || 0);
        const deliveryCharge = parseFloat(irctc['Delivery Charge'] || 0);
        const discount = parseFloat(irctc['discount'] || rf['Discount'] || 0);
        const sellingAmount = parseFloat(rf['Selling Amount'] || irctc['Amount Payable'] || 0);
        const vendorPrice = parseFloat(rf['Vendor Price'] || 0);
        
        const rfComm = parseFloat(rf['Relfood Commission'] || 0);
        const irctcComm = parseFloat(rf['IRCTC Commission'] || 0);
        const totalComm = parseFloat(rf['Total Commission RF + IRCTC'] || (rfComm + irctcComm));

        return {
          orderId,
          rfOrderId: rf['Relfood OrderId'] || '',
          vendorCode: irctc['Vendor Id'] || irctc['Outlet Id'] || '',
          vendorName: irctc['Vendor Name'] || rf['Vendor Name'] || '',
          outletName: irctc['Outlet Name'] || '',
          stationCode: rf['Station Code'] || irctc['Delivery Station'] || '',
          stationZone: rf['Station Zone'] || '',
          trainNo: irctc['Train No.'] || rf['Train Number'] || '',
          pnr: irctc['PNR No.'] || '',
          bookingDate: irctc['Date of Booking'] || rf['Booking Date'] || '',
          deliveryDate: irctc['Delivery Date'] || rf['Delivery Date'] || '',
          deliveryTime: rf['Delivery Time'] || '',
          orderStatus: irctc['Order Status'] || rf['Order Status'] || '',
          paymentType: irctc['Transaction Type'] || rf['Payment Type'] || '',
          
          // Financials
          basePrice,
          gst,
          deliveryCharge,
          discount,
          sellingAmount,
          vendorPrice,
          rfComm,
          irctcComm,
          totalComm,

          // Feedback & Complaints
          rating: fb['Rating'] || '',
          feedbackRemarks: fb['Remarks'] || irctc['Comments'] || '',
          feedbackType: fb['Type'] || irctc['Feedback Type'] || '',
          complaintStatus: irctc['Complaint Order status'] || ''
        };
      });

      onDataProcessed(masterRows);
      setIsProcessing(false);
      onClose();
    } catch (error) {
      console.error(error);
      alert('Error parsing files: ' + error.message);
      setIsProcessing(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
      <div className="w-full max-w-xl rounded-2xl bg-slate-900 p-6 border border-slate-700 shadow-2xl text-white">
        <h2 className="text-xl font-bold mb-4 flex items-center gap-2">
          <span>📊</span> Upload Multi-Report Data
        </h2>
        
        <p className="text-xs text-slate-400 mb-6">
          Apni 3 reports upload karein. System in teeno ko Order ID se auto-merge karke 9 reports taiyar karega.
        </p>

        <div className="space-y-4">
          {/* 1. RF Report */}
          <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
            <label className="text-sm font-semibold text-emerald-400 block mb-1">
              1. RF Report Export (.xls / HTML) *
            </label>
            <input
              type="file"
              accept=".xls,.xlsx,.html,.csv"
              onChange={(e) => setRfFile(e.target.files[0])}
              className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-emerald-600 file:text-white hover:file:bg-emerald-500 cursor-pointer"
            />
          </div>

          {/* 2. IRCTC Report */}
          <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
            <label className="text-sm font-semibold text-blue-400 block mb-1">
              2. IRCTC Report (.csv) *
            </label>
            <input
              type="file"
              accept=".csv"
              onChange={(e) => setIrctcFile(e.target.files[0])}
              className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-blue-600 file:text-white hover:file:bg-blue-500 cursor-pointer"
            />
          </div>

          {/* 3. Feedback Report */}
          <div className="p-3 bg-slate-800/60 rounded-xl border border-slate-700">
            <label className="text-sm font-semibold text-amber-400 block mb-1">
              3. Feedback Report (.csv) (Optional)
            </label>
            <input
              type="file"
              accept=".csv"
              onChange={(e) => setFeedbackFile(e.target.files[0])}
              className="text-xs text-slate-300 file:mr-3 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:bg-amber-600 file:text-white hover:file:bg-amber-500 cursor-pointer"
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
            onClick={onClose}
            className="px-4 py-2 text-sm rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 transition"
          >
            Cancel
          </button>
          <button
            onClick={handleProcessAndMerge}
            disabled={isProcessing || !rfFile || !irctcFile}
            className="px-5 py-2 text-sm font-medium rounded-lg bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-400 hover:to-teal-500 text-white disabled:opacity-50 transition shadow-lg"
          >
            {isProcessing ? 'Processing...' : 'Merge & Generate Reports'}
          </button>
        </div>
      </div>
    </div>
  );
}
