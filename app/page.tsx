'use client';

import React, { useState, useMemo } from 'react';
import Papa from 'papaparse';

export default function Page() {
  const [data, setData] = useState<any[]>([]);
  const [activeTab, setActiveTab] = useState<string>('Vendor RDS');
  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [searchTerm, setSearchTerm] = useState<string>('');

  // Upload Modal State
  const [rfFile, setRfFile] = useState<File | null>(null);
  const [irctcFile, setIrctcFile] = useState<File | null>(null);
  const [feedbackFile, setFeedbackFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [statusText, setStatusText] = useState<string>('');

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

  // 3-Files Process & Auto Merge Engine
  const handleProcessAndMerge = async () => {
    if (!rfFile || !irctcFile) {
      alert('Kripya RF Report aur IRCTC Report zaroor select karein!');
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

      setStatusText('Merging 15,000+ orders & calculating reports...');

      // Lookup Map for RF Report by IRCTC Order ID
      const rfMap = new Map();
      rfData.forEach((row) => {
        const orderId = String(row['IRCTC OrderId'] || '').trim();
        if (orderId) rfMap.set(orderId, row);
      });

      // Lookup Map for Feedback Report by Order ID
      const fbMap = new Map();
      feedbackData.forEach((row) => {
        const orderId = String(row['Order ID'] || '').trim();
        if (orderId) fbMap.set(orderId, row);
      });

      // Master Records Merging
      const masterRows = irctcData.map((irctc) => {
        const orderId = String(irctc['Order Id'] || '').trim();
        const rf = rfMap.get(orderId) || {};
        const fb = fbMap.get(orderId) || {};

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
          vendorName: irctc['Vendor Name'] || rf['Vendor Name'] || 'Unknown Vendor',
          outletName: irctc['Outlet Name'] || '',
          stationCode: rf['Station Code'] || irctc['Delivery Station'] || '',
          stationZone: rf['Station Zone'] || '',
          trainNo: irctc['Train No.'] || rf['Train Number'] || '',
          pnr: irctc['PNR No.'] || '',
          bookingDate: irctc['Date of Booking'] || rf['Booking Date'] || '',
          deliveryDate: irctc['Delivery Date'] || rf['Delivery Date'] || '',
          deliveryTime: rf['Delivery Time'] || '',
          orderStatus: irctc['Order Status'] || rf['Order Status'] || 'Delivered',
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
          feedbackType: fb['Type'] || irctc['Feedback Type'] || 'None',
          complaintStatus: irctc['Complaint Order status'] || ''
        };
      });

      setData(masterRows);
      setIsProcessing(false);
      setIsModalOpen(false);
    } catch (error: any) {
      console.error(error);
      alert('Error parsing files: ' + error.message);
      setIsProcessing(false);
    }
  };

  // Export to CSV Function
  const handleExportCSV = () => {
    if (data.length === 0) return alert('No data to export!');
    const csv = Papa.unparse(data);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${activeTab.replace(/\s+/g, '_')}_Report.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Grouped Calculations for Tabs
  const vendorRDSGrouped = useMemo(() => {
    const groups: { [key: string]: any } = {};
    data.forEach((row) => {
      const key = row.vendorName;
      if (!groups[key]) {
        groups[key] = {
          vendorName: key,
          vendorCode: row.vendorCode,
          orders: 0,
          basePrice: 0,
          gst: 0,
          sellingAmount: 0,
          vendorPrice: 0,
          rfComm: 0,
          irctcComm: 0,
          netPayable: 0
        };
      }
      groups[key].orders += 1;
      groups[key].basePrice += row.basePrice;
      groups[key].gst += row.gst;
      groups[key].sellingAmount += row.sellingAmount;
      groups[key].vendorPrice += row.vendorPrice;
      groups[key].rfComm += row.rfComm;
      groups[key].irctcComm += row.irctcComm;
      groups[key].netPayable += (row.vendorPrice - row.rfComm);
    });
    return Object.values(groups);
  }, [data]);

  const stationGrouped = useMemo(() => {
    const groups: { [key: string]: any } = {};
    data.forEach((row) => {
      const key = row.stationCode || 'N/A';
      if (!groups[key]) {
        groups[key] = {
          stationCode: key,
          stationZone: row.stationZone,
          orders: 0,
          sellingAmount: 0,
          rfComm: 0
        };
      }
      groups[key].orders += 1;
      groups[key].sellingAmount += row.sellingAmount;
      groups[key].rfComm += row.rfComm;
    });
    return Object.values(groups);
  }, [data]);

  const complaintsList = useMemo(() => {
    return data.filter((row) => row.feedbackRemarks || (row.rating && Number(row.rating) <= 3) || row.feedbackType !== 'None');
  }, [data]);

  const tabs = [
    'Vendor RDS',
    'VendorReport',
    'Vendor Date Wise',
    'StationReport',
    'DeliveryDateReport',
    'DateStationReport',
    'Main Report',
    'StnDatewiseReport',
    'Complaints'
  ];

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100 p-4 md:p-6 font-sans">
      {/* Top Navbar */}
      <header className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pb-6 border-b border-slate-800">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-400 font-bold text-xl">
            📄
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h1 className="text-xl font-bold tracking-wide">RELFOOD</h1>
              <span className="px-2 py-0.5 rounded text-[10px] font-semibold bg-emerald-500/20 text-emerald-400 border border-emerald-500/30">
                ENGINE v2.0
              </span>
            </div>
            <p className="text-xs text-slate-400">IRCTC Settlement, Remittance & Multi-Report Generator</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          {data.length > 0 && (
            <button
              onClick={handleExportCSV}
              className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs font-semibold text-slate-200 border border-slate-700 transition"
            >
              📥 Export CSV
            </button>
          )}
          <button
            onClick={() => setIsModalOpen(true)}
            className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-semibold text-white flex items-center gap-2 transition shadow-lg shadow-emerald-950"
          >
            <span>☁️</span> Upload 3 Raw Reports
          </button>
        </div>
      </header>

      {/* Tabs Navigation */}
      <div className="flex items-center gap-2 overflow-x-auto py-4 scrollbar-none border-b border-slate-800/80">
        {tabs.map((tab) => (
          <button
            key={tab}
            onClick={() => setActiveTab(tab)}
            className={`px-4 py-2 rounded-xl text-xs font-semibold whitespace-nowrap transition-all ${
              activeTab === tab
                ? 'bg-emerald-500 text-slate-950 shadow-md shadow-emerald-500/20 font-bold'
                : 'bg-slate-900/80 text-slate-400 hover:text-slate-200 hover:bg-slate-800 border border-slate-800'
            }`}
          >
            {tab}
          </button>
        ))}
      </div>

      {/* Main View Area */}
      <main className="mt-6">
        {data.length === 0 ? (
          <div className="flex flex-col items-center justify-center p-16 rounded-2xl border border-dashed border-slate-800 bg-slate-900/30 text-center">
            <div className="text-5xl mb-4">☁️</div>
            <h3 className="text-lg font-bold text-slate-300 mb-1">No Data Uploaded Yet</h3>
            <p className="text-xs text-slate-500 max-w-md mb-6">
              Upload raw reports (RF XLS, IRCTC CSV, Feedback CSV). The system will automatically calculate AC to BJ formulas and generate all 9 reports.
            </p>
            <button
              onClick={() => setIsModalOpen(true)}
              className="px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white transition"
            >
              Start Multi-Report Upload
            </button>
          </div>
        ) : (
          <div>
            {/* Search and Summary */}
            <div className="flex justify-between items-center mb-4">
              <span className="text-xs text-slate-400">
                Showing <strong className="text-emerald-400">{data.length}</strong> loaded orders
              </span>
              <input
                type="text"
                placeholder="Search vendor, order id, station..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="px-3 py-1.5 rounded-lg bg-slate-900 border border-slate-700 text-xs text-slate-200 focus:outline-none focus:border-emerald-500 w-64"
              />
            </div>

            {/* Table Container */}
            <div className="overflow-x-auto rounded-xl border border-slate-800 bg-slate-900/60 shadow-xl">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="border-b border-slate-800 bg-slate-900/90 text-slate-400">
                    {activeTab === 'Vendor RDS' && (
                      <>
                        <th className="p-3 font-semibold">Vendor Name</th>
                        <th className="p-3 font-semibold">Vendor Code</th>
                        <th className="p-3 font-semibold text-right">Orders</th>
                        <th className="p-3 font-semibold text-right">Base Price (₹)</th>
                        <th className="p-3 font-semibold text-right">GST (₹)</th>
                        <th className="p-3 font-semibold text-right">Selling Amt (₹)</th>
                        <th className="p-3 font-semibold text-right">RF Comm (₹)</th>
                        <th className="p-3 font-semibold text-right">Net Payable (₹)</th>
                      </>
                    )}
                    {activeTab === 'StationReport' && (
                      <>
                        <th className="p-3 font-semibold">Station Code</th>
                        <th className="p-3 font-semibold">Station Zone</th>
                        <th className="p-3 font-semibold text-right">Total Orders</th>
                        <th className="p-3 font-semibold text-right">Total Selling (₹)</th>
                        <th className="p-3 font-semibold text-right">RF Comm (₹)</th>
                      </>
                    )}
                    {activeTab === 'Complaints' && (
                      <>
                        <th className="p-3 font-semibold">Order ID</th>
                        <th className="p-3 font-semibold">Vendor</th>
                        <th className="p-3 font-semibold">Rating</th>
                        <th className="p-3 font-semibold">Type</th>
                        <th className="p-3 font-semibold">Remarks</th>
                      </>
                    )}
                    {activeTab !== 'Vendor RDS' && activeTab !== 'StationReport' && activeTab !== 'Complaints' && (
                      <>
                        <th className="p-3 font-semibold">Order ID</th>
                        <th className="p-3 font-semibold">Date</th>
                        <th className="p-3 font-semibold">Vendor</th>
                        <th className="p-3 font-semibold">Station</th>
                        <th className="p-3 font-semibold">Train No</th>
                        <th className="p-3 font-semibold text-right">Selling (₹)</th>
                        <th className="p-3 font-semibold text-right">RF Comm (₹)</th>
                        <th className="p-3 font-semibold">Status</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60">
                  {activeTab === 'Vendor RDS' &&
                    vendorRDSGrouped
                      .filter((v) => v.vendorName.toLowerCase().includes(searchTerm.toLowerCase()))
                      .map((row, i) => (
                        <tr key={i} className="hover:bg-slate-800/40 text-slate-300">
                          <td className="p-3 font-medium text-white">{row.vendorName}</td>
                          <td className="p-3 text-slate-400">{row.vendorCode}</td>
                          <td className="p-3 text-right">{row.orders}</td>
                          <td className="p-3 text-right">₹{row.basePrice.toFixed(2)}</td>
                          <td className="p-3 text-right">₹{row.gst.toFixed(2)}</td>
                          <td className="p-3 text-right">₹{row.sellingAmount.toFixed(2)}</td>
                          <td className="p-3 text-right text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                          <td className="p-3 text-right font-bold text-amber-400">₹{row.netPayable.toFixed(2)}</td>
                        </tr>
                      ))}

                  {activeTab === 'StationReport' &&
                    stationGrouped
                      .filter((s) => s.stationCode.toLowerCase().includes(searchTerm.toLowerCase()))
                      .map((row, i) => (
                        <tr key={i} className="hover:bg-slate-800/40 text-slate-300">
                          <td className="p-3 font-bold text-white">{row.stationCode}</td>
                          <td className="p-3 text-slate-400">{row.stationZone || 'N/A'}</td>
                          <td className="p-3 text-right">{row.orders}</td>
                          <td className="p-3 text-right">₹{row.sellingAmount.toFixed(2)}</td>
                          <td className="p-3 text-right text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                        </tr>
                      ))}

                  {activeTab === 'Complaints' &&
                    complaintsList
                      .filter((c) => c.orderId.includes(searchTerm) || c.vendorName.toLowerCase().includes(searchTerm.toLowerCase()))
                      .map((row, i) => (
                        <tr key={i} className="hover:bg-slate-800/40 text-slate-300">
                          <td className="p-3 font-medium text-white">{row.orderId}</td>
                          <td className="p-3">{row.vendorName}</td>
                          <td className="p-3 text-amber-400 font-bold">{row.rating ? `${row.rating} ★` : '-'}</td>
                          <td className="p-3 text-rose-400">{row.feedbackType}</td>
                          <td className="p-3 text-slate-400">{row.feedbackRemarks || '-'}</td>
                        </tr>
                      ))}

                  {activeTab !== 'Vendor RDS' && activeTab !== 'StationReport' && activeTab !== 'Complaints' &&
                    data
                      .slice(0, 100)
                      .map((row, i) => (
                        <tr key={i} className="hover:bg-slate-800/40 text-slate-300">
                          <td className="p-3 font-medium text-white">{row.orderId}</td>
                          <td className="p-3 text-slate-400">{row.deliveryDate}</td>
                          <td className="p-3">{row.vendorName}</td>
                          <td className="p-3">{row.stationCode}</td>
                          <td className="p-3">{row.trainNo}</td>
                          <td className="p-3 text-right">₹{row.sellingAmount.toFixed(2)}</td>
                          <td className="p-3 text-right text-emerald-400">₹{row.rfComm.toFixed(2)}</td>
                          <td className="p-3">
                            <span className="px-2 py-0.5 rounded text-[10px] bg-emerald-950 text-emerald-400 border border-emerald-800">
                              {row.orderStatus}
                            </span>
                          </td>
                        </tr>
                      ))}
                </tbody>
              </table>
            </div>
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
              Teeno raw files select karein. System in teeno ko Order ID se auto-merge karke saari reports taiyar karega.
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
                {isProcessing ? 'Processing...' : 'Merge & Generate Reports'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
