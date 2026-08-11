'use client';

import React, { useState } from 'react';
import { FileSpreadsheet, Download, FileText, UploadCloud } from 'lucide-react';
import Papa from 'papaparse';

interface VendorRdsSummary {
  vendorCode: string;
  vendorName: string;
  stationCode: string;
  vendorPrice: number;
  basePrice: number;
  grossComm: number;
  penalty: number;
  totalGrossCommWithPenalty: number;
  igst: number;
  netPayment: number;
}

export default function Home() {
  const [reportType, setReportType] = useState('vendor_rds');
  const [data, setData] = useState<VendorRdsSummary[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  const calculateVendorRDS = (rows: any[]): VendorRdsSummary[] => {
    const summaryMap: { [key: string]: VendorRdsSummary } = {};

    rows.forEach((row) => {
      const vendorCode = row['Vendor Code'] || row['vendor_code'] || row['Outlet ID'] || 'UNKNOWN';
      const vendorName = row['Vendor Name'] || row['vendor_name'] || row['Outlet Name'] || 'Unknown Vendor';
      const stationCode = row['Station Code'] || row['station_code'] || row['Station'] || '-';
      
      const vendorPrice = parseFloat(row['Vendor Price'] || row['vendor_price'] || '0') || 0;
      const basePrice = parseFloat(row['Base Price'] || row['base_price'] || '0') || 0;
      const penalty = parseFloat(row['Penalty'] || row['penalty'] || '0') || 0;
      const grossComm = parseFloat(row['Gross Commission'] || row['gross_comm'] || '0') || 0;

      if (!summaryMap[vendorCode]) {
        summaryMap[vendorCode] = {
          vendorCode,
          vendorName,
          stationCode,
          vendorPrice: 0,
          basePrice: 0,
          grossComm: 0,
          penalty: 0,
          totalGrossCommWithPenalty: 0,
          igst: 0,
          netPayment: 0,
        };
      }

      summaryMap[vendorCode].vendorPrice += vendorPrice;
      summaryMap[vendorCode].basePrice += basePrice;
      summaryMap[vendorCode].grossComm += grossComm;
      summaryMap[vendorCode].penalty += penalty;
    });

    return Object.values(summaryMap).map((v) => {
      const totalComm = v.grossComm + v.penalty;
      const igst = totalComm * 0.18;
      const netPayment = v.vendorPrice - (totalComm + igst);
      return {
        ...v,
        totalGrossCommWithPenalty: totalComm,
        igst,
        netPayment,
      };
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results: any) => {
        const calculated = calculateVendorRDS(results.data);
        setData(calculated);
        setIsLoading(false);
      },
      error: () => setIsLoading(false),
    });
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <main className="min-h-screen bg-slate-900 text-slate-100 p-8">
      <header className="max-w-6xl mx-auto flex justify-between items-center pb-6 border-b border-slate-800 mb-8">
        <div className="flex items-center gap-3">
          <FileSpreadsheet className="w-8 h-8 text-emerald-400" />
          <h1 className="text-2xl font-bold tracking-tight">RELFOOD Reports Engine</h1>
        </div>
        <div className="text-sm text-slate-400">Production Dashboard</div>
      </header>

      <div className="max-w-6xl mx-auto space-y-8">
        <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 shadow-xl grid md:grid-cols-2 gap-6">
          <div>
            <label className="block text-sm font-medium text-slate-300 mb-2">
              1. Select Report Type
            </label>
            <select
              value={reportType}
              onChange={(e) => setReportType(e.target.value)}
              className="w-full bg-slate-900 border border-slate-700 rounded-lg p-3 text-slate-200 focus:outline-none focus:ring-2 focus:ring-emerald-500"
            >
              <option value="vendor_rds">Vendor RDS (Settlement & Remittance)</option>
              <option value="vendor_report">Vendor Summary Report</option>
              <option value="station_report">Station Wise Report</option>
              <option value="date_report">Delivery Date Report</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium text-slate-300 mb-2">
              2. Upload Data CSV File
            </label>
            <div className="relative border-2 border-dashed border-slate-700 rounded-lg p-3 hover:border-emerald-500 transition-colors text-center cursor-pointer bg-slate-900">
              <input
                type="file"
                accept=".csv"
                onChange={handleFileUpload}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <div className="flex items-center justify-center gap-2 text-slate-400">
                <UploadCloud className="w-5 h-5 text-emerald-400" />
                <span className="text-sm">Click or Drag & Drop Data CSV</span>
              </div>
            </div>
          </div>
        </div>

        {isLoading && (
          <div className="text-center py-12 text-slate-400">
            Processing Data & Applying Excel Formulas...
          </div>
        )}

        {!isLoading && data.length > 0 && (
          <div className="bg-slate-800 border border-slate-700 rounded-xl p-6 shadow-xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-700 pb-4">
              <h2 className="text-lg font-semibold text-emerald-400 flex items-center gap-2">
                <FileText className="w-5 h-5" /> RELFOOD Settlement Summary ({data.length} Outlets)
              </h2>

              <button
                onClick={handlePrint}
                className="bg-emerald-500 hover:bg-emerald-600 text-slate-950 font-bold px-4 py-2 rounded-lg flex items-center gap-2 text-sm transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" /> Save / Print PDF Report
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs text-slate-300">
                <thead className="bg-slate-900 uppercase text-slate-400 font-medium">
                  <tr>
                    <th className="p-3">Vendor Code</th>
                    <th className="p-3">Vendor Name</th>
                    <th className="p-3">Station</th>
                    <th className="p-3">Vendor Price</th>
                    <th className="p-3">Base Price</th>
                    <th className="p-3">Gross Comm</th>
                    <th className="p-3">GST (18%)</th>
                    <th className="p-3 text-emerald-400 font-bold">Net Payable</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-700">
                  {data.map((row, idx) => (
                    <tr key={idx} className="hover:bg-slate-700/50">
                      <td className="p-3 font-mono text-slate-400">{row.vendorCode}</td>
                      <td className="p-3 font-medium text-slate-200">{row.vendorName}</td>
                      <td className="p-3">{row.stationCode}</td>
                      <td className="p-3">₹{row.vendorPrice.toFixed(2)}</td>
                      <td className="p-3">₹{row.basePrice.toFixed(2)}</td>
                      <td className="p-3">₹{row.totalGrossCommWithPenalty.toFixed(2)}</td>
                      <td className="p-3">₹{row.igst.toFixed(2)}</td>
                      <td className="p-3 font-bold text-emerald-400">
                        ₹{row.netPayment.toFixed(2)}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
