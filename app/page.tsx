'use client';

import React, { useState, useMemo } from 'react';
import { 
  FileSpreadsheet, 
  Download, 
  FileText, 
  UploadCloud, 
  Filter, 
  Search,
  RefreshCw,
  Printer
} from 'lucide-react';
import Papa from 'papaparse';

// Raw Input Row (A to AB) + Calculated (AC to BJ)
interface OrderRecord {
  // Raw Columns (A to AB)
  orderId: string;
  orderDate: string;
  deliveryDate: string;
  deliveryTime: string;
  pnr: string;
  trainNo: string;
  trainName: string;
  stationCode: string;
  stationName: string;
  state: string;
  vendorCode: string;
  vendorName: string;
  outletName: string;
  vendorGstin: string;
  fssaiNo: string;
  basePrice: number;
  itemTotal: number;
  deliveryCharge: number;
  discount: number;
  customerPaid: number;
  vendorPrice: number;
  paymentMode: string;
  orderStatus: string;
  delayMinutes: number;
  complaintType: string;

  // Calculated Columns (AC to BJ)
  grossCommission: number;
  platformFee: number;
  delayPenalty: number;
  complaintPenalty: number;
  cancellationPenalty: number;
  totalPenalties: number;
  totalGrossCommWithPenalties: number;
  cgst: number;
  sgst: number;
  igst: number;
  totalGst: number;
  tcs: number;
  tds: number;
  netPayableVendor: number;
  relfoodMargin: number;
}

export default function ReportsDashboard() {
  const [reportType, setReportType] = useState<string>('Vendor RDS');
  const [records, setRecords] = useState<OrderRecord[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  
  // Filter States
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedStation, setSelectedStation] = useState<string>('ALL');
  const [selectedVendor, setSelectedVendor] = useState<string>('ALL');
  const [selectedDate, setSelectedDate] = useState<string>('ALL');

  // Excel AC-BJ Formulas Calculation Engine
  const processCalculations = (rows: any[]): OrderRecord[] => {
    return rows.map((r) => {
      // 1. Inputs reading (Col A - AB)
      const orderId = String(r['Order ID'] || r['order_id'] || r['Booking ID'] || '').trim();
      const orderDate = String(r['Order Date'] || r['order_date'] || '').trim();
      const deliveryDate = String(r['Delivery Date'] || r['delivery_date'] || orderDate || '').trim();
      const deliveryTime = String(r['Delivery Time'] || r['delivery_time'] || '').trim();
      const pnr = String(r['PNR'] || r['pnr'] || '').trim();
      const trainNo = String(r['Train No'] || r['train_no'] || '').trim();
      const trainName = String(r['Train Name'] || r['train_name'] || '').trim();
      const stationCode = String(r['Station Code'] || r['station_code'] || r['Station'] || '').trim().toUpperCase();
      const stationName = String(r['Station Name'] || r['station_name'] || stationCode).trim();
      const state = String(r['State'] || r['state'] || '').trim();
      const vendorCode = String(r['Vendor Code'] || r['vendor_code'] || r['Outlet ID'] || 'UNKNOWN').trim().toUpperCase();
      const vendorName = String(r['Vendor Name'] || r['vendor_name'] || r['Outlet Name'] || 'Unknown Vendor').trim();
      const outletName = String(r['Outlet Name'] || r['outlet_name'] || vendorName).trim();
      const vendorGstin = String(r['Vendor GSTIN'] || r['GSTIN'] || '').trim();
      const fssaiNo = String(r['FSSAI'] || r['fssai'] || '').trim();

      const basePrice = parseFloat(r['Base Price'] || r['base_price'] || '0') || 0;
      const itemTotal = parseFloat(r['Item Total'] || r['item_total'] || basePrice.toString()) || 0;
      const deliveryCharge = parseFloat(r['Delivery Charge'] || r['delivery_charge'] || '0') || 0;
      const discount = parseFloat(r['Discount'] || r['discount'] || '0') || 0;
      const customerPaid = parseFloat(r['Customer Paid'] || r['customer_paid'] || (itemTotal + deliveryCharge - discount).toString()) || 0;
      const vendorPrice = parseFloat(r['Vendor Price'] || r['vendor_price'] || basePrice.toString()) || 0;
      const paymentMode = String(r['Payment Mode'] || r['payment_mode'] || 'PREPAID').trim();
      const orderStatus = String(r['Order Status'] || r['order_status'] || r['Status'] || 'Delivered').trim();
      const delayMinutes = parseFloat(r['Delay Minutes'] || r['delay'] || '0') || 0;
      const complaintType = String(r['Complaint Type'] || r['complaint'] || '').trim();

      // 2. Calculated Logic (Col AC to BJ)
      const commissionRate = 0.10; // 10% Standard Commission
      const grossCommission = parseFloat(r['Gross Commission'] || '') || (basePrice * commissionRate);
      const platformFee = 5.0; // Flat convenience/platform fee
      
      const delayPenalty = delayMinutes > 30 ? 50 : 0;
      const complaintPenalty = complaintType && complaintType !== 'None' ? 100 : 0;
      const cancellationPenalty = orderStatus.toLowerCase().includes('cancel') ? 150 : 0;
      const totalPenalties = parseFloat(r['Penalty'] || '') || (delayPenalty + complaintPenalty + cancellationPenalty);

      const totalGrossCommWithPenalties = grossCommission + totalPenalties;

      // GST Calculation (18% Total - Splitting into CGST+SGST or IGST)
      const isInterState = state.toLowerCase().includes('delhi') ? false : true;
      const igst = isInterState ? (totalGrossCommWithPenalties * 0.18) : 0;
      const cgst = !isInterState ? (totalGrossCommWithPenalties * 0.09) : 0;
      const sgst = !isInterState ? (totalGrossCommWithPenalties * 0.09) : 0;
      const totalGst = igst + cgst + sgst;

      // Statutory deductions
      const tcs = basePrice * 0.01; // 1% TCS
      const tds = vendorPrice * 0.001; // 0.1% TDS

      // Final Net Settlement
      const netPayableVendor = orderStatus.toLowerCase().includes('cancel') 
        ? 0 
        : (vendorPrice - (totalGrossCommWithPenalties + totalGst + tcs + tds));

      // RELFOOD Margin
      const relfoodMargin = totalGrossCommWithPenalties - 15; // Net commission minus IRCTC/Ops cost

      return {
        orderId: orderId || `ORD-${Math.floor(100000 + Math.random() * 900000)}`,
        orderDate,
        deliveryDate,
        deliveryTime,
        pnr,
        trainNo,
        trainName,
        stationCode,
        stationName,
        state,
        vendorCode,
        vendorName,
        outletName,
        vendorGstin,
        fssaiNo,
        basePrice,
        itemTotal,
        deliveryCharge,
        discount,
        customerPaid,
        vendorPrice,
        paymentMode,
        orderStatus,
        delayMinutes,
        complaintType,
        grossCommission,
        platformFee,
        delayPenalty,
        complaintPenalty,
        cancellationPenalty,
        totalPenalties,
        totalGrossCommWithPenalties,
        cgst,
        sgst,
        igst,
        totalGst,
        tcs,
        tds,
        netPayableVendor,
        relfoodMargin
      };
    }).filter(item => item.vendorCode !== 'UNKNOWN' || item.vendorPrice > 0);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsLoading(true);
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results: any) => {
        const calculated = processCalculations(results.data);
        setRecords(calculated);
        setIsLoading(false);
      },
      error: () => setIsLoading(false),
    });
  };

  // Dropdown options
  const stationList = useMemo(() => {
    const set = new Set<string>();
    records.forEach(r => { if (r.stationCode) set.add(r.stationCode); });
    return Array.from(set).sort();
  }, [records]);

  const vendorList = useMemo(() => {
    const set = new Set<string>();
    records.forEach(r => { if (r.vendorName) set.add(r.vendorName); });
    return Array.from(set).sort();
  }, [records]);

  const dateList = useMemo(() => {
    const set = new Set<string>();
    records.forEach(r => { if (r.deliveryDate) set.add(r.deliveryDate); });
    return Array.from(set).sort();
  }, [records]);

  // Filtered master records
  const filteredRecords = useMemo(() => {
    return records.filter(r => {
      const matchSearch = searchTerm === '' || 
        r.vendorName.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.vendorCode.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.stationCode.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.orderId.toLowerCase().includes(searchTerm.toLowerCase());

      const matchStation = selectedStation === 'ALL' || r.stationCode === selectedStation;
      const matchVendor = selectedVendor === 'ALL' || r.vendorName === selectedVendor;
      const matchDate = selectedDate === 'ALL' || r.deliveryDate === selectedDate;

      return matchSearch && matchStation && matchVendor && matchDate;
    });
  }, [records, searchTerm, selectedStation, selectedVendor, selectedDate]);

  // Aggregated Reports Computation
  const reportData = useMemo(() => {
    if (filteredRecords.length === 0) return { columns: [], rows: [] };

    switch (reportType) {
      case 'Vendor RDS': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          if (!map[r.vendorCode]) {
            map[r.vendorCode] = {
              vendorCode: r.vendorCode,
              vendorName: r.vendorName,
              stationCode: r.stationCode,
              ordersCount: 0,
              vendorPrice: 0,
              basePrice: 0,
              grossComm: 0,
              penalties: 0,
              gst: 0,
              tcs: 0,
              netPayable: 0
            };
          }
          map[r.vendorCode].ordersCount += 1;
          map[r.vendorCode].vendorPrice += r.vendorPrice;
          map[r.vendorCode].basePrice += r.basePrice;
          map[r.vendorCode].grossComm += r.grossCommission;
          map[r.vendorCode].penalties += r.totalPenalties;
          map[r.vendorCode].gst += r.totalGst;
          map[r.vendorCode].tcs += r.tcs;
          map[r.vendorCode].netPayable += r.netPayableVendor;
        });

        return {
          columns: ['Vendor Code', 'Vendor Name', 'Station', 'Orders', 'Vendor Price', 'Gross Comm', 'Penalty', 'GST (18%)', 'Net Settlement'],
          rows: Object.values(map).map(v => [
            v.vendorCode,
            v.vendorName,
            v.stationCode,
            v.ordersCount,
            `₹${v.vendorPrice.toFixed(2)}`,
            `₹${v.grossComm.toFixed(2)}`,
            `₹${v.penalties.toFixed(2)}`,
            `₹${v.gst.toFixed(2)}`,
            `₹${v.netPayable.toFixed(2)}`
          ])
        };
      }

      case 'VendorReport': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          if (!map[r.vendorCode]) {
            map[r.vendorCode] = {
              vendorCode: r.vendorCode,
              vendorName: r.vendorName,
              totalOrders: 0,
              delivered: 0,
              cancelled: 0,
              totalSales: 0,
              totalMargin: 0
            };
          }
          map[r.vendorCode].totalOrders += 1;
          if (r.orderStatus.toLowerCase().includes('deliver')) map[r.vendorCode].delivered += 1;
          if (r.orderStatus.toLowerCase().includes('cancel')) map[r.vendorCode].cancelled += 1;
          map[r.vendorCode].totalSales += r.customerPaid;
          map[r.vendorCode].totalMargin += r.relfoodMargin;
        });

        return {
          columns: ['Vendor Code', 'Vendor Name', 'Total Orders', 'Delivered', 'Cancelled', 'Gross Sales', 'RELFOOD Margin'],
          rows: Object.values(map).map(v => [
            v.vendorCode,
            v.vendorName,
            v.totalOrders,
            v.delivered,
            v.cancelled,
            `₹${v.totalSales.toFixed(2)}`,
            `₹${v.totalMargin.toFixed(2)}`
          ])
        };
      }

      case 'Vendor Date Wise': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          const key = `${r.vendorCode}_${r.deliveryDate}`;
          if (!map[key]) {
            map[key] = {
              vendorCode: r.vendorCode,
              vendorName: r.vendorName,
              deliveryDate: r.deliveryDate || 'N/A',
              orders: 0,
              sales: 0,
              netPayable: 0
            };
          }
          map[key].orders += 1;
          map[key].sales += r.vendorPrice;
          map[key].netPayable += r.netPayableVendor;
        });

        return {
          columns: ['Delivery Date', 'Vendor Code', 'Vendor Name', 'Orders', 'Vendor Sales', 'Net Payout'],
          rows: Object.values(map).map(v => [
            v.deliveryDate,
            v.vendorCode,
            v.vendorName,
            v.orders,
            `₹${v.sales.toFixed(2)}`,
            `₹${v.netPayable.toFixed(2)}`
          ])
        };
      }

      case 'StationReport': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          if (!map[r.stationCode]) {
            map[r.stationCode] = {
              stationCode: r.stationCode,
              stationName: r.stationName,
              orders: 0,
              totalSales: 0,
              vendorPayout: 0,
              relfoodEarnings: 0
            };
          }
          map[r.stationCode].orders += 1;
          map[r.stationCode].totalSales += r.customerPaid;
          map[r.stationCode].vendorPayout += r.netPayableVendor;
          map[r.stationCode].relfoodEarnings += r.relfoodMargin;
        });

        return {
          columns: ['Station Code', 'Station Name', 'Total Orders', 'Gross Sales', 'Vendor Payout', 'RELFOOD Revenue'],
          rows: Object.values(map).map(s => [
            s.stationCode,
            s.stationName,
            s.orders,
            `₹${s.totalSales.toFixed(2)}`,
            `₹${s.vendorPayout.toFixed(2)}`,
            `₹${s.relfoodEarnings.toFixed(2)}`
          ])
        };
      }

      case 'DeliveryDateReport': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          const d = r.deliveryDate || 'No Date';
          if (!map[d]) {
            map[d] = {
              date: d,
              orders: 0,
              sales: 0,
              commission: 0,
              gst: 0,
              netPayout: 0
            };
          }
          map[d].orders += 1;
          map[d].sales += r.customerPaid;
          map[d].commission += r.totalGrossCommWithPenalties;
          map[d].gst += r.totalGst;
          map[d].netPayout += r.netPayableVendor;
        });

        return {
          columns: ['Delivery Date', 'Total Orders', 'Customer Sales', 'Gross Commission', 'Total GST', 'Net Vendor Payout'],
          rows: Object.values(map).map(d => [
            d.date,
            d.orders,
            `₹${d.sales.toFixed(2)}`,
            `₹${d.commission.toFixed(2)}`,
            `₹${d.gst.toFixed(2)}`,
            `₹${d.netPayout.toFixed(2)}`
          ])
        };
      }

      case 'DateStationReport': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          const key = `${r.deliveryDate}_${r.stationCode}`;
          if (!map[key]) {
            map[key] = {
              date: r.deliveryDate || 'N/A',
              stationCode: r.stationCode,
              orders: 0,
              sales: 0,
              payout: 0
            };
          }
          map[key].orders += 1;
          map[key].sales += r.customerPaid;
          map[key].payout += r.netPayableVendor;
        });

        return {
          columns: ['Delivery Date', 'Station Code', 'Orders', 'Gross Sales', 'Vendor Payout'],
          rows: Object.values(map).map(v => [
            v.date,
            v.stationCode,
            v.orders,
            `₹${v.sales.toFixed(2)}`,
            `₹${v.payout.toFixed(2)}`
          ])
        };
      }

      case 'Main Report': {
        return {
          columns: ['Order ID', 'Delivery Date', 'Station', 'Vendor', 'Base Price', 'Vendor Price', 'Gross Comm', 'Penalty', 'GST', 'Net Payable', 'Status'],
          rows: filteredRecords.map(r => [
            r.orderId,
            r.deliveryDate,
            r.stationCode,
            r.vendorName,
            `₹${r.basePrice.toFixed(2)}`,
            `₹${r.vendorPrice.toFixed(2)}`,
            `₹${r.grossCommission.toFixed(2)}`,
            `₹${r.totalPenalties.toFixed(2)}`,
            `₹${r.totalGst.toFixed(2)}`,
            `₹${r.netPayableVendor.toFixed(2)}`,
            r.orderStatus
          ])
        };
      }

      case 'StnDatewiseReport': {
        const map: { [key: string]: any } = {};
        filteredRecords.forEach(r => {
          const key = `${r.stationCode}_${r.deliveryDate}`;
          if (!map[key]) {
            map[key] = {
              stationCode: r.stationCode,
              date: r.deliveryDate || 'N/A',
              orders: 0,
              sales: 0,
              margin: 0
            };
          }
          map[key].orders += 1;
          map[key].sales += r.customerPaid;
          map[key].margin += r.relfoodMargin;
        });

        return {
          columns: ['Station Code', 'Delivery Date', 'Total Orders', 'Sales Amount', 'RELFOOD Margin'],
          rows: Object.values(map).map(v => [
            v.stationCode,
            v.date,
            v.orders,
            `₹${v.sales.toFixed(2)}`,
            `₹${v.margin.toFixed(2)}`
          ])
        };
      }

      case 'Complaints': {
        const complaintsData = filteredRecords.filter(r => r.totalPenalties > 0 || r.complaintType || r.delayMinutes > 20);
        return {
          columns: ['Order ID', 'Vendor Name', 'Station', 'Complaint / Issue', 'Delay (Mins)', 'Penalty Deduction', 'Order Status'],
          rows: complaintsData.map(r => [
            r.orderId,
            r.vendorName,
            r.stationCode,
            r.complaintType || (r.delayMinutes > 20 ? 'Late Delivery' : 'Standard Fine'),
            `${r.delayMinutes} min`,
            `₹${r.totalPenalties.toFixed(2)}`,
            r.orderStatus
          ])
        };
      }

      default:
        return { columns: [], rows: [] };
    }
  }, [reportType, filteredRecords]);

  // Export CSV Action
  const exportToCsv = () => {
    if (reportData.rows.length === 0) return;
    const csvContent = [
      reportData.columns.join(','),
      ...reportData.rows.map(row => row.map(val => `"${val}"`).join(','))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `${reportType.replace(/\s+/g, '_')}_Report.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const reportTabs = [
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
    <main className="min-h-screen bg-slate-950 text-slate-100 p-6 md:p-8 font-sans">
      {/* Header */}
      <header className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center justify-between pb-6 border-b border-slate-800 gap-4 mb-6">
        <div className="flex items-center gap-3">
          <div className="bg-emerald-500/10 p-2.5 rounded-xl border border-emerald-500/20">
            <FileSpreadsheet className="w-8 h-8 text-emerald-400" />
          </div>
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-white flex items-center gap-2">
              RELFOOD <span className="text-xs bg-emerald-500/20 text-emerald-300 font-mono px-2 py-0.5 rounded border border-emerald-500/30">ENGINE v2.0</span>
            </h1>
            <p className="text-xs text-slate-400">IRCTC Settlement, Remittance & Multi-Report Generator</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="relative border-2 border-dashed border-emerald-500/40 rounded-xl px-4 py-2 hover:border-emerald-400 transition-colors cursor-pointer bg-slate-900/60 text-center">
            <input
              type="file"
              accept=".csv"
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <div className="flex items-center gap-2 text-emerald-300 text-sm font-medium">
              <UploadCloud className="w-4 h-4 text-emerald-400" />
              <span>{records.length > 0 ? `Loaded ${records.length} Records` : 'Upload Raw Data CSV'}</span>
            </div>
          </div>

          {records.length > 0 && (
            <button
              onClick={() => setRecords([])}
              className="p-2 text-slate-400 hover:text-rose-400 transition-colors"
              title="Clear Data"
            >
              <RefreshCw className="w-5 h-5" />
            </button>
          )}
        </div>
      </header>

      <div className="max-w-7xl mx-auto space-y-6">
        {/* Tab Navigation Menu (Matching Photo Tabs) */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-2 border-b border-slate-800 scrollbar-thin">
          {reportTabs.map((tab) => (
            <button
              key={tab}
              onClick={() => setReportType(tab)}
              className={`px-4 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all ${
                reportType === tab
                  ? 'bg-emerald-500 text-slate-950 shadow-lg shadow-emerald-500/20'
                  : 'bg-slate-900/80 text-slate-400 hover:bg-slate-800 hover:text-slate-200 border border-slate-800'
              }`}
            >
              {tab}
            </button>
          ))}
        </div>

        {/* Global Multi-Filters Bar */}
        {records.length > 0 && (
          <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-4 grid grid-cols-1 md:grid-cols-4 gap-3 shadow-sm">
            {/* Search */}
            <div className="relative">
              <Search className="w-4 h-4 absolute left-3 top-3 text-slate-500" />
              <input
                type="text"
                placeholder="Search vendor, order, station..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500"
              />
            </div>

            {/* Station Filter */}
            <select
              value={selectedStation}
              onChange={(e) => setSelectedStation(e.target.value)}
              className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500"
            >
              <option value="ALL">All Stations ({stationList.length})</option>
              {stationList.map(stn => <option key={stn} value={stn}>{stn}</option>)}
            </select>

            {/* Vendor Filter */}
            <select
              value={selectedVendor}
              onChange={(e) => setSelectedVendor(e.target.value)}
              className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500"
            >
              <option value="ALL">All Vendors ({vendorList.length})</option>
              {vendorList.map(v => <option key={v} value={v}>{v}</option>)}
            </select>

            {/* Date Filter */}
            <select
              value={selectedDate}
              onChange={(e) => setSelectedDate(e.target.value)}
              className="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-emerald-500"
            >
              <option value="ALL">All Delivery Dates ({dateList.length})</option>
              {dateList.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
        )}

        {/* Loading Spinner */}
        {isLoading && (
          <div className="text-center py-16 bg-slate-900 border border-slate-800 rounded-2xl">
            <RefreshCw className="w-8 h-8 text-emerald-400 animate-spin mx-auto mb-3" />
            <p className="text-sm text-slate-300 font-medium">Executing AC-BJ Formula Calculations...</p>
          </div>
        )}

        {/* Empty State */}
        {!isLoading && records.length === 0 && (
          <div className="text-center py-20 bg-slate-900/40 border border-dashed border-slate-800 rounded-2xl">
            <UploadCloud className="w-12 h-12 text-slate-600 mx-auto mb-3" />
            <h3 className="text-base font-semibold text-slate-300 mb-1">No Data Uploaded Yet</h3>
            <p className="text-xs text-slate-500 max-w-sm mx-auto mb-4">
              Upload raw CSV (Columns A to AB). The system will automatically calculate formulas AC to BJ and generate all 9 reports.
            </p>
          </div>
        )}

        {/* Report Output Display */}
        {!isLoading && records.length > 0 && (
          <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl">
            {/* Report Header & Action Buttons */}
            <div className="p-4 border-b border-slate-800 flex flex-col md:flex-row md:items-center justify-between gap-3 bg-slate-900/60">
              <div className="flex items-center gap-2">
                <FileText className="w-5 h-5 text-emerald-400" />
                <h2 className="text-sm font-bold text-white uppercase tracking-wide">
                  {reportType} <span className="text-slate-400 font-normal">({reportData.rows.length} Rows Generated)</span>
                </h2>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={exportToCsv}
                  className="bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-semibold px-3 py-1.5 rounded-lg text-xs flex items-center gap-1.5 transition-colors cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" /> Export Excel/CSV
                </button>
                <button
                  onClick={() => window.print()}
                  className="bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 px-3 py-1.5 rounded-lg text-xs flex items-center gap-1.5 transition-colors cursor-pointer"
                >
                  <Printer className="w-3.5 h-3.5" /> Print
                </button>
              </div>
            </div>

            {/* Table */}
            <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
              <table className="w-full text-left text-xs text-slate-300">
                <thead className="bg-slate-950 uppercase text-slate-400 font-medium sticky top-0 z-10 border-b border-slate-800">
                  <tr>
                    {reportData.columns.map((col, idx) => (
                      <th key={idx} className="p-3.5 whitespace-nowrap">{col}</th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/60">
                  {reportData.rows.length === 0 ? (
                    <tr>
                      <td colSpan={reportData.columns.length} className="text-center py-8 text-slate-500">
                        No records match the selected filter criteria.
                      </td>
                    </tr>
                  ) : (
                    reportData.rows.map((row, rIdx) => (
                      <tr key={rIdx} className="hover:bg-slate-800/40 transition-colors">
                        {row.map((val: any, cIdx: number) => (
                          <td 
                            key={cIdx} 
                            className={`p-3.5 whitespace-nowrap ${
                              String(val).startsWith('₹') && cIdx === row.length - 1 
                                ? 'font-bold text-emerald-400' 
                                : ''
                            }`}
                          >
                            {val}
                          </td>
                        ))}
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
