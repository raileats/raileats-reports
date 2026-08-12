import React, { useMemo } from 'react';
import * as XLSX from 'xlsx';

interface Props {
  masterData: any[];
}

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

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

const getSource = (row: any): string => {
  const channel = String(row['Source'] || row['Channel'] || row['Booking Channel'] || '').toUpperCase();
  const orderId = String(row['IRCTC Order ID'] || row['Order ID'] || '').toUpperCase();

  if (channel.includes('MMT') || channel.includes('MAKEMYTRIP')) return 'MakeMyTrip';
  if (channel.includes('APP') || channel.includes('REL_APP')) return 'REL_Food_App';
  if (channel.includes('WEB') || channel.includes('WEBSITE')) return 'RELFood_WEBSITE';
  return 'RELFood_IRCTC';
};

const formatDisplayDate = (dateStr: string): string => {
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString('en-GB', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
};

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

export const MainDateWiseReport: React.FC<Props> = ({ masterData }) => {
  // Process all dates & MTD running totals
  const processedBlocks = useMemo(() => {
    if (!masterData || masterData.length === 0) return [];

    // 1. Group by Date
    const dateMap: Record<string, any[]> = {};
    masterData.forEach((row) => {
      const rawDate = row['Delivery Date'] || row['Booking Date'] || 'Unknown Date';
      const dateKey = String(rawDate).split(' ')[0].split('T')[0];
      if (!dateMap[dateKey]) dateMap[dateKey] = [];
      dateMap[dateKey].push(row);
    });

    const sortedDates = Object.keys(dateMap).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

    // MTD Accumulators
    const mtdBySource: Record<string, MetricStats> = {};
    SOURCES.forEach((s) => (mtdBySource[s] = createEmptyStats()));
    let mtdTotal = createEmptyStats();

    return sortedDates.map((dateKey) => {
      const rows = dateMap[dateKey];
      const dayStats: Record<string, MetricStats> = {};
      SOURCES.forEach((s) => (dayStats[s] = createEmptyStats()));
      const dayTotal = createEmptyStats();

      rows.forEach((r) => {
        const src = getSource(r);
        const isDelivered = r['Final Status'] === 'Delivered';
        const isUndelivered = r['Final Status'] === 'Not Delivered' || String(r['IRCTC Status'] || '').toUpperCase().includes('UNDELIVERED');
        const sellingPrice = parseFloat(r['Final Selling Price'] || r['Selling Price'] || 0) || 0;
        const discount = parseFloat(r['Final Total Discount'] || r['Total Discount'] || 0) || 0;
        const rfComm = parseFloat(r['Final RF Commission'] || r['RF Comm'] || 0) || 0;
        const prepaid = parseFloat(r['PPD'] || r['Prepaid'] || 0) || 0;
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

      // Combine day totals & add to MTD
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

        // Update MTD per source
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

      // Update MTD Grand total
      mtdTotal.orders += dayTotal.orders;
      mtdTotal.deliveredOrders += dayTotal.deliveredOrders;
      mtdTotal.meals += dayTotal.meals;
      mtdTotal.value += dayTotal.value;
      mtdTotal.prepaidValue += dayTotal.prepaidValue;
      mtdTotal.discount += dayTotal.discount;
      mtdTotal.revenue += dayTotal.revenue;
      mtdTotal.complaints += dayTotal.complaints;
      mtdTotal.feedback += dayTotal.feedback;
      mtdTotal.undelivered += dayTotal.undelivered;

      return {
        dateLabel: formatDisplayDate(dateKey),
        dayTotal: { ...dayTotal },
        dayStats: JSON.parse(JSON.stringify(dayStats)),
        mtdTotal: { ...mtdTotal },
        mtdBySource: JSON.parse(JSON.stringify(mtdBySource)),
        outletsCount: dayTotal.outletsSet.size,
      };
    });
  }, [masterData]);

  const renderRow = (label: string, ftd: MetricStats, mtd: MetricStats, isTotal: boolean = false) => {
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
      <tr className={isTotal ? 'font-bold bg-white text-black' : 'text-gray-800'}>
        <td className={`p-1 border text-xs text-center ${isTotal ? 'bg-[#990000] text-white font-bold' : 'bg-red-600 text-white font-semibold'}`}>
          {label}
        </td>

        {/* ORDERS */}
        <td className="p-1 border text-xs text-center">{ftd.orders}</td>
        <td className="p-1 border text-xs text-center">{mtd.orders}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{orderAsp}</td>
        <td className="p-1 border text-xs text-center">{delPct}</td>

        {/* MEALS */}
        <td className="p-1 border text-xs text-center">{ftd.meals}</td>
        <td className="p-1 border text-xs text-center">{mtd.meals}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{mealAsp}</td>
        <td className="p-1 border text-xs text-center">{mpo}</td>

        {/* VALUE */}
        <td className="p-1 border text-xs text-center">{Math.round(ftd.value)}</td>
        <td className="p-1 border text-xs text-center">{Math.round(mtd.value)}</td>
        <td className="p-1 border text-xs text-center">0</td>

        {/* PREPAID */}
        <td className="p-1 border text-xs text-center">{Math.round(ftd.prepaidValue)}</td>
        <td className="p-1 border text-xs text-center">{Math.round(mtd.prepaidValue)}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{prepaidPct}</td>

        {/* DISCOUNT */}
        <td className="p-1 border text-xs text-center">{Math.round(ftd.discount)}</td>
        <td className="p-1 border text-xs text-center">{Math.round(mtd.discount)}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{discountPct}</td>

        {/* REVENUE */}
        <td className="p-1 border text-xs text-center">{Math.round(ftd.revenue)}</td>
        <td className="p-1 border text-xs text-center">{Math.round(mtd.revenue)}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{revenuePct}</td>

        {/* Complaints */}
        <td className="p-1 border text-xs text-center">{ftd.complaints}</td>
        <td className="p-1 border text-xs text-center">{mtd.complaints}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{complaintPct}</td>

        {/* Feedback */}
        <td className="p-1 border text-xs text-center">{ftd.feedback}</td>
        <td className="p-1 border text-xs text-center">{mtd.feedback}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{feedbackPct}</td>

        {/* IRCTC Undelivered */}
        <td className="p-1 border text-xs text-center">{ftd.undelivered}</td>
        <td className="p-1 border text-xs text-center">{mtd.undelivered}</td>
        <td className="p-1 border text-xs text-center">0</td>
        <td className="p-1 border text-xs text-center">{undeliveredPct}</td>
      </tr>
    );
  };

  if (processedBlocks.length === 0) {
    return <div className="p-4 text-center text-gray-500">Koi Data Uplabdh Nahi Hai</div>;
  }

  return (
    <div className="w-full overflow-x-auto space-y-6 p-2 bg-white">
      {processedBlocks.map((blk, idx) => (
        <div key={idx} className="border border-gray-400 shadow-sm overflow-hidden">
          <table className="w-full border-collapse text-[11px]">
            <thead>
              {/* BANNER 1: RED DATE HEADER */}
              <tr>
                <th colSpan={38} className="bg-red-600 text-white font-bold py-1 text-center text-sm tracking-wide">
                  {blk.dateLabel}
                </th>
                <th className="bg-red-600 text-white text-[10px] text-center px-1 font-bold">Outlets</th>
              </tr>

              {/* GROUP HEADERS */}
              <tr className="text-white font-bold text-center">
                <th className="bg-black text-white p-1 border">Source</th>
                <th colSpan={5} className="bg-[#5da0dc] border text-white">ORDERS</th>
                <th colSpan={5} className="bg-[#78b778] border text-white">MEALS</th>
                <th colSpan={3} className="bg-[#f2a879] border text-white">VALUE</th>
                <th colSpan={4} className="bg-[#7db4db] border text-white">PREPAID</th>
                <th colSpan={4} className="bg-[#e5989b] border text-white">DISCOUNT</th>
                <th colSpan={4} className="bg-[#83b0df] border text-white">REVENUE</th>
                <th colSpan={4} className="bg-[#7ea8db] border text-white">Complaints</th>
                <th colSpan={4} className="bg-[#9ec899] border text-white">Feedback</th>
                <th colSpan={4} className="bg-[#444444] border text-white">IRCTC Undelivered</th>
                <th rowSpan={2} className="bg-[#f0c808] text-black font-bold border text-center text-sm w-12">
                  {blk.outletsCount}
                </th>
              </tr>

              {/* SUB-COLUMN HEADERS */}
              <tr className="text-[10px] text-center font-semibold bg-gray-100">
                <th className="border p-0.5"></th>
                {/* Orders */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">ASP</th><th className="border">Del%</th>
                {/* Meals */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">ASP</th><th className="border">MPO</th>
                {/* Value */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th>
                {/* Prepaid */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
                {/* Discount */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
                {/* Revenue */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
                {/* Complaints */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
                {/* Feedback */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
                {/* Undelivered */}
                <th className="border">FTD</th><th className="border">MTD</th><th className="border">LMTD</th><th className="border">%</th>
              </tr>
            </thead>

            <tbody>
              {/* Total Row */}
              {renderRow('Total', blk.dayTotal, blk.mtdTotal, true)}

              {/* Source Rows */}
              {SOURCES.map((src) => (
                <React.Fragment key={src}>
                  {renderRow(src, blk.dayStats[src], blk.mtdBySource[src], false)}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
};
