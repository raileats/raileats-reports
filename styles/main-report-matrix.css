'use client';

import React from 'react';

interface MainReportMatrixProps {
  blocks: any[];
  searchTerm: string;
}

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

const cell = 'border border-gray-300 px-0.5 py-0.5 text-[7px] leading-[1.0] text-center whitespace-nowrap overflow-hidden';
const headCell = 'border border-gray-300 px-0.5 py-1 text-[7px] leading-none text-center whitespace-nowrap';

export default function MainReportMatrix({ blocks, searchTerm }: MainReportMatrixProps) {
  const renderMainReportRow = (label: string, ftd: any, mtd: any, isTotal = false) => {
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
      <tr className={`portal-data-row ${isTotal ? 'font-bold bg-white text-black' : 'bg-white text-gray-800'}`}>
        <td className={`sticky left-0 z-10 min-w-0 ${isTotal ? 'bg-[#b40000] text-white' : 'bg-[#ef0000] text-white'} border border-gray-300 px-1 py-1 text-[7px] font-bold text-center whitespace-nowrap overflow-hidden`}>
          {label}
        </td>

        <td className={cell}>{ftd.orders}</td><td className={cell}>{mtd.orders}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{orderAsp}</td><td className={cell}>{delPct}</td>
        <td className={cell}>{ftd.meals}</td><td className={cell}>{mtd.meals}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{mealAsp}</td><td className={cell}>{mpo}</td>
        <td className={`${cell} font-semibold`}>{Math.round(ftd.value)}</td><td className={`${cell} font-semibold`}>{Math.round(mtd.value)}</td><td className={`${cell} text-gray-400`}>0</td>
        <td className={cell}>{Math.round(ftd.prepaidValue)}</td><td className={cell}>{Math.round(mtd.prepaidValue)}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{prepaidPct}</td>
        <td className={cell}>{Math.round(ftd.discount)}</td><td className={cell}>{Math.round(mtd.discount)}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{discountPct}</td>
        <td className={`${cell} font-bold text-emerald-700`}>{Math.round(ftd.revenue)}</td><td className={`${cell} font-bold text-emerald-700`}>{Math.round(mtd.revenue)}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{revenuePct}</td>
        <td className={cell}>{ftd.complaints}</td><td className={cell}>{mtd.complaints}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{complaintPct}</td>
        <td className={cell}>{ftd.feedback}</td><td className={cell}>{mtd.feedback}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{feedbackPct}</td>
        <td className={`${cell} text-rose-600 font-bold`}>{ftd.undelivered}</td><td className={`${cell} text-rose-600 font-bold`}>{mtd.undelivered}</td><td className={`${cell} text-gray-400`}>0</td><td className={cell}>{undeliveredPct}</td>
        <td className="border border-gray-300 bg-[#f0c808] text-black font-extrabold text-center text-[7px] px-0 py-0.5 whitespace-nowrap">{ftd.outletsSet?.size ?? 0}</td>
      </tr>
    );
  };

  const q = searchTerm.trim().toLowerCase();
  const filteredBlocks = blocks.filter((blk) => String(blk.dateLabel || '').toLowerCase().includes(q));

  return (
    <div className="main-report-matrix w-full space-y-4 max-h-[76vh] overflow-y-auto overflow-x-hidden pr-0" style={{ width: '100%', maxWidth: '100%' }}>
      {filteredBlocks.map((blk, bIdx) => (
        <div key={`${blk.rawDate || blk.dateLabel}-${bIdx}`} className="w-full overflow-hidden rounded-xl border border-slate-700 shadow-sm bg-white" style={{ width: '100%', maxWidth: '100%' }}>
          <table className="portal-report-table portal-table-main w-full table-fixed border-collapse whitespace-nowrap" style={{ width: '100%', minWidth: '0', maxWidth: '100%', tableLayout: 'fixed', fontSize: '7px' }}>
            <colgroup>
              <col style={{ width: '9%' }} />
              {Array.from({ length: 37 }).map((_, i) => <col key={i} style={{ width: '2.34%' }} />)}
              <col style={{ width: '2.0%' }} />
            </colgroup>
            <thead>
              <tr>
                <th colSpan={39} className="bg-red-600 text-white font-bold py-1 text-center text-[9px] tracking-wide">
                  {blk.dateLabel}
                </th>
              </tr>
              <tr className="text-white font-bold text-center text-[8px]">
                <th className="bg-black text-white border border-gray-400 py-0.5 sticky left-0 z-20 min-w-0">Source</th>
                <th colSpan={5} className="bg-[#5da0dc] border border-gray-300 py-0.5">ORDERS</th>
                <th colSpan={5} className="bg-[#78b778] border border-gray-300 py-0.5">MEALS</th>
                <th colSpan={3} className="bg-[#f2a879] border border-gray-300 py-0.5">VALUE</th>
                <th colSpan={4} className="bg-[#7db4db] border border-gray-300 py-0.5">PREPAID</th>
                <th colSpan={4} className="bg-[#e5989b] border border-gray-300 py-0.5">DISCOUNT</th>
                <th colSpan={4} className="bg-[#83b0df] border border-gray-300 py-0.5">REVENUE</th>
                <th colSpan={4} className="bg-[#7ea8db] border border-gray-300 py-0.5">Complaints</th>
                <th colSpan={4} className="bg-[#9ec899] border border-gray-300 py-0.5">Feedback</th>
                <th colSpan={4} className="bg-[#444444] border border-gray-300 py-0.5">IRCTC Undelivered</th>
                <th className="bg-[#f0c808] text-black border border-gray-400 py-0.5">Outlets</th>
              </tr>
              <tr className="text-[7px] text-center font-bold bg-gray-100 text-gray-800">
                <th className="border border-gray-300 py-0.5 sticky left-0 z-20 min-w-0 bg-gray-200"></th>
                {['FTD','MTD','LMTD','ASP','Del%','FTD','MTD','LMTD','ASP','MPO','FTD','MTD','LMTD','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%','FTD','MTD','LMTD','%'].map((x, i) => <th key={i} className={headCell}>{x}</th>)}
              </tr>
            </thead>
            <tbody>
              {renderMainReportRow('Total', blk.dayTotal, blk.mtdTotal, true)}
              {SOURCES.map((src) => <React.Fragment key={src}>{renderMainReportRow(src, blk.dayStats[src], blk.mtdBySource[src], false)}</React.Fragment>)}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}
