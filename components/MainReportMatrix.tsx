'use client';

import React from 'react';

interface MainReportMatrixProps {
  blocks: any[];
  searchTerm: string;
}

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

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
      <tr className={`portal-data-row main-matrix-row ${isTotal ? 'is-total' : ''}`}>
        <td className={`main-matrix-source ${isTotal ? 'is-total-source' : 'is-source'}`}>
          {label}
        </td>

        <td className="main-matrix-cell">{ftd.orders}</td>
        <td className="main-matrix-cell">{mtd.orders}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{orderAsp}</td>
        <td className="main-matrix-cell">{delPct}</td>

        <td className="main-matrix-cell">{ftd.meals}</td>
        <td className="main-matrix-cell">{mtd.meals}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{mealAsp}</td>
        <td className="main-matrix-cell">{mpo}</td>

        <td className="main-matrix-cell">{Math.round(ftd.value)}</td>
        <td className="main-matrix-cell">{Math.round(mtd.value)}</td>
        <td className="main-matrix-cell">0</td>

        <td className="main-matrix-cell">{Math.round(ftd.prepaidValue)}</td>
        <td className="main-matrix-cell">{Math.round(mtd.prepaidValue)}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{prepaidPct}</td>

        <td className="main-matrix-cell">{Math.round(ftd.discount)}</td>
        <td className="main-matrix-cell">{Math.round(mtd.discount)}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{discountPct}</td>

        <td className="main-matrix-cell">{Math.round(ftd.revenue)}</td>
        <td className="main-matrix-cell">{Math.round(mtd.revenue)}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{revenuePct}</td>

        <td className="main-matrix-cell">{ftd.complaints}</td>
        <td className="main-matrix-cell">{mtd.complaints}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{complaintPct}</td>

        <td className="main-matrix-cell">{ftd.feedback}</td>
        <td className="main-matrix-cell">{mtd.feedback}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{feedbackPct}</td>

        <td className="main-matrix-cell">{ftd.undelivered}</td>
        <td className="main-matrix-cell">{mtd.undelivered}</td>
        <td className="main-matrix-cell">0</td>
        <td className="main-matrix-cell">{undeliveredPct}</td>
      </tr>
    );
  };

  const q = searchTerm.trim().toLowerCase();
  const filteredBlocks = blocks.filter((blk) => String(blk.dateLabel || '').toLowerCase().includes(q));

  return (
    <div className="main-report-matrix space-y-3 max-h-[78vh] overflow-y-auto pr-0">
      <style jsx global>{`
        .main-report-block {
          width: 100% !important;
          max-width: 100% !important;
          overflow-x: hidden !important;
          background: #fff !important;
          border: 1px solid #9aa4b2 !important;
          border-radius: 8px !important;
        }
        .main-report-fit {
          width: 100% !important;
          min-width: 0 !important;
          max-width: 100% !important;
          table-layout: fixed !important;
          border-collapse: collapse !important;
          font-size: 9px !important;
        }
        .main-report-fit th, .main-report-fit td {
          box-sizing: border-box !important;
          min-width: 0 !important;
          max-width: none !important;
          padding: 2px 3px !important;
          height: 18px !important;
          line-height: 1 !important;
          white-space: nowrap !important;
          overflow: hidden !important;
          text-overflow: clip !important;
          border: 1px solid #9ca3af !important;
          text-align: center !important;
        }
        .main-report-fit th:first-child, .main-report-fit td:first-child { width: 5.7% !important; }
        .main-report-fit th:nth-child(39), .main-report-fit td:nth-child(39) { width: 2.8% !important; }
        .main-report-fit th:not(:first-child):not(:nth-child(39)), .main-report-fit td:not(:first-child):not(:nth-child(39)) { width: 2.41% !important; }

        .main-date-banner {
          background: #ff0000 !important; color: #fff !important; border-color: #ff0000 !important;
          height: 25px !important; padding: 3px !important; font-size: 13px !important;
          font-weight: 900 !important; text-align: center !important;
        }
        .main-date-outlets {
          background: #ff0000 !important; color:#fff !important; border-color:#ff0000 !important;
          font-size:9px !important; font-weight:900 !important;
        }
        .main-group-source { background:#000 !important; color:#fff !important; font-weight:900 !important; }
        .main-group-orders { background:#5da0dc !important; color:#fff !important; }
        .main-group-meals { background:#78b778 !important; color:#fff !important; }
        .main-group-value { background:#f2a879 !important; color:#fff !important; }
        .main-group-prepaid { background:#7db4db !important; color:#fff !important; }
        .main-group-discount { background:#e5989b !important; color:#fff !important; }
        .main-group-revenue { background:#7f7f7f !important; color:#fff !important; }
        .main-group-complaints { background:#5f95cf !important; color:#fff !important; }
        .main-group-feedback { background:#8bc05d !important; color:#fff !important; }
        .main-group-undelivered { background:#333 !important; color:#fff !important; }
        .main-group-outlets { background:#ffff00 !important; color:#111 !important; font-weight:900 !important; }

        .main-subheaders { background:#eef2f5 !important; color:#111827 !important; font-size:8px !important; font-weight:900 !important; }
        .main-sub-source { background:#c9dbe7 !important; color:#111827 !important; font-weight:900 !important; }
        .main-sub-cell { font-size:8px !important; font-weight:900 !important; }
        .main-subheaders .main-sub-cell:nth-child(n+2):nth-child(-n+6) { background:#b9dcec !important; }
        .main-subheaders .main-sub-cell:nth-child(n+7):nth-child(-n+11) { background:#d2e8c1 !important; }
        .main-subheaders .main-sub-cell:nth-child(n+12):nth-child(-n+14) { background:#f9d2b9 !important; }
        .main-subheaders .main-sub-cell:nth-child(n+15):nth-child(-n+18) { background:#c5e0ef !important; }
        .main-subheaders .main-sub-cell:nth-child(n+19):nth-child(-n+22) { background:#f1c5c8 !important; }
        .main-subheaders .main-sub-cell:nth-child(n+23):nth-child(-n+26) { background:#d0d0d0 !important; }
        .main-subheaders .main-sub-cell:nth-child(n+27):nth-child(-n+30) { background:#c5d9ee !important; }
        .main-subheaders .main-sub-cell:nth-child(n+31):nth-child(-n+34) { background:#d2e8bf !important; }
        .main-subheaders .main-sub-cell:nth-child(n+35):nth-child(-n+38) { background:#777 !important; color:#fff !important; }

        .main-report-fit tbody .main-matrix-row td:nth-child(n+2):nth-child(-n+6) { background:#d9edf7 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+7):nth-child(-n+11) { background:#e3f0d5 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+12):nth-child(-n+14) { background:#fce4d6 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+15):nth-child(-n+18) { background:#d9edf7 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+19):nth-child(-n+22) { background:#f4d7d9 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+23):nth-child(-n+26) { background:#d9d9d9 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+27):nth-child(-n+30) { background:#cfe2f3 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+31):nth-child(-n+34) { background:#d9ead3 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(n+35):nth-child(-n+38) { background:#666 !important; color:#fff !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(39) { background:#ffff00 !important; color:#111 !important; font-weight:900 !important; }

        .main-report-fit tbody .main-matrix-row .main-matrix-source {
          position: sticky !important; left: 0 !important; z-index: 6 !important;
          width:5.7% !important; color:#fff !important; font-weight:900 !important;
        }
        .main-report-fit tbody .main-matrix-row .is-total-source { background:#000 !important; }
        .main-report-fit tbody .main-matrix-row .is-source { background:#ff0000 !important; }
        .main-report-fit tbody .main-matrix-row td { color:#111827; font-weight:600; }
        .main-report-fit tbody .main-matrix-row.is-total td { font-weight:900 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(25),
        .main-report-fit tbody .main-matrix-row td:nth-child(26) { color:#047857 !important; font-weight:900 !important; }
        .main-report-fit tbody .main-matrix-row td:nth-child(35),
        .main-report-fit tbody .main-matrix-row td:nth-child(36) { color:#ff3b3b !important; font-weight:900 !important; }
        .main-report-fit tbody .main-matrix-row:hover td { filter: brightness(0.97); }
        .main-report-fit thead th { position:static !important; }
        .main-report-fit thead tr:first-child th { position:static !important; }
        .main-report-fit thead tr:nth-child(2) th:first-child,
        .main-report-fit thead tr:nth-child(3) th:first-child { position:sticky !important; left:0 !important; z-index:9 !important; }
      `}</style>

      {filteredBlocks.map((blk, bIdx) => (
        <div
          key={`${blk.rawDate || blk.dateLabel}-${bIdx}`}
          className="main-report-block w-full overflow-hidden rounded-xl border shadow-sm"
        >
          <table className="portal-report-table portal-table-main border-separate border-spacing-0 text-[10px] whitespace-nowrap">
            <thead>
              <tr>
                <th colSpan={38} className="main-date-banner">
                  {blk.dateLabel}
                </th>
                <th className="bg-red-600 text-white text-[10px] text-center px-1 font-bold">Outlets</th>
              </tr>

              <tr className="text-white font-bold text-center text-[10px]">
                <th className="bg-black text-white p-2 border border-gray-400 sticky left-0 z-20 shadow-[2px_0_5px_rgba(0,0,0,0.4)]">Source</th>
                <th colSpan={5} className="main-group-orders">ORDERS</th>
                <th colSpan={5} className="main-group-meals">MEALS</th>
                <th colSpan={3} className="main-group-value">VALUE</th>
                <th colSpan={4} className="main-group-prepaid">PREPAID</th>
                <th colSpan={4} className="main-group-discount">DISCOUNT</th>
                <th colSpan={4} className="main-group-revenue">REVENUE</th>
                <th colSpan={4} className="main-group-complaints">Complaints</th>
                <th colSpan={4} className="main-group-feedback">Feedback</th>
                <th colSpan={4} className="main-group-undelivered">IRCTC Undelivered</th>
                <th rowSpan={2} className="bg-[#f0c808] text-black font-extrabold border border-gray-400 text-center text-base align-middle">{blk.outletsCount}</th>
              </tr>

              <tr className="main-subheaders">
                <th className="border border-gray-300 p-1 sticky left-0 z-20 bg-gray-200 shadow-[2px_0_5px_rgba(0,0,0,0.3)]"></th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">ASP</th><th className="main-sub-cell">Del%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">ASP</th><th className="main-sub-cell">MPO</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
                <th className="main-sub-cell">FTD</th><th className="main-sub-cell">MTD</th><th className="main-sub-cell">LMTD</th><th className="main-sub-cell">%</th>
              </tr>
            </thead>
            <tbody>
              {renderMainReportRow('Total', blk.dayTotal, blk.mtdTotal, true)}
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
  );
}
