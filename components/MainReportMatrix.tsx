'use client';

import React from 'react';
import './main-report-matrix.css';

interface MainReportMatrixProps {
  blocks: any[];
  searchTerm: string;
}

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

const cell = 'mr-cell';
const headCell = 'mr-subhead';

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
      <tr className={`mr-data-row ${isTotal ? 'mr-total-row' : 'mr-source-row'}`}>
        <td className={`mr-source-cell ${isTotal ? 'mr-total-label' : ''}`}>{label}</td>

        <td className={cell}>{ftd.orders}</td><td className={cell}>{mtd.orders}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{orderAsp}</td><td className={cell}>{delPct}</td>
        <td className={cell}>{ftd.meals}</td><td className={cell}>{mtd.meals}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{mealAsp}</td><td className={cell}>{mpo}</td>
        <td className={`${cell} mr-value-number`}>{Math.round(ftd.value)}</td><td className={`${cell} mr-value-number`}>{Math.round(mtd.value)}</td><td className={`${cell} mr-muted`}>0</td>
        <td className={cell}>{Math.round(ftd.prepaidValue)}</td><td className={cell}>{Math.round(mtd.prepaidValue)}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{prepaidPct}</td>
        <td className={cell}>{Math.round(ftd.discount)}</td><td className={cell}>{Math.round(mtd.discount)}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{discountPct}</td>
        <td className={`${cell} mr-revenue-number`}>{Math.round(ftd.revenue)}</td><td className={`${cell} mr-revenue-number`}>{Math.round(mtd.revenue)}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{revenuePct}</td>
        <td className={cell}>{ftd.complaints}</td><td className={cell}>{mtd.complaints}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{complaintPct}</td>
        <td className={cell}>{ftd.feedback}</td><td className={cell}>{mtd.feedback}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{feedbackPct}</td>
        <td className={`${cell} mr-undelivered-number`}>{ftd.undelivered}</td><td className={`${cell} mr-undelivered-number`}>{mtd.undelivered}</td><td className={`${cell} mr-muted`}>0</td><td className={cell}>{undeliveredPct}</td>
        <td className="mr-outlet-cell">{ftd.outletsSet?.size ?? 0}</td>
      </tr>
    );
  };

  const q = searchTerm.trim().toLowerCase();
  const filteredBlocks = blocks.filter((blk) =>
    String(blk.dateLabel || '').toLowerCase().includes(q)
  );

  const subHeaders = [
    'FTD','MTD','LMTD','ASP','Del%',
    'FTD','MTD','LMTD','ASP','MPO',
    'FTD','MTD','LMTD',
    'FTD','MTD','LMTD','%',
    'FTD','MTD','LMTD','%',
    'FTD','MTD','LMTD','%',
    'FTD','MTD','LMTD','%',
    'FTD','MTD','LMTD','%',
    'FTD','MTD','LMTD','%',
  ];

  return (
    <div className="main-report-matrix">
      {filteredBlocks.map((blk, bIdx) => (
        <div
          key={`${blk.rawDate || blk.dateLabel}-${bIdx}`}
          className="mr-date-block"
        >
          <table className="portal-report-table portal-table-main mr-table">
            <colgroup>
              <col className="mr-source-col" />
              {Array.from({ length: 37 }).map((_, i) => <col key={i} className="mr-data-col" />)}
              <col className="mr-outlet-col" />
            </colgroup>

            <thead>
              <tr>
                <th colSpan={39} className="mr-date-header">
                  {blk.dateLabel}
                </th>
              </tr>

              <tr className="mr-group-header">
                <th className="mr-source-header">Source</th>
                <th colSpan={5} className="mr-orders">ORDERS</th>
                <th colSpan={5} className="mr-meals">MEALS</th>
                <th colSpan={3} className="mr-value">VALUE</th>
                <th colSpan={4} className="mr-prepaid">PREPAID</th>
                <th colSpan={4} className="mr-discount">DISCOUNT</th>
                <th colSpan={4} className="mr-revenue">REVENUE</th>
                <th colSpan={4} className="mr-complaints">Complaints</th>
                <th colSpan={4} className="mr-feedback">Feedback</th>
                <th colSpan={4} className="mr-undelivered">IRCTC Undelivered</th>
                <th className="mr-outlets-header">Outlets</th>
              </tr>

              <tr className="mr-sub-header">
                <th className="mr-sub-source"></th>
                {subHeaders.map((x, i) => (
                  <th key={i} className={headCell}>{x}</th>
                ))}
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
