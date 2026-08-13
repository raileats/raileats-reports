'use client';

import React from 'react';

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
    <>
      <style jsx global>{`

/* Main Report Matrix — Excel-style compact colours */
.main-report-matrix {
  width: 100%;
  max-width: 100%;
  overflow-x: hidden !important;
  overflow-y: auto;
  box-sizing: border-box;
  background: #fff !important;
}

.mr-date-block {
  width: 100%;
  margin: 0 0 12px 0;
  overflow: hidden;
  background: #fff !important;
  border: 0 !important;
  border-radius: 0 !important;
  box-shadow: none !important;
}

.mr-table {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  table-layout: fixed !important;
  border-collapse: collapse !important;
  border-spacing: 0 !important;
  font-size: 7px !important;
  line-height: 1 !important;
  color: #000 !important;
}

.mr-table th,
.mr-table td {
  box-sizing: border-box !important;
  border: 1px solid #222 !important;
  padding: 2px 1px !important;
  height: 14px !important;
  min-width: 0 !important;
  max-width: 100% !important;
  overflow: hidden !important;
  text-overflow: clip !important;
  white-space: nowrap !important;
  text-align: center !important;
  vertical-align: middle !important;
}

/* Column sizing: source + 37 metric columns + outlets */
.mr-source-col { width: 7% !important; }
.mr-data-col { width: 2.43% !important; }
.mr-outlet-col { width: 2.34% !important; }

/* Excel-like red date strip */
.mr-date-header {
  background: #ff0000 !important;
  color: #fff !important;
  border: 0 !important;
  height: 24px !important;
  padding: 3px 4px !important;
  font-size: 13px !important;
  font-weight: 800 !important;
  text-align: center !important;
}

/* Group colours copied from the supplied Excel reference */
.mr-source-header {
  background: #000 !important;
  color: #fff !important;
  font-weight: 800 !important;
}

.mr-orders {
  background: #77bdd6 !important;
  color: #000 !important;
}

.mr-meals {
  background: #b6d67a !important;
  color: #000 !important;
}

.mr-value {
  background: #f6bd8f !important;
  color: #000 !important;
}

.mr-prepaid {
  background: #8fc5df !important;
  color: #000 !important;
}

.mr-discount {
  background: #d98f93 !important;
  color: #000 !important;
}

.mr-revenue {
  background: #9a9a9a !important;
  color: #000 !important;
}

.mr-complaints {
  background: #5b8fc8 !important;
  color: #000 !important;
}

.mr-feedback {
  background: #86c445 !important;
  color: #000 !important;
}

.mr-undelivered {
  background: #555 !important;
  color: #fff !important;
}

.mr-outlets-header {
  background: #f2a900 !important;
  color: #fff !important;
}

.mr-sub-header th {
  background: #e8eef2 !important;
  color: #000 !important;
  font-weight: 800 !important;
  height: 16px !important;
}

.mr-sub-source {
  background: #000 !important;
}

/* Body matches Excel: very light cells, red source rows */
.mr-data-row td {
  background: #edf4f7 !important;
  color: #000 !important;
  font-weight: 400 !important;
}

.mr-data-row.mr-total-row td {
  background: #eef4f7 !important;
  color: #000 !important;
  font-weight: 800 !important;
}

.mr-source-cell {
  background: #ef0000 !important;
  color: #fff !important;
  font-weight: 800 !important;
  text-align: left !important;
  padding-left: 3px !important;
}

.mr-total-label {
  background: #000 !important;
  color: #fff !important;
  text-align: left !important;
}

.mr-muted {
  color: #333 !important;
}

.mr-value-number {
  background: #f8e2d1 !important;
}

.mr-revenue-number {
  color: #008c69 !important;
  font-weight: 800 !important;
}

.mr-undelivered-number {
  background: #555 !important;
  color: #fff !important;
}

.mr-outlet-cell {
  background: #fff200 !important;
  color: #000 !important;
  font-weight: 900 !important;
  border-color: #222 !important;
}

/* Keep source column visually fixed without forcing horizontal scroll. */
.mr-source-cell,
.mr-source-header,
.mr-sub-source {
  position: sticky !important;
  left: 0 !important;
  z-index: 10 !important;
}

.mr-source-header,
.mr-sub-source {
  z-index: 20 !important;
}

/* Make the whole matrix fit the viewport, matching the supplied compact Excel view. */
@media (max-width: 1400px) {
  .mr-table {
    font-size: 6px !important;
  }

  .mr-table th,
  .mr-table td {
    height: 13px !important;
    padding: 1px !important;
  }

  .mr-date-header {
    font-size: 11px !important;
    height: 21px !important;
  }
}

@media (max-width: 1000px) {
  .mr-table {
    font-size: 5px !important;
  }

  .mr-date-header {
    font-size: 10px !important;
  }
}


      `}</style>
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
    </>
  );
}
