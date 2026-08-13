'use client';

import React from 'react';
import './main-report-matrix.css';

type SourceRow = {
  source: string;
  orders: [number, number, number, number, string];
  meals: [number, number, number, number, number];
  value: [number, number, number];
  prepaid: [number, number, number, string];
  discount: [number, number, number, string];
  revenue: [number, number, number, string];
  complaints: [number, number, number, string];
  feedback: [number, number, number, string];
  undelivered: [number, number, number, string];
  outlets: number;
};

type DayData = {
  date: string;
  rows: SourceRow[];
};

const reportData: DayData[] = [
  {
    date: 'Saturday, 1 August, 2026',
    rows: [
      {
        source: 'Total',
        orders: [1166, 0, 0, 319, '99.46%'],
        meals: [2155, 0, 0, 228, 1.40],
        value: [352803, 0, 0],
        prepaid: [163676, 0, 0, '46.4%'],
        discount: [12340, 12340, 0, '4.00%'],
        revenue: [63619, 63619, 0, '18%'],
        complaints: [76, 0, 0, '0.00%'],
        feedback: [229, 0, 0, '0.00%'],
        undelivered: [6, 0, 0, '0.00%'],
        outlets: 184,
      },
      {
        source: 'RELFood_IRCTC',
        orders: [1104, 0, 0, 319, '99%'],
        meals: [2151, 0, 0, 164, 1.95],
        value: [352547, 0, 0],
        prepaid: [163562, 0, 0, '46.4%'],
        discount: [12340, 12340, 0, '3.5%'],
        revenue: [63619, 63619, 0, '18.0%'],
        complaints: [76, 0, 0, '0.00%'],
        feedback: [229, 0, 0, '0.00%'],
        undelivered: [6, 0, 0, '0.00%'],
        outlets: 184,
      },
      {
        source: 'RELFood_WEBSITE',
        orders: [2, 0, 0, 128, '100%'],
        meals: [4, 0, 0, 64, 2.00],
        value: [256, 0, 0],
        prepaid: [204, 0, 0, '79.7%'],
        discount: [0, 0, 0, '0.0%'],
        revenue: [0, 0, 0, '0.0%'],
        complaints: [0, 0, 0, '0.0%'],
        feedback: [0, 0, 0, '0.0%'],
        undelivered: [0, 0, 0, '0.0%'],
        outlets: 0,
      },
      {
        source: 'REL_Food_App',
        orders: [0, 0, 0, 0, '0%'],
        meals: [0, 0, 0, 0, 0],
        value: [0, 0, 0],
        prepaid: [0, 0, 0, '0.0%'],
        discount: [0, 0, 0, '0.0%'],
        revenue: [0, 0, 0, '0.0%'],
        complaints: [0, 0, 0, '0.0%'],
        feedback: [0, 0, 0, '0.0%'],
        undelivered: [0, 0, 0, '0.0%'],
        outlets: 0,
      },
      {
        source: 'MakeMyTrip',
        orders: [0, 0, 0, 0, '0%'],
        meals: [0, 0, 0, 0, 0],
        value: [0, 0, 0],
        prepaid: [0, 0, 0, '0.0%'],
        discount: [0, 0, 0, '0.0%'],
        revenue: [0, 0, 0, '0.0%'],
        complaints: [0, 0, 0, '0.0%'],
        feedback: [0, 0, 0, '0.0%'],
        undelivered: [0, 0, 0, '0.0%'],
        outlets: 0,
      },
    ],
  },

  // Isi tarah 2 August se 10 August tak data add karo.
];

/* ---------- GROUP CONFIG ---------- */

const groups = [
  {
    key: 'orders',
    title: 'ORDERS',
    className: 'orders',
    columns: ['FTD', 'MTD', 'LMTD', 'ASP', 'Del%'],
  },
  {
    key: 'meals',
    title: 'MEALS',
    className: 'meals',
    columns: ['FTD', 'MTD', 'LMTD', 'ASP', 'MPO'],
  },
  {
    key: 'value',
    title: 'VALUE',
    className: 'value',
    columns: ['FTD', 'MTD', 'LMTD'],
  },
  {
    key: 'prepaid',
    title: 'PREPAID',
    className: 'prepaid',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
  {
    key: 'discount',
    title: 'DISCOUNT',
    className: 'discount',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
  {
    key: 'revenue',
    title: 'REVENUE',
    className: 'revenue',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
  {
    key: 'complaints',
    title: 'Complaints',
    className: 'complaints',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
  {
    key: 'feedback',
    title: 'Feedback',
    className: 'feedback',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
  {
    key: 'undelivered',
    title: 'IRCTC Undelivered',
    className: 'undelivered',
    columns: ['FTD', 'MTD', 'LMTD', '%'],
  },
];

/* ---------- COMPONENT ---------- */

export default function MainReportMatrix() {
  return (
    <div className="main-report-page">

      <div className="report-toolbar">
        <div>
          <h1>RELFOOD ENTERPRISE PORTAL</h1>
          <p>
            Multi-Report Aggregation, Universal XLS/PDF Engine & Real-time
            Calculations
          </p>
        </div>

        <div className="toolbar-actions">
          <button>🌙 Night</button>
          <button className="excel-btn">📊 Excel (.xlsx)</button>
          <button className="pdf-btn">📄 Download PDF</button>
          <button className="upload-btn">☁ Upload 7 Reports</button>
        </div>
      </div>

      <div className="report-info">
        <span className="view-badge">Viewing: MAIN REPORT</span>
        <span>Total Records: <b>11650</b></span>

        <input
          className="search"
          placeholder="Search records..."
        />
      </div>

      <div className="matrix-scroll">

        {reportData.map((day) => (
          <div className="day-block" key={day.date}>

            {/* DATE BAR */}
            <div className="date-bar">
              {day.date}
            </div>

            <table className="matrix-table">

              <thead>
                {/* GROUP HEADERS */}
                <tr>
                  <th
                    className="source-header sticky-source"
                    rowSpan={2}
                  >
                    Source
                  </th>

                  {groups.map((group) => (
                    <th
                      key={group.key}
                      className={`group-header ${group.className}`}
                      colSpan={group.columns.length}
                    >
                      {group.title}
                    </th>
                  ))}

                  <th
                    className="outlets-header"
                    rowSpan={2}
                  >
                    Outlets
                  </th>
                </tr>

                {/* SUB HEADERS */}
                <tr>
                  {groups.flatMap((group) =>
                    group.columns.map((column) => (
                      <th
                        key={`${group.key}-${column}`}
                        className={`sub-header ${group.className}`}
                      >
                        {column}
                      </th>
                    ))
                  )}
                </tr>
              </thead>

              <tbody>
                {day.rows.map((row) => (
                  <tr
                    key={row.source}
                    className={row.source === 'Total' ? 'total-row' : ''}
                  >
                    <td className="source-cell sticky-source">
                      {row.source}
                    </td>

                    {groups.map((group) => {
                      const values =
                        row[group.key as keyof SourceRow] as any[];

                      return values.map((value, index) => (
                        <td
                          key={`${row.source}-${group.key}-${index}`}
                          className={group.className}
                        >
                          {formatValue(value)}
                        </td>
                      ));
                    })}

                    <td className="outlet-cell">
                      {row.outlets}
                    </td>
                  </tr>
                ))}
              </tbody>

            </table>
          </div>
        ))}

      </div>

      <div className="report-note">
        Note: FTD = Full Day Total, MTD = Month Till Date,
        LMTD = Last Month Till Date, ASP = Average Selling Price,
        MPO = Meals Per Order.
      </div>

    </div>
  );
}


/* ---------- VALUE FORMATTER ---------- */

function formatValue(value: any) {
  if (value === null || value === undefined || value === '') {
    return '0';
  }

  if (typeof value === 'number') {
    if (Number.isInteger(value)) {
      return value.toLocaleString('en-IN');
    }

    return value.toLocaleString('en-IN', {
      maximumFractionDigits: 2,
    });
  }

  return String(value);
}
