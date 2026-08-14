'use client';

import React from 'react';

interface MainReportMatrixProps {
  blocks: any[];
  searchTerm: string;
}

const SOURCES = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

const n = (v: any) => Number(v || 0);
const round = (v: any) => Math.round(n(v));
const pct = (a: any, b: any, digits = 0) => b > 0 ? `${((n(a) / n(b)) * 100).toFixed(digits)}%` : `${(0).toFixed(digits)}%`;

export default function MainReportMatrix({ blocks, searchTerm }: MainReportMatrixProps) {
  const q = searchTerm.trim().toLowerCase();
  const filteredBlocks = blocks.filter((b) => String(b.dateLabel || '').toLowerCase().includes(q));

  const row = (label: string, ftd: any, mtd: any, total = false, outletCount = 0) => {
    const orderAsp = n(ftd.orders) > 0 ? round(n(ftd.value) / n(ftd.orders)) : 0;
    const mealAsp = n(ftd.meals) > 0 ? round(n(ftd.value) / n(ftd.meals)) : 0;
    const mpo = n(ftd.orders) > 0 ? (n(ftd.meals) / n(ftd.orders)).toFixed(2) : '0.00';
    const prepaidPct = pct(ftd.prepaidValue, ftd.value, 2);
    const discountPct = pct(ftd.discount, ftd.value, 2);
    const revenuePct = pct(ftd.revenue, ftd.value, 1);
    const complaintPct = pct(ftd.complaints, ftd.deliveredOrders, 2);
    const feedbackPct = pct(ftd.feedback, ftd.deliveredOrders, 2);
    const undeliveredPct = pct(ftd.undelivered, ftd.orders, 2);

    const values = [
      ftd.orders, mtd.orders, 0, orderAsp, pct(ftd.deliveredOrders, ftd.orders, 0),
      ftd.meals, mtd.meals, 0, mealAsp, mpo,
      round(ftd.value), round(mtd.value), 0,
      round(ftd.prepaidValue), round(mtd.prepaidValue), 0, prepaidPct,
      round(ftd.discount), round(mtd.discount), 0, discountPct,
      round(ftd.revenue), round(mtd.revenue), 0, revenuePct,
      ftd.complaints, mtd.complaints, 0, complaintPct,
      ftd.feedback, mtd.feedback, 0, feedbackPct,
      ftd.undelivered, mtd.undelivered, 0, undeliveredPct,
    ];

    return (
      <tr className={total ? 'mr-total-row' : ''}>
        <td className={total ? 'mr-total-label' : 'mr-source-cell'}>{label}</td>
        {values.map((v, i) => <td key={i} className={`mr-data-cell mr-group-${i}`}>{v}</td>)}
        {total && (
          <td rowSpan={5} className="mr-outlet-empty">{outletCount}</td>
        )}
      </tr>
    );
  };

  return (
    <div className="mr-root">
      <style jsx global>{`
        .mr-root{width:100%;max-width:100%;overflow-x:hidden;overflow-y:auto;box-sizing:border-box;background:#fff}
        .mr-block{width:100%;max-width:100%;overflow:hidden;margin:0 0 12px 0;background:#fff}
        .mr-table{width:100%!important;min-width:0!important;max-width:100%!important;table-layout:fixed!important;border-collapse:collapse!important;border-spacing:0!important;font-family:Arial,Helvetica,sans-serif;font-size:9px;line-height:1}
        .mr-table th,.mr-table td{box-sizing:border-box;min-width:0!important;padding:2px 1px;height:15px;line-height:1;white-space:nowrap;overflow:hidden;text-overflow:clip;text-align:center;vertical-align:middle;border:1px solid #222}
        .mr-table .mr-source-col{width:5.55%!important}.mr-table .mr-outlet-col{width:2.75%!important}.mr-table .mr-data-col{width:2.478%!important}
        .mr-date{background:#ff0000!important;color:#fff!important;font-size:13px!important;font-weight:900!important;height:25px!important;border:1px solid #ff0000!important;padding:3px!important;text-align:center}
        .mr-source-head{background:#000!important;color:#fff!important;font-weight:900}
        .mr-orders{background:#72b9d2!important;color:#000!important}.mr-meals{background:#b5d37b!important;color:#000!important}.mr-value{background:#f3b37e!important;color:#000!important}.mr-prepaid{background:#8dc5de!important;color:#000!important}.mr-discount{background:#d99397!important;color:#000!important}.mr-revenue{background:#999!important;color:#000!important}.mr-complaints{background:#5b91c8!important;color:#000!important}.mr-feedback{background:#87c34f!important;color:#000!important}.mr-undelivered{background:#555!important;color:#fff!important}.mr-outlet-head{background:#f2a900!important;color:#fff!important;font-weight:900}
        .mr-sub{background:#dce7ec!important;color:#000!important;font-weight:900;font-size:7px!important;height:16px!important}
        .mr-sub.orders{background:#b9dce8!important}.mr-sub.meals{background:#d2e7bf!important}.mr-sub.value{background:#f8d8c0!important}.mr-sub.prepaid{background:#c4e0ec!important}.mr-sub.discount{background:#efc7ca!important}.mr-sub.revenue{background:#cfcfcf!important}.mr-sub.complaints{background:#c4d9ef!important}.mr-sub.feedback{background:#d4e8c5!important}.mr-sub.undelivered{background:#666!important;color:#fff!important}
        .mr-total-row td{font-weight:900!important}.mr-total-label{background:#000!important;color:#fff!important;font-weight:900!important}.mr-source-cell{background:#ef0000!important;color:#fff!important;font-weight:900!important;text-align:left!important;padding-left:4px!important}
        .mr-data-cell{color:#000;background:#edf4f7!important}.mr-group-0,.mr-group-1,.mr-group-2,.mr-group-3,.mr-group-4{background:#cce5ee!important}.mr-group-5,.mr-group-6,.mr-group-7,.mr-group-8,.mr-group-9{background:#d9e8c8!important}.mr-group-10,.mr-group-11,.mr-group-12{background:#f8dfca!important}.mr-group-13,.mr-group-14,.mr-group-15,.mr-group-16{background:#d6eaf2!important}.mr-group-17,.mr-group-18,.mr-group-19,.mr-group-20{background:#f0d6d9!important}.mr-group-21,.mr-group-22,.mr-group-23,.mr-group-24{background:#d0d0d0!important}.mr-group-25,.mr-group-26,.mr-group-27,.mr-group-28{background:#c8ddf0!important}.mr-group-29,.mr-group-30,.mr-group-31,.mr-group-32{background:#d9eacb!important}.mr-group-33,.mr-group-34,.mr-group-35,.mr-group-36{background:#666!important;color:#fff!important}.mr-group-20,.mr-group-21{color:#008c69!important;font-weight:900!important}.mr-group-33,.mr-group-34{color:#ff3030!important;font-weight:900!important}.mr-outlet-empty{background:#fff200!important;color:#111!important;font-weight:900!important}
        @media(max-width:1200px){.mr-table{font-size:8px}.mr-table th,.mr-table td{height:14px;padding:1px}.mr-date{font-size:12px!important;height:23px!important}}
        @media(max-width:900px){.mr-table{font-size:7px}.mr-date{font-size:10px!important}}
      `}</style>

      {filteredBlocks.map((blk, index) => (
        <div className="mr-block" key={`${blk.rawDate || blk.dateLabel}-${index}`}>
          <table className="mr-table">
            <colgroup>
              <col className="mr-source-col" />
              {Array.from({length:37}).map((_,i)=><col className="mr-data-col" key={i}/>)}
              <col className="mr-outlet-col" />
            </colgroup>
            <thead>
              <tr><th colSpan={39} className="mr-date">{blk.dateLabel}</th></tr>
              <tr>
                <th className="mr-source-head">Source</th>
                <th colSpan={5} className="mr-orders">ORDERS</th>
                <th colSpan={5} className="mr-meals">MEALS</th>
                <th colSpan={3} className="mr-value">VALUE</th>
                <th colSpan={4} className="mr-prepaid">PREPAID</th>
                <th colSpan={4} className="mr-discount">DISCOUNT</th>
                <th colSpan={4} className="mr-revenue">REVENUE</th>
                <th colSpan={4} className="mr-complaints">Complaints</th>
                <th colSpan={4} className="mr-feedback">Feedback</th>
                <th colSpan={4} className="mr-undelivered">IRCTC Undelivered</th>
                <th rowSpan={2} className="mr-outlet-head">Outlets</th>
              </tr>
              <tr>
                <th className="mr-sub"> </th>
                {['FTD','MTD','LMTD','ASP','Del%'].map((x,i)=><th key={'o'+i} className="mr-sub orders">{x}</th>)}
                {['FTD','MTD','LMTD','ASP','MPO'].map((x,i)=><th key={'m'+i} className="mr-sub meals">{x}</th>)}
                {['FTD','MTD','LMTD'].map((x,i)=><th key={'v'+i} className="mr-sub value">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'p'+i} className="mr-sub prepaid">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'d'+i} className="mr-sub discount">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'r'+i} className="mr-sub revenue">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'c'+i} className="mr-sub complaints">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'f'+i} className="mr-sub feedback">{x}</th>)}
                {['FTD','MTD','LMTD','%'].map((x,i)=><th key={'u'+i} className="mr-sub undelivered">{x}</th>)}
              </tr>
            </thead>
            <tbody>
              {row('Total', blk.dayTotal, blk.mtdTotal, true, Number(blk.outletsCount || 0))}
              {SOURCES.map(src => row(src, blk.dayStats[src] || {}, blk.mtdBySource[src] || {}, false))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}
