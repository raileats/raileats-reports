import * as XLSX from 'xlsx';

// Universal date normalizer
const normalizeToDeliveryDate = (dateVal: any): string => {
  if (!dateVal) return '';
  const str = String(dateVal).trim();

  if (str.includes('-')) {
    const parts = str.split(' ')[0].split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return `${parseInt(parts[2], 10)}/${parseInt(parts[1], 10)}/${parts[0]}`;
      } else {
        return `${parseInt(parts[0], 10)}/${parseInt(parts[1], 10)}/${parts[2]}`;
      }
    }
  }

  if (str.includes('/')) {
    const parts = str.split(' ')[0].split('/');
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        return `${parseInt(parts[0], 10)}/${parseInt(parts[1], 10)}/${parts[2]}`;
      }
    }
  }

  return str;
};

// Convert "D/M/YYYY" or "DD/MM/YYYY" to comparable timestamp/number
const dateToTimestamp = (dateStr: string): number => {
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    return new Date(year, month - 1, day).getTime();
  }
  return 0;
};

export const generateLastDayStationReportWorkbook = (
  masterData: any[],
  outletsMasterInfo: Record<string, any>
) => {
  if (!masterData || masterData.length === 0) {
    alert('Master Data khali hai! Pehle reports process karein.');
    return;
  }

  // 1. Find the Maximum/Last Delivery Date across all records
  let maxTimestamp = -1;
  let lastDeliveryDateStr = '';

  masterData.forEach((row) => {
    const dStr = normalizeToDeliveryDate(row['Delivery Date'] || row['Booking Date'] || '');
    if (dStr) {
      const ts = dateToTimestamp(dStr);
      if (ts > maxTimestamp) {
        maxTimestamp = ts;
        lastDeliveryDateStr = dStr;
      }
    }
  });

  if (!lastDeliveryDateStr) {
    alert('Koi valid Delivery Date nahi mili!');
    return;
  }

  // 2. Filter rows belonging ONLY to the last delivery date
  const lastDayRows = masterData.filter((row) => {
    const dStr = normalizeToDeliveryDate(row['Delivery Date'] || row['Booking Date'] || '');
    return dStr === lastDeliveryDateStr;
  });

  if (lastDayRows.length === 0) {
    alert(`Last date (${lastDeliveryDateStr}) par koi records nahi mile.`);
    return;
  }

  // 3. Group by Station Code
  const stationMap: Record<
    string,
    {
      stationCode: string;
      stationName: string;
      vendorPrice: number;
      finalBasePrice: number;
      finalTotalComm: number;
      finalIrctcComm: number;
      finalRfComm: number;
      finalGst: number;
      finalDiscount: number;
      finalVendorDiscount: number;
      finalRfDiscount: number;
      deliveryCharges: number;
      finalSellingPrice: number;
      finalOrderTotal: number;
      discountedBasePrice: number;
      ppd: number;
      cod: number;
      meals: number;
      deliveredOrdersCount: number;
      deliveredOutlets: Set<string>;
      stationAllOutlets: Set<string>;
    }
  > = {};

  // First, map all known outlets from outletsMasterInfo to stations if available, to calculate Total Station Vendors accurately
  const stationTotalOutletsMap: Record<string, Set<string>> = {};
  Object.values(outletsMasterInfo).forEach((out: any) => {
    const stCode = String(out.station || '').trim().toUpperCase();
    const outId = String(out.outletId || '').trim();
    if (stCode && outId) {
      if (!stationTotalOutletsMap[stCode]) stationTotalOutletsMap[stCode] = new Set();
      stationTotalOutletsMap[stCode].add(outId);
    }
  });

  lastDayRows.forEach((r) => {
    const stationCode = String(r['Station Code'] || 'UNKNOWN').trim().toUpperCase();
    const stationName = String(r['Station Name'] || r['Station'] || stationCode).trim().toUpperCase();
    const outletId = String(r['Outlet ID'] || '').trim();
    const isDelivered = String(r['Final Status'] || '').trim().toLowerCase() === 'delivered';

    if (!stationMap[stationCode]) {
      stationMap[stationCode] = {
        stationCode,
        stationName,
        vendorPrice: 0,
        finalBasePrice: 0,
        finalTotalComm: 0,
        finalIrctcComm: 0,
        finalRfComm: 0,
        finalGst: 0,
        finalDiscount: 0,
        finalVendorDiscount: 0,
        finalRfDiscount: 0,
        deliveryCharges: 0,
        finalSellingPrice: 0,
        finalOrderTotal: 0,
        discountedBasePrice: 0,
        ppd: 0,
        cod: 0,
        meals: 0,
        deliveredOrdersCount: 0,
        deliveredOutlets: new Set<string>(),
        stationAllOutlets: stationTotalOutletsMap[stationCode] || new Set<string>(),
      };
    }

    const st = stationMap[stationCode];
    if (stationName && st.stationName === st.stationCode) {
      st.stationName = stationName;
    }

    st.vendorPrice += parseFloat(r['Final Vendor Price'] || 0) || 0;
    st.finalBasePrice += parseFloat(r['Final Base Price'] || 0) || 0;
    st.finalTotalComm += parseFloat(r['Final Total Commission'] || 0) || 0;
    st.finalIrctcComm += parseFloat(r['Final IRCTC Commission'] || 0) || 0;
    st.finalRfComm += parseFloat(r['Final RF Commission'] || 0) || 0;
    st.finalGst += parseFloat(r['Final GST'] || 0) || 0;
    st.finalDiscount += parseFloat(r['Final Total Discount'] || r['Discount'] || 0) || 0;
    st.finalVendorDiscount += parseFloat(r['Final Vendor Discount'] || 0) || 0;
    st.finalRfDiscount += parseFloat(r['Final RF Discount'] || 0) || 0;
    st.deliveryCharges += parseFloat(r['Delivery Charges'] || 0) || 0;
    st.finalSellingPrice += parseFloat(r['Final Selling Price'] || 0) || 0;
    st.finalOrderTotal += parseFloat(r['Final Order Total'] || 0) || 0;
    st.discountedBasePrice += parseFloat(r['Discounted Base Price'] || 0) || 0;
    st.ppd += parseFloat(r['PPD'] || 0) || 0;
    st.cod += parseFloat(r['COD'] || 0) || 0;
    st.meals += parseInt(r['Meals'] || '1', 10) || 1;

    if (isDelivered) {
      st.deliveredOrdersCount += 1;
      if (outletId) {
        st.deliveredOutlets.add(outletId);
        st.stationAllOutlets.add(outletId);
      }
    }
  });

  // Convert to array and sort by Final Base Price descending (or Total Commission descending)
  const sortedStations = Object.values(stationMap).sort(
    (a, b) => b.finalBasePrice - a.finalBasePrice
  );

  // Calculate Grand Totals for Top Summary Row
  let sumVendorPrice = 0;
  let sumFinalBasePrice = 0;
  let sumFinalTotalComm = 0;
  let sumFinalIrctcComm = 0;
  let sumFinalRfComm = 0;
  let sumFinalGst = 0;
  let sumFinalDiscount = 0;
  let sumFinalVendorDiscount = 0;
  let sumFinalRfDiscount = 0;
  let sumDeliveryCharges = 0;
  let sumFinalSellingPrice = 0;
  let sumFinalOrderTotal = 0;
  let sumDiscountedBasePrice = 0;
  let sumPpd = 0;
  let sumCod = 0;
  let sumMeals = 0;
  let sumDeliveredOrders = 0;
  let sumDeliveredOutlets = 0;
  let sumStationVendors = 0;

  sortedStations.forEach((st) => {
    sumVendorPrice += st.vendorPrice;
    sumFinalBasePrice += st.finalBasePrice;
    sumFinalTotalComm += st.finalTotalComm;
    sumFinalIrctcComm += st.finalIrctcComm;
    sumFinalRfComm += st.finalRfComm;
    sumFinalGst += st.finalGst;
    sumFinalDiscount += st.finalDiscount;
    sumFinalVendorDiscount += st.finalVendorDiscount;
    sumFinalRfDiscount += st.finalRfDiscount;
    sumDeliveryCharges += st.deliveryCharges;
    sumFinalSellingPrice += st.finalSellingPrice;
    sumFinalOrderTotal += st.finalOrderTotal;
    sumDiscountedBasePrice += st.discountedBasePrice;
    sumPpd += st.ppd;
    sumCod += st.cod;
    sumMeals += st.meals;
    sumDeliveredOrders += st.deliveredOrdersCount;
    sumDeliveredOutlets += st.deliveredOutlets.size;
    sumStationVendors += st.stationAllOutlets.size;
  });

  const overallCheckPct =
    sumFinalBasePrice > 0
      ? Number(((sumFinalTotalComm / sumFinalBasePrice) * 100).toFixed(6))
      : 0;

  // Row 1: Top Summary Row (Red numbers as in screenshot)
  const topSummaryRow = [
    '', // Station Code placeholder
    '', // Rank
    '', // Delivery Date
    '', // Station Name
    Number(sumVendorPrice.toFixed(2)),
    Number(sumFinalBasePrice.toFixed(2)),
    Number(sumFinalTotalComm.toFixed(2)),
    Number(sumFinalIrctcComm.toFixed(2)),
    Number(sumFinalRfComm.toFixed(2)),
    Number(sumFinalGst.toFixed(2)),
    Number(sumFinalDiscount.toFixed(2)),
    Number(sumFinalVendorDiscount.toFixed(2)),
    Number(sumFinalRfDiscount.toFixed(2)),
    Number(sumDeliveryCharges.toFixed(2)),
    Number(sumFinalSellingPrice.toFixed(2)),
    Number(sumFinalOrderTotal.toFixed(2)),
    Number(sumDiscountedBasePrice.toFixed(2)),
    Number(sumPpd.toFixed(2)),
    Number(sumCod.toFixed(2)),
    sumMeals,
    sumCheckPctFormatted(overallCheckPct),
    sumDeliveredOrders,
    sumDeliveredOutlets,
    sumStationVendors,
  ];

  // Helper for check percentage
  function sumCheckPctFormatted(val: number) {
    return val;
  }

  // Row 2: Headers
  const headers = [
    'Station Code',
    'Rank',
    'Delivery Date',
    'Station Name',
    'Vendor Price',
    'Final Base Price',
    'Final Total Commission',
    'Final IRCTC Comm',
    'Final RF Commission',
    'Final GST',
    'Final Discount',
    'Final Vendor Discount',
    'Final RF Discount',
    'Delivery Charges',
    'Final Selling Price',
    'Final Order Total',
    'Discounted Base Price',
    'PPD',
    'COD',
    'Meals',
    'Check',
    'Count of Delivered Orders',
    'Count of Delivered Outlets',
    'Total Station Vendors',
  ];

  // Build rows data
  const rowsData: any[][] = sortedStations.map((st, index) => {
    const vPrice = Number(st.vendorPrice.toFixed(2));
    const bPrice = Number(st.finalBasePrice.toFixed(2));
    const tComm = Number(st.finalTotalComm.toFixed(2));
    const checkPct = bPrice > 0 ? Number((tComm / bPrice).toFixed(6)) : 0;

    return [
      st.stationCode,
      index + 1,
      lastDeliveryDateStr,
      st.stationName,
      vPrice,
      bPrice,
      tComm,
      Number(st.finalIrctcComm.toFixed(2)),
      Number(st.finalRfComm.toFixed(2)),
      Number(st.finalGst.toFixed(2)),
      Number(st.finalDiscount.toFixed(2)),
      Number(st.finalVendorDiscount.toFixed(2)),
      Number(st.finalRfDiscount.toFixed(2)),
      Number(st.deliveryCharges.toFixed(2)),
      Number(st.finalSellingPrice.toFixed(2)),
      Number(st.finalOrderTotal.toFixed(2)),
      Number(st.discountedBasePrice.toFixed(2)),
      Number(st.ppd.toFixed(2)),
      Number(st.cod.toFixed(2)),
      st.meals,
      checkPct,
      st.deliveredOrdersCount,
      st.deliveredOutlets.size,
      st.stationAllOutlets.size,
    ];
  });

  const fullSheetData = [topSummaryRow, headers, ...rowsData];
  const worksheet = XLSX.utils.aoa_to_sheet(fullSheetData);

  // Set column widths
  worksheet['!cols'] = [
    { wch: 14 }, // Station Code
    { wch: 8 },  // Rank
    { wch: 14 }, // Delivery Date
    { wch: 22 }, // Station Name
    { wch: 14 }, // Vendor Price
    { wch: 15 }, // Final Base Price
    { wch: 22 }, // Final Total Commission
    { wch: 16 }, // Final IRCTC Comm
    { wch: 18 }, // Final RF Commission
    { wch: 12 }, // Final GST
    { wch: 14 }, // Final Discount
    { wch: 19 }, // Final Vendor Discount
    { wch: 16 }, // Final RF Discount
    { wch: 15 }, // Delivery Charges
    { wch: 18 }, // Final Selling Price
    { wch: 16 }, // Final Order Total
    { wch: 20 }, // Discounted Base Price
    { wch: 14 }, // PPD
    { wch: 14 }, // COD
    { wch: 10 }, // Meals
    { wch: 12 }, // Check
    { wch: 24 }, // Count of Delivered Orders
    { wch: 26 }, // Count of Delivered Outlets
    { wch: 22 }, // Total Station Vendors
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Last Day Station Report');

  XLSX.writeFile(
    workbook,
    `LAST_DAY_STATION_REPORT_${lastDeliveryDateStr.replace(/\//g, '-')}.xlsx`
  );
};
