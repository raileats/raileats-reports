import * as XLSX from 'xlsx';

// Universal date normalizer: converts any date string to standard format "D/M/YYYY"
export const normalizeToDeliveryDate = (dateVal: any): string => {
  if (!dateVal) return '';
  const str = String(dateVal).trim();

  // Agar format "YYYY-MM-DD" ya "YYYY-MM-DD HH:mm:ss" ho
  if (str.includes('-')) {
    const parts = str.split(' ')[0].split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD
        const year = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const day = parseInt(parts[2], 10);
        return `${day}/${month}/${year}`;
      } else {
        // DD-MM-YYYY
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        return `${day}/${month}/${year}`;
      }
    }
  }

  // Agar format "DD/MM/YYYY" ya "D/M/YYYY" ho
  if (str.includes('/')) {
    const parts = str.split(' ')[0].split('/');
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        return `${day}/${month}/${year}`;
      }
    }
  }

  return str;
};

export const generateDateWiseReportWorkbook = (masterData: any[]) => {
  if (!masterData || masterData.length === 0) {
    alert('Master Data khali hai! Pehle reports process karein.');
    return;
  }

  // Sirf Delivered orders par aggregate karein (screenshot metrics ke according)
  const deliveredRows = masterData.filter(
    (row) => String(row['Final Status'] || '').trim().toLowerCase() === 'delivered'
  );

  // Determine Month & Year from data
  let detectedMonth = 8;
  let detectedYear = 2026;
  for (const r of deliveredRows) {
    const dStr = normalizeToDeliveryDate(r['Delivery Date'] || r['Booking Date'] || '');
    if (dStr.includes('/')) {
      const parts = dStr.split('/');
      if (parts.length === 3) {
        detectedMonth = parseInt(parts[1], 10) || 8;
        detectedYear = parseInt(parts[2], 10) || 2026;
        break;
      }
    }
  }

  const daysInMonth = new Date(detectedYear, detectedMonth, 0).getDate();

  // Group by Normalized Delivery Date
  const dateMap: Record<
    string,
    {
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
      ordersCount: number;
      uniqueOutlets: Set<string>;
    }
  > = {};

  deliveredRows.forEach((r) => {
    const dKey = normalizeToDeliveryDate(r['Delivery Date'] || r['Booking Date'] || '');
    if (!dKey) return;

    if (!dateMap[dKey]) {
      dateMap[dKey] = {
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
        ordersCount: 0,
        uniqueOutlets: new Set<string>(),
      };
    }

    const item = dateMap[dKey];
    item.vendorPrice += parseFloat(r['Final Vendor Price'] || 0) || 0;
    item.finalBasePrice += parseFloat(r['Final Base Price'] || 0) || 0;
    item.finalTotalComm += parseFloat(r['Final Total Commission'] || 0) || 0;
    item.finalIrctcComm += parseFloat(r['Final IRCTC Commission'] || 0) || 0;
    item.finalRfComm += parseFloat(r['Final RF Commission'] || 0) || 0;
    item.finalGst += parseFloat(r['Final GST'] || 0) || 0;
    item.finalDiscount += parseFloat(r['Final Total Discount'] || r['Discount'] || 0) || 0;
    item.finalVendorDiscount += parseFloat(r['Final Vendor Discount'] || 0) || 0;
    item.finalRfDiscount += parseFloat(r['Final RF Discount'] || 0) || 0;
    item.deliveryCharges += parseFloat(r['Delivery Charges'] || 0) || 0;
    item.finalSellingPrice += parseFloat(r['Final Selling Price'] || 0) || 0;
    item.finalOrderTotal += parseFloat(r['Final Order Total'] || 0) || 0;
    item.discountedBasePrice += parseFloat(r['Discounted Base Price'] || 0) || 0;
    item.ppd += parseFloat(r['PPD'] || 0) || 0;
    item.cod += parseFloat(r['COD'] || 0) || 0;
    item.meals += parseInt(r['Meals'] || '1', 10) || 1;
    item.ordersCount += 1;

    const outId = String(r['Outlet ID'] || r['Outlet Id'] || '').trim();
    if (outId) item.uniqueOutlets.add(outId);
  });

  // Calculate Month Totals for Top Summary Row
  let totalVendorPrice = 0;
  let totalFinalBasePrice = 0;
  let totalFinalTotalComm = 0;
  let totalFinalIrctcComm = 0;
  let totalFinalRfComm = 0;
  let totalFinalGst = 0;
  let totalFinalDiscount = 0;
  let totalFinalVendorDiscount = 0;
  let totalFinalRfDiscount = 0;
  let totalDeliveryCharges = 0;
  let totalFinalSellingPrice = 0;
  let totalFinalOrderTotal = 0;
  let totalDiscountedBasePrice = 0;
  let totalPpd = 0;
  let totalCod = 0;
  let totalMeals = 0;
  let totalOrdersCount = 0;
  const monthAllOutlets = new Set<string>();

  // Build Day-wise rows 1 to daysInMonth (e.g. 1/8/2026 to 31/8/2026)
  const rowsData: any[][] = [];

  for (let day = 1; day <= daysInMonth; day++) {
    const dateStr = `${day}/${detectedMonth}/${detectedYear}`;
    const d = dateMap[dateStr];

    if (d && d.ordersCount > 0) {
      const vPrice = Number(d.vendorPrice.toFixed(2));
      const bPrice = Number(d.finalBasePrice.toFixed(2));
      const tComm = Number(d.finalTotalComm.toFixed(2));
      const irctcComm = Number(d.finalIrctcComm.toFixed(2));
      const rfComm = Number(d.finalRfComm.toFixed(2));
      const gst = Number(d.finalGst.toFixed(2));
      const disc = Number(d.finalDiscount.toFixed(2));
      const vDisc = Number(d.finalVendorDiscount.toFixed(2));
      const rfDisc = Number(d.finalRfDiscount.toFixed(2));
      const deliv = Number(d.deliveryCharges.toFixed(2));
      const sellPrice = Number(d.finalSellingPrice.toFixed(2));
      const ordTotal = Number(d.finalOrderTotal.toFixed(2));
      const discBase = Number(d.discountedBasePrice.toFixed(2));
      const ppdVal = Number(d.ppd.toFixed(2));
      const codVal = Number(d.cod.toFixed(2));
      const checkPct = bPrice > 0 ? Number(((tComm / bPrice) * 100).toFixed(2)) : 0;
      const outletCount = d.uniqueOutlets.size;

      totalVendorPrice += vPrice;
      totalFinalBasePrice += bPrice;
      totalFinalTotalComm += tComm;
      totalFinalIrctcComm += irctcComm;
      totalFinalRfComm += rfComm;
      totalFinalGst += gst;
      totalFinalDiscount += disc;
      totalFinalVendorDiscount += vDisc;
      totalFinalRfDiscount += rfDisc;
      totalDeliveryCharges += deliv;
      totalFinalSellingPrice += sellPrice;
      totalFinalOrderTotal += ordTotal;
      totalDiscountedBasePrice += discBase;
      totalPpd += ppdVal;
      totalCod += codVal;
      totalMeals += d.meals;
      totalOrdersCount += d.ordersCount;
      d.uniqueOutlets.forEach((id) => monthAllOutlets.add(id));

      rowsData.push([
        dateStr,
        vPrice,
        bPrice,
        tComm,
        irctcComm,
        rfComm,
        gst,
        disc,
        vDisc,
        rfDisc,
        deliv,
        sellPrice,
        ordTotal,
        discBase,
        ppdVal,
        codVal,
        d.meals,
        `${checkPct.toFixed(2)}%`,
        d.ordersCount,
        outletCount,
      ]);
    } else {
      // Empty date row with 0s exactly as shown in screenshot
      rowsData.push([
        dateStr,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        '0.00%',
        0,
        0,
      ]);
    }
  }

  const overallCheckPct =
    totalFinalBasePrice > 0
      ? Number(((totalFinalTotalComm / totalFinalBasePrice) * 100).toFixed(2))
      : 0;

  // Row 1: Top Summary Row (Red numbers)
  const topSummaryRow = [
    '',
    Number(totalVendorPrice.toFixed(2)),
    Number(totalFinalBasePrice.toFixed(2)),
    Number(totalFinalTotalComm.toFixed(2)),
    Number(totalFinalIrctcComm.toFixed(2)),
    Number(totalFinalRfComm.toFixed(2)),
    Number(totalFinalGst.toFixed(2)),
    Number(totalFinalDiscount.toFixed(2)),
    Number(totalFinalVendorDiscount.toFixed(2)),
    Number(totalFinalRfDiscount.toFixed(2)),
    Number(totalDeliveryCharges.toFixed(2)),
    Number(totalFinalSellingPrice.toFixed(2)),
    Number(totalFinalOrderTotal.toFixed(2)),
    Number(totalDiscountedBasePrice.toFixed(2)),
    Number(totalPpd.toFixed(2)),
    Number(totalCod.toFixed(2)),
    totalMeals,
    `${overallCheckPct.toFixed(2)}%`,
    totalOrdersCount,
    monthAllOutlets.size,
  ];

  // Row 2: Table Column Headers (Matching image layout)
  const headers = [
    'Delivery Date',
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
    'Delivered Orders Outlet Count',
  ];

  const fullSheetData = [topSummaryRow, headers, ...rowsData];
  const worksheet = XLSX.utils.aoa_to_sheet(fullSheetData);

  // Column auto-width styling
  worksheet['!cols'] = [
    { wch: 14 }, // Delivery Date
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
    { wch: 10 }, // Check
    { wch: 24 }, // Count of Delivered Orders
    { wch: 26 }, // Delivered Orders Outlet Count
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Date Wise Summary');

  XLSX.writeFile(
    workbook,
    `DATE_WISE_SUMMARY_REPORT_${detectedMonth}_${detectedYear}.xlsx`
  );
};
