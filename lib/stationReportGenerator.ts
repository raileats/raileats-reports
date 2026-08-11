import * as XLSX from 'xlsx';
import { MasterOrderRow, OutletMasterInfo } from './vendorRdsGenerator';

export interface StationReportRow {
  'Station Code': string;
  'Rank': number;
  'Station Name': string;
  'Vendor Price': number;
  'Final Base Price': number;
  'Final Total Commission': number;
  'Final IRCTC Comm': number;
  'Final RF Commission': number;
  'Final GST': number;
  'Final Discount': number;
  'Final Vendor Discount': number;
  'Final RF Discount': number;
  'Delivery Charges': number;
  'Final Selling Price': number;
  'Final Order Total': number;
  'Discounted Base Price': number;
  'PPD': number;
  'COD': number;
  'Meals': number;
  'Check': string;
  'Count of Delivered Orders': number;
  'Count of Delivered Outlets': number;
  'Total Station Vendors': number;
}

/**
 * Generates Station Code Wise Summary Report matching exact screenshot format
 */
export const generateStationWiseData = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {}
): StationReportRow[] => {
  // 1. Group Master Orders Station Wise
  const stationOrdersMap: Record<string, MasterOrderRow[]> = {};
  
  // Total vendor counts per station from Master Outlets
  const totalStationVendorsMap: Record<string, Set<string>> = {};
  Object.values(outletsMasterInfo).forEach((out) => {
    const stCode = (out.station || '').trim().toUpperCase();
    const outId = String(out.outletId || '').trim();
    if (stCode && outId) {
      if (!totalStationVendorsMap[stCode]) {
        totalStationVendorsMap[stCode] = new Set();
      }
      totalStationVendorsMap[stCode].add(outId);
    }
  });

  masterOrders.forEach((row) => {
    const stCode = String(row['Station Code'] || '').trim().toUpperCase();
    if (!stCode) return;

    if (!stationOrdersMap[stCode]) {
      stationOrdersMap[stCode] = [];
    }
    stationOrdersMap[stCode].push(row);
  });

  const rawStationList: any[] = [];

  Object.entries(stationOrdersMap).forEach(([stationCode, orders]) => {
    let vendorPriceSum = 0;
    let basePriceSum = 0;
    let totalCommSum = 0;
    let irctcCommSum = 0;
    let rfCommSum = 0;
    let gstSum = 0;
    let totalDiscountSum = 0;
    let vendorDiscountSum = 0;
    let rfDiscountSum = 0;
    let deliveryChargesSum = 0;
    let sellingPriceSum = 0;
    let orderTotalSum = 0;
    let discountedBaseSum = 0;
    let ppdSum = 0;
    let codSum = 0;
    let mealsSum = 0;
    let deliveredOrdersCount = 0;

    const deliveredOutletSet = new Set<string>();
    let stationName = '';

    orders.forEach((ord) => {
      const status = String(ord['Final Status'] || '').trim().toLowerCase();
      const outletId = String(ord['Outlet ID'] || '').trim();

      // Station Name extraction from Outlet Master or row remarks
      if (!stationName) {
        const outInfo = outletsMasterInfo[outletId];
        if (outInfo && outInfo.station) {
          stationName = outInfo.station;
        } else if (ord['Station Code']) {
          stationName = ord['Station Code'];
        }
      }

      if (status === 'delivered') {
        vendorPriceSum += Number(ord['Final Vendor Price'] || 0);
        basePriceSum += Number(ord['Final Base Price'] || 0);
        totalCommSum += Number(ord['Final Total Commission'] || 0);
        irctcCommSum += Number(ord['Final IRCTC Commission'] || 0);
        rfCommSum += Number(ord['Final RF Commission'] || 0);
        gstSum += Number(ord['Final GST'] || 0);
        totalDiscountSum += Number(ord['Final Total Discount'] || 0);
        vendorDiscountSum += Number(ord['Final Vendor Discount'] || 0);
        rfDiscountSum += Number(ord['Final RF Discount'] || 0);
        deliveryChargesSum += Number(ord['Delivery Charges'] || 0);
        sellingPriceSum += Number(ord['Final Selling Price'] || 0);
        orderTotalSum += Number(ord['Final Order Total'] || 0);
        discountedBaseSum += Number(ord['Discounted Base Price'] || 0);
        ppdSum += Number(ord['PPD'] || 0);
        codSum += Number(ord['COD'] || 0);
        mealsSum += Number(ord['Meals'] || 0);
        deliveredOrdersCount += Number(ord['Orders Count'] || 1);

        if (outletId) {
          deliveredOutletSet.add(outletId);
        }
      }
    });

    const vendorPrice = Number(vendorPriceSum.toFixed(2));
    const finalBasePrice = Number(basePriceSum.toFixed(2));
    const totalCommission = Number(totalCommSum.toFixed(2));
    const irctcComm = Number(irctcCommSum.toFixed(2));
    const rfComm = Number(rfCommSum.toFixed(2));
    const gst = Number(gstSum.toFixed(2));
    const totalDiscount = Number(totalDiscountSum.toFixed(2));
    const vendorDiscount = Number(vendorDiscountSum.toFixed(2));
    const rfDiscount = Number(rfDiscountSum.toFixed(2));
    const deliveryCharges = Number(deliveryChargesSum.toFixed(2));
    const sellingPrice = Number(sellingPriceSum.toFixed(2));
    const orderTotal = Number(orderTotalSum.toFixed(2));
    const discountedBase = Number(discountedBaseSum.toFixed(2));
    const ppd = Number(ppdSum.toFixed(2));
    const cod = Number(codSum.toFixed(2));

    // Check % (RF Commission Margin / Check)
    const checkPct = finalBasePrice > 0 ? `${Math.round((totalCommission / finalBasePrice) * 100)}%` : '0%';

    const countOfDeliveredOutlets = deliveredOutletSet.size;
    const totalStationVendors = totalStationVendorsMap[stationCode]
      ? totalStationVendorsMap[stationCode].size
      : Math.max(countOfDeliveredOutlets, 1);

    rawStationList.push({
      stationCode,
      stationName: stationName || stationCode,
      vendorPrice,
      finalBasePrice,
      totalCommission,
      irctcComm,
      rfComm,
      gst,
      totalDiscount,
      vendorDiscount,
      rfDiscount,
      deliveryCharges,
      sellingPrice,
      orderTotal,
      discountedBase,
      ppd,
      cod,
      meals: mealsSum,
      checkPct,
      deliveredOrdersCount,
      countOfDeliveredOutlets,
      totalStationVendors,
    });
  });

  // 2. Sort by Highest Volume (Delivered Orders / Revenue)
  rawStationList.sort((a, b) => b.finalBasePrice - a.finalBasePrice);

  // 3. Map into exact Columns and Assign Rank
  return rawStationList.map((st, idx) => ({
    'Station Code': st.stationCode,
    'Rank': idx + 1,
    'Station Name': st.stationName,
    'Vendor Price': st.vendorPrice,
    'Final Base Price': st.finalBasePrice,
    'Final Total Commission': st.totalCommission,
    'Final IRCTC Comm': st.irctcComm,
    'Final RF Commission': st.rfComm,
    'Final GST': st.gst,
    'Final Discount': st.totalDiscount,
    'Final Vendor Discount': st.vendorDiscount,
    'Final RF Discount': st.rfDiscount,
    'Delivery Charges': st.deliveryCharges,
    'Final Selling Price': st.sellingPrice,
    'Final Order Total': st.orderTotal,
    'Discounted Base Price': st.discountedBase,
    'PPD': st.ppd,
    'COD': st.cod,
    'Meals': st.meals,
    'Check': st.checkPct,
    'Count of Delivered Orders': st.deliveredOrdersCount,
    'Count of Delivered Outlets': st.countOfDeliveredOutlets,
    'Total Station Vendors': st.totalStationVendors,
  }));
};

/**
 * Direct Excel Export Workbook Generator
 */
export const generateStationReportWorkbook = (
  masterOrders: MasterOrderRow[],
  outletsMasterInfo: Record<string, OutletMasterInfo> = {},
  fileNamePrefix: string = 'STATION_WISE_REPORT'
) => {
  const stationData = generateStationWiseData(masterOrders, outletsMasterInfo);
  if (stationData.length === 0) {
    alert('Station report export karne ke liye koi data available nahi hai.');
    return;
  }

  const worksheet = XLSX.utils.json_to_sheet(stationData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Station Wise Summary');

  const todayStr = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `${fileNamePrefix}_${todayStr}.xlsx`);
};
