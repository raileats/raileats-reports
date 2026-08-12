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
  'Not Delivered Order': number;
  'Not Delivered %': string;
  'PPD % of Final Selling Price': string;
  'Feedback Good': number;
  'Feedback Bad': number;
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
  Object.values(outletsMasterInfo || {}).forEach((out: any) => {
    const stCode = String(
      out?.station || out?.stationCode || out?.stn_code || out?.deliveryStation || out?.stationName || ''
    ).trim().toUpperCase();
    const outId = String(out?.outletId || out?.restaurantId || out?.outlet_id || '').trim();
    if (stCode && outId) {
      if (!totalStationVendorsMap[stCode]) {
        totalStationVendorsMap[stCode] = new Set();
      }
      totalStationVendorsMap[stCode].add(outId);
    }
  });

  masterOrders.forEach((row: any) => {
    const stCode = String(
      row['Station Code'] ||
      row['StationCode'] ||
      row['Delivery Station'] ||
      row['DeliveryStation'] ||
      row['Station Name'] ||
      row['Station'] ||
      ''
    ).trim().toUpperCase();

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
    let notDeliveredOrdersCount = 0;
    let feedbackGoodCount = 0;
    let feedbackBadCount = 0;

    const deliveredOutletSet = new Set<string>();
    let stationName = '';

    orders.forEach((ord: any) => {
      const finalStatus = String(ord['Final Status'] || ord['final_status'] || ord['Delivery Status'] || ord['Status'] || '').trim().toLowerCase();
      const irctcStatus = String(ord['IRCTC Status'] || ord['irctc_status'] || ord['Order Status'] || '').trim().toLowerCase();
      const rfStatus = String(ord['RF Status'] || ord['rf_status'] || '').trim().toLowerCase();
      const outletId = String(ord['Outlet ID'] || ord['OutletId'] || ord['outlet_id'] || '').trim();

      // Station Name extraction from Delivery Station / Outlet Master / row
      if (!stationName) {
        const outInfo: any = outletsMasterInfo[outletId];
        if (ord['Delivery Station']) {
          stationName = String(ord['Delivery Station']).trim();
        } else if (ord['Station Name']) {
          stationName = String(ord['Station Name']).trim();
        } else if (outInfo && outInfo.station) {
          stationName = String(outInfo.station).trim();
        } else if (ord['Station Code']) {
          stationName = String(ord['Station Code']).trim();
        }
      }

      // Feedback & Complaint Check
      const feedbackVal = String(ord['Feedback'] || ord['feedback'] || ord['Customer Feedback'] || '').trim();
      const complaintVal = String(ord['Complaint'] || ord['complaint'] || ord['Complaint Comment'] || ord['Comment'] || '').trim();

      if (feedbackVal.length > 0) {
        feedbackGoodCount += 1;
      }
      if (complaintVal.length > 0) {
        feedbackBadCount += 1;
      }

      // Check Status
      const isDelivered = finalStatus === 'delivered' || finalStatus === 'success';
      const isNotDelivered =
        finalStatus === 'not delivered' ||
        finalStatus === 'cancelled' ||
        finalStatus === 'undelivered' ||
        irctcStatus.includes('undelivered') ||
        irctcStatus.includes('cancel') ||
        rfStatus.includes('undelivered') ||
        rfStatus.includes('cancel');

      if (isNotDelivered) {
        notDeliveredOrdersCount += 1;
      }

      if (isDelivered) {
        vendorPriceSum += Number(ord['Final Vendor Price'] || ord['Vendor Price'] || 0);
        basePriceSum += Number(ord['Final Base Price'] || ord['Base Price'] || 0);
        totalCommSum += Number(ord['Final Total Commission'] || ord['Total Commission'] || 0);
        irctcCommSum += Number(ord['Final IRCTC Commission'] || ord['IRCTC Comm'] || ord['Vendor Comm'] || 0);
        rfCommSum += Number(ord['Final RF Commission'] || ord['RF Commission'] || 0);
        gstSum += Number(ord['Final GST'] || ord['Total Gst'] || ord['GST'] || 0);
        totalDiscountSum += Number(ord['Final Total Discount'] || ord['Discount'] || ord['discount'] || 0);
        vendorDiscountSum += Number(ord['Final Vendor Discount'] || ord['Vendor Discount'] || 0);
        rfDiscountSum += Number(ord['Final RF Discount'] || ord['RF Discount'] || 0);
        deliveryChargesSum += Number(ord['Delivery Charges'] || ord['Delivery Charge'] || 0);
        sellingPriceSum += Number(ord['Final Selling Price'] || ord['Selling Price'] || 0);
        orderTotalSum += Number(ord['Final Order Total'] || ord['Order Total'] || 0);
        discountedBaseSum += Number(ord['Discounted Base Price'] || 0);
        ppdSum += Number(ord['PPD'] || ord['ppd'] || 0);
        codSum += Number(ord['COD'] || ord['cod'] || 0);
        mealsSum += Number(ord['Meals'] || ord['Meal Count'] || 1);
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

    // Percentages
    const checkPct = finalBasePrice > 0 ? `${((totalCommission / finalBasePrice) * 100).toFixed(2)}%` : '0.00%';
    const notDeliveredPct = deliveredOrdersCount > 0 
      ? `${((notDeliveredOrdersCount / deliveredOrdersCount) * 100).toFixed(2)}%`
      : notDeliveredOrdersCount > 0 ? '100.00%' : '0.00%';
    const ppdPct = sellingPrice > 0 
      ? `${((ppd / sellingPrice) * 100).toFixed(2)}%` 
      : '0.00%';

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
      notDeliveredOrdersCount,
      notDeliveredPct,
      ppdPct,
      feedbackGoodCount,
      feedbackBadCount,
      countOfDeliveredOutlets,
      totalStationVendors,
    });
  });

  // 2. Sort by Highest Volume (Base Price)
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
    'Not Delivered Order': st.notDeliveredOrdersCount,
    'Not Delivered %': st.notDeliveredPct,
    'PPD % of Final Selling Price': st.ppdPct,
    'Feedback Good': st.feedbackGoodCount,
    'Feedback Bad': st.feedbackBadCount,
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
