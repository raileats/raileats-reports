import * as XLSX from 'xlsx';

export const generateMainReportWorkbook = (masterData: any[]) => {
  if (!masterData || masterData.length === 0) {
    alert('Master Data uplabdh nahi hai!');
    return;
  }

  // Group by Delivery Date
  const dateGroups: { [date: string]: any[] } = {};

  masterData.forEach((row) => {
    let rawDate = (row['Delivery Date'] || row['Booking Date'] || 'Unknown').trim();
    if (rawDate.includes('T')) rawDate = rawDate.split('T')[0];
    if (rawDate.includes(' ')) rawDate = rawDate.split(' ')[0];
    if (!dateGroups[rawDate]) dateGroups[rawDate] = [];
    dateGroups[rawDate].push(row);
  });

  const sortedDates = Object.keys(dateGroups).sort((a, b) => new Date(a).getTime() - new Date(b).getTime());

  const sheetData: any[][] = [];

  // Track MTD metrics
  let mtdOrders = 0;
  let mtdMeals = 0;
  let mtdValue = 0;
  let mtdPrepaid = 0;
  let mtdDiscount = 0;
  let mtdRevenue = 0;
  let mtdComplaints = 0;
  let mtdFeedback = 0;
  let mtdUndelivered = 0;

  // Source-wise MTD tracking
  const sourceMtdMap: { [source: string]: { orders: number; meals: number; value: number; prepaid: number; discount: number; revenue: number } } = {
    'RELFood_IRCTC': { orders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
    'RELFood_WEBSITE': { orders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
    'REL_Food_App': { orders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
    'MakeMyTrip': { orders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
  };

  const sourcesList = ['RELFood_IRCTC', 'RELFood_WEBSITE', 'REL_Food_App', 'MakeMyTrip'];

  sortedDates.forEach((dateStr) => {
    const rows = dateGroups[dateStr];

    // Format Date Title (e.g. Saturday, 1 August, 2026)
    let formattedDateTitle = dateStr;
    try {
      const d = new Date(dateStr);
      if (!isNaN(d.getTime())) {
        formattedDateTitle = d.toLocaleDateString('en-US', {
          weekday: 'long',
          day: 'numeric',
          month: 'long',
          year: 'numeric',
        });
      }
    } catch (e) {
      // fallback
    }

    // 1. Red Date Banner Header
    sheetData.push([formattedDateTitle]);

    // 2. Main Metric Group Header
    sheetData.push([
      'Source',
      'ORDERS', '', '', '',
      'MEALS', '', '', '',
      'VALUE', '', '',
      'PREPAID', '', '', '',
      'DISCOUNT', '', '', '',
      'REVENUE', '', '', '',
      'Complaints', '', '', '',
      'Feedback', '', '', '',
      'IRCTC Undelivered', '', '', '',
      'Outlets'
    ]);

    // 3. Sub-headers
    sheetData.push([
      '',
      'FTD', 'MTD', 'LMTD', 'ASP', 'Del%',
      'FTD', 'MTD', 'LMTD', 'ASP', 'MPO',
      'FTD', 'MTD', 'LMTD',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      'FTD', 'MTD', 'LMTD', '%',
      ''
    ]);

    // Daily totals calculation
    let ftdOrders = 0;
    let ftdDeliveredOrders = 0;
    let ftdMeals = 0;
    let ftdValue = 0;
    let ftdPrepaid = 0;
    let ftdDiscount = 0;
    let ftdRevenue = 0;
    let ftdComplaints = 0;
    let ftdFeedback = 0;
    let ftdUndelivered = 0;

    const uniqueOutlets = new Set<string>();

    // Source breakdown map for this date
    const daySourceMap: { [key: string]: any } = {
      'RELFood_IRCTC': { orders: 0, delOrders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
      'RELFood_WEBSITE': { orders: 0, delOrders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
      'REL_Food_App': { orders: 0, delOrders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
      'MakeMyTrip': { orders: 0, delOrders: 0, meals: 0, value: 0, prepaid: 0, discount: 0, revenue: 0 },
    };

    rows.forEach((r) => {
      const orderId = String(r['IRCTC Order ID'] || '');
      const outletId = String(r['Outlet ID'] || '');
      if (outletId) uniqueOutlets.add(outletId);

      const status = String(r['Final Status'] || '').trim().toLowerCase();
      const val = parseFloat(r['Final Base Price'] || r['Final Selling Price'] || 0) || 0;
      const ppd = parseFloat(r['PPD'] || 0) || 0;
      const disc = parseFloat(r['Final Total Discount'] || 0) || 0;
      const rfComm = parseFloat(r['Final RF Commission'] || 0) || 0;
      const meal = parseInt(r['Meals'] || 1, 10) || 1;

      ftdOrders += 1;
      if (status === 'delivered') ftdDeliveredOrders += 1;
      if (status.includes('not') || status.includes('undelivered')) ftdUndelivered += 1;

      ftdMeals += meal;
      ftdValue += val;
      ftdPrepaid += ppd;
      ftdDiscount += disc;
      ftdRevenue += rfComm;

      if (r['Remarks'] && String(r['Remarks']).length > 2) ftdComplaints += 1;
      if (r['Rating'] && String(r['Rating']).trim() !== '') ftdFeedback += 1;

      // Classify Source
      let matchedSource = 'RELFood_IRCTC';
      const rawVendor = String(r['Vendor Name'] || '').toUpperCase();
      if (rawVendor.includes('WEBSITE') || orderId.startsWith('WEB')) {
        matchedSource = 'RELFood_WEBSITE';
      } else if (rawVendor.includes('APP') || orderId.startsWith('APP')) {
        matchedSource = 'REL_Food_App';
      } else if (rawVendor.includes('MMT') || rawVendor.includes('MAKEMYTRIP')) {
        matchedSource = 'MakeMyTrip';
      }

      const sm = daySourceMap[matchedSource];
      sm.orders += 1;
      if (status === 'delivered') sm.delOrders += 1;
      sm.meals += meal;
      sm.value += val;
      sm.prepaid += ppd;
      sm.discount += disc;
      sm.revenue += rfComm;
    });

    // Update Overall MTD
    mtdOrders += ftdOrders;
    mtdMeals += ftdMeals;
    mtdValue += ftdValue;
    mtdPrepaid += ftdPrepaid;
    mtdDiscount += ftdDiscount;
    mtdRevenue += ftdRevenue;
    mtdComplaints += ftdComplaints;
    mtdFeedback += ftdFeedback;
    mtdUndelivered += ftdUndelivered;

    const totalASP = ftdOrders > 0 ? Number((ftdValue / ftdOrders).toFixed(0)) : 0;
    const totalDelPct = ftdOrders > 0 ? `${((ftdDeliveredOrders / ftdOrders) * 100).toFixed(2)}%` : '0%';
    const totalMealASP = ftdMeals > 0 ? Number((ftdValue / ftdMeals).toFixed(0)) : 0;
    const totalMPO = ftdOrders > 0 ? Number((ftdMeals / ftdOrders).toFixed(2)) : 0;
    const totalPrepaidPct = ftdValue > 0 ? `${((ftdPrepaid / ftdValue) * 100).toFixed(2)}%` : '0%';
    const totalDiscPct = ftdValue > 0 ? `${((ftdDiscount / ftdValue) * 100).toFixed(2)}%` : '0%';
    const totalRevPct = ftdValue > 0 ? `${((ftdRevenue / ftdValue) * 100).toFixed(2)}%` : '0%';
    const totalUndelPct = ftdOrders > 0 ? `${((ftdUndelivered / ftdOrders) * 100).toFixed(2)}%` : '0%';

    // 4. Total Row
    sheetData.push([
      'Total',
      ftdOrders, mtdOrders, 0, totalASP, totalDelPct,
      ftdMeals, mtdMeals, 0, totalMealASP, totalMPO,
      Number(ftdValue.toFixed(0)), Number(mtdValue.toFixed(0)), 0,
      Number(ftdPrepaid.toFixed(0)), Number(mtdPrepaid.toFixed(0)), 0, totalPrepaidPct,
      Number(ftdDiscount.toFixed(0)), Number(mtdDiscount.toFixed(0)), 0, totalDiscPct,
      Number(ftdRevenue.toFixed(0)), Number(mtdRevenue.toFixed(0)), 0, totalRevPct,
      ftdComplaints, mtdComplaints, 0, '0.00%',
      ftdFeedback, mtdFeedback, 0, '0.00%',
      ftdUndelivered, mtdUndelivered, 0, totalUndelPct,
      uniqueOutlets.size
    ]);

    // 5. Source Breakdown Rows
    sourcesList.forEach((src) => {
      const sm = daySourceMap[src];
      const mtd = sourceMtdMap[src];

      mtd.orders += sm.orders;
      mtd.meals += sm.meals;
      mtd.value += sm.value;
      mtd.prepaid += sm.prepaid;
      mtd.discount += sm.discount;
      mtd.revenue += sm.revenue;

      const asp = sm.orders > 0 ? Number((sm.value / sm.orders).toFixed(0)) : 0;
      const delPct = sm.orders > 0 ? `${((sm.delOrders / sm.orders) * 100).toFixed(0)}%` : '0%';
      const mealAsp = sm.meals > 0 ? Number((sm.value / sm.meals).toFixed(0)) : 0;
      const mpo = sm.orders > 0 ? Number((sm.meals / sm.orders).toFixed(2)) : 0;
      const prepPct = sm.value > 0 ? `${((sm.prepaid / sm.value) * 100).toFixed(2)}%` : '0%';
      const discPct = sm.value > 0 ? `${((sm.discount / sm.value) * 100).toFixed(2)}%` : '0%';
      const revPct = sm.value > 0 ? `${((sm.revenue / sm.value) * 100).toFixed(2)}%` : '0%';

      sheetData.push([
        src,
        sm.orders, mtd.orders, 0, asp, delPct,
        sm.meals, mtd.meals, 0, mealAsp, mpo,
        Number(sm.value.toFixed(0)), Number(mtd.value.toFixed(0)), 0,
        Number(sm.prepaid.toFixed(0)), Number(mtd.prepaid.toFixed(0)), 0, prepPct,
        Number(sm.discount.toFixed(0)), Number(mtd.discount.toFixed(0)), 0, discPct,
        Number(sm.revenue.toFixed(0)), Number(mtd.revenue.toFixed(0)), 0, revPct,
        0, 0, 0, '0.00%',
        0, 0, 0, '0.00%',
        0, 0, 0, '0.00%',
        ''
      ]);
    });

    // Blank row spacer
    sheetData.push([]);
  });

  const ws = XLSX.utils.aoa_to_sheet(sheetData);

  // Set column widths
  ws['!cols'] = [
    { wch: 18 }, // Source
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 9 }, // Orders
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, // Meals
    { wch: 10 }, { wch: 10 }, { wch: 8 }, // Value
    { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 9 }, // Prepaid
    { wch: 9 }, { wch: 9 }, { wch: 8 }, { wch: 8 }, // Discount
    { wch: 9 }, { wch: 9 }, { wch: 8 }, { wch: 8 }, // Revenue
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, // Complaints
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, // Feedback
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, // IRCTC Undelivered
    { wch: 8 }, // Outlets
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Main Report');

  const today = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `MAIN_EXECUTIVE_REPORT_${today}.xlsx`);
};
