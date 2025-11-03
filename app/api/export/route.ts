import { NextRequest } from 'next/server';
import ExcelJS from 'exceljs';

type Trip = {
  date: string; // ISO date
  month: string; // YYYY-MM
  vendor: string;
  vehicleId: string;
  truckType: string;
  origin: string;
  onTime: boolean;
  reasonForDelay: string;
  breakdown: boolean;
  delayMinutes: number;
};

export const dynamic = 'force-dynamic';

export async function GET(_req: NextRequest) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Agentic Dashboard';
  workbook.created = new Date();

  const dataSheet = workbook.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
  const listsSheet = workbook.addWorksheet('Lists');
  const dashboard = workbook.addWorksheet('Dashboard');

  dataSheet.columns = [
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Month', key: 'month', width: 10 },
    { header: 'Vendor', key: 'vendor', width: 16 },
    { header: 'VehicleId', key: 'vehicleId', width: 16 },
    { header: 'TruckType', key: 'truckType', width: 16 },
    { header: 'Origin', key: 'origin', width: 14 },
    { header: 'OnTime', key: 'onTime', width: 10 },
    { header: 'ReasonForDelay', key: 'reasonForDelay', width: 24 },
    { header: 'Breakdown', key: 'breakdown', width: 12 },
    { header: 'DelayMinutes', key: 'delayMinutes', width: 14 },
  ];

  const vendors = ['Acme Logistics', 'TransGo', 'RoadRunner', 'CargoX'];
  const truckTypes = ['Flatbed', 'Reefer', 'Box', 'Tanker'];
  const origins = ['NYC', 'LAX', 'DAL', 'ATL', 'SEA'];
  const reasons = ['Traffic', 'Weather', 'Mechanical', 'Route', 'Loading Delay', 'Other'];
  const months = generateRecentMonths(6);

  const trips = generateSampleData({ months, vendors, truckTypes, origins, reasons, rows: 400 });

  // Write data rows
  for (const t of trips) {
    dataSheet.addRow({
      date: t.date,
      month: t.month,
      vendor: t.vendor,
      vehicleId: t.vehicleId,
      truckType: t.truckType,
      origin: t.origin,
      onTime: t.onTime,
      reasonForDelay: t.reasonForDelay,
      breakdown: t.breakdown,
      delayMinutes: t.delayMinutes,
    });
  }

  // Table for structured references
  dataSheet.addTable({
    name: 'DataTable',
    ref: 'A1',
    headerRow: true,
    columns: dataSheet.columns.map((c) => ({ name: c.header as string })),
    rows: trips.map((t) => [
      t.date,
      t.month,
      t.vendor,
      t.vehicleId,
      t.truckType,
      t.origin,
      t.onTime,
      t.reasonForDelay,
      t.breakdown,
      t.delayMinutes,
    ]),
  });

  // Build Lists sheet with uniques and named ranges
  writeVerticalList(listsSheet, 'A', 'Months', unique(trips.map(t => t.month)));
  writeVerticalList(listsSheet, 'B', 'Vendors', unique(trips.map(t => t.vendor)));
  writeVerticalList(listsSheet, 'C', 'Reasons', unique(reasons));
  writeVerticalList(listsSheet, 'D', 'TruckTypes', unique(truckTypes));
  writeVerticalList(listsSheet, 'E', 'Origins', unique(origins));

  // Name the selection cells on Dashboard
  dashboard.getCell('B2').value = 'Selections';
  dashboard.getCell('B2').font = { bold: true, size: 12, color: { argb: 'FF0EA5E9' } };

  dashboard.getCell('B4').value = 'Month';
  dashboard.getCell('C4').dataValidation = { type: 'list', allowBlank: false, formulae: ['=Lists!$A$2:INDEX(Lists!$A:$A,COUNTA(Lists!$A:$A))'] };
  dashboard.getCell('B5').value = 'Vendor';
  dashboard.getCell('C5').dataValidation = { type: 'list', allowBlank: false, formulae: ['=Lists!$B$2:INDEX(Lists!$B:$B,COUNTA(Lists!$B:$B))'] };

  // Headline KPIs
  dashboard.getCell('B8').value = 'KPIs';
  dashboard.getCell('B8').font = { bold: true, size: 12 };

  // Helper named cells (Excel will treat names when user defines; here we just reference C4/C5 directly in formulas)
  const selMonth = 'Dashboard!$C$4';
  const selVendor = 'Dashboard!$C$5';

  // KPI formulas using structured references
  dashboard.getCell('B10').value = 'Total Trips';
  dashboard.getCell('C10').value = {
    formula: `=ROWS(FILTER(DataTable[Month], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})))`,
  };

  dashboard.getCell('B11').value = 'On-time Vehicles';
  dashboard.getCell('C11').value = {
    formula: `=ROWS(FILTER(DataTable[OnTime], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})*(DataTable[OnTime]=TRUE)))`,
  };

  dashboard.getCell('B12').value = 'Delayed Vehicles';
  dashboard.getCell('C12').value = {
    formula: `=ROWS(FILTER(DataTable[OnTime], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})*(DataTable[OnTime]=FALSE)))`,
  };

  dashboard.getCell('B13').value = 'On-time %';
  dashboard.getCell('C13').numFmt = '0.0%';
  dashboard.getCell('C13').value = {
    formula: `=IFERROR(C11/C10,0)`
  };

  dashboard.getCell('B14').value = 'Breakdowns';
  dashboard.getCell('C14').value = {
    formula: `=ROWS(FILTER(DataTable[Breakdown], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})*(DataTable[Breakdown]=TRUE)))`,
  };

  // Reasons for delay table
  dashboard.getCell('E8').value = 'Reasons for Delay';
  dashboard.getCell('E8').font = { bold: true };
  dashboard.getCell('E10').value = 'Reason';
  dashboard.getCell('F10').value = 'Count';

  const reasonStartRow = 11;
  const reasonsCount = listsSheet.getColumn('C').values.filter(v => !!v).length - 1;
  for (let i = 0; i < reasonsCount; i++) {
    const row = reasonStartRow + i;
    // reason name from Lists
    dashboard.getCell(`E${row}`).value = { formula: `=INDEX(Lists!$C:$C, ${i + 2})` };
    dashboard.getCell(`F${row}`).value = {
      formula: `=ROWS(FILTER(DataTable[ReasonForDelay], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})*(DataTable[ReasonForDelay]=E${row})))`
    };
  }

  // Truck type at origin frequency
  dashboard.getCell('H8').value = 'Truck Type at Origin';
  dashboard.getCell('H8').font = { bold: true };
  dashboard.getCell('H10').value = 'Truck Type';
  dashboard.getCell('I10').value = 'Count';
  const truckTypeCount = listsSheet.getColumn('D').values.filter(v => !!v).length - 1;
  for (let i = 0; i < truckTypeCount; i++) {
    const row = reasonStartRow + i;
    dashboard.getCell(`H${row}`).value = { formula: `=INDEX(Lists!$D:$D, ${i + 2})` };
    dashboard.getCell(`I${row}`).value = {
      formula: `=ROWS(FILTER(DataTable[TruckType], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor})*(DataTable[TruckType]=H${row})))`
    };
  }

  // Distinct truck numbers used at origin (list + count)
  dashboard.getCell('B17').value = 'Truck numbers used (distinct)';
  dashboard.getCell('B17').font = { bold: true };
  dashboard.getCell('B18').value = 'Count';
  dashboard.getCell('C18').value = {
    formula: `=ROWS(UNIQUE(FILTER(DataTable[VehicleId], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor}))))`
  };
  dashboard.getCell('B20').value = 'List';
  dashboard.getCell('B21').value = {
    formula: `=LET(x, UNIQUE(FILTER(DataTable[VehicleId], (DataTable[Month]=${selMonth})*(DataTable[Vendor]=${selVendor}))), IF(ROWS(x)>0, x, ""))`
  };

  // Styles
  applyDashboardStyles(dashboard);
  styleHeaderRow(dataSheet);

  // Finalize workbook and return
  const buffer = await workbook.xlsx.writeBuffer();
  return new Response(Buffer.from(buffer), {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="Logistics_Dashboard.xlsx"'
    }
  });
}

function generateRecentMonths(n: number): string[] {
  const arr: string[] = [];
  const today = new Date();
  for (let i = 0; i < n; i++) {
    const d = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() - i, 1));
    const m = `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`;
    arr.push(m);
  }
  return arr.reverse();
}

function generateSampleData(opts: { months: string[]; vendors: string[]; truckTypes: string[]; origins: string[]; reasons: string[]; rows: number }): Trip[] {
  const { months, vendors, truckTypes, origins, reasons, rows } = opts;
  const rand = (min: number, max: number) => Math.floor(Math.random() * (max - min + 1)) + min;
  const pick = <T,>(arr: T[]) => arr[rand(0, arr.length - 1)];

  const trips: Trip[] = [];
  let vehicleCounter = 1000;

  for (let i = 0; i < rows; i++) {
    const month = pick(months);
    const day = rand(1, 26);
    const date = `${month}-${String(day).padStart(2, '0')}`;
    const vendor = pick(vendors);
    const truckType = pick(truckTypes);
    const origin = pick(origins);

    // Determine on-time / delay
    const isDelayed = Math.random() < 0.25; // 25% delayed
    const breakdown = isDelayed && Math.random() < 0.15; // 15% of delayed are breakdowns
    const delayMinutes = isDelayed ? rand(10, 240) : 0;
    const reasonForDelay = isDelayed ? pick(reasons) : '';

    const vehicleId = `TRK-${vehicleCounter + rand(0, 300)}`;

    trips.push({
      date,
      month,
      vendor,
      vehicleId,
      truckType,
      origin,
      onTime: !isDelayed,
      reasonForDelay,
      breakdown,
      delayMinutes,
    });
  }

  return trips;
}

function writeVerticalList(sheet: ExcelJS.Worksheet, col: string, title: string, values: string[]) {
  sheet.getCell(`${col}1`).value = title;
  sheet.getCell(`${col}1`).font = { bold: true };
  values.forEach((v, idx) => {
    sheet.getCell(`${col}${idx + 2}`).value = v;
  });
}

function unique<T>(arr: T[]): T[] {
  return Array.from(new Set(arr));
}

function applyDashboardStyles(ws: ExcelJS.Worksheet) {
  // Column widths
  ws.getColumn('A').width = 2;
  ws.getColumn('B').width = 20;
  ws.getColumn('C').width = 24;
  ws.getColumn('D').width = 2;
  ws.getColumn('E').width = 22;
  ws.getColumn('F').width = 12;
  ws.getColumn('G').width = 2;
  ws.getColumn('H').width = 22;
  ws.getColumn('I').width = 12;

  // Title
  ws.getCell('B1').value = 'Logistics Operations Dashboard';
  ws.mergeCells('B1', 'I1');
  ws.getCell('B1').font = { bold: true, size: 18, color: { argb: 'FF0F172A' } };

  // Card-like KPI styles
  for (const r of [10, 11, 12, 13, 14]) {
    ws.getCell(`B${r}`).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    ws.getCell(`B${r}`).font = { bold: true };
    ws.getCell(`C${r}`).numFmt = r === 13 ? '0.0%' : '0';
    ws.getCell(`C${r}`).border = { top: { style: 'thin' }, bottom: { style: 'thin' } };
  }

  // Section headers background
  for (const cell of ['B2', 'B8', 'E8', 'H8', 'B17']) {
    ws.getCell(cell).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
  }
}

function styleHeaderRow(ws: ExcelJS.Worksheet) {
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { vertical: 'middle' };
  headerRow.height = 18;
}
