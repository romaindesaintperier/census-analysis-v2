import ExcelJS from 'exceljs';
import { AnalysisData } from '@/types/employee';
import { formatCurrency, formatPercent, defaultBenchmarks } from './analysis';

export async function exportToExcel(data: AnalysisData): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Org Due Diligence Tool';
  workbook.created = new Date();

  // 1. Summary Sheet
  addSummarySheet(workbook, data);

  // 2. Headcount Sheet
  addHeadcountSheet(workbook, data);

  // 3. Org Chart Sheet
  addOrgChartSheet(workbook, data);

  // 4. Spans & Layers Sheet
  addSpansLayersSheet(workbook, data);

  // 5. Tenure Sheet
  addTenureSheet(workbook, data);

  // 6. Automation Sheet
  addAutomationSheet(workbook, data);

  // 7. Offshoring Sheet
  addOffshoringSheet(workbook, data);

  // 8. Compensation Sheet
  addCompensationSheet(workbook, data);

  // 9. Complete Employee Data Sheet (all employees, all fields)
  addCompleteEmployeeDataSheet(workbook, data);

  // 10. All Managers Sheet (complete span data)
  addAllManagersSheet(workbook, data);

  // Generate and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `org-analysis-${new Date().toISOString().split('T')[0]}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
}

function styleHeader(row: ExcelJS.Row) {
  row.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4A5568' } };
  row.alignment = { horizontal: 'center', vertical: 'middle' };
  row.height = 24;
}

function addSummarySheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Summary');
  const { totals, quickWins, functionStats, employees } = data;

  // Calculate average tenure
  const now = new Date();
  const avgTenure = employees.reduce((sum, emp) => {
    const hireDate = new Date(emp.hireDate);
    return sum + (now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24 * 365);
  }, 0) / employees.length;

  // Key Metrics Section
  sheet.addRow(['KEY METRICS']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  sheet.addRow(['Metric', 'Value']);
  styleHeader(sheet.getRow(3));

  sheet.addRow(['Total Headcount', totals.headcount]);
  sheet.addRow(['Total FLRR', formatCurrency(totals.totalFLRR)]);
  sheet.addRow(['Average FLRR', formatCurrency(totals.avgFLRR)]);
  sheet.addRow(['Org Layers', totals.layers]);
  sheet.addRow(['Average Span of Control', totals.avgSpan.toFixed(1)]);
  sheet.addRow(['Manager:IC Ratio', `1:${totals.managerToICRatio > 0 ? (1 / totals.managerToICRatio).toFixed(1) : '0'}`]);
  sheet.addRow(['Average Tenure (years)', avgTenure.toFixed(1)]);
  sheet.addRow(['Avg Variable Compensation', formatPercent(totals.avgVariablePercent)]);
  sheet.addRow(['Total Managers', totals.totalManagers]);
  sheet.addRow(['Manager Percentage', formatPercent(totals.managerPercent)]);
  sheet.addRow(['CEO Direct Reports', totals.ceoDirectReports]);

  sheet.addRow([]);
  sheet.addRow([]);

  // Quick Wins Section
  sheet.addRow(['VALUE CREATION OPPORTUNITIES']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const qwHeaderRow = sheet.addRow(['Title', 'Category', 'Impact', 'Description', 'Metric']);
  styleHeader(qwHeaderRow);

  quickWins.forEach(qw => {
    sheet.addRow([qw.title, qw.category, qw.impact, qw.description, qw.metric || '']);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Function Summary Section
  sheet.addRow(['FUNCTION SUMMARY']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const funcHeaderRow = sheet.addRow(['Function', 'Headcount', 'Total FLRR', 'Avg FLRR', 'Best-Cost %']);
  styleHeader(funcHeaderRow);

  functionStats.forEach(func => {
    sheet.addRow([
      func.function,
      func.headcount,
      formatCurrency(func.totalFLRR),
      formatCurrency(func.avgFLRR),
      formatPercent(func.bestCostPercent)
    ]);
  });

  // Set column widths
  sheet.columns = [
    { width: 30 },
    { width: 15 },
    { width: 15 },
    { width: 50 },
    { width: 20 }
  ];
}

function addHeadcountSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Headcount');
  const { employees, functionStats } = data;

  // Function Breakdown
  sheet.addRow(['FUNCTION BREAKDOWN']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const funcHeaderRow = sheet.addRow(['Function', 'Headcount', 'Total FLRR', 'Avg FLRR']);
  styleHeader(funcHeaderRow);

  functionStats.forEach(func => {
    sheet.addRow([func.function, func.headcount, formatCurrency(func.totalFLRR), formatCurrency(func.avgFLRR)]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Location Breakdown
  const locationStats = Array.from(
    employees.reduce((map, emp) => {
      const loc = emp.location;
      const existing = map.get(loc) || { location: loc, headcount: 0, flrr: 0 };
      existing.headcount++;
      existing.flrr += emp.flrr;
      map.set(loc, existing);
      return map;
    }, new Map<string, { location: string; headcount: number; flrr: number }>())
  ).map(([, v]) => v).sort((a, b) => b.headcount - a.headcount);

  sheet.addRow(['LOCATION BREAKDOWN']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const locHeaderRow = sheet.addRow(['Location', 'Headcount', 'Total FLRR']);
  styleHeader(locHeaderRow);

  locationStats.forEach(loc => {
    sheet.addRow([loc.location, loc.headcount, formatCurrency(loc.flrr)]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Business Unit Breakdown
  const businessUnitStats = Array.from(
    employees.reduce((map, emp) => {
      const bu = emp.businessUnit || 'Unknown';
      const existing = map.get(bu) || { businessUnit: bu, headcount: 0, flrr: 0 };
      existing.headcount++;
      existing.flrr += emp.flrr;
      map.set(bu, existing);
      return map;
    }, new Map<string, { businessUnit: string; headcount: number; flrr: number }>())
  ).map(([, v]) => v).sort((a, b) => b.headcount - a.headcount);

  sheet.addRow(['BUSINESS UNIT BREAKDOWN']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const buHeaderRow = sheet.addRow(['Business Unit', 'Headcount', 'Total FLRR']);
  styleHeader(buHeaderRow);

  businessUnitStats.forEach(bu => {
    sheet.addRow([bu.businessUnit, bu.headcount, formatCurrency(bu.flrr)]);
  });

  sheet.columns = [
    { width: 30 },
    { width: 15 },
    { width: 18 },
    { width: 18 }
  ];
}

function addOrgChartSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Org Chart');
  const { employees, layerStats, totals } = data;

  // Org Structure Summary
  sheet.addRow(['ORG STRUCTURE SUMMARY']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  sheet.addRow(['Metric', 'Value']);
  styleHeader(sheet.getRow(3));

  sheet.addRow(['Total Layers', totals.layers]);
  sheet.addRow(['Total Employees', totals.headcount]);
  sheet.addRow(['Total Managers', totals.totalManagers]);
  sheet.addRow(['CEO Direct Reports', totals.ceoDirectReports]);
  sheet.addRow(['Average Span of Control', totals.avgSpan.toFixed(1)]);

  sheet.addRow([]);
  sheet.addRow([]);

  // Layer Stats
  sheet.addRow(['LAYER BREAKDOWN']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const layerHeaderRow = sheet.addRow(['Layer', 'Headcount', 'Managers', 'ICs', 'Total FLRR', 'Avg FLRR', 'Avg Tenure (years)']);
  styleHeader(layerHeaderRow);

  layerStats.forEach(layer => {
    sheet.addRow([
      `Layer ${layer.layer}`,
      layer.headcount,
      layer.managers,
      layer.ics,
      formatCurrency(layer.totalFLRR),
      formatCurrency(layer.avgFLRR),
      layer.avgTenure.toFixed(1)
    ]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Employee Hierarchy List
  sheet.addRow(['EMPLOYEE LIST']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const empHeaderRow = sheet.addRow(['Employee ID', 'Title', 'Function', 'Manager ID', 'Location', 'Country', 'FLRR']);
  styleHeader(empHeaderRow);

  employees.forEach(emp => {
    sheet.addRow([emp.employeeId, emp.title, emp.function, emp.managerId || 'None', emp.location, emp.country, formatCurrency(emp.flrr)]);
  });

  sheet.columns = [
    { width: 15 },
    { width: 25 },
    { width: 15 },
    { width: 15 },
    { width: 20 },
    { width: 15 },
    { width: 15 }
  ];
}

function addSpansLayersSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Spans & Layers');
  const { layerStats, spanStats, functionSpanStats, totals } = data;
  const benchmarks = defaultBenchmarks;

  // Key Metrics
  sheet.addRow(['SPANS & LAYERS SUMMARY']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  sheet.addRow(['Metric', 'Value', 'Benchmark']);
  styleHeader(sheet.getRow(3));

  sheet.addRow(['Total Managers', totals.totalManagers, '']);
  sheet.addRow(['Manager %', formatPercent(totals.managerPercent), '']);
  sheet.addRow(['CEO Direct Reports', totals.ceoDirectReports, '']);
  sheet.addRow(['Org Layers', totals.layers, `≤ ${benchmarks.maxLayers}`]);
  sheet.addRow(['Average Span of Control', totals.avgSpan.toFixed(1), `${benchmarks.minSpan}-${benchmarks.maxSpan}`]);

  sheet.addRow([]);
  sheet.addRow([]);

  // Span Distribution
  const spanDistribution = [
    { range: '1', count: spanStats.filter(s => s.directReports === 1).length },
    { range: '2-4', count: spanStats.filter(s => s.directReports >= 2 && s.directReports <= 4).length },
    { range: '5-7', count: spanStats.filter(s => s.directReports >= 5 && s.directReports <= 7).length },
    { range: '8-10', count: spanStats.filter(s => s.directReports >= 8 && s.directReports <= 10).length },
    { range: '11+', count: spanStats.filter(s => s.directReports >= 11).length },
  ];

  sheet.addRow(['SPAN DISTRIBUTION']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const spanDistHeaderRow = sheet.addRow(['Span Range', 'Manager Count']);
  styleHeader(spanDistHeaderRow);

  spanDistribution.forEach(s => {
    sheet.addRow([s.range, s.count]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Layer Stats
  sheet.addRow(['LAYER BREAKDOWN']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const layerHeaderRow = sheet.addRow(['Layer', 'Headcount', 'Managers', 'ICs', 'Total FLRR', 'Avg Tenure (years)']);
  styleHeader(layerHeaderRow);

  layerStats.forEach(layer => {
    sheet.addRow([
      `Layer ${layer.layer}`,
      layer.headcount,
      layer.managers,
      layer.ics,
      formatCurrency(layer.totalFLRR),
      layer.avgTenure.toFixed(1)
    ]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Function Span Stats
  sheet.addRow(['SPAN OF CONTROL BY FUNCTION']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const funcSpanHeaderRow = sheet.addRow(['Function', 'Employees', 'Managers', 'Manager %', 'Avg Span', 'Layers']);
  styleHeader(funcSpanHeaderRow);

  functionSpanStats.forEach(func => {
    sheet.addRow([
      func.function,
      func.totalEmployees,
      func.managerCount,
      formatPercent(func.managerPercent),
      func.avgSpan.toFixed(1),
      func.layers
    ]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Single-Report Managers List
  const singleReportManagers = spanStats.filter(s => s.directReports === 1);
  if (singleReportManagers.length > 0) {
    sheet.addRow(['SINGLE-REPORT MANAGERS']);
    sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
    sheet.addRow([]);

    const singleHeaderRow = sheet.addRow(['Manager ID', 'Function', 'Layer', 'Direct Reports']);
    styleHeader(singleHeaderRow);

    singleReportManagers.forEach(mgr => {
      sheet.addRow([mgr.managerId, mgr.function, mgr.layer, mgr.directReports]);
    });
  }

  sheet.columns = [
    { width: 20 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 18 },
    { width: 15 }
  ];
}

function addTenureSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Tenure');
  const { employees, tenureBands, functionStats } = data;
  const now = new Date();

  // Calculate tenure for each employee
  const employeesWithTenure = employees.map(emp => {
    const hireDate = new Date(emp.hireDate);
    const years = (now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24 * 365);
    return { ...emp, tenureYears: years };
  });

  const avgTenure = employeesWithTenure.reduce((sum, e) => sum + e.tenureYears, 0) / employees.length;
  const recentJoiners = employeesWithTenure.filter(e => e.tenureYears < 1);
  const veterans = employeesWithTenure.filter(e => e.tenureYears >= 5);

  // Summary
  sheet.addRow(['TENURE SUMMARY']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  sheet.addRow(['Metric', 'Value']);
  styleHeader(sheet.getRow(3));

  sheet.addRow(['Average Tenure (years)', avgTenure.toFixed(1)]);
  sheet.addRow(['Recent Joiners (<1 year)', recentJoiners.length]);
  sheet.addRow(['Veterans (5+ years)', veterans.length]);

  sheet.addRow([]);
  sheet.addRow([]);

  // Tenure Distribution
  sheet.addRow(['TENURE DISTRIBUTION']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const tenureHeaderRow = sheet.addRow(['Tenure Band', 'Headcount', 'Total FLRR']);
  styleHeader(tenureHeaderRow);

  tenureBands.forEach(band => {
    sheet.addRow([band.band, band.headcount, formatCurrency(band.totalFLRR)]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Tenure by Function
  const tenureByFunction = functionStats.map(func => {
    const funcEmps = employeesWithTenure.filter(e => e.function === func.function);
    const avgFuncTenure = funcEmps.reduce((sum, e) => sum + e.tenureYears, 0) / funcEmps.length;
    return {
      function: func.function,
      avgTenure: avgFuncTenure,
      headcount: func.headcount,
    };
  }).sort((a, b) => b.avgTenure - a.avgTenure);

  sheet.addRow(['AVERAGE TENURE BY FUNCTION']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const funcTenureHeaderRow = sheet.addRow(['Function', 'Avg Tenure (years)', 'Headcount']);
  styleHeader(funcTenureHeaderRow);

  tenureByFunction.forEach(func => {
    sheet.addRow([func.function, func.avgTenure.toFixed(1), func.headcount]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Recent Joiners List (ALL recent joiners, not limited)
  if (recentJoiners.length > 0) {
    sheet.addRow(['RECENT JOINERS (<1 YEAR)']);
    sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
    sheet.addRow([]);

    const recentHeaderRow = sheet.addRow(['Employee ID', 'Title', 'Function', 'Hire Date', 'FLRR']);
    styleHeader(recentHeaderRow);

    // Export ALL recent joiners, no limit
    recentJoiners.forEach(emp => {
      sheet.addRow([emp.employeeId, emp.title, emp.function, emp.hireDate, formatCurrency(emp.flrr)]);
    });
  }

  sheet.addRow([]);
  sheet.addRow([]);

  // Veterans List (ALL veterans)
  if (veterans.length > 0) {
    sheet.addRow(['VETERANS (5+ YEARS)']);
    sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
    sheet.addRow([]);

    const veteransHeaderRow = sheet.addRow(['Employee ID', 'Title', 'Function', 'Hire Date', 'Tenure (years)', 'FLRR']);
    styleHeader(veteransHeaderRow);

    veterans.forEach(emp => {
      sheet.addRow([emp.employeeId, emp.title, emp.function, emp.hireDate, emp.tenureYears.toFixed(1), formatCurrency(emp.flrr)]);
    });
  }

  sheet.columns = [
    { width: 25 },
    { width: 25 },
    { width: 20 },
    { width: 15 },
    { width: 15 },
    { width: 15 }
  ];
}

function addAutomationSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Automation');
  const { employees, totals } = data;

  // Automation patterns (simplified from component)
  const highRiskPatterns = [
    { pattern: /data\s*entry/i, score: 95 },
    { pattern: /clerk/i, score: 85 },
    { pattern: /bookkeep/i, score: 90 },
    { pattern: /payroll/i, score: 80 },
    { pattern: /accounts\s*(payable|receivable)/i, score: 85 },
    { pattern: /receptionist/i, score: 75 },
    { pattern: /processor/i, score: 70 },
  ];

  const mediumRiskPatterns = [
    { pattern: /accountant/i, score: 45 },
    { pattern: /coordinator/i, score: 40 },
    { pattern: /recruiter/i, score: 45 },
  ];

  const lowRiskPatterns = [
    { pattern: /director/i, score: 15 },
    { pattern: /manager/i, score: 20 },
    { pattern: /engineer/i, score: 25 },
  ];

  function getAutomationScore(title: string): number {
    const normalizedTitle = title.toLowerCase().trim();
    for (const { pattern, score } of highRiskPatterns) {
      if (pattern.test(normalizedTitle)) return score;
    }
    for (const { pattern, score } of mediumRiskPatterns) {
      if (pattern.test(normalizedTitle)) return score;
    }
    for (const { pattern, score } of lowRiskPatterns) {
      if (pattern.test(normalizedTitle)) return score;
    }
    return 35;
  }

  // Group by title
  const titleGroups = new Map<string, { headcount: number; totalFLRR: number; score: number }>();
  for (const emp of employees) {
    const title = emp.title || 'Unknown Title';
    const existing = titleGroups.get(title) || { headcount: 0, totalFLRR: 0, score: getAutomationScore(title) };
    titleGroups.set(title, {
      headcount: existing.headcount + 1,
      totalFLRR: existing.totalFLRR + emp.flrr,
      score: existing.score,
    });
  }

  const opportunities = Array.from(titleGroups.entries())
    .map(([title, stats]) => ({
      title,
      ...stats,
      level: stats.score >= 60 ? 'High' : stats.score >= 35 ? 'Medium' : 'Low',
    }))
    .sort((a, b) => b.score - a.score);

  const high = opportunities.filter(o => o.level === 'High');
  const medium = opportunities.filter(o => o.level === 'Medium');
  const low = opportunities.filter(o => o.level === 'Low');

  const highHeadcount = high.reduce((sum, o) => sum + o.headcount, 0);
  const highFLRR = high.reduce((sum, o) => sum + o.totalFLRR, 0);
  const mediumHeadcount = medium.reduce((sum, o) => sum + o.headcount, 0);
  const mediumFLRR = medium.reduce((sum, o) => sum + o.totalFLRR, 0);
  const lowHeadcount = low.reduce((sum, o) => sum + o.headcount, 0);
  const lowFLRR = low.reduce((sum, o) => sum + o.totalFLRR, 0);

  // Summary
  sheet.addRow(['AUTOMATION OPPORTUNITY SUMMARY']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  sheet.addRow(['Opportunity Level', 'Headcount', '% of Total', 'FLRR Exposure']);
  styleHeader(sheet.getRow(3));

  sheet.addRow(['High', highHeadcount, formatPercent((highHeadcount / totals.headcount) * 100), formatCurrency(highFLRR)]);
  sheet.addRow(['Medium', mediumHeadcount, formatPercent((mediumHeadcount / totals.headcount) * 100), formatCurrency(mediumFLRR)]);
  sheet.addRow(['Low', lowHeadcount, formatPercent((lowHeadcount / totals.headcount) * 100), formatCurrency(lowFLRR)]);

  sheet.addRow([]);
  sheet.addRow([]);

  // Detailed List
  sheet.addRow(['AUTOMATION OPPORTUNITY BY JOB TITLE']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const detailHeaderRow = sheet.addRow(['Job Title', 'Opportunity Level', 'Score', 'Headcount', 'Total FLRR']);
  styleHeader(detailHeaderRow);

  opportunities.forEach(opp => {
    sheet.addRow([opp.title, opp.level, `${opp.score}%`, opp.headcount, formatCurrency(opp.totalFLRR)]);
  });

  sheet.columns = [
    { width: 35 },
    { width: 18 },
    { width: 12 },
    { width: 12 },
    { width: 18 }
  ];
}

function addOffshoringSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Offshoring');
  const { countryStats, functionStats, employees } = data;

  // Country Summary
  sheet.addRow(['COUNTRY SUMMARY']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const countryHeaderRow = sheet.addRow(['Country', 'Headcount', 'Total FLRR', 'Avg FLRR']);
  styleHeader(countryHeaderRow);

  countryStats.forEach(cs => {
    sheet.addRow([cs.country, cs.headcount, formatCurrency(cs.totalFLRR), formatCurrency(cs.avgFLRR)]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Function Distribution by Country
  sheet.addRow(['FUNCTION BREAKDOWN']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const funcHeaderRow = sheet.addRow(['Function', 'Headcount', 'Total FLRR', 'Avg FLRR']);
  styleHeader(funcHeaderRow);

  functionStats.forEach(func => {
    sheet.addRow([func.function, func.headcount, formatCurrency(func.totalFLRR), formatCurrency(func.avgFLRR)]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Employee Distribution by Country (ALL employees, not limited)
  sheet.addRow(['EMPLOYEE DISTRIBUTION BY COUNTRY']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);
  sheet.addRow(['Note: Country cost tags (Best-cost vs High-cost) require user input in the dashboard.']);
  sheet.addRow([]);

  const empHeaderRow = sheet.addRow(['Employee ID', 'Title', 'Function', 'Country', 'Location', 'FLRR']);
  styleHeader(empHeaderRow);

  // Export ALL employees, no limit
  employees.forEach(emp => {
    sheet.addRow([emp.employeeId, emp.title, emp.function, emp.country, emp.location, formatCurrency(emp.flrr)]);
  });

  sheet.columns = [
    { width: 15 },
    { width: 25 },
    { width: 15 },
    { width: 15 },
    { width: 20 },
    { width: 15 }
  ];
}

function addCompensationSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('Compensation');
  const { employees, functionStats } = data;
  const benchmarks = defaultBenchmarks;

  // Calculate variable compensation by function
  const compData = functionStats.map(func => {
    const funcEmps = employees.filter(e => e.function === func.function);
    const totalBase = funcEmps.reduce((sum, e) => sum + e.baseSalary, 0);
    const totalBonus = funcEmps.reduce((sum, e) => sum + e.bonus, 0);
    const totalComp = totalBase + totalBonus;
    const avgVariablePercent = totalComp > 0 ? (totalBonus / totalComp) * 100 : 0;
    const target = benchmarks.targetVariableByRole[func.function] || benchmarks.targetVariableByRole['default'];

    return {
      function: func.function,
      headcount: func.headcount,
      totalBase,
      totalBonus,
      totalFLRR: func.totalFLRR,
      variablePercent: avgVariablePercent,
      targetPercent: target,
      gap: target - avgVariablePercent,
      status: avgVariablePercent < target * 0.7 ? 'Below Target' : 'On Target',
    };
  }).sort((a, b) => b.gap - a.gap);

  // Summary
  sheet.addRow(['COMPENSATION ANALYSIS']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const headerRow = sheet.addRow(['Function', 'Headcount', 'Total Base', 'Total Bonus', 'Total FLRR', 'Variable %', 'Target %', 'Status']);
  styleHeader(headerRow);

  compData.forEach(row => {
    sheet.addRow([
      row.function,
      row.headcount,
      formatCurrency(row.totalBase),
      formatCurrency(row.totalBonus),
      formatCurrency(row.totalFLRR),
      formatPercent(row.variablePercent),
      formatPercent(row.targetPercent),
      row.status
    ]);
  });

  sheet.addRow([]);
  sheet.addRow([]);

  // Benchmarks Reference
  sheet.addRow(['VARIABLE COMPENSATION BENCHMARKS']);
  sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
  sheet.addRow([]);

  const benchHeaderRow = sheet.addRow(['Role', 'Target Variable %']);
  styleHeader(benchHeaderRow);

  Object.entries(benchmarks.targetVariableByRole)
    .filter(([key]) => key !== 'default')
    .forEach(([role, target]) => {
      sheet.addRow([role, `${target}%`]);
    });

  sheet.addRow([]);
  sheet.addRow([]);

  // Compensation Differences
  const titleMap = new Map<string, typeof employees>();
  employees.forEach(emp => {
    const key = emp.title;
    const existing = titleMap.get(key) || [];
    existing.push(emp);
    titleMap.set(key, existing);
  });

  const differences: { title: string; func: string; count: number; min: number; max: number; avg: number; variance: number; variancePercent: number }[] = [];

  titleMap.forEach((emps, title) => {
    if (emps.length < 2) return;
    const comps = emps.map(e => e.baseSalary + e.bonus);
    const minComp = Math.min(...comps);
    const maxComp = Math.max(...comps);
    const avgComp = comps.reduce((a, b) => a + b, 0) / comps.length;
    const variance = maxComp - minComp;
    const variancePercent = avgComp > 0 ? (variance / avgComp) * 100 : 0;

    if (variancePercent > 30 && variance > 10000) {
      differences.push({
        title,
        func: emps[0].function,
        count: emps.length,
        min: minComp,
        max: maxComp,
        avg: avgComp,
        variance,
        variancePercent,
      });
    }
  });

  if (differences.length > 0) {
    sheet.addRow(['SIGNIFICANT COMPENSATION DIFFERENCES (>30% variance)']);
    sheet.getRow(sheet.rowCount).font = { bold: true, size: 14 };
    sheet.addRow([]);

    const diffHeaderRow = sheet.addRow(['Job Title', 'Function', '# Employees', 'Min Comp', 'Max Comp', 'Avg Comp', 'Variance', 'Variance %']);
    styleHeader(diffHeaderRow);

    // Export ALL compensation differences, no limit
    differences.sort((a, b) => b.variancePercent - a.variancePercent).forEach(diff => {
      sheet.addRow([
        diff.title,
        diff.func,
        diff.count,
        formatCurrency(diff.min),
        formatCurrency(diff.max),
        formatCurrency(diff.avg),
        formatCurrency(diff.variance),
        formatPercent(diff.variancePercent)
      ]);
    });
  }

  sheet.columns = [
    { width: 25 },
    { width: 12 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 12 },
    { width: 12 },
    { width: 15 }
  ];
}

// New sheet: Complete Employee Data (all employees, all fields)
function addCompleteEmployeeDataSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('All Employee Data');
  const { employees } = data;
  const now = new Date();

  sheet.addRow(['COMPLETE EMPLOYEE DATA']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([`Total: ${employees.length} employees`]);
  sheet.addRow([]);

  const headerRow = sheet.addRow([
    'Employee ID',
    'Manager ID',
    'Title',
    'Function',
    'Business Unit',
    'Location',
    'Country',
    'Hire Date',
    'Tenure (years)',
    'Base Salary',
    'Bonus',
    'Total Comp',
    'FLRR',
    'Variable %'
  ]);
  styleHeader(headerRow);

  // Export ALL employees with ALL fields
  employees.forEach(emp => {
    const hireDate = new Date(emp.hireDate);
    const tenureYears = (now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24 * 365);
    const totalComp = emp.baseSalary + emp.bonus;
    const variablePercent = totalComp > 0 ? (emp.bonus / totalComp) * 100 : 0;

    sheet.addRow([
      emp.employeeId,
      emp.managerId || 'None',
      emp.title,
      emp.function,
      emp.businessUnit || 'Unknown',
      emp.location,
      emp.country,
      emp.hireDate,
      tenureYears.toFixed(1),
      formatCurrency(emp.baseSalary),
      formatCurrency(emp.bonus),
      formatCurrency(totalComp),
      formatCurrency(emp.flrr),
      formatPercent(variablePercent)
    ]);
  });

  sheet.columns = [
    { width: 15 },
    { width: 15 },
    { width: 25 },
    { width: 15 },
    { width: 15 },
    { width: 20 },
    { width: 15 },
    { width: 12 },
    { width: 12 },
    { width: 15 },
    { width: 12 },
    { width: 15 },
    { width: 15 },
    { width: 12 }
  ];
}

// New sheet: All Managers with complete span data
function addAllManagersSheet(workbook: ExcelJS.Workbook, data: AnalysisData) {
  const sheet = workbook.addWorksheet('All Managers');
  const { spanStats, employees } = data;

  sheet.addRow(['ALL MANAGERS - SPAN OF CONTROL']);
  sheet.getRow(1).font = { bold: true, size: 14 };
  sheet.addRow([`Total: ${spanStats.length} managers`]);
  sheet.addRow([]);

  const headerRow = sheet.addRow([
    'Manager ID',
    'Manager Title',
    'Function',
    'Layer',
    'Direct Reports',
    'Manager FLRR'
  ]);
  styleHeader(headerRow);

  // Export ALL managers
  spanStats.forEach(span => {
    const manager = employees.find(e => e.employeeId === span.managerId);
    sheet.addRow([
      span.managerId,
      span.managerName,
      span.function,
      span.layer,
      span.directReports,
      manager ? formatCurrency(manager.flrr) : 'N/A'
    ]);
  });

  sheet.columns = [
    { width: 15 },
    { width: 30 },
    { width: 15 },
    { width: 10 },
    { width: 15 },
    { width: 15 }
  ];
}
