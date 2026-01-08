import jsPDF from 'jspdf';
import { AnalysisData } from '@/types/employee';
import { formatCurrency, formatPercent, defaultBenchmarks } from './analysis';

interface TableColumn {
  header: string;
  key: string;
  width: number;
  align?: 'left' | 'center' | 'right';
}

export async function exportToPDF(data: AnalysisData): Promise<void> {
  const pdf = new jsPDF({
    orientation: 'portrait',
    unit: 'mm',
    format: 'a4',
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 15;
  const contentWidth = pageWidth - 2 * margin;
  let yPos = margin;

  const colors = {
    primary: [59, 130, 246] as [number, number, number],
    header: [55, 65, 81] as [number, number, number],
    text: [31, 41, 55] as [number, number, number],
    muted: [107, 114, 128] as [number, number, number],
    border: [229, 231, 235] as [number, number, number],
    success: [34, 197, 94] as [number, number, number],
    warning: [234, 179, 8] as [number, number, number],
  };

  function checkPageBreak(requiredSpace: number): void {
    if (yPos + requiredSpace > pageHeight - margin) {
      pdf.addPage();
      yPos = margin;
    }
  }

  function addTabHeader(title: string): void {
    pdf.setFontSize(18);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(...colors.primary);
    pdf.text(title, margin, yPos);
    yPos += 10;
    
    // Underline
    pdf.setDrawColor(...colors.primary);
    pdf.setLineWidth(0.5);
    pdf.line(margin, yPos - 3, pageWidth - margin, yPos - 3);
    yPos += 5;
  }

  function addSectionHeader(title: string): void {
    checkPageBreak(15);
    pdf.setFontSize(12);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(...colors.header);
    pdf.text(title, margin, yPos);
    yPos += 7;
  }

  function addText(text: string, fontSize: number = 10, isBold: boolean = false): void {
    checkPageBreak(8);
    pdf.setFontSize(fontSize);
    pdf.setFont('helvetica', isBold ? 'bold' : 'normal');
    pdf.setTextColor(...colors.text);
    
    const lines = pdf.splitTextToSize(text, contentWidth);
    lines.forEach((line: string) => {
      checkPageBreak(6);
      pdf.text(line, margin, yPos);
      yPos += 5;
    });
  }

  function addKeyValue(key: string, value: string): void {
    checkPageBreak(8);
    pdf.setFontSize(10);
    pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(...colors.muted);
    pdf.text(key + ':', margin, yPos);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(...colors.text);
    pdf.text(value, margin + 60, yPos);
    yPos += 6;
  }

  function addTable(columns: TableColumn[], rows: Record<string, string | number>[], maxRows: number = 50): void {
    const rowHeight = 7;
    const headerHeight = 8;
    
    // Calculate if table fits, otherwise it will span pages
    checkPageBreak(headerHeight + rowHeight * Math.min(rows.length, 3));

    // Header
    pdf.setFillColor(...colors.header);
    pdf.rect(margin, yPos - 5, contentWidth, headerHeight, 'F');
    pdf.setFontSize(9);
    pdf.setFont('helvetica', 'bold');
    pdf.setTextColor(255, 255, 255);
    
    let xPos = margin + 2;
    columns.forEach(col => {
      const align = col.align || 'left';
      let textX = xPos;
      if (align === 'center') textX = xPos + col.width / 2;
      if (align === 'right') textX = xPos + col.width - 2;
      pdf.text(col.header, textX, yPos, { align });
      xPos += col.width;
    });
    yPos += headerHeight;

    // Rows
    pdf.setFont('helvetica', 'normal');
    pdf.setTextColor(...colors.text);
    
    const displayRows = rows.slice(0, maxRows);
    displayRows.forEach((row, index) => {
      checkPageBreak(rowHeight);
      
      // Alternate row background
      if (index % 2 === 0) {
        pdf.setFillColor(249, 250, 251);
        pdf.rect(margin, yPos - 5, contentWidth, rowHeight, 'F');
      }
      
      pdf.setFontSize(8);
      xPos = margin + 2;
      columns.forEach(col => {
        const value = String(row[col.key] ?? '');
        const align = col.align || 'left';
        let textX = xPos;
        if (align === 'center') textX = xPos + col.width / 2;
        if (align === 'right') textX = xPos + col.width - 2;
        
        // Truncate if too long
        const maxChars = Math.floor(col.width / 2);
        const displayValue = value.length > maxChars ? value.substring(0, maxChars - 2) + '..' : value;
        pdf.text(displayValue, textX, yPos, { align });
        xPos += col.width;
      });
      yPos += rowHeight;
    });

    if (rows.length > maxRows) {
      pdf.setFontSize(8);
      pdf.setTextColor(...colors.muted);
      pdf.text(`... and ${rows.length - maxRows} more rows`, margin, yPos);
      yPos += 6;
    }

    yPos += 5;
  }

  function addNewPage(tabName: string): void {
    pdf.addPage();
    yPos = margin;
    addTabHeader(tabName);
  }

  // ============ 1. SUMMARY TAB ============
  addTabHeader('Executive Summary');
  
  const { totals, quickWins, functionStats, employees } = data;
  const now = new Date();
  const avgTenure = employees.reduce((sum, emp) => {
    const hireDate = new Date(emp.hireDate);
    return sum + (now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24 * 365);
  }, 0) / employees.length;

  addSectionHeader('Key Metrics');
  addKeyValue('Total Headcount', String(totals.headcount));
  addKeyValue('Total FLRR', formatCurrency(totals.totalFLRR));
  addKeyValue('Average FLRR', formatCurrency(totals.avgFLRR));
  addKeyValue('Org Layers', String(totals.layers));
  addKeyValue('Average Span', totals.avgSpan.toFixed(1));
  addKeyValue('Manager:IC Ratio', `1:${totals.managerToICRatio > 0 ? (1 / totals.managerToICRatio).toFixed(1) : '0'}`);
  addKeyValue('Average Tenure', `${avgTenure.toFixed(1)} years`);
  addKeyValue('Avg Variable Comp', formatPercent(totals.avgVariablePercent));
  yPos += 5;

  if (quickWins.length > 0) {
    addSectionHeader('Value Creation Opportunities');
    addTable(
      [
        { header: 'Title', key: 'title', width: 50 },
        { header: 'Category', key: 'category', width: 25, align: 'center' },
        { header: 'Impact', key: 'impact', width: 20, align: 'center' },
        { header: 'Metric', key: 'metric', width: 35, align: 'right' },
      ],
      quickWins.map(qw => ({ title: qw.title, category: qw.category, impact: qw.impact, metric: qw.metric || '' })),
      10
    );
  }

  addSectionHeader('Function Summary');
  addTable(
    [
      { header: 'Function', key: 'function', width: 35 },
      { header: 'Headcount', key: 'headcount', width: 25, align: 'right' },
      { header: 'Total FLRR', key: 'totalFLRR', width: 35, align: 'right' },
      { header: 'Avg FLRR', key: 'avgFLRR', width: 35, align: 'right' },
    ],
    functionStats.map(f => ({
      function: f.function,
      headcount: f.headcount,
      totalFLRR: formatCurrency(f.totalFLRR),
      avgFLRR: formatCurrency(f.avgFLRR),
    }))
  );

  // ============ 2. HEADCOUNT TAB ============
  addNewPage('Headcount Breakdown');

  // Location stats
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

  addSectionHeader('Function Breakdown');
  addTable(
    [
      { header: 'Function', key: 'function', width: 40 },
      { header: 'Headcount', key: 'headcount', width: 30, align: 'right' },
      { header: 'Total FLRR', key: 'flrr', width: 40, align: 'right' },
    ],
    functionStats.map(f => ({ function: f.function, headcount: f.headcount, flrr: formatCurrency(f.totalFLRR) }))
  );

  addSectionHeader('Location Breakdown');
  addTable(
    [
      { header: 'Location', key: 'location', width: 50 },
      { header: 'Headcount', key: 'headcount', width: 30, align: 'right' },
      { header: 'Total FLRR', key: 'flrr', width: 40, align: 'right' },
    ],
    locationStats.map(l => ({ location: l.location, headcount: l.headcount, flrr: formatCurrency(l.flrr) })),
    15
  );

  addSectionHeader('Business Unit Breakdown');
  addTable(
    [
      { header: 'Business Unit', key: 'bu', width: 50 },
      { header: 'Headcount', key: 'headcount', width: 30, align: 'right' },
      { header: 'Total FLRR', key: 'flrr', width: 40, align: 'right' },
    ],
    businessUnitStats.map(b => ({ bu: b.businessUnit, headcount: b.headcount, flrr: formatCurrency(b.flrr) })),
    15
  );

  // ============ 3. ORG CHART TAB ============
  addNewPage('Org Chart');

  const { layerStats } = data;

  addSectionHeader('Org Structure Summary');
  addKeyValue('Total Layers', String(totals.layers));
  addKeyValue('Total Employees', String(totals.headcount));
  addKeyValue('Total Managers', String(totals.totalManagers));
  addKeyValue('CEO Direct Reports', String(totals.ceoDirectReports));
  addKeyValue('Average Span', totals.avgSpan.toFixed(1));
  yPos += 5;

  addSectionHeader('Layer Breakdown');
  addTable(
    [
      { header: 'Layer', key: 'layer', width: 20, align: 'center' },
      { header: 'Headcount', key: 'headcount', width: 25, align: 'right' },
      { header: 'Managers', key: 'managers', width: 25, align: 'right' },
      { header: 'ICs', key: 'ics', width: 20, align: 'right' },
      { header: 'Total FLRR', key: 'flrr', width: 35, align: 'right' },
    ],
    layerStats.map(l => ({
      layer: `L${l.layer}`,
      headcount: l.headcount,
      managers: l.managers,
      ics: l.ics,
      flrr: formatCurrency(l.totalFLRR),
    }))
  );

  // ============ 4. SPANS & LAYERS TAB ============
  addNewPage('Spans & Layers');

  const { spanStats, functionSpanStats } = data;
  const benchmarks = defaultBenchmarks;

  addSectionHeader('Key Metrics');
  addKeyValue('Total Managers', String(totals.totalManagers));
  addKeyValue('Manager %', formatPercent(totals.managerPercent));
  addKeyValue('Org Layers', `${totals.layers} (benchmark: ≤${benchmarks.maxLayers})`);
  addKeyValue('Avg Span', `${totals.avgSpan.toFixed(1)} (benchmark: ${benchmarks.minSpan}-${benchmarks.maxSpan})`);
  yPos += 5;

  const spanDistribution = [
    { range: '1', count: spanStats.filter(s => s.directReports === 1).length },
    { range: '2-4', count: spanStats.filter(s => s.directReports >= 2 && s.directReports <= 4).length },
    { range: '5-7', count: spanStats.filter(s => s.directReports >= 5 && s.directReports <= 7).length },
    { range: '8-10', count: spanStats.filter(s => s.directReports >= 8 && s.directReports <= 10).length },
    { range: '11+', count: spanStats.filter(s => s.directReports >= 11).length },
  ];

  addSectionHeader('Span Distribution');
  addTable(
    [
      { header: 'Span Range', key: 'range', width: 40, align: 'center' },
      { header: 'Manager Count', key: 'count', width: 40, align: 'right' },
    ],
    spanDistribution
  );

  addSectionHeader('Span of Control by Function');
  addTable(
    [
      { header: 'Function', key: 'function', width: 35 },
      { header: 'Employees', key: 'employees', width: 25, align: 'right' },
      { header: 'Managers', key: 'managers', width: 25, align: 'right' },
      { header: 'Avg Span', key: 'avgSpan', width: 25, align: 'right' },
      { header: 'Layers', key: 'layers', width: 20, align: 'right' },
    ],
    functionSpanStats.map(f => ({
      function: f.function,
      employees: f.totalEmployees,
      managers: f.managerCount,
      avgSpan: f.avgSpan.toFixed(1),
      layers: f.layers,
    }))
  );

  const singleReportManagers = spanStats.filter(s => s.directReports === 1);
  if (singleReportManagers.length > 0) {
    addSectionHeader('Single-Report Managers');
    addTable(
      [
        { header: 'Manager ID', key: 'id', width: 40 },
        { header: 'Function', key: 'function', width: 40 },
        { header: 'Layer', key: 'layer', width: 25, align: 'center' },
      ],
      singleReportManagers.map(m => ({ id: m.managerId, function: m.function, layer: m.layer })),
      20
    );
  }

  // ============ 5. TENURE TAB ============
  addNewPage('Tenure Analysis');

  const { tenureBands } = data;
  const employeesWithTenure = employees.map(emp => {
    const hireDate = new Date(emp.hireDate);
    const years = (now.getTime() - hireDate.getTime()) / (1000 * 60 * 60 * 24 * 365);
    return { ...emp, tenureYears: years };
  });

  const recentJoiners = employeesWithTenure.filter(e => e.tenureYears < 1);
  const veterans = employeesWithTenure.filter(e => e.tenureYears >= 5);

  addSectionHeader('Tenure Summary');
  addKeyValue('Average Tenure', `${avgTenure.toFixed(1)} years`);
  addKeyValue('Recent Joiners (<1 year)', String(recentJoiners.length));
  addKeyValue('Veterans (5+ years)', String(veterans.length));
  yPos += 5;

  addSectionHeader('Tenure Distribution');
  addTable(
    [
      { header: 'Tenure Band', key: 'band', width: 40 },
      { header: 'Headcount', key: 'headcount', width: 30, align: 'right' },
      { header: 'Total FLRR', key: 'flrr', width: 40, align: 'right' },
    ],
    tenureBands.map(b => ({ band: b.band, headcount: b.headcount, flrr: formatCurrency(b.totalFLRR) }))
  );

  const tenureByFunction = functionStats.map(func => {
    const funcEmps = employeesWithTenure.filter(e => e.function === func.function);
    const avgFuncTenure = funcEmps.reduce((sum, e) => sum + e.tenureYears, 0) / funcEmps.length;
    return { function: func.function, avgTenure: avgFuncTenure.toFixed(1), headcount: func.headcount };
  }).sort((a, b) => parseFloat(b.avgTenure) - parseFloat(a.avgTenure));

  addSectionHeader('Average Tenure by Function');
  addTable(
    [
      { header: 'Function', key: 'function', width: 40 },
      { header: 'Avg Tenure (years)', key: 'avgTenure', width: 35, align: 'right' },
      { header: 'Headcount', key: 'headcount', width: 30, align: 'right' },
    ],
    tenureByFunction
  );

  // ============ 6. AUTOMATION TAB ============
  addNewPage('Automation Analysis');

  // Simplified automation scoring
  const highRiskPatterns = [
    { pattern: /data\s*entry/i, score: 95 },
    { pattern: /clerk/i, score: 85 },
    { pattern: /bookkeep/i, score: 90 },
    { pattern: /payroll/i, score: 80 },
    { pattern: /processor/i, score: 70 },
  ];

  function getAutomationScore(title: string): number {
    const t = title.toLowerCase();
    for (const { pattern, score } of highRiskPatterns) {
      if (pattern.test(t)) return score;
    }
    if (/director|manager|engineer/i.test(t)) return 20;
    return 35;
  }

  const titleGroups = new Map<string, { headcount: number; totalFLRR: number; score: number }>();
  for (const emp of employees) {
    const title = emp.title || 'Unknown';
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

  addSectionHeader('Automation Opportunity Summary');
  addKeyValue('High Opportunity', `${high.reduce((s, o) => s + o.headcount, 0)} employees (${formatCurrency(high.reduce((s, o) => s + o.totalFLRR, 0))} FLRR)`);
  addKeyValue('Medium Opportunity', `${medium.reduce((s, o) => s + o.headcount, 0)} employees (${formatCurrency(medium.reduce((s, o) => s + o.totalFLRR, 0))} FLRR)`);
  addKeyValue('Low Opportunity', `${low.reduce((s, o) => s + o.headcount, 0)} employees (${formatCurrency(low.reduce((s, o) => s + o.totalFLRR, 0))} FLRR)`);
  yPos += 5;

  addSectionHeader('High Automation Opportunity Roles');
  addTable(
    [
      { header: 'Job Title', key: 'title', width: 50 },
      { header: 'Score', key: 'score', width: 20, align: 'center' },
      { header: 'Headcount', key: 'headcount', width: 25, align: 'right' },
      { header: 'FLRR', key: 'flrr', width: 35, align: 'right' },
    ],
    high.map(o => ({ title: o.title, score: `${o.score}%`, headcount: o.headcount, flrr: formatCurrency(o.totalFLRR) })),
    15
  );

  // ============ 7. OFFSHORING TAB ============
  addNewPage('Offshoring Analysis');

  const { countryStats } = data;

  addSectionHeader('Country Summary');
  addText('Note: Country cost tags (Best-cost vs High-cost) require user input in the dashboard.', 9);
  yPos += 3;

  addTable(
    [
      { header: 'Country', key: 'country', width: 40 },
      { header: 'Headcount', key: 'headcount', width: 25, align: 'right' },
      { header: 'Total FLRR', key: 'totalFLRR', width: 35, align: 'right' },
      { header: 'Avg FLRR', key: 'avgFLRR', width: 30, align: 'right' },
    ],
    countryStats.map(c => ({
      country: c.country,
      headcount: c.headcount,
      totalFLRR: formatCurrency(c.totalFLRR),
      avgFLRR: formatCurrency(c.avgFLRR),
    })),
    20
  );

  addSectionHeader('Function Distribution');
  addTable(
    [
      { header: 'Function', key: 'function', width: 40 },
      { header: 'Headcount', key: 'headcount', width: 25, align: 'right' },
      { header: 'Total FLRR', key: 'totalFLRR', width: 35, align: 'right' },
      { header: 'Avg FLRR', key: 'avgFLRR', width: 30, align: 'right' },
    ],
    functionStats.map(f => ({
      function: f.function,
      headcount: f.headcount,
      totalFLRR: formatCurrency(f.totalFLRR),
      avgFLRR: formatCurrency(f.avgFLRR),
    }))
  );

  // ============ 8. COMPENSATION TAB ============
  addNewPage('Compensation Analysis');

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
      totalBase: formatCurrency(totalBase),
      totalBonus: formatCurrency(totalBonus),
      variablePercent: formatPercent(avgVariablePercent),
      targetPercent: formatPercent(target),
      status: avgVariablePercent < target * 0.7 ? 'Below Target' : 'On Target',
    };
  });

  addSectionHeader('Compensation Breakdown by Function');
  addTable(
    [
      { header: 'Function', key: 'function', width: 30 },
      { header: 'HC', key: 'headcount', width: 15, align: 'right' },
      { header: 'Total Base', key: 'totalBase', width: 30, align: 'right' },
      { header: 'Total Bonus', key: 'totalBonus', width: 30, align: 'right' },
      { header: 'Var %', key: 'variablePercent', width: 18, align: 'right' },
      { header: 'Status', key: 'status', width: 25, align: 'center' },
    ],
    compData
  );

  addSectionHeader('Variable Compensation Benchmarks');
  const benchmarkRows = Object.entries(benchmarks.targetVariableByRole)
    .filter(([key]) => key !== 'default')
    .map(([role, target]) => ({ role, target: `${target}%` }));

  addTable(
    [
      { header: 'Role', key: 'role', width: 50 },
      { header: 'Target Variable %', key: 'target', width: 40, align: 'right' },
    ],
    benchmarkRows
  );

  // Footer on last page
  pdf.setFontSize(8);
  pdf.setTextColor(...colors.muted);
  pdf.text(`Generated on ${new Date().toLocaleDateString()}`, margin, pageHeight - 10);

  // Save
  pdf.save(`org-analysis-${new Date().toISOString().split('T')[0]}.pdf`);
}
