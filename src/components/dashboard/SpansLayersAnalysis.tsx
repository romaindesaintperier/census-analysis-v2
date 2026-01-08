import { useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { AnalysisData, Benchmarks } from '@/types/employee';
import { formatCurrency, formatPercent } from '@/lib/analysis';
import { defaultBenchmarks } from '@/lib/analysis';
import { AlertTriangle, CheckCircle, Users, Layers, User } from 'lucide-react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';

interface SpansLayersAnalysisProps {
  data: AnalysisData;
  benchmarks?: Benchmarks;
}

export function SpansLayersAnalysis({ data, benchmarks = defaultBenchmarks }: SpansLayersAnalysisProps) {
  const { employees, layerStats, spanStats, functionSpanStats, totals } = data;

  // Span distribution data
  const spanDistribution = [
    { range: '1', count: spanStats.filter(s => s.directReports === 1).length },
    { range: '2-4', count: spanStats.filter(s => s.directReports >= 2 && s.directReports <= 4).length },
    { range: '5-7', count: spanStats.filter(s => s.directReports >= 5 && s.directReports <= 7).length },
    { range: '8-10', count: spanStats.filter(s => s.directReports >= 8 && s.directReports <= 10).length },
    { range: '11+', count: spanStats.filter(s => s.directReports >= 11).length },
  ];

  const getSpanColor = (range: string) => {
    if (range === '1') return 'hsl(var(--destructive))';
    if (range === '2-4') return 'hsl(var(--warning))';
    if (range === '5-7' || range === '8-10') return 'hsl(var(--success))';
    return 'hsl(var(--chart-3))';
  };

  const outliers = {
    singleReport: spanStats.filter(s => s.directReports === 1),
    narrow: spanStats.filter(s => s.directReports > 1 && s.directReports < benchmarks.minSpan),
    wide: spanStats.filter(s => s.directReports > benchmarks.maxSpan),
  };

  // Get direct report info for single-report managers
  const singleReportManagersWithDetails = useMemo(() => {
    return outliers.singleReport.map(mgr => {
      // Find the direct report (employee whose managerId matches this manager)
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitles: directReports.map(dr => dr.title),
        directReportTitle: directReports[0]?.title || 'Unknown',
      };
    });
  }, [outliers.singleReport, employees]);

  // Get direct report info for managers below min span
  const belowMinSpanManagersWithDetails = useMemo(() => {
    return outliers.narrow.map(mgr => {
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitles: directReports.map(dr => dr.title),
      };
    });
  }, [outliers.narrow, employees]);

  // Get direct report info for managers above max span
  const aboveMaxSpanManagersWithDetails = useMemo(() => {
    return outliers.wide.map(mgr => {
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitles: directReports.map(dr => dr.title),
      };
    });
  }, [outliers.wide, employees]);

  // Org-wide average span
  const orgAvgSpan = totals.avgSpan;
  const orgAvgLayers = layerStats.length > 0 ? Math.max(...layerStats.map(l => l.layer)) + 1 : 0;

  return (
    <div className="space-y-6 animate-fade-in">
      {/* Key Metrics Row */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <Card>
          <CardContent className="p-4">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-primary/10">
                <Users className="w-5 h-5 text-primary" />
              </div>
              <div>
                <p className="text-sm text-muted-foreground">Total Managers</p>
                <p className="text-2xl font-semibold">{totals.totalManagers}</p>
                <p className="text-xs text-muted-foreground">{formatPercent(totals.managerPercent)} of total</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-4">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-accent/10">
                <User className="w-5 h-5 text-accent" />
              </div>
              <div>
                <p className="text-sm text-muted-foreground">CEO Direct Reports</p>
                <p className="text-2xl font-semibold">{totals.ceoDirectReports}</p>
                <p className="text-xs text-muted-foreground">Layer 1 headcount</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-4">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-secondary">
                <Layers className="w-5 h-5 text-foreground" />
              </div>
              <div>
                <p className="text-sm text-muted-foreground">Org Layers</p>
                <p className="text-2xl font-semibold">{orgAvgLayers}</p>
                <p className="text-xs text-muted-foreground">CEO = Layer 0</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-4">
            <div className="flex items-center gap-3">
              <div className="p-2 rounded-lg bg-success/10">
                <CheckCircle className="w-5 h-5 text-success" />
              </div>
              <div>
                <p className="text-sm text-muted-foreground">Avg Span of Control</p>
                <p className="text-2xl font-semibold">{orgAvgSpan.toFixed(1)}</p>
                <p className="text-xs text-muted-foreground">Target: {benchmarks.minSpan}-{benchmarks.maxSpan}</p>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Layer Analysis - 3 charts side by side */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* FLRR by Layer */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base">FLRR by Layer</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={layerStats} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis type="number" tickFormatter={(v) => formatCurrency(v)} fontSize={11} />
                  <YAxis type="category" dataKey="layer" tickFormatter={(v) => `L${v}`} fontSize={11} width={30} />
                  <Tooltip 
                    formatter={(value: number) => formatCurrency(value)}
                    labelFormatter={(label) => `Layer ${label}`}
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                  <Bar dataKey="totalFLRR" fill="hsl(var(--primary))" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>

        {/* Headcount by Layer */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base">Headcount by Layer</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={layerStats} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis type="number" fontSize={11} />
                  <YAxis type="category" dataKey="layer" tickFormatter={(v) => `L${v}`} fontSize={11} width={30} />
                  <Tooltip 
                    formatter={(value: number) => [value, 'Employees']}
                    labelFormatter={(label) => `Layer ${label}`}
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                  <Bar dataKey="headcount" fill="hsl(var(--chart-2))" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>

        {/* Avg Tenure by Layer */}
        <Card>
          <CardHeader className="pb-2">
            <CardTitle className="text-base">Avg Tenure by Layer</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={layerStats} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis type="number" unit=" yrs" fontSize={11} />
                  <YAxis type="category" dataKey="layer" tickFormatter={(v) => `L${v}`} fontSize={11} width={30} />
                  <Tooltip 
                    formatter={(value: number) => [`${value.toFixed(1)} years`, 'Avg Tenure']}
                    labelFormatter={(label) => `Layer ${label}`}
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                  <Bar dataKey="avgTenure" fill="hsl(var(--chart-3))" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Layer Stats Summary - Box Layout */}
      <Card>
        <CardHeader className="pb-2">
          <CardTitle className="text-base">Layer Breakdown</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="flex flex-wrap gap-3">
            {layerStats.map((layer, index) => (
              <div 
                key={layer.layer}
                className="flex-1 min-w-[120px] p-4 rounded-lg border border-border"
                style={{ backgroundColor: `hsl(var(--chart-${(index % 5) + 1}) / 0.1)` }}
              >
                <div className="text-lg font-bold text-foreground mb-2">L{layer.layer}</div>
                <div className="space-y-1 text-sm">
                  <div className="flex justify-between">
                    <span className="text-muted-foreground">Headcount:</span>
                    <span className="font-medium text-foreground">{layer.headcount}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted-foreground">Managers:</span>
                    <span className="font-medium" style={{ color: `hsl(var(--chart-${(index % 5) + 1}))` }}>{layer.managers}</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-muted-foreground">ICs:</span>
                    <span className="font-medium text-foreground">{layer.ics}</span>
                  </div>
                </div>
              </div>
            ))}
          </div>
          <div className="flex flex-wrap gap-4 mt-4 text-sm text-muted-foreground">
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 rounded-full bg-primary"></div>
              <span>Managers</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-3 h-3 rounded-full bg-muted-foreground"></div>
              <span>Individual Contributors</span>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Span of Control Distribution & Benchmark */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle>Span of Control Distribution</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={spanDistribution}>
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis dataKey="range" />
                  <YAxis />
                  <Tooltip 
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                  <Bar dataKey="count" radius={[4, 4, 0, 0]}>
                    {spanDistribution.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={getSpanColor(entry.range)} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
            <div className="mt-4 flex items-center gap-4 text-sm">
              <span className="flex items-center gap-1">
                <div className="w-3 h-3 rounded-full bg-success" />
                Optimal ({benchmarks.minSpan}-{benchmarks.maxSpan})
              </span>
              <span className="flex items-center gap-1">
                <div className="w-3 h-3 rounded-full bg-warning" />
                Narrow
              </span>
              <span className="flex items-center gap-1">
                <div className="w-3 h-3 rounded-full bg-destructive" />
                Single report
              </span>
            </div>
          </CardContent>
        </Card>

        {/* Benchmark Comparison */}
        <Card>
          <CardHeader>
            <CardTitle>Benchmark Comparison</CardTitle>
          </CardHeader>
          <CardContent className="space-y-4">
            <div className="p-4 rounded-lg border bg-card">
              <div className="flex items-center justify-between mb-2">
                <span className="font-medium">Max Layers</span>
                <Badge variant={layerStats.length <= benchmarks.maxLayers ? 'default' : 'destructive'}>
                  {layerStats.length <= benchmarks.maxLayers ? (
                    <CheckCircle className="w-3 h-3 mr-1" />
                  ) : (
                    <AlertTriangle className="w-3 h-3 mr-1" />
                  )}
                  {orgAvgLayers} / {benchmarks.maxLayers}
                </Badge>
              </div>
              <p className="text-sm text-muted-foreground">
                {layerStats.length <= benchmarks.maxLayers 
                  ? 'Within best practice range' 
                  : `${layerStats.length - benchmarks.maxLayers} excess layers`}
              </p>
            </div>

            <div className="p-4 rounded-lg border bg-card">
              <div className="flex items-center justify-between mb-2">
                <span className="font-medium">Optimal Span Range</span>
                <span className="text-muted-foreground">{benchmarks.minSpan}-{benchmarks.maxSpan}</span>
              </div>
              <p className="text-sm text-muted-foreground">
                Average span: {totals.avgSpan.toFixed(1)}
              </p>
            </div>

            <div className="space-y-2">
              <h4 className="font-medium">Issues Identified</h4>
              {outliers.singleReport.length > 0 && (
                <div className="flex items-center gap-2 text-sm text-destructive">
                  <AlertTriangle className="w-4 h-4" />
                  {outliers.singleReport.length} single-report managers
                </div>
              )}
              {outliers.narrow.length > 0 && (
                <div className="flex items-center gap-2 text-sm text-warning">
                  <AlertTriangle className="w-4 h-4" />
                  {outliers.narrow.length} managers below min span
                </div>
              )}
              {outliers.wide.length > 0 && (
                <div className="flex items-center gap-2 text-sm text-muted-foreground">
                  <AlertTriangle className="w-4 h-4" />
                  {outliers.wide.length} managers above max span
                </div>
              )}
              {outliers.singleReport.length === 0 && outliers.narrow.length === 0 && outliers.wide.length === 0 && (
                <div className="flex items-center gap-2 text-sm text-success">
                  <CheckCircle className="w-4 h-4" />
                  All spans within acceptable range
                </div>
              )}
            </div>
          </CardContent>
        </Card>
      </div>


      {/* Function Spans & Layers Table */}
      <Card>
        <CardHeader>
          <CardTitle>Span of Control by Function</CardTitle>
        </CardHeader>
        <CardContent>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Function</TableHead>
                <TableHead className="text-right">Employees</TableHead>
                <TableHead className="text-right">Managers</TableHead>
                <TableHead className="text-right">Manager %</TableHead>
                <TableHead className="text-right">Avg Span</TableHead>
                <TableHead className="text-right">Layers</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {functionSpanStats.map((func) => (
                <TableRow key={func.function}>
                  <TableCell className="font-medium">{func.function}</TableCell>
                  <TableCell className="text-right">{func.totalEmployees}</TableCell>
                  <TableCell className="text-right">{func.managerCount}</TableCell>
                  <TableCell className="text-right">{formatPercent(func.managerPercent)}</TableCell>
                  <TableCell className="text-right">
                    <span className={func.avgSpan < benchmarks.minSpan ? 'text-warning' : func.avgSpan > benchmarks.maxSpan ? 'text-destructive' : 'text-success'}>
                      {func.avgSpan.toFixed(1)}
                    </span>
                  </TableCell>
                  <TableCell className="text-right">{func.layers}</TableCell>
                </TableRow>
              ))}
              {/* Org Average Row */}
              <TableRow className="bg-secondary/30">
                <TableCell className="font-bold">Organization Average</TableCell>
                <TableCell className="text-right font-bold">{totals.headcount}</TableCell>
                <TableCell className="text-right font-bold">{totals.totalManagers}</TableCell>
                <TableCell className="text-right font-bold">{formatPercent(totals.managerPercent)}</TableCell>
                <TableCell className="text-right font-bold">{orgAvgSpan.toFixed(1)}</TableCell>
                <TableCell className="text-right font-bold">{orgAvgLayers}</TableCell>
              </TableRow>
            </TableBody>
          </Table>
        </CardContent>
      </Card>

      {/* Single Report Managers Table */}
      {singleReportManagersWithDetails.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-destructive" />
              Single-Report Managers ({singleReportManagersWithDetails.length})
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Function</TableHead>
                    <TableHead>Manager Title</TableHead>
                    <TableHead>Employee ID</TableHead>
                    <TableHead className="text-center">Layer</TableHead>
                    <TableHead>Direct Report Title</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {singleReportManagersWithDetails.map((mgr) => (
                    <TableRow key={mgr.managerId}>
                      <TableCell className="font-medium">{mgr.function}</TableCell>
                      <TableCell>{mgr.managerName}</TableCell>
                      <TableCell className="text-muted-foreground">{mgr.managerId}</TableCell>
                      <TableCell className="text-center">{mgr.layer}</TableCell>
                      <TableCell>{mgr.directReportTitle}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Managers Below Minimum Span Table */}
      {belowMinSpanManagersWithDetails.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-warning" />
              Managers Below Minimum Span ({belowMinSpanManagersWithDetails.length})
              <Badge variant="outline" className="ml-2 text-warning border-warning">
                &lt; {benchmarks.minSpan} reports
              </Badge>
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Function</TableHead>
                    <TableHead>Manager Title</TableHead>
                    <TableHead>Employee ID</TableHead>
                    <TableHead className="text-center">Layer</TableHead>
                    <TableHead className="text-center">Direct Reports</TableHead>
                    <TableHead>Direct Report Titles</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {belowMinSpanManagersWithDetails.map((mgr) => (
                    <TableRow key={mgr.managerId}>
                      <TableCell className="font-medium">{mgr.function}</TableCell>
                      <TableCell>{mgr.managerName}</TableCell>
                      <TableCell className="text-muted-foreground">{mgr.managerId}</TableCell>
                      <TableCell className="text-center">{mgr.layer}</TableCell>
                      <TableCell className="text-center">
                        <Badge variant="outline" className="text-warning border-warning">
                          {mgr.directReports}
                        </Badge>
                      </TableCell>
                      <TableCell className="max-w-[300px]">
                        <div className="text-sm">
                          {mgr.directReportTitles.join(', ')}
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      )}

      {/* Managers Above Maximum Span Table */}
      {aboveMaxSpanManagersWithDetails.length > 0 && (
        <Card>
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <AlertTriangle className="w-5 h-5 text-chart-3" />
              Managers Above Maximum Span ({aboveMaxSpanManagersWithDetails.length})
              <Badge variant="outline" className="ml-2">
                &gt; {benchmarks.maxSpan} reports
              </Badge>
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Function</TableHead>
                    <TableHead>Manager Title</TableHead>
                    <TableHead>Employee ID</TableHead>
                    <TableHead className="text-center">Layer</TableHead>
                    <TableHead className="text-center">Direct Reports</TableHead>
                    <TableHead>Direct Report Titles</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {aboveMaxSpanManagersWithDetails.map((mgr) => (
                    <TableRow key={mgr.managerId}>
                      <TableCell className="font-medium">{mgr.function}</TableCell>
                      <TableCell>{mgr.managerName}</TableCell>
                      <TableCell className="text-muted-foreground">{mgr.managerId}</TableCell>
                      <TableCell className="text-center">{mgr.layer}</TableCell>
                      <TableCell className="text-center">
                        <Badge variant="outline">
                          {mgr.directReports}
                        </Badge>
                      </TableCell>
                      <TableCell className="max-w-[300px]">
                        <div className="text-sm truncate" title={mgr.directReportTitles.join(', ')}>
                          {mgr.directReportTitles.slice(0, 5).join(', ')}
                          {mgr.directReportTitles.length > 5 && ` +${mgr.directReportTitles.length - 5} more`}
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );
}
