import { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Badge } from '@/components/ui/badge';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip, BarChart, Bar, XAxis, YAxis, CartesianGrid } from 'recharts';
import { AnalysisData, CountryStats } from '@/types/employee';
import { formatCurrency, formatPercent } from '@/lib/analysis';
import { Globe, DollarSign, AlertTriangle, Calculator, Scissors } from 'lucide-react';

interface OffshoringAnalysisProps {
  data: AnalysisData;
}

const COLORS = {
  bestCost: 'hsl(var(--success))',
  highCost: 'hsl(var(--chart-1))',
  untagged: 'hsl(var(--muted-foreground))',
};

// Top 20 roles typically suitable for offshoring based on industry research
const OFFSHORING_ROLES = [
  { title: 'Software Developer', reason: 'Highly standardized work, strong global talent pool' },
  { title: 'Customer Service Representative', reason: 'Process-driven, multilingual talent available' },
  { title: 'Data Entry Specialist', reason: 'Routine, transaction-based tasks' },
  { title: 'Accountant', reason: 'Standardized processes, time-zone advantages' },
  { title: 'IT Support Specialist', reason: 'Remote support capabilities, 24/7 coverage' },
  { title: 'Quality Assurance Analyst', reason: 'Testing can be done remotely' },
  { title: 'Financial Analyst', reason: 'Analytical work with clear deliverables' },
  { title: 'HR Administrator', reason: 'Transactional HR processes' },
  { title: 'Graphic Designer', reason: 'Creative work with clear briefs' },
  { title: 'Technical Writer', reason: 'Documentation can be standardized' },
  { title: 'Marketing Coordinator', reason: 'Digital marketing tasks' },
  { title: 'Payroll Specialist', reason: 'Routine processing tasks' },
  { title: 'Business Analyst', reason: 'Requirements gathering and documentation' },
  { title: 'Database Administrator', reason: 'Remote management capabilities' },
  { title: 'Network Engineer', reason: 'Remote monitoring and maintenance' },
  { title: 'Content Writer', reason: 'Clear deliverables, global talent' },
  { title: 'Procurement Specialist', reason: 'Transaction-based work' },
  { title: 'Claims Processor', reason: 'Standardized processes' },
  { title: 'Research Analyst', reason: 'Desk research and analysis' },
  { title: 'Compliance Analyst', reason: 'Documentation and monitoring' },
];

export function OffshoringAnalysis({ data }: OffshoringAnalysisProps) {
  const { employees, countryStats } = data;

  // BCC FLRR assumption (default $40k)
  const [bccFLRRAssumption, setBccFLRRAssumption] = useState(40000);

  // Country tags state
  const [countryTags, setCountryTags] = useState<Map<string, 'Best-cost' | 'High-cost'>>(
    new Map()
  );

  // Function offshoring potentials (user inputs)
  const [functionOffshoringPotentials, setFunctionOffshoringPotentials] = useState<Record<string, number>>({});

  // Calculate updated country stats with tags
  const taggedCountryStats = useMemo(() => {
    return countryStats.map(cs => ({
      ...cs,
      tag: countryTags.get(cs.country) || 'Untagged' as const,
    }));
  }, [countryStats, countryTags]);

  // Calculate function-level offshoring stats
  const functionOffshoringStats = useMemo(() => {
    const funcMap = new Map<string, {
      employees: typeof employees;
      bccHeadcount: number;
      hccHeadcount: number;
      bccFLRR: number;
      hccFLRR: number;
    }>();

    employees.forEach(emp => {
      const existing = funcMap.get(emp.function) || {
        employees: [],
        bccHeadcount: 0,
        hccHeadcount: 0,
        bccFLRR: 0,
        hccFLRR: 0,
      };
      existing.employees.push(emp);
      
      const tag = countryTags.get(emp.country);
      if (tag === 'Best-cost') {
        existing.bccHeadcount++;
        existing.bccFLRR += emp.flrr;
      } else if (tag === 'High-cost') {
        existing.hccHeadcount++;
        existing.hccFLRR += emp.flrr;
      }
      
      funcMap.set(emp.function, existing);
    });

    return Array.from(funcMap.entries()).map(([func, stats]) => {
      const totalHeadcount = stats.employees.length;
      const avgFLRR = stats.employees.reduce((sum, e) => sum + e.flrr, 0) / totalHeadcount;
      const offshoringPotential = functionOffshoringPotentials[func] || 0;
      
      // Calculate potential savings
      // offshoring potential % * HCC headcount * (avg HCC FLRR - BCC assumption)
      const avgHccFLRR = stats.hccHeadcount > 0 ? stats.hccFLRR / stats.hccHeadcount : 0;
      const headcountToOffshore = Math.round(stats.hccHeadcount * (offshoringPotential / 100));
      const potentialSavings = headcountToOffshore * (avgHccFLRR - bccFLRRAssumption);

      return {
        function: func,
        totalHeadcount,
        bccHeadcount: stats.bccHeadcount,
        hccHeadcount: stats.hccHeadcount,
        bccPercent: totalHeadcount > 0 ? (stats.bccHeadcount / totalHeadcount) * 100 : 0,
        hccPercent: totalHeadcount > 0 ? (stats.hccHeadcount / totalHeadcount) * 100 : 0,
        avgFLRR,
        bccFLRR: stats.bccFLRR,
        hccFLRR: stats.hccFLRR,
        offshoringPotential,
        potentialSavings: potentialSavings > 0 ? potentialSavings : 0,
        headcountToOffshore,
      };
    }).sort((a, b) => b.hccHeadcount - a.hccHeadcount);
  }, [employees, countryTags, functionOffshoringPotentials, bccFLRRAssumption]);

  // Calculate totals
  const totals = useMemo(() => {
    const bccCount = employees.filter(e => countryTags.get(e.country) === 'Best-cost').length;
    const hccCount = employees.filter(e => countryTags.get(e.country) === 'High-cost').length;
    const bccFLRR = employees.filter(e => countryTags.get(e.country) === 'Best-cost').reduce((sum, e) => sum + e.flrr, 0);
    const hccFLRR = employees.filter(e => countryTags.get(e.country) === 'High-cost').reduce((sum, e) => sum + e.flrr, 0);
    const totalPotentialSavings = functionOffshoringStats.reduce((sum, f) => sum + f.potentialSavings, 0);
    const totalHeadcountToOffshore = functionOffshoringStats.reduce((sum, f) => sum + f.headcountToOffshore, 0);
    
    // Severance = 25% of FLRR for HCC headcount to be offshored
    const avgHccFLRR = hccCount > 0 ? hccFLRR / hccCount : 0;
    const severanceCost = totalHeadcountToOffshore * avgHccFLRR * 0.25;

    return { bccCount, hccCount, bccFLRR, hccFLRR, totalPotentialSavings, severanceCost, totalHeadcountToOffshore };
  }, [employees, countryTags, functionOffshoringStats]);

  const handleCountryTagChange = (country: string, tag: 'Best-cost' | 'High-cost' | 'Untagged') => {
    setCountryTags(prev => {
      const next = new Map(prev);
      if (tag === 'Untagged') {
        next.delete(country);
      } else {
        next.set(country, tag);
      }
      return next;
    });
  };

  const handleOffshoringPotentialChange = (func: string, value: string) => {
    const numValue = Math.min(100, Math.max(0, parseInt(value) || 0));
    setFunctionOffshoringPotentials(prev => ({
      ...prev,
      [func]: numValue,
    }));
  };

  const pieData = [
    { name: 'Best-cost', value: totals.bccCount, flrr: totals.bccFLRR },
    { name: 'High-cost', value: totals.hccCount, flrr: totals.hccFLRR },
    { name: 'Untagged', value: employees.length - totals.bccCount - totals.hccCount, flrr: 0 },
  ].filter(d => d.value > 0);

  return (
    <div className="space-y-6 animate-fade-in">
      {/* User Input Required Message */}
      <div className="p-4 rounded-lg bg-warning/10 border border-warning/30 text-warning-foreground">
        <p className="text-sm font-medium text-warning">
          User input required to classify countries into high-cost vs. best-cost, and to estimate offshoring potential.
        </p>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card className="bg-success/10 border-success/30">
          <CardContent className="p-6">
            <div className="flex items-center gap-3">
              <Globe className="w-8 h-8 text-success" />
              <div>
                <p className="text-sm text-muted-foreground">Best-Cost Countries</p>
                <p className="text-2xl font-bold">{totals.bccCount}</p>
                <p className="text-sm text-muted-foreground">{formatCurrency(totals.bccFLRR)} FLRR</p>
                <p className="text-sm text-muted-foreground">
                  Avg FLRR per employee: {totals.bccCount > 0 ? formatCurrency(totals.bccFLRR / totals.bccCount) : 'N/A'}
                </p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center gap-3">
              <Globe className="w-8 h-8 text-primary" />
              <div>
                <p className="text-sm text-muted-foreground">High-Cost Countries</p>
                <p className="text-2xl font-bold">{totals.hccCount}</p>
                <p className="text-sm text-muted-foreground">{formatCurrency(totals.hccFLRR)} FLRR</p>
                <p className="text-sm text-muted-foreground">
                  Avg FLRR per employee: {totals.hccCount > 0 ? formatCurrency(totals.hccFLRR / totals.hccCount) : 'N/A'}
                </p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card className="gradient-aubergine text-primary-foreground">
          <CardContent className="p-6">
            <div className="flex items-center gap-3">
              <Calculator className="w-8 h-8" />
              <div>
                <p className="text-sm opacity-80">Potential Run-Rate Savings</p>
                <p className="text-2xl font-bold">{formatCurrency(totals.totalPotentialSavings)}</p>
                <p className="text-sm opacity-80">{totals.totalHeadcountToOffshore} roles</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card className="border-warning/30 bg-warning/5">
          <CardContent className="p-6">
            <div className="flex items-center gap-3">
              <Scissors className="w-8 h-8 text-warning" />
              <div>
                <p className="text-sm text-muted-foreground">Est. Severance Cost</p>
                <p className="text-2xl font-bold">{formatCurrency(totals.severanceCost)}</p>
                <p className="text-sm text-muted-foreground">25% of offshored FLRR</p>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* BCC FLRR Assumption Input */}
      <Card>
        <CardHeader className="pb-3">
          <CardTitle className="flex items-center gap-2 text-base">
            <DollarSign className="w-5 h-5" />
            Best-Cost Country FLRR Assumption
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-2">
              <span className="text-sm text-muted-foreground">FLRR per headcount in BCC:</span>
              <div className="relative w-32">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-muted-foreground">$</span>
                <Input
                  type="number"
                  value={bccFLRRAssumption}
                  onChange={(e) => setBccFLRRAssumption(parseInt(e.target.value) || 0)}
                  className="pl-7"
                />
              </div>
            </div>
            <p className="text-sm text-muted-foreground">
              This assumption is used to calculate potential savings when shifting headcount from HCC to BCC.
            </p>
          </div>
        </CardContent>
      </Card>

      {/* Country Tagging Table */}
      <Card>
        <CardHeader>
          <CardTitle>Country Summary & Tagging</CardTitle>
        </CardHeader>
        <CardContent>
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead>Country</TableHead>
                <TableHead className="text-right"># Headcount</TableHead>
                <TableHead className="text-right">Avg FLRR</TableHead>
                <TableHead className="text-center">Cost Tag</TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {taggedCountryStats.map((cs) => (
                <TableRow key={cs.country}>
                  <TableCell className="font-medium">{cs.country}</TableCell>
                  <TableCell className="text-right">{cs.headcount}</TableCell>
                  <TableCell className="text-right">{formatCurrency(cs.avgFLRR)}</TableCell>
                  <TableCell>
                    <Select
                      value={countryTags.get(cs.country) || 'Untagged'}
                      onValueChange={(value) => handleCountryTagChange(cs.country, value as any)}
                    >
                      <SelectTrigger className="w-[140px] mx-auto">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="Untagged">Untagged</SelectItem>
                        <SelectItem value="Best-cost">Best-Cost</SelectItem>
                        <SelectItem value="High-cost">High-Cost</SelectItem>
                      </SelectContent>
                    </Select>
                  </TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </CardContent>
      </Card>

      {/* Distribution Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle>Headcount by Cost Type</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie
                    data={pieData}
                    dataKey="value"
                    nameKey="name"
                    cx="50%"
                    cy="50%"
                    outerRadius={80}
                    innerRadius={50}
                    label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
                  >
                    {pieData.map((entry, index) => (
                      <Cell 
                        key={`cell-${index}`} 
                        fill={entry.name === 'Best-cost' ? COLORS.bestCost : entry.name === 'High-cost' ? COLORS.highCost : COLORS.untagged} 
                      />
                    ))}
                  </Pie>
                  <Tooltip 
                    formatter={(value: number, name: string, props: any) => [
                      `${value} employees (${formatCurrency(props.payload.flrr)})`,
                      name
                    ]}
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="flex justify-center gap-6 mt-4">
              <span className="flex items-center gap-2 text-sm">
                <div className="w-3 h-3 rounded-full" style={{ backgroundColor: COLORS.bestCost }} />
                Best-cost
              </span>
              <span className="flex items-center gap-2 text-sm">
                <div className="w-3 h-3 rounded-full" style={{ backgroundColor: COLORS.highCost }} />
                High-cost
              </span>
              <span className="flex items-center gap-2 text-sm">
                <div className="w-3 h-3 rounded-full" style={{ backgroundColor: COLORS.untagged }} />
                Untagged
              </span>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Best-Cost Country Penetration by Function</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="h-[250px]">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={functionOffshoringStats.slice(0, 8)} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" />
                  <XAxis type="number" domain={[0, 100]} tickFormatter={(v) => `${v}%`} />
                  <YAxis type="category" dataKey="function" width={100} tick={{ fontSize: 12 }} />
                  <Tooltip 
                    formatter={(value: number) => `${value.toFixed(1)}%`}
                    contentStyle={{ 
                      backgroundColor: 'hsl(var(--card))', 
                      border: '1px solid hsl(var(--border))',
                      borderRadius: '8px',
                    }}
                  />
                  <Bar dataKey="bccPercent" fill="hsl(var(--success))" radius={[0, 4, 4, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Function Offshoring Analysis Table */}
      <Card>
        <CardHeader>
          <CardTitle>Offshoring Opportunity by Function</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="overflow-x-auto">
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Function</TableHead>
                  <TableHead className="text-right">Total HC</TableHead>
                  <TableHead className="text-right">BCC %</TableHead>
                  <TableHead className="text-right">HCC %</TableHead>
                  <TableHead className="text-right">Avg FLRR</TableHead>
                  <TableHead className="text-center">Incremental Offshoring %</TableHead>
                  <TableHead className="text-center">Benchmark %</TableHead>
                  <TableHead className="text-right">Potential Savings</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {functionOffshoringStats.map((func) => {
                  // Benchmark offshoring percentages based on function type
                  const getBenchmark = (funcName: string): number => {
                    const lower = funcName.toLowerCase();
                    if (lower.includes('it') || lower.includes('technology') || lower.includes('engineering')) return 40;
                    if (lower.includes('finance') || lower.includes('accounting')) return 35;
                    if (lower.includes('hr') || lower.includes('human')) return 30;
                    if (lower.includes('customer') || lower.includes('support') || lower.includes('service')) return 50;
                    if (lower.includes('operations') || lower.includes('admin')) return 35;
                    if (lower.includes('sales')) return 15;
                    if (lower.includes('marketing')) return 25;
                    if (lower.includes('legal')) return 15;
                    if (lower.includes('research') || lower.includes('r&d')) return 20;
                    return 25; // default benchmark
                  };
                  const benchmark = getBenchmark(func.function);
                  
                  return (
                    <TableRow key={func.function}>
                      <TableCell className="font-medium">{func.function}</TableCell>
                      <TableCell className="text-right">{func.totalHeadcount}</TableCell>
                      <TableCell className="text-right">
                        <span className={func.bccPercent > 40 ? 'text-success' : ''}>
                          {formatPercent(func.bccPercent)}
                        </span>
                      </TableCell>
                      <TableCell className="text-right">
                        <span className={func.hccPercent > 60 ? 'text-warning' : ''}>
                          {formatPercent(func.hccPercent)}
                        </span>
                      </TableCell>
                      <TableCell className="text-right">{formatCurrency(func.avgFLRR)}</TableCell>
                      <TableCell>
                        <div className="flex items-center justify-center">
                          <Input
                            type="number"
                            min="0"
                            max="100"
                            value={func.offshoringPotential}
                            onChange={(e) => handleOffshoringPotentialChange(func.function, e.target.value)}
                            className="w-20 text-center"
                            placeholder="0"
                          />
                          <span className="ml-1 text-muted-foreground">%</span>
                        </div>
                      </TableCell>
                      <TableCell className="text-center text-muted-foreground">
                        {benchmark}%
                      </TableCell>
                      <TableCell className="text-right font-medium text-primary">
                        {formatCurrency(func.potentialSavings)}
                      </TableCell>
                    </TableRow>
                  );
                })}
                {/* Total Row */}
                <TableRow className="bg-secondary/50 font-semibold border-t-2">
                  <TableCell>Total</TableCell>
                  <TableCell className="text-right">
                    {functionOffshoringStats.reduce((sum, f) => sum + f.totalHeadcount, 0)}
                  </TableCell>
                  <TableCell className="text-right">
                    {formatPercent(
                      functionOffshoringStats.reduce((sum, f) => sum + f.bccHeadcount, 0) / 
                      Math.max(1, functionOffshoringStats.reduce((sum, f) => sum + f.totalHeadcount, 0)) * 100
                    )}
                  </TableCell>
                  <TableCell className="text-right">
                    {formatPercent(
                      functionOffshoringStats.reduce((sum, f) => sum + f.hccHeadcount, 0) / 
                      Math.max(1, functionOffshoringStats.reduce((sum, f) => sum + f.totalHeadcount, 0)) * 100
                    )}
                  </TableCell>
                  <TableCell className="text-right">
                    {formatCurrency(
                      functionOffshoringStats.reduce((sum, f) => sum + f.avgFLRR * f.totalHeadcount, 0) / 
                      Math.max(1, functionOffshoringStats.reduce((sum, f) => sum + f.totalHeadcount, 0))
                    )}
                  </TableCell>
                  <TableCell className="text-center">-</TableCell>
                  <TableCell className="text-center">-</TableCell>
                  <TableCell className="text-right text-primary">
                    {formatCurrency(functionOffshoringStats.reduce((sum, f) => sum + f.potentialSavings, 0))}
                  </TableCell>
                </TableRow>
              </TableBody>
            </Table>
          </div>
        </CardContent>
      </Card>

      {/* Top Offshoring Roles */}
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <AlertTriangle className="w-5 h-5 text-primary" />
            Top 20 Job Titles with Offshoring Potential
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-sm text-muted-foreground mb-4">
            Based on industry research and best practices, the following roles are typically suitable for offshoring to lower-cost countries:
          </p>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {OFFSHORING_ROLES.map((role, index) => (
              <div key={role.title} className="p-3 rounded-lg border bg-card hover:bg-secondary/30 transition-colors">
                <div className="flex items-start gap-3">
                  <Badge variant="secondary" className="text-xs shrink-0">
                    #{index + 1}
                  </Badge>
                  <div>
                    <p className="font-medium text-sm">{role.title}</p>
                    <p className="text-xs text-muted-foreground">{role.reason}</p>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
