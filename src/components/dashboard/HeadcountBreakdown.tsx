import { useState, useMemo } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Treemap, Cell, LabelList } from 'recharts';
import { AnalysisData } from '@/types/employee';
import { formatCurrency, formatNumber } from '@/lib/analysis';
import { ToggleGroup, ToggleGroupItem } from '@/components/ui/toggle-group';
import { Users, DollarSign } from 'lucide-react';

interface HeadcountBreakdownProps {
  data: AnalysisData;
}

// Custom label renderer that positions labels inside or outside bars based on space
const createCustomLabel = (metric: 'headcount' | 'flrr', maxValue: number) => {
  return (props: any) => {
    const { x, y, width, height, value } = props;
    if (value === undefined || value === null) return null;

    const displayValue = metric === 'headcount' ? formatNumber(value) : formatCurrency(value);
    const labelWidth = displayValue.length * 7; // Approximate width per character
    const barWidth = width || 0;
    const threshold = maxValue * 0.15; // 15% of max value as threshold for inside/outside

    // Determine if label should be inside or outside the bar
    const isInsideBar = value > threshold && barWidth > labelWidth + 10;
    
    if (isInsideBar) {
      return (
        <text
          x={x + barWidth - 8}
          y={y + height / 2}
          textAnchor="end"
          dominantBaseline="middle"
          fill="hsl(var(--primary-foreground))"
          fontSize={11}
          fontWeight={500}
        >
          {displayValue}
        </text>
      );
    } else {
      return (
        <text
          x={x + barWidth + 6}
          y={y + height / 2}
          textAnchor="start"
          dominantBaseline="middle"
          fill="hsl(var(--foreground))"
          fontSize={11}
          fontWeight={500}
        >
          {displayValue}
        </text>
      );
    }
  };
};

export function HeadcountBreakdown({ data }: HeadcountBreakdownProps) {
  const [metric, setMetric] = useState<'headcount' | 'flrr'>('headcount');
  const { employees, functionStats } = data;

  // Location data
  const locationStats = useMemo(() => Array.from(
    employees.reduce((map, emp) => {
      const loc = emp.location;
      const existing = map.get(loc) || { location: loc, headcount: 0, flrr: 0 };
      existing.headcount++;
      existing.flrr += emp.flrr;
      map.set(loc, existing);
      return map;
    }, new Map<string, { location: string; headcount: number; flrr: number }>())
  ).map(([, v]) => v).sort((a, b) => b.headcount - a.headcount), [employees]);

  // Business unit data
  const businessUnitStats = useMemo(() => Array.from(
    employees.reduce((map, emp) => {
      const bu = emp.businessUnit || 'Unknown';
      const existing = map.get(bu) || { businessUnit: bu, headcount: 0, flrr: 0 };
      existing.headcount++;
      existing.flrr += emp.flrr;
      map.set(bu, existing);
      return map;
    }, new Map<string, { businessUnit: string; headcount: number; flrr: number }>())
  ).map(([, v]) => v).sort((a, b) => b.headcount - a.headcount), [employees]);

  // Calculate max values for each dataset for label positioning
  const maxFunctionValue = useMemo(() => {
    return Math.max(...functionStats.map(d => metric === 'headcount' ? d.headcount : d.totalFLRR));
  }, [functionStats, metric]);

  const maxLocationValue = useMemo(() => {
    return Math.max(...locationStats.map(d => metric === 'headcount' ? d.headcount : d.flrr));
  }, [locationStats, metric]);

  const maxBusinessUnitValue = useMemo(() => {
    return Math.max(...businessUnitStats.map(d => metric === 'headcount' ? d.headcount : d.flrr));
  }, [businessUnitStats, metric]);

  // Calculate dynamic chart height based on number of items
  const getChartHeight = (itemCount: number) => Math.max(400, itemCount * 35);

  // Treemap data for function breakdown
  const treemapData = functionStats.map(d => ({
    name: d.function,
    size: metric === 'headcount' ? d.headcount : d.totalFLRR,
    headcount: d.headcount,
    flrr: d.totalFLRR,
  }));

  const CustomTreemapContent = (props: any) => {
    const { x, y, width, height, name, headcount, flrr } = props;
    if (width < 50 || height < 30 || !name) return null;

    return (
      <g>
        <rect
          x={x}
          y={y}
          width={width}
          height={height}
          fill="hsl(var(--primary))"
          stroke="hsl(var(--background))"
          strokeWidth={2}
          rx={4}
        />
        <text
          x={x + width / 2}
          y={y + height / 2 - 8}
          textAnchor="middle"
          fill="hsl(var(--primary-foreground))"
          fontSize={width > 100 ? 12 : 10}
          fontWeight={500}
        >
          {name}
        </text>
        <text
          x={x + width / 2}
          y={y + height / 2 + 8}
          textAnchor="middle"
          fill="hsl(var(--primary-foreground))"
          fontSize={10}
          opacity={0.8}
        >
          {metric === 'headcount' ? `${headcount ?? 0}` : formatCurrency(flrr ?? 0)}
        </text>
      </g>
    );
  };

  return (
    <div className="space-y-6 animate-fade-in">
      {/* Metric Toggle */}
      <div className="flex justify-end">
        <ToggleGroup type="single" value={metric} onValueChange={(v) => v && setMetric(v as 'headcount' | 'flrr')}>
          <ToggleGroupItem value="headcount" aria-label="Headcount view">
            <Users className="w-4 h-4 mr-2" />
            Headcount
          </ToggleGroupItem>
          <ToggleGroupItem value="flrr" aria-label="FLRR view">
            <DollarSign className="w-4 h-4 mr-2" />
            FLRR
          </ToggleGroupItem>
        </ToggleGroup>
      </div>

      {/* Treemap Overview */}
      <Card>
        <CardHeader>
          <CardTitle>Function Overview</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="h-[400px]">
            <ResponsiveContainer width="100%" height="100%">
              <Treemap
                data={treemapData}
                dataKey="size"
                aspectRatio={4 / 3}
                content={<CustomTreemapContent />}
              />
            </ResponsiveContainer>
          </div>
        </CardContent>
      </Card>

      {/* Detailed Breakdowns */}
      <Tabs defaultValue="function" className="w-full">
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="function">Function</TabsTrigger>
          <TabsTrigger value="location">Location</TabsTrigger>
          <TabsTrigger value="businessUnit">Business Unit</TabsTrigger>
        </TabsList>

        <TabsContent value="function" className="mt-4">
          <Card>
            <CardContent className="pt-6">
              <div style={{ height: getChartHeight(functionStats.length) }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart 
                    data={functionStats} 
                    layout="vertical"
                    margin={{ top: 5, right: 80, left: 10, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" horizontal={false} />
                    <XAxis 
                      type="number" 
                      tickFormatter={(v) => metric === 'headcount' ? String(v) : formatCurrency(v)}
                      domain={[0, 'dataMax']}
                    />
                    <YAxis type="category" dataKey="function" width={120} tick={{ fontSize: 12 }} />
                    <Tooltip 
                      formatter={(value: number) => metric === 'headcount' ? formatNumber(value) : formatCurrency(value)}
                      contentStyle={{ 
                        backgroundColor: 'hsl(var(--card))', 
                        border: '1px solid hsl(var(--border))',
                        borderRadius: '8px',
                      }}
                    />
                    <Bar 
                      dataKey={metric === 'headcount' ? 'headcount' : 'totalFLRR'} 
                      fill="hsl(var(--primary))" 
                      radius={[0, 4, 4, 0]}
                      label={createCustomLabel(metric, maxFunctionValue)}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="location" className="mt-4">
          <Card>
            <CardContent className="pt-6">
              <div style={{ height: getChartHeight(Math.min(locationStats.length, 15)) }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart 
                    data={locationStats.slice(0, 15)} 
                    layout="vertical"
                    margin={{ top: 5, right: 80, left: 10, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" horizontal={false} />
                    <XAxis 
                      type="number" 
                      tickFormatter={(v) => metric === 'headcount' ? String(v) : formatCurrency(v)}
                      domain={[0, 'dataMax']}
                    />
                    <YAxis type="category" dataKey="location" width={120} tick={{ fontSize: 12 }} />
                    <Tooltip 
                      formatter={(value: number) => metric === 'headcount' ? formatNumber(value) : formatCurrency(value)}
                      contentStyle={{ 
                        backgroundColor: 'hsl(var(--card))', 
                        border: '1px solid hsl(var(--border))',
                        borderRadius: '8px',
                      }}
                    />
                    <Bar 
                      dataKey={metric} 
                      fill="hsl(var(--chart-3))" 
                      radius={[0, 4, 4, 0]}
                      label={createCustomLabel(metric, maxLocationValue)}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </CardContent>
          </Card>
        </TabsContent>

        <TabsContent value="businessUnit" className="mt-4">
          <Card>
            <CardContent className="pt-6">
              <div style={{ height: getChartHeight(Math.min(businessUnitStats.length, 15)) }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart 
                    data={businessUnitStats.slice(0, 15)} 
                    layout="vertical"
                    margin={{ top: 5, right: 80, left: 10, bottom: 5 }}
                  >
                    <CartesianGrid strokeDasharray="3 3" stroke="hsl(var(--border))" horizontal={false} />
                    <XAxis 
                      type="number" 
                      tickFormatter={(v) => metric === 'headcount' ? String(v) : formatCurrency(v)}
                      domain={[0, 'dataMax']}
                    />
                    <YAxis type="category" dataKey="businessUnit" width={120} tick={{ fontSize: 12 }} />
                    <Tooltip 
                      formatter={(value: number) => metric === 'headcount' ? formatNumber(value) : formatCurrency(value)}
                      contentStyle={{ 
                        backgroundColor: 'hsl(var(--card))', 
                        border: '1px solid hsl(var(--border))',
                        borderRadius: '8px',
                      }}
                    />
                    <Bar 
                      dataKey={metric} 
                      fill="hsl(var(--chart-4))" 
                      radius={[0, 4, 4, 0]}
                      label={createCustomLabel(metric, maxBusinessUnitValue)}
                    />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </CardContent>
          </Card>
        </TabsContent>
      </Tabs>

      {/* Summary Table */}
      <Card>
        <CardHeader>
          <CardTitle>Function Summary</CardTitle>
        </CardHeader>
        <CardContent>
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="border-b">
                  <th className="text-left py-3 px-4 font-medium">Function</th>
                  <th className="text-right py-3 px-4 font-medium">Headcount</th>
                  <th className="text-right py-3 px-4 font-medium">Total FLRR</th>
                  <th className="text-right py-3 px-4 font-medium">Avg FLRR</th>
                  <th className="text-right py-3 px-4 font-medium">Best-Cost %</th>
                </tr>
              </thead>
              <tbody>
                {functionStats.map((func) => (
                  <tr key={func.function} className="border-b hover:bg-secondary/30">
                    <td className="py-3 px-4 font-medium">{func.function}</td>
                    <td className="py-3 px-4 text-right">{formatNumber(func.headcount)}</td>
                    <td className="py-3 px-4 text-right">{formatCurrency(func.totalFLRR)}</td>
                    <td className="py-3 px-4 text-right">{formatCurrency(func.avgFLRR)}</td>
                    <td className="py-3 px-4 text-right">
                      <span className={func.bestCostPercent > 40 ? 'text-success' : 'text-muted-foreground'}>
                        {func.bestCostPercent.toFixed(0)}%
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
