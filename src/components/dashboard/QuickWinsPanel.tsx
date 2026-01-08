import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Lightbulb, TrendingUp, Users, DollarSign, Layers } from 'lucide-react';
import { QuickWin } from '@/types/employee';
import { useMemo } from 'react';

interface QuickWinsPanelProps {
  quickWins: QuickWin[];
}

const categoryIcons = {
  offshoring: DollarSign,
  spans: Users,
  compensation: TrendingUp,
  structure: Layers,
};

const impactColors = {
  high: 'bg-destructive/10 text-destructive border-destructive/30',
  medium: 'bg-warning/10 text-warning border-warning/30',
  low: 'bg-success/10 text-success border-success/30',
};

interface DisplayQuickWin extends QuickWin {
  subItems?: QuickWin[];
}

export function QuickWinsPanel({ quickWins }: QuickWinsPanelProps) {
  // Group compensation-related quick wins into a single consolidated entry
  const displayQuickWins = useMemo((): DisplayQuickWin[] => {
    const compensationWins = quickWins.filter(w => w.category === 'compensation');
    const otherWins = quickWins.filter(w => w.category !== 'compensation');

    const result: DisplayQuickWin[] = [...otherWins];

    if (compensationWins.length > 0) {
      // Determine highest impact among compensation wins
      const hasHighImpact = compensationWins.some(w => w.impact === 'high');
      const hasMediumImpact = compensationWins.some(w => w.impact === 'medium');
      const groupImpact: 'high' | 'medium' | 'low' = hasHighImpact ? 'high' : hasMediumImpact ? 'medium' : 'low';

      result.push({
        id: 'comp-grouped',
        title: 'Variable Compensation Optimization (multiple initiatives)',
        description: `${compensationWins.length} function${compensationWins.length > 1 ? 's' : ''} with variable compensation below target.`,
        impact: groupImpact,
        category: 'compensation',
        subItems: compensationWins,
      });
    }

    // Sort by impact
    const impactOrder = { high: 0, medium: 1, low: 2 };
    return result.sort((a, b) => impactOrder[a.impact] - impactOrder[b.impact]);
  }, [quickWins]);

  if (quickWins.length === 0) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
          <Lightbulb className="w-5 h-5 text-accent" />
            Value Creation Opportunities
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-muted-foreground text-center py-4">
            No immediate optimization opportunities identified
          </p>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Lightbulb className="w-5 h-5 text-accent" />
          Value Creation Opportunities
          <Badge variant="secondary" className="ml-2">
            {quickWins.length} opportunities
          </Badge>
        </CardTitle>
      </CardHeader>
      <CardContent className="space-y-4">
        {displayQuickWins.map((win) => {
          const Icon = categoryIcons[win.category];
          return (
            <div 
              key={win.id} 
              className="p-4 rounded-lg border bg-card hover:bg-secondary/30 transition-colors"
            >
              <div className="flex items-start gap-3">
                <div className="p-2 rounded-md bg-primary/10">
                  <Icon className="w-4 h-4 text-primary" />
                </div>
                <div className="flex-1 space-y-1">
                  <div className="flex items-center gap-2">
                    <h4 className="font-medium text-foreground">{win.title}</h4>
                    <Badge className={impactColors[win.impact]}>
                      {win.impact} impact
                    </Badge>
                  </div>
                  <p className="text-sm text-muted-foreground">{win.description}</p>
                  {win.metric && (
                    <p className="text-sm font-medium text-primary">{win.metric}</p>
                  )}
                  {/* Show sub-items for grouped compensation wins */}
                  {win.subItems && win.subItems.length > 0 && (
                    <ul className="mt-2 space-y-1 pl-2 border-l-2 border-primary/20">
                      {win.subItems.map((subItem) => (
                        <li key={subItem.id} className="text-sm text-muted-foreground">
                          <span className="font-medium text-foreground">{subItem.title.replace('Low Variable Comp: ', '')}</span>
                          {subItem.metric && <span className="text-primary ml-1">({subItem.metric})</span>}
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>
            </div>
          );
        })}
      </CardContent>
    </Card>
  );
}