import { Employee, SpanStats, Benchmarks } from '@/types/employee';
import { defaultBenchmarks } from './analysis';

export interface ManagerWithDetails extends SpanStats {
  directReportTitle?: string;
  directReportTitles?: string;
}

export interface SpansLayersOutliers {
  singleReportManagers: ManagerWithDetails[];
  belowMinSpanManagers: ManagerWithDetails[];
  aboveMaxSpanManagers: ManagerWithDetails[];
}

/**
 * Shared helper for calculating spans & layers outliers.
 * Used by UI (SpansLayersAnalysis), PDF export, and Excel export for consistency.
 */
export function calculateSpansLayersOutliers(
  spanStats: SpanStats[],
  employees: Employee[],
  benchmarks: Benchmarks = defaultBenchmarks
): SpansLayersOutliers {
  // Single-report managers with direct report title
  const singleReportManagers = spanStats
    .filter(s => s.directReports === 1)
    .map(mgr => {
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitle: directReports[0]?.title || 'Unknown',
      };
    });

  // Managers below minimum span with direct report titles
  const belowMinSpanManagers = spanStats
    .filter(s => s.directReports > 1 && s.directReports < benchmarks.minSpan)
    .map(mgr => {
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitles: directReports.map(dr => dr.title).join(', '),
      };
    });

  // Managers above maximum span with direct report titles
  const aboveMaxSpanManagers = spanStats
    .filter(s => s.directReports > benchmarks.maxSpan)
    .map(mgr => {
      const directReports = employees.filter(emp => emp.managerId === mgr.managerId);
      return {
        ...mgr,
        directReportTitles: directReports.map(dr => dr.title).join(', '),
      };
    });

  return {
    singleReportManagers,
    belowMinSpanManagers,
    aboveMaxSpanManagers,
  };
}
