import { useState, useMemo } from 'react';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Button } from '@/components/ui/button';
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from '@/components/ui/alert-dialog';
import { Employee, AnalysisData } from '@/types/employee';
import { analyzeEmployeeData } from '@/lib/analysis';
import { ExecutiveSummary } from './ExecutiveSummary';
import { SpansLayersAnalysis } from './SpansLayersAnalysis';
import { OffshoringAnalysis } from './OffshoringAnalysis';
import { CompensationAnalysis } from './CompensationAnalysis';
import { TenureAnalysis } from './TenureAnalysis';
import { HeadcountBreakdown } from './HeadcountBreakdown';
import { OrgChart } from './OrgChart';
import { AutomationAnalysis } from './AutomationAnalysis';
import { exportToExcel } from '@/lib/excel-export';
import { exportToPDF } from '@/lib/pdf-export';
import { 
  LayoutDashboard, 
  Layers, 
  Globe, 
  DollarSign, 
  Clock, 
  Users,
  Upload,
  FileSpreadsheet,
  FileText,
  Network,
  Bot,
  Loader2
} from 'lucide-react';
import { toast } from 'sonner';

interface DashboardProps {
  employees: Employee[];
  onReset: () => void;
}

export function Dashboard({ employees, onReset }: DashboardProps) {
  const [activeTab, setActiveTab] = useState('summary');
  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [isExportingPDF, setIsExportingPDF] = useState(false);
  const [showLeaveConfirmation, setShowLeaveConfirmation] = useState(false);

  const analysisData: AnalysisData = useMemo(() => {
    return analyzeEmployeeData(employees);
  }, [employees]);

  const handleExportExcel = async () => {
    setIsExportingExcel(true);
    try {
      await exportToExcel(analysisData);
      toast.success('Excel export completed');
    } catch (error) {
      console.error('Excel export failed:', error);
      toast.error('Failed to export Excel file');
    } finally {
      setIsExportingExcel(false);
    }
  };

  const handleExportPDF = async () => {
    setIsExportingPDF(true);
    try {
      await exportToPDF(analysisData);
      toast.success('PDF export completed');
    } catch (error) {
      console.error('PDF export failed:', error);
      toast.error('Failed to export PDF file');
    } finally {
      setIsExportingPDF(false);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      {/* Header */}
      <header className="sticky top-0 z-50 border-b bg-card/95 backdrop-blur supports-[backdrop-filter]:bg-card/60">
        <div className="container mx-auto px-4 py-4">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold text-foreground">
                Organizational Due Diligence
              </h1>
              <p className="text-sm text-muted-foreground">
                {analysisData.totals.headcount} employees analyzed
              </p>
            </div>
            <div className="flex items-center gap-3">
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleExportExcel}
                disabled={isExportingExcel}
              >
                {isExportingExcel ? (
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                ) : (
                  <FileSpreadsheet className="w-4 h-4 mr-2" />
                )}
                Export Excel
              </Button>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={handleExportPDF}
                disabled={isExportingPDF}
              >
                {isExportingPDF ? (
                  <Loader2 className="w-4 h-4 mr-2 animate-spin" />
                ) : (
                  <FileText className="w-4 h-4 mr-2" />
                )}
                Export PDF
              </Button>
              <Button variant="ghost" size="sm" onClick={() => setShowLeaveConfirmation(true)}>
                <Upload className="w-4 h-4 mr-2" />
                New Analysis
              </Button>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="container mx-auto px-4 py-6">
        <Tabs value={activeTab} onValueChange={setActiveTab} className="w-full">
          <TabsList className="grid w-full grid-cols-8 mb-6">
            <TabsTrigger value="summary" className="flex items-center gap-2">
              <LayoutDashboard className="w-4 h-4" />
              <span className="hidden md:inline">Summary</span>
            </TabsTrigger>
            <TabsTrigger value="headcount" className="flex items-center gap-2">
              <Users className="w-4 h-4" />
              <span className="hidden md:inline">Headcount</span>
            </TabsTrigger>
            <TabsTrigger value="orgchart" className="flex items-center gap-2">
              <Network className="w-4 h-4" />
              <span className="hidden md:inline">Org Chart</span>
            </TabsTrigger>
            <TabsTrigger value="spans" className="flex items-center gap-2">
              <Layers className="w-4 h-4" />
              <span className="hidden md:inline">Spans & Layers</span>
            </TabsTrigger>
            <TabsTrigger value="tenure" className="flex items-center gap-2">
              <Clock className="w-4 h-4" />
              <span className="hidden md:inline">Tenure</span>
            </TabsTrigger>
            <TabsTrigger value="automation" className="flex items-center gap-2">
              <Bot className="w-4 h-4" />
              <span className="hidden md:inline">Automation</span>
            </TabsTrigger>
            <TabsTrigger value="offshoring" className="flex items-center gap-2">
              <Globe className="w-4 h-4" />
              <span className="hidden md:inline">Offshoring</span>
            </TabsTrigger>
            <TabsTrigger value="compensation" className="flex items-center gap-2">
              <DollarSign className="w-4 h-4" />
              <span className="hidden md:inline">Compensation</span>
            </TabsTrigger>
          </TabsList>

          <TabsContent value="summary">
            <ExecutiveSummary data={analysisData} />
          </TabsContent>

          <TabsContent value="orgchart">
            <OrgChart data={analysisData} />
          </TabsContent>

          <TabsContent value="spans">
            <SpansLayersAnalysis data={analysisData} />
          </TabsContent>

          <TabsContent value="offshoring">
            <OffshoringAnalysis data={analysisData} />
          </TabsContent>

          <TabsContent value="compensation">
            <CompensationAnalysis data={analysisData} />
          </TabsContent>

          <TabsContent value="tenure">
            <TenureAnalysis data={analysisData} />
          </TabsContent>

          <TabsContent value="automation">
            <AutomationAnalysis data={analysisData} />
          </TabsContent>

          <TabsContent value="headcount">
            <HeadcountBreakdown data={analysisData} />
          </TabsContent>
        </Tabs>
      </main>

      {/* Leave Page Confirmation Dialog */}
      <AlertDialog open={showLeaveConfirmation} onOpenChange={setShowLeaveConfirmation}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>Are you sure you'd like to leave this page?</AlertDialogTitle>
            <AlertDialogDescription>
              You will lose the current version of the dashboards, but you can export them to PDF or Excel.
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>Cancel</AlertDialogCancel>
            <AlertDialogAction
              onClick={onReset}
              className="bg-destructive text-destructive-foreground hover:bg-destructive/90"
            >
              Leave page
            </AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </div>
  );
}
