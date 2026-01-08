import { useState } from 'react';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { CheckCircle } from 'lucide-react';
import { ColumnMapping, fieldLabels } from '@/lib/excel-parser';

interface ColumnMappingModalProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  headers: string[];
  mapping: ColumnMapping;
  onConfirm: (mapping: ColumnMapping) => void;
}

const UNMAPPED = '__unmapped__';

export function ColumnMappingModal({
  open,
  onOpenChange,
  headers,
  mapping,
  onConfirm,
}: ColumnMappingModalProps) {
  const [currentMapping, setCurrentMapping] = useState<ColumnMapping>(mapping);

  const handleMappingChange = (field: keyof ColumnMapping, value: string) => {
    setCurrentMapping(prev => ({
      ...prev,
      [field]: value === UNMAPPED ? null : value,
    }));
  };

  const mappedCount = (Object.keys(currentMapping) as (keyof ColumnMapping)[]).filter(f => currentMapping[f]).length;
  const totalFields = Object.keys(fieldLabels).length;

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-w-2xl max-h-[80vh] overflow-y-auto">
        <DialogHeader>
          <DialogTitle>Column Mapping</DialogTitle>
        </DialogHeader>

        <div className="space-y-4">
          {/* Status Summary */}
          <p className="text-sm text-muted-foreground">
            {mappedCount} of {totalFields} fields mapped. Unmapped fields will be skipped in relevant analyses.
          </p>

          {/* Mapping Fields */}
          <div className="grid gap-3">
            {(Object.keys(fieldLabels) as (keyof ColumnMapping)[]).map(field => (
              <div key={field} className="flex items-center gap-3">
                <div className="w-56">
                  <span className="text-sm font-medium">{fieldLabels[field]}</span>
                </div>
                <Select
                  value={currentMapping[field] || UNMAPPED}
                  onValueChange={(v) => handleMappingChange(field, v)}
                >
                  <SelectTrigger className="flex-1">
                    <SelectValue placeholder="Select column..." />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value={UNMAPPED}>
                      <span className="text-muted-foreground">-- Not mapped --</span>
                    </SelectItem>
                    {headers.map(header => (
                      <SelectItem key={header} value={header}>
                        {header}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {currentMapping[field] && (
                  <CheckCircle className="w-4 h-4 text-success flex-shrink-0" />
                )}
              </div>
            ))}
          </div>

          {/* Manager ID Note */}
          {!currentMapping.managerId && (
            <div className="p-3 rounded-lg bg-secondary/50 border">
              <p className="text-sm text-muted-foreground">
                <strong>Note:</strong> Manager ID or Name is required to map organizational hierarchy for Spans & Layers analysis.
              </p>
            </div>
          )}
        </div>

        <DialogFooter>
          <Button variant="outline" onClick={() => onOpenChange(false)}>
            Cancel
          </Button>
          <Button onClick={() => onConfirm(currentMapping)}>
            Confirm & Analyze
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
