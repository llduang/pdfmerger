'use client';

import { cn } from '@/lib/utils';

interface MergeProgressProps {
  visible: boolean;
  percent: number;
  fileName: string;
}

export function MergeProgress({ visible, percent, fileName }: MergeProgressProps) {
  if (!visible) return null;

  return (
    <div className="mt-5 space-y-3">
      <div className="relative h-5 w-full overflow-hidden rounded-full bg-gray-200">
        <div
          className={cn(
            'h-full rounded-full transition-all duration-300 ease-out',
            'bg-gradient-to-r from-purple-600 to-purple-500'
          )}
          style={{ width: `${Math.min(percent, 100)}%` }}
        />
        <div className="absolute inset-0 flex items-center justify-center">
          <span className="text-xs font-bold text-white drop-shadow-sm">
            {Math.round(percent)}%
          </span>
        </div>
      </div>
      <p className="text-center text-sm text-gray-500 font-medium">
        正在处理: {fileName}
      </p>
    </div>
  );
}
