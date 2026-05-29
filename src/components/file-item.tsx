'use client';

import { FileText, Image as ImageIcon, GripVertical, Trash2, FileCode } from 'lucide-react';
import { cn } from '@/lib/utils';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Tooltip, TooltipContent, TooltipTrigger } from '@/components/ui/tooltip';
import type { ProcessedFile } from '@/lib/process-files';
import { formatFileSize } from '@/lib/process-files';

interface FileItemProps {
  file: ProcessedFile;
  index: number;
  onDelete: (id: string) => void;
  dragHandleProps?: {
    listeners?: Record<string, unknown>;
    attributes?: Record<string, unknown>;
    style?: React.CSSProperties;
    ref?: (el: HTMLElement | null) => void;
  };
  isDragging?: boolean;
}

function getFileIcon(category: string) {
  switch (category) {
    case 'PDF':
      return <FileText className="w-6 h-6 text-red-500" />;
    case 'Word':
      return <FileCode className="w-6 h-6 text-blue-600" />;
    case 'Image':
      return <ImageIcon className="w-6 h-6 text-green-600" />;
    default:
      return <FileText className="w-6 h-6 text-gray-500" />;
  }
}

export function FileItem({ file, index, onDelete, dragHandleProps, isDragging }: FileItemProps) {
  const isImage = file.category === 'Image' && file.preview && file.preview.startsWith('data:');

  return (
    <div
      className={cn(
        'flex items-center gap-3 p-3 rounded-lg transition-all duration-200',
        'bg-slate-50 hover:bg-purple-50/60 border-2 border-transparent hover:border-purple-200',
        isDragging && 'opacity-50 scale-[0.98]'
      )}
    >
      {/* Drag handle */}
      <div
        {...dragHandleProps?.listeners}
        {...dragHandleProps?.attributes}
        ref={dragHandleProps?.ref}
        className="flex-shrink-0 cursor-grab active:cursor-grabbing text-gray-400 hover:text-gray-600 transition-colors touch-none"
        aria-label="Drag to reorder"
      >
        <GripVertical className="w-5 h-5" />
      </div>

      {/* Preview */}
      <div className="flex-shrink-0 w-12 h-12 rounded-lg bg-gray-100 overflow-hidden flex items-center justify-center">
        {isImage ? (
          <img
            src={file.preview as string}
            alt={file.name}
            className="w-full h-full object-cover"
          />
        ) : (
          getFileIcon(file.category)
        )}
      </div>

      {/* Info */}
      <div className="flex-1 min-w-0">
        <div className="text-sm font-semibold text-gray-800 truncate">
          {index + 1}. {file.name}
        </div>
        <div className="text-xs text-gray-500 mt-0.5 flex items-center flex-wrap gap-1">
          <span>{formatFileSize(file.size)}</span>
          <span className="text-gray-300">|</span>
          <span>{file.category}</span>
          {file.pages > 0 && (
            <>
              <span className="text-gray-300">|</span>
              <span>{file.pages}页</span>
            </>
          )}
          {file.category === 'Word' && file.pages === -1 && (
            <>
              <span className="text-gray-300">|</span>
              <span>待合并计算</span>
            </>
          )}
          {file.width && file.height && (
            <>
              <span className="text-gray-300">|</span>
              <span>{Math.round(file.width)}×{Math.round(file.height)}</span>
            </>
          )}
        </div>
      </div>

      {/* Orientation badge */}
      <Tooltip>
        <TooltipTrigger asChild>
          <Badge
            variant="secondary"
            className={cn(
              'flex-shrink-0 text-xs font-semibold rounded-full',
              file.orientation === 'portrait'
                ? 'bg-green-100 text-green-700 hover:bg-green-100'
                : 'bg-amber-100 text-amber-700 hover:bg-amber-100'
            )}
          >
            {file.orientation === 'portrait' ? '竖向' : '横向'}
          </Badge>
        </TooltipTrigger>
        <TooltipContent>
          <p>页面方向：{file.orientation === 'portrait' ? '竖向（Portrait）' : '横向（Landscape）'}</p>
        </TooltipContent>
      </Tooltip>

      {/* Delete button */}
      <Button
        variant="ghost"
        size="sm"
        onClick={(e) => {
          e.stopPropagation();
          onDelete(file.id);
        }}
        className="flex-shrink-0 text-red-400 hover:text-red-600 hover:bg-red-50 h-8 w-8 p-0"
        aria-label={`Delete ${file.name}`}
      >
        <Trash2 className="w-4 h-4" />
      </Button>
    </div>
  );
}
