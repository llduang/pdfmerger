'use client';

import { useCallback, useRef, useState } from 'react';
import { Upload } from 'lucide-react';
import { cn } from '@/lib/utils';

interface FileUploadZoneProps {
  onFilesSelected: (files: FileList) => void;
  disabled?: boolean;
}

const ACCEPTED_TYPES = '.pdf,.docx,.jpg,.jpeg,.png,.gif,.bmp,.webp';

export function FileUploadZone({ onFilesSelected, disabled }: FileUploadZoneProps) {
  const [isDragOver, setIsDragOver] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const openFilePicker = useCallback(() => {
    if (disabled) return;
    const input = inputRef.current;
    if (!input) return;
    // Reset value so re-selecting the same files works
    input.value = '';
    input.click();
  }, [disabled]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) setIsDragOver(true);
  }, [disabled]);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
    if (disabled) return;
    if (e.dataTransfer.files.length > 0) {
      onFilesSelected(e.dataTransfer.files);
    }
  }, [disabled, onFilesSelected]);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      onFilesSelected(e.target.files);
    }
  }, [onFilesSelected]);

  return (
    <>
      {/* Hidden file input — completely detached from the clickable area */}
      <input
        ref={inputRef}
        type="file"
        multiple
        accept={ACCEPTED_TYPES}
        onChange={handleInputChange}
        disabled={disabled}
        style={{ display: 'none' }}
        tabIndex={-1}
      />

      {/* Clickable upload zone — uses onClick to programmatically open the file picker */}
      <div
        onClick={openFilePicker}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        role="button"
        tabIndex={0}
        aria-label="拖拽文件到这里，或点击选择"
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            openFilePicker();
          }
        }}
        className={cn(
          'rounded-xl border-[3px] border-dashed p-10 md:p-16 text-center cursor-pointer transition-all duration-300 select-none',
          isDragOver
            ? 'border-purple-500 bg-purple-50/50 scale-[1.01]'
            : 'border-purple-300 bg-slate-50/50 hover:border-purple-400 hover:bg-purple-50/30',
          disabled && 'opacity-50 cursor-not-allowed pointer-events-none'
        )}
      >
        <div className="flex flex-col items-center gap-4">
          <div className={cn(
            'flex items-center justify-center w-16 h-16 rounded-full transition-colors duration-300',
            isDragOver ? 'bg-purple-200' : 'bg-purple-100'
          )}>
            <Upload className={cn(
              'w-8 h-8 transition-colors duration-300',
              isDragOver ? 'text-purple-700' : 'text-purple-500'
            )} />
          </div>

          <div>
            <h3 className="text-lg font-semibold text-gray-800 mb-1">
              {isDragOver ? '松开鼠标添加文件' : '拖拽文件到这里，或点击选择'}
            </h3>
            <p className="text-sm text-gray-500">
              支持 PDF、Word(.docx)、JPG、PNG、GIF、BMP、WEBP 格式
            </p>
          </div>
        </div>
      </div>
    </>
  );
}
