'use client';

import { Settings } from 'lucide-react';
import { Label } from '@/components/ui/label';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';

export type PageSizeOption = 'a4' | 'letter' | 'a3' | 'auto';
export type OrientationOption = 'all-portrait' | 'all-landscape' | 'keep-original';
export type QualityOption = '0.9' | '0.7' | '0.5';

interface MergeOptionsProps {
  pageSize: PageSizeOption;
  orientation: OrientationOption;
  imageQuality: QualityOption;
  hasWordFiles: boolean;
  onPageSizeChange: (value: PageSizeOption) => void;
  onOrientationChange: (value: OrientationOption) => void;
  onImageQualityChange: (value: QualityOption) => void;
}

export function MergeOptions({
  pageSize,
  orientation,
  imageQuality,
  hasWordFiles,
  onPageSizeChange,
  onOrientationChange,
  onImageQualityChange,
}: MergeOptionsProps) {
  return (
    <div className="bg-slate-50/60 rounded-xl p-4 md:p-5">
      <div className="flex items-center gap-2 mb-4">
        <Settings className="w-5 h-5 text-purple-600" />
        <h3 className="text-base font-semibold text-gray-800">合并选项</h3>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        {/* Paper Size */}
        <div className="space-y-2">
          <Label htmlFor="pageSize" className="text-sm font-medium text-gray-600">
            目标纸张大小
          </Label>
          <Select
            value={pageSize}
            onValueChange={(v) => onPageSizeChange(v as PageSizeOption)}
          >
            <SelectTrigger id="pageSize" className="w-full">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="a4">A4 (210×297mm)</SelectItem>
              <SelectItem value="letter">Letter (216×279mm)</SelectItem>
              <SelectItem value="a3">A3 (297×420mm)</SelectItem>
              <SelectItem value="auto" disabled={hasWordFiles}>
                自动（仅限PDF和图片）
              </SelectItem>
            </SelectContent>
          </Select>
          {hasWordFiles && pageSize === 'auto' && (
            <p className="text-xs text-amber-600">含Word文档时自动使用A4</p>
          )}
        </div>

        {/* Orientation */}
        <div className="space-y-2">
          <Label htmlFor="orientation" className="text-sm font-medium text-gray-600">
            页面方向处理
          </Label>
          <Select
            value={orientation}
            onValueChange={(v) => onOrientationChange(v as OrientationOption)}
          >
            <SelectTrigger id="orientation" className="w-full">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="all-portrait">统一为竖向（适合打印）</SelectItem>
              <SelectItem value="all-landscape">统一为横向</SelectItem>
              <SelectItem value="keep-original">保持原方向</SelectItem>
            </SelectContent>
          </Select>
        </div>

        {/* Image Quality */}
        <div className="space-y-2">
          <Label htmlFor="imageQuality" className="text-sm font-medium text-gray-600">
            图片质量
          </Label>
          <Select
            value={imageQuality}
            onValueChange={(v) => onImageQualityChange(v as QualityOption)}
          >
            <SelectTrigger id="imageQuality" className="w-full">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="0.9">高质量</SelectItem>
              <SelectItem value="0.7">标准</SelectItem>
              <SelectItem value="0.5">较小文件</SelectItem>
            </SelectContent>
          </Select>
        </div>
      </div>
    </div>
  );
}
