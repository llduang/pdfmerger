'use client';

import { useCallback } from 'react';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { Inbox } from 'lucide-react';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Badge } from '@/components/ui/badge';
import { FileItem } from './file-item';
import type { ProcessedFile } from '@/lib/process-files';

interface FileListProps {
  files: ProcessedFile[];
  onReorder: (fromIndex: number, toIndex: number) => void;
  onDelete: (id: string) => void;
}

function SortableFileItem({
  file,
  index,
  onDelete,
}: {
  file: ProcessedFile;
  index: number;
  onDelete: (id: string) => void;
}) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: file.id });

  const style: React.CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
    opacity: isDragging ? 0.5 : 1,
    zIndex: isDragging ? 50 : undefined,
  };

  return (
    <FileItem
      file={file}
      index={index}
      onDelete={onDelete}
      dragHandleProps={{
        listeners,
        attributes,
        style,
        ref: setNodeRef,
      }}
      isDragging={isDragging}
    />
  );
}

export function FileList({ files, onReorder, onDelete }: FileListProps) {
  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 5,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const handleDragEnd = useCallback(
    (event: DragEndEvent) => {
      const { active, over } = event;
      if (over && active.id !== over.id) {
        const oldIndex = files.findIndex((f) => f.id === active.id);
        const newIndex = files.findIndex((f) => f.id === over.id);
        if (oldIndex !== -1 && newIndex !== -1) {
          onReorder(oldIndex, newIndex);
        }
      }
    },
    [files, onReorder]
  );

  if (files.length === 0) {
    return (
      <div className="text-center py-10 text-gray-400">
        <Inbox className="w-12 h-12 mx-auto mb-3 text-gray-300" />
        <p>还没有添加文件</p>
      </div>
    );
  }

  return (
    <div>
      <DndContext
        sensors={sensors}
        collisionDetection={closestCenter}
        onDragEnd={handleDragEnd}
      >
        <SortableContext
          items={files.map((f) => f.id)}
          strategy={verticalListSortingStrategy}
        >
          <ScrollArea className="max-h-96">
            <div className="flex flex-col gap-2 pr-3">
              {files.map((file, index) => (
                <SortableFileItem
                  key={file.id}
                  file={file}
                  index={index}
                  onDelete={onDelete}
                />
              ))}
            </div>
          </ScrollArea>
        </SortableContext>
      </DndContext>
    </div>
  );
}

interface FileListHeaderProps {
  fileCount: number;
}

export function FileListHeader({ fileCount }: FileListHeaderProps) {
  return (
    <div className="flex items-center justify-between mb-3">
      <div className="flex items-center gap-2">
        <h3 className="text-base font-semibold text-gray-800">文件列表</h3>
        <Badge className="bg-purple-600 hover:bg-purple-700 text-white rounded-full px-2.5 py-0.5 text-xs">
          {fileCount} 个文件
        </Badge>
      </div>
    </div>
  );
}
