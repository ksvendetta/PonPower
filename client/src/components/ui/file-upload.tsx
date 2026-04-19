import { useRef, useState } from 'react';
import { cn } from '@/lib/utils';
import { Upload, FileSpreadsheet, X, CheckCircle2 } from 'lucide-react';

interface FileUploadProps {
  onFileSelect: (file: File, handle?: FileSystemFileHandle | null) => void;
  accept?: string;
  className?: string;
  label?: string;
  helper?: string;
  /**
   * File System Access API "types" filter for showOpenFilePicker.
   * When supplied, clicking the tile opens the native picker (returning a
   * writable handle). Drag-and-drop still works via dataTransfer.
   */
  pickerTypes?: Array<{
    description?: string;
    accept: Record<string, string[]>;
  }>;
}

export function FileUpload({
  onFileSelect,
  accept = ".xlsx, .xls",
  className,
  label = "Upload Excel Sheet",
  helper = "Drag and drop your .xlsx file here, or click to browse",
  pickerTypes,
}: FileUploadProps) {
  const [dragActive, setDragActive] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    const items = e.dataTransfer.items;
    if (items && items.length > 0) {
      const item = items[0];
      // Try to get a writable handle from the drop (Chromium)
      const anyItem = item as any;
      if (typeof anyItem.getAsFileSystemHandle === "function") {
        try {
          const handle = await anyItem.getAsFileSystemHandle();
          if (handle && handle.kind === "file") {
            const file = await (handle as FileSystemFileHandle).getFile();
            emit(file, handle as FileSystemFileHandle);
            return;
          }
        } catch {
          // fall through to File-only path
        }
      }
    }
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      emit(e.dataTransfer.files[0], null);
    }
  };

  const handleClick = async () => {
    // Prefer showOpenFilePicker so we get a writable handle back.
    const w = window as any;
    if (pickerTypes && typeof w.showOpenFilePicker === "function") {
      try {
        const [handle] = await w.showOpenFilePicker({
          multiple: false,
          types: pickerTypes,
          excludeAcceptAllOption: false,
        });
        const file = await handle.getFile();
        emit(file, handle);
        return;
      } catch (err: any) {
        if (err && err.name === "AbortError") return;
        // fall through to input click on other failures
      }
    }
    inputRef.current?.click();
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      emit(e.target.files[0], null);
    }
  };

  const emit = (file: File, handle: FileSystemFileHandle | null) => {
    setFileName(file.name);
    onFileSelect(file, handle);
  };

  const clearFile = (e: React.MouseEvent) => {
    e.stopPropagation();
    setFileName(null);
    if (inputRef.current) {
      inputRef.current.value = '';
    }
  };

  return (
    <div className={cn("w-full", className)}>
      <div
        className={cn(
          "relative flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-xl transition-all duration-200 ease-in-out cursor-pointer overflow-hidden group",
          dragActive
            ? "border-primary bg-primary/5"
            : "border-border hover:border-primary/50 hover:bg-secondary/30",
          fileName && "border-green-500/50 bg-green-500/5"
        )}
        onDragEnter={handleDrag}
        onDragLeave={handleDrag}
        onDragOver={handleDrag}
        onDrop={handleDrop}
        onClick={handleClick}
      >
        <input
          ref={inputRef}
          className="hidden"
          type="file"
          accept={accept}
          onChange={handleChange}
        />

        <div className="flex flex-col items-center justify-center pt-5 pb-6 px-4 text-center z-10">
          {fileName ? (
            <>
              <div className="w-16 h-16 rounded-full bg-green-500/20 flex items-center justify-center mb-4 animate-in zoom-in duration-300">
                <FileSpreadsheet className="w-8 h-8 text-green-500" />
              </div>
              <p className="mb-2 text-sm text-foreground font-medium truncate max-w-[200px]">
                {fileName}
              </p>
              <p className="text-xs text-green-500 font-semibold flex items-center gap-1">
                <CheckCircle2 className="w-3 h-3" /> Ready to process
              </p>
              <button
                onClick={clearFile}
                className="mt-4 p-2 rounded-full hover:bg-destructive/10 text-muted-foreground hover:text-destructive transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </>
          ) : (
            <>
              <div className={cn(
                "w-16 h-16 rounded-full bg-secondary flex items-center justify-center mb-4 transition-transform duration-300 group-hover:scale-110",
                dragActive && "scale-110 bg-primary/20 text-primary"
              )}>
                <Upload className="w-8 h-8 text-muted-foreground group-hover:text-primary transition-colors" />
              </div>
              <p className="mb-2 text-lg font-semibold text-foreground">
                {label}
              </p>
              <p className="text-sm text-muted-foreground max-w-xs mx-auto">
                {helper}
              </p>
            </>
          )}
        </div>

        {/* Background pattern effect */}
        <div className="absolute inset-0 opacity-[0.03] pointer-events-none bg-[radial-gradient(#fff_1px,transparent_1px)] [background-size:16px_16px]" />
      </div>
    </div>
  );
}
