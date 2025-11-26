import { useRef, useState } from 'react';
import { cn } from '@/lib/utils';
import { Upload, FileSpreadsheet, X, CheckCircle2 } from 'lucide-react';

interface FileUploadProps {
  onFileSelect: (file: File) => void;
  accept?: string;
  className?: string;
}

export function FileUpload({ onFileSelect, accept = ".xlsx, .xls", className }: FileUploadProps) {
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

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    e.preventDefault();
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const handleFile = (file: File) => {
    setFileName(file.name);
    onFileSelect(file);
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
        onClick={() => inputRef.current?.click()}
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
                Upload Excel Sheet
              </p>
              <p className="text-sm text-muted-foreground max-w-xs mx-auto">
                Drag and drop your .xlsx file here, or click to browse
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
