
import React, { useState, useCallback, useMemo } from 'react';
import type { ExcelRow } from './types';

// Declare XLSX from CDN to satisfy TypeScript
declare const XLSX: any;

// --- Helper Functions (Excel Service Logic) ---

const SOURCE_COLUMN_NAME = 'TÊN NGUỒN';

/**
 * Processes an array of Files, reading each sheet and returning its data.
 * @param files The array of Excel files to process.
 * @returns A promise that resolves to an array of merged data rows.
 */
const processMultipleFiles = async (files: File[]): Promise<ExcelRow[]> => {
  const allRows: ExcelRow[] = [];

  for (const file of files) {
    try {
      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array' });

      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (jsonData.length > 0) {
          jsonData.forEach(row => {
            allRows.push({
              [SOURCE_COLUMN_NAME]: sheetName,
              ...row
            });
          });
        }
      }
    } catch (error) {
      console.error(`Error processing file ${file.name}:`, error);
      throw new Error(`Không thể xử lý file: ${file.name}`);
    }
  }

  return allRows;
};

/**
 * Processes selected sheets from a single Excel file.
 * @param file The single Excel file.
 * @param selectedSheetNames The names of the sheets to merge.
 * @returns A promise that resolves to an array of merged data rows.
 */
const processSheetsInFile = async (file: File, selectedSheetNames: string[]): Promise<ExcelRow[]> => {
    const allRows: ExcelRow[] = [];
    try {
        const arrayBuffer = await file.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });

        for (const sheetName of selectedSheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;

            const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            if (jsonData.length > 0) {
                jsonData.forEach(row => {
                    allRows.push({
                        [SOURCE_COLUMN_NAME]: sheetName,
                        ...row
                    });
                });
            }
        }
    } catch (error) {
        console.error(`Error processing file ${file.name}:`, error);
        throw new Error(`Không thể xử lý file: ${file.name}`);
    }
    return allRows;
};


/**
 * Extracts all unique headers from the merged data.
 * @param data The array of merged data rows.
 * @returns An array of string headers.
 */
const getHeadersFromData = (data: ExcelRow[]): string[] => {
  if (data.length === 0) return [];
  const headerSet = new Set<string>();
  data.forEach(row => {
    Object.keys(row).forEach(key => headerSet.add(key));
  });

  const headers = Array.from(headerSet);
  if (headers.includes(SOURCE_COLUMN_NAME)) {
    return [SOURCE_COLUMN_NAME, ...headers.filter(h => h !== SOURCE_COLUMN_NAME)];
  }
  return headers;
};

/**
 * Exports data to an Excel file and triggers a download.
 * @param data The data to export.
 * @param fileName The desired name of the output file.
 */
const exportToExcel = (data: ExcelRow[], fileName: string): void => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'MergedData');
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};


// --- UI Components ---

const UploadIcon: React.FC<{className?: string}> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
    </svg>
);

const DownloadIcon: React.FC<{className?: string}> = ({ className }) => (
    <svg xmlns="http://www.w3.org/2000/svg" className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
        <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
    </svg>
);

const Spinner: React.FC = () => (
  <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
);

interface FileUploaderProps {
    onFilesSelected: (files: File[]) => void;
    disabled: boolean;
    multiple: boolean;
    uploadText: string;
}

const FileUploader: React.FC<FileUploaderProps> = ({ onFilesSelected, disabled, multiple, uploadText }) => {
    const [isDragging, setIsDragging] = useState(false);

    const handleDragEnter = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        if (!disabled) setIsDragging(true);
    };

    const handleDragLeave = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
    };

    const handleDragOver = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
    };

    const handleDrop = (e: React.DragEvent<HTMLLabelElement>) => {
        e.preventDefault();
        e.stopPropagation();
        setIsDragging(false);
        if (!disabled && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            onFilesSelected(Array.from(e.dataTransfer.files));
        }
    };
    
    const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        if (e.target.files && e.target.files.length > 0) {
            onFilesSelected(Array.from(e.target.files));
        }
        // Reset input value to allow re-uploading the same file
        e.target.value = '';
    };

    const borderStyle = isDragging ? 'border-blue-400' : 'border-slate-600';

    return (
        <label
            onDragEnter={handleDragEnter}
            onDragLeave={handleDragLeave}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            className={`flex flex-col items-center justify-center w-full h-64 border-2 ${borderStyle} border-dashed rounded-lg cursor-pointer bg-slate-800 hover:bg-slate-700 transition-colors ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
        >
            <div className="flex flex-col items-center justify-center pt-5 pb-6 text-center">
                <UploadIcon className="w-10 h-10 mb-3 text-slate-400"/>
                <p className="mb-2 text-sm text-slate-400">
                    <span className="font-semibold text-blue-400">Nhấn để chọn</span> hoặc kéo thả file
                </p>
                <p className="text-xs text-slate-500">{uploadText}</p>
            </div>
            <input id="dropzone-file" type="file" className="hidden" multiple={multiple} accept=".xlsx, .xls" onChange={handleFileChange} disabled={disabled} />
        </label>
    );
};


interface DataTableProps {
    data: ExcelRow[];
    headers: string[];
}
const DataTable: React.FC<DataTableProps> = ({ data, headers }) => {
    return (
        <div className="w-full max-h-[50vh] overflow-auto rounded-lg border border-slate-700 shadow-md">
            <table className="w-full text-sm text-left text-slate-300">
                <thead className="text-xs text-slate-300 uppercase bg-slate-800 sticky top-0">
                    <tr>
                        {headers.map((header) => (
                            <th key={header} scope="col" className="px-6 py-3 whitespace-nowrap">
                                {header}
                            </th>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {data.slice(0, 100).map((row, rowIndex) => (
                        <tr key={rowIndex} className="bg-slate-900 border-b border-slate-700 hover:bg-slate-800/50">
                            {headers.map(header => (
                                <td key={`${rowIndex}-${header}`} className="px-6 py-4 whitespace-nowrap">
                                    {row[header] === null ? '' : String(row[header])}
                                </td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

// --- Main App Component ---

type MergeMode = 'files' | 'sheets';

export default function App() {
    const [mergeMode, setMergeMode] = useState<MergeMode>('files');
    const [files, setFiles] = useState<File[]>([]);
    const [singleFileSheets, setSingleFileSheets] = useState<{ name: string; selected: boolean }[]>([]);
    const [mergedData, setMergedData] = useState<ExcelRow[] | null>(null);
    const [isLoading, setIsLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const headers = useMemo(() => (mergedData ? getHeadersFromData(mergedData) : []), [mergedData]);

    const resetState = () => {
        setFiles([]);
        setSingleFileSheets([]);
        setMergedData(null);
        setError(null);
    };

    const handleModeChange = (mode: MergeMode) => {
        setMergeMode(mode);
        resetState();
    };

    const extractSheetNames = async (file: File) => {
        setIsLoading(true);
        setError(null);
        try {
            const arrayBuffer = await file.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array', bookSheets: true });
            setSingleFileSheets(workbook.SheetNames.map((name: string) => ({ name, selected: true })));
        } catch (e) {
            setError('Không thể đọc được file. Vui lòng kiểm tra lại file có bị lỗi không.');
            setFiles([]);
        } finally {
            setIsLoading(false);
        }
    };

    const handleFilesSelected = (selectedFiles: File[]) => {
        setMergedData(null);
        setError(null);
        if (mergeMode === 'sheets') {
            const singleFile = selectedFiles[0];
            if (singleFile) {
                setFiles([singleFile]);
                extractSheetNames(singleFile);
            }
        } else {
            setFiles(selectedFiles);
        }
    };

    const handleSheetSelectionChange = (sheetName: string, isSelected: boolean) => {
        setSingleFileSheets(prevSheets =>
            prevSheets.map(sheet =>
                sheet.name === sheetName ? { ...sheet, selected: isSelected } : sheet
            )
        );
    };
    
    const toggleAllSheets = (select: boolean) => {
        setSingleFileSheets(prev => prev.map(s => ({ ...s, selected: select })));
    };


    const handleMerge = useCallback(async () => {
        if (files.length === 0) {
            setError('Vui lòng chọn file.');
            return;
        }
        setIsLoading(true);
        setError(null);
        setMergedData(null);

        try {
            let data: ExcelRow[] = [];
            if (mergeMode === 'files') {
                data = await processMultipleFiles(files);
            } else { // 'sheets' mode
                const selectedSheets = singleFileSheets.filter(s => s.selected).map(s => s.name);
                if (selectedSheets.length === 0) {
                    setError('Vui lòng chọn ít nhất một sheet để gộp.');
                    setIsLoading(false);
                    return;
                }
                data = await processSheetsInFile(files[0], selectedSheets);
            }

            if (data.length === 0) {
                setError('Không tìm thấy dữ liệu trong các file/sheet đã chọn.');
            } else {
                setMergedData(data);
            }
        } catch (e) {
            setError(e instanceof Error ? e.message : 'Đã xảy ra lỗi không xác định.');
            console.error(e);
        } finally {
            setIsLoading(false);
        }
    }, [files, mergeMode, singleFileSheets]);
    
    const handleDownload = () => {
        if(mergedData && mergedData.length > 0) {
            exportToExcel(mergedData, 'DuLieuGop');
        }
    };
    
    const isMergeDisabled = files.length === 0 || isLoading || (mergeMode === 'sheets' && singleFileSheets.length > 0 && !singleFileSheets.some(s => s.selected));

    return (
        <div className="min-h-screen bg-slate-900 text-slate-200 flex flex-col items-center p-4 sm:p-6 lg:p-8">
            <div className="w-full max-w-4xl mx-auto">
                <header className="text-center mb-8">
                    <h1 className="text-3xl sm:text-4xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-teal-300">
                        Công Cụ Gộp File Excel
                    </h1>
                    <p className="mt-2 text-slate-400">
                        Gộp nhiều sheet hoặc file thành một, tự động thêm cột và xử lý các cấu trúc cột khác nhau.
                    </p>
                </header>

                <main className="space-y-6">
                    <div className="p-6 bg-slate-800/50 rounded-lg border border-slate-700 shadow-lg">
                        <h2 className="text-xl font-semibold mb-4 text-slate-100">1. Chọn chế độ gộp</h2>
                        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <button onClick={() => handleModeChange('files')} className={`p-4 rounded-lg border-2 text-left transition-all ${mergeMode === 'files' ? 'bg-blue-600/30 border-blue-500' : 'bg-slate-800 border-slate-600 hover:border-slate-500'}`}>
                                <h3 className="font-bold">Gộp nhiều file Excel</h3>
                                <p className="text-sm text-slate-400 mt-1">Chọn và gộp nhiều file excel riêng biệt thành một file duy nhất.</p>
                            </button>
                            <button onClick={() => handleModeChange('sheets')} className={`p-4 rounded-lg border-2 text-left transition-all ${mergeMode === 'sheets' ? 'bg-blue-600/30 border-blue-500' : 'bg-slate-800 border-slate-600 hover:border-slate-500'}`}>
                                <h3 className="font-bold">Gộp nhiều sheet trong 1 file</h3>
                                <p className="text-sm text-slate-400 mt-1">Chọn một file excel và gộp các sheet được chỉ định từ file đó.</p>
                            </button>
                        </div>
                    </div>
                    
                    <div className="p-6 bg-slate-800/50 rounded-lg border border-slate-700 shadow-lg">
                        <h2 className="text-xl font-semibold mb-4 text-slate-100">2. Tải lên File</h2>
                        <FileUploader 
                            onFilesSelected={handleFilesSelected} 
                            disabled={isLoading} 
                            multiple={mergeMode === 'files'}
                            uploadText={mergeMode === 'files' ? 'File Excel (.xlsx, .xls) - Có thể chọn nhiều file' : 'File Excel (.xlsx, .xls) - Chỉ chọn một file'}
                        />
                        {files.length > 0 && (
                             <div className="mt-4 text-sm text-slate-400">
                                <p className="font-semibold">File đã chọn:</p>
                                <ul className="list-disc list-inside">
                                    {files.map(f => <li key={f.name}>{f.name} ({Math.round(f.size/1024)} KB)</li>)}
                                </ul>
                            </div>
                        )}
                    </div>
                    
                    {mergeMode === 'sheets' && singleFileSheets.length > 0 && (
                        <div className="p-6 bg-slate-800/50 rounded-lg border border-slate-700 shadow-lg">
                            <div className="flex justify-between items-center mb-4">
                               <h2 className="text-xl font-semibold text-slate-100">3. Chọn các sheet cần gộp</h2>
                               <div className="flex gap-2">
                                    <button onClick={() => toggleAllSheets(true)} className="text-xs px-2 py-1 bg-slate-700 hover:bg-slate-600 rounded">Chọn tất cả</button>
                                    <button onClick={() => toggleAllSheets(false)} className="text-xs px-2 py-1 bg-slate-700 hover:bg-slate-600 rounded">Bỏ chọn tất cả</button>
                               </div>
                            </div>
                            <div className="max-h-60 overflow-y-auto space-y-2 pr-2">
                                {singleFileSheets.map(sheet => (
                                    <label key={sheet.name} className="flex items-center p-2 bg-slate-800 rounded-md cursor-pointer hover:bg-slate-700">
                                        <input
                                            type="checkbox"
                                            checked={sheet.selected}
                                            onChange={(e) => handleSheetSelectionChange(sheet.name, e.target.checked)}
                                            className="w-4 h-4 text-blue-600 bg-gray-700 border-gray-600 rounded focus:ring-blue-600 ring-offset-gray-800 focus:ring-2"
                                        />
                                        <span className="ml-3 text-sm font-medium text-slate-300">{sheet.name}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                    )}


                    <div className="flex justify-center">
                        <button
                            onClick={handleMerge}
                            disabled={isMergeDisabled}
                            className="flex items-center justify-center gap-2 px-8 py-3 font-semibold text-white bg-blue-600 rounded-lg shadow-md hover:bg-blue-700 disabled:bg-slate-600 disabled:cursor-not-allowed transition-all duration-300 transform hover:scale-105 disabled:scale-100 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 focus:ring-offset-slate-900"
                        >
                            {isLoading && files.length > 0 ? <Spinner /> : `Gộp File`}
                        </button>
                    </div>

                    {error && (
                        <div className="p-4 text-center text-red-300 bg-red-900/50 border border-red-700 rounded-lg">
                            {error}
                        </div>
                    )}
                    
                    {mergedData && mergedData.length > 0 && (
                        <div className="p-6 bg-slate-800/50 rounded-lg border border-slate-700 shadow-lg space-y-4">
                            <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-4">
                                <div>
                                    <h2 className="text-xl font-semibold text-slate-100">Kết quả & Tải về</h2>
                                    <p className="text-sm text-slate-400">Hiển thị 100 dòng đầu tiên. Tổng số dòng đã gộp: {mergedData.length}.</p>
                                </div>
                                <button
                                    onClick={handleDownload}
                                    className="flex items-center justify-center gap-2 px-6 py-2 font-semibold text-white bg-teal-600 rounded-lg shadow-md hover:bg-teal-700 disabled:bg-slate-600 transition-all duration-300 transform hover:scale-105 focus:outline-none focus:ring-2 focus:ring-teal-500 focus:ring-offset-2 focus:ring-offset-slate-900"
                                >
                                    <DownloadIcon className="w-5 h-5"/>
                                    Tải File đã gộp
                                </button>
                            </div>
                            <DataTable data={mergedData} headers={headers} />
                        </div>
                    )}
                </main>

                <footer className="text-center mt-12 text-sm text-slate-500">
                    <p>&copy; {new Date().getFullYear()} Excel Merger Tool. All rights reserved.</p>
                </footer>
            </div>
        </div>
    );
}
