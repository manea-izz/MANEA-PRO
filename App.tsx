
import React, { useState, useCallback, useEffect, useRef } from 'react';
import { extractDataFromFile, getCompanyInfo, isValidSwift } from './services/geminiService';
import { ProcessableFile, EnrichedData, ExtractedData } from './types';
import { UploadIcon, CheckIcon, CrossIcon, InfoIcon, PdfIcon, ImageIcon, FileIcon, TrashIcon, CopyIcon, ClearIcon, WordIcon, ExcelIcon, TextIcon, WhatsAppIcon, FacebookIcon, InstagramIcon, ChevronDownIcon, ChevronUpIcon } from './components/icons';
import Spinner from './components/Spinner';
import { Part } from '@google/genai';

// --- Type declarations for external libraries ---
declare const mammoth: any;
declare const XLSX: any;

// --- File Parsing Helper Functions ---

const toBase64 = (file: File, onProgress?: (percent: number) => void): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    if (onProgress) {
        reader.onprogress = (event) => {
            if (event.lengthComputable) {
                const percent = Math.round((event.loaded / event.total) * 100);
                onProgress(percent);
            }
        };
    }
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = (error) => reject(error);
  });

const toText = (file: File): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsText(file);
    reader.onload = () => resolve(reader.result as string);
    reader.onerror = (error) => reject(error);
  });
  
const wordToText = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value;
};

const excelToText = async (file: File): Promise<string> => {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'buffer' });
    let fullText = '';
    workbook.SheetNames.forEach((sheetName: string) => {
      const worksheet = workbook.Sheets[sheetName];
      const sheetText = XLSX.utils.sheet_to_csv(worksheet);
      fullText += `--- ${sheetName} ---\n${sheetText}\n\n`;
    });
    return fullText;
};


const prepareContentPart = async (file: File, onProgress?: (percent: number) => void): Promise<Part> => {
  const { type, name } = file;
  if (type.startsWith('image/') || type === 'application/pdf') {
    const base64 = await toBase64(file, onProgress);
    return { inlineData: { mimeType: type, data: base64 } };
  } else if (type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || name.endsWith('.docx')) {
    const text = await wordToText(file);
    if (onProgress) onProgress(100);
    return { text };
  } else if (type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || name.endsWith('.xlsx') || name.endsWith('.xls')) {
    const text = await excelToText(file);
    if (onProgress) onProgress(100);
    return { text };
  } else if (type.startsWith('text/') || name.endsWith('.txt')) {
    const text = await toText(file);
    if (onProgress) onProgress(100);
    return { text };
  }
  throw new Error(`نوع الملف غير مدعوم: ${type || name}`);
};

// --- Child Components ---

const ProgressBar: React.FC<{ progress: number; label: string; estimatedTime?: string | null }> = ({ progress, label, estimatedTime }) => (
    <div className="w-full max-w-md mx-auto mt-6 animate-slide-in-fade-in">
        <div className="flex justify-between mb-2">
            <span className="text-sm font-medium text-brand-blue-light">{label}</span>
            <span className="text-sm font-medium text-brand-blue-light">{Math.min(100, Math.max(0, progress))}%</span>
        </div>
        <div className="w-full bg-brand-gray-700 rounded-full h-2.5 overflow-hidden">
            <div 
                className="bg-gradient-to-r from-brand-blue to-brand-blue-light h-2.5 rounded-full transition-all duration-300 ease-out shadow-[0_0_10px_rgba(0,180,216,0.5)]" 
                style={{ width: `${Math.min(100, Math.max(0, progress))}%` }}
            ></div>
        </div>
        {estimatedTime && (
            <p className="text-xs text-brand-gray-400 mt-2 text-left dir-rtl">
                ⏱️ الوقت المتبقي التقريبي: {estimatedTime}
            </p>
        )}
    </div>
);

const SwiftTooltip: React.FC<{ swift: string; bankName: string; country: string; children: React.ReactNode }> = ({ swift, bankName, country, children }) => {
    const [isVisible, setIsVisible] = useState(false);
    const valid = isValidSwift(swift);

    if (!swift) return <>{children}</>;

    return (
        <div className="relative inline-block group" onMouseEnter={() => setIsVisible(true)} onMouseLeave={() => setIsVisible(false)}>
            {children}
            {isVisible && (
                <div className="absolute z-[100] top-full left-1/2 -translate-x-1/2 mt-3 w-80 p-0 overflow-visible bg-brand-gray-900/95 backdrop-blur-3xl border border-white/10 rounded-[2rem] shadow-[0_30px_60px_-15px_rgba(0,0,0,0.8)] animate-slide-in-fade-in dir-rtl text-right ring-1 ring-white/10">
                    {/* Floating Accent Bar */}
                    <div className={`h-1.5 w-full relative overflow-hidden rounded-t-[2rem] ${valid ? 'bg-green-500' : 'bg-red-500'}`}>
                        <div className="absolute inset-0 bg-gradient-to-r from-transparent via-white/40 to-transparent animate-[shimmer_2s_infinite]"></div>
                    </div>
                    
                    <div className="p-7">
                        <div className="flex items-center justify-between mb-5 pb-4 border-b border-white/5">
                            <div className="flex items-center gap-3">
                                <div className={`p-2.5 rounded-2xl shadow-xl transition-transform duration-500 group-hover:scale-110 ${valid ? 'bg-green-500/10 text-green-400 ring-1 ring-green-500/20' : 'bg-red-500/10 text-red-400 ring-1 ring-red-500/20'}`}>
                                    {valid ? <CheckIcon className="w-5 h-5" /> : <CrossIcon className="w-5 h-5" />}
                                </div>
                                <div>
                                    <h4 className="text-[11px] font-black text-white uppercase tracking-[0.15em]">نظام التحقق الذكي</h4>
                                    <p className={`text-[9px] font-bold ${valid ? 'text-green-400/80' : 'text-red-400/80'}`}>
                                        {valid ? 'معتمد ومطابق' : 'غير مطابق للمواصفات'}
                                    </p>
                                </div>
                            </div>
                            <span className="text-[10px] font-mono font-black text-brand-blue-light bg-brand-blue/10 px-2.5 py-1 rounded-xl ring-1 ring-brand-blue/20">{swift}</span>
                        </div>

                        <div className="space-y-6">
                            <div className="flex items-start gap-4 group/item">
                                <div className="mt-0.5 p-2.5 bg-brand-gray-800 rounded-2xl ring-1 ring-white/5 group-hover/item:ring-brand-blue/40 transition-all duration-300 shadow-lg">
                                    <FileIcon className="w-4 h-4 text-brand-blue-light" />
                                </div>
                                <div className="flex-1">
                                    <p className="text-[9px] text-brand-gray-500 uppercase font-black tracking-[0.2em] mb-1.5">اسم المصرف</p>
                                    <p className="text-sm text-brand-gray-100 font-bold leading-relaxed">{bankName || 'غير مدرج في المستند'}</p>
                                </div>
                            </div>
                            
                            <div className="flex items-start gap-4 group/item">
                                <div className="mt-0.5 p-2.5 bg-brand-gray-800 rounded-2xl ring-1 ring-white/5 group-hover/item:ring-brand-blue/40 transition-all duration-300 shadow-lg">
                                    <InfoIcon className="w-4 h-4 text-brand-blue-light" />
                                </div>
                                <div className="flex-1">
                                    <p className="text-[9px] text-brand-gray-500 uppercase font-black tracking-[0.2em] mb-1.5">الدولة المستضيفة</p>
                                    <p className="text-sm text-brand-gray-100 font-bold leading-relaxed">{country || 'غير مدرج في المستند'}</p>
                                </div>
                            </div>

                            <div className="pt-5 mt-2 border-t border-white/5 flex items-center justify-between">
                                <div className="flex flex-col">
                                    <p className="text-[10px] text-brand-gray-500 font-black uppercase tracking-widest">البنية الهيكلية</p>
                                    <div className="flex items-center gap-1.5 mt-1">
                                        <div className={`w-2 h-2 rounded-full ${valid ? 'bg-green-500 animate-pulse' : 'bg-red-500'}`}></div>
                                        <p className={`text-[11px] font-black ${valid ? 'text-green-400' : 'text-red-400'}`}>
                                            {valid ? 'BIC-11 صالح' : 'تنسيق معيب'}
                                        </p>
                                    </div>
                                </div>
                                <div className={`flex items-center gap-2 px-4 py-2 rounded-2xl text-[10px] font-black uppercase tracking-tighter shadow-2xl transition-colors ${valid ? 'bg-green-500/10 text-green-400 border border-green-500/20' : 'bg-red-500/10 text-red-400 border border-red-500/20'}`}>
                                    {valid ? 'موثوق' : 'مرفوض'}
                                </div>
                            </div>
                        </div>
                    </div>
                    {/* Stylized Caret at Top */}
                    <div className="absolute -top-1.5 left-1/2 -translate-x-1/2 w-4 h-4 bg-brand-gray-900 border-l border-t border-white/10 rotate-45 shadow-[-2px_-2px_15px_rgba(0,0,0,0.3)]"></div>
                </div>
            )}
        </div>
    );
};

// New SwiftBadge component for shared styled display
const SwiftBadge: React.FC<{ swift: string; bankName: string; country: string }> = ({ swift, bankName, country }) => {
    const valid = isValidSwift(swift);
    return (
        <SwiftTooltip swift={swift} bankName={bankName} country={country}>
            <div className={`inline-flex items-center justify-center gap-2 px-4 py-1.5 rounded-full border shadow-md transition-all duration-300 cursor-help ring-1 ring-inset ${
                valid 
                ? 'bg-green-500/10 border-green-500/30 text-green-400 ring-green-500/20 hover:bg-green-500/20' 
                : 'bg-red-500/10 border-red-500/30 text-red-400 ring-red-500/20 hover:bg-red-500/20'
            }`}>
                <span className="font-mono font-black text-xs tracking-widest uppercase">{swift}</span>
                {valid ? <CheckIcon className="w-3.5 h-3.5" /> : <CrossIcon className="w-3.5 h-3.5" />}
            </div>
        </SwiftTooltip>
    );
};

const DropZone: React.FC<{ onFilesSelect: (files: File[]) => void; multiple: boolean; disabled: boolean; label: string }> = ({ onFilesSelect, multiple, disabled, label }) => {
  const [isDragging, setIsDragging] = useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      onFilesSelect(Array.from(e.target.files));
      e.target.value = '';
    }
  };

  const handleDragEvents = (e: React.DragEvent<HTMLDivElement>, isEntering: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    if (!disabled) setIsDragging(isEntering);
  };
  
  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    handleDragEvents(e, false);
    if (!disabled && e.dataTransfer.files?.length) {
      onFilesSelect(Array.from(e.dataTransfer.files));
      e.dataTransfer.clearData();
    }
  };
  
  return (
    <div
      onClick={() => !disabled && fileInputRef.current?.click()}
      onDragEnter={(e) => handleDragEvents(e, true)}
      onDragLeave={(e) => handleDragEvents(e, false)}
      onDragOver={(e) => e.preventDefault()}
      onDrop={handleDrop}
      className={`relative w-full h-48 flex flex-col items-center justify-center p-8 text-center bg-brand-gray-800/50 border-2 border-dashed rounded-xl cursor-pointer transition-all duration-300 ${isDragging ? 'border-brand-blue-light bg-brand-gray-700/50' : 'border-brand-gray-700 hover:border-brand-blue-light'} ${disabled ? 'opacity-50 cursor-not-allowed' : ''}`}
    >
      <div className="flex flex-col items-center justify-center text-brand-gray-400 pointer-events-none">
        <UploadIcon className="w-12 h-12 text-brand-gray-500 mb-3"/>
        <p className="mt-2 text-lg font-semibold">{label}</p>
        <p className="text-sm">أو اسحب وأفلت الملفات هنا</p>
        <p className="text-xs mt-1 text-brand-gray-500">يمكنك لصق الصور أو المستندات مباشرة</p>
      </div>
      <input 
        ref={fileInputRef}
        type='file' 
        className="hidden" 
        multiple={multiple} 
        onChange={handleFileChange} 
        disabled={disabled} 
        accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.txt"
      />
    </div>
  );
};

const ResultCard: React.FC<{ title: string; data: EnrichedData | null; showCompanyInfo?: boolean }> = ({ title, data, showCompanyInfo = true }) => {
    const [copiedSection, setCopiedSection] = useState<'data' | 'info' | null>(null);
    const [isDataExpanded, setIsDataExpanded] = useState(true);
    const [isInfoExpanded, setIsInfoExpanded] = useState(false);

    if (!data) return null;
    
    const dataFields: { key: keyof ExtractedData; label: string }[] = [
      { key: 'beneficiaryName', label: 'اسم المستفيد' },
      { key: 'accountNumber', label: 'رقم الحساب' },
      { key: 'swiftCode', label: 'سويفت البنك' },
      { key: 'bankName', label: 'البنك' },
      { key: 'country', label: 'الدولة' },
      { key: 'city', label: 'المدينة' },
      { key: 'state', label: 'المقاطعة / الولاية' },
      { key: 'address', label: 'العنوان' },
    ];
    
    const handleCopy = (section: 'data' | 'info') => {
        if (!data) return;
        let textToCopy = '';

        if (section === 'data') {
            const header = "📋 --- البيانات المستخرجة ---\n\n";
            const footer = "\n\n------------------------------";
            const fieldsText = dataFields
                .map(({ key, label }) => {
                    const value = data[key as keyof ExtractedData];
                    return value ? `▪️ ${label}:\n   ${value}` : null;
                })
                .filter(Boolean)
                .join('\n\n');
            if (fieldsText) {
                textToCopy = header + fieldsText + footer;
            }
        } else { // section === 'info'
            if (data.companyInfo) {
                textToCopy = data.companyInfo.trim();
            }
        }

        if (textToCopy) {
            navigator.clipboard.writeText(textToCopy);
            setCopiedSection(section);
            setTimeout(() => setCopiedSection(null), 2000);
        }
    };
  
    return (
      <div className="bg-brand-gray-800 p-6 rounded-xl shadow-lg w-full h-fit flex flex-col transition-all duration-300 border border-brand-gray-700">
        <h3 className="text-xl font-bold text-brand-blue-light mb-4">{title}</h3>
        
        {/* Extracted Data Section */}
        <div className="mb-4 border border-brand-gray-700 rounded-lg overflow-hidden shadow-sm">
            <button 
                onClick={() => setIsDataExpanded(!isDataExpanded)}
                className="w-full flex justify-between items-center p-3 bg-brand-gray-700/50 hover:bg-brand-gray-700 transition-colors"
            >
                <div className="flex items-center gap-2">
                    <h4 className="text-md font-semibold text-brand-gray-200 uppercase">البيانات المستخرجة (UPPERCASE)</h4>
                </div>
                 <div className="flex items-center gap-2">
                     <div onClick={(e) => { e.stopPropagation(); handleCopy('data'); }} className={`cursor-pointer p-1 rounded hover:bg-brand-gray-600 ${copiedSection === 'data' ? 'text-green-400' : 'text-brand-gray-400'}`} title="نسخ البيانات">
                         {copiedSection === 'data' ? <CheckIcon /> : <CopyIcon className="h-4 w-4" />}
                     </div>
                     {isDataExpanded ? <ChevronUpIcon className="text-brand-gray-400" /> : <ChevronDownIcon className="text-brand-gray-400" />}
                 </div>
            </button>
            
            {isDataExpanded && (
                <div className="p-4 bg-brand-gray-800/30 space-y-3 animate-slide-in-fade-in">
                     {dataFields.map(({key, label}) => data[key as keyof ExtractedData] ? (
                        <div key={key} className="border-b border-brand-gray-700/50 last:border-0 pb-2 last:pb-0">
                          <p className="text-[10px] font-bold text-brand-gray-400 uppercase tracking-widest mb-2 text-center">{label}</p>
                          <div className="flex items-center justify-center">
                            {key === 'swiftCode' ? (
                                <SwiftBadge swift={data.swiftCode} bankName={data.bankName} country={data.country} />
                            ) : (
                                <p className="text-brand-gray-100 text-center font-bold text-sm break-all uppercase leading-relaxed">{data[key as keyof ExtractedData]}</p>
                            )}
                          </div>
                       </div>
                     ) : (
                        key !== 'swiftCode' ? (
                          <div key={key} className="border-b border-brand-gray-700/50 last:border-0 pb-2 last:pb-0">
                             <p className="text-[10px] font-bold text-brand-gray-400 uppercase tracking-widest mb-1 text-center">{label}</p>
                             <div className="flex items-center justify-center">
                                <span className="text-brand-gray-600 opacity-30 select-none text-xl font-light">−</span>
                             </div>
                          </div>
                        ) : null
                     ))}
                </div>
            )}
        </div>

        {/* Invoice Audit Section */}
        {showCompanyInfo && data.companyInfo && (
            <div className="border border-brand-blue-light/30 bg-brand-blue-light/5 rounded-lg overflow-hidden shadow-[0_4px_20px_-4px_rgba(0,180,216,0.15)] transition-all duration-300">
                <button 
                    onClick={() => setIsInfoExpanded(!isInfoExpanded)}
                    className="w-full flex justify-between items-center p-3 bg-brand-blue/10 hover:bg-brand-blue/20 transition-colors"
                >
                    <div className="flex items-center gap-2">
                         <InfoIcon className="text-brand-blue-light" />
                        <h4 className="text-md font-bold text-brand-blue-light">فحص الفاتورة</h4>
                    </div>
                     <div className="flex items-center gap-2">
                         <div onClick={(e) => { e.stopPropagation(); handleCopy('info'); }} className={`cursor-pointer p-1 rounded hover:bg-brand-blue/20 ${copiedSection === 'info' ? 'text-green-400' : 'text-brand-blue-light'}`} title="نسخ المعلومات">
                             {copiedSection === 'info' ? <CheckIcon /> : <CopyIcon className="h-4 w-4" />}
                         </div>
                         {isInfoExpanded ? <ChevronUpIcon className="text-brand-blue-light" /> : <ChevronDownIcon className="text-brand-blue-light" />}
                     </div>
                </button>
                
                {isInfoExpanded && (
                    <div className="p-4 bg-brand-gray-900/40 animate-slide-in-fade-in border-t border-brand-blue-light/20">
                         <div className="text-sm text-brand-gray-200 whitespace-pre-wrap leading-relaxed prose prose-invert max-w-none">
                            {data.companyInfo}
                         </div>
                         {data.sources?.length && (
                            <div className="mt-4 pt-3 border-t border-white/5">
                                <h5 className="text-[10px] font-bold text-brand-blue-light uppercase tracking-widest mb-2">المصادر الموثقة:</h5>
                                <div className="flex flex-wrap gap-2">
                                    {data.sources.map((s, i) => (
                                      <a key={i} href={s.uri} target="_blank" rel="noopener noreferrer" className="text-[10px] bg-brand-blue/20 hover:bg-brand-blue text-brand-blue-light hover:text-white px-2.5 py-1 rounded-md transition-all truncate max-w-[220px] border border-brand-blue/30">
                                        {s.title}
                                      </a>
                                    ))}
                                </div>
                            </div>
                        )}
                    </div>
                )}
            </div>
        )}
      </div>
    );
};

const ComparisonTable: React.FC<{ files: ProcessableFile[] }> = ({ files }) => {
    const results = files.filter(f => f.status === 'done' && f.data).map(f => ({ data: f.data!, fileName: f.file.name }));
    const [copiedColumn, setCopiedColumn] = useState<number | null>(null);

    if (!results || results.length === 0) return null;

    const fields: { key: keyof ExtractedData; label: string; isMono?: boolean }[] = [
      { key: 'beneficiaryName', label: 'اسم المستفيد' },
      { key: 'accountNumber', label: 'رقم الحساب' },
      { key: 'swiftCode', label: 'سويفت البنك' },
      { key: 'bankName', label: 'البنك' },
      { key: 'country', label: 'الدولة' },
      { key: 'city', label: 'المدينة' },
      { key: 'state', label: 'المقاطعة / الولاية' },
      { key: 'address', label: 'العنوان' },
    ];

    const handleCopyFile = (index: number, data: EnrichedData, fileName: string) => {
        const lines = fields.map(field => {
            const value = data[field.key];
            return value ? `${field.label}: ${String(value).toUpperCase()}` : null;
        }).filter(Boolean);

        const formattedText = `بيانات الملف: ${fileName.toUpperCase()}\n━━━━━━━━━━━━━━━━━━\n` + lines.join('\n');
        
        navigator.clipboard.writeText(formattedText);
        setCopiedColumn(index);
        setTimeout(() => setCopiedColumn(null), 2000);
    };

    return (
        <div className="w-full bg-brand-gray-800 rounded-2xl shadow-2xl border border-brand-gray-700 mt-8 animate-slide-in-fade-in ring-1 ring-white/5 overflow-visible">
            <div className="p-6 border-b border-brand-gray-700 bg-brand-gray-900/50 backdrop-blur-sm flex justify-between items-center rounded-t-2xl">
                 <div className="flex items-center gap-3">
                    <div className="p-2 bg-brand-blue/10 rounded-lg text-brand-blue-light ring-1 ring-brand-blue/20">
                         <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                        </svg>
                    </div>
                    <div>
                         <h3 className="text-xl font-bold text-white uppercase">البيانات المستخرجة (UPPERCASE)</h3>
                         <p className="text-sm text-brand-gray-400">عرض شامل للبيانات المستخرجة من {results.length} ملفات</p>
                    </div>
                </div>
            </div>

            <div className="overflow-x-auto">
                <table className="w-full text-sm text-center text-brand-gray-100 border-collapse table-fixed">
                    <thead className="text-xs text-brand-gray-400 uppercase bg-brand-gray-900">
                        <tr>
                            <th scope="col" className="px-6 py-5 font-bold w-48 sticky right-0 z-20 bg-brand-gray-900 border-b border-brand-gray-700 text-brand-blue-light tracking-wider shadow-lg text-right">
                                الحقل المطلوب
                            </th>
                            {results.map((res, index) => (
                                <th key={index} scope="col" className="px-4 py-5 font-semibold border-b border-brand-gray-700">
                                    <div className="flex items-center justify-between bg-brand-gray-800/80 border border-brand-gray-700 rounded-lg p-2 group hover:border-brand-blue-light/30 transition-colors">
                                        <div className="flex items-center gap-2 mx-auto min-w-0">
                                             <span className="text-brand-blue-light font-bold opacity-50 flex-shrink-0">#{index + 1}</span>
                                             <span className="break-words text-brand-gray-200 block uppercase" title={res.fileName}>{res.fileName}</span>
                                        </div>
                                        <button 
                                            onClick={() => handleCopyFile(index, res.data, res.fileName)}
                                            className={`p-1.5 rounded-md transition-all flex-shrink-0 ${copiedColumn === index ? 'bg-green-500/20 text-green-400' : 'bg-brand-gray-700 text-brand-gray-400 hover:bg-brand-blue hover:text-white'}`}
                                            title="نسخ بيانات الملف"
                                        >
                                            {copiedColumn === index ? <CheckIcon /> : <CopyIcon className="w-4 h-4" />}
                                        </button>
                                    </div>
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-brand-gray-700/30">
                        {fields.map((field) => {
                            return (
                                <tr key={field.key} className="group transition-colors hover:bg-brand-gray-700/30 even:bg-brand-gray-800/30 odd:bg-brand-gray-800/10">
                                    <th scope="row" className="px-6 py-5 font-bold text-brand-gray-200 sticky right-0 z-10 border-l border-brand-gray-700 bg-brand-gray-800 shadow-[4px_0_24px_-2px_rgba(0,0,0,0.5)] group-hover:bg-brand-gray-800/90 transition-colors text-right align-top">
                                        <div className="flex items-center gap-2 mt-1">
                                             <span className="w-1.5 h-1.5 rounded-full bg-brand-blue-light/50 group-hover:bg-brand-blue-light transition-colors flex-shrink-0"></span>
                                             <span className="tracking-wide">{field.label}</span>
                                        </div>
                                    </th>
                                    {results.map((res, index) => {
                                        const val = res.data[field.key];
                                        return (
                                            <td key={index} className={`px-6 py-5 text-sm align-top border-l border-brand-gray-700/20 last:border-0 font-bold leading-relaxed text-brand-gray-200 group-hover:text-white transition-colors text-center uppercase`}>
                                               <div className="break-words whitespace-pre-wrap w-full uppercase flex flex-col items-center justify-center gap-2">
                                                  {val ? (
                                                      field.key === 'swiftCode' ? (
                                                        <SwiftBadge swift={res.data.swiftCode} bankName={res.data.bankName} country={res.data.country} />
                                                      ) : val
                                                  ) : (
                                                      <span className="text-brand-gray-600 opacity-30 select-none text-xl font-light">−</span>
                                                  )}
                                               </div>
                                            </td>
                                        );
                                    })}
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            </div>
        </div>
    );
};

// --- Main App Component ---

function App() {
  const [singleFile, setSingleFile] = useState<File | null>(null);
  const [processableFiles, setProcessableFiles] = useState<ProcessableFile[]>([]);
  const [singleResult, setSingleResult] = useState<EnrichedData | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'single' | 'multi'>('single');
  
  // States for progress bar
  const [progress, setProgress] = useState(0);
  const [processingStatus, setProcessingStatus] = useState<'idle' | 'reading' | 'analyzing'>('idle');
  const [estimatedTime, setEstimatedTime] = useState<string | null>(null);

  useEffect(() => {
    const handlePaste = (event: ClipboardEvent) => {
      if (isLoading) return;
      const files = Array.from(event.clipboardData?.files || []);
      if (files.length > 0) {
        handleFilesSelected(files);
      }
    };
    window.addEventListener('paste', handlePaste);
    return () => window.removeEventListener('paste', handlePaste);
  }, [activeTab, isLoading]);

  const handleFilesSelected = (files: File[]) => {
      setError(null);
      if (activeTab === 'single') {
        setSingleFile(files[0]);
        setSingleResult(null);
      } else {
        const newProcessableFiles = files.map(file => ({ file, status: 'pending' as const, id: `${file.name}-${Date.now()}-${Math.random()}` }));
        setProcessableFiles(prev => [...prev, ...newProcessableFiles]);
      }
  };

  const handleClear = () => {
    setError(null);
    if (activeTab === 'single') {
      setSingleFile(null);
      setSingleResult(null);
      setProcessingStatus('idle');
      setProgress(0);
      setEstimatedTime(null);
    } else {
      setProcessableFiles([]);
    }
  };

  const handleRemoveMultiFile = (idToRemove: string) => {
    setProcessableFiles(files => files.filter(f => f.id !== idToRemove));
  };

  const handleProcessSingleFile = useCallback(async () => {
    if (!singleFile) return;
    setIsLoading(true);
    setError(null);
    setSingleResult(null);
    setProcessingStatus('reading');
    setProgress(0);
    setEstimatedTime(null);

    let intervalId: any;

    try {
      const contentPart = await prepareContentPart(singleFile, (percent) => {
          setProgress(Math.round(percent * 0.3));
      });

      setProcessingStatus('analyzing');
      
      const sizeMB = singleFile.size / (1024 * 1024);
      // Faster estimation for Flash model
      const estimatedSeconds = Math.ceil(3 + (sizeMB * 1.2));
      setEstimatedTime(`${estimatedSeconds} ثانية`);

      let currentSimulated = 30;
      const maxSimulated = 98;
      const stepTime = (estimatedSeconds * 1000) / (maxSimulated - currentSimulated);
      
      intervalId = setInterval(() => {
          currentSimulated += 1;
          if (currentSimulated <= maxSimulated) {
              setProgress(currentSimulated);
              const remaining = Math.ceil(estimatedSeconds * (1 - ((currentSimulated - 30) / (maxSimulated - 30))));
              if (remaining > 0) setEstimatedTime(`${remaining} ثانية`);
              else setEstimatedTime("لحظات أخيرة...");
          }
      }, stepTime / 2); // Faster simulation steps

      const extractedData = await extractDataFromFile(contentPart);
      const { info, sources } = await getCompanyInfo(extractedData.beneficiaryName, extractedData.bankName, extractedData.goodsDescription);
      
      setProgress(100);
      setEstimatedTime("تم!");
      
      await new Promise(r => setTimeout(r, 200)); // Shorter delay after finish

      setSingleResult({ ...extractedData, companyInfo: info, sources });
    } catch (err: any) {
      setError(err.message || 'حدث خطأ غير متوقع.');
    } finally {
      clearInterval(intervalId);
      setIsLoading(false);
      setProcessingStatus('idle');
    }
  }, [singleFile]);
  
  const handleProcessMultiFile = useCallback(async () => {
    const filesToProcess = processableFiles.filter(pf => pf.status === 'pending');
    
    if (filesToProcess.length === 0 && processableFiles.filter(pf => pf.status === 'done').length < 2) {
      setError("يرجى رفع ملفين على الأقل للاستخراج المتعدد.");
      return;
    }
    
    if (filesToProcess.length === 0) return;

    setIsLoading(true);
    setError(null);

    // Increased concurrency for faster multi-file processing
    const CONCURRENCY_LIMIT = 5;
    const queue = [...filesToProcess];
    const activePromises: Promise<void>[] = [];

    const processOneFile = async (pf: ProcessableFile) => {
        if (pf.status === 'done' && pf.data) return;
        
        setProcessableFiles(prev => prev.map(f => f.id === pf.id ? { ...f, status: 'processing' } : f));

        try {
            const part = await prepareContentPart(pf.file);
            const data = await extractDataFromFile(part);
            const { info, sources } = await getCompanyInfo(data.beneficiaryName, data.bankName, data.goodsDescription);
            const enrichedData: EnrichedData = { ...data, companyInfo: info, sources };
            
            setProcessableFiles(prev => prev.map(f => f.id === pf.id ? { ...f, status: 'done', data: enrichedData } : f));
        } catch (err: any) {
            setProcessableFiles(prev => prev.map(f => f.id === pf.id ? { ...f, status: 'error', error: err.message } : f));
        }
    };

    while (queue.length > 0 || activePromises.length > 0) {
        while (queue.length > 0 && activePromises.length < CONCURRENCY_LIMIT) {
            const file = queue.shift()!;
            const promise = processOneFile(file).then(() => {
                activePromises.splice(activePromises.indexOf(promise), 1);
            });
            activePromises.push(promise);
        }

        if (activePromises.length > 0) {
            await Promise.race(activePromises);
        }
    }

    setIsLoading(false);
  }, [processableFiles]);
  
  const getFileIcon = (file: File) => {
    const { type, name } = file;
    if (type.startsWith('image/')) return <ImageIcon />;
    if (type === 'application/pdf') return <PdfIcon />;
    if (type.includes('word') || name.endsWith('.docx')) return <WordIcon />;
    if (type.includes('excel') || type.includes('spreadsheet') || name.endsWith('.xlsx') || name.endsWith('.xls')) return <ExcelIcon />;
    if (type.startsWith('text/') || name.endsWith('.txt')) return <TextIcon />;
    return <FileIcon />;
  };

  const renderStatusIndicator = (status: ProcessableFile['status']) => {
    switch (status) {
        case 'processing': return <div className="animate-spin rounded-full h-4 w-4 border-2 border-transparent border-t-brand-blue-light border-r-brand-blue-light"></div>;
        case 'done': return <CheckIcon />;
        case 'error': return <CrossIcon />;
        default: return <div className="h-2 w-2 rounded-full bg-brand-gray-600"></div>;
    }
  };

  return (
    <div className="min-h-screen text-brand-gray-100 p-4 sm:p-8 flex flex-col overflow-x-hidden">
      <header className="text-center mb-10">
        <h1 className="text-4xl sm:text-5xl font-bold text-transparent bg-clip-text bg-gradient-to-r from-brand-blue-light to-brand-blue uppercase tracking-tight">مانع برو</h1>
        <p className="text-lg text-brand-gray-400 mt-2">أداة فحص واستخراج البيانات الذكية (Official)</p>
      </header>
      <main className="flex-grow w-full max-w-6xl mx-auto bg-brand-gray-800/20 p-4 sm:p-8 rounded-2xl border border-brand-gray-700/50 shadow-2xl shadow-black/20">
        <div className="flex justify-center mb-8 bg-brand-gray-800 p-1 rounded-full w-fit mx-auto shadow-md">
            {['single', 'multi'].map(tab => (
              <button key={tab} onClick={() => setActiveTab(tab as any)} className={`px-6 py-2 text-md font-medium transition-all rounded-full ${activeTab === tab ? 'bg-brand-blue text-white shadow-lg' : 'text-brand-gray-400 hover:text-white'}`}>
                {tab === 'single' ? 'فحص ملف واحد' : 'استخراج متعدد'}
              </button>
            ))}
        </div>
        {error && <p className="text-red-400 my-4 text-center bg-red-900/20 border border-red-900/50 p-3 rounded-lg max-w-2xl mx-auto text-sm">{error}</p>}
        {activeTab === 'single' ? (
             <div className="w-full max-w-2xl mx-auto">
                <div className="flex flex-col items-center gap-4">
                    <DropZone onFilesSelect={handleFilesSelected} multiple={false} disabled={isLoading || !!singleFile} label="اختر ملفًا"/>
                    {singleFile && <div className="flex items-center justify-between w-full max-w-md bg-brand-gray-800 px-4 py-3 rounded-lg border border-brand-gray-700 shadow-md animate-slide-in-fade-in">
                        <div className="flex items-center gap-3 overflow-hidden">{getFileIcon(singleFile)}<span className="text-sm text-brand-gray-300 truncate font-medium uppercase" title={singleFile.name}>{singleFile.name}</span></div>
                        <button onClick={() => { setSingleFile(null); handleClear(); }} className="text-gray-500 hover:text-red-400 p-1.5 rounded-full hover:bg-brand-gray-700 transition-colors"><TrashIcon className="h-4 w-4" /></button>
                    </div>}
                    
                    {isLoading && processingStatus !== 'idle' && (
                        <ProgressBar 
                            progress={progress} 
                            label={processingStatus === 'reading' ? 'جاري قراءة الملف...' : 'جاري تحليل البيانات بالذكاء الاصطناعي...'}
                            estimatedTime={estimatedTime}
                        />
                    )}

                    <div className="w-full flex items-stretch gap-2 mt-2">
                        <button onClick={handleProcessSingleFile} disabled={!singleFile || isLoading} className="flex-grow bg-brand-blue hover:bg-brand-blue-light text-white font-bold py-3 px-4 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-brand-blue/20">{isLoading ? <Spinner/> : 'فحص واستخراج البيانات'}</button>
                        {(singleResult || error || singleFile) && !isLoading && <button onClick={handleClear} className="flex-shrink-0 bg-brand-gray-700 hover:bg-brand-gray-600 text-white font-bold p-3 rounded-lg transition-colors shadow-lg"><ClearIcon/></button>}
                    </div>
                </div>
                {singleResult && <div className="mt-8 animate-slide-in-fade-in uppercase"><ResultCard title="البيانات المستخرجة" data={singleResult} /></div>}
            </div>
        ) : (
             <div className="w-full max-w-6xl mx-auto">
                <div className="max-w-2xl mx-auto w-full flex flex-col items-center gap-6">
                    <DropZone onFilesSelect={handleFilesSelected} multiple={true} disabled={isLoading} label="اختر ملفين أو أكثر"/>
                    
                    {processableFiles.length > 0 && (
                        <div className="w-full flex flex-wrap gap-3 justify-center">
                                {processableFiles.map(pf => (
                                    <div key={pf.id} className="flex items-center gap-3 bg-brand-gray-800 border border-brand-gray-700/60 hover:border-brand-gray-600 rounded-md px-3 py-2 transition-all duration-200 shadow-sm group min-w-[180px] max-w-[240px]">
                                        <div className="flex-shrink-0 text-brand-gray-400">
                                            {getFileIcon(pf.file)}
                                        </div>
                                        <div className="flex-grow min-w-0">
                                            <p className="text-xs font-medium text-brand-gray-200 truncate uppercase" title={pf.file.name}>
                                                {pf.file.name}
                                            </p>
                                            <div className="flex items-center gap-2 mt-0.5">
                                                <span className="text-[10px] text-brand-gray-500">{(pf.file.size / 1024).toFixed(0)} KB</span>
                                            </div>
                                        </div>
                                         <div className="flex-shrink-0 flex items-center gap-2">
                                            {renderStatusIndicator(pf.status)}
                                            <button 
                                                onClick={() => handleRemoveMultiFile(pf.id)} 
                                                className="text-brand-gray-600 hover:text-red-400 opacity-0 group-hover:opacity-100 transition-opacity"
                                                title="إزالة"
                                            >
                                                <TrashIcon className="h-3 w-3" />
                                            </button>
                                         </div>
                                    </div>
                                ))}
                        </div>
                    )}
                    
                    <div className="w-full flex items-stretch gap-2 mt-4">
                        <button onClick={handleProcessMultiFile} disabled={processableFiles.length < 2 || isLoading} className="flex-grow bg-brand-blue hover:bg-brand-blue-light text-white font-bold py-3 px-4 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-brand-blue/20">{isLoading ? <Spinner/> : 'استخراج البيانات'}</button>
                        {(processableFiles.length > 0) && !isLoading && <button onClick={handleClear} className="flex-shrink-0 bg-brand-gray-700 hover:bg-brand-gray-600 text-white font-bold p-3 rounded-lg transition-colors shadow-lg"><ClearIcon/></button>}
                    </div>
                </div>
                {processableFiles.some(f => f.status === 'done' && f.data) && <div className="uppercase"><ComparisonTable files={processableFiles} /></div>}
            </div>
        )}
      </main>
      <footer className="text-center mt-12 text-sm text-brand-gray-600 pb-12 px-4 border-t border-brand-gray-800/30 pt-8">
        <p className="font-bold text-brand-gray-400 mb-6">جميع حقوق الملكية محفوظة © ٢٠٢٥ - تم التطوير بواسطة مانع عزالدين</p>
        <div className="flex justify-center items-center gap-8 mt-6">
            <a 
              href="https://wa.me/967772655825" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="group relative flex items-center justify-center w-12 h-12 bg-brand-gray-800 rounded-2xl border border-white/5 text-brand-gray-500 hover:text-[#25D366] hover:border-[#25D366]/30 transition-all duration-500 transform hover:-translate-y-1 hover:shadow-[0_10px_20px_-5px_rgba(37,211,102,0.3)]"
              title="واتساب"
            >
                <WhatsAppIcon className="h-6 w-6" />
                <span className="absolute -bottom-8 opacity-0 group-hover:opacity-100 text-[10px] font-bold text-[#25D366] transition-opacity">WhatsApp</span>
            </a>
            <a 
              href="https://www.facebook.com/9l7iz" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="group relative flex items-center justify-center w-12 h-12 bg-brand-gray-800 rounded-2xl border border-white/5 text-brand-gray-500 hover:text-[#1877F2] hover:border-[#1877F2]/30 transition-all duration-500 transform hover:-translate-y-1 hover:shadow-[0_10px_20px_-5px_rgba(24,119,242,0.3)]"
              title="فيسبوك"
            >
                <FacebookIcon className="h-6 w-6" />
                <span className="absolute -bottom-8 opacity-0 group-hover:opacity-100 text-[10px] font-bold text-[#1877F2] transition-opacity">Facebook</span>
            </a>
            <a 
              href="https://www.instagram.com/9l7iz" 
              target="_blank" 
              rel="noopener noreferrer" 
              className="group relative flex items-center justify-center w-12 h-12 bg-brand-gray-800 rounded-2xl border border-white/5 text-brand-gray-500 hover:text-[#E4405F] hover:border-[#E4405F]/30 transition-all duration-500 transform hover:-translate-y-1 hover:shadow-[0_10px_20px_-5px_rgba(228,64,95,0.3)]"
              title="إنستغرام"
            >
                <InstagramIcon className="h-6 w-6" />
                <span className="absolute -bottom-8 opacity-0 group-hover:opacity-100 text-[10px] font-bold text-[#E4405F] transition-opacity">Instagram</span>
            </a>
        </div>
      </footer>
    </div>
  );
}

export default App;
