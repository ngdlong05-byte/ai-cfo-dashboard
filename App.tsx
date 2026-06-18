import React, { useState, useRef, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import * as htmlToImage from 'html-to-image';
import { UploadCloud, FileSpreadsheet, Bot, AlertTriangle, Loader2, HelpCircle, X, Download, Send } from 'lucide-react';
import ReactMarkdown from 'react-markdown';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Treemap
} from 'recharts';


const COLORS = ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];

interface DataRow {
  [key: string]: any;
}

interface TimeSeriesData {
  date: string;
  grossIncome: number;
  sales: number;
  cogs: number;
}

interface GenderAggregatedData {
  productLine: string;
  Male: number;
  Female: number;
  Other: number;
}

interface LocationAggregatedData {
  location: string;
  sales: number;
  grossIncome: number;
}

interface AggregatedData {
  productLine: string;
  sales: number;
  cogs: number;
  grossIncome: number;
  grossMarginPercent: number;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [fullJsonExcelData, setFullJsonExcelData] = useState<DataRow[] | null>(null);
  const [rawData, setRawData] = useState<DataRow[]>([]);
  const [aggregatedData, setAggregatedData] = useState<AggregatedData[]>([]);
  const [timeSeriesData, setTimeSeriesData] = useState<TimeSeriesData[]>([]);
  const [genderData, setGenderData] = useState<GenderAggregatedData[]>([]);
  const [locationData, setLocationData] = useState<LocationAggregatedData[]>([]);
  const [aiResult, setAiResult] = useState<string>('');
  const [loadingFile, setLoadingFile] = useState(false);
  const [loadingAi, setLoadingAi] = useState(false);
  const [error, setError] = useState<string>('');
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [dataQualityReport, setDataQualityReport] = useState<{ totalRows: number, missingValues: number, duplicateRows: number } | null>(null);
  const [showDataQualityModal, setShowDataQualityModal] = useState(false);

  // Chat state
  const [chatQuestion, setChatQuestion] = useState('');
  const [chatHistory, setChatHistory] = useState<{ role: 'user' | 'ai', content: string }[]>([]);
  const [chatLoading, setChatLoading] = useState(false);
  const [globalDateBounds, setGlobalDateBounds] = useState<{ min: string, max: string } | null>(null);
  const [filterDateRange, setFilterDateRange] = useState<{ start: string, end: string } | null>(null);

  const dashboardRef = useRef<HTMLDivElement>(null);

  // Column Mapping state
  const [availableColumns, setAvailableColumns] = useState<string[]>([]);
  const [isMapping, setIsMapping] = useState(false);
  const [mappings, setMappings] = useState({
    productLine: '',
    sales: '',
    cogs: '',
    grossIncome: '',
    date: '',
    gender: '',
    location: ''
  });
  const [appliedMappings, setAppliedMappings] = useState({
    productLine: '',
    sales: '',
    cogs: '',
    grossIncome: '',
    date: '',
    gender: '',
    location: ''
  });

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setError('');
    setRawData([]);
    setAggregatedData([]);
    setTimeSeriesData([]);
    setGenderData([]);
    setLocationData([]);
    setAiResult('');
    setChatHistory([]);
    setChatQuestion('');
    setLoadingFile(true);
    setIsMapping(false);

    const fn = uploadedFile.name.toLowerCase();
    
    if (fn.endsWith('.csv') || fn.endsWith('.tsv') || fn.endsWith('.txt')) {
      Papa.parse(uploadedFile, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        preview: 5,
        complete: (results) => {
          triggerProcessing(results.data as DataRow[], false);
        },
        error: (err) => {
          setError(`CSV Parse Error: ${err.message}`);
          setLoadingFile(false);
        }
      });
    } else {
      const fileReader = new FileReader();

      fileReader.onload = (event) => {
        try {
          const result = event.target?.result;
          if (!result) throw new Error("Failed to read file.");

          let parsedData: DataRow[] = [];

          if (fn.endsWith('.xlsx') || fn.endsWith('.xls')) {
            const workbook = XLSX.read(result, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            parsedData = XLSX.utils.sheet_to_json(worksheet);
            triggerProcessing(parsedData, true);
          } else if (fn.endsWith('.json')) {
            try {
              const jsonData = JSON.parse(result as string);
              parsedData = Array.isArray(jsonData) ? jsonData : [jsonData];
              triggerProcessing(parsedData, true);
            } catch(e: any) {
               setError(`JSON Parse Error: ${e.message}`);
               setLoadingFile(false);
            }
          } else {
            throw new Error('Unsupported file format. Please upload .csv, .tsv, .txt, .xlsx, .xls, or .json');
          }
        } catch (err: any) {
          setError(err.message || "An error occurred while reading the file.");
          setLoadingFile(false);
        }
      };

      if (fn.endsWith('.json')) {
        fileReader.readAsText(uploadedFile);
      } else {
        fileReader.readAsBinaryString(uploadedFile);
      }
    }
  };

  const triggerProcessing = (data: DataRow[], isFullData: boolean = false) => {
    if (data.length === 0) {
      setError("The uploaded file is empty or could not be properly parsed.");
      setLoadingFile(false);
      return;
    }

    if (isFullData) setFullJsonExcelData(data);
    setRawData(data.slice(0, 5));
    
    const sampleKeys = Object.keys(data[0] || {});
    setAvailableColumns(sampleKeys);
    
    // Auto map columns
    const normalizeKey = (k: string) => k.toLowerCase().replace(/[\s_]/g, '');
    const productLineKey = sampleKeys.find(k => normalizeKey(k).includes('productline') || normalizeKey(k).includes('category')) || sampleKeys[0] || '';
    const salesKey = sampleKeys.find(k => normalizeKey(k).includes('sales') || normalizeKey(k) === 'total' || normalizeKey(k).includes('revenue')) || '';
    const cogsKey = sampleKeys.find(k => normalizeKey(k).includes('cogs') || normalizeKey(k).includes('cost')) || '';
    const grossIncomeKey = sampleKeys.find(k => normalizeKey(k).includes('grossincome') || normalizeKey(k).includes('gross') || normalizeKey(k).includes('profit')) || '';
    const dateKey = sampleKeys.find(k => normalizeKey(k) === 'date' || normalizeKey(k).includes('time')) || '';
    const genderKey = sampleKeys.find(k => normalizeKey(k).includes('gender') || normalizeKey(k).includes('sex')) || '';
    const locationKey = sampleKeys.find(k => normalizeKey(k).includes('location') || normalizeKey(k).includes('city') || normalizeKey(k).includes('branch')) || '';

    const initialMappings = {
      productLine: productLineKey,
      sales: salesKey,
      cogs: cogsKey,
      grossIncome: grossIncomeKey,
      date: dateKey,
      gender: genderKey,
      location: locationKey
    };
    
    setMappings(initialMappings);

    // Always show mapping UI to let user confirm or skip columns
    setIsMapping(true);
    setLoadingFile(false);
  };

  useEffect(() => {
    if (aggregatedData.length > 0 && !loadingFile && !aiResult && !error && !loadingAi) {
      runAiAnalysis();
    }
  }, [aggregatedData, loadingFile, aiResult, error, loadingAi]);

  useEffect(() => {
    if (fullJsonExcelData && fullJsonExcelData.length > 0) {
      let missingVals = 0;
      let dupes = 0;
      const seen = new Set<string>();
      
      for (const row of fullJsonExcelData) {
        let isMissing = false;
        // Check missing values across mapped columns or all depending on requirement
        // Let's just check all available columns for simplicity, or just look for null/undefined/''
        for (const key in row) {
          if (row[key] === null || row[key] === undefined || row[key] === '') {
            isMissing = true;
          }
        }
        if (isMissing) missingVals++;

        const rowString = JSON.stringify(row);
        if (seen.has(rowString)) {
          dupes++;
        } else {
          seen.add(rowString);
        }
      }

      setDataQualityReport({
        totalRows: fullJsonExcelData.length,
        missingValues: missingVals,
        duplicateRows: dupes
      });
    } else {
      setDataQualityReport(null);
    }
  }, [fullJsonExcelData]);

  const handleApplyMapping = () => {
    const activeMappings = Object.values(mappings).filter(v => v !== '');
    if (activeMappings.length === 0) {
      setError("Please map at least one column to proceed.");
      return;
    }
    setError('');
    setIsMapping(false);
    setLoadingFile(true);
    setAppliedMappings(mappings);
    processAggregation(mappings);
  };

  const processAggregation = (currMappings: typeof mappings, dateFilter?: {start: string, end: string}) => {
    const aggregates: Record<string, { sales: number; cogs: number; grossIncome: number }> = {};
    const timeAggregates: Record<string, { sales: number; cogs: number; grossIncome: number }> = {};
    const genderAggregates: Record<string, { Male: number; Female: number; Other: number }> = {};
    const locationAggregates: Record<string, { sales: number; grossIncome: number }> = {};

    const filterStartTime = dateFilter ? new Date(dateFilter.start).getTime() : -Infinity;
    const filterEndTime = dateFilter ? new Date(dateFilter.end).getTime() : Infinity;

    const processRow = (row: DataRow) => {
      let rowDateStr = '';
      let rowTime = NaN;
      if (currMappings.date) {
        const dateVal = row[currMappings.date];
        if (dateVal) {
          // Attempt to parse string to a Date object safely
          const jsDate = new Date(dateVal);
          rowTime = jsDate.getTime();
          if (!isNaN(rowTime)) {
            // Convert to YYYY-MM-DD safely
            // Using local year, month, day to avoid timezone shifting issues
            const yyyy = jsDate.getFullYear();
            const mm = String(jsDate.getMonth() + 1).padStart(2, '0');
            const dd = String(jsDate.getDate()).padStart(2, '0');
            rowDateStr = `${yyyy}-${mm}-${dd}`;
          }
        }
      }

      if (dateFilter) {
        if (isNaN(rowTime) || rowTime < filterStartTime || rowTime > filterEndTime) {
          return;
        }
      }

      const pLine = String(row[currMappings.productLine] || 'Unknown');
      
      const parseNumber = (val: any) => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const cleaned = String(val).replace(/[^0-9.-]+/g, '');
        return Number(cleaned) || 0;
      };

      const sales = parseNumber(row[currMappings.sales]);
      const cogs = parseNumber(row[currMappings.cogs]);
      const gInc = parseNumber(row[currMappings.grossIncome]);

      if (!aggregates[pLine]) {
        aggregates[pLine] = { sales: 0, cogs: 0, grossIncome: 0 };
      }
      aggregates[pLine].sales += sales;
      aggregates[pLine].cogs += cogs;
      aggregates[pLine].grossIncome += gInc;

      if (rowDateStr) {
        if (!timeAggregates[rowDateStr]) timeAggregates[rowDateStr] = { sales: 0, cogs: 0, grossIncome: 0 };
        timeAggregates[rowDateStr].sales += sales;
        timeAggregates[rowDateStr].cogs += cogs;
        timeAggregates[rowDateStr].grossIncome += gInc;
      }

      if (currMappings.gender) {
        const genderVal = String(row[currMappings.gender] || 'Other');
        let gKey = 'Other';
        if (genderVal.toLowerCase() === 'male' || genderVal.toLowerCase() === 'm') gKey = 'Male';
        else if (genderVal.toLowerCase() === 'female' || genderVal.toLowerCase() === 'f') gKey = 'Female';
        
        if (!genderAggregates[pLine]) genderAggregates[pLine] = { Male: 0, Female: 0, Other: 0 };
        (genderAggregates[pLine] as any)[gKey]++;
      }

      if (currMappings.location) {
        const locVal = String(row[currMappings.location] || 'Unknown');
        if (!locationAggregates[locVal]) locationAggregates[locVal] = { sales: 0, grossIncome: 0 };
        locationAggregates[locVal].sales += sales;
        locationAggregates[locVal].grossIncome += gInc;
      }
    };

    const finalizeAggregation = () => {
      const aggArray: AggregatedData[] = Object.keys(aggregates).map(pLine => {
        const acc = aggregates[pLine];
        const gmPercent = acc.sales > 0 ? (acc.grossIncome / acc.sales) * 100 : 0;
        return {
          productLine: pLine,
          sales: acc.sales,
          cogs: acc.cogs,
          grossIncome: acc.grossIncome,
          grossMarginPercent: Number(gmPercent.toFixed(2))
        };
      });

      setAggregatedData(aggArray);

      if (currMappings.date) {
        const timeArr = Object.keys(timeAggregates).map(d => ({
          date: d,
          sales: Number(timeAggregates[d].sales.toFixed(2)),
          cogs: Number(timeAggregates[d].cogs.toFixed(2)),
          grossIncome: Number(timeAggregates[d].grossIncome.toFixed(2))
        })).sort((a, b) => {
          const timeA = new Date(a.date).getTime();
          const timeB = new Date(b.date).getTime();
          if (!isNaN(timeA) && !isNaN(timeB)) {
            return timeA - timeB;
          }
          return a.date.localeCompare(b.date);
        });
        setTimeSeriesData(timeArr);
        
        if (!dateFilter && timeArr.length > 0) {
          const minDate = timeArr[0].date;
          const maxDate = timeArr[timeArr.length - 1].date;
          setGlobalDateBounds({ min: minDate, max: maxDate });
          if (!filterDateRange) {
            setFilterDateRange({ start: minDate, end: maxDate });
          }
        }
      } else {
        setTimeSeriesData([]);
        if (!dateFilter) {
          setGlobalDateBounds(null);
          setFilterDateRange(null);
        }
      }

      if (currMappings.gender) {
        const genArr = Object.keys(genderAggregates).map(pLine => ({
          productLine: pLine,
          Male: genderAggregates[pLine].Male,
          Female: genderAggregates[pLine].Female,
          Other: genderAggregates[pLine].Other
        }));
        setGenderData(genArr);
      } else {
        setGenderData([]);
      }

      if (currMappings.location) {
        const locArr = Object.keys(locationAggregates).map(loc => ({
          location: loc,
          sales: Number(locationAggregates[loc].sales.toFixed(2)),
          grossIncome: Number(locationAggregates[loc].grossIncome.toFixed(2))
        }));
        setLocationData(locArr);
      } else {
        setLocationData([]);
      }

      setLoadingFile(false);
    };

    if (fullJsonExcelData) {
      fullJsonExcelData.forEach(processRow);
      finalizeAggregation();
    } else if (file) {
      Papa.parse(file, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true,
        complete: (results) => {
          const parsedData = results.data as DataRow[];
          setFullJsonExcelData(parsedData);
          parsedData.forEach((row: any) => processRow(row));
          finalizeAggregation();
        },
        error: (err) => {
          setError(`Processing Error: ${err.message}`);
          setLoadingFile(false);
        }
      });
    } else {
      setError("No file loaded");
      setLoadingFile(false);
    }
  };

  const handleDownloadDashboard = async () => {
    if (!dashboardRef.current) return;
    try {
      const dataUrl = await htmlToImage.toPng(dashboardRef.current, {
        backgroundColor: '#020617',
        pixelRatio: 2,
      });
      const a = document.createElement('a');
      a.href = dataUrl;
      a.download = 'AI_CFO_Dashboard.png';
      a.click();
    } catch (err) {
      console.error("Failed to download dashboard", err);
    }
  };

  const handleDownloadSection = async (elementId: string, fileName: string) => {
    const element = document.getElementById(elementId);
    if (!element) return;
    try {
      const dataUrl = await htmlToImage.toPng(element, {
        backgroundColor: '#0f172a',
        pixelRatio: 2,
        style: {
          padding: '24px',
          borderRadius: '16px'
        }
      });
      const a = document.createElement('a');
      a.href = dataUrl;
      a.download = `${fileName}.png`;
      a.click();
    } catch (error) {
      console.error('Failed to download section', error);
    }
  };

  const handleApplyDateFilter = () => {
    if (filterDateRange) {
      setLoadingFile(true);
      setAiResult('');
      setChatHistory([]);
      // Let React update the loading state before processing
      setTimeout(() => {
        processAggregation(appliedMappings, filterDateRange);
      }, 0);
    }
  };

  const handleCleanData = () => {
    if (!fullJsonExcelData) return;
    
    setLoadingFile(true);
    setTimeout(() => {
      const seen = new Set<string>();
      const cleanedData: DataRow[] = [];
      
      for (const row of fullJsonExcelData) {
        // Remove exact duplicates
        const rowString = JSON.stringify(row);
        if (seen.has(rowString)) continue;
        seen.add(rowString);
        
        // Impute missing values
        const newRow: DataRow = { ...row };
        for (const key in newRow) {
          if (newRow[key] === null || newRow[key] === undefined || String(newRow[key]).trim() === '') {
             if (key === appliedMappings.sales || key === appliedMappings.cogs || key === appliedMappings.grossIncome) {
               newRow[key] = 0;
             } else {
               newRow[key] = 'Unknown';
             }
          }
        }
        cleanedData.push(newRow);
      }
      
      setFullJsonExcelData(cleanedData);
      processAggregation(appliedMappings, filterDateRange || undefined);
    }, 0);
  };

  const handleDownloadCleanedData = () => {
    if (!fullJsonExcelData || fullJsonExcelData.length === 0) return;
    
    const headers = Object.keys(fullJsonExcelData[0]);
    const csvContent = fullJsonExcelData.map(row => {
      return headers.map(header => {
        const val = row[header];
        if (typeof val === 'string') {
          return `"${val.replace(/"/g, '""')}"`; // Escape quotes safely
        }
        return val;
      }).join(",");
    });
    
    csvContent.unshift(headers.join(","));
    const finalCsv = csvContent.join("\n");
    
    const blob = new Blob([finalCsv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "cleaned_data.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleDownloadCSV = () => {
    if (!aggregatedData.length) return;
    const headers = ["Product Line"];
    if (appliedMappings.sales) headers.push("Sales");
    if (appliedMappings.cogs) headers.push("COGS");
    if (appliedMappings.grossIncome) headers.push("Gross Income");
    if (appliedMappings.sales && appliedMappings.grossIncome) headers.push("Gross Margin (%)");
    
    const csvRows = [headers.join(',')];
    aggregatedData.forEach(row => {
      const rowData = [row.productLine];
      if (appliedMappings.sales) rowData.push(row.sales.toString());
      if (appliedMappings.cogs) rowData.push(row.cogs.toString());
      if (appliedMappings.grossIncome) rowData.push(row.grossIncome.toString());
      if (appliedMappings.sales && appliedMappings.grossIncome) rowData.push(row.grossMarginPercent.toString());
      csvRows.push(rowData.join(','));
    });
    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ai_cfo_aggregated_data.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleAskQuestion = async () => {
    if (!chatQuestion.trim() || !aggregatedData.length) return;
    
    const headers = ["Product Line"];
    if (appliedMappings.sales) headers.push("Sales");
    if (appliedMappings.cogs) headers.push("COGS");
    if (appliedMappings.grossIncome) headers.push("Gross Income");
    if (appliedMappings.sales && appliedMappings.grossIncome) headers.push("Gross Margin (%)");
    
    let dataText = headers.join(" | ") + "\n";
    dataText += headers.map(() => "---").join("|") + "\n";
    
    aggregatedData.forEach(row => {
      const rowData = [row.productLine];
      if (appliedMappings.sales) rowData.push(row.sales.toFixed(2));
      if (appliedMappings.cogs) rowData.push(row.cogs.toFixed(2));
      if (appliedMappings.grossIncome) rowData.push(row.grossIncome.toFixed(2));
      if (appliedMappings.sales && appliedMappings.grossIncome) rowData.push(row.grossMarginPercent + "%");
      dataText += rowData.join(" | ") + "\n";
    });

    if (genderData.length > 0) {
      dataText += "\n=== GENDER RATIO (Count) ===\nProduct Line | Male | Female | Other\n---|---|---|---\n";
      genderData.forEach(row => {
        dataText += `${row.productLine} | ${row.Male} | ${row.Female} | ${row.Other}\n`;
      });
    }

    if (locationData.length > 0) {
      dataText += "\n=== LOCATION PERFORMANCE ===\nLocation | Sales | Gross Income\n---|---|---\n";
      locationData.forEach(row => {
        dataText += `${row.location} | ${row.sales.toFixed(2)} | ${row.grossIncome.toFixed(2)}\n`;
      });
    }

    const newQuestion = chatQuestion.trim();
    setChatQuestion('');
    setChatHistory(prev => [...prev, { role: 'user', content: newQuestion }]);
    setChatLoading(true);

    try {
      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ 
          dataText, 
          question: newQuestion,
          history: chatHistory 
        }),
      });

      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Failed to get answer');

      setChatHistory(prev => [...prev, { role: 'ai', content: data.result }]);
    } catch (error: any) {
      let errMsg = error.message;
      try {
        const parsed = JSON.parse(errMsg);
        if (parsed?.error?.message) errMsg = parsed.error.message;
      } catch (e) {}
      setChatHistory(prev => [...prev, { role: 'ai', content: `Error: ${errMsg}` }]);
    } finally {
      setChatLoading(false);
    }
  };

  const runAiAnalysis = async () => {
    if (aggregatedData.length === 0) return;
    setLoadingAi(true);
    setAiResult('');
    setError('');

    try {
      const headers = ["Product Line"];
      if (appliedMappings.sales) headers.push("Sales");
      if (appliedMappings.cogs) headers.push("COGS");
      if (appliedMappings.grossIncome) headers.push("Gross Income");
      if (appliedMappings.sales && appliedMappings.grossIncome) headers.push("Gross Margin (%)");
      
      let dataText = headers.join(" | ") + "\n";
      dataText += headers.map(() => "---").join("|") + "\n";

      aggregatedData.forEach(row => {
        const rowData = [row.productLine];
        if (appliedMappings.sales) rowData.push(row.sales.toFixed(2));
        if (appliedMappings.cogs) rowData.push(row.cogs.toFixed(2));
        if (appliedMappings.grossIncome) rowData.push(row.grossIncome.toFixed(2));
        if (appliedMappings.sales && appliedMappings.grossIncome) rowData.push(row.grossMarginPercent + "%");
        dataText += rowData.join(" | ") + "\n";
      });

      if (genderData.length > 0) {
        dataText += "\n=== GENDER RATIO (Count) ===\nProduct Line | Male | Female | Other\n---|---|---|---\n";
        genderData.forEach(row => {
          dataText += `${row.productLine} | ${row.Male} | ${row.Female} | ${row.Other}\n`;
        });
      }

      if (locationData.length > 0) {
        dataText += "\n=== LOCATION PERFORMANCE ===\nLocation | Sales | Gross Income\n---|---|---\n";
        locationData.forEach(row => {
          dataText += `${row.location} | ${row.sales.toFixed(2)} | ${row.grossIncome.toFixed(2)}\n`;
        });
      }

      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ dataText })
      });

      const resultData = await response.json();
      if (!response.ok) {
        throw new Error(resultData.error || 'Failed to analyze data.');
      }

      setAiResult(resultData.result);
    } catch (err: any) {
      let errMsg = err.message || 'Failed to analyze data.';
      try {
        const parsed = JSON.parse(errMsg);
        if (parsed?.error?.message) errMsg = parsed.error.message;
      } catch (e) {}
      setError(`AI Analysis Error: ${errMsg}`);
    } finally {
      setLoadingAi(false);
    }
  };

  return (
    <div className="h-screen w-full bg-[#020617] text-slate-100 relative overflow-hidden flex flex-col font-sans">
      {/* Background Mesh Orbs */}
      <div className="absolute top-[-100px] left-[-100px] w-[500px] h-[500px] bg-emerald-500/20 rounded-full blur-[120px] pointer-events-none"></div>
      <div className="absolute bottom-[-100px] right-[-100px] w-[500px] h-[500px] bg-blue-600/20 rounded-full blur-[120px] pointer-events-none"></div>

      {/* Header Section */}
      <header className="relative z-10 flex items-center justify-between px-8 py-6 border-b border-white/10 bg-white/5 backdrop-blur-md shrink-0">
        <div className="flex items-center space-x-3">
          <div className="p-2 bg-emerald-500/20 rounded-lg border border-emerald-500/40">
            <Bot className="w-6 h-6 text-emerald-400" />
          </div>
          <h1 className="text-xl font-bold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white to-slate-400">
            AI CFO <span className="text-emerald-400 font-light">— Smart Costing & COGS Analyzer</span>
          </h1>
        </div>
        <div className="flex items-center space-x-4">
          <button 
            onClick={() => setShowHelpModal(true)}
            className="p-2 rounded-full hover:bg-white/10 transition-colors text-slate-400 hover:text-emerald-400 flex items-center justify-center shrink-0"
            title="Format Instructions"
          >
            <HelpCircle className="w-5 h-5" />
          </button>
          <div className="text-xs font-mono text-slate-400 border border-white/10 px-3 py-1 rounded-full bg-white/5">
            v2.5-FLASH ACTIVE
          </div>
        </div>
      </header>

      {/* Main Layout */}
      <main className="relative z-10 flex flex-col lg:flex-row flex-1 p-6 lg:space-x-6 space-y-6 lg:space-y-0 overflow-hidden max-w-[1400px] mx-auto w-full">
        
        {/* Left Panel: Sidebar & Upload */}
        <aside className="w-full lg:w-1/3 flex flex-col space-y-6 h-full overflow-y-auto">
          {/* Error Banner */}
          {error && (
            <div className="bg-red-500/10 text-red-400 p-4 rounded-2xl flex items-start space-x-3 border border-red-500/20 backdrop-blur-md">
              <AlertTriangle className="w-5 h-5 flex-shrink-0 mt-0.5" />
              <div>
                <h3 className="font-semibold text-sm">Error</h3>
                <p className="text-sm mt-1 opacity-90">{error}</p>
              </div>
            </div>
          )}

          {/* Upload Box */}
          <label htmlFor="dropzone-file" className="bg-white/10 backdrop-blur-xl border border-white/10 rounded-2xl p-6 flex flex-col items-center justify-center border-dashed border-2 hover:border-emerald-500/50 cursor-pointer transition-colors relative group">
            <UploadCloud className="w-10 h-10 text-slate-400 mb-4 group-hover:text-emerald-400 transition-colors" />
            {file ? (
              <div className="flex flex-col items-center">
                 <p className="text-sm font-medium text-slate-200 flex items-center text-center"><FileSpreadsheet className="w-4 h-4 mr-2 text-emerald-400" /> {file.name}</p>
                 <div className="mt-4 w-full py-2 px-4 bg-white/5 group-hover:bg-white/10 border border-white/20 rounded-lg text-xs font-semibold text-slate-300 transition-all text-center">REPLACE FILE</div>
              </div>
            ) : (
              <div className="flex flex-col items-center text-center">
                 <p className="mb-2 text-sm text-slate-200 font-medium"><span className="text-emerald-400">Click to upload</span> or drag and drop</p>
                 <p className="text-xs text-slate-500">Supports .csv, .tsv, .xlsx, .xls, .json</p>
              </div>
            )}
            <input id="dropzone-file" type="file" className="hidden" accept=".csv, .tsv, .txt, .xlsx, .xls, .json" onChange={handleFileUpload} />
            {loadingFile && (
               <div className="absolute inset-0 bg-slate-900/50 backdrop-blur-sm rounded-2xl flex items-center justify-center">
                  <Loader2 className="w-6 h-6 text-emerald-400 animate-spin" />
               </div>
            )}
          </label>

          {/* Column Mapping Form */}
          {isMapping && availableColumns.length > 0 && (
            <div className="p-4 bg-blue-500/10 border border-blue-500/20 rounded-xl space-y-4">
              <div className="flex items-start">
                <h3 className="text-sm font-semibold text-blue-300">
                  Map Columns
                  <span className="block text-xs font-normal text-blue-200/70 mt-1">
                    Select the columns from your file that match the required data fields.
                  </span>
                </h3>
              </div>
              
              <div className="space-y-3">
                 {[
                   { label: 'Product Line (Category)', key: 'productLine' },
                   { label: 'Sales/Total (Optional)', key: 'sales' },
                   { label: 'COGS/Cost (Optional)', key: 'cogs' },
                   { label: 'Gross Income/Profit (Optional)', key: 'grossIncome' },
                   { label: 'Date (Optional)', key: 'date' },
                   { label: 'Gender (Optional)', key: 'gender' },
                   { label: 'Location/Branch (Optional)', key: 'location' }
                 ].map(field => (
                   <div key={field.key} className="flex flex-col space-y-1">
                     <label className="text-[10px] uppercase font-bold text-slate-400">{field.label}</label>
                     <select 
                       value={(mappings as any)[field.key] || ''}
                       onChange={(e) => setMappings({...mappings, [field.key]: e.target.value})}
                       className="w-full bg-slate-900/80 border border-white/10 rounded-lg p-2 text-xs text-slate-200 outline-none focus:border-emerald-500/50 appearance-none"
                     >
                       <option value="">-- Select Column --</option>
                       {availableColumns.map(col => (
                         <option key={col} value={col}>{col}</option>
                       ))}
                     </select>
                   </div>
                 ))}
                 
                 <button 
                  onClick={handleApplyMapping}
                  className="w-full py-3 mt-2 bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold rounded-lg transition-colors"
                 >
                   Apply Mapping
                 </button>
              </div>
            </div>
          )}

          {/* Summary Metrics Mini-Cards */}
          {aggregatedData.length > 0 && !loadingFile && (
            <div className="grid grid-cols-2 gap-4">
              <div className="bg-white/5 backdrop-blur-lg border border-white/10 rounded-xl p-4">
                {appliedMappings.sales ? (
                  <>
                    <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-1">Total Sales</p>
                    <p className="text-lg font-bold">
                       ${aggregatedData.reduce((sum, r) => sum + r.sales, 0).toLocaleString('en-US', {maximumFractionDigits:0})}
                    </p>
                  </>
                ) : (
                  <>
                    <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-1">Total Categories</p>
                    <p className="text-lg font-bold">{aggregatedData.length}</p>
                  </>
                )}
              </div>
              <div className="bg-white/5 backdrop-blur-lg border border-white/10 rounded-xl p-4">
                {appliedMappings.sales && appliedMappings.grossIncome ? (
                  <>
                    <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-1">Avg Margin</p>
                    <p className="text-lg font-bold text-emerald-400">
                       {(aggregatedData.reduce((sum, r) => sum + r.grossMarginPercent, 0) / aggregatedData.length).toFixed(2)}%
                    </p>
                  </>
                ) : appliedMappings.grossIncome ? (
                  <>
                    <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-1">Gross Income</p>
                    <p className="text-lg font-bold text-emerald-400">
                       ${aggregatedData.reduce((sum, r) => sum + r.grossIncome, 0).toLocaleString('en-US', {maximumFractionDigits:0})}
                    </p>
                  </>
                ) : (
                  <>
                    <p className="text-[10px] text-slate-400 uppercase tracking-widest mb-1">Processed</p>
                    <p className="text-lg font-bold text-emerald-400">Done</p>
                  </>
                )}
              </div>
            </div>
          )}

          {/* Data Quality Report */}
          {dataQualityReport && !loadingFile && !isMapping && (
             <div className="bg-slate-900/50 border border-slate-700/50 rounded-xl p-4 mt-4 space-y-3">
               <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest flex items-center justify-between">
                 <span>Data Quality</span>
                 {(dataQualityReport.missingValues > 0 || dataQualityReport.duplicateRows > 0) && (
                   <AlertTriangle className="w-4 h-4 text-amber-400" />
                 )}
               </h3>
               <div className="flex gap-4 text-xs text-slate-400 flex-wrap">
                 <div>Total Rows: <span className="text-white font-mono">{dataQualityReport.totalRows}</span></div>
                 <div>Missing: <span className={dataQualityReport.missingValues > 0 ? "text-amber-400 font-mono" : "text-white font-mono"}>{dataQualityReport.missingValues}</span></div>
                 <div>Dupes: <span className={dataQualityReport.duplicateRows > 0 ? "text-amber-400 font-mono" : "text-white font-mono"}>{dataQualityReport.duplicateRows}</span></div>
               </div>
               
               <div className="flex items-center gap-2 pt-2 border-t border-slate-800">
                   <button onClick={handleCleanData} className="flex-1 py-1.5 bg-blue-600/20 hover:bg-blue-600/40 text-blue-400 text-xs font-medium rounded-lg transition-colors border border-blue-500/20">Clean Data</button>
                   <button onClick={handleDownloadCleanedData} className="flex-1 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-medium rounded-lg transition-colors border border-slate-700">Download CSV</button>
               </div>
             </div>
          )}

          {/* Action Button */}
          {aggregatedData.length > 0 && !loadingFile && !isMapping && (
            <div className="space-y-3">
              <button
                onClick={runAiAnalysis}
                disabled={loadingAi}
                className="w-full py-4 bg-emerald-500 text-slate-950 font-bold rounded-xl shadow-[0_0_20px_rgba(16,185,129,0.3)] hover:scale-[1.02] active:scale-[0.98] transition-all flex items-center justify-center space-x-2 disabled:opacity-70 disabled:hover:scale-100 disabled:cursor-not-allowed"
              >
                {loadingAi ? <Loader2 className="w-5 h-5 animate-spin" /> : <Bot className="w-5 h-5" />}
                <span>{loadingAi ? 'ĐANG PHÂN TÍCH...' : 'PHÂN TÍCH BẰNG AI CFO'}</span>
              </button>
              <button 
                onClick={() => setIsMapping(true)}
                className="w-full py-2 bg-white/5 hover:bg-white/10 text-slate-300 text-xs font-medium rounded-lg transition-colors border border-white/10"
              >
                Re-map Columns
              </button>
            </div>
          )}

          <div className="flex-1"></div>

          {/* Raw Data Preview */}
          {rawData.length > 0 && !loadingFile && (
            <div className="p-4 bg-slate-900/50 rounded-xl border border-white/5">
              <p className="text-[10px] text-slate-500 mb-2 font-mono uppercase">Raw Data Preview (Head 5)</p>
              <div className="text-[9px] text-slate-400 font-mono space-y-1">
                {rawData.slice(0, 5).map((row, i) => (
                  <div key={i} className="truncate">
                    {Object.values(row).slice(0, 5).join(' | ')}
                  </div>
                ))}
              </div>
            </div>
          )}
        </aside>

        {/* Right Panel: Data Table, Charts & AI Insight */}
        <section className="flex-[2] flex flex-col space-y-6 h-full overflow-y-auto hide-scrollbar pb-6 relative">
          
          {/* Empty State / Intro */}
          {aggregatedData.length === 0 && !loadingFile && (
            <div className="flex-1 flex flex-col items-center justify-center p-12 text-center h-full border border-white/5 bg-white/5 rounded-2xl backdrop-blur-sm relative z-10 min-h-[500px]">
               <Bot className="w-16 h-16 text-emerald-500/50 mb-6" />
               <h2 className="text-2xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-emerald-400 to-blue-400 mb-4">Welcome to AI CFO</h2>
               <p className="text-slate-400 max-w-lg leading-relaxed text-sm mb-6">
                 Upload your financial data file to instantly generate visual dashboards, aggregate metrics by product line, and receive AI-driven insights on cost management and profit margins.
               </p>
               
               <div className="grid grid-cols-1 md:grid-cols-3 gap-4 w-full max-w-2xl text-left border-t border-white/10 pt-8 mt-4">
                 <div className="bg-slate-900/50 p-5 rounded-xl border border-white/5">
                   <div className="w-8 h-8 rounded-full bg-emerald-500/20 flex items-center justify-center mb-3">
                     <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
                   </div>
                   <h3 className="font-bold text-slate-200 text-sm mb-2">1. Upload Data</h3>
                   <p className="text-xs text-slate-500 leading-relaxed">Simply drop your CSV, Excel, or JSON files. The system will automatically detect the columns.</p>
                 </div>
                 
                 <div className="bg-slate-900/50 p-5 rounded-xl border border-white/5">
                   <div className="w-8 h-8 rounded-full bg-blue-500/20 flex items-center justify-center mb-3">
                     <BarChart className="w-4 h-4 text-blue-400" />
                   </div>
                   <h3 className="font-bold text-slate-200 text-sm mb-2">2. Map Columns</h3>
                   <p className="text-xs text-slate-500 leading-relaxed">Map your columns such as Sales, COGS, Date, Gender, Location, and Product Line for the dynamic dashboard to render.</p>
                 </div>
                 
                 <div className="bg-slate-900/50 p-5 rounded-xl border border-white/5">
                   <div className="w-8 h-8 rounded-full bg-purple-500/20 flex items-center justify-center mb-3">
                     <Bot className="w-4 h-4 text-purple-400" />
                   </div>
                   <h3 className="font-bold text-slate-200 text-sm mb-2">3. Analyze & Ask</h3>
                   <p className="text-xs text-slate-500 leading-relaxed">Let our Gemini AI analyze your performance, spot anomalies, and ask follow-up questions to understand your data deeply.</p>
                 </div>
               </div>
            </div>
          )}

          {/* Download Dashboard Action */}
          {aggregatedData.length > 0 && !loadingFile && (
            <div className="flex justify-end sticky top-0 z-20 mb-[-1rem]">
              <button
                onClick={handleDownloadDashboard}
                className="bg-slate-800/90 hover:bg-slate-700/90 text-slate-200 px-4 py-2 rounded-xl text-xs font-bold transition-colors flex items-center shadow-lg border border-white/10 backdrop-blur-md"
              >
                <Download className="w-4 h-4 mr-2" /> Export Dashboard
              </button>
            </div>
          )}

          <div ref={dashboardRef} className="flex flex-col space-y-6 bg-[#020617] p-2 rounded-2xl relative z-10">
          
          {/* Global Date Filter */}
          {globalDateBounds && filterDateRange && (
            <div className={`bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-4 flex flex-col sm:flex-row items-center gap-4 justify-between shrink-0 ${loadingFile ? 'opacity-50 pointer-events-none' : ''}`}>
              <div className="flex flex-col">
                 <h2 className="text-xs font-bold text-slate-300 uppercase tracking-widest mb-1">Date Range Filter</h2>
                 <p className="text-[10px] text-slate-500">Filter your dashboard metrics by a specific time period</p>
              </div>
              <div className="flex items-center gap-3">
                <div className="flex flex-col gap-1">
                  <label className="text-[10px] text-slate-400 font-medium">Start Date</label>
                  <input 
                    type="date"
                    min={globalDateBounds.min}
                    max={filterDateRange.end || globalDateBounds.max}
                    value={filterDateRange.start}
                    onChange={(e) => setFilterDateRange(prev => prev ? {...prev, start: e.target.value} : null)}
                    className="bg-slate-900 border border-white/10 rounded-lg px-3 py-1.5 text-xs text-slate-200 outline-none focus:border-emerald-500/50"
                  />
                </div>
                <div className="text-slate-600 mt-5">-</div>
                <div className="flex flex-col gap-1">
                  <label className="text-[10px] text-slate-400 font-medium">End Date</label>
                  <input 
                    type="date"
                    min={filterDateRange.start || globalDateBounds.min}
                    max={globalDateBounds.max}
                    value={filterDateRange.end}
                    onChange={(e) => setFilterDateRange(prev => prev ? {...prev, end: e.target.value} : null)}
                    className="bg-slate-900 border border-white/10 rounded-lg px-3 py-1.5 text-xs text-slate-200 outline-none focus:border-emerald-500/50"
                  />
                </div>
                <button
                  onClick={handleApplyDateFilter}
                  className="mt-5 bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-1.5 rounded-lg text-xs font-bold transition-colors shadow-lg border border-emerald-500/40"
                >
                  Apply Filter
                </button>
              </div>
            </div>
          )}

          {/* Main Data Table */}
          {aggregatedData.length > 0 && !loadingFile && (
            <div id="chart-aggregation-summary" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl overflow-hidden flex flex-col shrink-0 lg:min-h-[40%] group relative">
              <div className="px-6 py-4 border-b border-white/10 bg-white/5 shrink-0 flex items-center justify-between">
                <h2 className="text-sm font-semibold text-slate-300">Product Line Aggregation Summary</h2>
                <div className="flex gap-2">
                  <button onClick={() => handleDownloadSection('chart-aggregation-summary', 'Aggregation_Summary')} className="p-1.5 rounded-lg hover:bg-white/10 text-slate-400 hover:text-emerald-400 transition-colors flex items-center justify-center border border-transparent hover:border-white/10" title="Download Image">
                    <Download className="w-4 h-4" />
                  </button>
                  <button
                    onClick={handleDownloadCSV}
                    className="p-1.5 rounded-lg hover:bg-white/10 text-slate-400 hover:text-emerald-400 transition-colors flex items-center justify-center border border-transparent hover:border-white/10"
                    title="Download CSV"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                  </button>
                </div>
              </div>
              <div className="flex-1 overflow-auto">
                <table className="w-full text-left text-xs whitespace-nowrap">
                  <thead>
                    <tr className="text-slate-400 border-b border-white/10 bg-slate-900/50 sticky top-0 backdrop-blur-sm z-10">
                      <th className="px-6 py-4 font-medium uppercase tracking-wider text-[10px]">Product Line</th>
                      {appliedMappings.sales && <th className="px-6 py-4 font-medium uppercase tracking-wider text-[10px] text-right">Sales ($)</th>}
                      {appliedMappings.cogs && <th className="px-6 py-4 font-medium uppercase tracking-wider text-[10px] text-right">COGS ($)</th>}
                      {appliedMappings.grossIncome && <th className="px-6 py-4 font-medium uppercase tracking-wider text-[10px] text-right">Profit ($)</th>}
                      {appliedMappings.sales && appliedMappings.grossIncome && <th className="px-6 py-4 font-medium uppercase tracking-wider text-[10px] text-right text-emerald-400">Margin (%)</th>}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-white/5 text-slate-300">
                    {aggregatedData.map((row, i) => (
                      <tr key={i} className="hover:bg-white/5 transition-colors">
                        <td className="px-6 py-3 font-medium">{row.productLine}</td>
                        {appliedMappings.sales && <td className="px-6 py-3 text-right">{row.sales.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>}
                        {appliedMappings.cogs && <td className="px-6 py-3 text-right opacity-80">{row.cogs.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</td>}
                        {appliedMappings.grossIncome && <td className="px-6 py-3 text-right">{row.grossIncome.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>}
                        {appliedMappings.sales && appliedMappings.grossIncome && <td className="px-6 py-3 text-right font-semibold text-emerald-400">{row.grossMarginPercent}%</td>}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Charts Section */}
          {aggregatedData.length > 0 && !loadingFile && (
            <div className="flex flex-col gap-6 shrink-0">
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                {/* Bar Chart */}
                <div id="chart-financial-overview" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 h-[350px] flex flex-col group relative">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Financial Overview</h3>
                    <button onClick={() => handleDownloadSection('chart-financial-overview', 'Financial_Overview')} className="p-1.5 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg text-slate-400 hover:text-white transition-colors" title="Download Chart">
                      <Download className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="flex-1 min-h-0 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={aggregatedData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                        <XAxis dataKey="productLine" stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} />
                        <YAxis stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(value) => `$${value/1000}k`} />
                        <RechartsTooltip 
                          cursor={{ fill: '#ffffff10' }}
                          contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #1e293b', borderRadius: '8px', color: '#f8fafc' }}
                          itemStyle={{ fontSize: '12px' }}
                        />
                        <Legend wrapperStyle={{ fontSize: '10px', paddingTop: '10px' }} />
                        {appliedMappings.sales && <Bar dataKey="sales" name="Sales ($)" fill="#3b82f6" radius={[4, 4, 0, 0]} />}
                        {appliedMappings.cogs && <Bar dataKey="cogs" name="COGS ($)" fill="#ef4444" radius={[4, 4, 0, 0]} />}
                        {appliedMappings.grossIncome && <Bar dataKey="grossIncome" name="Gross Income ($)" fill="#10b981" radius={[4, 4, 0, 0]} />}
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Pie Chart */}
                {appliedMappings.sales && appliedMappings.grossIncome && (
                <div id="chart-margin-distribution" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 h-[350px] flex flex-col group relative">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Margin Distribution (%)</h3>
                    <button onClick={() => handleDownloadSection('chart-margin-distribution', 'Margin_Distribution')} className="p-1.5 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg text-slate-400 hover:text-white transition-colors" title="Download Chart">
                      <Download className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="flex-1 min-h-0 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <RechartsTooltip 
                          contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #1e293b', borderRadius: '8px', color: '#f8fafc' }}
                          itemStyle={{ fontSize: '12px' }}
                          formatter={(value) => `${value}%`}
                        />
                        <Legend wrapperStyle={{ fontSize: '10px', paddingTop: '10px' }} />
                        <Pie
                          data={aggregatedData}
                          dataKey="grossMarginPercent"
                          nameKey="productLine"
                          cx="50%"
                          cy="50%"
                          innerRadius={60}
                          outerRadius={90}
                          paddingAngle={2}
                        >
                          {aggregatedData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} stroke="rgba(255,255,255,0.1)" strokeWidth={1} />
                          ))}
                        </Pie>
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>
                )}
              </div>

              {/* Line Chart */}
              {timeSeriesData.length > 0 && (
                <div id="chart-performance-trend" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 h-[350px] flex flex-col group relative">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Performance Trend</h3>
                    <button onClick={() => handleDownloadSection('chart-performance-trend', 'Performance_Trend')} className="p-1.5 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg text-slate-400 hover:text-white transition-colors" title="Download Chart">
                      <Download className="w-4 h-4" />
                    </button>
                  </div>
                  <div className="flex-1 min-h-0 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={timeSeriesData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                        <XAxis dataKey="date" stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} />
                        <YAxis stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(value) => `$${value/1000}k`} />
                        <RechartsTooltip 
                          cursor={{ stroke: '#ffffff20' }}
                          contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #1e293b', borderRadius: '8px', color: '#f8fafc' }}
                          itemStyle={{ fontSize: '12px' }}
                        />
                        <Legend wrapperStyle={{ fontSize: '10px', paddingTop: '10px' }} />
                        {appliedMappings.sales && <Line type="monotone" dataKey="sales" name="Sales ($)" stroke="#3b82f6" strokeWidth={2} dot={{ fill: '#3b82f6', r: 3 }} activeDot={{ r: 5 }} />}
                        {appliedMappings.cogs && <Line type="monotone" dataKey="cogs" name="COGS ($)" stroke="#ef4444" strokeWidth={2} dot={{ fill: '#ef4444', r: 3 }} activeDot={{ r: 5 }} />}
                        {appliedMappings.grossIncome && <Line type="monotone" dataKey="grossIncome" name="Gross Income ($)" stroke="#10b981" strokeWidth={2} dot={{ fill: '#10b981', r: 3 }} activeDot={{ r: 5 }} />}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              )}

              {/* Dynamic Row for Gender & Location Data */}
              {(genderData.length > 0 || locationData.length > 0) && (
                <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                  {/* Gender Stacked Bar Chart */}
                  {genderData.length > 0 && (
                    <div id="chart-gender-distribution" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 h-[350px] flex flex-col group relative">
                      <div className="flex justify-between items-center mb-4">
                        <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Gender Distribution by Category</h3>
                        <button onClick={() => handleDownloadSection('chart-gender-distribution', 'Gender_Distribution')} className="p-1.5 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg text-slate-400 hover:text-white transition-colors" title="Download Chart">
                          <Download className="w-4 h-4" />
                        </button>
                      </div>
                      <div className="flex-1 min-h-0 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={genderData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                            <CartesianGrid strokeDasharray="3 3" stroke="#ffffff20" vertical={false} />
                            <XAxis dataKey="productLine" stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} />
                            <YAxis stroke="#94a3b8" fontSize={10} tickLine={false} axisLine={false} />
                            <RechartsTooltip 
                              cursor={{ fill: '#ffffff10' }}
                              contentStyle={{ backgroundColor: '#0f172a', border: '1px solid #1e293b', borderRadius: '8px', color: '#f8fafc' }}
                              itemStyle={{ fontSize: '12px' }}
                            />
                            <Legend wrapperStyle={{ fontSize: '10px', paddingTop: '10px' }} />
                            <Bar dataKey="Male" stackId="a" fill="#3b82f6" name="Male" />
                            <Bar dataKey="Female" stackId="a" fill="#ec4899" name="Female" />
                            <Bar dataKey="Other" stackId="a" fill="#94a3b8" name="Other" radius={[4, 4, 0, 0]} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  )}

                  {/* Location Performance List */}
                  {locationData.length > 0 && (
                    <div id="chart-performance-location" className="bg-white/5 backdrop-blur-xl border border-white/10 rounded-2xl p-6 h-[350px] flex flex-col group relative">
                      <div className="flex justify-between items-center mb-4">
                        <h3 className="text-xs font-bold text-slate-300 uppercase tracking-widest">Performance by Location</h3>
                        <button onClick={() => handleDownloadSection('chart-performance-location', 'Performance_by_Location')} className="p-1.5 bg-white/5 hover:bg-white/10 border border-white/10 rounded-lg text-slate-400 hover:text-white transition-colors" title="Download Table">
                          <Download className="w-4 h-4" />
                        </button>
                      </div>
                      <div className="flex-1 w-full overflow-y-auto pr-2 custom-scrollbar">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="sticky top-0 bg-[#0f172a] z-10 shadow-sm shadow-[#020617]">
                              <th className="py-3 px-2 text-[10px] font-bold text-slate-400 uppercase tracking-wider border-b border-white/10">Location</th>
                              {appliedMappings.sales && <th className="py-3 px-2 text-[10px] font-bold text-slate-400 uppercase tracking-wider text-right border-b border-white/10">Total Sales</th>}
                              {appliedMappings.grossIncome && <th className="py-3 px-2 text-[10px] font-bold text-slate-400 uppercase tracking-wider text-right border-b border-white/10">Gross Income</th>}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-white/5">
                            {locationData.map((loc, idx) => (
                              <tr key={idx} className="hover:bg-white/5 transition-colors">
                                <td className="py-3 px-2 text-sm text-slate-200 font-medium">{loc.location}</td>
                                {appliedMappings.sales && <td className="py-3 px-2 text-sm text-emerald-400 text-right">${loc.sales.toFixed(2)}</td>}
                                {appliedMappings.grossIncome && <td className="py-3 px-2 text-sm text-slate-300 text-right">${loc.grossIncome.toFixed(2)}</td>}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* AI Analysis Result */}
          {aiResult && (
            <div id="chart-ai-insights" className="shrink-0 bg-gradient-to-br from-emerald-500/10 to-blue-600/10 backdrop-blur-2xl border border-emerald-500/20 rounded-2xl p-6 overflow-hidden relative flex flex-col min-h-[300px] group">
              <div className="absolute top-0 right-0 p-4 opacity-10 pointer-events-none">
                <Bot className="w-32 h-32 text-white" />
              </div>
              <div className="flex justify-between items-center mb-4 shrink-0">
                <h3 className="text-xs font-bold text-emerald-400 uppercase tracking-widest flex items-center">
                  <span className="mr-2">●</span> AI CFO INSIGHTS
                </h3>
                <button onClick={() => handleDownloadSection('chart-ai-insights', 'AI_CFO_Insights')} className="p-1.5 bg-emerald-500/10 hover:bg-emerald-500/20 border border-emerald-500/20 rounded-lg text-emerald-400 hover:text-white transition-colors z-10" title="Download Insights">
                  <Download className="w-4 h-4" />
                </button>
              </div>
              <div className="text-sm leading-relaxed overflow-y-auto pr-2 flex-1 prose prose-invert prose-emerald max-w-none prose-p:text-slate-200 prose-strong:text-white prose-li:text-slate-200 hide-scrollbar">
                <ReactMarkdown>{aiResult}</ReactMarkdown>
                
                {/* Chat History */}
                {chatHistory.length > 0 && (
                  <div className="mt-6 space-y-4 border-t border-emerald-500/20 pt-6">
                    {chatHistory.map((msg, idx) => (
                      <div key={idx} className={`flex flex-col ${msg.role === 'user' ? 'items-end' : 'items-start'}`}>
                        <div className={`text-[10px] uppercase font-bold mb-1 ${msg.role === 'user' ? 'text-blue-400' : 'text-emerald-400'}`}>
                          {msg.role === 'user' ? 'You' : 'AI CFO'}
                        </div>
                        <div className={`px-4 py-3 rounded-2xl max-w-[85%] text-sm ${msg.role === 'user' ? 'bg-blue-600/20 border border-blue-500/30 text-blue-50' : 'bg-emerald-900/40 border border-emerald-500/20 text-slate-200'}`}>
                          <ReactMarkdown>{msg.content}</ReactMarkdown>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Chat Input */}
              <div className="mt-4 pt-4 border-t border-emerald-500/20 flex gap-2 shrink-0 z-10 relative">
                <input
                  type="text"
                  value={chatQuestion}
                  onChange={e => setChatQuestion(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && handleAskQuestion()}
                  placeholder="Ask a follow-up question (e.g. What about Q2 projection?)..."
                  className="flex-1 bg-black/40 border border-emerald-500/30 rounded-xl px-4 py-2.5 text-sm text-slate-200 placeholder:text-slate-500 max-w-full focus:outline-none focus:border-emerald-500/60"
                />
                <button
                  onClick={handleAskQuestion}
                  disabled={chatLoading || !chatQuestion.trim()}
                  className="bg-emerald-600 text-slate-50 px-5 py-2.5 rounded-xl font-bold text-sm hover:bg-emerald-500 transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center shrink-0 border border-emerald-500/40 hover:border-emerald-400"
                >
                  {chatLoading ? <Loader2 className="w-4 h-4 animate-spin" /> : <Send className="w-4 h-4" />}
                </button>
              </div>
            </div>
          )}
          </div>
        </section>
      </main>

      {/* Footer Status */}
      <footer className="px-8 py-3 border-t border-white/5 bg-black/20 backdrop-blur-sm flex justify-between text-[10px] text-slate-500 font-mono relative z-10 shrink-0">
        <div>STATUS: {loadingAi ? 'ANALYZING...' : (aiResult ? 'ANALYZED' : 'READY FOR EXPORT')}</div>
        <div>ENGINE: GOOGLE_GEMINI_2.5_FLASH</div>
        <div>SYSTEM: ONLINE</div>
      </footer>

      {/* Help Modal */}
      {showHelpModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
          <div className="bg-[#0f172a] border border-white/10 rounded-2xl w-full max-w-lg overflow-hidden shadow-2xl relative flex flex-col">
            <div className="px-6 py-4 border-b border-white/10 flex justify-between items-center bg-white/5">
              <h3 className="text-lg font-semibold text-slate-200">File Format Instructions</h3>
              <button 
                onClick={() => setShowHelpModal(false)}
                className="text-slate-400 hover:text-white transition-colors p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 text-sm text-slate-300 space-y-4">
              <p>For the AI CFO Analyzer to work correctly, your uploaded CSV or XLSX file should contain the following columns:</p>
              <ul className="space-y-3 list-disc pl-5">
                <li>
                  <strong className="text-emerald-400">Product Line</strong>: The category or line of the product (e.g., "Electronic accessories"). 
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "product line"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">Sales</strong> (or Total): The total sales amount for the transaction. 
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "sales", "total"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">COGS</strong> (Cost of Goods Sold): The cost associated with the transaction. 
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "cogs", "cost"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">Gross Income</strong>: The profit generated from the transaction. 
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "gross income", "gross"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">Date</strong> (Optional): Used to generate a Gross Income trend over time.
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "date", "time"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">Gender</strong> (Optional): Analyzes the gender ratio by product line.
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "gender", "sex"</span>
                </li>
                <li>
                  <strong className="text-emerald-400">Location</strong> (Optional): Analyzes sales and profit by location/branch.
                  <br/><span className="text-slate-500 text-xs">Keywords matched: "location", "city", "branch"</span>
                </li>
              </ul>
              <div className="p-3 bg-blue-500/10 border border-blue-500/20 rounded-xl mt-4 text-blue-300 text-xs leading-relaxed">
                <strong className="block mb-1 text-blue-400">💡 Pro Tip:</strong> 
                The system is case-insensitive and ignores spaces when trying to match column headers. Please ensure numeric columns contain valid numbers.
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
