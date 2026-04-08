import React, { useState, useCallback, useRef, useEffect } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload, FileText, FileSpreadsheet, Loader2, Save, CheckCircle2, AlertCircle, Download, Clock, Calendar, Users, History, Trash2, Settings, X, GripVertical } from 'lucide-react';
import { useAuth } from './AuthContext';
import { db } from '../firebase';
import { collection, addDoc } from 'firebase/firestore';
import { GoogleGenAI, Type } from '@google/genai';
import * as XLSX from 'xlsx';
import Tesseract from 'tesseract.js';
import { DragDropContext, Droppable, Draggable, DropResult } from '@hello-pangea/dnd';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

interface TimecardEntry {
  date: string;
  [key: string]: string; // Dynamic time columns
}

interface TimecardData {
  employeeName: string;
  employeeId?: string;
  period: string;
  totalHours: string;
  entries: TimecardEntry[];
  payslip?: Record<string, string>;
}

interface ExtractionSession {
  id: string;
  extractedAt: string;
  fileNames: string[];
  timecards: TimecardData[];
  extractionTimeMs: number;
  employeeNames: string[];
  yearsCovered: string[];
  saved?: boolean;
  payslipColumns?: string[];
}

type ExtractionType = 'both' | 'timecard' | 'payslip';

interface UploadedFile {
  id: string;
  file: File;
  selected: boolean;
  progress?: number;
  status?: 'pending' | 'processing' | 'success' | 'error';
  error?: string;
  ocrText?: string;
}

interface PreAnalysisData {
  timeColumns: string[];
  payslipMappings: { targetName: string; sourceName: string; }[];
}

interface FieldMapping {
  targetName: string;
  sourceName: string;
}

export default function Dashboard() {
  const { user } = useAuth();
  
  // File State
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  
  // Configuration State
  const [extractionType, setExtractionType] = useState<ExtractionType>('both');
  
  // Pre-analysis State
  const [isPreAnalyzing, setIsPreAnalyzing] = useState(false);
  const [showPreAnalysisModal, setShowPreAnalysisModal] = useState(false);
  const [preAnalysisData, setPreAnalysisData] = useState<PreAnalysisData | null>(null);
  const [selectedTimeColumns, setSelectedTimeColumns] = useState<string[]>([]);
  const [fieldMappings, setFieldMappings] = useState<FieldMapping[]>([]);
  
  // Extraction State
  const [isExtracting, setIsExtracting] = useState(false);
  const [elapsedTime, setElapsedTime] = useState(0);
  const [progress, setProgress] = useState(0);
  const timerRef = useRef<NodeJS.Timeout | null>(null);
  
  // History & Error State
  const [history, setHistory] = useState<ExtractionSession[]>([]);
  const [savingId, setSavingId] = useState<string | null>(null);
  const [errorModal, setErrorModal] = useState<{ isOpen: boolean; title: string; message: string }>({ isOpen: false, title: '', message: '' });

  // Cleanup timer on unmount
  useEffect(() => {
    return () => {
      if (timerRef.current) clearInterval(timerRef.current);
    };
  }, []);

  const showError = (title: string, message: string) => {
    setErrorModal({ isOpen: true, title, message });
  };

  const onDropFiles = useCallback(async (acceptedFiles: File[]) => {
    const newFiles: UploadedFile[] = acceptedFiles.map(file => ({
      id: Math.random().toString(36).substring(7),
      file,
      selected: true,
      status: file.type.startsWith('image/') ? 'processing' : 'success',
      progress: file.type.startsWith('image/') ? 0 : 100
    }));
    
    setUploadedFiles(prev => [...prev, ...newFiles]);

    // Process OCR for image files
    for (const newFile of newFiles) {
      if (newFile.file.type.startsWith('image/')) {
        try {
          const worker = await Tesseract.createWorker('por', 1, {
            logger: m => {
              if (m.status === 'recognizing text') {
                setUploadedFiles(prev => prev.map(f => 
                  f.id === newFile.id ? { ...f, progress: Math.round(m.progress * 100) } : f
                ));
              }
            }
          });
          
          const { data: { text } } = await worker.recognize(newFile.file);
          await worker.terminate();
          
          setUploadedFiles(prev => prev.map(f => 
            f.id === newFile.id ? { ...f, status: 'success', progress: 100, ocrText: text } : f
          ));
        } catch (error: any) {
          console.error("OCR Error:", error);
          setUploadedFiles(prev => prev.map(f => 
            f.id === newFile.id ? { ...f, status: 'error', error: 'Falha no OCR: ' + (error.message || 'Erro desconhecido') } : f
          ));
        }
      }
    }
  }, []);

  const onDropTemplate = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      setTemplateFile(acceptedFiles[0]);
    }
  }, []);

  const { getRootProps: getFilesProps, getInputProps: getFilesInputProps, isDragActive: isFilesDrag } = useDropzone({
    onDrop: onDropFiles,
    accept: {
      'image/*': ['.png', '.jpg', '.jpeg'],
      'application/pdf': ['.pdf']
    }
  });

  const { getRootProps: getTemplateProps, getInputProps: getTemplateInputProps, isDragActive: isTemplateDrag } = useDropzone({
    onDrop: onDropTemplate,
    accept: {
      'application/pdf': ['.pdf'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    maxFiles: 1
  });

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => {
        if (typeof reader.result === 'string') {
          const base64 = reader.result.split(',')[1];
          resolve(base64);
        } else {
          reject(new Error("Falha ao converter o arquivo."));
        }
      };
      reader.onerror = error => reject(error);
    });
  };

  const toggleFileSelection = (id: string) => {
    setUploadedFiles(prev => prev.map(f => f.id === id ? { ...f, selected: !f.selected } : f));
  };

  const removeFile = (id: string) => {
    setUploadedFiles(prev => prev.filter(f => f.id !== id));
  };

  const toggleAllFiles = (select: boolean) => {
    setUploadedFiles(prev => prev.map(f => ({ ...f, selected: select })));
  };

  const handlePreAnalyze = async () => {
    const selectedFiles = uploadedFiles.filter(f => f.selected);
    if (selectedFiles.length === 0) {
      showError("Nenhum arquivo selecionado", "Por favor, selecione pelo menos um arquivo para extração.");
      return;
    }

    setIsPreAnalyzing(true);
    
    try {
      const firstFileObj = selectedFiles[0];
      const firstFile = firstFileObj.file;
      const fileBase64 = await fileToBase64(firstFile);
      
      const parts: any[] = [
        { inlineData: { data: fileBase64, mimeType: firstFile.type } }
      ];

      if (firstFileObj.ocrText) {
        parts.push({ text: `\nTexto extraído via OCR do Documento 1:\n${firstFileObj.ocrText}\n` });
      }

      let prompt = `Analise os documentos fornecidos.\nO Documento 1 (fornecido acima) é o arquivo principal (cartão de ponto/holerite).\n`;

      if (templateFile) {
        if (templateFile.type === 'application/pdf') {
          const templateBase64 = await fileToBase64(templateFile);
          parts.push({
            inlineData: { data: templateBase64, mimeType: templateFile.type }
          });
        } else {
          const arrayBuffer = await templateFile.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          const csvData = XLSX.utils.sheet_to_csv(workbook.Sheets[workbook.SheetNames[0]]);
          parts.push({ text: `Documento 2 (Modelo em CSV):\n${csvData}` });
        }
        
        prompt += `O Documento 2 é o modelo de formatação desejado.
        
TAREFAS:
1. Identifique as colunas de horários de ponto no Documento 1 (ex: Entrada, Saída).
2. Identifique as verbas/campos de holerite presentes no Documento 2 (Modelo).
3. Para cada verba encontrada no Documento 2, procure a verba correspondente no Documento 1.

REGRA GERAL PARA SALÁRIO: Sempre que o modelo pedir "salário" ou "salário hora", o valor correto no Documento 1 é o SALÁRIO HORA que aparece no cabeçalho seguido de "/H" ou "/ H" (exemplo: "4,7778/ H"). Mapeie o "sourceName" para o nome exato do campo que contém esse valor (ex: "Salário:" ou "Salário Hora").

Retorne um JSON estrito com:
- timeColumns: array de strings com os nomes das colunas de horário do Documento 1.
- payslipMappings: array de objetos contendo:
  - targetName: o nome da verba como aparece no Documento 2 (Modelo).
  - sourceName: o nome da mesma verba como aparece no Documento 1. Se não encontrar correspondência no Documento 1, retorne uma string vazia "".`;
      } else {
        prompt += `TAREFAS:
1. Identifique as colunas de horários de ponto no Documento 1 (ex: Entrada, Saída).
2. Identifique as verbas/campos de holerite presentes no Documento 1.

REGRA GERAL PARA SALÁRIO: O salário do funcionário é o SALÁRIO HORA que aparece no cabeçalho do documento seguido de "/H" ou "/ H" (exemplo: "4,7778/ H"). Certifique-se de incluir esse campo nos mapeamentos.

Retorne um JSON estrito com:
- timeColumns: array de strings com os nomes das colunas de horário.
- payslipMappings: array de objetos contendo:
  - targetName: o nome da verba encontrada.
  - sourceName: o mesmo nome da verba encontrada.`;
      }

      parts.push({ text: prompt });

      const response = await ai.models.generateContent({
        model: 'gemini-3.1-pro-preview',
        contents: { parts },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              timeColumns: { type: Type.ARRAY, items: { type: Type.STRING } },
              payslipMappings: { 
                type: Type.ARRAY, 
                items: { 
                  type: Type.OBJECT, 
                  properties: {
                    targetName: { type: Type.STRING },
                    sourceName: { type: Type.STRING }
                  },
                  required: ["targetName", "sourceName"]
                } 
              }
            },
            required: ["timeColumns", "payslipMappings"]
          }
        }
      });

      const jsonStr = response.text?.trim();
      if (jsonStr) {
        const data: PreAnalysisData = JSON.parse(jsonStr);
        
        // Make column names unique to prevent React key errors and JSON schema overwrites
        const makeUnique = (arr: string[]) => {
          const counts: Record<string, number> = {};
          arr.forEach(item => { counts[item] = (counts[item] || 0) + 1; });
          const seen: Record<string, number> = {};
          return arr.map(item => {
            if (counts[item] > 1) {
              seen[item] = (seen[item] || 0) + 1;
              return `${item} ${seen[item]}`;
            }
            return item;
          });
        };

        const uniqueTimeColumns = makeUnique(data.timeColumns || []);
        
        const payslipTargetCounts: Record<string, number> = {};
        const rawMappings = data.payslipMappings || [];
        rawMappings.forEach(m => { payslipTargetCounts[m.targetName] = (payslipTargetCounts[m.targetName] || 0) + 1; });
        const payslipSeen: Record<string, number> = {};
        
        const uniquePayslipMappings = rawMappings.map(m => {
          let newTarget = m.targetName;
          if (payslipTargetCounts[m.targetName] > 1) {
            payslipSeen[m.targetName] = (payslipSeen[m.targetName] || 0) + 1;
            newTarget = `${m.targetName} ${payslipSeen[m.targetName]}`;
          }
          return { targetName: newTarget, sourceName: m.sourceName };
        });

        data.timeColumns = uniqueTimeColumns;
        data.payslipMappings = uniquePayslipMappings;

        setPreAnalysisData(data);
        setSelectedTimeColumns(uniqueTimeColumns);
        setFieldMappings(uniquePayslipMappings);
        setShowPreAnalysisModal(true);
      } else {
        throw new Error("Não foi possível analisar o documento.");
      }
    } catch (err: any) {
      console.error("Pre-analysis error:", err);
      showError("Erro na Análise Prévia", err.message || "Falha ao analisar o documento. Verifique se o arquivo é válido e tente novamente.");
    } finally {
      setIsPreAnalyzing(false);
    }
  };

  const handleExtract = async () => {
    setShowPreAnalysisModal(false);
    setIsExtracting(true);
    setElapsedTime(0);
    setProgress(0);
    
    const startTime = Date.now();
    const selectedFiles = uploadedFiles.filter(f => f.selected);
    
    // Start progress timer
    timerRef.current = setInterval(() => {
      const currentElapsed = Date.now() - startTime;
      setElapsedTime(currentElapsed);
      // Estimate 20s per file
      const estimatedTotalTime = selectedFiles.length * 20000;
      setProgress(Math.min(95, (currentElapsed / estimatedTotalTime) * 100));
    }, 100);
    
    try {
      let allExtractedData: TimecardData[] = [];

      for (const fileObj of selectedFiles) {
        const fileBase64 = await fileToBase64(fileObj.file);
        
        const parts: any[] = [
          {
            inlineData: {
              data: fileBase64,
              mimeType: fileObj.file.type
            }
          }
        ];

        if (fileObj.ocrText) {
          parts.push({ text: `\nTexto extraído via OCR do Documento 1:\n${fileObj.ocrText}\n` });
        }

        let prompt = `Extraia os dados do documento fornecido para TODOS os funcionários.
O usuário solicitou a extração do tipo: ${extractionType === 'both' ? 'Cartão de Ponto e Holerite' : extractionType === 'timecard' ? 'Apenas Cartão de Ponto' : 'Apenas Holerite'}.

REGRAS IMPORTANTES:
1. O idioma de saída de todos os textos deve ser Português do Brasil (pt-BR).
2. FORMATO DOS NÚMEROS/HORÁRIOS: Modifique os dois-pontos ":" para vírgula "," (exemplo: 6:52 vira 6,50). Use vírgula para decimais em valores monetários.
`;

        if (extractionType === 'timecard' || extractionType === 'both') {
          prompt += `\nREGRAS PARA CARTÃO DE PONTO:
- INCLUA TODOS OS DIAS DO MÊS/PERÍODO: Não pule nenhum dia. Mesmo dias em branco, faltas, feriados, férias, atestados, DSR, etc., devem constar na lista. Se for falta/feriado/etc, coloque essa informação nos campos de horário.
- FORMATO DA DATA: Use o formato "DD/MM/YYYY ddd" (exemplo: "01/02/2018 qui"). Use as abreviações dos dias da semana em minúsculo e em português.
- COLUNAS SOLICITADAS: Extraia APENAS as seguintes colunas de horário: ${selectedTimeColumns.join(', ')}. Crie as chaves no objeto JSON exatamente com esses nomes.
- Inclua também a coluna "hours" (horas trabalhadas no dia).
`;
        }

        if (extractionType === 'payslip' || extractionType === 'both') {
          const mappingInstructions = fieldMappings
            .filter(m => m.sourceName.trim() !== '')
            .map(m => {
              const targetLower = m.targetName.toLowerCase().trim();
              const sourceLower = m.sourceName.toLowerCase().trim();
              const isSalary = ['salário', 'salario', 'salário hora', 'salario hora'].includes(targetLower) || 
                               ['salário', 'salario', 'salário hora', 'salario hora'].includes(sourceLower);
              const extraInstruction = isSalary ? ' (ATENÇÃO: O salário é o SALÁRIO HORA. Procure no cabeçalho pelo valor numérico seguido de "/H" ou "/ H", ex: "4,7778/ H")' : '';
              return `- Procure por "${m.sourceName}" e extraia seu valor para a chave "${m.targetName}"${extraInstruction}`;
            })
            .join('\n');
          prompt += `\nREGRAS PARA HOLERITE:
- Extraia os valores das seguintes verbas usando o mapeamento abaixo:
${mappingInstructions}
- Se uma verba não for encontrada, deixe o valor vazio.
- REGRA GERAL PARA SALÁRIO: Apenas para o campo estritamente chamado "salário" ou "salário hora", o valor correto é o SALÁRIO HORA que aparece no cabeçalho do documento seguido de "/H" ou "/ H" (exemplo: "4,7778/ H"). Para "salário família", "salário total", etc., extraia o valor normal da tabela.
`;
        }

        if (templateFile) {
          if (templateFile.type === 'application/pdf') {
            const templateBase64 = await fileToBase64(templateFile);
            parts.push({
              inlineData: {
                data: templateBase64,
                mimeType: templateFile.type
              }
            });
          } else if (
            templateFile.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
            templateFile.type === 'application/vnd.ms-excel' ||
            templateFile.name.endsWith('.xlsx') ||
            templateFile.name.endsWith('.xls')
          ) {
            const arrayBuffer = await templateFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const csvData = XLSX.utils.sheet_to_csv(worksheet);
            
            parts.push({
              text: `Aqui está o modelo em formato CSV:\n${csvData}`
            });
          }
          prompt += "\nO segundo documento é um modelo Excel. Use-o APENAS para entender o formato de saída desejado.";
        }

        parts.push({ text: prompt });

        // Define dynamic schema based on selections
        const entryProperties: Record<string, any> = {
          date: { type: Type.STRING, description: "Data no formato DD/MM/YYYY ddd" },
          hours: { type: Type.STRING, description: "Horas trabalhadas no dia (com vírgula)" }
        };
        
        if (extractionType === 'timecard' || extractionType === 'both') {
          selectedTimeColumns.forEach(col => {
            entryProperties[col] = { type: Type.STRING, description: `Valor para a coluna ${col}` };
          });
        }

        const timecardProperties: Record<string, any> = {
          employeeName: { type: Type.STRING, description: "Nome do funcionário" },
          employeeId: { type: Type.STRING, description: "Matrícula do funcionário, se houver" },
          period: { type: Type.STRING, description: "Período de apuração ou mês/ano" },
          totalHours: { type: Type.STRING, description: "Total de horas trabalhadas no período (com vírgula)" },
        };

        if (extractionType === 'timecard' || extractionType === 'both') {
          timecardProperties.entries = {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: entryProperties,
              required: ["date", "hours", ...selectedTimeColumns]
            }
          };
        } else {
          timecardProperties.entries = { type: Type.ARRAY, items: { type: Type.OBJECT, properties: {} } }; // Empty if not requested
        }

        if (extractionType === 'payslip' || extractionType === 'both') {
          const payslipProps: Record<string, any> = {};
          fieldMappings.forEach(m => {
            if (m.targetName.trim() !== '') {
              payslipProps[m.targetName] = { type: Type.STRING, description: `Valor extraído de ${m.sourceName || m.targetName}` };
            }
          });
          timecardProperties.payslip = {
            type: Type.OBJECT,
            properties: payslipProps
          };
        }

        const response = await ai.models.generateContent({
          model: 'gemini-3.1-pro-preview',
          contents: { parts },
          config: {
            responseMimeType: 'application/json',
            responseSchema: {
              type: Type.ARRAY,
              items: {
                type: Type.OBJECT,
                properties: timecardProperties,
                required: ["employeeName", "period"]
              }
            }
          }
        });

        const jsonStr = response.text?.trim();
        if (jsonStr) {
          const data: TimecardData[] = JSON.parse(jsonStr);
          
          // Validation and sanitization
          data.forEach(tc => {
            if (tc.totalHours) {
              tc.totalHours = tc.totalHours.replace('.', ',');
            }
            tc.entries?.forEach(entry => {
              if (entry.hours) {
                entry.hours = entry.hours.replace('.', ',');
              }
              // Ensure time columns use comma
              selectedTimeColumns.forEach(col => {
                if (entry[col]) {
                  entry[col] = entry[col].replace('.', ',');
                }
              });
            });
            if (tc.payslip) {
              Object.keys(tc.payslip).forEach(k => {
                if (tc.payslip![k]) {
                  // If it looks like a monetary value or number, ensure it uses comma for decimal
                  if (/^\d+\.\d{2,4}$/.test(tc.payslip![k])) {
                    tc.payslip![k] = tc.payslip![k].replace('.', ',');
                  }
                }
              });
            }
          });

          allExtractedData = [...allExtractedData, ...data];
        } else {
          throw new Error(`Nenhum dado extraído do arquivo ${fileObj.file.name}.`);
        }
      }
      
      const endTime = Date.now();
      const totalTimeMs = endTime - startTime;
      setProgress(100);
      
      // Process data for summary
      const employeeNames = Array.from(new Set(allExtractedData.map(tc => tc.employeeName).filter(Boolean)));
      const yearsSet = new Set<string>();
      
      allExtractedData.forEach(tc => {
        const periodMatch = tc.period?.match(/\b(20\d{2})\b/);
        if (periodMatch) yearsSet.add(periodMatch[1]);
        
        tc.entries?.forEach(entry => {
          const dateMatch = entry.date?.match(/\b(20\d{2})\b/);
          if (dateMatch) yearsSet.add(dateMatch[1]);
        });
      });

      const newSession: ExtractionSession = {
        id: Date.now().toString() + Math.random().toString(36).substring(7),
        extractedAt: new Date().toISOString(),
        fileNames: selectedFiles.map(f => f.file.name),
        timecards: allExtractedData,
        extractionTimeMs: totalTimeMs,
        employeeNames,
        yearsCovered: Array.from(yearsSet).sort(),
        payslipColumns: fieldMappings.map(m => m.targetName).filter(name => name.trim() !== '')
      };

      setHistory(prev => [newSession, ...prev]);
      
    } catch (err: any) {
      console.error("Extraction error:", err);
      showError("Erro na Extração", err.message || "Falha ao extrair dados. Tente novamente.");
    } finally {
      if (timerRef.current) clearInterval(timerRef.current);
      setIsExtracting(false);
    }
  };

  const handleSave = async (session: ExtractionSession) => {
    if (!user) return;
    
    setSavingId(session.id);
    
    try {
      for (const timecard of session.timecards) {
        const timecardData = {
          ...timecard,
          uid: user.uid,
          createdAt: new Date().toISOString()
        };
        await addDoc(collection(db, 'users', user.uid, 'timecards'), timecardData);
      }
      
      setHistory(prev => prev.map(s => s.id === session.id ? { ...s, saved: true } : s));
    } catch (err: any) {
      console.error("Save error:", err);
      showError("Erro ao Salvar", err.message || "Falha ao salvar dados no banco. Tente novamente.");
    } finally {
      setSavingId(null);
    }
  };

  const exportToExcel = (session: ExtractionSession) => {
    const wb = XLSX.utils.book_new();
    const wsData: any[][] = [];
    
    const setCell = (r: number, c: number, val: any) => {
      while (wsData.length <= r) wsData.push([]);
      while (wsData[r].length <= c) wsData[r].push("");
      wsData[r][c] = val;
    };

    // Group by employee
    const employees = Array.from(new Set(session.timecards.map(t => t.employeeName)));
    
    // Determine columns from session data
    const timeColsSet = new Set<string>();
    
    // Initialize payslipCols with the exact order from the session (or template if missing)
    const payslipCols: string[] = session.payslipColumns ? [...session.payslipColumns] : fieldMappings.map(m => m.targetName).filter(name => name.trim() !== '');
    
    session.timecards.forEach(tc => {
      tc.entries?.forEach(entry => {
        Object.keys(entry).forEach(k => {
          if (k !== 'date' && k !== 'hours') timeColsSet.add(k);
        });
      });
      if (tc.payslip) {
        Object.keys(tc.payslip).forEach(k => {
          if (!payslipCols.includes(k)) {
            payslipCols.push(k);
          }
        });
      }
    });
    
    const timeCols = Array.from(timeColsSet);
    
    const payslipStartCol = timeCols.length + 3; // Data, ...timeCols, Horas, [Empty]
    
    let currentRow = 0;

    employees.forEach(empName => {
      const empTimecards = session.timecards.filter(t => t.employeeName === empName);
      
      // Employee Header
      setCell(currentRow, 0, "Funcionário:");
      setCell(currentRow, 1, empName);
      
      // Find Employee ID if available
      const empId = empTimecards.find(t => t.employeeId)?.employeeId;
      if (empId) {
        setCell(currentRow, 2, "Matrícula:");
        setCell(currentRow, 3, empId);
      }
      
      currentRow += 2;
      const sectionStartRow = currentRow;
      
      // --- RIGHT SIDE: PAYSLIPS ---
      let psRow = sectionStartRow;
      if (payslipCols.length > 0) {
        // Headers
        setCell(psRow, payslipStartCol, "Período");
        payslipCols.forEach((col, i) => {
          setCell(psRow, payslipStartCol + 1 + i, col);
        });
        psRow++;
        
        // Data
        empTimecards.forEach(tc => {
          if (tc.payslip && Object.keys(tc.payslip).length > 0) {
            setCell(psRow, payslipStartCol, tc.period);
            payslipCols.forEach((col, i) => {
              setCell(psRow, payslipStartCol + 1 + i, tc.payslip![col] || "");
            });
            psRow++;
          }
        });
      }
      
      // --- LEFT SIDE: TIMECARDS ---
      let tcRow = sectionStartRow;
      empTimecards.forEach(tc => {
        if (tc.entries && tc.entries.length > 0) {
          setCell(tcRow, 0, "Período:");
          setCell(tcRow, 1, tc.period);
          if (tc.totalHours) {
            setCell(tcRow, 2, "Total de Horas:");
            setCell(tcRow, 3, tc.totalHours);
          }
          tcRow++;
          
          // Headers
          setCell(tcRow, 0, "Data");
          timeCols.forEach((col, i) => setCell(tcRow, i + 1, col));
          setCell(tcRow, timeCols.length + 1, "Horas");
          tcRow++;
          
          // Entries
          tc.entries.forEach(entry => {
            setCell(tcRow, 0, entry.date || "");
            timeCols.forEach((col, i) => setCell(tcRow, i + 1, entry[col] || ""));
            setCell(tcRow, timeCols.length + 1, entry.hours || "");
            tcRow++;
          });
          
          tcRow += 1; // Space between timecards
        }
      });
      
      // Move currentRow to the maximum of tcRow and psRow for the next employee
      currentRow = Math.max(tcRow, psRow) + 2;
    });

    // Fill missing cells to make it a perfect rectangle
    let maxCols = 0;
    wsData.forEach(row => {
      if (row.length > maxCols) maxCols = row.length;
    });
    
    for (let r = 0; r < wsData.length; r++) {
      while (wsData[r].length < maxCols) {
        wsData[r].push("");
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, "Extração");

    const fileName = session.fileNames.length === 1 
      ? `${session.fileNames[0].replace(/\.[^/.]+$/, "")}_extraido.xlsx`
      : `Extracao_Multipla_${new Date().getTime()}.xlsx`;

    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="max-w-6xl mx-auto p-6">
      <div className="mb-8">
        <h1 className="text-3xl font-bold text-gray-900">Extrator de Cartão de Ponto e Holerite</h1>
        <p className="text-gray-600 mt-2">
          Envie seus documentos em PDF ou Imagem. O sistema analisará os arquivos e permitirá que você escolha quais dados extrair.
        </p>
      </div>

      <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 mb-8">
        <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
          <Settings className="w-5 h-5 text-gray-700" />
          Tipo de Extração
        </h2>
        <div className="flex flex-wrap gap-4">
          <label className="flex items-center gap-2 cursor-pointer p-3 border rounded-lg hover:bg-gray-50 transition-colors">
            <input 
              type="radio" 
              name="extractionType" 
              value="both" 
              checked={extractionType === 'both'} 
              onChange={() => setExtractionType('both')}
              className="w-4 h-4 text-indigo-600"
            />
            <span className="font-medium text-gray-700">Cartão de Ponto + Holerite</span>
          </label>
          <label className="flex items-center gap-2 cursor-pointer p-3 border rounded-lg hover:bg-gray-50 transition-colors">
            <input 
              type="radio" 
              name="extractionType" 
              value="timecard" 
              checked={extractionType === 'timecard'} 
              onChange={() => setExtractionType('timecard')}
              className="w-4 h-4 text-indigo-600"
            />
            <span className="font-medium text-gray-700">Apenas Cartão de Ponto</span>
          </label>
          <label className="flex items-center gap-2 cursor-pointer p-3 border rounded-lg hover:bg-gray-50 transition-colors">
            <input 
              type="radio" 
              name="extractionType" 
              value="payslip" 
              checked={extractionType === 'payslip'} 
              onChange={() => setExtractionType('payslip')}
              className="w-4 h-4 text-indigo-600"
            />
            <span className="font-medium text-gray-700">Apenas Holerite</span>
          </label>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-8">
        {/* Files Upload */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 flex flex-col">
          <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
            <FileText className="w-5 h-5 text-indigo-600" />
            Documentos (Obrigatório)
          </h2>
          <div 
            {...getFilesProps()} 
            className={`border-2 border-dashed rounded-lg p-6 text-center cursor-pointer transition-colors mb-4 ${
              isFilesDrag ? 'border-indigo-500 bg-indigo-50' : 'border-gray-300 hover:border-indigo-400 hover:bg-gray-50'
            }`}
          >
            <input {...getFilesInputProps()} />
            <Upload className="w-8 h-8 mx-auto text-gray-400 mb-2" />
            <p className="text-sm text-gray-600">Arraste e solte arquivos PDF/Imagens aqui</p>
            <p className="text-xs text-gray-400 mt-1">Você pode selecionar vários arquivos (ex: um por ano)</p>
          </div>

          {uploadedFiles.length > 0 && (
            <div className="flex-1 overflow-y-auto max-h-60 border border-gray-200 rounded-lg p-2">
              <div className="flex justify-between items-center mb-2 px-2 pb-2 border-b border-gray-100">
                <span className="text-sm font-medium text-gray-700">{uploadedFiles.length} arquivo(s)</span>
                <div className="flex gap-2">
                  <button onClick={() => toggleAllFiles(true)} className="text-xs text-indigo-600 hover:underline">Selecionar Todos</button>
                  <button onClick={() => toggleAllFiles(false)} className="text-xs text-gray-500 hover:underline">Desmarcar Todos</button>
                </div>
              </div>
              <ul className="space-y-1">
                {uploadedFiles.map(f => (
                  <li key={f.id} className="flex flex-col p-2 hover:bg-gray-50 rounded-md group">
                    <div className="flex items-center justify-between">
                      <label className="flex items-center gap-3 cursor-pointer flex-1 overflow-hidden">
                        <input 
                          type="checkbox" 
                          checked={f.selected} 
                          onChange={() => toggleFileSelection(f.id)}
                          className="w-4 h-4 text-indigo-600 rounded border-gray-300"
                          disabled={f.status === 'processing'}
                        />
                        <span className="text-sm text-gray-700 truncate">{f.file.name}</span>
                      </label>
                      <div className="flex items-center gap-2">
                        {f.status === 'processing' && (
                          <span className="text-xs text-emerald-600 flex items-center gap-1">
                            <Loader2 className="w-3 h-3 animate-spin" />
                            OCR {f.progress}%
                          </span>
                        )}
                        {f.status === 'error' && (
                          <span className="text-xs text-red-500 flex items-center gap-1" title={f.error}>
                            <AlertCircle className="w-3 h-3" />
                            Erro
                          </span>
                        )}
                        <button 
                          onClick={() => removeFile(f.id)}
                          className="text-gray-400 hover:text-red-500 opacity-0 group-hover:opacity-100 transition-opacity p-1"
                          title="Remover arquivo"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                    {f.status === 'processing' && (
                      <div className="w-full bg-gray-200 rounded-full h-1.5 mt-2">
                        <div className="bg-emerald-500 h-1.5 rounded-full transition-all duration-300" style={{ width: `${f.progress}%` }}></div>
                      </div>
                    )}
                    {f.status === 'error' && (
                      <p className="text-xs text-red-500 mt-1 ml-7">{f.error}</p>
                    )}
                  </li>
                ))}
              </ul>
            </div>
          )}
        </div>

        {/* Template Upload */}
        <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
          <h2 className="text-lg font-semibold mb-4 flex items-center gap-2">
            <FileSpreadsheet className="w-5 h-5 text-emerald-600" />
            Modelo de Formatação (Opcional)
          </h2>
          <div 
            {...getTemplateProps()} 
            className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors h-[282px] flex flex-col justify-center ${
              isTemplateDrag ? 'border-emerald-500 bg-emerald-50' : 'border-gray-300 hover:border-emerald-400 hover:bg-gray-50'
            }`}
          >
            <input {...getTemplateInputProps()} />
            <Upload className="w-8 h-8 mx-auto text-gray-400 mb-3" />
            {templateFile ? (
              <div>
                <p className="text-sm font-medium text-emerald-600">{templateFile.name}</p>
                <button 
                  onClick={(e) => { e.stopPropagation(); setTemplateFile(null); }}
                  className="mt-3 text-xs text-red-500 hover:text-red-700 font-medium"
                >
                  Remover Modelo
                </button>
              </div>
            ) : (
              <div>
                <p className="text-sm text-gray-600">Arraste e solte um modelo Excel aqui</p>
                <p className="text-xs text-gray-400 mt-1">Usado apenas para formatação de saída</p>
              </div>
            )}
          </div>
        </div>
      </div>

      <div className="flex flex-col items-center justify-center mb-12">
        <button
          onClick={handlePreAnalyze}
          disabled={uploadedFiles.filter(f => f.selected).length === 0 || isPreAnalyzing || isExtracting}
          className={`flex items-center gap-2 px-8 py-3 rounded-lg font-medium text-white transition-all ${
            uploadedFiles.filter(f => f.selected).length === 0 || isPreAnalyzing || isExtracting 
              ? 'bg-gray-400 cursor-not-allowed' 
              : 'bg-indigo-600 hover:bg-indigo-700 shadow-md hover:shadow-lg'
          }`}
        >
          {isPreAnalyzing ? (
            <>
              <Loader2 className="w-5 h-5 animate-spin" />
              Analisando Documento...
            </>
          ) : isExtracting ? (
            <>
              <Loader2 className="w-5 h-5 animate-spin" />
              Extraindo Dados...
            </>
          ) : (
            <>
              <Settings className="w-5 h-5" />
              Analisar e Extrair
            </>
          )}
        </button>

        {/* Progress Indicator */}
        {isExtracting && (
          <div className="w-full max-w-md mt-6">
            <div className="flex justify-between text-sm text-gray-600 mb-2 font-medium">
              <span>A IA está processando os documentos...</span>
              <span className="tabular-nums">{(elapsedTime / 1000).toFixed(1)}s</span>
            </div>
            <div className="w-full bg-gray-200 rounded-full h-2.5 overflow-hidden">
              <div 
                className="bg-indigo-600 h-2.5 rounded-full transition-all duration-300 ease-out" 
                style={{ width: `${progress}%` }}
              ></div>
            </div>
            <p className="text-xs text-gray-500 text-center mt-2">
              Processando {uploadedFiles.filter(f => f.selected).length} arquivo(s). Isso pode levar alguns minutos.
            </p>
          </div>
        )}
      </div>

      {/* Extraction History Dashboard */}
      {history.length > 0 && (
        <div className="space-y-6">
          <div className="flex items-center gap-2 border-b border-gray-200 pb-4">
            <History className="w-6 h-6 text-gray-700" />
            <h2 className="text-2xl font-bold text-gray-900">Histórico de Extrações</h2>
          </div>
          
          <div className="grid grid-cols-1 gap-6">
            {history.map((session) => (
              <div key={session.id} className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden hover:shadow-md transition-shadow">
                <div className="p-5 border-b border-gray-100 bg-gray-50 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div>
                    <h3 className="text-lg font-semibold text-gray-900 flex items-center gap-2">
                      <FileText className="w-4 h-4 text-indigo-500" />
                      {session.fileNames.length === 1 ? session.fileNames[0] : `${session.fileNames.length} arquivos processados`}
                    </h3>
                    <p className="text-sm text-gray-500 mt-1">
                      Extraído em {new Date(session.extractedAt).toLocaleString('pt-BR')}
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-3">
                    <button
                      onClick={() => exportToExcel(session)}
                      className="flex items-center gap-2 px-4 py-2 rounded-md font-medium text-sm transition-colors bg-emerald-50 text-emerald-700 hover:bg-emerald-100"
                    >
                      <Download className="w-4 h-4" />
                      Exportar para Excel
                    </button>
                    <button
                      onClick={() => handleSave(session)}
                      disabled={savingId === session.id || session.saved}
                      className={`flex items-center gap-2 px-4 py-2 rounded-md font-medium text-sm transition-colors ${
                        session.saved
                          ? 'bg-indigo-100 text-indigo-700 cursor-default'
                          : savingId === session.id
                          ? 'bg-gray-100 text-gray-500 cursor-not-allowed'
                          : 'bg-indigo-50 text-indigo-700 hover:bg-indigo-100'
                      }`}
                    >
                      {session.saved ? (
                        <>
                          <CheckCircle2 className="w-4 h-4" />
                          Salvo
                        </>
                      ) : savingId === session.id ? (
                        <>
                          <Loader2 className="w-4 h-4 animate-spin" />
                          Salvando...
                        </>
                      ) : (
                        <>
                          <Save className="w-4 h-4" />
                          Salvar no Banco
                        </>
                      )}
                    </button>
                  </div>
                </div>
                
                <div className="p-5 grid grid-cols-1 sm:grid-cols-3 gap-6">
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-blue-50 rounded-lg text-blue-600">
                      <Users className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-xs font-medium text-gray-500 uppercase tracking-wider">Funcionários Encontrados</p>
                      <p className="text-sm font-medium text-gray-900 mt-1">
                        {session.employeeNames.length > 0 ? session.employeeNames.join(', ') : 'Desconhecido'}
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-purple-50 rounded-lg text-purple-600">
                      <Calendar className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-xs font-medium text-gray-500 uppercase tracking-wider">Anos Identificados</p>
                      <p className="text-sm font-medium text-gray-900 mt-1">
                        {session.yearsCovered.length > 0 ? session.yearsCovered.join(', ') : 'Não detectado'}
                      </p>
                    </div>
                  </div>
                  
                  <div className="flex items-start gap-3">
                    <div className="p-2 bg-orange-50 rounded-lg text-orange-600">
                      <Clock className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-xs font-medium text-gray-500 uppercase tracking-wider">Tempo de Extração</p>
                      <p className="text-sm font-medium text-gray-900 mt-1">
                        {(session.extractionTimeMs / 1000).toFixed(1)} segundos
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Pre-Analysis Modal */}
      {showPreAnalysisModal && preAnalysisData && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50">
              <h3 className="text-xl font-bold text-gray-900">Configurar Extração</h3>
              <button 
                onClick={() => setShowPreAnalysisModal(false)}
                className="text-gray-400 hover:text-gray-600 transition-colors"
              >
                <X className="w-6 h-6" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto flex-1 space-y-8">
              <p className="text-gray-600 text-sm">
                Analisamos o primeiro documento e encontramos os seguintes dados. Selecione o que deseja extrair e renomeie as verbas se necessário.
              </p>

              {(extractionType === 'timecard' || extractionType === 'both') && (
                <div>
                  <h4 className="font-semibold text-gray-900 mb-3 flex items-center gap-2">
                    <Clock className="w-4 h-4 text-indigo-600" />
                    Colunas de Horário Encontradas
                  </h4>
                  {preAnalysisData.timeColumns.length > 0 ? (
                    <div className="grid grid-cols-2 gap-3">
                      {preAnalysisData.timeColumns.map(col => (
                        <label key={col} className="flex items-center gap-3 p-3 border rounded-lg cursor-pointer hover:bg-gray-50">
                          <input 
                            type="checkbox" 
                            checked={selectedTimeColumns.includes(col)}
                            onChange={(e) => {
                              if (e.target.checked) {
                                setSelectedTimeColumns(prev => [...prev, col]);
                              } else {
                                setSelectedTimeColumns(prev => prev.filter(c => c !== col));
                              }
                            }}
                            className="w-4 h-4 text-indigo-600 rounded border-gray-300"
                          />
                          <span className="text-sm font-medium text-gray-700">{col}</span>
                        </label>
                      ))}
                    </div>
                  ) : (
                    <p className="text-sm text-amber-600 bg-amber-50 p-3 rounded-lg border border-amber-200">
                      Nenhuma coluna de horário foi identificada automaticamente.
                    </p>
                  )}
                </div>
              )}

              {(extractionType === 'payslip' || extractionType === 'both') && (
                <div>
                  <h4 className="font-semibold text-gray-900 mb-3 flex items-center gap-2">
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    Verbas do Holerite Encontradas
                  </h4>
                  <p className="text-xs text-gray-500 mb-4">
                    A primeira coluna mostra as verbas do modelo (Desejadas). A segunda mostra o nome encontrado no documento. Se estiver em branco ou incorreto, digite o nome exato como aparece no documento.
                  </p>
                  {fieldMappings.length > 0 ? (
                    <DragDropContext onDragEnd={(result) => {
                      if (!result.destination) return;
                      const items = Array.from(fieldMappings);
                      const [reorderedItem] = items.splice(result.source.index, 1);
                      items.splice(result.destination.index, 0, reorderedItem);
                      setFieldMappings(items);
                    }}>
                      <Droppable droppableId="fieldMappings">
                        {(provided) => (
                          <div {...provided.droppableProps} ref={provided.innerRef} className="space-y-3">
                            {fieldMappings.map((mapping, index) => (
                              <Draggable key={`mapping-${index}-${mapping.targetName}`} draggableId={`mapping-${index}-${mapping.targetName}`} index={index}>
                                {(provided) => (
                                  <div 
                                    ref={provided.innerRef}
                                    {...provided.draggableProps}
                                    className="flex items-center gap-4 bg-white p-2 border border-gray-100 rounded-md shadow-sm"
                                  >
                                    <div {...provided.dragHandleProps} className="text-gray-400 hover:text-gray-600 cursor-grab active:cursor-grabbing">
                                      <GripVertical className="w-5 h-5" />
                                    </div>
                                    <div className="flex-1">
                                      <label className="block text-xs font-medium text-gray-500 mb-1">Nome no Modelo (Desejado)</label>
                                      <input 
                                        type="text" 
                                        value={mapping.targetName} 
                                        disabled 
                                        className="w-full p-2 text-sm bg-gray-50 border border-gray-200 rounded-md text-gray-600"
                                      />
                                    </div>
                                    <div className="text-gray-400 mt-5">=</div>
                                    <div className="flex-1">
                                      <label className="block text-xs font-medium text-gray-500 mb-1">Nome no Documento</label>
                                      <input 
                                        type="text" 
                                        value={mapping.sourceName} 
                                        onChange={(e) => {
                                          const newMappings = [...fieldMappings];
                                          newMappings[index].sourceName = e.target.value;
                                          setFieldMappings(newMappings);
                                        }}
                                        placeholder="Digite o nome no documento"
                                        className="w-full p-2 text-sm border border-gray-300 rounded-md focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500"
                                      />
                                    </div>
                                  </div>
                                )}
                              </Draggable>
                            ))}
                            {provided.placeholder}
                          </div>
                        )}
                      </Droppable>
                    </DragDropContext>
                  ) : (
                    <p className="text-sm text-amber-600 bg-amber-50 p-3 rounded-lg border border-amber-200">
                      Nenhuma verba de holerite foi identificada automaticamente.
                    </p>
                  )}
                </div>
              )}
            </div>
            
            <div className="p-6 border-t border-gray-100 bg-gray-50 flex justify-end gap-3">
              <button 
                onClick={() => setShowPreAnalysisModal(false)}
                className="px-5 py-2.5 text-sm font-medium text-gray-700 hover:bg-gray-200 rounded-lg transition-colors"
              >
                Cancelar
              </button>
              <button 
                onClick={handleExtract}
                className="px-5 py-2.5 text-sm font-medium text-white bg-indigo-600 hover:bg-indigo-700 rounded-lg transition-colors shadow-sm"
              >
                Confirmar e Iniciar Extração
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Error Modal */}
      {errorModal.isOpen && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="p-6 text-center">
              <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <AlertCircle className="w-8 h-8 text-red-600" />
              </div>
              <h3 className="text-xl font-bold text-gray-900 mb-2">{errorModal.title}</h3>
              <p className="text-gray-600 text-sm">{errorModal.message}</p>
            </div>
            <div className="p-4 border-t border-gray-100 bg-gray-50 flex justify-center">
              <button 
                onClick={() => setErrorModal({ isOpen: false, title: '', message: '' })}
                className="px-6 py-2 bg-gray-900 text-white font-medium rounded-lg hover:bg-gray-800 transition-colors w-full"
              >
                Entendi
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
