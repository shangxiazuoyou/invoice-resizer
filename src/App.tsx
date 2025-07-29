import { useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { PDFDocument, rgb, StandardFonts, PDFPage } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import * as XLSX from 'xlsx'
import html2canvas from 'html2canvas'
import './App.css'

// 错误类型常量
const ErrorType = {
  FILE_INVALID: 'FILE_INVALID',
  FILE_CORRUPTED: 'FILE_CORRUPTED', 
  FILE_ENCRYPTED: 'FILE_ENCRYPTED',
  FILE_TOO_LARGE: 'FILE_TOO_LARGE',
  FILE_EMPTY: 'FILE_EMPTY',
  IMAGE_LOAD_FAILED: 'IMAGE_LOAD_FAILED',
  IMAGE_FORMAT_UNSUPPORTED: 'IMAGE_FORMAT_UNSUPPORTED',
  PDF_PARSE_FAILED: 'PDF_PARSE_FAILED',
  PDF_NO_PAGES: 'PDF_NO_PAGES',
  EXCEL_PARSE_FAILED: 'EXCEL_PARSE_FAILED',
  EXCEL_NO_SHEETS: 'EXCEL_NO_SHEETS',
  EXCEL_FORMAT_UNSUPPORTED: 'EXCEL_FORMAT_UNSUPPORTED',
  SIZE_INVALID: 'SIZE_INVALID',
  MEMORY_INSUFFICIENT: 'MEMORY_INSUFFICIENT',
  PROCESSING_FAILED: 'PROCESSING_FAILED',
  CANVAS_ERROR: 'CANVAS_ERROR',
  NETWORK_ERROR: 'NETWORK_ERROR'
} as const

type ErrorType = typeof ErrorType[keyof typeof ErrorType]

// 自定义错误类
class ProcessingError extends Error {
  public readonly type: ErrorType
  public readonly userMessage: string
  public readonly suggestion: string
  public readonly fileName?: string

  constructor(type: ErrorType, userMessage: string, suggestion: string, fileName?: string, originalError?: Error) {
    super(userMessage)
    this.type = type
    this.userMessage = userMessage
    this.suggestion = suggestion
    this.fileName = fileName
    this.name = 'ProcessingError'
    
    if (originalError) {
      this.stack = originalError.stack
    }
  }
}

// 错误信息配置
const ERROR_CONFIGS = {
  [ErrorType.FILE_INVALID]: {
    message: '文件格式不支持',
    suggestion: '请上传PDF文件、PNG/JPG/WebP格式的图片或Excel文件(.xlsx/.xls)'
  },
  [ErrorType.FILE_CORRUPTED]: {
    message: '文件已损坏或格式不正确',
    suggestion: '请重新下载原始文件，或尝试其他格式的版本'
  },
  [ErrorType.FILE_ENCRYPTED]: {
    message: 'PDF文件已加密',
    suggestion: '请先解除PDF密码保护，或将PDF另存为新文件'
  },
  [ErrorType.FILE_TOO_LARGE]: {
    message: '文件过大',
    suggestion: '请压缩文件大小至50MB以下，或分批处理'
  },
  [ErrorType.FILE_EMPTY]: {
    message: '文件为空或无有效内容',
    suggestion: '请检查文件是否正常，或重新获取原始文件'
  },
  [ErrorType.IMAGE_LOAD_FAILED]: {
    message: '图片加载失败',
    suggestion: '请确认图片文件完整且未损坏，建议重新保存图片'
  },
  [ErrorType.IMAGE_FORMAT_UNSUPPORTED]: {
    message: '图片格式不支持',
    suggestion: '请将图片转换为PNG、JPG或WebP格式'
  },
  [ErrorType.PDF_PARSE_FAILED]: {
    message: 'PDF解析失败',
    suggestion: '请确认PDF文件正常，或尝试用PDF阅读器重新保存'
  },
  [ErrorType.PDF_NO_PAGES]: {
    message: 'PDF文件没有页面',
    suggestion: '请检查PDF文件是否有实际内容'
  },
  [ErrorType.EXCEL_PARSE_FAILED]: {
    message: 'Excel文件解析失败',
    suggestion: '请确认Excel文件正常，或尝试另存为xlsx格式'
  },
  [ErrorType.EXCEL_NO_SHEETS]: {
    message: 'Excel文件没有工作表',
    suggestion: '请检查Excel文件是否包含有效的工作表'
  },
  [ErrorType.EXCEL_FORMAT_UNSUPPORTED]: {
    message: 'Excel文件格式不支持',
    suggestion: '请将文件转换为.xlsx或.xls格式'
  },
  [ErrorType.SIZE_INVALID]: {
    message: '目标尺寸设置无效',
    suggestion: '请输入1-30cm范围内的有效数值'
  },
  [ErrorType.MEMORY_INSUFFICIENT]: {
    message: '内存不足',
    suggestion: '请关闭其他浏览器标签页，或分批处理较少文件'
  },
  [ErrorType.PROCESSING_FAILED]: {
    message: '文件处理失败',
    suggestion: '请重试，或尝试处理更小的文件'
  },
  [ErrorType.CANVAS_ERROR]: {
    message: '图片处理出错',
    suggestion: '请刷新页面重试，或更换浏览器'
  },
  [ErrorType.NETWORK_ERROR]: {
    message: '网络连接异常',
    suggestion: '请检查网络连接后重试'
  }
}

interface SizePreset {
  name: string
  width: number
  height: number
}

const SIZE_PRESETS: SizePreset[] = [
  { name: '小报销单 (5.5×8cm)', width: 155.91, height: 226.77 },
  { name: '中报销单 (8×12cm)', width: 226.77, height: 340.16 },
  { name: '标准报销单 (10×15cm)', width: 283.46, height: 425.20 },
  { name: '大报销单 (15×20cm)', width: 425.20, height: 566.93 },
  { name: '超大报销单 (21×11cm)', width: 595.28, height: 311.81 }
]

interface ErrorState {
  error: ProcessingError | null
  isVisible: boolean
}

function App() {
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([])
  const [selectedPreset, setSelectedPreset] = useState<SizePreset>(SIZE_PRESETS[0])
  const [customWidth, setCustomWidth] = useState('')
  const [customHeight, setCustomHeight] = useState('')
  const [useCustomSize, setUseCustomSize] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedPdfUrl, setProcessedPdfUrl] = useState<string | null>(null)
  const [errorState, setErrorState] = useState<ErrorState>({ error: null, isVisible: false })

  const showError = (error: ProcessingError) => {
    setErrorState({ error, isVisible: true })
    console.error('处理错误:', error)
  }

  const hideError = () => {
    setErrorState({ error: null, isVisible: false })
  }

  const onDrop = (acceptedFiles: File[]) => {
    const validFiles = acceptedFiles.filter(file => 
      file.type === 'application/pdf' || 
      file.type.startsWith('image/') ||
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel' ||
      file.name.toLowerCase().endsWith('.xlsx') ||
      file.name.toLowerCase().endsWith('.xls')
    )
    if (validFiles.length > 0) {
      setUploadedFiles(prev => [...prev, ...validFiles])
      setProcessedPdfUrl(null)
    }
  }

  const removeFile = (index: number) => {
    setUploadedFiles(prev => prev.filter((_, i) => i !== index))
  }

  const clearAllFiles = () => {
    setUploadedFiles([])
    setProcessedPdfUrl(null)
  }

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/pdf': ['.pdf'],
      'image/*': ['.png', '.jpg', '.jpeg', '.webp'],
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: true
  })

  const cmToPt = (cm: number) => cm * 28.346

  const validatePdfFile = async (file: File): Promise<void> => {
    const MAX_FILE_SIZE = 50 * 1024 * 1024 // 50MB
    
    if (file.size === 0) {
      throw new ProcessingError(ErrorType.FILE_EMPTY, ERROR_CONFIGS[ErrorType.FILE_EMPTY].message, ERROR_CONFIGS[ErrorType.FILE_EMPTY].suggestion, file.name)
    }
    
    if (file.size > MAX_FILE_SIZE) {
      throw new ProcessingError(ErrorType.FILE_TOO_LARGE, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].message, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].suggestion, file.name)
    }

    try {
      const arrayBuffer = await file.arrayBuffer()
      const header = new Uint8Array(arrayBuffer.slice(0, 8))
      const headerString = Array.from(header).map(byte => String.fromCharCode(byte)).join('')
      
      if (!headerString.startsWith('%PDF')) {
        throw new ProcessingError(ErrorType.FILE_INVALID, ERROR_CONFIGS[ErrorType.FILE_INVALID].message, ERROR_CONFIGS[ErrorType.FILE_INVALID].suggestion, file.name)
      }
      
      // 检查是否加密
      const textContent = new TextDecoder().decode(arrayBuffer.slice(0, Math.min(arrayBuffer.byteLength, 2048)))
      if (textContent.includes('/Encrypt') && !textContent.includes('/V 0')) {
        throw new ProcessingError(ErrorType.FILE_ENCRYPTED, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].message, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].suggestion, file.name)
      }
    } catch (error) {
      if (error instanceof ProcessingError) {
        throw error
      }
      throw new ProcessingError(ErrorType.FILE_CORRUPTED, ERROR_CONFIGS[ErrorType.FILE_CORRUPTED].message, ERROR_CONFIGS[ErrorType.FILE_CORRUPTED].suggestion, file.name, error as Error)
    }
  }

  const validateImageFile = (file: File): void => {
    const MAX_FILE_SIZE = 20 * 1024 * 1024 // 20MB for images
    const SUPPORTED_TYPES = ['image/png', 'image/jpeg', 'image/jpg', 'image/webp']
    
    if (file.size === 0) {
      throw new ProcessingError(ErrorType.FILE_EMPTY, ERROR_CONFIGS[ErrorType.FILE_EMPTY].message, ERROR_CONFIGS[ErrorType.FILE_EMPTY].suggestion, file.name)
    }
    
    if (file.size > MAX_FILE_SIZE) {
      throw new ProcessingError(ErrorType.FILE_TOO_LARGE, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].message, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].suggestion, file.name)
    }
    
    if (!file.type.startsWith('image/') || !SUPPORTED_TYPES.includes(file.type)) {
      throw new ProcessingError(ErrorType.IMAGE_FORMAT_UNSUPPORTED, ERROR_CONFIGS[ErrorType.IMAGE_FORMAT_UNSUPPORTED].message, ERROR_CONFIGS[ErrorType.IMAGE_FORMAT_UNSUPPORTED].suggestion, file.name)
    }
  }

  const validateExcelFile = (file: File): void => {
    const MAX_FILE_SIZE = 30 * 1024 * 1024 // 30MB for Excel files
    const SUPPORTED_TYPES = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ]
    
    if (file.size === 0) {
      throw new ProcessingError(ErrorType.FILE_EMPTY, ERROR_CONFIGS[ErrorType.FILE_EMPTY].message, ERROR_CONFIGS[ErrorType.FILE_EMPTY].suggestion, file.name)
    }
    
    if (file.size > MAX_FILE_SIZE) {
      throw new ProcessingError(ErrorType.FILE_TOO_LARGE, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].message, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].suggestion, file.name)
    }
    
    const isValidType = SUPPORTED_TYPES.includes(file.type) || 
                       file.name.toLowerCase().endsWith('.xlsx') || 
                       file.name.toLowerCase().endsWith('.xls')
    
    if (!isValidType) {
      throw new ProcessingError(ErrorType.EXCEL_FORMAT_UNSUPPORTED, ERROR_CONFIGS[ErrorType.EXCEL_FORMAT_UNSUPPORTED].message, ERROR_CONFIGS[ErrorType.EXCEL_FORMAT_UNSUPPORTED].suggestion, file.name)
    }
  }

  const imageToCanvas = (file: File): Promise<HTMLCanvasElement> => {
    return new Promise((resolve, reject) => {
      const img = new Image()
      const canvas = document.createElement('canvas')
      const ctx = canvas.getContext('2d')
      
      if (!ctx) {
        reject(new ProcessingError(ErrorType.CANVAS_ERROR, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].message, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].suggestion, file.name))
        return
      }

      const timeout = setTimeout(() => {
        reject(new ProcessingError(ErrorType.IMAGE_LOAD_FAILED, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].message, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].suggestion, file.name))
      }, 30000) // 30秒超时

      img.onload = () => {
        clearTimeout(timeout)
        try {
          // 检查图片尺寸是否合理
          if (img.width === 0 || img.height === 0) {
            reject(new ProcessingError(ErrorType.IMAGE_LOAD_FAILED, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].message, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].suggestion, file.name))
            return
          }
          
          // 检查图片是否过大（内存限制）
          if (img.width * img.height > 50 * 1024 * 1024) { // 50M pixels
            reject(new ProcessingError(ErrorType.MEMORY_INSUFFICIENT, ERROR_CONFIGS[ErrorType.MEMORY_INSUFFICIENT].message, ERROR_CONFIGS[ErrorType.MEMORY_INSUFFICIENT].suggestion, file.name))
            return
          }
          
          canvas.width = img.width
          canvas.height = img.height
          ctx.drawImage(img, 0, 0)
          resolve(canvas)
        } catch (error) {
          reject(new ProcessingError(ErrorType.CANVAS_ERROR, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].message, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].suggestion, file.name, error as Error))
        }
      }
      
      img.onerror = () => {
        clearTimeout(timeout)
        reject(new ProcessingError(ErrorType.IMAGE_LOAD_FAILED, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].message, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].suggestion, file.name))
      }
      
      try {
        img.src = URL.createObjectURL(file)
      } catch (error) {
        clearTimeout(timeout)
        reject(new ProcessingError(ErrorType.PROCESSING_FAILED, ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].message, ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].suggestion, file.name, error as Error))
      }
    })
  }

  const canvasToImageBytes = async (canvas: HTMLCanvasElement, fileName: string): Promise<Uint8Array> => {
    return new Promise((resolve, reject) => {
      try {
        canvas.toBlob((blob) => {
          if (!blob) {
            reject(new ProcessingError(ErrorType.CANVAS_ERROR, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].message, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].suggestion, fileName))
            return
          }
          
          const reader = new FileReader()
          reader.onload = () => {
            const arrayBuffer = reader.result as ArrayBuffer
            resolve(new Uint8Array(arrayBuffer))
          }
          reader.onerror = () => {
            reject(new ProcessingError(ErrorType.PROCESSING_FAILED, ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].message, ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].suggestion, fileName))
          }
          reader.readAsArrayBuffer(blob)
        }, 'image/png')
      } catch (error) {
        reject(new ProcessingError(ErrorType.CANVAS_ERROR, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].message, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].suggestion, fileName, error as Error))
      }
    })
  }

  const excelToCanvas = async (file: File): Promise<HTMLCanvasElement> => {
    return new Promise((resolve, reject) => {
      try {
        validateExcelFile(file)
        console.log('处理Excel文件:', file.name)

        const reader = new FileReader()
        reader.onload = async (e) => {
          try {
            const data = new Uint8Array(e.target?.result as ArrayBuffer)
            const workbook = XLSX.read(data, { type: 'array' })
            
            if (workbook.SheetNames.length === 0) {
              reject(new ProcessingError(ErrorType.EXCEL_NO_SHEETS, ERROR_CONFIGS[ErrorType.EXCEL_NO_SHEETS].message, ERROR_CONFIGS[ErrorType.EXCEL_NO_SHEETS].suggestion, file.name))
              return
            }

            // 取第一个工作表
            const firstSheetName = workbook.SheetNames[0]
            const worksheet = workbook.Sheets[firstSheetName]
            
            // 转换为HTML表格
            const htmlTable = XLSX.utils.sheet_to_html(worksheet, {
              id: 'excel-table',
              editable: false
            })

            // 创建临时容器
            const tempContainer = document.createElement('div')
            tempContainer.style.position = 'absolute'
            tempContainer.style.top = '-10000px'
            tempContainer.style.left = '-10000px'
            tempContainer.style.width = '595px' // A4宽度的像素值(210mm)
            tempContainer.style.height = 'auto'
            tempContainer.style.backgroundColor = 'white'
            tempContainer.style.padding = '20px'
            tempContainer.style.fontFamily = 'Arial, sans-serif'
            tempContainer.style.fontSize = '11px'
            tempContainer.style.boxSizing = 'border-box'
            
            tempContainer.innerHTML = htmlTable
            
            // 设置表格样式
            const table = tempContainer.querySelector('table')
            if (table) {
              table.style.borderCollapse = 'collapse'
              table.style.border = '1px solid #333'
              table.style.width = '100%'
              table.style.tableLayout = 'auto'
              table.style.wordBreak = 'break-word'
              
              // 设置单元格样式
              const cells = table.querySelectorAll('td, th')
              cells.forEach(cell => {
                const cellElement = cell as HTMLElement
                cellElement.style.border = '1px solid #333'
                cellElement.style.padding = '6px'
                cellElement.style.textAlign = 'left'
                cellElement.style.verticalAlign = 'top'
                cellElement.style.whiteSpace = 'normal' // 允许换行
                cellElement.style.wordWrap = 'break-word'
                cellElement.style.backgroundColor = 'white'
                cellElement.style.fontSize = '10px'
                cellElement.style.lineHeight = '1.3'
                cellElement.style.maxWidth = '150px' // 限制单元格最大宽度
              })
              
              // 设置表头样式
              const headers = table.querySelectorAll('th')
              headers.forEach(header => {
                const headerElement = header as HTMLElement
                headerElement.style.backgroundColor = '#f5f5f5'
                headerElement.style.fontWeight = 'bold'
              })
            }
            
            document.body.appendChild(tempContainer)
            
            // 使用html2canvas截图
            const originalCanvas = await html2canvas(tempContainer, {
              backgroundColor: 'white',
              scale: 2,
              useCORS: true,
              allowTaint: true,
              logging: false
            })
            
            // 清理临时容器
            document.body.removeChild(tempContainer)
            
            // 创建旋转90度的canvas
            const rotatedCanvas = document.createElement('canvas')
            const ctx = rotatedCanvas.getContext('2d')
            
            if (!ctx) {
              reject(new ProcessingError(ErrorType.CANVAS_ERROR, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].message, ERROR_CONFIGS[ErrorType.CANVAS_ERROR].suggestion, file.name))
              return
            }
            
            // 旋转后的尺寸：宽高互换
            rotatedCanvas.width = originalCanvas.height
            rotatedCanvas.height = originalCanvas.width
            
            // 移动到画布中心，旋转90度，然后绘制
            ctx.translate(rotatedCanvas.width / 2, rotatedCanvas.height / 2)
            ctx.rotate(Math.PI / 2) // 顺时针旋转90度
            ctx.drawImage(originalCanvas, -originalCanvas.width / 2, -originalCanvas.height / 2)
            
            resolve(rotatedCanvas)
          } catch (error) {
            reject(new ProcessingError(ErrorType.EXCEL_PARSE_FAILED, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].suggestion, file.name, error as Error))
          }
        }

        reader.onerror = () => {
          reject(new ProcessingError(ErrorType.EXCEL_PARSE_FAILED, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].suggestion, file.name))
        }

        reader.readAsArrayBuffer(file)
      } catch (error) {
        if (error instanceof ProcessingError) {
          reject(error)
        } else {
          reject(new ProcessingError(ErrorType.EXCEL_PARSE_FAILED, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].suggestion, file.name, error as Error))
        }
      }
    })
  }

  const processPdf = async () => {
    if (uploadedFiles.length === 0) return

    setIsProcessing(true)
    try {
      console.log('开始处理', uploadedFiles.length, '个PDF文件')
      
      const targetWidth = useCustomSize ? 
        cmToPt(parseFloat(customWidth) || 0) : 
        selectedPreset.width
      const targetHeight = useCustomSize ? 
        cmToPt(parseFloat(customHeight) || 0) : 
        selectedPreset.height

      if (targetWidth <= 0 || targetHeight <= 0 || isNaN(targetWidth) || isNaN(targetHeight)) {
        throw new ProcessingError(ErrorType.SIZE_INVALID, ERROR_CONFIGS[ErrorType.SIZE_INVALID].message, ERROR_CONFIGS[ErrorType.SIZE_INVALID].suggestion)
      }

      console.log('报销单尺寸:', targetWidth, 'x', targetHeight, 'pt')

      const outputPdf = await PDFDocument.create()
      outputPdf.registerFontkit(fontkit)
      
      const helveticaFont = await outputPdf.embedFont(StandardFonts.Helvetica)
      
      // A4纸尺寸 (210×297mm = 595×842pt)
      const A4_WIDTH = 595.28
      const A4_HEIGHT = 841.89

      // 收集所有需要处理的页面
      const allPages = []
      
      for (let fileIndex = 0; fileIndex < uploadedFiles.length; fileIndex++) {
        const file = uploadedFiles[fileIndex]
        console.log(`处理第${fileIndex + 1}个文件:`, file.name)
        
        if (file.type.startsWith('image/')) {
          // 处理图片文件
          try {
            validateImageFile(file)
            console.log('处理图片文件:', file.name)
            
            // 将图片转换为Canvas
            const canvas = await imageToCanvas(file)
            const imageBytes = await canvasToImageBytes(canvas, file.name)
            
            // 创建临时PDF以获取页面对象
            const tempPdf = await PDFDocument.create()
            const pngImage = await tempPdf.embedPng(imageBytes)
            const tempPage = tempPdf.addPage([pngImage.width, pngImage.height])
            tempPage.drawImage(pngImage, { x: 0, y: 0 })
            
            const pageWidth = pngImage.width
            const pageHeight = pngImage.height
            
            // 按报销单尺寸计算缩放比例
            const scaleX = targetWidth / pageWidth
            const scaleY = targetHeight / pageHeight
            const scale = Math.min(scaleX, scaleY)
            
            const scaledWidth = pageWidth * scale
            const scaledHeight = pageHeight * scale
            
            allPages.push({
              page: tempPage,
              originalWidth: pageWidth,
              originalHeight: pageHeight,
              scaledWidth,
              scaledHeight,
              fileName: file.name,
              pageNumber: 1,
              isImage: true,
              imageBytes
            })
            
            console.log(`图片处理完成: ${file.name}, 尺寸: ${pageWidth}x${pageHeight}`)
          } catch (error) {
            console.error(`处理图片文件失败: ${file.name}`, error)
            if (error instanceof ProcessingError) {
              throw error // 向上传播详细错误
            }
            throw new ProcessingError(ErrorType.IMAGE_LOAD_FAILED, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].message, ERROR_CONFIGS[ErrorType.IMAGE_LOAD_FAILED].suggestion, file.name, error as Error)
          }
        } else if (file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                   file.type === 'application/vnd.ms-excel' ||
                   file.name.toLowerCase().endsWith('.xlsx') ||
                   file.name.toLowerCase().endsWith('.xls')) {
          // 处理Excel文件
          try {
            console.log('处理Excel文件:', file.name)
            
            // 将Excel转换为Canvas
            const canvas = await excelToCanvas(file)
            const imageBytes = await canvasToImageBytes(canvas, file.name)
            
            // 创建临时PDF以获取页面对象
            const tempPdf = await PDFDocument.create()
            const pngImage = await tempPdf.embedPng(imageBytes)
            const tempPage = tempPdf.addPage([pngImage.width, pngImage.height])
            tempPage.drawImage(pngImage, { x: 0, y: 0 })
            
            const pageWidth = pngImage.width
            const pageHeight = pngImage.height
            
            // 按报销单尺寸计算缩放比例
            const scaleX = targetWidth / pageWidth
            const scaleY = targetHeight / pageHeight
            const scale = Math.min(scaleX, scaleY)
            
            const scaledWidth = pageWidth * scale
            const scaledHeight = pageHeight * scale
            
            allPages.push({
              page: tempPage,
              originalWidth: pageWidth,
              originalHeight: pageHeight,
              scaledWidth,
              scaledHeight,
              fileName: file.name,
              pageNumber: 1,
              isImage: true,
              imageBytes
            })
            
            console.log(`Excel处理完成: ${file.name}, 尺寸: ${pageWidth}x${pageHeight}`)
          } catch (error) {
            console.error(`处理Excel文件失败: ${file.name}`, error)
            if (error instanceof ProcessingError) {
              throw error // 向上传播详细错误
            }
            throw new ProcessingError(ErrorType.EXCEL_PARSE_FAILED, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.EXCEL_PARSE_FAILED].suggestion, file.name, error as Error)
          }
        } else {
          // 处理PDF文件
          try {
            await validatePdfFile(file)
            console.log('处理PDF文件:', file.name)

            const arrayBuffer = await file.arrayBuffer()
            console.log('文件读取成功，大小:', arrayBuffer.byteLength)
            
            const pdfDoc = await PDFDocument.load(arrayBuffer)
            console.log('PDF加载成功')
            
            const pages = pdfDoc.getPages()
            console.log('PDF页数:', pages.length)
            
            if (pages.length === 0) {
              throw new ProcessingError(ErrorType.PDF_NO_PAGES, ERROR_CONFIGS[ErrorType.PDF_NO_PAGES].message, ERROR_CONFIGS[ErrorType.PDF_NO_PAGES].suggestion, file.name)
            }

          for (let i = 0; i < pages.length; i++) {
            const page = pages[i]
            const { width: pageWidth, height: pageHeight } = page.getSize()
            
            // 按报销单尺寸计算缩放比例
            const scaleX = targetWidth / pageWidth
            const scaleY = targetHeight / pageHeight
            const scale = Math.min(scaleX, scaleY)
            
            const scaledWidth = pageWidth * scale
            const scaledHeight = pageHeight * scale
            
            allPages.push({
              page,
              originalWidth: pageWidth,
              originalHeight: pageHeight,
              scaledWidth,
              scaledHeight,
              fileName: file.name,
              pageNumber: i + 1,
              isImage: false
            })
          }
          } catch (error) {
            console.error(`处理PDF文件失败: ${file.name}`, error)
            if (error instanceof ProcessingError) {
              throw error // 向上传播详细错误
            }
            
            // 根据错误类型进行分类
            if (error instanceof Error) {
              if (error.message.includes('Invalid PDF') || error.message.includes('corrupted')) {
                throw new ProcessingError(ErrorType.PDF_PARSE_FAILED, ERROR_CONFIGS[ErrorType.PDF_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.PDF_PARSE_FAILED].suggestion, file.name, error)
              }
              if (error.message.includes('encrypted')) {
                throw new ProcessingError(ErrorType.FILE_ENCRYPTED, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].message, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].suggestion, file.name, error)
              }
            }
            
            throw new ProcessingError(ErrorType.PDF_PARSE_FAILED, ERROR_CONFIGS[ErrorType.PDF_PARSE_FAILED].message, ERROR_CONFIGS[ErrorType.PDF_PARSE_FAILED].suggestion, file.name, error as Error)
          }
        }
      }

      console.log(`总共需要处理 ${allPages.length} 个页面`)

      // 智能布局：在A4纸上排列发票
      const margin = 20 // A4纸边距
      const gap = 15 // 发票之间的间隙
      let currentPage: PDFPage | null = null
      let currentX = margin
      let currentY = margin
      let rowHeight = 0

      for (let i = 0; i < allPages.length; i++) {
        const pageInfo = allPages[i]
        const invoiceWidth = pageInfo.scaledWidth + 10 // 包含裁剪线的宽度
        const invoiceHeight = pageInfo.scaledHeight + 10 // 包含裁剪线的高度

        // 检查是否需要换行
        if (currentX + invoiceWidth > A4_WIDTH - margin) {
          // 换行
          currentX = margin
          currentY += rowHeight + gap
          rowHeight = 0
        }

        // 检查是否需要新页面 - 确保发票不会与页脚重叠
        if (currentY + invoiceHeight > A4_HEIGHT - margin - 80) { // 80是页脚预留空间
          // 创建新页面
          currentPage = outputPdf.addPage([A4_WIDTH, A4_HEIGHT])
          currentX = margin
          currentY = margin
          rowHeight = 0
          
          // 添加页面说明
          const pageText = `Target Size: ${useCustomSize ? customWidth : (selectedPreset.width/28.346).toFixed(1)}x${useCustomSize ? customHeight : (selectedPreset.height/28.346).toFixed(1)}cm - Cut along dashed lines`
          currentPage.drawText(pageText, {
            x: margin,
            y: A4_HEIGHT - 20,
            size: 10,
            color: rgb(0.5, 0.5, 0.5),
            font: helveticaFont
          })
          
          // 添加页脚说明
          // currentPage.drawText('Print this A4 page and cut along the dashed lines to get your receipts for reimbursement', {
          //   x: margin,
          //   y: 25,
          //   size: 8,
          //   color: rgb(0.6, 0.6, 0.6),
          //   font: helveticaFont
          // })
        }

        // 如果还没有页面，创建第一页
        if (!currentPage) {
          currentPage = outputPdf.addPage([A4_WIDTH, A4_HEIGHT])
          
          // 添加页面说明
          const pageText = `Target Size: ${useCustomSize ? customWidth : (selectedPreset.width/28.346).toFixed(1)}x${useCustomSize ? customHeight : (selectedPreset.height/28.346).toFixed(1)}cm - Cut along dashed lines`
          currentPage.drawText(pageText, {
            x: margin,
            y: A4_HEIGHT - 20,
            size: 10,
            color: rgb(0.5, 0.5, 0.5),
            font: helveticaFont
          })
          
          // 添加页脚说明
          // currentPage.drawText('Print this A4 page and cut along the dashed lines to get your receipts for reimbursement', {
          //   x: margin,
          //   y: 25,
          //   size: 8,
          //   color: rgb(0.6, 0.6, 0.6),
          //   font: helveticaFont
          // })
        }

        console.log(`在位置 (${currentX}, ${currentY}) 放置发票: ${pageInfo.fileName} 第${pageInfo.pageNumber}页`)

        // 计算发票在当前位置的坐标
        const invoiceX = currentX + 5 // 5px的裁剪线边距
        const invoiceY = currentY + 5
        
        if (currentPage) {
          if (pageInfo.isImage && pageInfo.imageBytes) {
            // 处理图片文件：直接嵌入图片
            const pngImage = await outputPdf.embedPng(pageInfo.imageBytes)
            currentPage.drawImage(pngImage, {
              x: invoiceX,
              y: invoiceY,
              width: pageInfo.scaledWidth,
              height: pageInfo.scaledHeight
            })
          } else {
            // 处理PDF文件：嵌入页面
            const embeddedPage = await outputPdf.embedPage(pageInfo.page)
            currentPage.drawPage(embeddedPage, {
              x: invoiceX,
              y: invoiceY,
              width: pageInfo.scaledWidth,
              height: pageInfo.scaledHeight
            })
          }
        }
        
        // 绘制裁剪线框 - 避免与页脚重叠
        if (currentPage) {
          currentPage.drawRectangle({
            x: currentX,
            y: currentY,
            width: pageInfo.scaledWidth + 10,
            height: pageInfo.scaledHeight + 10,
            borderColor: rgb(0.7, 0.7, 0.7),
            borderWidth: 1,
            borderDashArray: [5, 5]
          })
        }
        
        // 添加裁剪角标记
        const cornerLength = 8
        const corners = [
          { x: currentX, y: currentY }, // 左下
          { x: currentX + pageInfo.scaledWidth + 10, y: currentY }, // 右下
          { x: currentX, y: currentY + pageInfo.scaledHeight + 10 }, // 左上
          { x: currentX + pageInfo.scaledWidth + 10, y: currentY + pageInfo.scaledHeight + 10 } // 右上
        ]
        
        corners.forEach(corner => {
          if (currentPage) {
            // 水平线
            currentPage.drawLine({
              start: { x: corner.x - cornerLength/2, y: corner.y },
              end: { x: corner.x + cornerLength/2, y: corner.y },
              thickness: 1,
              color: rgb(0.5, 0.5, 0.5)
            })
            // 垂直线
            currentPage.drawLine({
              start: { x: corner.x, y: corner.y - cornerLength/2 },
              end: { x: corner.x, y: corner.y + cornerLength/2 },
              thickness: 1,
              color: rgb(0.5, 0.5, 0.5)
            })
          }
        })

        // 更新位置
        currentX += invoiceWidth + gap
        rowHeight = Math.max(rowHeight, invoiceHeight)
      }


      console.log('开始保存PDF')
      const pdfBytes = await outputPdf.save()
      console.log('PDF保存成功，大小:', pdfBytes.length)
      
      const blob = new Blob([pdfBytes], { type: 'application/pdf' })
      const url = URL.createObjectURL(blob)
      setProcessedPdfUrl(url)
      
      console.log('处理完成')
    } catch (error) {
      console.error('处理文件时出错:', error)
      
      if (error instanceof ProcessingError) {
        showError(error)
      } else if (error instanceof Error) {
        // 处理未预期的错误
        let errorType: ErrorType = ErrorType.PROCESSING_FAILED
        
        if (error.message.includes('memory') || error.message.includes('Memory')) {
          errorType = ErrorType.MEMORY_INSUFFICIENT
        } else if (error.message.includes('network') || error.message.includes('Network')) {
          errorType = ErrorType.NETWORK_ERROR
        }
        
        const processingError = new ProcessingError(
          errorType,
          ERROR_CONFIGS[errorType].message,
          ERROR_CONFIGS[errorType].suggestion,
          undefined,
          error
        )
        showError(processingError)
      } else {
        // 处理完全未知的错误
        const unknownError = new ProcessingError(
          ErrorType.PROCESSING_FAILED,
          ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].message,
          ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].suggestion
        )
        showError(unknownError)
      }
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <div className="app">
      <h1>发票/支付截图/Excel尺寸调整工具</h1>
      
      {errorState.isVisible && errorState.error && (
        <div className="error-section">
          <div className="error-content">
            <div className="error-header">
              <span className="error-icon">⚠️</span>
              <h4 className="error-title">
                {errorState.error.fileName ? `处理文件 "${errorState.error.fileName}" 时出错` : '处理出错'}
              </h4>
              <button onClick={hideError} className="error-close">×</button>
            </div>
            <div className="error-body">
              <p className="error-message">{errorState.error.userMessage}</p>
              <p className="error-suggestion">
                <strong>解决建议：</strong>{errorState.error.suggestion}
              </p>
              {errorState.error.type === ErrorType.FILE_TOO_LARGE && (
                <div className="error-details">
                  <p>文件大小限制：</p>
                  <ul>
                    <li>PDF文件：最大50MB</li>
                    <li>图片文件：最大20MB</li>
                  </ul>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
      
      <div className="upload-section">
        <div {...getRootProps()} className={`dropzone ${isDragActive ? 'active' : ''}`}>
          <input {...getInputProps()} />
          {uploadedFiles.length > 0 ? (
            <p>已选择 {uploadedFiles.length} 个文件</p>
          ) : (
            <p>拖放PDF文件、图片或Excel文件到这里，或点击选择文件（支持多选）</p>
          )}
        </div>
        
        {uploadedFiles.length > 0 && (
          <div className="file-list">
            <div className="file-list-header">
              <h4>已选择的文件 ({uploadedFiles.length})</h4>
              <button onClick={clearAllFiles} className="clear-btn">清空所有</button>
            </div>
            <div className="file-items">
              {uploadedFiles.map((file, index) => (
                <div key={index} className="file-item">
                  <span className="file-name">{file.name}</span>
                  <span className="file-size">({(file.size / 1024 / 1024).toFixed(1)}MB)</span>
                  <button onClick={() => removeFile(index)} className="remove-btn">删除</button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>

      <div className="size-section">
        <h3>选择报销单粘贴区域尺寸</h3>
        
        <div className="size-options">
          <label>
            <input
              type="radio"
              checked={!useCustomSize}
              onChange={() => setUseCustomSize(false)}
            />
            预设尺寸
          </label>
          
          {!useCustomSize && (
            <select 
              value={SIZE_PRESETS.indexOf(selectedPreset)}
              onChange={(e) => setSelectedPreset(SIZE_PRESETS[parseInt(e.target.value)])}
            >
              {SIZE_PRESETS.map((preset, index) => (
                <option key={index} value={index}>
                  {preset.name}
                </option>
              ))}
            </select>
          )}
        </div>

        <div className="size-options">
          <label>
            <input
              type="radio"
              checked={useCustomSize}
              onChange={() => setUseCustomSize(true)}
            />
自定义报销单尺寸 (cm)
          </label>
          
          {useCustomSize && (
            <div className="custom-size-inputs">
              <input
                type="number"
                placeholder="宽度"
                value={customWidth}
                onChange={(e) => setCustomWidth(e.target.value)}
                step="0.1"
                min="1"
                max="20"
              />
              ×
              <input
                type="number"
                placeholder="高度"
                value={customHeight}
                onChange={(e) => setCustomHeight(e.target.value)}
                step="0.1"
                min="1"
                max="30"
              />
            </div>
          )}
        </div>
      </div>

      <div className="action-section">
        <button 
          onClick={processPdf}
          disabled={uploadedFiles.length === 0 || isProcessing || (useCustomSize && (!customWidth || !customHeight))}
          className="process-btn"
        >
{isProcessing ? '处理中...' : '生成A4打印PDF'}
        </button>
      </div>

      {processedPdfUrl && (
        <div className="result-section">
          <h3>处理完成</h3>
          <div className="result-actions">
            <a 
              href={processedPdfUrl} 
              download="resized-documents.pdf"
              className="download-btn"
            >
              下载A4打印文件
            </a>
            <iframe 
              src={processedPdfUrl} 
              className="pdf-preview"
              title="PDF预览"
            />
          </div>
        </div>
      )}
    </div>
  )
}

export default App
