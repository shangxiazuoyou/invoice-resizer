import { useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { PDFDocument, rgb, StandardFonts, PDFPage } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'
import './App.css'

interface SizePreset {
  name: string
  width: number
  height: number
}

const SIZE_PRESETS: SizePreset[] = [
  { name: '小报销单 (5.5×8cm)', width: 155.91, height: 226.77 },
  { name: '中报销单 (8×12cm)', width: 226.77, height: 340.16 },
  { name: '标准报销单 (10×15cm)', width: 283.46, height: 425.20 },
  { name: '大报销单 (15×20cm)', width: 425.20, height: 566.93 }
]

function App() {
  const [uploadedFiles, setUploadedFiles] = useState<File[]>([])
  const [selectedPreset, setSelectedPreset] = useState<SizePreset>(SIZE_PRESETS[0])
  const [customWidth, setCustomWidth] = useState('')
  const [customHeight, setCustomHeight] = useState('')
  const [useCustomSize, setUseCustomSize] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [processedPdfUrl, setProcessedPdfUrl] = useState<string | null>(null)

  const onDrop = (acceptedFiles: File[]) => {
    const pdfFiles = acceptedFiles.filter(file => file.type === 'application/pdf')
    if (pdfFiles.length > 0) {
      setUploadedFiles(prev => [...prev, ...pdfFiles])
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
      'application/pdf': ['.pdf']
    },
    multiple: true
  })

  const cmToPt = (cm: number) => cm * 28.346

  const validatePdfFile = async (file: File): Promise<boolean> => {
    try {
      const arrayBuffer = await file.arrayBuffer()
      const header = new Uint8Array(arrayBuffer.slice(0, 5))
      const pdfSignature = '%PDF'
      const headerString = Array.from(header).map(byte => String.fromCharCode(byte)).join('')
      return headerString.startsWith(pdfSignature)
    } catch {
      return false
    }
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

      if (targetWidth <= 0 || targetHeight <= 0) {
        throw new Error('目标尺寸必须大于0')
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
        
        const isValidPdf = await validatePdfFile(file)
        if (!isValidPdf) {
          console.warn(`跳过无效PDF文件: ${file.name}`)
          continue
        }

        const arrayBuffer = await file.arrayBuffer()
        console.log('文件读取成功，大小:', arrayBuffer.byteLength)
        
        const pdfDoc = await PDFDocument.load(arrayBuffer)
        console.log('PDF加载成功')
        
        const pages = pdfDoc.getPages()
        console.log('PDF页数:', pages.length)
        
        if (pages.length === 0) {
          console.warn(`跳过空PDF文件: ${file.name}`)
          continue
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
            pageNumber: i + 1
          })
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

        // 嵌入并绘制发票页面
        const embeddedPage = await outputPdf.embedPage(pageInfo.page)
        
        // 计算发票在当前位置的坐标
        const invoiceX = currentX + 5 // 5px的裁剪线边距
        const invoiceY = currentY + 5
        
        if (currentPage) {
          currentPage.drawPage(embeddedPage, {
            x: invoiceX,
            y: invoiceY,
            width: pageInfo.scaledWidth,
            height: pageInfo.scaledHeight
          })
          
          // 绘制裁剪线框 - 避免与页脚重叠
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
      console.error('处理PDF时出错:', error)
      
      let errorMessage = '处理PDF时出错: '
      if (error instanceof Error) {
        errorMessage += error.message
      } else {
        errorMessage += '未知错误'
      }
      
      if (error instanceof Error && error.message.includes('Invalid PDF')) {
        errorMessage = 'PDF文件已损坏或格式不正确，请尝试重新下载发票'
      } else if (error instanceof Error && error.message.includes('encrypted')) {
        errorMessage = 'PDF文件已加密，请先解除密码保护'
      } else if (error instanceof Error && error.message.includes('目标尺寸')) {
        errorMessage = '请输入有效的目标尺寸'
      }
      
      alert(errorMessage)
    } finally {
      setIsProcessing(false)
    }
  }

  return (
    <div className="app">
      <h1>发票PDF尺寸调整工具</h1>
      
      <div className="upload-section">
        <div {...getRootProps()} className={`dropzone ${isDragActive ? 'active' : ''}`}>
          <input {...getInputProps()} />
          {uploadedFiles.length > 0 ? (
            <p>已选择 {uploadedFiles.length} 个PDF文件</p>
          ) : (
            <p>拖放PDF文件到这里，或点击选择文件（支持多选）</p>
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
        <h3>选择报销单发票粘贴区域尺寸</h3>
        
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
              download="resized-invoice.pdf"
              className="download-btn"
            >
              下载PDF
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
