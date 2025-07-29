// PDF处理Worker
importScripts('https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js');

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
  SIZE_INVALID: 'SIZE_INVALID',
  MEMORY_INSUFFICIENT: 'MEMORY_INSUFFICIENT',
  PROCESSING_FAILED: 'PROCESSING_FAILED',
  CANVAS_ERROR: 'CANVAS_ERROR',
  NETWORK_ERROR: 'NETWORK_ERROR'
};

// 错误信息配置
const ERROR_CONFIGS = {
  [ErrorType.FILE_INVALID]: {
    message: '文件格式不支持',
    suggestion: '请上传PDF文件或PNG/JPG/WebP格式的图片'
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
  [ErrorType.PROCESSING_FAILED]: {
    message: '文件处理失败',
    suggestion: '请重试，或尝试处理更小的文件'
  },
  [ErrorType.MEMORY_INSUFFICIENT]: {
    message: '内存不足',
    suggestion: '请关闭其他浏览器标签页，或分批处理较少文件'
  }
};

// 自定义错误类
class ProcessingError extends Error {
  constructor(type, userMessage, suggestion, fileName, originalError) {
    super(userMessage);
    this.type = type;
    this.userMessage = userMessage;
    this.suggestion = suggestion;
    this.fileName = fileName;
    this.name = 'ProcessingError';
    
    if (originalError) {
      this.stack = originalError.stack;
    }
  }
}

// 工具函数
const cmToPt = (cm) => cm * 28.346;

// 验证PDF文件
const validatePdfFile = async (arrayBuffer, fileName) => {
  const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB
  
  if (arrayBuffer.byteLength === 0) {
    throw new ProcessingError(ErrorType.FILE_EMPTY, ERROR_CONFIGS[ErrorType.FILE_EMPTY].message, ERROR_CONFIGS[ErrorType.FILE_EMPTY].suggestion, fileName);
  }
  
  if (arrayBuffer.byteLength > MAX_FILE_SIZE) {
    throw new ProcessingError(ErrorType.FILE_TOO_LARGE, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].message, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].suggestion, fileName);
  }

  try {
    const header = new Uint8Array(arrayBuffer.slice(0, 8));
    const headerString = Array.from(header).map(byte => String.fromCharCode(byte)).join('');
    
    if (!headerString.startsWith('%PDF')) {
      throw new ProcessingError(ErrorType.FILE_INVALID, ERROR_CONFIGS[ErrorType.FILE_INVALID].message, ERROR_CONFIGS[ErrorType.FILE_INVALID].suggestion, fileName);
    }
    
    // 检查是否加密
    const textContent = new TextDecoder().decode(arrayBuffer.slice(0, Math.min(arrayBuffer.byteLength, 2048)));
    if (textContent.includes('/Encrypt') && !textContent.includes('/V 0')) {
      throw new ProcessingError(ErrorType.FILE_ENCRYPTED, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].message, ERROR_CONFIGS[ErrorType.FILE_ENCRYPTED].suggestion, fileName);
    }
  } catch (error) {
    if (error instanceof ProcessingError) {
      throw error;
    }
    throw new ProcessingError(ErrorType.FILE_CORRUPTED, ERROR_CONFIGS[ErrorType.FILE_CORRUPTED].message, ERROR_CONFIGS[ErrorType.FILE_CORRUPTED].suggestion, fileName, error);
  }
};

// 验证图片文件
const validateImageFile = (arrayBuffer, fileName, fileType) => {
  const MAX_FILE_SIZE = 20 * 1024 * 1024; // 20MB for images
  const SUPPORTED_TYPES = ['image/png', 'image/jpeg', 'image/jpg', 'image/webp'];
  
  if (arrayBuffer.byteLength === 0) {
    throw new ProcessingError(ErrorType.FILE_EMPTY, ERROR_CONFIGS[ErrorType.FILE_EMPTY].message, ERROR_CONFIGS[ErrorType.FILE_EMPTY].suggestion, fileName);
  }
  
  if (arrayBuffer.byteLength > MAX_FILE_SIZE) {
    throw new ProcessingError(ErrorType.FILE_TOO_LARGE, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].message, ERROR_CONFIGS[ErrorType.FILE_TOO_LARGE].suggestion, fileName);
  }
  
  if (!fileType.startsWith('image/') || !SUPPORTED_TYPES.includes(fileType)) {
    throw new ProcessingError(ErrorType.IMAGE_FORMAT_UNSUPPORTED, ERROR_CONFIGS[ErrorType.IMAGE_FORMAT_UNSUPPORTED].message, ERROR_CONFIGS[ErrorType.IMAGE_FORMAT_UNSUPPORTED].suggestion, fileName);
  }
};

// 处理PDF文件
const processPdfFile = async (arrayBuffer, fileName, targetWidth, targetHeight) => {
  await validatePdfFile(arrayBuffer, fileName);
  
  const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
  const pages = pdfDoc.getPages();
  
  if (pages.length === 0) {
    throw new ProcessingError(ErrorType.PDF_NO_PAGES, ERROR_CONFIGS[ErrorType.PDF_NO_PAGES].message, ERROR_CONFIGS[ErrorType.PDF_NO_PAGES].suggestion, fileName);
  }

  const processedPages = [];
  
  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    const { width: pageWidth, height: pageHeight } = page.getSize();
    
    // 按报销单尺寸计算缩放比例
    const scaleX = targetWidth / pageWidth;
    const scaleY = targetHeight / pageHeight;
    const scale = Math.min(scaleX, scaleY);
    
    const scaledWidth = pageWidth * scale;
    const scaledHeight = pageHeight * scale;
    
    processedPages.push({
      pageData: await pdfDoc.copyPages(pdfDoc, [i]),
      originalWidth: pageWidth,
      originalHeight: pageHeight,
      scaledWidth,
      scaledHeight,
      fileName: fileName,
      pageNumber: i + 1,
      isImage: false
    });
  }
  
  return processedPages;
};

// 处理图片文件
const processImageFile = async (arrayBuffer, fileName, fileType, targetWidth, targetHeight) => {
  validateImageFile(arrayBuffer, fileName, fileType);
  
  // 在Worker中处理图片比较复杂，我们将这部分留给主线程处理
  // 这里主要做验证和基础处理
  return {
    arrayBuffer,
    fileName,
    fileType,
    needsMainThreadProcessing: true // 标记需要主线程处理
  };
};

// 主要的处理函数
const processFiles = async (data) => {
  const { files, targetWidth, targetHeight, useCustomSize, customWidth, customHeight, selectedPresetWidth, selectedPresetHeight } = data;
  
  // 计算目标尺寸
  const finalTargetWidth = useCustomSize ? cmToPt(parseFloat(customWidth) || 0) : selectedPresetWidth;
  const finalTargetHeight = useCustomSize ? cmToPt(parseFloat(customHeight) || 0) : selectedPresetHeight;

  if (finalTargetWidth <= 0 || finalTargetHeight <= 0 || isNaN(finalTargetWidth) || isNaN(finalTargetHeight)) {
    throw new ProcessingError(ErrorType.SIZE_INVALID, ERROR_CONFIGS[ErrorType.SIZE_INVALID].message, ERROR_CONFIGS[ErrorType.SIZE_INVALID].suggestion);
  }

  const allPages = [];
  const imagesToProcess = [];
  
  // 发送进度更新
  self.postMessage({
    type: 'progress',
    message: '开始处理文件...',
    progress: 0
  });

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    
    // 发送进度更新
    self.postMessage({
      type: 'progress',
      message: `处理文件 ${fileIndex + 1}/${files.length}: ${file.name}`,
      progress: (fileIndex / files.length) * 100
    });

    try {
      if (file.type.startsWith('image/')) {
        // 图片文件需要主线程处理
        const imageResult = await processImageFile(file.arrayBuffer, file.name, file.type, finalTargetWidth, finalTargetHeight);
        imagesToProcess.push({
          ...imageResult,
          targetWidth: finalTargetWidth,
          targetHeight: finalTargetHeight
        });
      } else {
        // PDF文件在Worker中处理
        const pdfPages = await processPdfFile(file.arrayBuffer, file.name, finalTargetWidth, finalTargetHeight);
        allPages.push(...pdfPages);
      }
    } catch (error) {
      // 发送错误到主线程
      self.postMessage({
        type: 'error', 
        error: {
          type: error.type || ErrorType.PROCESSING_FAILED,
          userMessage: error.userMessage || error.message,
          suggestion: error.suggestion || ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].suggestion,
          fileName: error.fileName || file.name
        }
      });
      return;
    }
  }

  // 返回处理结果
  self.postMessage({
    type: 'success',
    data: {
      processedPages: allPages,
      imagesToProcess,
      targetWidth: finalTargetWidth,
      targetHeight: finalTargetHeight
    }
  });
};

// 监听主线程消息
self.onmessage = async function(e) {
  const { type, data } = e.data;
  
  try {
    switch (type) {
      case 'process':
        await processFiles(data);
        break;
      default:
        console.warn('Unknown message type:', type);
    }
  } catch (error) {
    self.postMessage({
      type: 'error',
      error: {
        type: error.type || ErrorType.PROCESSING_FAILED,
        userMessage: error.userMessage || error.message || '处理过程中发生未知错误',
        suggestion: error.suggestion || ERROR_CONFIGS[ErrorType.PROCESSING_FAILED].suggestion,
        fileName: error.fileName
      }
    });
  }
};