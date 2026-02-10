<script setup>
import { ref, computed, watch, onUnmounted } from 'vue';
import { useI18n } from 'vue-i18n';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } from 'docx';
import { createWorker } from 'tesseract.js';

// 使用i18n
const { t, locale, availableLocales } = useI18n();

// 语言列表
const languages = [
  { code: 'zh', name: '中文' },
  { code: 'en', name: 'English' },
  { code: 'ja', name: '日本語' },
  { code: 'ko', name: '한국어' },
  { code: 'fr', name: 'Français' },
  { code: 'de', name: 'Deutsch' },
  { code: 'es', name: 'Español' },
  { code: 'ru', name: 'Русский' },
  { code: 'pt', name: 'Português' },
  { code: 'it', name: 'Italiano' }
];

// 当前选择的语言
const selectedLanguage = ref(locale.value);

// 切换语言
const changeLanguage = (lang) => {
  selectedLanguage.value = lang;
  locale.value = lang;
  
  // 检查并更新PDF转换完成的提示信息
  const currentText = resultText.value;
  const pdfCompleteTexts = [
    '转化PDF完成,请点击下载',
    'PDF conversion completed, please click to download',
    'PDF変換が完了しました、ダウンロードするにはクリックしてください',
    'PDF 변환이 완료되었습니다, 다운로드하려면 클릭하세요',
    'Conversion PDF terminée, cliquez pour télécharger',
    'PDF-Konvertierung abgeschlossen, klicken Sie zum Herunterladen',
    '¡Conversión PDF completada, haz clic para descargar',
    'Преобразование PDF завершено, нажмите для скачивания',
    'Conversão PDF concluída, clique para baixar',
    'Conversione PDF completata, fai clic per scaricare'
  ];
  
  if (pdfCompleteTexts.includes(currentText)) {
    resultText.value = t('pdfConvertComplete');
  }
};

const imageUrl = ref('');
const imageUrls = ref([]); // 用于批处理多张图片
const resultText = ref('');
const isLoading = ref(false);
const errorMessage = ref('');
const progress = ref(0);
const progressText = ref('');
const outputFormat = ref('text'); // 'text', 'excel', 'pdf' or 'word'
const worker = ref(null); // Tesseract.js worker实例
const currentLanguage = ref('chi_sim'); // 默认中文语言包
const batchResults = ref([]); // 批处理结果
const isBatchMode = ref(false); // 是否处于批处理模式

// 监听输出格式变化，清空结果
watch(outputFormat, () => {
  resultText.value = '';
});

// 处理图片上传
const handleImageUpload = (event) => {
  const files = event.target.files;
  if (files.length > 0) {
    // 检查文件类型
    const validFiles = [];
    const invalidFiles = [];
    
    for (const file of files) {
      if (!file.type.startsWith('image/')) {
        invalidFiles.push(file.name);
      } else if (file.size > 1024 * 1024) { // 限制1MB
        invalidFiles.push(`${file.name} (文件过大)`);
      } else {
        validFiles.push(file);
      }
    }
    
    if (invalidFiles.length > 0) {
      alert(`请选择图片类型的文件：\n${invalidFiles.join('\n')}`);
    }
    
    if (validFiles.length > 0) {
      if (validFiles.length === 1) {
        // 单张图片模式
        isBatchMode.value = false;
        imageUrls.value = [];
        batchResults.value = [];
        
        const reader = new FileReader();
        reader.onload = (e) => {
          imageUrl.value = e.target.result;
          resultText.value = '';
          errorMessage.value = '';
        };
        reader.readAsDataURL(validFiles[0]);
      } else {
        // 批处理模式
        isBatchMode.value = true;
        imageUrl.value = '';
        resultText.value = '';
        errorMessage.value = '';
        batchResults.value = [];
        
        imageUrls.value = [];
        for (const file of validFiles) {
          const reader = new FileReader();
          reader.onload = (e) => {
            imageUrls.value.push(e.target.result);
          };
          reader.readAsDataURL(file);
        }
      }
    }
  }
};

// 处理拖拽上传
const handleDragOver = (event) => {
  event.preventDefault();
  event.stopPropagation();
};

const handleDrop = (event) => {
  event.preventDefault();
  event.stopPropagation();
  
  const files = event.dataTransfer.files;
  if (files.length > 0) {
    // 检查文件类型
    const validFiles = [];
    const invalidFiles = [];
    
    for (const file of files) {
      if (!file.type.startsWith('image/')) {
        invalidFiles.push(file.name);
      } else if (file.size > 1024 * 1024) { // 限制1MB
        invalidFiles.push(`${file.name} (文件过大)`);
      } else {
        validFiles.push(file);
      }
    }
    
    if (invalidFiles.length > 0) {
      alert(`请选择图片类型的文件：\n${invalidFiles.join('\n')}`);
    }
    
    if (validFiles.length > 0) {
      if (validFiles.length === 1) {
        // 单张图片模式
        isBatchMode.value = false;
        imageUrls.value = [];
        batchResults.value = [];
        
        const reader = new FileReader();
        reader.onload = (e) => {
          imageUrl.value = e.target.result;
          resultText.value = '';
          errorMessage.value = '';
        };
        reader.readAsDataURL(validFiles[0]);
      } else {
        // 批处理模式
        isBatchMode.value = true;
        imageUrl.value = '';
        resultText.value = '';
        errorMessage.value = '';
        batchResults.value = [];
        
        imageUrls.value = [];
        for (const file of validFiles) {
          const reader = new FileReader();
          reader.onload = (e) => {
            imageUrls.value.push(e.target.result);
          };
          reader.readAsDataURL(file);
        }
      }
    }
  }
};

// 图片预处理函数
const preprocessImage = (imageSrc) => {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    
    // 设置超时，防止图片加载时间过长
    const timeoutId = setTimeout(() => {
      reject(new Error('图片加载超时'));
    }, 10000); // 10秒超时
    
    img.onload = () => {
      clearTimeout(timeoutId);
      try {
        // 计算新尺寸（保持宽高比，限制最大宽度）
        const maxWidth = 1200;
        let width = img.width;
        let height = img.height;
        
        if (width > maxWidth) {
          const ratio = maxWidth / width;
          width *= ratio;
          height *= ratio;
        }
        
        canvas.width = width;
        canvas.height = height;
        
        // 绘制并预处理图片
        ctx.drawImage(img, 0, 0, width, height);
        
        // 获取图片数据
        const imageData = ctx.getImageData(0, 0, width, height);
        const data = imageData.data;
        
        // 简单的二值化处理
        for (let i = 0; i < data.length; i += 4) {
          const gray = (data[i] + data[i + 1] + data[i + 2]) / 3;
          const threshold = 128;
          const value = gray > threshold ? 255 : 0;
          data[i] = value;     // R
          data[i + 1] = value; // G
          data[i + 2] = value; // B
          // data[i + 3] 保持不变 (alpha)
        }
        
        // 将处理后的数据放回canvas
        ctx.putImageData(imageData, 0, 0);
        
        // 返回处理后的图片URL
        resolve(canvas.toDataURL('image/png'));
      } catch (error) {
        reject(error);
      }
    };
    
    img.onerror = () => {
      clearTimeout(timeoutId);
      reject(new Error('图片加载失败'));
    };
    
    img.src = imageSrc;
  });
};

// 初始化Tesseract.js worker
const initWorker = async () => {
  console.log('检查worker实例:', worker.value ? '已存在' : '不存在');
  
  if (!worker.value) {
    console.log('开始初始化新的worker，语言:', currentLanguage.value);
    // 添加超时机制
    const timeoutPromise = new Promise((_, reject) => {
      setTimeout(() => {
        reject(new Error('Tesseract.js初始化超时，请检查网络连接'));
      }, 60000); // 增加到60秒超时，确保有足够时间下载资源
    });
    
    // 同时执行初始化和超时检查
    try {
      // 使用最新的API方式创建worker，直接指定语言
      worker.value = await Promise.race([
        createWorker(currentLanguage.value, 1, {
          logger: (m) => {
            // 处理进度更新
            console.log('Tesseract.js状态:', m.status, '进度:', m.progress);
            if (m.status === 'recognizing text') {
              progress.value = m.progress * 100;
              progressText.value = `${t('recognizing')} ${Math.round(progress.value)}%`;
            } else if (m.status === 'loading tesseract core') {
              progress.value = m.progress * 30;
              progressText.value = `${t('loadingCore')} ${Math.round(progress.value)}%`;
            } else if (m.status === 'loading language pack') {
              progress.value = 30 + (m.progress * 40);
              progressText.value = `${t('loadingLanguage')} ${Math.round(progress.value)}%`;
            } else if (m.status === 'initializing api') {
              progress.value = 70 + (m.progress * 20);
              progressText.value = `${t('initializing')} ${Math.round(progress.value)}%`;
            }
          }
        }),
        timeoutPromise
      ]);
      console.log('Worker初始化成功');
    } catch (error) {
      console.error('Worker初始化失败:', error);
      throw error;
    }
  } else {
    console.log('重用已存在的worker实例');
  }
  return worker.value;
};

// 转换图片为文本、Excel或PDF
const imageToText = async () => {
  if (!imageUrl.value && imageUrls.value.length === 0) {
    errorMessage.value = t('pleaseUpload');
    return;
  }

  isLoading.value = true;
  progress.value = 0;
  progressText.value = t('initializing');

  errorMessage.value = '';
  resultText.value = '';

  try {
    // 如果选择PDF格式，直接调用图片转PDF功能
    if (outputFormat.value === 'pdf') {
      if (isBatchMode.value) {
        errorMessage.value = t('pdfBatchNotSupported');
        isLoading.value = false;
        return;
      }
      
      await imageToPdf();
      // 显示转换完成提示
      resultText.value = t('pdfConvertComplete');
      setTimeout(() => {
        isLoading.value = false;
        progress.value = 0;
        progressText.value = '';
      }, 500);
      return;
    }

    // 初始化worker
    progressText.value = t('initializing');
    progress.value = 0;
    
    try {
      const tessWorker = await initWorker();
      
      // 最新版本的Tesseract.js不需要单独调用loadLanguage和initialize
      // 语言已经在createWorker时指定了
      console.log('Worker准备就绪，开始处理图片');
      progressText.value = t('preprocessing');
      progress.value = 30;

      if (isBatchMode.value) {
        // 批处理模式
        batchResults.value = [];
        progressText.value = t('processingBatch');
        
        for (let i = 0; i < imageUrls.value.length; i++) {
          progressText.value = `${t('processingImage')} ${i + 1}/${imageUrls.value.length}`;
          progress.value = (i / imageUrls.value.length) * 100;
          
          try {
            // 预处理图片
            const processedImage = await preprocessImage(imageUrls.value[i]);
            
            // 识别图片
            const { data: { text } } = await tessWorker.recognize(processedImage);
            
            batchResults.value.push({
              index: i + 1,
              text: text || '未识别到文本',
              success: true
            });
          } catch (error) {
            console.error(`Error processing image ${i + 1}:`, error);
            batchResults.value.push({
              index: i + 1,
              text: `识别失败: ${error.message}`,
              success: false
            });
          }
        }
        
        // 合并批处理结果
        resultText.value = batchResults.value
          .map(item => `=== ${t('image')} ${item.index} ===\n${item.text}\n`) 
          .join('');
      } else {
        // 单张图片模式
        try {
          // 预处理图片
          progressText.value = t('preprocessing');
          const processedImage = await preprocessImage(imageUrl.value);
          
          // 识别图片
          progressText.value = t('recognizing');
          const { data: { text } } = await tessWorker.recognize(processedImage);
          
          resultText.value = text || '未识别到文本';
        } catch (error) {
          console.error('Error in single image processing:', error);
          throw error;
        }
      }
      
      // 处理Word格式 - 不再自动下载，由用户手动点击下载按钮
      // if (outputFormat.value === 'word' && resultText.value) {
      //   await downloadWord();
      // }
      
      progressText.value = t('completed');
      progress.value = 100;
    } catch (initError) {
      console.error('Tesseract.js初始化失败:', initError);
      // 清除worker实例，以便下次可以重新尝试初始化
      if (worker.value) {
        try {
          await worker.value.terminate();
        } catch (e) {
          console.error('终止worker失败:', e);
        }
        worker.value = null;
      }
      // 抛出错误，让外层catch处理
      throw new Error(`Tesseract.js初始化失败: ${initError.message}。请检查网络连接后重试。`);
    } finally {
      // 不要在这里终止worker，以便后续重用
      // 只有在组件卸载时才终止worker
    }
  } catch (error) {
    console.error('Error:', error);
    errorMessage.value = `${t('processFailed')}: ${error.message || t('checkNetwork')}`;
  } finally {
    isLoading.value = false;
    // 重置进度
    setTimeout(() => {
      progress.value = 0;
      progressText.value = '';
    }, 1000);
  }
};

// 清理worker实例
onUnmounted(async () => {
  if (worker.value) {
    try {
      await worker.value.terminate();
      console.log('Worker已终止');
    } catch (error) {
      console.error('终止Worker失败:', error);
    } finally {
      worker.value = null; // 确保worker实例被正确清除
    }
  }
});

// 将图片直接转换为PDF
const imageToPdf = async () => {
  if (!imageUrl.value) {
    throw new Error(t('pleaseUpload'));
  }


  // 创建PDF文档
  const doc = new jsPDF();
  
  // 获取图片尺寸
  const img = new Image();
  img.src = imageUrl.value;
  
  await new Promise((resolve, reject) => {
    img.onload = () => {
      // 计算图片在PDF中的尺寸（保持宽高比）
      const pdfWidth = doc.internal.pageSize.getWidth() - 20; // 左右边距各10
      const pdfHeight = doc.internal.pageSize.getHeight() - 20; // 上下边距各10
      
      let imgWidth = img.width;
      let imgHeight = img.height;
      
      // 调整图片尺寸以适应PDF页面
      if (imgWidth > pdfWidth) {
        const ratio = pdfWidth / imgWidth;
        imgWidth = pdfWidth;
        imgHeight = imgHeight * ratio;
      }
      
      if (imgHeight > pdfHeight) {
        const ratio = pdfHeight / imgHeight;
        imgHeight = pdfHeight;
        imgWidth = imgWidth * ratio;
      }
      
      // 计算图片位置（居中）
      const x = (doc.internal.pageSize.getWidth() - imgWidth) / 2;
      const y = (doc.internal.pageSize.getHeight() - imgHeight) / 2;
      
      // 添加图片到PDF
      doc.addImage(img, 'PNG', x, y, imgWidth, imgHeight);
      resolve();
    };
    
    img.onerror = reject;
  });
  
  // 返回doc对象，供downloadPDF函数使用
  return doc;

};

// 图片转Word
const imageToWord = async () => {
  if (!imageUrl.value) {
    throw new Error(t('pleaseUpload'));
  }


  // 创建Word文档，包含图片
  const doc = new Document({
    sections: [{
      properties: {},
      children: [
        new Paragraph({
          text: t('imageConversion'),
          heading: 'heading1',
          spacing: {
            before: 240,
            after: 240
          }
        }),
        new Paragraph({
          text: t('imageConvertedToWord'),
          spacing: {
            before: 0,
            after: 120
          }
        }),
        new Paragraph({
          text: t('conversionTime') + new Date().toLocaleString(),
          spacing: {
            before: 0,
            after: 240
          }
        }),
        // 注意：在浏览器环境中直接嵌入图片到Word文档需要特殊处理
        // 这里我们使用一个更简单的方法，添加图片的描述信息
        new Paragraph({
          text: t('imageInfo'),
          heading: 'heading3',
          spacing: {
            before: 120,
            after: 120
          }
        }),
        new Paragraph({
          text: t('imageFormat') + imageUrl.value.split('.').pop().split('?')[0],
          spacing: {
            before: 0,
            after: 60
          }
        }),
        new Paragraph({
          text: t('conversionMethod') + t('directConversion'),
          spacing: {
            before: 0,
            after: 60
          }
        })
      ]
    }]
  });


  // 生成Word文档但不自动下载
  try {
    Packer.toBlob(doc).then(blob => {
      // 无弹窗提示
    });
  } catch (error) {
    console.error('Error generating Word document:', error);
    alert(t('wordError'));
  }

};

// 复制到剪贴板
const copyToClipboard = () => {
  if (resultText.value) {
    navigator.clipboard.writeText(resultText.value)
      .then(() => {
        alert(t('copiedToClipboard'));
      })
      .catch(() => {
        alert(t('copyFailed'));
      });
  }

};

// 下载文本
const downloadText = () => {
  if (!resultText.value) {
    alert(t('pleaseConvert'));
    return;
  }


  // 创建文本文件
  const blob = new Blob([resultText.value], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'image-to-text.txt';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};

// 下载Excel (.xlsx格式)
const downloadExcel = () => {
  if (!resultText.value) {
    alert(t('pleaseConvert'));
    return;
  }


  // 将文本转换为二维数组，优化表格处理
  const lines = resultText.value.split('\n');
  const data = [];
  
  // 处理每一行，智能识别表格结构
  lines.forEach(line => {
    // 移除行首尾空白
    const trimmedLine = line.trim();
    
    // 如果行为空，添加空行以保持表格结构
    if (!trimmedLine) {
      data.push(['']);
      return;
    }
    
    // 尝试更智能的列分割，优先保持表格结构
    let columns = [];
    
    // 1. 优先按制表符分割（OCR.space表格模式会返回制表符分隔的数据）
    if (trimmedLine.includes('\t')) {
      columns = trimmedLine.split('\t');
    }
    // 2. 对于表格数据，尝试按固定宽度分割
    else if (isTableData(trimmedLine)) {
      columns = splitByFixedWidth(trimmedLine);
    }
    // 3. 按多个空格分割
    else {
      columns = trimmedLine.split(/\s{2,}/);
    }
    
    // 清理列数据
    columns = columns.map(col => col.trim());
    
    // 智能处理行号：检查第一列是否为纯数字行号
    if (columns.length > 0) {
      const firstCol = columns[0];
      // 如果第一列是纯数字，且长度合理（1-5位），可能是行号
      if (/^\d{1,5}$/.test(firstCol)) {
        // 确保行号列独立，不与其他内容混合
        // 这里可以添加更多逻辑来验证行号的连续性
      }
    }
    
    // 添加到数据数组
    if (columns.length > 0) {
      data.push(columns);
    }
  });

  // 后处理：尝试修复行号错位问题
  fixRowNumbering(data);

  // 创建Excel工作簿
  const wb = XLSX.utils.book_new();
  
  // 创建工作表
  const ws = XLSX.utils.aoa_to_sheet(data);
  
  // 自动调整列宽
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let C = range.s.c; C <= range.e.c; ++C) {
    let maxLen = 0;
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const cellAddress = XLSX.utils.encode_cell({c: C, r: R});
      if (ws[cellAddress] && ws[cellAddress].v) {
        maxLen = Math.max(maxLen, ws[cellAddress].v.toString().length);
      }
    }
    // 设置列宽（1字符约等于1.1宽度单位）
    if (!ws['!cols']) ws['!cols'] = [];
    ws['!cols'][C] = {wch: Math.max(10, maxLen * 1.1)};
  }
  
  // 将工作表添加到工作簿
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  
  // 生成Excel文件并下载
  XLSX.writeFile(wb, 'image-to-excel.xlsx');
};

// 修复行号错位问题
const fixRowNumbering = (data) => {
  // 分析表格结构，寻找可能的行号列
  if (data.length < 2) return;
  
  // 检查第一列是否可能是行号列
  const possibleRowNumbers = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i].length > 0) {
      const firstCol = data[i][0];
      if (/^\d{1,5}$/.test(firstCol)) {
        possibleRowNumbers.push({ row: i, value: parseInt(firstCol) });
      }
    }
  }
  
  // 如果找到足够的行号，检查连续性并修复错位
  if (possibleRowNumbers.length > 2) {
    // 按行号值排序
    possibleRowNumbers.sort((a, b) => a.value - b.value);
    
    // 检查行号的连续性和位置
    for (let i = 0; i < possibleRowNumbers.length; i++) {
      const { row, value } = possibleRowNumbers[i];
      
      // 检查下一行是否存在，并且是否以数字开头（可能是错位的行号）
      if (row + 1 < data.length) {
        const nextRow = data[row + 1];
        if (nextRow.length > 0) {
          const nextFirstCol = nextRow[0];
          // 如果下一行的第一列也是数字，可能是行号错位
          if (/^\d{1,5}$/.test(nextFirstCol)) {
            const nextValue = parseInt(nextFirstCol);
            // 检查行号是否连续
            if (nextValue === value + 1) {
              // 行号看起来是连续的，保持现状
            } else {
              // 可能存在错位，检查是否需要调整
              // 这里可以添加更复杂的逻辑
            }
          }
        }
      }
    }
    
    // 特殊处理：检查是否有行号单独占一行的情况
    // 例如：行号"41"在一行，内容在另一行
    for (let i = 0; i < data.length - 1; i++) {
      const currentRow = data[i];
      const nextRow = data[i + 1];
      
      // 如果当前行只有一个元素，且是数字（可能是行号）
      if (currentRow.length === 1 && /^\d{1,5}$/.test(currentRow[0])) {
        // 如果下一行有多个元素（可能是内容）
        if (nextRow.length > 1) {
          // 将行号移动到下一行的开头，合并为一行
          nextRow.unshift(currentRow[0]);
          // 清空当前行
          data[i] = [''];
        }
      }
    }
  }
};

// 判断是否为表格数据
const isTableData = (line) => {
  // 表格数据通常包含多个连续的空格，或者有明显的列对齐特征
  const spacesCount = (line.match(/\s{3,}/g) || []).length;
  return spacesCount >= 2;
};

// 按固定宽度分割表格行
const splitByFixedWidth = (line) => {
  // 寻找可能的列分隔位置（连续空格）
  const separators = [];
  let currentSpaceStart = -1;
  
  for (let i = 0; i < line.length; i++) {
    if (line[i] === ' ' || line[i] === '\t') {
      if (currentSpaceStart === -1) {
        currentSpaceStart = i;
      }
    } else {
      if (currentSpaceStart !== -1 && i - currentSpaceStart >= 2) {
        separators.push(Math.floor((currentSpaceStart + i) / 2));
      }
      currentSpaceStart = -1;
    }
  }
  
  // 如果找到分隔符，按分隔符分割
  if (separators.length > 0) {
    const columns = [];
    let start = 0;
    
    for (const sep of separators) {
      columns.push(line.substring(start, sep).trim());
      start = sep;
    }
    
    columns.push(line.substring(start).trim());
    return columns;
  }
  
  // 否则按多个空格分割
  return line.split(/\s{2,}/);
};

// 下载PDF（保留此函数以兼容现有UI）
const downloadPDF = async () => {
  try {
    // 调用imageToPdf函数获取doc对象
    const doc = await imageToPdf();
    // 保存PDF文件
    doc.save('image-to-pdf.pdf');
  } catch (error) {
    console.error('Error converting image to PDF:', error);
    errorMessage.value = error.message;
  }
};

// 下载Word
const downloadWord = () => {
  if (!resultText.value) {
    alert(t('pleaseConvert'));
    return;
  }


  // 将识别结果转换为Word文档，尽量保持原始排版
  const paragraphs = resultText.value.split('\n');
  const docElements = [];
  
  // 处理每个段落
  paragraphs.forEach(paragraph => {
    const trimmedPara = paragraph.trim();
    if (trimmedPara) {
      // 检查是否可能是标题（长度较短且全大写或首字母大写）
      if (trimmedPara.length < 50 && (trimmedPara === trimmedPara.toUpperCase() || /^[A-Z][a-z0-9\s]+$/.test(trimmedPara))) {
        // 添加标题段落
        docElements.push(
          new Paragraph({
            text: trimmedPara,
            heading: 'heading3',
            spacing: {
              before: 240,
              after: 120
            }
          })
        );
      } else {
        // 检查是否可能是表格数据
        if (trimmedPara.includes('\t') || (trimmedPara.match(/\s{3,}/g) || []).length >= 2) {
          // 尝试转换为简单表格
          const cells = trimmedPara.includes('\t') ? trimmedPara.split('\t') : trimmedPara.split(/\s{3,}/);
          
          // 创建表格行和单元格
          const tableRows = [
            new TableRow({
              children: cells.map(cell => new TableCell({
                children: [new Paragraph(cell.trim())]
              }))
            })
          ];
          
          // 添加表格
          docElements.push(
            new Table({
              rows: tableRows,
              width: {
                size: 100,
                type: WidthType.PERCENTAGE
              },
              spacing: {
                before: 120,
                after: 120
              }
            })
          );
        } else {
          // 普通段落
          docElements.push(
            new Paragraph({
              text: trimmedPara,
              spacing: {
                before: 0,
                after: 120
              }
            })
          );
        }
      }
    } else {
      // 空行
      docElements.push(new Paragraph({ text: '' }));
    }
  });

  // 创建Word文档
  const doc = new Document({
    sections: [{
      properties: {},
      children: docElements
    }]
  });

  // 生成Word文档并触发下载
  try {
    Packer.toBlob(doc).then(blob => {
      // 创建下载链接并触发下载
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'image-to-word.docx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    });
  } catch (error) {
    console.error('Error generating Word document:', error);
    alert('生成Word文档时出错，请重试。');
  }
};

// 清除所有内容
const clearAll = () => {
  imageUrl.value = '';
  imageUrls.value = [];
  resultText.value = '';
  errorMessage.value = '';
  batchResults.value = [];
  isBatchMode.value = false;
  // 重置文件输入框，确保可以再次选择相同的文件
  const fileInput = document.getElementById('image-upload');
  if (fileInput) {
    fileInput.value = '';
  }
};
</script>

<template>
  <div class="min-h-screen bg-gray-50 text-gray-800">
    <!-- 顶部导航 -->
    <header class="bg-white shadow-sm">
      <div class="container mx-auto px-4 py-4">
        <div class="flex justify-between items-center">
          <!-- 标题 -->
          <div class="flex-1">
            <h1 class="text-2xl font-bold text-blue-600">{{ t('title') }}</h1>
            <p class="text-gray-600 mt-1 text-sm">{{ t('subtitle') }}</p>
          </div>
          
          <!-- 语言选择器 -->
          <div class="flex items-center ml-4">
            <label for="language-select" class="mr-1 text-xs text-gray-600">{{ t('language') }}:</label>
            <select 
              id="language-select" 
              v-model="selectedLanguage" 
              @change="changeLanguage(selectedLanguage)"
              class="border border-gray-300 rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <option v-for="lang in languages" :key="lang.code" :value="lang.code">
                {{ lang.name }}
              </option>
            </select>
          </div>
        </div>
      </div>
    </header>

    <!-- 主要内容 -->
    <main class="container mx-auto px-4 py-8">
      <div class="max-w-5xl mx-auto">
        
        
        <!-- 功能介绍 -->
        <div class="bg-white rounded-lg shadow-sm p-4 mb-6">
          <h2 class="text-lg font-semibold mb-3 text-center">{{ t('howToUse') }}</h2>
          <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
            <!-- 步骤1 -->
            <div class="flex items-start">
              <div class="bg-blue-100 rounded-full w-10 h-10 flex items-center justify-center mr-2 flex-shrink-0">
                <span class="text-blue-600 font-bold">1</span>
              </div>
              <div class="text-left">
                <h3 class="font-medium text-sm mb-1">{{ t('step1') }}</h3>
                <p class="text-gray-600 text-xs">{{ t('step1Desc') }}</p>
              </div>
            </div>
            
            <!-- 步骤2 -->
            <div class="flex items-start">
              <div class="bg-blue-100 rounded-full w-10 h-10 flex items-center justify-center mr-2 flex-shrink-0">
                <span class="text-blue-600 font-bold">2</span>
              </div>
              <div class="text-left">
                <h3 class="font-medium text-sm mb-1">{{ t('step2') }}</h3>
                <p class="text-gray-600 text-xs">{{ t('step2Desc') }}</p>
              </div>
            </div>
            
            <!-- 步骤3 -->
            <div class="flex items-start">
              <div class="bg-blue-100 rounded-full w-10 h-10 flex items-center justify-center mr-2 flex-shrink-0">
                <span class="text-blue-600 font-bold">3</span>
              </div>
              <div class="text-left">
                <h3 class="font-medium text-sm mb-1">{{ t('step3') }}</h3>
                <p class="text-gray-600 text-xs">{{ t('step3Desc') }}</p>
              </div>
            </div>
          </div>
        </div>

        <!-- 图片上传和结果区域 -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
          <!-- 图片上传区域 -->
          <div class="bg-white rounded-lg shadow-sm p-6">
            <h2 class="text-lg font-semibold mb-4 text-center">{{ t('uploadImage') }}</h2>
            
            <!-- 拖拽上传区域 -->
            <div 
              class="border-2 border-dashed border-gray-300 rounded-lg p-8 flex flex-col items-center justify-center"
              @dragover="handleDragOver"
              @drop="handleDrop"
            >
              <input 
                type="file" 
                accept="image/*" 
                @change="handleImageUpload" 
                class="hidden" 
                id="image-upload"
              />
              
              <!-- 当有单张图片时显示预览 -->
              <div v-if="imageUrl && !isBatchMode" class="w-full">
                <img 
                  :src="imageUrl" 
                  alt="Preview" 
                  class="w-full h-auto max-h-64 object-contain mx-auto"
                />
              </div>
              
              <!-- 当有多个图片时显示网格预览 -->
              <div v-else-if="imageUrls.length > 0 && isBatchMode" class="w-full">
                <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  <div v-for="(imgUrl, index) in imageUrls" :key="index" class="relative">
                    <img 
                      :src="imgUrl" 
                      :alt="`Preview ${index + 1}`" 
                      class="w-full h-32 object-cover rounded"
                    />
                    <div class="absolute top-1 right-1 bg-blue-500 text-white text-xs px-2 py-1 rounded">
                      {{ index + 1 }}
                    </div>
                  </div>
                </div>
                <p class="text-center text-sm text-gray-500 mt-2">
                  {{ t('batchMode') }}: {{ imageUrls.length }} {{ t('images') }}
                </p>
              </div>
              
              <!-- 当没有图片时显示上传提示 -->
              <label 
                v-else
                for="image-upload" 
                class="cursor-pointer flex flex-col items-center justify-center"
              >
                <svg class="w-16 h-16 text-blue-500 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path>
                </svg>
                <span class="text-gray-600 mb-2">{{ t('dropImage') }}</span>
                <span class="text-gray-400 text-sm">{{ t('imageFormat') }}</span>
              </label>
            </div>
            
            <!-- 输出格式选择 -->
            <div class="mt-4">
              <!-- <label class="block text-sm font-medium text-gray-700 mb-2">{{ t('outputFormat') }}</label> -->
              <div class="flex flex-wrap gap-4">
                <label class="inline-flex items-center">
                  <input 
                    type="radio" 
                    v-model="outputFormat" 
                    value="text" 
                    :disabled="isLoading" 
                    class="form-radio text-blue-600"
                  >
                  <span class="ml-2 text-gray-700">{{ t('text') }}</span>
                </label>
                <label class="inline-flex items-center">
                  <input 
                    type="radio" 
                    v-model="outputFormat" 
                    value="excel" 
                    :disabled="isLoading" 
                    class="form-radio text-blue-600"
                  >
                  <span class="ml-2 text-gray-700">{{ t('excel') }}</span>
                </label>
                <label class="inline-flex items-center">
                  <input 
                    type="radio" 
                    v-model="outputFormat" 
                    value="pdf" 
                    :disabled="isLoading" 
                    class="form-radio text-blue-600"
                  >
                  <span class="ml-2 text-gray-700">{{ t('pdf') }}</span>
                </label>
                <label class="inline-flex items-center">
                  <input 
                    type="radio" 
                    v-model="outputFormat" 
                    value="word" 
                    :disabled="isLoading" 
                    class="form-radio text-blue-600"
                  >
                  <span class="ml-2 text-gray-700">{{ t('word') }}</span>
                </label>
              </div>
            </div>
            
            <!-- 操作按钮 -->
            <div class="mt-6 flex flex-col sm:flex-row gap-3 justify-center">
              <button 
                @click="imageToText" 
                :disabled="isLoading || !imageUrl" 
                class="flex-1 px-6 py-3 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {{ isLoading ? t('processing') : t('startConvert') }}
              </button>
              <button 
                @click="clearAll" 
                :disabled="isLoading" 
                class="flex-1 px-6 py-3 bg-gray-200 text-gray-700 rounded-md hover:bg-gray-300 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {{ t('clear') }}
              </button>
            </div>
            
            <!-- 进度条 -->
            <div v-if="isLoading" class="mt-4">
              <div class="text-sm mb-1 text-center text-gray-600">
                {{ progressText }}
              </div>
              <div class="w-full bg-gray-200 rounded-full h-2 overflow-hidden">
                <div 
                  class="h-2 rounded-full bg-gradient-to-r from-purple-300 via-purple-500 to-purple-700"
                  :style="{ width: progress + '%' }"
                ></div>
              </div>
            </div>
            
            <!-- 错误信息 -->
            <div v-if="errorMessage" class="mt-4 bg-red-50 text-red-600 p-3 rounded-md">
              {{ errorMessage }}
            </div>
          </div>
          
          <!-- 结果区域 -->
          <div class="bg-white rounded-lg shadow-sm p-6">
            <h2 class="text-lg font-semibold mb-4 text-center">{{ t('resultText') }}</h2>
            
            <!-- 结果显示 -->
            <div class="border border-gray-200 rounded-lg p-4 min-h-64 bg-gray-50 mb-4">
              <pre class="whitespace-pre-wrap text-gray-800 font-mono text-sm">{{ resultText || '' }}</pre>
            </div>
            
            <!-- 操作按钮 -->
            <div class="flex flex-col sm:flex-row gap-3">
              <template v-if="outputFormat === 'text'">
                <button 
                  @click="copyToClipboard" 
                  :disabled="!resultText" 
                  class="flex-1 px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {{ t('copyClipboard') }}
                </button>
                <button 
                  @click="downloadText" 
                  :disabled="!resultText" 
                  class="flex-1 px-4 py-2 bg-blue-500 text-white rounded-md hover:bg-blue-600 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {{ t('downloadText') }}
                </button>
              </template>
              <template v-else-if="outputFormat === 'excel'">
                <button 
                  @click="downloadExcel" 
                  :disabled="!resultText" 
                  class="flex-1 px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {{ t('downloadExcel') }}
                </button>
              </template>
              <template v-else-if="outputFormat === 'pdf'">
                <button 
                  @click="downloadPDF" 
                  :disabled="!resultText" 
                  class="flex-1 px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {{ t('downloadPDF') }}
                </button>
              </template>
              <template v-else-if="outputFormat === 'word'">
                <button 
                  @click="downloadWord" 
                  :disabled="!resultText" 
                  class="flex-1 px-4 py-2 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {{ t('downloadWord') }}
                </button>
              </template>
            </div>
          </div>
        </div>
        
        <!-- 广告位预留 -->
        <div class="rounded-lg shadow-sm p-3 mb-6 border border-dashed border-gray-200 mt-8">
          <div class="text-center">
            <p class="text-gray-400 text-sm">暂无内容</p>
          </div>
        </div>
        <!-- 功能特点 -->
        <div class="bg-white rounded-lg shadow-sm p-6 mt-8">
          <h2 class="text-lg font-semibold mb-4 text-center">{{ t('features') }}</h2>
          <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <div class="p-4 text-center">
              <div class="bg-blue-100 rounded-full w-12 h-12 flex items-center justify-center mx-auto mb-2">
                <svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                </svg>
              </div>
              <h3 class="font-medium mb-1">{{ t('feature4') }}</h3>
              <p class="text-gray-600 text-xs">{{ t('feature4Desc') }}</p>
            </div>
            <div class="p-4 text-center">
              <div class="bg-blue-100 rounded-full w-12 h-12 flex items-center justify-center mx-auto mb-2">
                <svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8c-1.657 0-3 .895-3 2s1.343 2 3 2 3 .895 3 2-1.343 2-3 2m0-8c1.11 0 2.08.402 2.599 1M12 8V7m0 1v8m0 0v1m0-1c-1.11 0-2.08-.402-2.599-1M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path>
                </svg>
              </div>
              <h3 class="font-medium mb-1">{{ t('feature1') }}</h3>
              <p class="text-gray-600 text-xs">{{ t('feature1Desc') }}</p>
            </div>
            <div class="p-4 text-center">
              <div class="bg-blue-100 rounded-full w-12 h-12 flex items-center justify-center mx-auto mb-2">
                <svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"></path>
                </svg>
              </div>
              <h3 class="font-medium mb-1">{{ t('feature5') }}</h3>
              <p class="text-gray-600 text-xs">{{ t('feature5Desc') }}</p>
            </div>
            <div class="p-4 text-center">
              <div class="bg-blue-100 rounded-full w-12 h-12 flex items-center justify-center mx-auto mb-2">
                <svg class="w-6 h-6 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z"></path>
                </svg>
              </div>
              <h3 class="font-medium mb-1">{{ t('feature3') }}</h3>
              <p class="text-gray-600 text-xs">{{ t('feature3Desc') }}</p>
            </div>
          </div>
        </div>
      </div>
    </main>

    <!-- 底部信息 -->
    <footer class="bg-white border-t border-gray-200 mt-12">
      <div class="container mx-auto px-4 py-6">
        <p class="text-center text-gray-600 text-sm">{{ t('copyright') }}</p>
      </div>
    </footer>
  </div>
</template>

<style scoped>
/* 自定义样式 */

@keyframes progress {
  0% {
    width: 0%;
  }
  70% {
    width: 100%;
  }
  95% {
    width: 100%;
  }
  100% {
    width: 0%;
  }
}

.animate-progress {
  animation: progress 4s ease-out infinite;
}
</style>