// exports.js

// 内部で利用する定数・ヘルパー関数
const ROWS_PER_PAGE = 45;

const getCellState = (grid, r, c) => {
  const DEFAULT_CELL = { 
    text: '', bold: false, underline: false, strikethrough: false, 
    align: 'left', vAlign: 'middle', borders: { t:0, r:0, b:0, l:0 }, 
    rowSpan: 1, colSpan: 1, hidden: false, fontSize: 14, 
    color: '#000000', bgColor: 'transparent', errorMsg: null 
  };
  return grid[`${r}-${c}`] || DEFAULT_CELL;
};

/**
 * Excel出力処理
 */
const exportToExcel = async ({ gridData, colWidths, rowHeights, pageCount }) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('翻訳データ');

  const totalRows = pageCount * ROWS_PER_PAGE;
  const totalCols = colWidths.length;

  worksheet.columns = colWidths.map(w => ({ width: w / 7 }));

  for (let r = 0; r < totalRows; r++) {
    const row = worksheet.getRow(r + 1);
    row.height = (rowHeights[r] || 15) * 0.75;

    for (let c = 0; c < totalCols; c++) {
      const cellState = getCellState(gridData, r, c);
      if (cellState.hidden) continue;

      const excelCell = row.getCell(c + 1);
      excelCell.value = stripHtml(cellState.text);
      excelCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
      excelCell.font = { name: 'MS Mincho', size: 10 };

      if ((cellState.colSpan && cellState.colSpan > 1) || (cellState.rowSpan && cellState.rowSpan > 1)) {
        worksheet.mergeCells(
          r + 1, 
          c + 1, 
          r + (cellState.rowSpan || 1), 
          c + (cellState.colSpan || 1)
        );
      }
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `TransLayer_Export_${Date.now()}.xlsx`;
  a.click();
};

/**
 * Word出力処理
 */
const exportToWord = async ({ gridData, colWidths, rowHeights, pageCount, fontFamily }) => {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, VerticalAlign, UnderlineType, HeightRule } = docx;
  
  const totalCols = colWidths.length;

  const getBorder = (b) => {
    if (!b) return { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
    return { style: BorderStyle.SINGLE, size: b * 4, color: "000000" };
  };

  const currentFont = fontFamily.split(',')[0].replace(/['"]/g, '').trim() || "MS Mincho";
  const sections = [];

  for (let p = 0; p < pageCount; p++) {
    const tableRows = [];
    const startRow = p * ROWS_PER_PAGE;
    const endRow = startRow + ROWS_PER_PAGE;

    for (let r = startRow; r < endRow; r++) {
      const cells = [];
      for (let c = 0; c < totalCols; c++) {
        const cellState = getCellState(gridData, r, c);
        if (cellState.hidden) continue;

        const textLines = stripHtml(cellState.text).split('\n');
        const paragraphs = textLines.map(line => 
          new Paragraph({
            alignment: cellState.align === 'center' ? AlignmentType.CENTER : cellState.align === 'right' ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { before: 0, after: 0, line: 240, lineRule: "auto" },
            children: [new TextRun({
              text: line,
              size: (cellState.fontSize || 14) * 2,
              font: currentFont,
              bold: cellState.bold,
              strike: cellState.strikethrough,
              underline: cellState.underline ? { type: UnderlineType.SINGLE } : undefined,
              color: cellState.color ? cellState.color.replace('#', '') : "000000"
            })],
          })
        );

        if (paragraphs.length === 0) {
          paragraphs.push(new Paragraph({ spacing: { before: 0, after: 0 }, children: [] }));
        }

        cells.push(new TableCell({
          children: paragraphs,
          columnSpan: cellState.colSpan || 1,
          rowSpan: cellState.rowSpan || 1,
          width: { size: colWidths[c] * 15, type: WidthType.DXA },
          verticalAlign: cellState.vAlign === 'top' ? VerticalAlign.TOP : cellState.vAlign === 'bottom' ? VerticalAlign.BOTTOM : VerticalAlign.CENTER,
          shading: cellState.bgColor && cellState.bgColor !== 'transparent' ? { fill: cellState.bgColor.replace('#', '') } : undefined,
          borders: {
            top: getBorder(cellState.borders.t),
            bottom: getBorder(cellState.borders.b),
            left: getBorder(cellState.borders.l),
            right: getBorder(cellState.borders.r),
          },
          margins: { top: 40, bottom: 40, left: 60, right: 60 }
        }));
      }
      if (cells.length > 0) {
        tableRows.push(new TableRow({ 
          children: cells,
          height: { value: (rowHeights[r] || 15) * 15, rule: HeightRule.EXACT }
        }));
      }
    }

    sections.push({
      properties: { 
        page: { 
          size: { width: 11906, height: 16838 },
          // ★重要修正: Word出力時の余白をなくす（marginを削りheader/footerを0にする）
          margin: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } 
        } 
      },
      children: [
        new Table({
          rows: tableRows,
          width: { size: colWidths.reduce((a, b) => a + b, 0) * 15, type: WidthType.DXA },
          layout: docx.TableLayoutType.FIXED
        })
      ],
    });
  }

  const doc = new Document({
    sections: sections
  });

  const blob = await Packer.toBlob(doc);
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `TransLayer_Export_${Date.now()}.docx`;
  a.click();
};

/**
 * PDFプレビュー画像生成処理
 * ※React側でズームリセット等のUI準備をした後に呼び出す
 */
const generatePdfPreview = async ({ pageCount, guideBox, exportBg }) => {
  const images = [];
  
  for (let i = 0; i < pageCount; i++) {
    const captureContainer = document.getElementById(`capture-container-page-${i}`); 
    if (!captureContainer) continue;
    
    const canvas = await html2canvas(captureContainer, { 
      scale: 2,
      useCORS: true, 
      backgroundColor: '#ffffff', 
      x: guideBox.x, 
      y: guideBox.y,
      width: guideBox.w, 
      height: guideBox.h,
      scrollX: 0,
      scrollY: 0,
      onclone: (clonedDoc) => {
        const container = clonedDoc.getElementById(`capture-container-page-${i}`);
        if (container) {
          container.style.transform = 'none';
          const allElements = container.querySelectorAll('*');
          allElements.forEach(el => {
            el.style.textRendering = 'auto';
            el.style.fontVariantLigatures = 'none';
          });
        }
        if (!exportBg) {
          const bgImages = clonedDoc.querySelectorAll('img[alt^="背景ページ"]');
          bgImages.forEach(img => {
            img.style.display = 'none';
          });
        }
      }
    }); 
    
    images.push(canvas.toDataURL('image/jpeg', 0.9));
  }
  
  return images;
};

/**
 * 提出パック（PDF）生成処理
 * ※React側でズームリセット等のUI準備をした後に呼び出す
 */
const generatePackPdf = async ({ pageCount, guideBox, exportBg, originalPdfBytes, packData, packCompression }) => {
  const { PDFDocument } = window.PDFLib;
  const mergedPdf = await PDFDocument.create();
  const a4Width = 595.28;
  const a4Height = 841.89;

  if (originalPdfBytes) {
    const originalPdf = await PDFDocument.load(originalPdfBytes);
    const copiedPages = await mergedPdf.copyPages(originalPdf, originalPdf.getPageIndices());
    copiedPages.forEach(page => mergedPdf.addPage(page));
  }

  for (let i = 0; i < pageCount; i++) {
    const captureContainer = document.getElementById(`capture-container-page-${i}`);
    if (!captureContainer) continue;
    
    const transCanvas = await html2canvas(captureContainer, { 
      scale: 2,
      useCORS: true, 
      backgroundColor: '#ffffff', 
      x: guideBox.x, 
      y: guideBox.y,
      width: guideBox.w, 
      height: guideBox.h,
      scrollX: 0,
      scrollY: 0,
      onclone: (clonedDoc) => {
        const container = clonedDoc.getElementById(`capture-container-page-${i}`);
        if (container) {
          container.style.transform = 'none';
          const allElements = container.querySelectorAll('*');
          allElements.forEach(el => {
            el.style.textRendering = 'auto';
            el.style.fontVariantLigatures = 'none';
          });
        }
        if (!exportBg) {
          const bgImages = clonedDoc.querySelectorAll('img[alt^="背景ページ"]');
          bgImages.forEach(img => {
            img.style.display = 'none';
          });
        }
      }
    });

    const transImageBytes = transCanvas.toDataURL('image/jpeg', packCompression);
    const transImage = await mergedPdf.embedJpg(transImageBytes);
    const transPage = mergedPdf.addPage([a4Width, a4Height]);
    transPage.drawImage(transImage, { x: 0, y: 0, width: a4Width, height: a4Height });
  }
  
  const pdfBytes = await mergedPdf.save();
  const blob = new Blob([pdfBytes], { type: 'application/pdf' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `提出パック_${packData.documentName || '未設定'}_${Date.now()}.pdf`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};
