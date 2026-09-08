// routes/export.js
const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');

router.post('/export-formulaire', async (req, res) => {
  try {
    const { projectName, requestor, date, vms = [] } = req.body;

    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Demande VM');

    // Ensure grid lines remain visible
    ws.views = [{ showGridLines: true }];

    // 1. Column Widths Setup
    ws.columns = [
      { width: 4 },   // Column A (Margin)
      { width: 28 },  // Column B (Hostname / Label)
      { width: 16 },  // Column C (CPU / Value)
      { width: 16 },  // Column D (RAM GB)
      { width: 28 },  // Column E (OS)
      { width: 18 }   // Column F (Disk GB)
    ];

    // Styling Palette & Shared Attributes
    const PALETTE = {
      primaryNavy: '1F4E79',
      whiteText: 'FFFFFF',
      accentLight: 'D9E1F2',
      zebraRow: 'F9FAFB',
      borderColor: 'D9D9D9'
    };

    const cellBorder = {
      top: { style: 'thin', color: { argb: PALETTE.borderColor } },
      left: { style: 'thin', color: { argb: PALETTE.borderColor } },
      bottom: { style: 'thin', color: { argb: PALETTE.borderColor } },
      right: { style: 'thin', color: { argb: PALETTE.borderColor } }
    };

    // 2. Header Title Banner
    ws.mergeCells('B2:F3');
    const titleCell = ws.getCell('B2');
    titleCell.value = 'FORMULAIRE DE CRÉATION DE MACHINES VIRTUELLES';
    titleCell.font = { name: 'Calibri', size: 14, bold: true, color: { argb: PALETTE.whiteText } };
    titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

    ['B2', 'C2', 'D2', 'E2', 'F2', 'B3', 'C3', 'D3', 'E3', 'F3'].forEach(coord => {
      ws.getCell(coord).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.primaryNavy } };
    });

    // 3. Metadata Form Fields
    const metaFields = [
      { label: 'Nom du Projet :', value: projectName || '' },
      { label: 'Demandeur :', value: requestor || '' },
      { label: 'Date de Demande :', value: date || '' }
    ];

    let rowCursor = 5;
    metaFields.forEach(field => {
      ws.getRow(rowCursor).height = 22;

      // Label Cell
      const lbl = ws.getCell(`B${rowCursor}`);
      lbl.value = field.label;
      lbl.font = { name: 'Calibri', size: 11, bold: true };
      lbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.accentLight } };
      lbl.alignment = { vertical: 'middle', horizontal: 'left' };
      lbl.border = cellBorder;

      // Value Field (Merged C:F)
      ws.mergeCells(`C${rowCursor}:F${rowCursor}`);
      const val = ws.getCell(`C${rowCursor}`);
      val.value = field.value;
      val.font = { name: 'Calibri', size: 11 };
      val.alignment = { vertical: 'middle', horizontal: 'left' };

      ['C', 'D', 'E', 'F'].forEach(col => {
        ws.getCell(`${col}${rowCursor}`).border = cellBorder;
      });

      rowCursor++;
    });

    rowCursor += 2; // Jump to Row 10 for table headers

    // 4. Data Table Headers
    const headers = ['Hostname', 'CPU (vCPU)', 'RAM (GB)', 'Système d\'exploitation', 'Disque (GB)'];
    const cols = ['B', 'C', 'D', 'E', 'F'];

    ws.getRow(rowCursor).height = 26;
    cols.forEach((col, idx) => {
      const headerCell = ws.getCell(`${col}${rowCursor}`);
      headerCell.value = headers[idx];
      headerCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: PALETTE.whiteText } };
      headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.primaryNavy } };
      headerCell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      headerCell.border = cellBorder;
    });

    const tableDataStartRow = rowCursor + 1;

    // 5. Dynamic Data Rows
    vms.forEach((vm, index) => {
      rowCursor++;
      ws.getRow(rowCursor).height = 20;

      const rowBg = index % 2 === 0 ? 'FFFFFF' : PALETTE.zebraRow;
      const rowValues = [
        vm.hostname || '',
        Number(vm.cpu) || 0,
        Number(vm.ram) || 0,
        vm.os || '',
        Number(vm.disk) || 0
      ];

      cols.forEach((col, idx) => {
        const cell = ws.getCell(`${col}${rowCursor}`);
        cell.value = rowValues[idx];
        cell.font = { name: 'Calibri', size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
        cell.border = cellBorder;
        cell.alignment = {
          vertical: 'middle',
          horizontal: idx === 0 || idx === 3 ? 'left' : 'center'
        };

        if (idx === 1 || idx === 2 || idx === 4) {
          cell.numFmt = '#,##0';
        }
      });
    });

    // 6. Formulas & Total Summary Row
    if (vms.length > 0) {
      rowCursor++;
      const summaryRow = rowCursor;
      ws.getRow(summaryRow).height = 22;

      // Label
      const totalLbl = ws.getCell(`B${summaryRow}`);
      totalLbl.value = 'TOTAL';
      totalLbl.font = { name: 'Calibri', size: 11, bold: true };
      totalLbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.accentLight } };
      totalLbl.alignment = { vertical: 'middle', horizontal: 'right' };
      totalLbl.border = cellBorder;

      // Dynamic CPU Formula Sum
      const totalCpu = ws.getCell(`C${summaryRow}`);
      totalCpu.value = { formula: `SUM(C${tableDataStartRow}:C${summaryRow - 1})` };
      totalCpu.font = { name: 'Calibri', size: 11, bold: true };
      totalCpu.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.accentLight } };
      totalCpu.alignment = { vertical: 'middle', horizontal: 'center' };
      totalCpu.numFmt = '#,##0';
      totalCpu.border = cellBorder;

      // Dynamic RAM Formula Sum
      const totalRam = ws.getCell(`D${summaryRow}`);
      totalRam.value = { formula: `SUM(D${tableDataStartRow}:D${summaryRow - 1})` };
      totalRam.font = { name: 'Calibri', size: 11, bold: true };
      totalRam.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.accentLight } };
      totalRam.alignment = { vertical: 'middle', horizontal: 'center' };
      totalRam.numFmt = '#,##0';
      totalRam.border = cellBorder;

      ['E', 'F'].forEach(col => {
        const emptyCell = ws.getCell(`${col}${summaryRow}`);
        emptyCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: PALETTE.accentLight } };
        emptyCell.border = cellBorder;
      });
    }

    // Stream Response
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="Formulaire_Creation_VM.xlsx"');

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Error constructing Excel file:', error);
    res.status(500).json({ error: 'Failed to build Excel document.' });
  }
});

module.exports = router;