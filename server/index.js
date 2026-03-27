const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// -------------------------------------------------------
// Helper: copy style from one cell to another
// -------------------------------------------------------
function copyCellStyle(srcCell, dstCell) {
  if (!srcCell) return;
  try {
    if (srcCell.style) dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
    if (srcCell.font) dstCell.font = { ...srcCell.font };
    if (srcCell.fill) dstCell.fill = { ...srcCell.fill };
    if (srcCell.border) dstCell.border = { ...srcCell.border };
    if (srcCell.alignment) dstCell.alignment = { ...srcCell.alignment };
  } catch (_) {}
}

app.post('/api/export', async (req, res) => {
  try {
    const { formData } = req.body;
    const templatePath = path.join(__dirname, '../../FlowLite_Fixed_Asset_Solar_Template.xlsx');

    // ----------------------------------------------------------
    // 1. Read the template to extract headers/structure only
    // ----------------------------------------------------------
    const templateWb = new ExcelJS.Workbook();
    await templateWb.xlsx.readFile(templatePath);

    // ----------------------------------------------------------
    // 2. Build a BRAND NEW workbook — no legacy data
    // ----------------------------------------------------------
    const newWb = new ExcelJS.Workbook();

    // Sheets to skip entirely (Budget removed per user request)
    const SKIP_SHEETS = ['Budget'];

    // Number of header rows per sheet (rows we preserve from template)
    const HEADER_ROWS = {
      'Location': 2,
      'Inverter': 2,
      'SCB': 2,
      'Tracker': 2,
      'Inverter Transformer': 1,
      'Meter': 1,
      'WMS': 1,
      'Other': 2,
      'Incomer': 1,
      'ICOG': 1,
      'HT Panel': 1,
    };

    // Copy each sheet from template, preserving only headers
    for (const tws of templateWb.worksheets) {
      if (SKIP_SHEETS.includes(tws.name)) continue;

      const nws = newWb.addWorksheet(tws.name, {
        properties: tws.properties,
        pageSetup: tws.pageSetup,
      });

      // Copy column widths
      tws.columns.forEach((col, idx) => {
        if (col.width) nws.getColumn(idx + 1).width = col.width;
      });

      // Copy only header rows
      const headerCount = HEADER_ROWS[tws.name] ?? 1;
      for (let r = 1; r <= headerCount; r++) {
        const srcRow = tws.getRow(r);
        const dstRow = nws.getRow(r);
        srcRow.eachCell({ includeEmpty: true }, (srcCell, colNum) => {
          const dstCell = dstRow.getCell(colNum);
          dstCell.value = srcCell.value;
          copyCellStyle(srcCell, dstCell);
        });
        dstRow.height = srcRow.height;
      }
    }

    // Convenience getter for the new workbook's sheets
    const getSheet = (name) => {
      let ws = newWb.getWorksheet(name);
      if (!ws) {
        ws = newWb.worksheets.find(w => w.name.toLowerCase().includes(name.toLowerCase()));
      }
      return ws;
    };

    // ----------------------------------------------------------
    // 3. Write user data into the clean sheets
    // ----------------------------------------------------------

    // Location Sheet (data starts row 3)
    const locSheet = getSheet('Location');
    if (locSheet && formData.locations) {
      formData.locations.forEach((loc, i) => {
        locSheet.getRow(i + 3).getCell(1).value = loc;
      });
    }

    // Inverter Sheet (data starts row 3)
    const invSheet = getSheet('Inverter');
    if (invSheet && formData.sheetData.inverter) {
      formData.sheetData.inverter.forEach((row, i) => {
        const xlRow = invSheet.getRow(i + 3);
        xlRow.getCell(1).value = row.inverterName;
        xlRow.getCell(2).value = row.location;
        xlRow.getCell(3).value = row.inverterType;
        xlRow.getCell(4).value = row.oem;
        xlRow.getCell(5).value = row.commissioningDate || null;
        xlRow.getCell(6).value = row.latitude;
        xlRow.getCell(7).value = row.longitude;
        xlRow.getCell(8).value = row.status;
        xlRow.getCell(9).value = row.dcCapacity ? Number(row.dcCapacity) : null;
        xlRow.getCell(10).value = row.acCapacity ? Number(row.acCapacity) : null;
        xlRow.getCell(11).value = row.totalModules ? Number(row.totalModules) : null;
      });
    }

    // SCB Sheet — generate string IDs (data starts row 3)
    const scbSheet = getSheet('SCB');
    if (scbSheet && formData.sheetData.scb) {
      let rowIndex = 3;
      formData.sheetData.scb.forEach(row => {
        if (!row.location || !row.inverterName || !row.scbQty || !row.maxStrings) return;
        for (let s = 1; s <= row.scbQty; s++) {
          for (let str = 1; str <= row.maxStrings; str++) {
            const stringName = `${row.location}_${row.inverterName}_SCB${s}_STRING${str}`;
            const xlRow = scbSheet.getRow(rowIndex++);
            xlRow.getCell(1).value = row.location;
            xlRow.getCell(2).value = row.inverterName;
            xlRow.getCell(3).value = row.scbQty;
            xlRow.getCell(4).value = row.maxStrings;
            xlRow.getCell(5).value = stringName;
          }
        }
      });
      // Add "Generated String Name" label to E2
      scbSheet.getCell('E2').value = 'Generated String Name';
    }

    // Tracker Sheet (data starts row 3)
    const trackerSheet = getSheet('Tracker');
    if (trackerSheet && formData.sheetData.tracker) {
      formData.sheetData.tracker.forEach((row, i) => {
        const xlRow = trackerSheet.getRow(i + 3);
        xlRow.getCell(1).value = row.equipmentName;
        xlRow.getCell(2).value = row.equipmentType;
        xlRow.getCell(3).value = row.latitude;
        xlRow.getCell(4).value = row.longitude;
        xlRow.getCell(5).value = row.oem;
        xlRow.getCell(6).value = row.qty ? Number(row.qty) : null;
        xlRow.getCell(7).value = row.location;
      });
    }

    // Inverter Transformer (data starts row 2)
    const itSheet = getSheet('Inverter Transformer');
    if (itSheet && formData.sheetData.inverterTransformer) {
      formData.sheetData.inverterTransformer.forEach((row, i) => {
        const xlRow = itSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.inverterTransformerName;
        xlRow.getCell(2).value = row.location;
      });
    }

    // Meter (data starts row 2)
    const meterSheet = getSheet('Meter');
    if (meterSheet && formData.sheetData.meter) {
      formData.sheetData.meter.forEach((row, i) => {
        const xlRow = meterSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.location;
        xlRow.getCell(2).value = row.meterName;
        xlRow.getCell(3).value = row.meterType;
      });
    }

    // WMS (data starts row 2)
    const wmsSheet = getSheet('WMS');
    if (wmsSheet && formData.sheetData.wms) {
      formData.sheetData.wms.forEach((row, i) => {
        const xlRow = wmsSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.sensorName;
        xlRow.getCell(2).value = row.location;
        xlRow.getCell(3).value = row.oem;
        xlRow.getCell(4).value = row.actualData;
        xlRow.getCell(5).value = row.gridCorrected;
      });
    }

    // Other (data starts row 3)
    const otherSheet = getSheet('Other');
    if (otherSheet && formData.sheetData.other) {
      formData.sheetData.other.forEach((row, i) => {
        const xlRow = otherSheet.getRow(i + 3);
        xlRow.getCell(1).value = row.equipmentName;
        xlRow.getCell(2).value = row.equipmentType;
        xlRow.getCell(3).value = row.location;
        xlRow.getCell(4).value = row.oem;
        xlRow.getCell(5).value = row.commissioningDate || null;
        xlRow.getCell(6).value = row.latitude;
        xlRow.getCell(7).value = row.longitude;
        xlRow.getCell(8).value = row.status;
        xlRow.getCell(9).value = row.quantity ? Number(row.quantity) : null;
        xlRow.getCell(10).value = row.unitOfMeasurement;
      });
    }

    // Incomer (data starts row 2)
    const incomerSheet = getSheet('Incomer');
    if (incomerSheet && formData.sheetData.incomer) {
      formData.sheetData.incomer.forEach((row, i) => {
        const xlRow = incomerSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.incomerName;
        xlRow.getCell(2).value = row.location;
      });
    }

    // ICOG (data starts row 2)
    const icogSheet = getSheet('ICOG');
    if (icogSheet && formData.sheetData.icog) {
      formData.sheetData.icog.forEach((row, i) => {
        const xlRow = icogSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.icogName;
        xlRow.getCell(2).value = row.location;
      });
    }

    // HT Panel (data starts row 2)
    const htPanelSheet = getSheet('HT Panel');
    if (htPanelSheet && formData.sheetData.htPanel) {
      formData.sheetData.htPanel.forEach((row, i) => {
        const xlRow = htPanelSheet.getRow(i + 2);
        xlRow.getCell(1).value = row.htPanelName;
        xlRow.getCell(2).value = row.location;
      });
    }

    // ----------------------------------------------------------
    // 4. Write and send
    // ----------------------------------------------------------
    const outputFileName = `Generated_Asset_Config_${Date.now()}.xlsx`;
    const outputPath = path.join(__dirname, outputFileName);
    await newWb.xlsx.writeFile(outputPath);

    res.download(outputPath, 'Final_Asset_Template.xlsx', (err) => {
      if (err) console.error('Error downloading file:', err);
      try { fs.unlinkSync(outputPath); } catch (_) {}
    });

  } catch (error) {
    console.error('Export Error:', error);
    res.status(500).json({ error: error.message || 'Failed to generate Excel file' });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});
