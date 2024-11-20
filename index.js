const express = require('express');
const ejs = require('ejs');
const { Sequelize, DataTypes } = require('sequelize');
const mssql = require('mssql');
const Stimulsoft = require('stimulsoft-reports-js');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const bodyParser = require('body-parser');



const app = express();
const port = 4000;

// Middleware untuk mem-parsing JSON
app.use(bodyParser.json());

// Konfigurasi database
const sequelize = new Sequelize('CompanyDB', 'sa', 'mulzahrian', {
  host: 'localhost',
  dialect: 'mssql'
});

// Middleware untuk memvalidasi template
function validateTemplate(req, res, next) {
  const filename = req.params.filename;
  req.templatePath = path.join(__dirname, 'templates', filename);

  if (!fs.existsSync(req.templatePath) || path.extname(req.templatePath) !== '.mrt') {
    return res.status(400).send('File template tidak ditemukan atau format tidak valid.');
  }
  next();
}

// Fungsi untuk memuat template dan mengambil data
async function loadTemplateAndData(templatePath, filter) {
  const report = new Stimulsoft.Report.StiReport();
  report.loadFile(templatePath);

  const dataSources = report.dictionary.dataSources.list;
  if (!dataSources || dataSources.length === 0) {
    throw new Error('No data sources found in the template.');
  }

  const originalSqlCommand = dataSources[0].sqlCommand;

  // Tambahkan filter ke SQL Command
  const modifiedSqlCommand = filter ? `${originalSqlCommand} ${filter}` : originalSqlCommand;
  dataSources[0].sqlCommand = modifiedSqlCommand;

  const [results] = await sequelize.query(modifiedSqlCommand);

  return { report, data: results, alias: dataSources[0].alias };
}

// Fungsi untuk menghasilkan laporan
async function generateReport(report, data, alias, format) {
  return new Promise((resolve, reject) => {
    report.regData(alias, data);
    report.renderAsync(() => {
      switch (format) {
        case 'pdf':
          report.exportDocumentAsync((pdfData) => {
            const filePath = path.join(__dirname, 'laporan.pdf');
            Stimulsoft.System.StiObject.saveAs(pdfData, filePath, 'application/pdf');
            resolve(filePath);
          }, Stimulsoft.Report.StiExportFormat.Pdf);
          break;
        case 'html':
          const htmlData = report.exportDocument(Stimulsoft.Report.StiExportFormat.Html);
          resolve(htmlData);
          break;
        case 'excel':
          report.exportDocumentAsync((excelData) => {
            const filePath = path.join(__dirname, 'laporan.xlsx');
            fs.writeFileSync(filePath, excelData);
            resolve(filePath);
          }, Stimulsoft.Report.StiExportFormat.Excel2007);
          break;
        default:
          reject(new Error('Format ekspor tidak valid'));
      }
    }, (error) => reject(error));
  });
}

// Rute untuk melihat laporan
app.post('/view/:filename', validateTemplate, async (req, res) => {
  const filter = req.body.filter; // Ambil filter dari request body

  try {
    const { report, data, alias } = await loadTemplateAndData(req.templatePath, filter);
    const htmlData = await generateReport(report, data, alias, 'html');
    res.setHeader('Content-Type', 'text/html');
    res.send(htmlData);
  } catch (error) {
    console.error('Error generating view:', error);
    res.status(500).send('Error generating view');
  }
});

// Rute untuk mengunduh laporan
app.post('/download/:format/:filename', validateTemplate, async (req, res) => {
  const format = req.params.format;
  const filter = req.body.filter; // Ambil filter dari request body

  try {
    const { report, data, alias } = await loadTemplateAndData(req.templatePath, filter);
    const filePath = await generateReport(report, data, alias, format);
    res.download(filePath);
  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).send('Error generating report');
  }
});

// Konfigurasi Multer untuk menyimpan file di folder "templates" dengan nama asli
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, 'templates')); // Folder tujuan
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname); // Menggunakan nama file asli
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    // Validasi hanya menerima file dengan ekstensi .mrt
    if (path.extname(file.originalname) !== '.mrt') {
      return cb(new Error('Hanya file .mrt yang diizinkan!'));
    }
    cb(null, true);
  }
});

// API untuk mengunggah file .mrt
app.post('/upload', upload.single('template'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('File tidak ditemukan atau format tidak valid.');
    }
    
    // Berikan respons berhasil
    res.status(200).send(`File berhasil diunggah: ${req.file.originalname}`);
  } catch (error) {
    console.error('Error uploading file:', error);
    res.status(500).send('Error uploading file');
  }
});


app.listen(port, () => {
    console.log(`Server listening on port ${port}`);
});