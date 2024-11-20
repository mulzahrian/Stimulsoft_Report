const express = require('express');
const ejs = require('ejs');
const { Sequelize, DataTypes } = require('sequelize');
const mssql = require('mssql');
const Stimulsoft = require('stimulsoft-reports-js');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const multer = require('multer');


const app = express();
const port = 4000;

// Konfigurasi database
const sequelize = new Sequelize('CompanyDB', 'sa', 'mulzahrian', {
    host: 'localhost',
    dialect: 'mssql'
  });

// Fungsi untuk mengambil data dari database
async function getData() {
    const templatePath = 'FixedQuery.mrt'
    const report = new Stimulsoft.Report.StiReport();
    report.loadFile(templatePath);

    const dataSources = report.dictionary.dataSources.list;
    if (!dataSources || dataSources.length === 0) {
      throw new Error('No data sources found in the template.');
    }

    const filter = 'and o.id = 1'
  
    // Ambil SqlCommand dari data source pertama
    const sqlCommand = dataSources[0].sqlCommand;
    const regData = dataSources[0].alias;
    console.log('ini ada sorucenya :',regData);
    const newCommand = sqlCommand + filter;
    console.log("ini querynya :", newCommand);


    const [results, metadata] = await sequelize.query(sqlCommand);
    return results;
  }

// Fungsi untuk membuat dan menyimpan laporan dalam format PDF atau Excel
async function generateAndSaveReport(data, format, templatePath = 'FixedQuery.mrt') {
    return new Promise((resolve, reject) => {
      const filter = 'and o.id = 4';

      const report = new Stimulsoft.Report.StiReport();  
      // Load template
      report.loadFile(templatePath);
  
      const dataSources = report.dictionary.dataSources.list;

    if (!dataSources || dataSources.length === 0) {
      throw new Error('No data sources found in the template.');
    }

    // Modifikasi SqlCommand
    const originalSqlCommand = dataSources[0].sqlCommand;
    const modifiedSqlCommand = `${originalSqlCommand} ${filter}`;

    // Update SqlCommand
    dataSources[0].sqlCommand = modifiedSqlCommand;
    const regData = dataSources[0].alias;

    console.log('Modified SqlCommand:', dataSources[0].sqlCommand);

      // Register data
      report.regData(regData, data); // Sesuaikan dengan nama data source dalam template
  
      // Render report
      report.renderAsync((renderedReport) => {
        if (renderedReport) {
          switch (format) {
            case 'pdf':
              report.exportDocumentAsync(
                (pdfData) => {
                  const pdfPath = path.join(__dirname, 'laporan.pdf');
                  Stimulsoft.System.StiObject.saveAs(pdfData, pdfPath, 'application/pdf');
                  console.log(`Laporan berhasil disimpan sebagai ${pdfPath}`);
                  resolve(pdfPath);
                },
                Stimulsoft.Report.StiExportFormat.Pdf
              );
              break;
            case 'html':
              const reportHtml = renderedReport.getHtml();
              resolve(reportHtml);
              break;
            case 'excel':
              report.exportDocumentAsync(
                (excelData) => {
                  const workbook = xlsx.read(excelData, { type: 'buffer' });
                  const excelPath = path.join(__dirname, 'laporan.xlsx');
                  xlsx.writeFile(workbook, excelPath);
                  console.log(`Laporan berhasil disimpan sebagai ${excelPath}`);
                  resolve(excelPath);
                },
                Stimulsoft.Report.StiExportFormat.Excel2007
              );
              break;
            default:
              reject(new Error('Format ekspor tidak valid'));
          }
        } else {
          console.error('Report rendering failed: No rendered report object');
          reject(new Error('Report rendering failed'));
        }
      }, (error) => {
        console.error('Report rendering error:', error);
        reject(error);
      });
    });
  }

// Rute untuk mengunduh laporan dalam format PDF atau Excel
app.get('/download/:format', async (req, res) => {
  const format = req.params.format;
  try {
    const data = await getData();
    const filePath = await generateAndSaveReport(data, format);
    res.download(filePath);
  } catch (error) {
    console.error('Error generating report:', error);
    res.status(500).send('Error generating report');
  }
});

// API untuk menampilkan laporan sebagai HTML
app.get('/view', async (req, res) => {
  try {
    const templatePath = 'FixedQuery.mrt'; // Path ke template .mrt
    const report = new Stimulsoft.Report.StiReport();
    
    // Load template
    report.loadFile(templatePath);

    // Ambil data dari database
    const data = await getData();

    // Registrasi data ke dalam template
    report.regData('getEmployee', data); // Sesuaikan dengan nama data source dalam template

    // Render report
    report.renderAsync(() => {
      // Ekspor laporan ke format HTML
      const reportHtml = report.exportDocument(Stimulsoft.Report.StiExportFormat.Html);
      res.setHeader('Content-Type', 'text/html');
      res.send(reportHtml);
    }, (error) => {
      console.error('Error rendering report:', error);
      res.status(500).send('Error rendering report');
    });
  } catch (error) {
    console.error('Error generating view:', error);
    res.status(500).send('Error generating view');
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