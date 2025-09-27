const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();

app.set('view engine', "ejs");
app.set("views", path.join(__dirname, "views"));

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, "./uploads/")
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}}`);
    }
});

const upload = multer({ storage });

app.get("/", (req, res) => {
    res.render('excel', { headers: null, rows: null });
});

let parsedData = {
    headers: [],
    rows: []
};

app.post('/upload', upload.single('excelFile'), async(req, res) => {
    try {
        const filePath = req.file.path;

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];

        const headers = worksheet.getRow(1).values.slice(1);
        const rows = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const rowData = {};
            headers.forEach((header, index) => {
                rowData[header] = row.getCell(index + 1).value || '';
            });
            rows.push(rowData);
        });

        fs.unlinkSync(filePath);

        // Store in memory
        parsedData.headers = headers;
        parsedData.rows = rows;

        // Redirect to /view
        res.redirect('/view');
    } catch (error) {
        console.error(error);
        res.send('Error processing file');
    }
});

app.get('/view', (req, res) => {
    const { final_Status, Country, QA_Status } = req.query;
    const { headers, rows } = parsedData;

    let filteredRows = rows;

    if (final_Status && final_Status !== 'All') {
        filteredRows = filteredRows.filter(r => r['Final_Status'] === final_Status);
    }
    if (Country && Country !== 'All') {
        filteredRows = filteredRows.filter(r => r['Country'] === Country);
    }
    if (QA_Status && QA_Status !== 'All') {
        filteredRows = filteredRows.filter(r => r['QA_Status'] === QA_Status);
    }

    // Extract unique values for filters
    const getUnique = (field) => [...new Set(rows.map(r => r[field]).filter(Boolean))];

    res.render('excel', {
        headers,
        rows: filteredRows,
        filters: {
            FinalStatusList: getUnique('Final_Status'),
            CountryList: getUnique('Country'),
            QAStatusList: getUnique('QA_Status')
        },
        selectedFilters: { final_Status, Country, QA_Status }
    });
});


const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
    console.log("🌐 Server Started...!");
});