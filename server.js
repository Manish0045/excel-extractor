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
        cb(null, "./uploads/");
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}`);
    }
});
const upload = multer({ storage });

let parsedData = {
    headers: [],
    rows: []
};

app.get("/", (req, res) => {
    res.render('excel', { headers: null, rows: null });
});

app.post('/upload', upload.single('excelFile'), async(req, res) => {
    try {
        const filePath = req.file.path;
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        const worksheet = workbook.worksheets[0];

        // Get headers exactly as they are (including trailing spaces)
        const headers = worksheet.getRow(1).values.slice(1);
        const rows = [];

        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const rowData = {};
            headers.forEach((header, index) => {
                const cell = row.getCell(index + 1);
                let value = cell.value;

                // Process hyperlink or rich text if needed
                if (value && typeof value === 'object') {
                    if (value.hyperlink && value.text) {
                        value = `<a href="${value.hyperlink}" target="_blank">${value.text}</a>`;
                    } else if (value.richText) {
                        value = value.richText.map(r => r.text).join('');
                    } else if (value.result) {
                        value = value.result;
                    } else if (value.text) {
                        value = value.text;
                    } else {
                        value = JSON.stringify(value);
                    }
                }

                rowData[header] = value || '';
            });

            rows.push(rowData);
        });

        fs.unlinkSync(filePath);

        parsedData.headers = headers;
        parsedData.rows = rows;

        res.redirect('/view');
    } catch (error) {
        console.error(error);
        res.send('Error processing file');
    }
});

app.get('/view', (req, res) => {
    const { Final_Status, Country, QA_Status, Job_Titles } = req.query;
    const { headers, rows } = parsedData;

    let filteredRows = rows;

    if (Final_Status && Final_Status !== 'All') {
        filteredRows = filteredRows.filter(r => r['Final_Status'] === Final_Status);
    }
    if (Country && Country !== 'All') {
        filteredRows = filteredRows.filter(r => r['Country'] === Country);
    }
    if (QA_Status && QA_Status !== 'All') {
        filteredRows = filteredRows.filter(r => r['QA_Status'] === QA_Status);
    }

    // Use exact header with trailing space as seen in your Excel
    const jobTitleHeader = 'Job_Title ';

    if (Job_Titles && Job_Titles.trim() !== '') {
        const jobTitleList = Job_Titles
            .split(/\r?\n|,/)
            .map(t => t.trim().toLowerCase())
            .filter(Boolean);

        filteredRows = filteredRows.filter(r =>
            r[jobTitleHeader] &&
            jobTitleList.some(filterTitle =>
                r[jobTitleHeader].toString().toLowerCase().includes(filterTitle)
            )
        );
    }

    const getUnique = (field) => [...new Set(rows.map(r => r[field]).filter(Boolean))];

    res.render('excel', {
        headers,
        rows: filteredRows,
        filters: {
            finalStatusList: getUnique('Final_Status'),
            countryList: getUnique('Country'),
            qaStatusList: getUnique('QA_Status')
        },
        selectedFilters: { Final_Status, Country, QA_Status, Job_Titles }
    });
});

const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
    console.log("🌐 Server Started on port " + PORT);
});