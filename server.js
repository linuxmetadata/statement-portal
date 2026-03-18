const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const session = require('express-session');

const app = express();

app.use(express.static(__dirname));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

/* ---------- SAFE FILE READ ---------- */
function readData() {
  if (!fs.existsSync('data.json')) return [];
  const txt = fs.readFileSync('data.json');
  if (!txt.toString().trim()) return [];
  return JSON.parse(txt);
}

/* ---------- STORAGE ---------- */
const storage = multer.diskStorage({
  destination: function (req, file, cb) {

    // Excel upload → temporary
    if (!req.body.type) {
      const dir = path.join(__dirname, 'uploads', 'temp');
      fs.mkdirSync(dir, { recursive: true });
      return cb(null, dir);
    }

    const state = req.body.state || "UNKNOWN";
    const type = req.body.type;

    const dir = path.join(__dirname, 'uploads', state, type);
    fs.mkdirSync(dir, { recursive: true });

    cb(null, dir);
  },

  filename: function (req, file, cb) {

    // Excel upload
    if (!req.body.type) {
      return cb(null, "excel_" + Date.now() + path.extname(file.originalname));
    }

    const code = req.body.code || "NO_CODE";
    const name = (req.body.name || "NO_NAME").replace(/[^a-zA-Z0-9]/g, "_");
    const type = req.body.type;
    const ext = path.extname(file.originalname);

    cb(null, `${code}_${name}_${type}${ext}`);
  }
});

const upload = multer({ storage });

const allowedTypes = ['.pdf','.xlsx','.xls','.txt','.html','.doc','.docx'];

/* ---------- LOGIN ---------- */
app.post('/login', (req, res) => {
  const { email } = req.body;

  const users = JSON.parse(fs.readFileSync('users.json'));
  const user = users.find(u => u.email === email);

  if (!user) return res.send("Invalid user");

  req.session.user = user;
  res.send("Login success");
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.send("Logged out");
});

/* ---------- EXCEL UPLOAD ---------- */
app.post('/uploadExcel', upload.single('file'), (req, res) => {

  const workbook = XLSX.readFile(req.file.path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  let data = XLSX.utils.sheet_to_json(sheet);

  data = data.map(row => ({
    ...row,
    Value: "",
    SSS: false,
    AWS: false
  }));

  fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
  fs.unlinkSync(req.file.path);

  res.send("Excel uploaded successfully");
});

/* ---------- GET DATA ---------- */
app.get('/getData', (req, res) => {

  if (!req.session.user) return res.json({ error: "login" });

  let data = readData();

  if (req.session.user.role !== "admin") {
    data = data.filter(r => r.BH_Email === req.session.user.email);
  }

  res.json(data);
});

/* ---------- SAVE VALUE ---------- */
app.post('/saveValue', (req, res) => {

  const { code, value } = req.body;
  let data = readData();

  data.forEach(row => {
    if (row.Stockist_Code === code) {
      row.Value = value;
    }
  });

  fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
  res.send("Saved");
});

/* ---------- FILE UPLOAD ---------- */
app.post('/uploadFile', upload.single('file'), async (req, res) => {

  const file = req.file;
  const ext = path.extname(file.originalname).toLowerCase();

  if (!allowedTypes.includes(ext)) {
    fs.unlinkSync(file.path);
    return res.send("Invalid file type");
  }

  // PDF validation
  if (ext === '.pdf') {
    try {
      const data = await pdfParse(fs.readFileSync(file.path));
      if (!data.text) throw "Invalid";
    } catch {
      fs.unlinkSync(file.path);
      return res.send("Invalid PDF");
    }
  }

  const { code, type } = req.body;
  let data = readData();

  for (let row of data) {
    if (row.Stockist_Code === code) {

      if (!row.Value) {
        fs.unlinkSync(file.path);
        return res.send("Enter value first");
      }

      if (row[type] === true) {
        fs.unlinkSync(file.path);
        return res.send("Already uploaded");
      }

      row[type] = true;
    }
  }

  fs.writeFileSync('data.json', JSON.stringify(data, null, 2));
  res.send("Uploaded successfully");
});

/* ---------- DASHBOARD ---------- */
app.get('/dashboard', (req, res) => {

  let data = readData();

  let total = data.length;
  let sss = data.filter(r => r.SSS).length;
  let aws = data.filter(r => r.AWS).length;

  res.json({
    total,
    sss,
    aws,
    pendingSSS: total - sss,
    pendingAWS: total - aws
  });
});
const ExcelJS = require('exceljs');

/* ---------- DOWNLOAD REPORT ---------- */
app.get('/downloadReport', async (req, res) => {

  let data = readData();

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Report');

  sheet.columns = [
    { header: 'Code', key: 'code' },
    { header: 'Name', key: 'name' },
    { header: 'Value', key: 'value' },
    { header: 'SSS', key: 'sss' },
    { header: 'AWS', key: 'aws' }
  ];

  data.forEach(r => {
    sheet.addRow({
      code: r.Stockist_Code,
      name: r.Stockist_Name,
      value: r.Value,
      sss: r.SSS ? "Done" : "Pending",
      aws: r.AWS ? "Done" : "Pending"
    });
  });

  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );

  res.setHeader(
    'Content-Disposition',
    'attachment; filename=report.xlsx'
  );

  await workbook.xlsx.write(res);
  res.end();
});
/* ---------- FILE DOWNLOAD ---------- */
app.get('/downloadFile', (req, res) => {

  if (!req.session.user || req.session.user.role !== "admin") {
    return res.send("Not allowed");
  }

  const filePath = req.query.path;

  res.download(filePath);
});

/* ---------- START ---------- */
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log("Server running");
});

