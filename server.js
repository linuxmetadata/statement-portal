// ================= IMPORTS =================
const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const session = require('express-session');
const ExcelJS = require('exceljs');
const archiver = require('archiver');

const app = express();

app.use(express.static(__dirname));
app.use(express.json());

app.use(session({
  secret: 'secret123',
  resave: false,
  saveUninitialized: true
}));

// ================= USERS =================
const allowedUsers = [/* your full list here */,"linuxmeta.data@gmail.com"];

function isAdmin(email){
  return email === "linuxmeta.data@gmail.com";
}

// ================= HELPERS =================
function readData() {
  if (!fs.existsSync('data.json')) return [];
  return JSON.parse(fs.readFileSync('data.json'));
}

function filterByEmail(data, email){
  email = email.toLowerCase();

  return data.filter(row =>
    (row.BH_Email || "").toLowerCase() === email ||
    (row.SM_Email || "").toLowerCase() === email ||
    (row.ZBM_Email || "").toLowerCase() === email ||
    (row.RBM_Email || "").toLowerCase() === email ||
    (row.ABM_Email || "").toLowerCase() === email
  );
}

// ================= STORAGE =================
const storage = multer.diskStorage({
  destination: (req,file,cb)=>{
    if(!req.body.type){
      let dir = "uploads/temp";
      fs.mkdirSync(dir,{recursive:true});
      return cb(null,dir);
    }
    let dir = `uploads/${req.body.state}/${req.body.type}`;
    fs.mkdirSync(dir,{recursive:true});
    cb(null,dir);
  },
  filename:(req,file,cb)=>{
    if(!req.body.type){
      return cb(null,"excel_"+Date.now()+path.extname(file.originalname));
    }
    let name=(req.body.name||"").replace(/[^a-zA-Z0-9]/g,"_");
    cb(null,`${req.body.code}_${name}_${req.body.type}${path.extname(file.originalname)}`);
  }
});
const upload = multer({storage});

const allowedTypes=['.pdf','.xlsx','.xls','.txt','.html','.doc','.docx'];

// ================= LOGIN =================
app.post('/login',(req,res)=>{
  let email=(req.body.email||"").toLowerCase();

  if(!allowedUsers.includes(email)){
    return res.send("Access Denied");
  }

  req.session.user={email};
  res.send("Login success");
});

// ================= GET DATA =================
app.get('/getData',(req,res)=>{
  if(!req.session.user) return res.json({error:"login"});

  let data=readData();

  // ✅ ADMIN FIX
  if(isAdmin(req.session.user.email)){
    return res.json(data);
  }

  res.json(filterByEmail(data,req.session.user.email));
});

// ================= DASHBOARD =================
app.get('/dashboard',(req,res)=>{

  let data = readData();

  // ✅ ADMIN FIX
  if(!isAdmin(req.session.user.email)){
    data = filterByEmail(data, req.session.user.email);
  }

  let total=data.length;
  let sss=data.filter(r=>r.SSS).length;
  let aws=data.filter(r=>r.AWS).length;

  res.json({
    total,
    sss,
    aws,
    pendingSSS: total-sss,
    pendingAWS: total-aws
  });
});

// ================= EXCEL UPLOAD =================
app.post('/uploadExcel',upload.single('file'),(req,res)=>{
  let wb=XLSX.readFile(req.file.path);
  let sheet=wb.Sheets[wb.SheetNames[0]];
  let data=XLSX.utils.sheet_to_json(sheet);

  data=data.map(r=>({...r,Value:"",SSS:false,AWS:false}));

  fs.writeFileSync('data.json',JSON.stringify(data,null,2));
  fs.unlinkSync(req.file.path);

  res.send("Excel uploaded");
});

// ================= SAVE VALUE =================
app.post('/saveValue',(req,res)=>{
  let data=readData();
  data.forEach(r=>{
    if(r.Stockist_Code===req.body.code) r.Value=req.body.value;
  });
  fs.writeFileSync('data.json',JSON.stringify(data,null,2));
  res.send("Saved");
});

// ================= FILE UPLOAD =================
app.post('/uploadFile',upload.single('file'),async(req,res)=>{

  let ext=path.extname(req.file.originalname).toLowerCase();

  if(!allowedTypes.includes(ext)){
    fs.unlinkSync(req.file.path);
    return res.send("Invalid file");
  }

  if(ext==='.pdf'){
    try{
      let d=await pdfParse(fs.readFileSync(req.file.path));
      if(!d.text) throw "";
    }catch{
      fs.unlinkSync(req.file.path);
      return res.send("Invalid PDF");
    }
  }

  let data=readData();

  for(let r of data){
    if(r.Stockist_Code===req.body.code){

      if(!r.Value){
        fs.unlinkSync(req.file.path);
        return res.send("Enter value first");
      }

      if(r[req.body.type]){
        fs.unlinkSync(req.file.path);
        return res.send("Already uploaded");
      }

      r[req.body.type]=true;
    }
  }

  fs.writeFileSync('data.json',JSON.stringify(data,null,2));
  res.send("Uploaded");
});

// ================= DOWNLOAD REPORT =================
app.get('/downloadReport',async(req,res)=>{

  if(!req.session.user || !isAdmin(req.session.user.email)){
    return res.send("Access denied");
  }

  let data=readData();
  let wb=new ExcelJS.Workbook();
  let sheet=wb.addWorksheet("Report");

  let headers=Object.keys(data[0]||{});
  sheet.columns=headers.map(h=>({header:h,key:h,width:20}));

  data.forEach(r=>{
    let row={...r};
    row.SSS=r.SSS?"Done":"Pending";
    row.AWS=r.AWS?"Done":"Pending";

    let added=sheet.addRow(row);

    let sssIndex=headers.indexOf("SSS")+1;
    let awsIndex=headers.indexOf("AWS")+1;

    added.getCell(sssIndex).fill={
      type:'pattern',pattern:'solid',
      fgColor:{argb:row.SSS==="Done"?'FF00FF00':'FFFF0000'}
    };

    added.getCell(awsIndex).fill={
      type:'pattern',pattern:'solid',
      fgColor:{argb:row.AWS==="Done"?'FF00FF00':'FFFF0000'}
    };
  });

  res.setHeader('Content-Disposition','attachment; filename=Report.xlsx');
  await wb.xlsx.write(res);
  res.end();
});

// ================= DOWNLOAD ZIP =================
app.get('/downloadAll',(req,res)=>{

  if(!req.session.user || !isAdmin(req.session.user.email)){
    return res.send("Access denied");
  }

  res.attachment("Files.zip");
  let archive=archiver('zip');
  archive.pipe(res);

  function addDir(dir,zipPath=""){
    if(!fs.existsSync(dir)) return;

    fs.readdirSync(dir).forEach(f=>{
      let full=path.join(dir,f);

      if(fs.statSync(full).isDirectory()){
        addDir(full,path.join(zipPath,f));
      }else{
        archive.file(full,{name:path.join(zipPath,f)});
      }
    });
  }

  addDir("uploads");
  archive.finalize();
});

// ================= START =================
app.get('/',(req,res)=>res.redirect('/login.html'));

const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log("Running"));