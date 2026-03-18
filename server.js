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

// ================= USER LIST =================
const allowedUsers = [
"brijesh.tiwari@linuxlaboratories.in",
"navin.kumar@linuxlaboratories.in",
"rajatnarayankarmokar@linuxlaboratories.in",
"kirandedhia@linuxlaboratories.in",
"jose.jacob@linuxlaboratories.in",
"manojkumar.patil@linuxlaboratories.in",
"devanshushah@linuxlaboratories.in",
"nandakishore.babu@linuxlaboratories.in",
"tarungohe@linuxlaboratories.in",
"randeepsingh_meta@linuxlaboratories.in",
"pankajrathor_meta@linuxlaboratories.in",
"niraj.barochia_meta@linuxlaboratories.in",
"harish.k.r@linuxlaboratories.in",
"rajugupta_meta@linuxlaboratories.in",
"bhawanishankar@linuxlaboratories.in",
"murugesanpalaniyappan@linuxlaboratories.in",
"gnanaprakash_meta@linuxlaboratories.in",
"debu.meta@linuxlaboratories.in",
"pasupuletivijay1986@gmail.com",
"nanhe.bhartendu@gmail.com",
"ramprajapati2007@ggmail.com",
"honeyvrm6@gmail.com",
"ssamani151@gmail.com",
"prabhatdwivedi19@gmail.com",
"shashimaahi@gmail.com",
"rajbahadurpatel172@gmail.com",
"pathanimrankhan051@gmail.com",
"santhoshkmr05@gmail.com",
"faze73@gmail.com",
"arundas.tinkufipzinda@gmail.com",
"vinitpy@gmail.com",
"vpvikaspatel163@gmail.com",
"raj.bvy@gmail.com",
"mohd786azamkhan@gmail.com",
"vishalrajbhar@rediffmail.com",
"ranjankumardalai02@gmail.com",
"goldysngh44@gmail.com",
"amankumarshridhar@gmail.com",
"vikeykamodiya421995@gmail.com",
"rajsinghajmer5@gmail.com",
"jagdish5586@gmail.com",
"nilsdesmukh@gmail.com",
"nilendrakathar@gmail.com",
"roshan.samarth3292@gmail.com",
"aamulraj2011@gmail.com",
"karthick1987.venkatesh@gmail.com",
"sri_1410@yahoo.com",
"vasanthanila.143@gmail.com",
"kumarswamy.kukkla@gmail.com",
"prabhu.chinna54@gmail.com",
"durgeshdubey1880@gmail.com",
"rathorneerajkumar@gmail.com",
"sonu.singhfmt@gmail.com",
"shovanghosh92@gmail.com",
"arunabha1981gon@gmail.com",
"niladrighatak1979@gmail.com",
"linuxmeta.data@gmail.com"
];

// ================= HELPERS =================
function isAdmin(email){
  return email === "linuxmeta.data@gmail.com";
}

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

// ================= DATA =================
app.get('/getData',(req,res)=>{
  if(!req.session.user) return res.json({error:"login"});
  let data=readData();
  res.json(filterByEmail(data,req.session.user.email));
});

// ================= DASHBOARD =================
app.get('/dashboard',(req,res)=>{
  let data=filterByEmail(readData(),req.session.user.email);
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

// ================= EXCEL =================
app.post('/uploadExcel',upload.single('file'),(req,res)=>{
  let wb=XLSX.readFile(req.file.path);
  let sheet=wb.Sheets[wb.SheetNames[0]];
  let data=XLSX.utils.sheet_to_json(sheet);

  data=data.map(r=>({...r,Value:"",SSS:false,AWS:false}));

  fs.writeFileSync('data.json',JSON.stringify(data,null,2));
  fs.unlinkSync(req.file.path);

  res.send("Excel uploaded");
});

// ================= SAVE =================
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

// ================= REPORT =================
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

// ================= ZIP =================
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