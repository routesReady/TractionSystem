require('dotenv').config();
const dns=require('dns');
dns.setServers(['8.8.8.8','8.8.4.4']);
const express=require('express'),cors=require('cors'),mongoose=require('mongoose'),multer=require('multer'),jwt=require('jsonwebtoken');
const bcrypt=require('bcryptjs');
const XLSX=require('xlsx');
const {SourceDataset,SavedRow}=require('./models');
const {trains,kavach,norm,tval,validDate,weekday,TEMPLATE_HEADERS}=require('./excel');
const app=express();
const upload=multer({storage:multer.memoryStorage(),limits:{fileSize:15*1024*1024}});
const allowedOrigins=String(process.env.CLIENT_ORIGIN||'http://localhost:5173')
  .split(',')
  .map(v=>v.trim().replace(/\\/+$/,''))
  .filter(Boolean);
app.use(cors({
  origin:(origin,callback)=>{
    if(!origin || allowedOrigins.includes(origin)) return callback(null,true);
    // Allow Vercel preview/production deployments for this project.
    if(/^https:\\/\\/traction-system-[a-z0-9-]+\\.vercel\\.app$/i.test(origin)) return callback(null,true);
    callback(new Error('CORS origin not allowed'));
  }
}));
app.use(express.json({limit:'2mb'}));

const auth=(req,res,next)=>{try{const h=req.headers.authorization||'';if(!h.startsWith('Bearer '))throw Error();req.user=jwt.verify(h.slice(7),process.env.JWT_SECRET);next();}catch(e){res.status(401).json({message:'Authentication required.'});}};
const excel=n=>/\.(xlsx|xls|xlsm)$/i.test(n||'');
const UPLOAD_SECRET=process.env.UPLOAD_PASSWORD||'';
const authenticatePassword=async password=>{
  if(!UPLOAD_SECRET||!password)return false;
  return password===UPLOAD_SECRET;
};
const save=async(key,fileName,rows,buffer,contentType)=>SourceDataset.findOneAndUpdate(
  {key},
  {key,fileName,rows,fileData:buffer,contentType:contentType||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',uploadedAt:new Date()},
  {upsert:true,new:true}
);

app.get('/api/health',(q,s)=>s.json({ok:true}));
app.get('/api/upload/status',async(q,s)=>{try{const a=await SourceDataset.findOne({key:'TRAINS_PER_DAY'}).lean(),b=await SourceDataset.findOne({key:'KAVACH_LOCO_DETAILS'}).lean();s.json({initialized:!!(a&&b),train:a&&{fileName:a.fileName,uploadedAt:a.uploadedAt,count:a.rows.length,hasFile:!!a.fileData},kavach:b&&{fileName:b.fileName,uploadedAt:b.uploadedAt,count:b.rows.length,hasFile:!!b.fileData}})}catch(e){s.status(500).json({message:e.message})}});

app.post('/api/upload/authenticate',async(q,s)=>{try{if(!(await authenticatePassword(q.body.password)))return s.status(401).json({message:'Incorrect upload password.'});s.json({token:jwt.sign({role:'uploader'},process.env.JWT_SECRET,{expiresIn:'30m'})})}catch(e){s.status(500).json({message:e.message})}});

// All uploads, including the first upload, require the configured upload password.
app.post('/api/upload/initial',auth,upload.fields([{name:'trainsFile',maxCount:1},{name:'kavachFile',maxCount:1}]),async(q,s)=>{try{
  const a=q.files?.trainsFile?.[0],b=q.files?.kavachFile?.[0];
  if(!a||!b||!excel(a.originalname)||!excel(b.originalname))return s.status(400).json({message:'Both Excel files are required.'});
  const existingTrain=await SourceDataset.exists({key:'TRAINS_PER_DAY'}),existingKavach=await SourceDataset.exists({key:'KAVACH_LOCO_DETAILS'});
  if(existingTrain&&existingKavach)return s.status(409).json({message:'Both source files already exist. Use Update Source Files.'});
  const tr=trains(a.buffer),ka=kavach(b.buffer);
  await save('TRAINS_PER_DAY',a.originalname,tr,a.buffer,a.mimetype);
  await save('KAVACH_LOCO_DETAILS',b.originalname,ka,b.buffer,b.mimetype);
  s.status(201).json({message:'Source files uploaded successfully.',trainCount:tr.length,kavachCount:ka.length});
}catch(e){s.status(500).json({message:e.message})}});

app.post('/api/upload/update',auth,upload.fields([{name:'trainsFile',maxCount:1},{name:'kavachFile',maxCount:1}]),async(q,s)=>{try{
  const a=q.files?.trainsFile?.[0],b=q.files?.kavachFile?.[0];
  if(!a&&!b)return s.status(400).json({message:'Select at least one updated file.'});
  if(a){if(!excel(a.originalname))return s.status(400).json({message:'Invalid train Excel file.'});await save('TRAINS_PER_DAY',a.originalname,trains(a.buffer),a.buffer,a.mimetype);}
  if(b){if(!excel(b.originalname))return s.status(400).json({message:'Invalid Kavach Excel file.'});await save('KAVACH_LOCO_DETAILS',b.originalname,kavach(b.buffer),b.buffer,b.mimetype);}
  s.json({message:'Source data updated successfully.'});
}catch(e){s.status(500).json({message:e.message})}});

function generatedSourceWorkbook(source){
  const data=source.key==='TRAINS_PER_DAY'
    ? source.rows.map((r,i)=>({
      'S.N':r.sourceRow||i+1,'DATE':r.date,'DAY':r.day,'TRAIN_No':r.trainNo,'DIR':r.dir,'TR_TYPE':r.trType,'T_O_Time':r.time,
      'LOCO_ATTACHED_DIV.':r.attachedDiv,'From_To':r.fromTo,'LOCO_link':r.locoLink,'LOCO_LINK_DIV':r.locoLinkDiv,
      'Day_of_wkg_in_territory':r.dayWork,'Loco_No':r.locoNo,'SHED':r.shed,'Loco_Type':r.locoType,'RLY':r.rly,
      'Loco_Kavach_Make':r.kavachMake,'Brake_system':r.brakeSystem,'KAVACH_WKG_SECTION':r.kavachSection,
      'Train_worked_with_Kavach_YES_NO':r.worked,'If_No_the_Reason \\n( Non Kavach Loco\\n/ Kavach defective\\n/ KAVACH Fitness cirtification N.Avl.\\n/ Crew incompetency)':r.noReason,'Remarks':r.remarks
    }))
    : source.rows.map(r=>({'Loco_No':r.locoNo,'SHED':r.shed,'Loco_Type':r.locoType,'RLY':r.rly,'Loco_Kavach_Make':r.kavachMake,'Brake_system':r.brakeSystem}));
  const ws=XLSX.utils.json_to_sheet(data,{header:source.key==='TRAINS_PER_DAY'?TEMPLATE_HEADERS:['Loco_No','SHED','Loco_Type','RLY','Loco_Kavach_Make','Brake_system']});
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,source.key==='TRAINS_PER_DAY'?'TRAINS_PER_DAY':'Kavach_Loco_Details');
  return XLSX.write(wb,{type:'buffer',bookType:'xlsx'});
}

app.get('/api/upload/download/:key',async(q,s)=>{try{
  const key=q.params.key==='trains'?'TRAINS_PER_DAY':q.params.key==='kavach'?'KAVACH_LOCO_DETAILS':null;
  if(!key)return s.status(400).json({message:'Unknown source file.'});
  const source=await SourceDataset.findOne({key}).lean();
  if(!source)return s.status(404).json({message:'Source file is not uploaded yet.'});
  const filename=source.fileName|| (key==='TRAINS_PER_DAY'?'TRAINS_PER_DAY.xlsx':'Kavach_Loco_Details.xlsx');
  const data=source.fileData?.length?source.fileData:generatedSourceWorkbook(source);
  s.setHeader('Content-Type',source.contentType||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  s.setHeader('Content-Disposition',`attachment; filename="${filename.replace(/"/g,'') }"`);
  s.send(data);
}catch(e){s.status(500).json({message:e.message})}});

app.get('/api/data/options',async(q,s)=>{try{
  const a=await SourceDataset.findOne({key:'TRAINS_PER_DAY'}).lean();
  if(!a)return s.status(409).json({message:'Upload TRAINS_PER_DAY.xlsx first.'});
  const unique=field=>Array.from(new Set(a.rows.map(r=>String(r[field]??'').trim()).filter(Boolean)));
  const worked=unique('worked');
  const reason=unique('noReason');
  // If the source contains no populated reason values, keep the standard workbook choices available.
  const defaultReasons=['Non Kavach Loco','Kavach defective','KAVACH Fitness cirtification N.Avl.','Crew incompetency'];
  s.json({workedOptions:worked.length?worked:['YES','NO'],reasonOptions:reason.length?reason:defaultReasons});
}catch(e){s.status(500).json({message:e.message})}});

app.get('/api/kavach/lookup', async (q, s) => {try {
  const locoNo=norm(q.query.locoNo);if(!locoNo)return s.status(400).json({message:'Enter a Loco_No.'});
  const b=await SourceDataset.findOne({key:'KAVACH_LOCO_DETAILS'}).lean();if(!b)return s.status(409).json({message:'Upload Kavach_Loco_Details.xlsx first.'});
  const k=b.rows.find(x=>norm(x.locoNo)===locoNo);if(!k)return s.status(404).json({message:`Loco_No ${q.query.locoNo} was not found in Kavach_Loco_Details.xlsx.`});
  s.json({found:true,locoNo:k.locoNo,shed:k.shed||'',locoType:k.locoType||'',rly:k.rly||'',kavachMake:k.kavachMake||'',brakeSystem:k.brakeSystem||''});
}catch(e){s.status(500).json({message:e.message})}});

function buildRow(x,searchDate,day,k){return {sourceRow:x.sourceRow,sn:x.sourceRow,date:searchDate,day,trainNo:x.trainNo,dir:x.dir,trType:x.trType,time:x.time,attachedDiv:x.attachedDiv,fromTo:x.fromTo,locoLink:x.locoLink,locoLinkDiv:x.locoLinkDiv,dayWork:x.dayWork,locoNo:k?k.locoNo:(x.locoNo||''),shed:k?.shed||x.shed||'',locoType:k?.locoType||x.locoType||'',rly:k?.rly||x.rly||'',kavachMake:k?.kavachMake||x.kavachMake||'',brakeSystem:k?.brakeSystem||x.brakeSystem||'',kavachSection:x.kavachSection||'',worked:x.worked||'',noReason:x.noReason||'',remarks:x.remarks||'',kavachFound:!!k};}

app.get('/api/data/search',async(q,s)=>{try{
  const {date,page=1,pageSize=10}=q.query;if(!validDate(date||''))return s.status(400).json({message:'Use DD-MM-YYYY.'});
  const a=await SourceDataset.findOne({key:'TRAINS_PER_DAY'}).lean(),b=await SourceDataset.findOne({key:'KAVACH_LOCO_DETAILS'}).lean();
  if(!a||!b)return s.status(409).json({message:'Upload both source files first.'});
  const km=new Map(b.rows.map(x=>[norm(x.locoNo),x])),day=weekday(date);
  const rows=a.rows.filter(x=>String(x.day).trim().toUpperCase()===day).sort((x,y)=>tval(x.time)-tval(y.time));
  const sourceRows=rows.map(x=>x.sourceRow);const saved=await SavedRow.find({searchDate:date,sourceRow:{$in:sourceRows}}).select('sourceRow row').lean();const savedMap=new Map(saved.map(x=>[String(x.sourceRow),x]));
  const enriched=rows.map(x=>{const savedRow=savedMap.get(String(x.sourceRow));const k=savedRow?.row?.locoNo?km.get(norm(savedRow.row.locoNo)):null;return {...buildRow(x,date,day,k),locoNo:savedRow?.row?.locoNo||'',worked:(savedRow?.row?.worked ?? x.worked ?? ''),noReason:(savedRow?.row?.noReason ?? x.noReason ?? ''),remarks:(savedRow?.row?.remarks ?? x.remarks ?? ''),added:!!savedRow,kavachFound:!!k};});
  const size=Math.max(1,Number(pageSize)),pg=Math.max(1,Number(page)),total=enriched.length;
  s.json({date,weekday:day,total,page:pg,pageSize:size,totalPages:Math.max(1,Math.ceil(total/size)),rows:enriched.slice((pg-1)*size,pg*size)});
}catch(e){s.status(500).json({message:e.message})}});

app.post('/api/data/rows',async(q,s)=>{try{
  const {searchDate,weekday:day,row}=q.body,locoNo=norm(row?.locoNo);
  if(!searchDate||!validDate(searchDate)||!day||!row?.sourceRow||!locoNo)return s.status(400).json({message:'Valid search date and Loco_No are required.'});
  if(String(row?.remarks||'').length>500)return s.status(400).json({message:'Remarks cannot exceed 500 characters.'});
  const b=await SourceDataset.findOne({key:'KAVACH_LOCO_DETAILS'}).lean();if(!b)return s.status(409).json({message:'Kavach source is not uploaded.'});
  const k=b.rows.find(x=>norm(x.locoNo)===locoNo);if(!k)return s.status(400).json({message:`Loco_No ${row.locoNo} was not found in Kavach_Loco_Details.xlsx.`});
  const savedRow={...row,locoNo:k.locoNo,shed:k.shed||'',locoType:k.locoType||'',rly:k.rly||'',kavachMake:k.kavachMake||'',brakeSystem:k.brakeSystem||'',kavachFound:true,remarks:String(row.remarks||'').slice(0,500)};
  const x=await SavedRow.findOneAndUpdate({searchDate,sourceRow:row.sourceRow},{searchDate,weekday:day,sourceRow:row.sourceRow,row:savedRow,addedAt:new Date()},{upsert:true,new:true,setDefaultsOnInsert:true});
  s.status(201).json({message:'Row added to database.',id:x._id,added:true,row:savedRow});
}catch(e){if(e.code===11000)return s.status(409).json({message:'This row is already added.'});s.status(500).json({message:e.message})}});

function parseDMY(s){if(!validDate(s||''))return null;const[d,m,y]=s.split('-').map(Number);return new Date(y,m-1,d);}
function filterSaved(all,startDate,endDate){return all.filter(x=>{const d=parseDMY(x.searchDate);return(!startDate||d>=parseDMY(startDate))&&(!endDate||d<=parseDMY(endDate));});}

app.get('/api/data/saved',async(q,s)=>{try{
  const{startDate,endDate,page=1,pageSize=20}=q.query;if(startDate&&!validDate(startDate))return s.status(400).json({message:'Start date must be DD-MM-YYYY.'});if(endDate&&!validDate(endDate))return s.status(400).json({message:'End date must be DD-MM-YYYY.'});if(startDate&&endDate&&parseDMY(startDate)>parseDMY(endDate))return s.status(400).json({message:'Start date cannot be after end date.'});
  const filtered=filterSaved(await SavedRow.find({}).sort({createdAt:-1}).lean(),startDate,endDate),size=Math.max(1,Number(pageSize)),pg=Math.max(1,Number(page)),total=filtered.length;
  s.json({total,page:pg,pageSize:size,totalPages:Math.max(1,Math.ceil(total/size)),rows:filtered.slice((pg-1)*size,pg*size)});
}catch(e){s.status(500).json({message:e.message})}});

app.get('/api/data/export',async(q,s)=>{try{
  const{startDate,endDate}=q.query;if(startDate&&!validDate(startDate))return s.status(400).json({message:'Start date must be DD-MM-YYYY.'});if(endDate&&!validDate(endDate))return s.status(400).json({message:'End date must be DD-MM-YYYY.'});if(startDate&&endDate&&parseDMY(startDate)>parseDMY(endDate))return s.status(400).json({message:'Start date cannot be after end date.'});
  const rows=filterSaved(await SavedRow.find({}).sort({createdAt:1}).lean(),startDate,endDate);
  const flat=rows.map((x,i)=>({'S.N':i+1,'DATE':x.searchDate,'DAY':x.weekday,'TRAIN_No':x.row?.trainNo||'','DIR':x.row?.dir||'','TR_TYPE':x.row?.trType||'','T_O_Time':x.row?.time||'','LOCO_ATTACHED_DIV.':x.row?.attachedDiv||'','From_To':x.row?.fromTo||'','LOCO_link':x.row?.locoLink||'','LOCO_LINK_DIV':x.row?.locoLinkDiv||'','Day_of_wkg_in_territory':x.row?.dayWork||'','Loco_No':x.row?.locoNo||'','SHED':x.row?.shed||'','Loco_Type':x.row?.locoType||'','RLY':x.row?.rly||'','Loco_Kavach_Make':x.row?.kavachMake||'','Brake_system':x.row?.brakeSystem||'','KAVACH_WKG_SECTION':x.row?.kavachSection||'','Train_worked_with_Kavach_YES_NO':x.row?.worked||'','If_No_the_Reason \\n( Non Kavach Loco\\n/ Kavach defective\\n/ KAVACH Fitness cirtification N.Avl.\\n/ Crew incompetency)':x.row?.noReason||'','Remarks':x.row?.remarks||''}));
  const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(flat,{header:TEMPLATE_HEADERS});XLSX.utils.book_append_sheet(wb,ws,'Added Data');ws['!cols']=TEMPLATE_HEADERS.map((h,i)=>({wch:i===20?46:i===21?42:Math.max(14,Math.min(28,String(h).length+3))}));
  const buffer=XLSX.write(wb,{type:'buffer',bookType:'xlsx'}),suffix=startDate||endDate?`_${startDate||'all'}_to_${endDate||'all'}`:'_all';
  s.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');s.setHeader('Content-Disposition',`attachment; filename="KV_Added_Data${suffix}.xlsx"`);s.send(buffer);
}catch(e){s.status(500).json({message:e.message})}});

mongoose.connect(process.env.MONGODB_URI).then(()=>app.listen(Number(process.env.PORT||5000),()=>console.log('API running on port '+(process.env.PORT||5000)))).catch(e=>{console.error(e);process.exit(1)});
