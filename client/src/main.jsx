import React,{useEffect,useState}from'react';
import{createRoot}from'react-dom/client';
import{motion}from'framer-motion';
import{CalendarDays,CheckCircle2,Database,Download,FileSpreadsheet,LockKeyhole,RefreshCw,Search,UploadCloud,XCircle,LoaderCircle,Menu,ShieldCheck,TrainFront,FileDown,Edit3,Save,RotateCcw}from'lucide-react';
import axios from'axios';
import'./styles.css';

const configuredApi=String(import.meta.env.VITE_API_URL||'http://localhost:5000/api').replace(/\/+$/,'');
const api=axios.create({baseURL:configuredApi.endsWith('/api')?configuredApi:`${configuredApi}/api`});

const cols=[
 ['sn','S.N'],['date','DATE'],['day','DAY'],['trainNo','TRAIN_No'],['dir','DIR'],['trType','TR_TYPE'],['time','T_O_Time'],
 ['attachedDiv','LOCO_ATTACHED_DIV.'],['fromTo','From_To'],['locoLink','LOCO_link'],['locoLinkDiv','LOCO_LINK_DIV'],
 ['dayWork','Day_of_wkg_in_territory'],['locoNo','Loco_No'],['shed','SHED'],['locoType','Loco_Type'],['rly','RLY'],
 ['kavachMake','Loco_Kavach_Make'],['brakeSystem','Brake_system'],['kavachSection','KAVACH_WKG_SECTION'],
 ['worked','Train_worked_with_Kavach_YES_NO'],['noReason','If_No_the_Reason'],['remarks','Remarks']
];
const editableKeys=new Set(['locoNo','worked','noReason','remarks']);
const derivedKeys=new Set(['shed','locoType','rly','kavachMake','brakeSystem']);

function dmyToISO(value){
 const m=String(value||'').match(/^(\d{2})-(\d{2})-(\d{4})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:'';
}
function isoToDMY(value){
 const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:'';
}
function DatePicker({value,onChange,label}){
 return <div className="field"><label><CalendarDays/> {label}</label><div className="date-input-wrap"><input type="date" value={dmyToISO(value)} onChange={e=>onChange(isoToDMY(e.target.value))}/><CalendarDays className="date-icon"/></div></div>
}

function App(){
 const[tab,setTab]=useState('dash'),[date,setDate]=useState('11-09-2026'),[res,setRes]=useState(),[msg,setMsg]=useState(''),[status,setStatus]=useState(),[loading,setLoading]=useState(false),[mobileNav,setMobileNav]=useState(false);
 const[options,setOptions]=useState({workedOptions:[],reasonOptions:[]});
 const statusLoad=async()=>{try{const{data}=await api.get('/upload/status');setStatus(data)}catch(e){setMsg('Backend is not reachable. Check the deployed API or start the server.')}};
 const loadOptions=async()=>{try{const{data}=await api.get('/data/options');setOptions(data)}catch(e){setOptions({workedOptions:[],reasonOptions:[]})}};
 useEffect(()=>{statusLoad();loadOptions()},[]);
 async function search(){setLoading(true);setMsg('');try{const{data}=await api.get('/data/search',{params:{date}});setRes(data);await loadOptions()}catch(e){setRes();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 function patchRow(sourceRow,patch){setRes(x=>x?{...x,rows:x.rows.map(r=>r.sourceRow===sourceRow?{...r,...patch}:r)}:x)}
 async function lookup(row){const locoNo=String(row.locoNo||'').trim().slice(0,6);patchRow(row.sourceRow,{locoNo,kavachFound:false,kavachLoading:Boolean(locoNo),kavachError:''});if(!locoNo){patchRow(row.sourceRow,{kavachLoading:false,kavachError:'Enter Loco_No'});return}try{const{data}=await api.get('/kavach/lookup',{params:{locoNo}});patchRow(row.sourceRow,{...data,kavachFound:true,kavachLoading:false,kavachError:''})}catch(e){patchRow(row.sourceRow,{kavachFound:false,kavachLoading:false,kavachError:e.response?.data?.message||'Loco_No not found in Kavach_Loco_Details.xlsx.'})}}
 async function saveRow(row){
  if(!String(row.locoNo||'').trim())return setMsg('Enter Loco_No before saving.');
  if(!row.kavachFound)return setMsg('Enter a valid Loco_No and complete the Kavach lookup.');
  if(String(row.remarks||'').length>500)return setMsg('Remarks cannot exceed 500 characters.');
  try{const{data}=await api.post('/data/rows',{searchDate:res.date,weekday:res.weekday,row});patchRow(row.sourceRow,{...data.row,added:true,editing:false,kavachFound:true});setMsg(`Loco ${data.row.locoNo} record ${row.added?'updated':'saved'} successfully.`)}catch(e){setMsg(e.response?.data?.message||e.message)}
 }
 function beginModify(row){patchRow(row.sourceRow,{editing:true});setMsg(`Modify mode enabled for Loco ${row.locoNo||'record'}.`)}
 const title=tab==='dash'?'Daily Traction Roster':tab==='saved'?'Saved Data & Excel Export':'Source Data Management';
 return <div className="app-shell">
  <header className="gov-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation</h1><p>West Central Railway, Kota Division</p></div><button className="mobile-menu" onClick={()=>setMobileNav(v=>!v)} aria-label="Open navigation"><Menu/></button></div>
   <div className="nav-strip"><div className="nav-inner">{[['dash',<TrainFront/>,'Dashboard'],['saved',<Download/>,'Saved Data / Export'],['src',<Database/>,'Data Sources']].map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>{setTab(id);setMobileNav(false)}}>{icon}{label}</button>)}<div className="nav-spacer"/><div className="official-chip"><ShieldCheck/> Official Workflow</div></div></div>
   {mobileNav&&<div className="mobile-nav">{[['dash','Dashboard'],['saved','Saved Data / Export'],['src','Data Sources']].map(([id,label])=><button key={id} className={tab===id?'active':''} onClick={()=>{setTab(id);setMobileNav(false)}}>{label}</button>)}</div>}
  </header>
  <main className="content"><div className="page-heading"><div><span>TRACTION OPERATION</span><h2>{title}</h2></div><button className="refresh-btn" onClick={()=>{statusLoad();loadOptions()}} title="Refresh"><RefreshCw/></button></div>
   {tab==='dash'?<Dashboard date={date} setDate={setDate} res={res} search={search} loading={loading} msg={msg} saveRow={saveRow} patchRow={patchRow} lookup={lookup} beginModify={beginModify} options={options}/>:tab==='saved'?<SavedData setMsg={setMsg}/>:<Upload status={status} refresh={()=>{statusLoad();loadOptions()}} setMsg={setMsg}/>} 
  </main><footer className="footer"><div>West Central Railway · Kota Division · Traction Operation</div><span>@mak</span></footer>
 </div>
}

function Dashboard({date,setDate,res,search,loading,msg,saveRow,patchRow,lookup,beginModify,options}){
 const workedOptions=options.workedOptions||[],reasonOptions=options.reasonOptions||[];
 return <>
  <section className="search-card"><DatePicker label="Date" value={date} onChange={setDate}/><button className="primary-btn" onClick={search} disabled={loading||!date}><Search/>{loading?'Fetching...':'Fetch Data'}</button></section>
  {msg&&<div className="alert"><XCircle/>{msg}</div>}
  {res&&<section className="results-card"><div className="result-head"><div><span className="eyebrow">DAILY ROSTER</span><h3>{res.weekday} <i>·</i> {res.date}</h3></div><div className="result-meta">{res.total} records</div></div>
   <div className="table-shell"><table className="data-table"><thead><tr>{cols.map(([key,label])=><th key={key} className={key==='worked'?'worked-head':key==='noReason'?'reason-head':''} title={label}>{key==='worked'?<>Train Worked<br/>with Kavach</>:key==='noReason'?'If No – Reason':label}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>
    {res.rows.length===0?<tr><td colSpan={cols.length+1} className="empty-cell">No train records found for the selected date.</td></tr>:res.rows.map((r,i)=><motion.tr initial={{opacity:0,y:5}} animate={{opacity:1,y:0}} transition={{duration:.18,delay:i*.01}} key={r.sourceRow}>
     {cols.map(([k,label])=><td key={k} className={`${k==='locoNo'?'loco-cell ':''}${k==='remarks'?'remarks-cell':''}`}>
      <CellEditor k={k} label={label} row={r} patchRow={patchRow} lookup={lookup} workedOptions={workedOptions} reasonOptions={reasonOptions}/>
     </td>)}
     <td className="action-cell">{r.added&&!r.editing?<button className="update modify" onClick={()=>beginModify(r)}><Edit3/>Modify</button>:<button className="update" disabled={!String(r.locoNo||'').trim()||!r.kavachFound||r.kavachLoading} onClick={()=>saveRow(r)}>{r.editing?<Save/>:<Database/>}{r.editing?'Update':r.kavachLoading?'Looking up...':r.kavachFound?'Update':'Enter Loco_No'}</button>}</td>
    </motion.tr>)}
   </tbody></table></div>
  </section>}
 </>
}

function CellEditor({k,label,row,patchRow,lookup,workedOptions,reasonOptions}){
 const editing=Boolean(row.editing),value=row[k]??'';
 if(k==='sn')return <b>{row.sn}</b>;
 if(k==='date')return <span>{row.date}</span>;
 if(k==='day')return <span>{row.day||'—'}</span>;
 if(k==='locoNo'){if(row.added&&!editing)return <span className="loco-value">{value||'—'}</span>;return <div className="locoeditor"><input className="loco-input" maxLength={6} inputMode="numeric" value={value} onChange={e=>patchRow(row.sourceRow,{locoNo:e.target.value.replace(/\D/g,'').slice(0,6),kavachFound:false,kavachError:''})} onBlur={()=>lookup(row)} onKeyDown={e=>{if(e.key==='Enter')lookup({...row,locoNo:e.currentTarget.value})}} placeholder="Loco No."/><button className="lookup" onClick={()=>lookup(row)} disabled={!String(value).trim()||row.kavachLoading} title="Find Loco_No in Kavach file">{row.kavachLoading?<LoaderCircle className="spin"/>:<Search/>}</button>{row.kavachError&&<small className="lookup-error">Not found</small>}</div>; }
 if(k==='worked')return editing||!row.added?<select className="cell-select worked-select" value={value} onChange={e=>patchRow(row.sourceRow,{worked:e.target.value,noReason:String(e.target.value).trim().toUpperCase()==='YES'?'':row.noReason})}><option value="">Select</option>{workedOptions.map(v=><option key={v} value={v}>{v}</option>)}</select>:<span>{value||'—'}</span>;
 if(k==='noReason')return editing||!row.added?<select className="cell-select reason-select" value={value} onChange={e=>patchRow(row.sourceRow,{noReason:e.target.value})}><option value="">Select</option>{reasonOptions.map(v=><option key={v} value={v}>{v}</option>)}</select>:<span className="wrap-text">{value||'—'}</span>;
 if(k==='remarks')return editing||!row.added?<div className="remarks-wrap"><textarea maxLength={500} value={value} onChange={e=>patchRow(row.sourceRow,{remarks:e.target.value})} placeholder="Add remarks..."/><small>{String(value).length}/500</small></div>:<span className="wrap-text">{value||'—'}</span>;
 if(editing&&editableKeys.has(k))return <input className="cell-editor" value={value} onChange={e=>patchRow(row.sourceRow,{[k]:e.target.value})} />;
 if(derivedKeys.has(k))return <span className="derived-cell">{value||'—'}</span>;
 return <span className="wrap-text">{value||'—'}</span>;
}

function SavedData({setMsg}){const[startDate,setStartDate]=useState(''),[endDate,setEndDate]=useState(''),[data,setData]=useState(),[loading,setLoading]=useState(false);
 async function load(){setLoading(true);setMsg('');try{const{data:x}=await api.get('/data/saved',{params:{startDate,endDate}});setData(x)}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 async function download(){setMsg('');try{const r=await api.get('/data/export',{params:{startDate,endDate},responseType:'blob'});const cd=r.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||'Traction_Operation_Data.xlsx';const url=URL.createObjectURL(new Blob([r.data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}catch(e){setMsg(await blobError(e,'Excel export failed.'))}}
 return <section className="panel"><div className="panel-intro compact"><div className="intro-icon"><Download/></div><div><span className="eyebrow">DATABASE EXPORT</span><h3>Saved records &amp; Excel download</h3><p>Filter saved records by date and export the complete data set.</p></div></div><div className="filterbar"><DatePicker label="From date" value={startDate} onChange={setStartDate}/><DatePicker label="To date" value={endDate} onChange={setEndDate}/><button className="secondary-btn" onClick={load} disabled={loading}><Search/>{loading?'Filtering...':'Filter'}</button><button className="primary-btn" onClick={download}><FileSpreadsheet/>Download Excel</button></div>{data&&<div className="savedresult"><strong>{data.total}</strong> saved row(s) match the selected date range.</div>}{data?.rows?.length>0&&<div className="table-shell saved-table"><table className="data-table"><thead><tr>{cols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':label}</th>)}</tr></thead><tbody>{data.rows.map(x=><tr key={x._id}>{cols.map(([k])=><td key={k}>{k==='sn'?x.row?.sn:k==='date'?x.searchDate:k==='day'?x.weekday:<span className="wrap-text">{x.row?.[k]||'—'}</span>}</td>)}</tr>)}</tbody></table></div>}</section>}

async function blobError(error,fallback){if(error.response?.data instanceof Blob){try{const text=await error.response.data.text();const json=JSON.parse(text);return json.message||fallback}catch{return fallback}}return error.response?.data?.message||error.message||fallback}

function Upload({status,refresh,setMsg}){const[tf,setTf]=useState(),[kf,setKf]=useState(),[pw,setPw]=useState(''),[busy,setBusy]=useState(false);const initialized=!!status?.initialized;
 async function submit(){if(!tf&&!kf)return setMsg('Select at least one Excel file.');if(!pw)return setMsg('Enter the upload password.');setBusy(true);try{const{data:a}=await api.post('/upload/authenticate',{password:pw}),f=new FormData();if(tf)f.append('trainsFile',tf);if(kf)f.append('kavachFile',kf);if(!initialized){if(!tf||!kf){setMsg('For first setup, select both source files.');setBusy(false);return}await api.post('/upload/initial',f,{headers:{Authorization:'Bearer '+a.token}})}else await api.post('/upload/update',f,{headers:{Authorization:'Bearer '+a.token}});setMsg('Source data updated successfully.');setPw('');setTf();setKf();await refresh()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 async function download(key){setMsg('');try{const r=await api.get(`/upload/download/${key}`,{responseType:'blob'});const cd=r.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||(key==='trains'?'TRAINS_PER_DAY.xlsx':'Kavach_Loco_Details.xlsx');const url=URL.createObjectURL(new Blob([r.data]));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}catch(e){setMsg(await blobError(e,'Source download failed.'))}}
 return <section className="panel"><div className="panel-intro compact"><div className="intro-icon"><Database/></div><div><span className="eyebrow">SOURCE DATA MANAGEMENT</span><h3>Upload &amp; maintain source workbooks</h3><p>Uploads require the configured password. Stored workbooks can be downloaded without a password.</p></div></div><div className="source-grid"><SourceCard title="TRAINS_PER_DAY.xlsx" file={tf} set={setTf} status={status?.train} onDownload={()=>download('trains')}/><SourceCard title="Kavach_Loco_Details.xlsx" file={kf} set={setKf} status={status?.kavach} onDownload={()=>download('kavach')}/></div><section className="security"><div className="security-head"><div className="security-icon"><LockKeyhole/></div><div><span className="eyebrow">UPLOAD SECURITY</span><h4>{initialized?'Password required for source updates':'Password required for initial upload'}</h4></div></div><div className="security-form"><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Enter upload password"/><button className="primary-btn" onClick={submit} disabled={busy}><LockKeyhole/>{busy?'Processing...':initialized?'Authenticate & Update':'Authenticate & Upload'}</button></div><div className="secure"><CheckCircle2/> Upload is password-protected. Source downloads do not require a password.</div></section></section>}
function SourceCard({title,file,set,status,onDownload}){return <section className="source-card"><div className="source-card-top"><FileSpreadsheet/><span>{status?.count??0} rows</span></div><h4>{title}</h4><p>{status?.fileName?`Stored: ${status.fileName}`:'Not uploaded yet'}</p><label className="picker"><UploadCloud/><span>{file?file.name:'Choose .xlsx / .xls / .xlsm file'}</span><input type="file" accept=".xlsx,.xls,.xlsm" onChange={e=>set(e.target.files?.[0])}/></label><button className="download-source" onClick={onDownload} disabled={!status}><FileDown/>Download stored file</button></section>}
createRoot(document.getElementById('root')).render(<App/>);
