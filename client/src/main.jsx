import React,{useEffect,useMemo,useState}from'react';
import{createRoot}from'react-dom/client';
import{motion}from'framer-motion';
import{CalendarDays,CheckCircle2,Database,Download,FileSpreadsheet,LockKeyhole,RefreshCw,Search,UploadCloud,XCircle,Zap,LoaderCircle,Menu,ShieldCheck,TrainFront,FileDown,ChevronLeft,ChevronRight}from'lucide-react';
import axios from'axios';
import'./styles.css';

const configuredApi=String(import.meta.env.VITE_API_URL||'http://localhost:5000/api').replace(/\/+$/,'');
const api=axios.create({
 baseURL: configuredApi.endsWith('/api') ? configuredApi : `${configuredApi}/api`
});
const size=10;
const cols=[
 ['sn','S.N'],['date','DATE'],['day','DAY'],['trainNo','TRAIN_No'],['dir','DIR'],['trType','TR_TYPE'],['time','T_O_Time'],
 ['attachedDiv','LOCO_ATTACHED_DIV.'],['fromTo','From_To'],['locoLink','LOCO_link'],['locoLinkDiv','LOCO_LINK_DIV'],
 ['dayWork','Day_of_wkg_in_territory'],['locoNo','Loco_No'],['shed','SHED'],['locoType','Loco_Type'],['rly','RLY'],
 ['kavachMake','Loco_Kavach_Make'],['brakeSystem','Brake_system'],['kavachSection','KAVACH_WKG_SECTION'],
 ['worked','Train_worked_with_Kavach_YES_NO'],['noReason','If_No_the_Reason\n( Non Kavach Loco\n/ Kavach defective\n/ KAVACH Fitness cirtification N.Avl.\n/ Crew incompetency)'],['remarks','Remarks']
];
const reasonDefault=['Non Kavach Loco','Kavach defective','KAVACH Fitness cirtification N.Avl.','Crew incompetency'];

function App(){
 const[tab,setTab]=useState('dash'),[date,setDate]=useState('11-09-2026'),[res,setRes]=useState(),[msg,setMsg]=useState(''),[status,setStatus]=useState(),[loading,setLoading]=useState(false),[mobileNav,setMobileNav]=useState(false);
 const[options,setOptions]=useState({workedOptions:['YES','NO'],reasonOptions:reasonDefault});
 const statusLoad=async()=>{try{const x=await api.get('/upload/status');setStatus(x.data)}catch(e){setMsg('Backend is not reachable. Start the server on port 5000.')}};
 useEffect(()=>{statusLoad();loadOptions()},[]);
 async function loadOptions(){try{const{data}=await api.get('/data/options');setOptions(data)}catch(e){/* source may not be uploaded yet */}}
 async function search(p=1){setLoading(true);setMsg('');try{const{data}=await api.get('/data/search',{params:{date,page:p,pageSize:size}});setRes(data);setPageSafe(p);await loadOptions()}catch(e){setRes();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 function setPageSafe(p){setRes(x=>x?{...x,page:p}:x)}
 function patchRow(sourceRow,patch){setRes(x=>x?{...x,rows:x.rows.map(r=>r.sourceRow===sourceRow?{...r,...patch}:r)}:x)}
 async function lookup(row){const locoNo=String(row.locoNo||'').trim();if(!locoNo){patchRow(row.sourceRow,{kavachFound:false,kavachLoading:false,kavachError:'Enter Loco_No'});return}patchRow(row.sourceRow,{kavachLoading:true,kavachFound:false,kavachError:''});try{const{data}=await api.get('/kavach/lookup',{params:{locoNo}});patchRow(row.sourceRow,{...data,kavachFound:true,kavachLoading:false})}catch(e){patchRow(row.sourceRow,{kavachFound:false,kavachLoading:false,kavachError:e.response?.data?.message||'Loco_No not found in Kavach_Loco_Details.xlsx.'})}}
 async function update(row){if(!String(row.locoNo||'').trim())return setMsg('Enter Loco_No before clicking Update.');if(!row.kavachFound)return setMsg('Enter a valid Loco_No and wait for Kavach data to be found.');if(String(row.remarks||'').length>500)return setMsg('Remarks cannot exceed 500 characters.');try{const{data}=await api.post('/data/rows',{searchDate:res.date,weekday:res.weekday,row});setRes(x=>x?{...x,rows:x.rows.map(r=>r.sourceRow===row.sourceRow?{...r,...data.row,added:true,kavachFound:true}:r)}:x);setMsg(`Loco ${row.locoNo} added successfully.`)}catch(e){setMsg(e.response?.data?.message||e.message)}}
 const title=tab==='dash'?'Train & Kavach Search':tab==='saved'?'Added Data & Excel Export':'Source Data Management';
 return <div className="app-shell">
  <header className="gov-header">
   <div className="gov-inner">
    <div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div>
    <div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY · KOTA DIVISION</div><h1>Traction System</h1><p>West Central Railway, Kota Division</p></div>
    <button className="mobile-menu" onClick={()=>setMobileNav(v=>!v)} aria-label="Open navigation"><Menu/></button>
   </div>
   <div className="nav-strip"><div className="nav-inner">{[['dash',<TrainFront/>,'Dashboard'],['saved',<Download/>,'Added Data / Export'],['src',<Database/>,'Data Sources']].map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>{setTab(id);setMobileNav(false)}}>{icon}{label}</button>)}<div className="nav-spacer"/><div className="official-chip"><ShieldCheck/> Official Workflow</div></div></div>
   {mobileNav&&<div className="mobile-nav">{[['dash','Dashboard'],['saved','Added Data / Export'],['src','Data Sources']].map(([id,label])=><button key={id} className={tab===id?'active':''} onClick={()=>{setTab(id);setMobileNav(false)}}>{label}</button>)}</div>}
  </header>
  <main className="content">
   <div className="page-heading"><div><span>TRACTION SYSTEM</span><h2>{title}</h2></div><button className="refresh-btn" onClick={statusLoad} title="Refresh source status"><RefreshCw/></button></div>
   {tab==='dash'?<Dashboard date={date} setDate={setDate} res={res} search={search} loading={loading} msg={msg} update={update} patchRow={patchRow} lookup={lookup} options={options}/>:tab==='saved'?<SavedData setMsg={setMsg}/>:<Upload status={status} refresh={()=>{statusLoad();loadOptions()}} setMsg={setMsg}/>} 
  </main>
  <footer className="footer"><div>West Central Railway · Kota Division · Traction System</div><span>@mak</span></footer>
 </div>
}

function Dashboard({date,setDate,res,search,loading,msg,update,patchRow,lookup,options}){
 const workedOptions=options.workedOptions||['YES','NO'], reasonOptions=options.reasonOptions||reasonDefault;
 return <>
  <section className="hero-card"><div className="hero-copy"><span className="eyebrow">DATE-DRIVEN TRAIN LOOKUP</span><h3>Daily train roster &amp; Kavach enrichment</h3><p>Select a date to load the daily roster. Enter a <b>Loco_No</b> in each row and the portal will fetch its corresponding Kavach details from the uploaded source file.</p></div><div className="hero-emblem"><Zap/></div></section>
  <section className="search-card"><div className="field"><label><CalendarDays/> Date</label><input value={date} onChange={e=>setDate(e.target.value)} placeholder="DD-MM-YYYY"/></div><button className="primary-btn" onClick={()=>search(1)} disabled={loading}><Search/>{loading?'Fetching...':'Fetch Data'}</button></section>
  {msg&&<div className="alert"><XCircle/>{msg}</div>}
  {res&&<section className="results-card"><div className="result-head"><div><span className="eyebrow">DAILY ROSTER</span><h3>{res.weekday} <i>·</i> {res.date}</h3></div><div className="result-meta">{res.total} records <span>•</span> Page {res.page}/{Math.max(1,res.totalPages)}</div></div>
   <div className="table-shell"><table className="data-table"><thead><tr>{cols.map(x=><th key={x[0]} title={x[1]}>{x[1]}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>
    {res.rows.map((r,i)=><motion.tr initial={{opacity:0,y:5}} animate={{opacity:1,y:0}} transition={{duration:.18,delay:i*.015}} key={r.sourceRow}>
     {cols.map(([k,label])=><td key={k} className={k==='locoNo'?'loco-cell':''}>
      {k==='locoNo'?<div className="locoeditor"><input value={r[k]||''} onChange={e=>patchRow(r.sourceRow,{locoNo:e.target.value,kavachFound:false,kavachError:''})} onBlur={()=>lookup(r)} onKeyDown={e=>{if(e.key==='Enter')lookup({...r,locoNo:e.currentTarget.value})}} placeholder="Enter Loco_No"/><button className="lookup" onClick={()=>lookup(r)} disabled={!String(r.locoNo||'').trim()||r.kavachLoading} title="Find Loco_No in Kavach file">{r.kavachLoading?<LoaderCircle className="spin"/>:<Search/>}</button>{r.kavachError&&<small className="lookup-error">Not found</small>}</div>
      :k==='worked'?<select className="cell-select" value={r[k]||''} onChange={e=>patchRow(r.sourceRow,{worked:e.target.value,noReason:e.target.value==='YES'?'':r.noReason})}>{<option value="">Select</option>}{workedOptions.filter(v=>v!==undefined).map(v=><option key={v} value={v}>{v}</option>)}</select>
      :k==='noReason'?<select className="cell-select reason-select" value={r[k]||''} onChange={e=>patchRow(r.sourceRow,{noReason:e.target.value})}><option value="">Select reason</option>{[...new Set([...(reasonOptions||[]),r[k]||''].filter(Boolean))].map(v=><option key={v} value={v}>{v}</option>)}</select>
      :k==='remarks'?<div className="remarks-wrap"><textarea maxLength={500} value={r[k]||''} onChange={e=>patchRow(r.sourceRow,{remarks:e.target.value})} placeholder="Add remarks..."/><small>{String(r[k]||'').length}/500</small></div>
      :k==='sn'?<b>{(res.page-1)*size+i+1}</b>:(r[k]||'—')}
     </td>)}
     <td className="action-cell"><button className={r.added?'update added':'update'} disabled={!String(r.locoNo||'').trim()||!r.kavachFound||r.added||r.kavachLoading} onClick={()=>update(r)}>{r.added?<CheckCircle2/>:<Database/>}{r.added?'Added':r.kavachLoading?'Looking up...':r.kavachFound?'Update':'Enter Loco_No'}</button></td>
    </motion.tr>)}
   </tbody></table></div>
   <div className="pages"><button disabled={res.page<=1} onClick={()=>search(res.page-1)}><ChevronLeft/>Previous</button><b>{res.page}</b><button disabled={res.page>=res.totalPages} onClick={()=>search(res.page+1)}>Next<ChevronRight/></button></div>
  </section>}
 </>
}

function SavedData({setMsg}){const[startDate,setStartDate]=useState(''),[endDate,setEndDate]=useState(''),[data,setData]=useState(),[loading,setLoading]=useState(false);async function load(){setLoading(true);setMsg('');try{const{data:x}=await api.get('/data/saved',{params:{startDate,endDate,page:1,pageSize:50}});setData(x)}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}async function download(){setMsg('');try{const r=await api.get('/data/export',{params:{startDate,endDate},responseType:'blob'});const blob=new Blob([r.data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`KV_Added_Data${startDate||endDate?`_${startDate||'all'}_to_${endDate||'all'}`:'_all'}.xlsx`;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}catch(e){if(e.response?.data instanceof Blob){try{const t=await e.response.data.text();setMsg(JSON.parse(t).message)}catch{setMsg('Excel export failed.')}}else setMsg(e.response?.data?.message||e.message)}}return <section className="panel"><div className="panel-intro"><div className="intro-icon"><Download/></div><div><span className="eyebrow">DATABASE EXPORT</span><h3>Added rows &amp; Excel download</h3><p>Filter saved records by date and export them using the same 22-column KV workbook structure.</p></div></div><div className="filterbar"><div className="field"><label><CalendarDays/> From date</label><input value={startDate} onChange={e=>setStartDate(e.target.value)} placeholder="DD-MM-YYYY"/></div><div className="field"><label><CalendarDays/> To date</label><input value={endDate} onChange={e=>setEndDate(e.target.value)} placeholder="DD-MM-YYYY"/></div><button className="secondary-btn" onClick={load} disabled={loading}><Search/>{loading?'Filtering...':'Filter'}</button><button className="primary-btn" onClick={download}><FileSpreadsheet/>Download Excel</button></div>{data&&<div className="savedresult"><strong>{data.total}</strong> saved row(s) match the selected date range.</div>}{data?.rows?.length>0&&<div className="table-shell saved-table"><table className="data-table"><thead><tr>{cols.map(x=><th key={x[0]} title={x[1]}>{x[1]}</th>)}</tr></thead><tbody>{data.rows.map(x=><tr key={x._id}>{cols.map(([k])=><td key={k}>{k==='sn'?x.row?.sn:k==='date'?x.searchDate:k==='day'?x.weekday:(x.row?.[k]||'—')}</td>)}</tr>)}</tbody></table></div>}</section>}

function Upload({status,refresh,setMsg}){const[tf,setTf]=useState(),[kf,setKf]=useState(),[pw,setPw]=useState(''),[busy,setBusy]=useState(false);const initialized=!!status?.initialized;async function submit(){if(!tf&&!kf)return setMsg('Select at least one Excel file.');if(initialized&&(!tf&&!kf))return setMsg('Select an updated file.');if(!pw)return setMsg('Enter the upload password.');setBusy(true);try{const{data:a}=await api.post('/upload/authenticate',{password:pw}),f=new FormData();if(tf)f.append('trainsFile',tf);if(kf)f.append('kavachFile',kf);if(!initialized){if(!tf||!kf){setMsg('For first setup, select both source files.');setBusy(false);return}await api.post('/upload/initial',f,{headers:{Authorization:'Bearer '+a.token}})}else await api.post('/upload/update',f,{headers:{Authorization:'Bearer '+a.token}});setMsg('Source data updated successfully.');setPw('');setTf();setKf();await refresh()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}async function download(key){setMsg('');try{const r=await api.get(`/upload/download/${key}`,{responseType:'blob'});const cd=r.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||(key==='trains'?'TRAINS_PER_DAY.xlsx':'Kavach_Loco_Details.xlsx');const url=URL.createObjectURL(new Blob([r.data]));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}catch(e){if(e.response?.data instanceof Blob){try{const t=await e.response.data.text();setMsg(JSON.parse(t).message)}catch{setMsg('Source download failed.')}}else setMsg(e.response?.data?.message||e.message)}}return <section className="panel"><div className="panel-intro"><div className="intro-icon"><Database/></div><div><span className="eyebrow">SOURCE DATA MANAGEMENT</span><h3>Upload &amp; maintain source workbooks</h3><p>Uploading requires the configured password. Downloading the currently stored workbooks is available without a password.</p></div></div><div className="source-grid"><SourceCard title="TRAINS_PER_DAY.xlsx" file={tf} set={setTf} status={status?.train} onDownload={()=>download('trains')}/><SourceCard title="Kavach_Loco_Details.xlsx" file={kf} set={setKf} status={status?.kavach} onDownload={()=>download('kavach')}/></div><section className="security"><div className="security-head"><div className="security-icon"><LockKeyhole/></div><div><span className="eyebrow">UPLOAD SECURITY</span><h4>{initialized?'Password required for source updates':'Password required for initial upload'}</h4></div></div><div className="security-form"><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Enter upload password"/><button className="primary-btn" onClick={submit} disabled={busy}><LockKeyhole/>{busy?'Processing...':initialized?'Authenticate & Update':'Authenticate & Upload'}</button></div><div className="secure"><CheckCircle2/> Upload is password-protected. Source downloads do not require a password.</div></section></section>}
function SourceCard({title,file,set,status,onDownload}){return <section className="source-card"><div className="source-card-top"><FileSpreadsheet/><span>{status?.count??0} rows</span></div><h4>{title}</h4><p>{status?.fileName?`Stored: ${status.fileName}`:'Not uploaded yet'}</p><label className="picker"><UploadCloud/><span>{file?file.name:'Choose .xlsx / .xls / .xlsm file'}</span><input type="file" accept=".xlsx,.xls,.xlsm" onChange={e=>set(e.target.files?.[0])}/></label><button className="download-source" onClick={onDownload} disabled={!status}><FileDown/>Download stored file</button></section>}
createRoot(document.getElementById('root')).render(<App/>);
