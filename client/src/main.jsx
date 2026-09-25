import React,{useEffect,useMemo,useRef,useState}from'react';
import{createRoot}from'react-dom/client';
import{motion}from'framer-motion';
import{CalendarDays,CheckCircle2,Database,Download,FileSpreadsheet,LockKeyhole,RefreshCw,Search,UploadCloud,XCircle,LoaderCircle,Menu,ShieldCheck,TrainFront,FileDown,Edit3,Save,BarChart3,Filter,ExternalLink,LogOut,UserRound,KeyRound,ShieldAlert,Check,X,Printer}from'lucide-react';
import axios from'axios';
import'./styles.css';

const configuredApi=String(import.meta.env.VITE_API_URL||'http://localhost:5000/api').replace(/\/+$/,'');
const api=axios.create({baseURL:configuredApi.endsWith('/api')?configuredApi:`${configuredApi}/api`,withCredentials:true});

const cols=[
 ['sn','S.N'],['date','DATE'],['day','DAY'],['kavachSection','KAVACH_WKG_SECTION'],['dir','DIR'],['trType','TR_TYPE'],['time','T_O_Time'],['attachedDiv','LOCO_ATTACHED_DIV.'],['fromTo','From_To'],['locoLink','LOCO_link'],['locoLinkDiv','LOCO_LINK_DIV'],['dayWork','Day_of_wkg_in_territory'],['locoNo','Loco_No'],['shed','SHED'],['locoType','Loco_Type'],['rly','RLY'],['kavachMake','Loco_Kavach_Make'],['brakeSystem','Brake_system'],['trainNo','TRAIN_No'],['worked','Train_worked_with_Kavach_YES_NO'],['noReason','If_No_the_Reason'],['remarks','Remarks'],['shedRemark','Shed_Remark'],['oem','OEM']
];
const dashboardCols=cols.filter(([k])=>k!=='shedRemark'&&k!=='oem');
// Remark pages intentionally use a lean, responsive subset of Dashboard data.
// These fields were removed to avoid unnecessary horizontal scrolling while
// preserving the operational fields needed by Shed/OEM users.
const REMARK_HIDDEN_KEYS=new Set(['day','attachedDiv','locoLink','locoLinkDiv','trType','locoType','rly']);
const roleCols=(isShed)=>{
  const filtered=cols
    .filter(([k])=>!REMARK_HIDDEN_KEYS.has(k)&&k!==(isShed?'oem':'shedRemark'))
    .map(([k,label])=>k==='date'?[k,'Date run Kota div']:[k,label]);
  // Keep the editable remark immediately beside the date on BOTH dedicated pages.
  const dateIndex=filtered.findIndex(([k])=>k==='date');
  const targetKey=isShed?'shedRemark':'oem';
  const target=filtered.find(([k])=>k===targetKey);
  const withoutTarget=filtered.filter(([k])=>k!==targetKey);
  withoutTarget.splice(Math.max(0,dateIndex+1),0,target);
  return withoutTarget;
};
const editableKeys=new Set(['locoNo','worked','noReason','remarks']);
const derivedKeys=new Set(['shed','locoType','rly','kavachMake','brakeSystem']);

function dmyToISO(value){const m=String(value||'').match(/^(\d{2})-(\d{2})-(\d{4})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:''}
function isoToDMY(value){const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:''}
function DatePicker({value,onChange,label}){return <div className="field"><label><CalendarDays/> {label}</label><div className="date-input-wrap"><input type="date" value={dmyToISO(value)} onChange={e=>onChange(isoToDMY(e.target.value))}/></div></div>}
function getInitialPage(){
 const path=window.location.pathname;
 if(path==='/login')return'login'; if(path==='/dashboard')return'dash'; if(path==='/stored-data')return'saved';
 if(path==='/summary')return'summary'; if(path==='/data-sources')return'src'; if(path==='/shed-remark')return'shed';
 if(path==='/oem-remark')return'oem'; if(path==='/summary-detail')return'detail'; if(path==='/password-reset-requests')return'resets';
 const p=new URLSearchParams(window.location.search);
 if(p.get('page')==='saved')return'saved'; if(p.get('page')==='summary')return'summary';
 return'summary';
}
const PAGE_PATH={login:'/login',dash:'/dashboard',saved:'/stored-data',summary:'/summary',detail:'/summary-detail',src:'/data-sources',shed:'/shed-remark',oem:'/oem-remark',resets:'/password-reset-requests'};
function navigatePage(page,params={},replace=false){
 const clean=Object.fromEntries(Object.entries(params).filter(([,v])=>v!==undefined&&v!==null&&v!==''));
 const q=new URLSearchParams(clean),path=PAGE_PATH[page]||'/dashboard',url=q.toString()?`${path}?${q}`:path;
 (replace?window.history.replaceState:window.history.pushState).call(window.history,{},'',url);
 // pushState/replaceState do not fire popstate. Notify the app so internal
 // navigation changes the rendered page without reloading the Vercel SPA.
 window.dispatchEvent(new PopStateEvent('popstate'));
}

function PublicShell({tab,setTab,msg,setMsg}){
 const [mobileNav,setMobileNav]=useState(false);
 const nav=[['summary',<BarChart3/>,'Summary'],['detail',<FileSpreadsheet/>,'Summary & Detail'],['saved',<Download/>,'Saved Data / Export']];
 function change(id){setTab(id);navigatePage(id);setMsg('');setMobileNav(false)}
 function goLogin(){setMobileNav(false);setTab('login');navigatePage('login',{},true)}
 return <div className="app-shell">
  <header className="gov-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div><div className="header-account public-account"><button className="login-header-btn desktop-login" onClick={goLogin}><LogOut/> Login</button><button className="mobile-menu public-mobile-menu" aria-label={mobileNav?'Close menu':'Open menu'} aria-expanded={mobileNav} onClick={()=>setMobileNav(v=>!v)}><Menu/></button></div></div>
   <div className="nav-strip"><div className="nav-inner">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>change(id)}>{icon}{label}</button>)}<div className="nav-spacer"/><div className="official-chip"><ShieldCheck/> PUBLIC ACCESS</div></div></div>
   {mobileNav&&<div className="mobile-nav public-mobile-nav">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'active':''} onClick={()=>change(id)}>{icon}{label}</button>)}<button onClick={goLogin}><LogOut/>Login</button></div>}
  </header>
  <main className="content">{msg&&<div className="app-message">{msg}</div>}{tab==='summary'?<Summary setMsg={setMsg}/>:tab==='detail'?<SummaryDetail setMsg={setMsg}/>:<SavedData setMsg={setMsg}/>}</main>
  <footer className="footer"><div>West Central Railway · Kota Division · Traction Operation</div><span>@mak</span></footer>
 </div>
}

function App(){
 const [user,setUser]=useState(null),[checking,setChecking]=useState(true),[tab,setTab]=useState(getInitialPage),[date,setDate]=useState('11-09-2026'),[res,setRes]=useState(),[msg,setMsg]=useState(''),[status,setStatus]=useState(),[loading,setLoading]=useState(false),[mobileNav,setMobileNav]=useState(false),[options,setOptions]=useState({workedOptions:[],reasonOptions:[]}),[refreshKey,setRefreshKey]=useState(0),[changePassword,setChangePassword]=useState(false);
 const statusLoad=async()=>{try{const{data}=await api.get('/upload/status');setStatus(data)}catch(e){if(e.response?.status!==401)setMsg('Backend is not reachable. Check the deployed API or start the server.')}};
 const loadOptions=async()=>{try{const{data}=await api.get('/data/options');setOptions(data)}catch{setOptions({workedOptions:[],reasonOptions:[]})}};
 const checkSession=async()=>{try{const{data}=await api.get('/auth/me');setUser(data.user);return data.user}catch{setUser(null);return null}finally{setChecking(false)}};
 useEffect(()=>{checkSession();const onPop=()=>setTab(getInitialPage());window.addEventListener('popstate',onPop);return()=>window.removeEventListener('popstate',onPop)},[]);
 useEffect(()=>{if(user?.role==='admin'){statusLoad();loadOptions()}},[user]); useEffect(()=>{if(!user)return;const allowed=user.role==='admin'?['dash','saved','summary','detail','src','resets']:user.role==='shed'?['shed']:['oem'];if(!allowed.includes(tab)){const home=user.role==='admin'?'dash':user.role==='shed'?'shed':'oem';setTab(home);navigatePage(home,{},true)}},[user,tab]);
 async function logout(){try{await api.post('/auth/logout')}catch{}try{localStorage.removeItem('token');localStorage.removeItem('role');localStorage.removeItem('user');sessionStorage.clear()}catch{}setUser(null);setRes();setMsg('');navigatePage('login',{},true);setTab('login');setMobileNav(false);window.scrollTo(0,0)}
 async function search(){setLoading(true);setMsg('');try{const{data}=await api.get('/data/search',{params:{date}});setRes(data);await loadOptions()}catch(e){setRes();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 function patchRow(sourceRow,patch){setRes(x=>x?{...x,rows:x.rows.map(r=>r.sourceRow===sourceRow?{...r,...patch}:r)}:x)}
 async function lookup(row){const locoNo=String(row.locoNo||'').trim().slice(0,6);patchRow(row.sourceRow,{locoNo,kavachFound:false,kavachLoading:Boolean(locoNo),kavachError:''});if(!locoNo){patchRow(row.sourceRow,{kavachLoading:false,kavachError:'Enter Loco_No'});return}try{const{data}=await api.get('/kavach/lookup',{params:{locoNo}});patchRow(row.sourceRow,{...data,kavachFound:true,kavachLoading:false,kavachError:''})}catch(e){patchRow(row.sourceRow,{kavachFound:false,kavachLoading:false,kavachError:e.response?.data?.message||'Loco_No not found in Kavach_Loco_Details.xlsx.'})}}
 async function saveRow(row){if(!String(row.locoNo||'').trim())return setMsg('Enter Loco_No before saving.');if(!row.kavachFound)return setMsg('Enter a valid Loco_No and complete the Kavach lookup.');if(!String(row.worked||'').trim())return setMsg('Please select Train Worked with Kavach before updating the record.');if(String(row.worked||'').trim().toUpperCase()==='NO'&&!String(row.noReason||'').trim())return setMsg('Please select If No – Reason before updating the record.');if(String(row.remarks||'').length>500)return setMsg('Remarks cannot exceed 500 characters.');try{const{data}=await api.post('/data/rows',{searchDate:res.date,weekday:res.weekday,row});patchRow(row.sourceRow,{...data.row,added:true,editing:false,kavachFound:true});setMsg(`Loco ${data.row.locoNo} record ${row.added?'updated':'saved'} successfully.`)}catch(e){setMsg(e.response?.data?.message||e.message)}}
 function beginModify(row){patchRow(row.sourceRow,{editing:true});setMsg(`Modify mode enabled for Loco ${row.locoNo||'record'}.`)}
 function changeTab(id){setTab(id);navigatePage(id);setMobileNav(false);setMsg('')}
 const title=tab==='dash'?'Daily Traction Roster':tab==='saved'?'Saved Data & Excel Export':tab==='summary'?'Summary':tab==='detail'?'Summary & Detail':tab==='src'?'Source Data Management':tab==='shed'?'Shed Remark':tab==='oem'?'OEM Remark':'Password Reset Requests';
 if(checking)return <div className="auth-loading"><LoaderCircle className="spin"/><span>Checking secure session...</span></div>;
 if(!user){
   if(tab==='summary'||tab==='saved'||tab==='detail') return <PublicShell tab={tab} setTab={setTab} msg={msg} setMsg={setMsg}/>;
   return <Login onHome={()=>{window.location.assign('/')}} onLogin={u=>{setUser(u);const p=u.role==='admin'?'dash':u.role==='shed'?'shed':'oem';setTab(p);navigatePage(p,{},true)}}/>;
 }
 const isAdmin=user.role==='admin';
 const nav=isAdmin?[['dash',<TrainFront/>,'Dashboard'],['saved',<Download/>,'Saved Data / Export'],['summary',<BarChart3/>,'Summary'],['detail',<FileSpreadsheet/>,'Summary & Detail'],['src',<Database/>,'Data Sources'],['resets',<KeyRound/>,'Password Reset Requests']]:user.role==='shed'?[['shed',<Database/>,'Shed Remark']]:[['oem',<Database/>,'OEM Remark']];
 return <div className="app-shell">
  <header className="gov-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div><div className="header-account"><button className="user-chip user-action" onClick={()=>setChangePassword(true)} title="Change password"><UserRound/>{user.username}<KeyRound/></button><button className="logout-btn" onClick={logout}><LogOut/>Logout</button><button className="mobile-menu" onClick={()=>setMobileNav(v=>!v)} aria-label="Open navigation"><Menu/></button></div></div>
   <div className="nav-strip"><div className="nav-inner">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>changeTab(id)}>{icon}{label}</button>)}<div className="nav-spacer"/><div className="official-chip"><ShieldCheck/> {user.role.toUpperCase()} · Official Workflow</div></div></div>
   {mobileNav&&<div className="mobile-nav">{nav.map(([id,,label])=><button key={id} className={tab===id?'active':''} onClick={()=>changeTab(id)}>{label}</button>)}<button className="mobile-logout" onClick={logout}><LogOut/> Logout</button></div>}
  </header>
  <main className="content"><div className="page-heading"><div><span>TRACTION OPERATION</span><h2>{title}</h2></div><button className="refresh-btn" onClick={()=>setRefreshKey(v=>v+1)} title="Refresh"><RefreshCw/></button></div>
   {tab==='dash'?<Dashboard key={refreshKey} refreshSignal={refreshKey} date={date} setDate={setDate} res={res} search={search} loading={loading} msg={msg} saveRow={saveRow} patchRow={patchRow} lookup={lookup} beginModify={beginModify} options={options}/>:tab==='saved'?<SavedData key={refreshKey} setMsg={setMsg}/>:tab==='summary'?<Summary key={refreshKey} setMsg={setMsg}/>:tab==='detail'?<SummaryDetail key={refreshKey} setMsg={setMsg}/>:tab==='src'?<Upload key={refreshKey} status={status} refresh={()=>{statusLoad();loadOptions()}} setMsg={setMsg}/>:tab==='shed'?<RemarkPage key={refreshKey} role="shed" setMsg={setMsg}/>:tab==='oem'?<RemarkPage key={refreshKey} role="oem" setMsg={setMsg}/>:<ResetRequests key={refreshKey} setMsg={setMsg}/>} 
   {msg&&tab!=='dash'&&<div className="alert"><XCircle/>{msg}</div>}
  </main>{changePassword&&<ChangePasswordModal onClose={()=>setChangePassword(false)} onDone={()=>{setChangePassword(false);logout()}}/>}<footer className="footer"><div>West Central Railway · Kota Division · Traction Operation</div><span>@mak</span></footer>
 </div>
}
function Login({onLogin,onHome}){
 const[username,setUsername]=useState(''),[password,setPassword]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState(''),[forgot,setForgot]=useState(false);
 async function submit(e){e.preventDefault();setBusy(true);setError('');try{const{data}=await api.post('/auth/login',{username,password});onLogin(data.user)}catch(e){setError(e.response?.data?.message||'Invalid username or password. Please try again.')}finally{setBusy(false)}}
 return <div className="login-shell"><header className="gov-header login-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div></div></header><main className="login-main"><form className="login-card" onSubmit={submit}><div className="login-icon"><ShieldCheck/></div><span className="eyebrow">SECURE LOGIN</span><h2>Sign in</h2><p className="login-subtitle">Access the Kavach Monitoring System using your authorized account.</p>{error&&<div className="login-error"><ShieldAlert/>{error}</div>}<div className="field"><label><UserRound/> Username</label><input autoComplete="username" value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter username"/></div><div className="field"><label><KeyRound/> Password</label><input type="password" autoComplete="current-password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password"/></div><button type="button" className="forgot-link" onClick={()=>setForgot(true)}>Forgot Password?</button><button className="primary-btn login-btn" disabled={busy||!username||!password}>{busy?<><LoaderCircle className="spin"/>Signing in...</>:<><ShieldCheck/>Login</>}</button><button type="button" className="secondary-btn back-home-btn" onClick={onHome}><span aria-hidden="true">←</span> Back to Home</button></form></main>{forgot&&<ResetRequestModal onClose={()=>setForgot(false)}/>}<footer className="footer login-footer"><div>West Central Railway · Kota Division · Traction Operation</div><span>@mak</span></footer></div>
}
function ResetRequestModal({onClose}){
 const[role,setRole]=useState('shed'),[username,setUsername]=useState(''),[mobile,setMobile]=useState(''),[busy,setBusy]=useState(false),[msg,setMsg]=useState('');
 async function submit(){setBusy(true);setMsg('');try{const{data}=await api.post('/auth/reset-request',{username,role,mobileNumber:mobile});setMsg(data.message)}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 return <div className="modal-backdrop"><div className="modal-card"><button className="modal-close" onClick={onClose}><X/></button><KeyRound className="modal-icon"/><h3>Request Admin Reset</h3><p>Enter your username, role and mobile number. The Admin will review the request.</p><div className="field"><label>Role</label><select value={role} onChange={e=>setRole(e.target.value)}><option value="shed">Shed User</option><option value="oem">OEM User</option></select></div><div className="field"><label>Username</label><input value={username} onChange={e=>setUsername(e.target.value)} placeholder="e.g. AJJE / MEDHA"/></div><div className="field"><label>Mobile Number</label><input value={mobile} onChange={e=>setMobile(e.target.value)} inputMode="tel" maxLength={15} placeholder="Mobile number"/></div>{msg&&<div className="modal-message">{msg}</div>}<button className="primary-btn" onClick={submit} disabled={busy||!username||!mobile}>{busy?'Submitting...':'Submit Request'}</button></div></div>
}
function RemarkPage({role,setMsg}){
 const[data,setData]=useState(),[loading,setLoading]=useState(true),[saving,setSaving]=useState('');
 const isShed=role==='shed';
 const shellRef=useRef(null),tableRef=useRef(null);
 async function load(){setLoading(true);try{const{data:x}=await api.get('/remark-records');setData(x)}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 useEffect(()=>{load()},[]);

 async function update(id,remark){setSaving(id);try{const{data:x}=await api.patch(`/remark-records/${id}`,{remark});setData(d=>({...d,rows:d.rows.map(r=>r._id===id?x.row:r)}));setMsg(x.message);return true}catch(e){setMsg(e.response?.data?.message||e.message);return false}finally{setSaving('')}}
 async function closeRecord(id){
  if(!window.confirm('Are you sure you want to close this record? Once closed, it will no longer appear on this page.')) return;
  setSaving(id);
  try{
   const{data:x}=await api.patch(`/remark-records/${id}/close`);
   setData(d=>({...d,rows:d.rows.filter(r=>r._id!==id)}));
   setMsg(x.message||'Record closed successfully.');
  }catch(e){setMsg(e.response?.data?.message||e.message)}finally{setSaving('')}
 }
 return <section className="panel role-panel"><div className="panel-intro compact"><div className="intro-icon"><Database/></div><div><span className="eyebrow">{isShed?'SHED WORKFLOW':'OEM WORKFLOW'}</span><h3>{isShed?'Shed Remark':'OEM Remark'}</h3><p>Logged in as <strong>{data?.user?.username||'—'}</strong>. Filters are automatically applied by the server. Only the relevant operational columns are shown to keep this page responsive; the other role's remark field is hidden. Only the {isShed?'Shed Remark':'OEM'} field is editable.</p></div></div>{loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading relevant records...</div>:data?.rows?.length?<><div className="role-table-shell" ref={shellRef}><table ref={tableRef} className="data-table role-table role-full-table"><thead><tr>{roleCols(isShed).map(([key,label])=><th key={key} className={key==='worked'?'worked-head':key==='noReason'?'reason-head':''}>{key==='worked'?<>Train Worked<br/>with Kavach</>:key==='noReason'?'If No – Reason':label}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>{data.rows.map((x,i)=><RemarkRow key={x._id} record={x} index={i} isShed={isShed} saving={saving} onUpdate={update} onClose={closeRecord}/>)}</tbody></table></div></>:<div className="empty-panel">No records currently match the automatic {isShed?'Shed':'OEM'} filters.</div>}</section>
}
function RemarkRow({record,index,isShed,saving,onUpdate,onClose}){
 const r=record.row||{};
 const initialRemark=isShed?(r.shedRemark||record.shedRemark||''):(r.oem||record.oem||'');
 const[remark,setRemark]=useState(initialRemark),[editing,setEditing]=useState(!initialRemark);
 const[hasSavedRemark,setHasSavedRemark]=useState(Boolean(initialRemark));
 const closed=isShed?Boolean(record.shedRemarkClosed):Boolean(record.oemRemarkClosed);
 useEffect(()=>{const next=isShed?(r.shedRemark||record.shedRemark||''):(r.oem||record.oem||'');setRemark(next);setHasSavedRemark(Boolean(next));setEditing(!next)},[record._id,record.updatedAt,isShed]);
 const valueFor=(k)=>k==='sn'?index+1:k==='date'?record.searchDate:k==='day'?record.weekday:(r[k]??'');
 return <tr>{roleCols(isShed).map(([k,label])=>{
   if(k==='shedRemark' || k==='oem'){
     const isTarget=(isShed&&k==='shedRemark')||(!isShed&&k==='oem');
     const v=isTarget?remark:(r[k]??'');
     return <td key={k} className="new-field-cell">{isTarget&&editing?<textarea className="role-remark-input" maxLength={isShed?300:200} value={v} onChange={e=>setRemark(e.target.value)} placeholder={isShed?'Enter shed remark':'Enter OEM remark'}/>:<span className="wrap-text">{v||'—'}</span>}</td>;
   }
   return <td key={k} className={k==='locoNo'?'loco-cell':''}><span className="wrap-text">{valueFor(k)||'—'}</span></td>;
 })}<td className="action-cell"><div className="remark-actions"><button className={`update ${editing?'':'modify'}`} disabled={saving===record._id} onClick={async()=>{if(editing){if(await onUpdate(record._id,remark)){setHasSavedRemark(true);setEditing(false)}}else setEditing(true)}}>{saving===record._id?<LoaderCircle className="spin"/>:editing?<Save/>:<Edit3/>}{saving===record._id?'Saving...':editing?'Update':hasSavedRemark?'Modify':'Update'}</button>{hasSavedRemark&&!closed&&<button type="button" className="close-btn" disabled={saving===record._id} onClick={()=>onClose(record._id)} title="Close this record"><X/></button>}</div></td></tr>
}
function ChangePasswordModal({onClose,onDone}){
 const[oldPassword,setOldPassword]=useState(''),[newPassword,setNewPassword]=useState(''),[confirmPassword,setConfirmPassword]=useState(''),[busy,setBusy]=useState(false),[msg,setMsg]=useState('');
 async function submit(){setBusy(true);setMsg('');try{const{data}=await api.post('/auth/change-password',{oldPassword,newPassword,confirmPassword});setMsg(data.message);setTimeout(onDone,700)}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 return <div className="modal-backdrop"><div className="modal-card"><button className="modal-close" onClick={onClose}><X/></button><KeyRound className="modal-icon"/><h3>Change Password</h3><p>Enter your current password and choose a new secure password.</p><div className="field"><label>Old Password</label><input type="password" value={oldPassword} onChange={e=>setOldPassword(e.target.value)}/></div><div className="field"><label>New Password</label><input type="password" value={newPassword} onChange={e=>setNewPassword(e.target.value)} placeholder="Minimum 6 characters"/></div><div className="field"><label>Confirm New Password</label><input type="password" value={confirmPassword} onChange={e=>setConfirmPassword(e.target.value)}/></div>{msg&&<div className="modal-message">{msg}</div>}<button className="primary-btn" onClick={submit} disabled={busy||!oldPassword||!newPassword||!confirmPassword}>{busy?'Changing...':'Change Password'}</button></div></div>
}
function ResetRequests({setMsg}){
 const[data,setData]=useState({requests:[]}),[loading,setLoading]=useState(true),[busy,setBusy]=useState('');
 async function load(){setLoading(true);try{const{data:x}=await api.get('/auth/reset-requests');setData(x)}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 useEffect(()=>{load()},[]);
 async function action(id,type){if(!window.confirm(type==='approve'?'Approve and reset this user password?':'Reject this password reset request?'))return;setBusy(id);try{const{data:x}=await api.post(`/auth/reset-requests/${id}/${type}`);setMsg(x.message);await load()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy('')}}
 return <section className="panel role-panel"><div className="panel-intro compact"><div className="intro-icon"><KeyRound/></div><div><span className="eyebrow">ADMIN SECURITY</span><h3>Password Reset Requests</h3><p>Review and process Shed/OEM reset requests.</p></div></div>{loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading requests...</div>:data.requests.length?<div className="role-table-shell"><table className="data-table role-table"><thead><tr><th>REQUEST ID</th><th>USERNAME</th><th>ROLE</th><th>SHED/OEM</th><th>MOBILE NUMBER</th><th>REQUEST DATE & TIME</th><th>STATUS</th><th>SMS</th><th>ACTION</th></tr></thead><tbody>{data.requests.map(r=><tr key={r.requestId}><td>{r.requestId}</td><td>{r.username}</td><td>{r.role}</td><td>{r.username}</td><td>{r.mobileNumber}</td><td>{new Date(r.requestedAt).toLocaleString()}</td><td>{r.status}</td><td>{r.smsStatus}</td><td>{r.status==='Pending'?<div className="request-actions"><button className="update approve" disabled={busy===r.requestId} onClick={()=>action(r.requestId,'approve')}><Check/>Approve</button><button className="update reject" disabled={busy===r.requestId} onClick={()=>action(r.requestId,'reject')}><X/>Reject</button></div>:'—'}</td></tr>)}</tbody></table></div>:<div className="empty-panel">No password reset requests found.</div>}</section>
}
function Dashboard({refreshSignal,date,setDate,res,search,loading,msg,saveRow,patchRow,lookup,beginModify,options}){
 const workedOptions=options.workedOptions||[],reasonOptions=options.reasonOptions||[]; useEffect(()=>{if(refreshSignal>0&&date)search()},[refreshSignal]);
 return <>
  <section className="search-card"><DatePicker label="Date" value={date} onChange={setDate}/><button className="primary-btn" onClick={search} disabled={loading||!date}><Search/>{loading?'Fetching...':'Fetch Data'}</button></section>
  {msg&&<div className="alert"><XCircle/>{msg}</div>}
  {res&&<section className="results-card dashboard-results"><div className="result-head"><div><span className="eyebrow">DAILY ROSTER</span><h3>{res.weekday} <i>·</i> {res.date}</h3></div><div className="result-meta">{res.total} records</div></div>
   <div className="table-shell dashboard-table-shell"><table className="data-table"><thead><tr>{dashboardCols.map(([key,label])=><th key={key} className={key==='worked'?'worked-head':key==='noReason'?'reason-head':''} title={label}>{key==='worked'?<>Train Worked<br/>with Kavach</>:key==='noReason'?'If No – Reason':label}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>
    {res.rows.length===0?<tr><td colSpan={dashboardCols.length+1} className="empty-cell">No train records found for the selected date.</td></tr>:res.rows.map((r,i)=><motion.tr initial={{opacity:0,y:5}} animate={{opacity:1,y:0}} transition={{duration:.18,delay:i*.01}} key={r.sourceRow}>
     {dashboardCols.map(([k,label])=><td key={k} className={`${k==='locoNo'?'loco-cell ':''}${k==='remarks'?'remarks-cell':''}${k==='shedRemark'||k==='oem'?' new-field-cell':''}`}><CellEditor k={k} label={label} row={r} patchRow={patchRow} lookup={lookup} workedOptions={workedOptions} reasonOptions={reasonOptions}/></td>)}
     <td className="action-cell">{r.added&&!r.editing?<button className="update modify" onClick={()=>beginModify(r)}><Edit3/>Modify</button>:<button className="update" disabled={!String(r.locoNo||'').trim()||!r.kavachFound||r.kavachLoading||!String(r.worked||'').trim()||(String(r.worked||'').trim().toUpperCase()==='NO'&&!String(r.noReason||'').trim())} onClick={()=>saveRow(r)}>{r.editing?<Save/>:<Database/>}{r.editing?'Update':r.kavachLoading?'Looking up...':r.kavachFound?'Update':'Enter Loco_No'}</button>}</td>
    </motion.tr>)}
   </tbody></table></div>
  </section>}
 </>
}

function CellEditor({k,label,row,patchRow,lookup,workedOptions,reasonOptions}){
 const editing=Boolean(row.editing),value=row[k]??'';
 if(k==='sn')return <b>{row.sn}</b>;if(k==='date')return <span>{row.date}</span>;if(k==='day')return <span>{row.day||'—'}</span>;
 if(k==='locoNo'){if(row.added&&!editing)return <span className="loco-value">{value||'—'}</span>;return <div className="locoeditor"><input className="loco-input" maxLength={6} inputMode="numeric" value={value} onChange={e=>patchRow(row.sourceRow,{locoNo:e.target.value.replace(/\D/g,'').slice(0,6),kavachFound:false,kavachError:''})} onBlur={()=>lookup(row)} onKeyDown={e=>{if(e.key==='Enter')lookup({...row,locoNo:e.currentTarget.value})}} placeholder="Loco No."/><button className="lookup" onClick={()=>lookup(row)} disabled={!String(value).trim()||row.kavachLoading} title="Find Loco_No in Kavach file">{row.kavachLoading?<LoaderCircle className="spin"/>:<Search/>}</button>{row.kavachError&&<small className="lookup-error">Not found</small>}</div>}
 if(k==='worked')return editing||!row.added?<select className="cell-select worked-select" value={value} onChange={e=>patchRow(row.sourceRow,{worked:e.target.value,noReason:String(e.target.value).trim().toUpperCase()==='YES'?'':row.noReason})}><option value="">Select</option>{workedOptions.map(v=><option key={v} value={v}>{v}</option>)}</select>:<span>{value||'—'}</span>;
 if(k==='noReason')return editing||!row.added?<select className="cell-select reason-select" value={value} onChange={e=>patchRow(row.sourceRow,{noReason:e.target.value})}><option value="">Select</option>{reasonOptions.map(v=><option key={v} value={v}>{v}</option>)}</select>:<span className="wrap-text">{value||'—'}</span>;
 if(k==='remarks')return editing||!row.added?<div className="remarks-wrap"><textarea maxLength={500} value={value} onChange={e=>patchRow(row.sourceRow,{remarks:e.target.value})} placeholder="Add remarks..."/><small>{String(value).length}/500</small></div>:<span className="wrap-text">{value||'—'}</span>;
 if((k==='shedRemark'||k==='oem')&&(editing||!row.added))return <input className="cell-editor new-field-input" value={value} maxLength={k==='shedRemark'?300:200} onChange={e=>patchRow(row.sourceRow,{[k]:e.target.value})} placeholder={k==='shedRemark'?'Shed remark':'OEM'}/>;
 if(editing&&editableKeys.has(k))return <input className="cell-editor" value={value} onChange={e=>patchRow(row.sourceRow,{[k]:e.target.value})}/>;
 if(derivedKeys.has(k))return <span className="derived-cell">{value||'—'}</span>;
 return <span className="wrap-text">{value||'—'}</span>;
}

function SavedData({setMsg}){
 const params=useMemo(()=>new URLSearchParams(window.location.search),[]),[startDate,setStartDate]=useState(params.get('date')||params.get('fromDate')||''),[endDate,setEndDate]=useState(params.get('endDate')||params.get('toDate')||params.get('date')||''),[dir,setDir]=useState(params.get('dir')||''),[metric,setMetric]=useState(params.get('metric')||''),[data,setData]=useState(),[loading,setLoading]=useState(false);
 const backSummary=params.get('fromSummary')==='1';
 async function load(overrides={}){const sd=overrides.startDate??startDate,ed=overrides.endDate??endDate,dr=overrides.dir??dir,mt=overrides.metric??metric;setLoading(true);setMsg('');try{const{data:x}=await api.get('/data/saved',{params:{startDate:sd,endDate:ed,dir:dr,metric:mt}});setData(x);setMetric(mt);navigatePage('saved',{...(sd?{date:sd}:{}),...(ed&&ed!==sd?{endDate:ed}:{}),...(dr?{dir:dr}:{}),...(mt?{metric:mt}:{}),...(backSummary?{fromSummary:'1',summaryFromDate:params.get('summaryFromDate')||'',summaryToDate:params.get('summaryToDate')||'',summaryDir:params.get('summaryDir')||''}:{})},true)}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 useEffect(()=>{if(startDate||dir)load()},[]);
 async function download(){setMsg('');try{const r=await api.get('/data/export',{params:{startDate,endDate,dir},responseType:'blob'});downloadBlob(r,'Traction_Operation_Data.xlsx')}catch(e){setMsg(await blobError(e,'Excel export failed.'))}}
 return <section className="panel"><div className="panel-intro compact"><div className="intro-icon"><Download/></div><div><span className="eyebrow">DATABASE EXPORT</span><h3>Saved records &amp; Excel download</h3><p>Filter saved records by date, DIR and export the matching data set.</p></div></div>{backSummary&&<div className="back-summary-row"><button className="secondary-btn" onClick={()=>window.history.back()}>← Back to Summary</button></div>}<div className="filterbar saved-filterbar"><DatePicker label="From date" value={startDate} onChange={setStartDate}/><DatePicker label="To date" value={endDate} onChange={setEndDate}/><div className="field"><label><Filter/> DIR</label><select className="top-select" value={dir} onChange={e=>setDir(e.target.value)}><option value="">All DIR</option><option value="UP">UP</option><option value="DN">DN</option></select></div><button className="secondary-btn" onClick={()=>load()} disabled={loading}><Search/>{loading?'Filtering...':'Filter'}</button><button className="primary-btn" onClick={download}><FileSpreadsheet/>Download Excel</button></div>{data&&<div className="savedresult"><strong>{data.total}</strong> saved row(s) match the selected filters.</div>}{data?.rows?.length>0&&<div className="table-shell saved-table"><table className="data-table"><thead><tr>{cols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':label}</th>)}</tr></thead><tbody>{data.rows.map((x,i)=><tr key={x._id}>{cols.map(([k])=><td key={k}>{k==='sn'?i+1:k==='date'?x.searchDate:k==='day'?x.weekday:<span className="wrap-text">{x.row?.[k]||'—'}</span>}</td>)}</tr>)}</tbody></table></div>}{data&&data.total===0&&<div className="empty-panel">No saved records match the selected date/DIR filters.</div>}</section>
}

function RecordsModal({open,onClose,fromDate,toDate,dir,metric,groupValue=''}){
 const [data,setData]=useState(null),[loading,setLoading]=useState(false),[error,setError]=useState('');
 useEffect(()=>{
  if(!open)return;
  let cancelled=false;
  (async()=>{setLoading(true);setError('');try{const{data:x}=await api.get('/summary-detail/records',{params:{fromDate,toDate,dir,metric,groupValue}});if(!cancelled)setData(x)}catch(e){if(!cancelled){setData(null);setError(e.response?.data?.message||e.message)}}finally{if(!cancelled)setLoading(false)}})();
  return()=>{cancelled=true};
 },[open,fromDate,toDate,dir,metric,groupValue]);
 if(!open)return null;
 return <div className="modal-backdrop records-modal-backdrop" role="dialog" aria-modal="true" aria-label="Corresponding records" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}>
  <div className="records-modal-card">
   <button className="modal-close" onClick={onClose} aria-label="Close"><X/></button>
   <div className="records-modal-head"><div><span className="eyebrow">CORRESPONDING DATABASE RECORDS</span><h3>{data?.metricLabel||'Corresponding Records'}</h3><p>{fromDate} to {toDate} · DIR: {dir||'ALL'} · <strong>{data?.total??'—'}</strong> record(s)</p></div></div>
   {loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading records...</div>:error?<div className="empty-panel">{error}</div>:data?.rows?.length?<div className="records-modal-table-shell"><table className="data-table records-modal-table"><thead><tr>{cols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':k==='date'?'Date run Kota div':label}</th>)}</tr></thead><tbody>{data.rows.map((x,i)=><tr key={x._id||i}>{cols.map(([k])=><td key={k}>{k==='sn'?i+1:k==='date'?x.searchDate:k==='day'?x.weekday:<span className="wrap-text">{x.row?.[k]||'—'}</span>}</td>)}</tr>)}</tbody></table></div>:<div className="empty-panel">No corresponding records found.</div>}
  </div>
 </div>
}

function SummaryCount({value,date,endDate,dir,metric,strong=false,summaryFromDate='',summaryToDate='',summaryDir='',onOpen}){
 const display=strong?<strong>{value}</strong>:value;
 const clickable=Number(value)>0;
 return <button type="button" disabled={!clickable} className={`count-link${strong?' count-link-strong':''}`} onClick={()=>clickable&&onOpen?.({fromDate:date,toDate:endDate,dir,metric})} title={clickable?`View ${metric} records`:'No records'}>{display}</button>;
}
function DetailCount({value,metric,scope='ALL',fromDate,toDate,onOpen,groupValue=''}){
 const clickable=Number(value)>0;
 return <button type="button" disabled={!clickable} className="count-link detail-count-link" onClick={()=>clickable&&onOpen?.({fromDate,toDate,dir:scope,metric,groupValue})}>{value||0}</button>;
}
function DetailMetricTable({data,onOpen}){
 const metrics=[['targetTrains','No of Target Trains'],['runningWithKavach','TOTAL NO. OF TRAIN RUNNING \nWITH KAVACH'],['workingCondition','TOTAL NO. OF TRAIN RUNNING \nWITH KAVACH IN WORKING CONDITION'],['nonKavach','NON KAVACH LOCO'],['defective','KAVACH DEFECTIVE LOCO'],['otherMake','OTHER MAKE'],['otherReason','OTHER REASON']];
 return <table className="excel-detail-table metric-report-table"><thead><tr><th className="sn-col"></th><th>DETAILS</th><th>TOTAL</th><th>UP</th><th>DN</th></tr></thead><tbody>{metrics.map(([k,label],i)=><tr key={k}><td className="sn-col">{i+1}</td><td className={k==='runningWithKavach'||k==='workingCondition'?'wrap-label':''}>{label.split('\n').map((x,j)=><React.Fragment key={j}>{j>0&&<br/>}{x}</React.Fragment>)}</td><td><DetailCount value={data.total[k]} metric={k} scope="ALL" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td><td><DetailCount value={data.up[k]} metric={k} scope="UP" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td><td><DetailCount value={data.dn[k]} metric={k} scope="DN" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td></tr>)}</tbody></table>
}
function DetailDivisionTable({title,rows,onOpen,fromDate,toDate,metric,showStation=false,groupByName=false}){
 return <section className="excel-detail-section"><h4>{title}</h4><table className="excel-detail-table division-report-table"><thead><tr><th className="sn-col"></th><th>{showStation?'Station':''}</th><th>TOTAL</th><th>UP</th><th>DN</th></tr></thead><tbody>{rows.length?rows.map((r,i)=><tr key={`${title}-${r.name}-${i}`}><td className="sn-col">{i+1}</td><td>{showStation?r.station||r.name:r.name}</td><td><DetailCount value={r.total} metric={metric} scope="ALL" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td><td><DetailCount value={r.up} metric={metric} scope="UP" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td><td><DetailCount value={r.dn} metric={metric} scope="DN" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td></tr>):<tr><td className="sn-col">—</td><td>No data</td><td>0</td><td>0</td><td>0</td></tr>}</tbody></table></section>
}
function DetailRecordsTable({title,rows,includeDate=false}){
 const headers=includeDate?['S.N','Date','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark']:['S.N','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark'];
 return <section className={`excel-detail-section records-report-section ${includeDate?'records-with-date':''}`}><h4>{title}</h4><div className="detail-record-scroll"><table className="excel-detail-table record-report-table"><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.length?rows.map(r=><tr key={`${title}-${r.sn}`}><td>{r.sn}</td>{includeDate&&<td>{r.date||''}</td>}<td>{r.trainNo||''}</td><td>{r.locoNo||''}</td><td>{r.shed||''}</td><td>{r.kavachMake||''}</td><td>{r.tlcRemark||''}</td><td>{r.shedRemark||''}</td><td>{r.oemRemark||''}</td></tr>):<tr><td>1</td><td colSpan={headers.length-1}>No data</td></tr>}</tbody></table></div></section>
}
function SummaryDetail({setMsg}){
 const params=new URLSearchParams(window.location.search); const [fromDate,setFromDate]=useState(params.get('fromDate')||''),[toDate,setToDate]=useState(params.get('toDate')||''),[data,setData]=useState(null),[loading,setLoading]=useState(false),[popup,setPopup]=useState(null);
 async function filter(){if(!fromDate||!toDate)return setMsg('Select From Date and To Date before filtering.');setLoading(true);setMsg('');try{const{data:x}=await api.get('/summary-detail',{params:{fromDate,toDate}});setData(x);navigatePage('detail',{fromDate,toDate})}catch(e){setData(null);setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 async function downloadPdf(){if(!data)return setMsg('Filter the report first.');try{const r=await api.get('/summary-detail/pdf',{params:{fromDate,toDate},responseType:'blob'});downloadBlobFile(r,'Summary_Detail.pdf','application/pdf')}catch(e){setMsg(await blobError(e,'PDF download failed.'))}}
 return <section className="panel detail-panel"><div className="panel-intro compact no-print"><div className="intro-icon"><FileSpreadsheet/></div><div><span className="eyebrow">DATE-WISE DETAIL REPORT</span><h3>Summary &amp; Detail</h3><p>Generate the report from existing database records for the selected date range.</p></div></div><div className="filterbar detail-filterbar no-print"><DatePicker label="From Date" value={fromDate} onChange={setFromDate}/><DatePicker label="To Date" value={toDate} onChange={setToDate}/><button className="secondary-btn" onClick={filter} disabled={loading||!fromDate||!toDate}><Search/>{loading?'Fetching...':'Filter'}</button>{data&&<><button className="primary-btn" type="button" onClick={downloadPdf}><FileDown/>Download PDF</button><button className="secondary-btn" type="button" onClick={()=>window.print()}><Printer/>Print A4</button></>}</div>{data&&<div className="summary-detail-a4"><div className="a4-report-title">Summary &amp; Detail - Date From&nbsp; {data.fromDate} &nbsp;To&nbsp; {data.toDate}</div><DetailMetricTable data={data} onOpen={setPopup}/><DetailDivisionTable title="KAVACH DEFECTIVE LOCO" rows={data.defectiveByShed} metric="defectiveByShed" groupByName fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailDivisionTable title="NON KAVACH LOCO (MISLINK BY LOCO LINK DIVISION)" rows={data.nonKavachLinkDivision} metric="nonKavachLink" groupByName fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailDivisionTable title="NON KAVACH LOCO (MISLINK BY LOCO ATTACH DIVISION)" rows={data.nonKavachAttachedDivision} metric="nonKavachAttach" groupByName fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailDivisionTable title="NON KAVACH LOCO" rows={data.nonKavachPlain} metric="nonKavachPlain" groupByName showStation fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailRecordsTable title="KAVACH DEFECTIVE LOCO DETAILS" rows={data.defectiveRows} includeDate/><DetailRecordsTable title="OTHER REASON" rows={data.otherRows} includeDate/></div>}{!data&&!loading&&<div className="detail-placeholder no-print">Select a date range and click <strong>Filter</strong> to fetch the report.</div>}<RecordsModal open={Boolean(popup)} onClose={()=>setPopup(null)} {...(popup||{})}/></section>
}
function Summary({setMsg}){
 const params=useMemo(()=>new URLSearchParams(window.location.search),[]),[fromDate,setFromDate]=useState(params.get('fromDate')||''),[toDate,setToDate]=useState(params.get('toDate')||''),[dir,setDir]=useState(params.get('dir')||'UP'),[data,setData]=useState(),[loading,setLoading]=useState(false),[hasFiltered,setHasFiltered]=useState(false),[popup,setPopup]=useState(null); useEffect(()=>{if(fromDate&&toDate&&dir)filter()},[]);
 async function filter(){if(!fromDate||!toDate||!dir)return setMsg('Select From Date, To Date and DIR before filtering.');setLoading(true);setMsg('');try{const{data:x}=await api.get('/summary',{params:{fromDate,toDate,dir}});setData(x);setHasFiltered(true);navigatePage('summary',{fromDate,toDate,dir})}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 async function download(){if(!data)return setMsg('Filter the summary first.');try{const r=await api.get('/summary/export',{params:{fromDate,toDate,dir},responseType:'blob'});downloadBlob(r,'Traction_Operation_Summary.xlsx')}catch(e){setMsg(await blobError(e,'Summary export failed.'))}}
 const cell=(value,r,metric,strong=false)=> <SummaryCount value={value} date={r.date} endDate={r.date} dir={r.dir} metric={metric} strong={strong} onOpen={setPopup}/>;
 return <section className="panel summary-panel"><div className="panel-intro compact"><div className="intro-icon"><BarChart3/></div><div><span className="eyebrow">DATE-WISE KAVACH SUMMARY</span><h3>Traction Operation Summary</h3><p>Generate the summary from saved database records using Date + DIR. The reference workbook is not used for calculations.</p></div></div><div className="filterbar summary-filterbar"><DatePicker label="From Date" value={fromDate} onChange={setFromDate}/><DatePicker label="To Date" value={toDate} onChange={setToDate}/><div className="field"><label><Filter/> DIR</label><select className="top-select" value={dir} onChange={e=>setDir(e.target.value)}><option value="ALL">All</option><option value="UP">UP</option><option value="DN">DN</option></select></div><button className="secondary-btn" onClick={filter} disabled={loading||!fromDate||!toDate}><Search/>{loading?'Generating...':'Filter'}</button><button className="primary-btn" onClick={download} disabled={!data}><FileSpreadsheet/>Download Summary</button></div>
 {hasFiltered&&data&&<><div className="summary-meta"><strong>{data.totalRecords}</strong> saved record(s) included for <strong>{data.dir}</strong> from <strong>{data.fromDate}</strong> to <strong>{data.toDate}</strong>.</div><div className="summary-table-wrap"><table className="summary-table"><thead><tr><th>DAY</th><th>DATE</th><th>DIR</th><th>TOTAL TRAINS LOCO IN KAVACH SECTION</th><th>NO. OF LOCO KAVACH FITTED</th><th>NO. OF LOCO KAVACH WORKING</th><th>KAVACH DEFECTIVE</th><th>KERNEX MAKE APPROVAL PENDING</th><th>DO NOT START REMARK/ STICKER</th><th>KAVACH ISOLATION PREVIOUS REMARKS</th><th>OTHER REASON</th><th>TOTAL</th><th>% WORKING</th></tr></thead><tbody>{data.rows.map(r=><tr key={`${r.date}-${r.dir}`}><td>{r.day}</td><td>{r.date}</td><td><span className="dir-pill">{r.dir}</span></td><td>{cell(r.totalTrains,r,'totalTrains')}</td><td>{cell(r.fitted,r,'fitted')}</td><td>{cell(r.working,r,'working')}</td><td>{cell(r.defective,r,'defective')}</td><td>{cell(r.kernex,r,'kernex')}</td><td>{cell(r.doNotStart,r,'doNotStart')}</td><td>{cell(r.isolation,r,'isolation')}</td><td>{cell(r.other,r,'other')}</td><td>{cell(r.total,r,'total',true)}</td><td><strong>{r.percentWorking}%</strong></td></tr>)}{data.totalRow&&<tr className="summary-total-row"><td><strong>TOTAL</strong></td><td></td><td><span className="dir-pill">{data.totalRow.dir}</span></td><td>{cell(data.totalRow.totalTrains,{date:data.fromDate,dir:data.dir},'totalTrains',true)}</td><td>{cell(data.totalRow.fitted,{date:data.fromDate,dir:data.dir},'fitted',true)}</td><td>{cell(data.totalRow.working,{date:data.fromDate,dir:data.dir},'working',true)}</td><td>{cell(data.totalRow.defective,{date:data.fromDate,dir:data.dir},'defective',true)}</td><td>{cell(data.totalRow.kernex,{date:data.fromDate,dir:data.dir},'kernex',true)}</td><td>{cell(data.totalRow.doNotStart,{date:data.fromDate,dir:data.dir},'doNotStart',true)}</td><td>{cell(data.totalRow.isolation,{date:data.fromDate,dir:data.dir},'isolation',true)}</td><td>{cell(data.totalRow.other,{date:data.fromDate,dir:data.dir},'other',true)}</td><td>{cell(data.totalRow.total,{date:data.fromDate,dir:data.dir},'total',true)}</td><td><strong>{data.totalRow.percentWorking}%</strong></td></tr>}</tbody></table></div>{data.rows.length===0&&<div className="empty-panel">No summary dates found.</div>}</>}
 <RecordsModal open={Boolean(popup)} onClose={()=>setPopup(null)} {...(popup||{})}/></section>
}

function downloadBlob(response,fallback){const cd=response.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||fallback;const url=URL.createObjectURL(new Blob([response.data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}
function downloadBlobFile(response,fallback,type){const cd=response.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||fallback;const url=URL.createObjectURL(new Blob([response.data],{type}));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)}
async function blobError(error,fallback){if(error.response?.data instanceof Blob){try{const text=await error.response.data.text();const json=JSON.parse(text);return json.message||fallback}catch{return fallback}}return error.response?.data?.message||error.message||fallback}

function Upload({status,refresh,setMsg}){const[tf,setTf]=useState(),[kf,setKf]=useState(),[pw,setPw]=useState(''),[busy,setBusy]=useState(false);const initialized=!!status?.initialized;
 async function submit(){if(!tf&&!kf)return setMsg('Select at least one Excel file.');if(!pw)return setMsg('Enter the upload password.');setBusy(true);try{const{data:a}=await api.post('/upload/authenticate',{password:pw}),f=new FormData();if(tf)f.append('trainsFile',tf);if(kf)f.append('kavachFile',kf);if(!initialized){if(!tf||!kf){setMsg('For first setup, select both source files.');setBusy(false);return}await api.post('/upload/initial',f,{headers:{Authorization:'Bearer '+a.token}})}else await api.post('/upload/update',f,{headers:{Authorization:'Bearer '+a.token}});setMsg('Source data updated successfully. Dropdown validation lists were refreshed from the workbook.');setPw('');setTf();setKf();await refresh()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 async function download(key){setMsg('');try{const r=await api.get(`/upload/download/${key}`,{responseType:'blob'});downloadBlob(r,key==='trains'?'TRAINS_PER_DAY.xlsx':'Kavach_Loco_Details.xlsx')}catch(e){setMsg(await blobError(e,'Source download failed.'))}}
 return <section className="panel"><div className="panel-intro compact"><div className="intro-icon"><Database/></div><div><span className="eyebrow">SOURCE DATA MANAGEMENT</span><h3>Upload &amp; maintain source workbooks</h3><p>Uploads require the configured password. Stored workbooks can be downloaded without a password.</p></div></div><div className="source-grid"><SourceCard title="TRAINS_PER_DAY.xlsx" file={tf} set={setTf} status={status?.train} onDownload={()=>download('trains')}/><SourceCard title="Kavach_Loco_Details.xlsx" file={kf} set={setKf} status={status?.kavach} onDownload={()=>download('kavach')}/></div><section className="security"><div className="security-head"><div className="security-icon"><LockKeyhole/></div><div><span className="eyebrow">UPLOAD SECURITY</span><h4>{initialized?'Password required for source updates':'Password required for initial upload'}</h4></div></div><div className="security-form"><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Enter upload password"/><button className="primary-btn" onClick={submit} disabled={busy}><LockKeyhole/>{busy?'Processing...':initialized?'Authenticate & Update':'Authenticate & Upload'}</button></div><div className="secure"><CheckCircle2/> Upload is password-protected. Source downloads do not require a password.</div></section></section>}
function SourceCard({title,file,set,status,onDownload}){return <section className="source-card"><div className="source-card-top"><FileSpreadsheet/><span>{status?.count??0} rows</span></div><h4>{title}</h4><p>{status?.fileName?`Stored: ${status.fileName}`:'Not uploaded yet'}</p><label className="picker"><UploadCloud/><span>{file?file.name:'Choose .xlsx / .xls / .xlsm file'}</span><input type="file" accept=".xlsx,.xls,.xlsm" onChange={e=>set(e.target.files?.[0])}/></label><button className="download-source" onClick={onDownload} disabled={!status}><FileDown/>Download stored file</button></section>}

createRoot(document.getElementById('root')).render(<App/>);
