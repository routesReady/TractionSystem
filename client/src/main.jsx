import React,{useEffect,useMemo,useRef,useState}from'react';
import{createRoot}from'react-dom/client';
import{motion}from'framer-motion';
import{CalendarDays,CheckCircle2,Database,Download,FileSpreadsheet,LockKeyhole,RefreshCw,Search,UploadCloud,XCircle,LoaderCircle,Menu,ShieldCheck,TrainFront,FileDown,Edit3,Save,BarChart3,Filter,ExternalLink,LogOut,UserRound,KeyRound,ShieldAlert,Check,X,Printer,Info,Mail,Phone,Workflow,ArrowRight,FileCheck2,UsersRound,ClipboardList}from'lucide-react';
import axios from'axios';
import'./styles.css';

const configuredApi=String(import.meta.env.VITE_API_URL||'http://localhost:5000/api').replace(/\/+$/,'');
const api=axios.create({baseURL:configuredApi.endsWith('/api')?configuredApi:`${configuredApi}/api`,withCredentials:true});

const cols=[
 ['sn','S.N'],['date','DATE'],['day','DAY'],['kavachSection','KAVACH_WKG_SECTION'],['dir','DIR'],['trType','TR_TYPE'],['time','T_O_Time'],['attachedDiv','LOCO_ATTACHED_DIV.'],['fromTo','From_To'],['locoLink','LOCO_link'],['locoLinkDiv','LOCO_LINK_DIV'],['dayWork','Day_of_wkg_in_territory'],['locoNo','Loco_No'],['shed','SHED'],['locoType','Loco_Type'],['rly','RLY'],['kavachMake','Loco_Kavach_Make'],['brakeSystem','Brake_system'],['trainNo','TRAIN_No'],['worked','Train_worked_with_Kavach_YES_NO'],['noReason','If_No_the_Reason'],['remarks','Remarks'],['shedRemark','Shed_Remark'],['oem','OEM']
];
const dashboardCols=cols.filter(([k])=>k!=='shedRemark'&&k!=='oem');
const headerToken=v=>String(v??'').toUpperCase().replace(/[^A-Z0-9]/g,'');
function dynamicSourceCols(sourceColumns, baseCols=cols){
 const known=new Set(baseCols.flatMap(([k,label])=>[headerToken(k),headerToken(label)]));
 const seen=new Set();
 return (sourceColumns||[]).filter(c=>c&&c.label).filter(c=>{
  const token=headerToken(c.label); if(!token||/^UNNAMEDCOLUMN\d*$/.test(token)||token.startsWith('IFNOTHEREASON')||known.has(token)||seen.has(token))return false; seen.add(token); return true;
 }).map(c=>[`source:${c.key}`,c.label]);
}
function dynamicCombinedSourceCols(trainColumns, kavachColumns, baseCols=cols){
 const train=dynamicSourceCols(trainColumns,baseCols);
 const known=new Set([...baseCols.flatMap(([k,label])=>[headerToken(k),headerToken(label)]),...train.map(([,label])=>headerToken(label))]);
 const kavach=(kavachColumns||[]).filter(c=>c&&c.label&&!/^UNNAMEDCOLUMN\d*$/.test(headerToken(c.label))&&!known.has(headerToken(c.label))).map(c=>[`kavach:${c.key}`,`Kavach — ${c.label}`]);
 return [...train,...kavach];
}
function placeLinkShedNextToLocoLink(baseColumns, extraColumns){
 const all=[...baseColumns,...extraColumns];
 const linkIndex=all.findIndex(([key,label])=>headerToken(key)==='LINKSHED'||headerToken(label)==='LINKSHED');
 if(linkIndex<0)return all;
 const [linkShed]=all.splice(linkIndex,1);
 const locoIndex=all.findIndex(([key,label])=>headerToken(key)==='LOCO_LINK'||headerToken(label)==='LOCO_LINK');
 all.splice(locoIndex>=0?locoIndex+1:all.length,0,linkShed);
 return all;
}
// Remark pages intentionally use a lean, responsive subset of Dashboard data.
// These fields were removed to avoid unnecessary horizontal scrolling while
// preserving the operational fields needed by Shed/OEM users.
const REMARK_HIDDEN_KEYS=new Set(['day','attachedDiv','locoLink','locoLinkDiv','trType','locoType','rly']);
const roleCols=(isShed,sourceColumns=[])=>{
  // Preserve the original V37 column layout: the role-specific editable
  // remark stays immediately after Date run Kota div, followed by the same
  // operational columns in their original order. The only addition is the
  // other role's saved remark immediately after TLC Remark.
  let filtered=cols
    .filter(([k])=>!REMARK_HIDDEN_KEYS.has(k))
    .map(([k,label])=>{
      if(k==='date') return [k,'Date run Kota div'];
      if(k==='remarks') return [k,'TLC Remark'];
      if(k==='shedRemark') return [k,'Shed Remark'];
      if(k==='oem') return [k,'OEM Remark'];
      return [k,label];
    });
  const targetField=isShed?'shedRemark':'oem';
  const otherRemarkField=isShed?'oem':'shedRemark';
  // Keep the editable remark in the same position as the previous version.
  const dateIndex=filtered.findIndex(([k])=>k==='date');
  const target=filtered.find(([k])=>k===targetField);
  filtered=filtered.filter(([k])=>k!==targetField);
  filtered.splice(Math.max(0,dateIndex+1),0,target);

  // Preserve all other columns' relative order. Only move the requested
  // operational identifiers beside the editable remark and add the other
  // role's remark after TLC Remark, before KAVACH_WKG_SECTION.
  const wanted=[targetField,'locoNo','shed','kavachMake','remarks',otherRemarkField,'kavachSection'];
  const moving=new Map(filtered.filter(([k])=>wanted.includes(k)).map(([k,label])=>[k,[k,label]]));
  // Remove the block first, then locate the date column in the remaining
  // array. This keeps the entire remark/operational block immediately after
  // `Date run Kota div` without an index shift caused by removed columns.
  filtered=filtered.filter(([k])=>!wanted.includes(k));
  const ordered=wanted.map(k=>moving.get(k)).filter(Boolean);
  const dateIndexAfterRemoval=filtered.findIndex(([k])=>k==='date');
  filtered.splice(dateIndexAfterRemoval >= 0 ? dateIndexAfterRemoval + 1 : 0,0,...ordered);
  const extras=dynamicSourceCols(sourceColumns, cols);
  return [...filtered, ...extras];
};
function remarkDateKey(value){
  const s=String(value??'').trim();
  let m=s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if(m)return `${m[3]}-${m[2]}-${m[1]}`;
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m)return `${m[1]}-${m[2]}-${m[3]}`;
  const d=new Date(s);
  return Number.isNaN(d.getTime())?'':`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function sortRemarkRows(rows){
  return [...(rows||[])].sort((a,b)=>{
    const dateDiff=remarkDateKey(b.searchDate).localeCompare(remarkDateKey(a.searchDate));
    if(dateDiff)return dateDiff;
    const rowDiff=Number(b.sourceRow??b.row?.sourceRow??0)-Number(a.sourceRow??a.row?.sourceRow??0);
    return rowDiff || String(b._id||'').localeCompare(String(a._id||''));
  });
}
const editableKeys=new Set(['locoNo','worked','noReason','remarks']);
const derivedKeys=new Set(['shed','locoType','rly','kavachMake','brakeSystem']);

function localTodayDMY(){const now=new Date();return `${String(now.getDate()).padStart(2,'0')}-${String(now.getMonth()+1).padStart(2,'0')}-${now.getFullYear()}`}
function dmyToISO(value){const m=String(value||'').match(/^(\d{2})-(\d{2})-(\d{4})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:''}
function isoToDMY(value){const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);return m?`${m[3]}-${m[2]}-${m[1]}`:''}
function DatePicker({value,onChange,label}){return <div className="field"><label><CalendarDays/> {label}</label><div className="date-input-wrap"><input type="date" value={dmyToISO(value)} onChange={e=>onChange(isoToDMY(e.target.value))}/></div></div>}
function getInitialPage(){
 const path=window.location.pathname;
 if(path==='/login')return'login'; if(path==='/dashboard')return'dash'; if(path==='/stored-data')return'saved';
 if(path==='/summary')return'summary'; if(path==='/about-us')return'about'; if(path==='/data-sources')return'src'; if(path==='/shed-remark')return'shed';
 if(path==='/oem-remark')return'oem'; if(path==='/summary-detail')return'detail'; if(path==='/remarks-pending')return'remarksPending'; if(path==='/password-reset-requests')return'resets'; if(path==='/active-sessions')return'activeSessions';
 const p=new URLSearchParams(window.location.search);
 if(p.get('page')==='saved')return'saved'; if(p.get('page')==='summary')return'summary';
 return'about';
}
const PAGE_PATH={login:'/login',dash:'/dashboard',saved:'/stored-data',summary:'/summary',detail:'/summary-detail',about:'/about-us',remarksPending:'/remarks-pending',src:'/data-sources',shed:'/shed-remark',oem:'/oem-remark',resets:'/password-reset-requests',activeSessions:'/active-sessions'};
function navigatePage(page,params={},replace=false){
 const clean=Object.fromEntries(Object.entries(params).filter(([,v])=>v!==undefined&&v!==null&&v!==''));
 const q=new URLSearchParams(clean),path=PAGE_PATH[page]||'/dashboard',url=q.toString()?`${path}?${q}`:path;
 (replace?window.history.replaceState:window.history.pushState).call(window.history,{},'',url);
 // pushState/replaceState do not fire popstate. Notify the app so internal
 // navigation changes the rendered page without reloading the Vercel SPA.
 window.dispatchEvent(new PopStateEvent('popstate'));
}

function PublicShell({tab,setTab,msg,setMsg,anchorTimestamp}){
 const [mobileNav,setMobileNav]=useState(false);
 const nav=[['summary',<BarChart3/>,'Summary'],['detail',<FileSpreadsheet/>,'Summary & Detail'],['saved',<Download/>,'Saved Data / Export'],['remarksPending',<ClipboardList/>,'Remarks Pending'],['about',<Info/>,'About Us']];
 function change(id){setTab(id);navigatePage(id);setMsg('');setMobileNav(false)}
 function goLogin(){setMobileNav(false);setTab('login');navigatePage('login',{},true)}
 return <div className="app-shell">
  <header className="gov-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div><div className="header-account public-account"><button className="login-header-btn desktop-login" onClick={goLogin}><LogOut/> Login</button><button className="mobile-menu public-mobile-menu" aria-label={mobileNav?'Close menu':'Open menu'} aria-expanded={mobileNav} onClick={()=>setMobileNav(v=>!v)}><Menu/></button></div></div>
   <div className="nav-strip"><div className="nav-inner">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>change(id)}>{icon}{label}</button>)}<div className="nav-spacer"/><div className="official-chip"><ShieldCheck/> PUBLIC ACCESS</div></div></div>
   {mobileNav&&<div className="mobile-nav public-mobile-nav">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'active':''} onClick={()=>change(id)}>{icon}{label}</button>)}<button onClick={goLogin}><LogOut/>Login</button></div>}
  </header>
  <main className="content">{msg&&<div className="app-message">{msg}</div>}{tab==='summary'?<Summary setMsg={setMsg}/>:tab==='detail'?<SummaryDetail setMsg={setMsg}/>:tab==='remarksPending'?<RemarksPending setMsg={setMsg} user={null}/>:tab==='about'?<AboutUs anchorTimestamp={anchorTimestamp}/>:<SavedData setMsg={setMsg}/>}</main>
  <footer className="footer"><div>West Central Railway · Kota Division · Traction Operation</div><span className="footer-credit"><span>@mak</span><span>Designed &amp; Developed by Arsh, +919509206604</span></span></footer>
 </div>
}

function App(){
 const [user,setUser]=useState(null),[sessionInfo,setSessionInfo]=useState(null),[checking,setChecking]=useState(true),[tab,setTab]=useState(getInitialPage),[date,setDate]=useState(localTodayDMY),[res,setRes]=useState(),[msg,setMsg]=useState(''),[status,setStatus]=useState(),[loading,setLoading]=useState(false),[mobileNav,setMobileNav]=useState(false),[options,setOptions]=useState({workedOptions:[],reasonOptions:[]}),[refreshKey,setRefreshKey]=useState(0),[changePassword,setChangePassword]=useState(false);
 const statusLoad=async()=>{try{const{data}=await api.get('/upload/status');setStatus(data)}catch(e){if(e.response?.status!==401)setMsg('Backend is not reachable. Check the deployed API or start the server.')}};
 const loadOptions=async()=>{try{const{data}=await api.get('/data/options');setOptions(data)}catch{setOptions({workedOptions:[],reasonOptions:[]})}};
 const checkSession=async()=>{
  // A deliberate logout is a local, persistent barrier against restoring an old
  // HttpOnly cookie from browser history/BFCache. Only a successful new login clears it.
  try{if(localStorage.getItem('traction-operation-logout')){setUser(null);userRef.current=null;setSessionInfo(null);setChecking(false);return null}}catch{}
  try{const{data}=await api.get('/auth/me');setUser(data.user);userRef.current=data.user;setSessionInfo(data.session||null);return data.user}catch{setUser(null);userRef.current=null;setSessionInfo(null);return null}finally{setChecking(false)}
 };
 const userRef=useRef(user);
 useEffect(()=>{userRef.current=user},[user]);
 useEffect(()=>{
  checkSession();
  const protectedPaths=new Set(['/dashboard','/data-sources','/shed-remark','/oem-remark','/password-reset-requests','/active-sessions']);
  const onPop=()=>{
   const requested=getInitialPage();
   let explicitlyLoggedOut=false;try{explicitlyLoggedOut=Boolean(localStorage.getItem('traction-operation-logout'))}catch{}
   if((!userRef.current||explicitlyLoggedOut) && protectedPaths.has(window.location.pathname)){
    setRes();setSessionInfo(null);setTab('about');navigatePage('about',{},true);return;
   }
   setTab(requested);
  };
  const onStorage=e=>{if(e.key==='traction-operation-logout'){
   setUser(null);userRef.current=null;setSessionInfo(null);setRes();setTab('about');navigatePage('about',{},true);
  }};
  const onPageShow=async e=>{
   // Browsers may restore a protected page from BFCache even after its URL was
   // changed. A local logout barrier wins immediately; otherwise revalidate.
   let explicitlyLoggedOut=false;try{explicitlyLoggedOut=Boolean(localStorage.getItem('traction-operation-logout'))}catch{}
   if(explicitlyLoggedOut && protectedPaths.has(window.location.pathname)){
    setUser(null);userRef.current=null;setRes();setSessionInfo(null);setTab('about');navigatePage('about',{},true);return;
   }
   if(e.persisted){const currentUser=await checkSession();if(!currentUser&&protectedPaths.has(window.location.pathname)){setRes();setSessionInfo(null);setTab('about');navigatePage('about',{},true)}}
  };
  window.addEventListener('popstate',onPop);window.addEventListener('storage',onStorage);window.addEventListener('pageshow',onPageShow);
  return()=>{window.removeEventListener('popstate',onPop);window.removeEventListener('storage',onStorage);window.removeEventListener('pageshow',onPageShow)}
 },[]);
 useEffect(()=>{if(checking||user)return;const protectedTabs=['dash','src','shed','oem','resets','activeSessions'];if(protectedTabs.includes(tab)){setTab('about');navigatePage('about',{},true)}},[checking,user,tab]);
 useEffect(()=>{if(['admin','admin_limited'].includes(user?.role)){statusLoad();loadOptions()}},[user]); useEffect(()=>{if(!user)return;const allowed=['admin','admin_limited'].includes(user.role)?(user.role==='admin'?['dash','saved','summary','detail','src','resets','remarksPending','activeSessions']:['dash','saved','summary','detail','src','remarksPending']):user.role==='shed'?['shed','remarksPending']:['oem','remarksPending'];if(!allowed.includes(tab)){const home=['admin','admin_limited'].includes(user.role)?'dash':user.role==='shed'?'shed':'oem';setTab(home);navigatePage(home,{},true)}},[user,tab]);
 useEffect(()=>{if(!user||['admin','admin_limited'].includes(user.role))return;let mounted=true,lastActivity=Date.parse(sessionInfo?.lastActivityAt||sessionInfo?.loginAt||new Date().toISOString());let lastSent=0;const resetActivity=async()=>{const now=Date.now();if(now-lastSent<10000)return;lastSent=now;try{const{data}=await api.post('/auth/activity');if(mounted&&data.lastActivityAt){lastActivity=Date.parse(data.lastActivityAt);setSessionInfo(v=>({...v,...data}));}}catch(e){if(e.response?.status===401&&mounted)logout(true)}};const activity=()=>{if(document.visibilityState==='visible')resetActivity()};const tick=()=>{if(!mounted)return;if(Date.now()-lastActivity>=300000){logout(true);return;}setSessionInfo(v=>({...v,_remaining:Math.max(0,300000-(Date.now()-lastActivity))}));};const poll=async()=>{try{const{data}=await api.get('/auth/session-status');if(data.lastActivityAt){lastActivity=Date.parse(data.lastActivityAt);if(mounted)setSessionInfo(v=>({...v,...data,_remaining:Math.max(0,300000-(Date.now()-lastActivity))}));}}catch(e){if(e.response?.status===401&&mounted)logout(true)}};const events=['pointerdown','keydown','touchstart','submit','change'];events.forEach(ev=>window.addEventListener(ev,activity,{passive:true}));const timer=setInterval(tick,1000),poller=setInterval(poll,20000);return()=>{mounted=false;events.forEach(ev=>window.removeEventListener(ev,activity));clearInterval(timer);clearInterval(poller)}},[user?.id,user?.role]);
 async function logout(expired=false){
  // Set the local barrier before awaiting the network so refresh/back cannot race
  // with the logout request. The backend invalidates the server session as well.
  try{localStorage.setItem('traction-operation-logout',String(Date.now()));['token','role','user','session','authToken'].forEach(k=>{localStorage.removeItem(k);sessionStorage.removeItem(k)})}catch{}
  setUser(null);userRef.current=null;setSessionInfo(null);setRes();setTab('about');setMsg(expired?'Your session expired after five minutes of inactivity. Please sign in again.':'');navigatePage('about',{},true);setMobileNav(false);window.scrollTo(0,0);
  try{await api.post('/auth/logout')}catch{}
 }
 async function search(){setLoading(true);setMsg('');try{const{data}=await api.get('/data/search',{params:{date}});setRes(data);await loadOptions()}catch(e){setRes();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 function patchRow(sourceRow,patch){setRes(x=>x?{...x,rows:x.rows.map(r=>r.sourceRow===sourceRow?{...r,...patch}:r)}:x)}
 async function lookup(row){const locoNo=String(row.locoNo||'').trim().slice(0,6);patchRow(row.sourceRow,{locoNo,kavachFound:false,kavachLoading:Boolean(locoNo),kavachError:''});if(!locoNo){patchRow(row.sourceRow,{kavachLoading:false,kavachError:'Enter Loco_No'});return}try{const{data}=await api.get('/kavach/lookup',{params:{locoNo}});patchRow(row.sourceRow,{...data,kavachFound:true,kavachLoading:false,kavachError:''})}catch(e){patchRow(row.sourceRow,{kavachFound:false,kavachLoading:false,kavachError:e.response?.data?.message||'Loco_No not found in Kavach_Loco_Details.xlsx.'})}}
 async function saveRow(row){if(!String(row.locoNo||'').trim())return setMsg('Enter Loco_No before saving.');if(!row.kavachFound)return setMsg('Enter a valid Loco_No and complete the Kavach lookup.');if(!String(row.worked||'').trim())return setMsg('Please select Train Worked with Kavach before updating the record.');if(String(row.worked||'').trim().toUpperCase()==='NO'&&!String(row.noReason||'').trim())return setMsg('Please select If No – Reason before updating the record.');if(String(row.remarks||'').length>500)return setMsg('Remarks cannot exceed 500 characters.');try{const{data}=await api.post('/data/rows',{searchDate:res.date,weekday:res.weekday,row});patchRow(row.sourceRow,{...data.row,added:true,editing:false,kavachFound:true});setMsg(`Loco ${data.row.locoNo} record ${row.added?'updated':'saved'} successfully.`)}catch(e){setMsg(e.response?.data?.message||e.message)}}
 function beginModify(row){patchRow(row.sourceRow,{editing:true});setMsg(`Modify mode enabled for Loco ${row.locoNo||'record'}.`)}
 function changeTab(id){if(id==='dash')setDate(localTodayDMY());setTab(id);navigatePage(id);setMobileNav(false);setMsg('')}
 const title=tab==='activeSessions'?'Active Sessions':tab==='dash'?'Daily Traction Roster':tab==='saved'?'Saved Data & Excel Export':tab==='summary'?'Summary':tab==='detail'?'Summary & Detail':tab==='remarksPending'?'Remarks Pending':tab==='src'?'Source Data Management':tab==='shed'?'Shed Remark':tab==='oem'?'OEM Remark':'Password Reset Requests';
 if(checking)return <div className="auth-loading"><LoaderCircle className="spin"/><span>Checking secure session...</span></div>;
 if(tab==='about') return <PublicShell tab={tab} setTab={setTab} msg={msg} setMsg={setMsg} anchorTimestamp={sessionInfo?.loginAt}/>;
 if(!user){
   if(tab==='summary'||tab==='saved'||tab==='detail'||tab==='about'||tab==='remarksPending') return <PublicShell tab={tab} setTab={setTab} msg={msg} setMsg={setMsg} anchorTimestamp={sessionInfo?.loginAt}/>;
   return <Login expiredMessage={msg} onHome={()=>{window.location.assign('/')}} onLogin={(u,session)=>{try{localStorage.removeItem('traction-operation-logout')}catch{}setUser(u);userRef.current=u;setSessionInfo(session||null);const p=['admin','admin_limited'].includes(u.role)?'dash':u.role==='shed'?'shed':'oem';setTab(p);navigatePage(p,{},true)}}/>;
 }
 const isAdmin=['admin','admin_limited'].includes(user.role);
 const nav=isAdmin?(user.role==='admin'?[['dash',<TrainFront/>,'Dashboard'],['saved',<Download/>,'Saved Data / Export'],['summary',<BarChart3/>,'Summary'],['detail',<FileSpreadsheet/>,'Summary & Detail'],['remarksPending',<ClipboardList/>,'Remarks Pending'],['src',<Database/>,'Data Sources'],['resets',<KeyRound/>,'Password Reset Requests'],['activeSessions',<UsersRound/>,'Active Sessions']]:[['dash',<TrainFront/>,'Dashboard'],['saved',<Download/>,'Saved Data / Export'],['summary',<BarChart3/>,'Summary'],['detail',<FileSpreadsheet/>,'Summary & Detail'],['remarksPending',<ClipboardList/>,'Remarks Pending'],['src',<Database/>,'Data Sources']]):user.role==='shed'?[['shed',<Database/>,'Shed Remark'],['remarksPending',<ClipboardList/>,'Remarks Pending']]:[['oem',<Database/>,'OEM Remark'],['remarksPending',<ClipboardList/>,'Remarks Pending']];
 return <div className="app-shell">
  <header className="gov-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div><div className="header-account"><button className="user-chip user-action" onClick={()=>setChangePassword(true)} title="Change password"><UserRound/>{user.username}<KeyRound/></button><button className="logout-btn" onClick={logout}><LogOut/>Logout</button><button className="mobile-menu" onClick={()=>setMobileNav(v=>!v)} aria-label="Open navigation"><Menu/></button></div></div>
   <div className="nav-strip"><div className="nav-inner">{nav.map(([id,icon,label])=><button key={id} className={tab===id?'nav-link active':'nav-link'} onClick={()=>changeTab(id)}>{icon}{label}</button>)}<div className="nav-spacer"/>{!['admin','admin_limited'].includes(user.role)&&<div className="session-countdown" title="Five-minute inactivity timeout"><span>Session expires in</span><strong>{(()=>{const ms=sessionInfo?._remaining??Math.max(0,300000-(Date.now()-Date.parse(sessionInfo?.lastActivityAt||sessionInfo?.loginAt||new Date().toISOString())));return `${String(Math.floor(ms/60000)).padStart(2,'0')}:${String(Math.floor(ms%60000/1000)).padStart(2,'0')}`})()}</strong></div>}<div className="official-chip"><ShieldCheck/> {user.role.toUpperCase()} · Official Workflow</div></div></div>
   {mobileNav&&<div className="mobile-nav">{nav.map(([id,,label])=><button key={id} className={tab===id?'active':''} onClick={()=>changeTab(id)}>{label}</button>)}<button className="mobile-logout" onClick={logout}><LogOut/> Logout</button></div>}
  </header>
  <main className="content"><div className="page-heading"><div><span>TRACTION OPERATION</span><h2>{title}</h2></div><button className="refresh-btn" onClick={()=>setRefreshKey(v=>v+1)} title="Refresh"><RefreshCw/></button></div>
   {tab==='dash'?<Dashboard key={refreshKey} refreshSignal={refreshKey} date={date} setDate={setDate} res={res} search={search} loading={loading} msg={msg} saveRow={saveRow} patchRow={patchRow} lookup={lookup} beginModify={beginModify} options={options}/>:tab==='saved'?<SavedData key={refreshKey} setMsg={setMsg}/>:tab==='summary'?<Summary key={refreshKey} setMsg={setMsg}/>:tab==='detail'?<SummaryDetail key={refreshKey} setMsg={setMsg}/>:tab==='remarksPending'?<RemarksPending key={refreshKey} setMsg={setMsg} user={user}/>:tab==='src'?<Upload key={refreshKey} status={status} refresh={()=>{statusLoad();loadOptions()}} setMsg={setMsg}/>:tab==='shed'?<RemarkPage key={refreshKey} role="shed" setMsg={setMsg}/>:tab==='oem'?<RemarkPage key={refreshKey} role="oem" setMsg={setMsg}/>:tab==='activeSessions'?<ActiveSessions setMsg={setMsg}/>:<ResetRequests key={refreshKey} setMsg={setMsg}/>} 
   {msg&&tab!=='dash'&&<div className="alert"><XCircle/>{msg}</div>}
  </main>{changePassword&&<ChangePasswordModal onClose={()=>setChangePassword(false)} onDone={()=>{setChangePassword(false);logout()}}/>}<footer className="footer"><div>West Central Railway · Kota Division · Traction Operation</div><span className="footer-credit"><span>@mak</span><span>Designed &amp; Developed by Arsh, +919509206604</span></span></footer>
 </div>
}
function Login({onLogin,onHome,expiredMessage=''}){
 const[username,setUsername]=useState(''),[password,setPassword]=useState(''),[busy,setBusy]=useState(false),[error,setError]=useState(expiredMessage||''),[forgot,setForgot]=useState(false);
 async function submit(e){e.preventDefault();setBusy(true);setError('');try{const{data}=await api.post('/auth/login',{username,password});onLogin(data.user,data.session)}catch(e){setError(e.response?.data?.message||'Invalid username or password. Please try again.')}finally{setBusy(false)}}
 return <div className="login-shell"><header className="gov-header login-header"><div className="gov-inner"><div className="brand-logo"><img src="/west-central-railway.png" alt="West Central Railway"/></div><div className="gov-title"><div className="gov-kicker">WEST CENTRAL RAILWAY</div><h1>Traction Operation - Kavach Monitoring System</h1><p>West Central Railway, Kota Division</p></div></div></header><main className="login-main"><form className="login-card" onSubmit={submit}><div className="login-icon"><ShieldCheck/></div><span className="eyebrow">SECURE LOGIN</span><h2>Sign in</h2><p className="login-subtitle">Access the Kavach Monitoring System using your authorized account.</p>{error&&<div className="login-error"><ShieldAlert/>{error}</div>}<div className="field"><label><UserRound/> Username</label><input autoComplete="username" value={username} onChange={e=>setUsername(e.target.value)} placeholder="Enter username"/></div><div className="field"><label><KeyRound/> Password</label><input type="password" autoComplete="current-password" value={password} onChange={e=>setPassword(e.target.value)} placeholder="Enter password"/></div><button type="button" className="forgot-link" onClick={()=>setForgot(true)}>Forgot Password?</button><button className="primary-btn login-btn" disabled={busy||!username||!password}>{busy?<><LoaderCircle className="spin"/>Signing in...</>:<><ShieldCheck/>Login</>}</button><button type="button" className="secondary-btn back-home-btn" onClick={onHome}><span aria-hidden="true">←</span> Back to Home</button></form></main>{forgot&&<ResetRequestModal onClose={()=>setForgot(false)}/>}<footer className="footer login-footer"><div>West Central Railway · Kota Division · Traction Operation</div><span className="footer-credit"><span>@mak</span><span>Designed &amp; Developed by Arsh, +919509206604</span></span></footer></div>
}
function ResetRequestModal({onClose}){
 const [role,setRole]=useState('shed'),[username,setUsername]=useState(''),[name,setName]=useState(''),[designation,setDesignation]=useState(''),[userType,setUserType]=useState('SHED User'),[hqStation,setHqStation]=useState(''),[division,setDivision]=useState(''),[mobile,setMobile]=useState(''),[emailAddress,setEmailAddress]=useState(''),[busy,setBusy]=useState(false),[msg,setMsg]=useState('');
 function changeRole(v){setRole(v);setUserType(v==='shed'?'SHED User':'OEM User')}
 async function submit(){if(!name.trim()||!designation.trim()||!username.trim()||!hqStation.trim()||!division.trim()||!/^\d{10}$/.test(mobile)||! /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailAddress.trim())){setMsg('Complete every field. Enter exactly 10 digits for mobile and a valid email address.');return}setBusy(true);setMsg('');try{const{data}=await api.post('/auth/reset-request',{username,role,name,designation,userType,hqStation,division,mobileNumber:mobile,emailAddress});setMsg(data.message)}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 return <div className="modal-backdrop"><div className="modal-card reset-request-card"><button className="modal-close" onClick={onClose}><X/></button><KeyRound className="modal-icon"/><h3>Request Admin Reset</h3><p>Complete the details below. The Admin will review your request.</p><div className="field"><label>Username</label><input value={username} onChange={e=>setUsername(e.target.value)} placeholder="Your account username" required/></div><div className="field"><label>Name</label><input value={name} onChange={e=>setName(e.target.value)} required/></div><div className="field"><label>Designation</label><input value={designation} onChange={e=>setDesignation(e.target.value)} required/></div><div className="field"><label>User Type</label><select value={userType} onChange={e=>{setUserType(e.target.value);setRole(e.target.value==='SHED User'?'shed':'oem')}}><option>SHED User</option><option>OEM User</option></select></div><div className="field"><label>HQ/Station</label><input value={hqStation} onChange={e=>setHqStation(e.target.value)} required/></div><div className="field"><label>Division</label><input value={division} onChange={e=>setDivision(e.target.value)} required/></div><div className="field"><label>Mobile Number</label><input value={mobile} onChange={e=>setMobile(e.target.value.replace(/\D/g,'').slice(0,10))} inputMode="numeric" maxLength={10} placeholder="10-digit mobile number" required/></div><div className="field"><label>Email Address</label><input type="email" value={emailAddress} onChange={e=>setEmailAddress(e.target.value)} placeholder="name@example.com" required/></div>{msg&&<div className="modal-message">{msg}</div>}<button className="primary-btn" onClick={submit} disabled={busy||!username.trim()||!name.trim()||!designation.trim()||!hqStation.trim()||!division.trim()||!/^\d{10}$/.test(mobile)||!emailAddress.trim()}>{busy?'Submitting...':'Submit Request'}</button></div></div>
}

function RemarkPage({role,setMsg}){
 const[data,setData]=useState(),[loading,setLoading]=useState(true),[saving,setSaving]=useState('');
 const isShed=role==='shed';
 const dynamicColumns=useMemo(()=>dynamicCombinedSourceCols(data?.rows?.[0]?.row?.sourceColumns||[],data?.rows?.[0]?.row?.kavachSourceColumns||[],cols),[data]);
 const visibleRoleCols=useMemo(()=>[...roleCols(isShed),...dynamicColumns],[isShed,dynamicColumns]);
 const shellRef=useRef(null),tableRef=useRef(null);
 async function load(){setLoading(true);try{const{data:x}=await api.get('/remark-records');setData({...x,rows:sortRemarkRows(x.rows)})}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 useEffect(()=>{load()},[]);

 async function update(id,remark){setSaving(id);try{const{data:x}=await api.patch(`/remark-records/${id}`,{remark});setData(d=>({...d,rows:sortRemarkRows(d.rows.map(r=>r._id===id?x.row:r))}));setMsg(x.message);return true}catch(e){setMsg(e.response?.data?.message||e.message);return false}finally{setSaving('')}}
 async function closeRecord(id){
  if(!window.confirm('Are you sure you want to close this record? Once closed, it will no longer appear on this page.')) return;
  setSaving(id);
  try{
   const{data:x}=await api.patch(`/remark-records/${id}/close`);
   setData(d=>({...d,rows:d.rows.filter(r=>r._id!==id)}));
   setMsg(x.message||'Record closed successfully.');
  }catch(e){setMsg(e.response?.data?.message||e.message)}finally{setSaving('')}
 }
 return <section className="panel role-panel"><div className="panel-intro compact"><div className="intro-icon"><Database/></div><div><span className="eyebrow">{isShed?'SHED WORKFLOW':'OEM WORKFLOW'}</span><h3>{isShed?'Shed Remark':'OEM Remark'}</h3><p>Logged in as <strong>{data?.user?.username||'—'}</strong>. Filters are automatically applied by the server. Both Shed and OEM remarks are shown for context. Only the {isShed?'Shed Remark':'OEM Remark'} field is editable on this page.</p></div></div>{loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading relevant records...</div>:data?.rows?.length?<><div className="role-table-shell" ref={shellRef}><table ref={tableRef} className={`data-table role-table role-full-table ${isShed?'role-shed-table':'role-oem-table'}`}><thead><tr>{visibleRoleCols.map(([key,label])=><th key={key} className={`role-col-${key} ${key==='worked'?'worked-head':key==='noReason'?'reason-head':''}`}>{key==='worked'?<>Train Worked<br/>with Kavach</>:key==='noReason'?'If No – Reason':label}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>{data.rows.map((x,i)=><RemarkRow key={x._id} record={x} index={i} isShed={isShed} saving={saving} onUpdate={update} onClose={closeRecord} dynamicColumns={dynamicColumns}/>)}</tbody></table></div></>:<div className="empty-panel">No records currently match the automatic {isShed?'Shed':'OEM'} filters.</div>}</section>
}
function RemarkRow({record,index,isShed,saving,onUpdate,onClose,dynamicColumns=[]}){
 const r=record.row||{};
 const initialRemark=isShed?(r.shedRemark||record.shedRemark||''):(r.oem||record.oem||'');
 const[remark,setRemark]=useState(initialRemark),[editing,setEditing]=useState(!initialRemark);
 const[hasSavedRemark,setHasSavedRemark]=useState(Boolean(initialRemark));
 const closed=isShed?Boolean(record.shedRemarkClosed):Boolean(record.oemRemarkClosed);
 useEffect(()=>{const next=isShed?(r.shedRemark||record.shedRemark||''):(r.oem||record.oem||'');setRemark(next);setHasSavedRemark(Boolean(next));setEditing(!next)},[record._id,record.updatedAt,isShed]);
 const valueFor=(k)=>k==='sn'?index+1:k==='date'?record.searchDate:k==='day'?record.weekday:(r[k]??'');
 return <tr>{[...roleCols(isShed),...dynamicColumns].map(([k,label])=>{
   if(k.startsWith('source:')||k.startsWith('kavach:')) { const isKavach=k.startsWith('kavach:'); const fieldKey=k.slice(isKavach?7:7); const raw=isKavach?r.kavachSourceData?.[fieldKey]:r.sourceData?.[fieldKey]; return <td key={k} className={`role-col-${k.replace(/[^a-zA-Z0-9_-]/g,'-')}`}><span className="wrap-text">{raw||'—'}</span></td>; }
   if(k==='shedRemark' || k==='oem'){
     const isTarget=(isShed&&k==='shedRemark')||(!isShed&&k==='oem');
     const v=isTarget?remark:(r[k]??(k==='shedRemark'?record.shedRemark:k==='oem'?record.oem:''));
     return <td key={k} className={`new-field-cell role-col-${k}`}>{isTarget&&editing?<textarea className="role-remark-input" maxLength={isShed?300:200} value={v} onChange={e=>setRemark(e.target.value)} placeholder={isShed?'Enter Shed Remark':'Enter OEM Remark'}/>:<span className="wrap-text">{v||'—'}</span>}</td>;
   }
   return <td key={k} className={`${k==='locoNo'?'loco-cell ':''}role-col-${k}`}><span className="wrap-text">{valueFor(k)||'—'}</span></td>;
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
 return <section className="panel role-panel"><div className="panel-intro compact"><div className="intro-icon"><KeyRound/></div><div><span className="eyebrow">ADMIN SECURITY</span><h3>Password Reset Requests</h3><p>Review and process Shed/OEM reset requests.</p></div></div>{loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading requests...</div>:data.requests.length?<div className="role-table-shell"><table className="data-table role-table"><thead><tr><th>REQUEST ID</th><th>USERNAME</th><th>NAME</th><th>DESIGNATION</th><th>USER TYPE</th><th>HQ/STATION</th><th>DIVISION</th><th>MOBILE</th><th>EMAIL</th><th>REQUEST DATE & TIME</th><th>STATUS</th><th>EMAIL STATUS</th><th>ACTION</th></tr></thead><tbody>{data.requests.map(r=><tr key={r.requestId}><td>{r.requestId}</td><td>{r.username}</td><td>{r.name||'—'}</td><td>{r.designation||'—'}</td><td>{r.userType||r.role}</td><td>{r.hqStation||'—'}</td><td>{r.division||'—'}</td><td>{r.mobileNumber}</td><td>{r.emailAddress||'—'}</td><td>{new Date(r.requestedAt).toLocaleString()}</td><td>{r.status}</td><td>{r.emailStatus||'—'}{r.emailError?` · ${r.emailError}`:''}</td><td>{r.status==='Pending'?<div className="request-actions"><button className="update approve" disabled={busy===r.requestId} onClick={()=>action(r.requestId,'approve')}><Check/>Approve</button><button className="update reject" disabled={busy===r.requestId} onClick={()=>action(r.requestId,'reject')}><X/>Reject</button></div>:'—'}</td></tr>)}</tbody></table></div>:<div className="empty-panel">No password reset requests found.</div>}</section>
}
function Dashboard({refreshSignal,date,setDate,res,search,loading,msg,saveRow,patchRow,lookup,beginModify,options}){
 const workedOptions=options.workedOptions||[],reasonOptions=options.reasonOptions||[]; const visibleDashboardCols=useMemo(()=>placeLinkShedNextToLocoLink(dashboardCols,dynamicCombinedSourceCols(res?.columns,res?.kavachColumns,dashboardCols)), [res]); useEffect(()=>{if(refreshSignal>0&&date)search()},[refreshSignal]);
 return <>
  <section className="search-card"><DatePicker label="Date" value={date} onChange={setDate}/><button className="primary-btn" onClick={search} disabled={loading||!date}><Search/>{loading?'Fetching...':'Fetch Data'}</button></section>
  {msg&&<div className="alert"><XCircle/>{msg}</div>}
  {res&&<section className="results-card dashboard-results"><div className="result-head"><div><span className="eyebrow">DAILY ROSTER</span><h3>{res.weekday} <i>·</i> {res.date}</h3></div><div className="result-meta">{res.total} records</div></div>
   <div className="table-shell dashboard-table-shell"><table className="data-table"><thead><tr>{visibleDashboardCols.map(([key,label])=><th key={key} className={key==='worked'?'worked-head':key==='noReason'?'reason-head':''} title={label}>{key==='worked'?<>Train Worked<br/>with Kavach</>:key==='noReason'?'If No – Reason':label}</th>)}<th className="action-header">ACTION</th></tr></thead><tbody>
    {res.rows.length===0?<tr><td colSpan={visibleDashboardCols.length+1} className="empty-cell">No train records found for the selected date.</td></tr>:res.rows.map((r,i)=><motion.tr initial={{opacity:0,y:5}} animate={{opacity:1,y:0}} transition={{duration:.18,delay:i*.01}} key={r.sourceRow}>
     {visibleDashboardCols.map(([k,label])=><td key={k} className={`${k==='locoNo'?'loco-cell ':''}${k==='remarks'?'remarks-cell':''}${k==='shedRemark'||k==='oem'?' new-field-cell':''}`}>{k.startsWith('source:')?<span className="wrap-text">{r.sourceData?.[k.slice(7)]||'—'}</span>:k.startsWith('kavach:')?<span className="wrap-text">{r.kavachSourceData?.[k.slice(7)]||'—'}</span>:<CellEditor k={k} label={label} row={r} patchRow={patchRow} lookup={lookup} workedOptions={workedOptions} reasonOptions={reasonOptions}/>} </td>)}
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
 const visibleSavedCols=useMemo(()=>placeLinkShedNextToLocoLink(cols,dynamicCombinedSourceCols(data?.rows?.[0]?.row?.sourceColumns||[],data?.rows?.[0]?.row?.kavachSourceColumns||[],cols)), [data]);
 async function load(overrides={}){const sd=overrides.startDate??startDate,ed=overrides.endDate??endDate,dr=overrides.dir??dir,mt=overrides.metric??metric;setLoading(true);setMsg('');try{const{data:x}=await api.get('/data/saved',{params:{startDate:sd,endDate:ed,dir:dr,metric:mt}});setData(x);setMetric(mt);navigatePage('saved',{...(sd?{date:sd}:{}),...(ed&&ed!==sd?{endDate:ed}:{}),...(dr?{dir:dr}:{}),...(mt?{metric:mt}:{}),...(backSummary?{fromSummary:'1',summaryFromDate:params.get('summaryFromDate')||'',summaryToDate:params.get('summaryToDate')||'',summaryDir:params.get('summaryDir')||''}:{})},true)}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 useEffect(()=>{if(startDate||dir)load()},[]);
 async function download(){setMsg('');try{const r=await api.get('/data/export',{params:{startDate,endDate,dir},responseType:'blob'});downloadBlob(r,'Traction_Operation_Data.xlsx')}catch(e){setMsg(await blobError(e,'Excel export failed.'))}}
 return <section className="panel"><div className="panel-intro compact"><div className="intro-icon"><Download/></div><div><span className="eyebrow">DATABASE EXPORT</span><h3>Saved records &amp; Excel download</h3><p>Filter saved records by date, DIR and export the matching data set.</p></div></div>{backSummary&&<div className="back-summary-row"><button className="secondary-btn" onClick={()=>window.history.back()}>← Back to Summary</button></div>}<div className="filterbar saved-filterbar"><DatePicker label="From date" value={startDate} onChange={setStartDate}/><DatePicker label="To date" value={endDate} onChange={setEndDate}/><div className="field"><label><Filter/> DIR</label><select className="top-select" value={dir} onChange={e=>setDir(e.target.value)}><option value="">All DIR</option><option value="UP">UP</option><option value="DN">DN</option></select></div><button className="secondary-btn" onClick={()=>load()} disabled={loading}><Search/>{loading?'Filtering...':'Filter'}</button><button className="primary-btn" onClick={download}><FileSpreadsheet/>Download Excel</button></div>{data&&<div className="savedresult"><strong>{data.total}</strong> saved row(s) match the selected filters.</div>}{data?.rows?.length>0&&<div className="table-shell saved-table"><table className="data-table"><thead><tr>{visibleSavedCols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':k==='shedRemark'?'Shed Remarks':k==='oem'?'OEM Remarks':label}</th>)}</tr></thead><tbody>{data.rows.map((x,i)=><tr key={x._id}>{visibleSavedCols.map(([k])=><td key={k}>{k==='sn'?i+1:k==='date'?x.searchDate:k==='day'?x.weekday:k.startsWith('source:')?<span className="wrap-text">{x.row?.sourceData?.[k.slice(7)]||'—'}</span>:k.startsWith('kavach:')?<span className="wrap-text">{x.row?.kavachSourceData?.[k.slice(7)]||'—'}</span>:<span className="wrap-text">{(k==='shedRemark'?(x.row?.shedRemark||x.shedRemark):k==='oem'?(x.row?.oem||x.oem):x.row?.[k])||'—'}</span>}</td>)}</tr>)}</tbody></table></div>}{data&&data.total===0&&<div className="empty-panel">No saved records match the selected date/DIR filters.</div>}</section>
}

function RecordsModal({open,onClose,fromDate,toDate,dir,metric,groupValue='',makeValue='',allowExcelDownload=false}){
 const [data,setData]=useState(null),[loading,setLoading]=useState(false),[error,setError]=useState(''),[downloadLoading,setDownloadLoading]=useState(false);
 useEffect(()=>{
  if(!open)return;
  let cancelled=false;
  (async()=>{setLoading(true);setError('');try{const{data:x}=await api.get('/summary/metric-records',{params:{fromDate,toDate,dir,metric,groupValue,makeValue}});if(!cancelled)setData(x)}catch(e){if(!cancelled){setData(null);setError(e.response?.data?.message||e.message)}}finally{if(!cancelled)setLoading(false)}})();
  return()=>{cancelled=true};
 },[open,fromDate,toDate,dir,metric,groupValue,makeValue]);
 async function downloadExcel(){
  if(!data?.total)return;
  setDownloadLoading(true);
  try{const r=await api.get('/summary/metric-export',{params:{fromDate,toDate,dir,metric,groupValue,makeValue},responseType:'blob'});downloadBlob(r,'Summary_Metric_Records.xlsx')}catch(e){setError(await blobError(e,'Excel download failed.'))}finally{setDownloadLoading(false)}
 }
 if(!open)return null;
 return <div className="modal-backdrop records-modal-backdrop" role="dialog" aria-modal="true" aria-label="Corresponding records" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}>
  <div className="records-modal-card">
   <button className="modal-close" onClick={onClose} aria-label="Close"><X/></button>
   <div className="records-modal-head"><div><span className="eyebrow">CORRESPONDING DATABASE RECORDS</span><h3>{data?.metricLabel||'Corresponding Records'}</h3><p>{fromDate} to {toDate} · DIR: {dir||'ALL'} · <strong>{data?.total??'—'}</strong> record(s)</p></div>{allowExcelDownload&&data?.total>0&&<button type="button" className="primary-btn records-modal-download" onClick={downloadExcel} disabled={downloadLoading}><FileSpreadsheet/>{downloadLoading?'Preparing Excel...':'Download Excel'}</button>}</div>
   {loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading records...</div>:error?<div className="empty-panel">{error}</div>:data?.rows?.length?<div className="records-modal-table-shell"><table className="data-table records-modal-table"><thead><tr>{cols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':k==='date'?'Date run Kota div':label}</th>)}</tr></thead><tbody>{data.rows.map((x,i)=><tr key={x._id||i}>{cols.map(([k])=><td key={k}>{k==='sn'?i+1:k==='date'?x.searchDate:k==='day'?x.weekday:k.startsWith('source:')?<span className="wrap-text">{x.row?.sourceData?.[k.slice(7)]||'—'}</span>:k.startsWith('kavach:')?<span className="wrap-text">{x.row?.kavachSourceData?.[k.slice(7)]||'—'}</span>:<span className="wrap-text">{(k==='shedRemark'?(x.row?.shedRemark||x.shedRemark):k==='oem'?(x.row?.oem||x.oem):x.row?.[k])||'—'}</span>}</td>)}</tr>)}</tbody></table></div>:<div className="empty-panel">No corresponding records found.</div>}
  </div>
 </div>
}

function RemarksRecordsModal({open,onClose,role,fromDate,toDate,groupValue='',status='pending'}){
  const [data,setData]=useState(null),[loading,setLoading]=useState(false),[error,setError]=useState(''),[downloadLoading,setDownloadLoading]=useState(false);
  useEffect(()=>{if(!open)return;let cancelled=false;(async()=>{setLoading(true);setError('');try{const{data:x}=await api.get('/remarks-pending/records',{params:{role,fromDate,toDate,groupValue,status}});if(!cancelled)setData(x)}catch(e){if(!cancelled){setData(null);setError(e.response?.data?.message||e.message)}}finally{if(!cancelled)setLoading(false)}})();return()=>{cancelled=true}},[open,role,fromDate,toDate,groupValue,status]);
  async function downloadExcel(){if(!data?.total)return;setDownloadLoading(true);try{const r=await api.get('/remarks-pending/export',{params:{role,fromDate,toDate,groupValue,status},responseType:'blob'});downloadBlob(r,`Remarks_${role}_${status}.xlsx`)}catch(e){setError(await blobError(e,'Excel download failed.'))}finally{setDownloadLoading(false)}}
  if(!open)return null;
  return <div className="modal-backdrop records-modal-backdrop" role="dialog" aria-modal="true" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}>
    <div className="records-modal-card">
      <button className="modal-close" onClick={onClose} aria-label="Close"><X/></button>
      <div className="records-modal-head"><div><span className="eyebrow">REMARK RECORDS</span><h3>{role.toUpperCase()} · {groupValue||'All'} · {status.toUpperCase()}</h3><p>{fromDate} to {toDate} · <strong>{data?.total??'—'}</strong> record(s)</p></div>{data?.total>0&&<button type="button" className="primary-btn records-modal-download" onClick={downloadExcel} disabled={downloadLoading}><FileSpreadsheet/>{downloadLoading?'Preparing Excel...':'Download Excel'}</button>}</div>
      {loading?<div className="loading-panel"><LoaderCircle className="spin"/>Loading records...</div>:error?<div className="empty-panel">{error}</div>:data?.rows?.length?<div className="records-modal-table-shell"><table className="data-table records-modal-table"><thead><tr>{cols.map(([k,label])=><th key={k}>{k==='worked'?<>Train Worked<br/>with Kavach</>:k==='noReason'?'If No – Reason':k==='date'?'Date run Kota div':label}</th>)}</tr></thead><tbody>{data.rows.map((x,i)=><tr key={x._id||i}>{cols.map(([k])=><td key={k}>{k==='sn'?i+1:k==='date'?x.searchDate:k==='day'?x.weekday:k.startsWith('source:')?<span className="wrap-text">{x.row?.sourceData?.[k.slice(7)]||'—'}</span>:k.startsWith('kavach:')?<span className="wrap-text">{x.row?.kavachSourceData?.[k.slice(7)]||'—'}</span>:<span className="wrap-text">{(k==='shedRemark'?(x.row?.shedRemark||x.shedRemark):k==='oem'?(x.row?.oem||x.oem):x.row?.[k])||'—'}</span>}</td>)}</tr>)}</tbody></table></div>:<div className="empty-panel">No corresponding records found.</div>}
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
 const metrics=[['targetTrains','No of Target Trains','totalTrains'],['runningWithKavach','TOTAL NO. OF TRAIN RUNNING \nWITH KAVACH','detailKavachMake'],['workingCondition','TOTAL NO. OF TRAIN RUNNING \nWITH KAVACH IN WORKING CONDITION','working'],['nonKavach','NON KAVACH LOCO','nonKavach'],['defective','KAVACH DEFECTIVE LOCO','defective'],['otherMake','OTHER MAKE','otherMake'],['otherReason','OTHER REASON','other']];
 return <table className="excel-detail-table metric-report-table"><thead><tr><th className="sn-col"></th><th>DETAILS</th><th>TOTAL</th><th>UP</th><th>DN</th></tr></thead><tbody>{metrics.map(([k,label,metric],i)=><tr key={k}><td className="sn-col">{i+1}</td><td className={k==='runningWithKavach'||k==='workingCondition'?'wrap-label':''}>{label.split('\n').map((x,j)=><React.Fragment key={j}>{j>0&&<br/>}{x}</React.Fragment>)}</td><td><DetailCount value={data.total[k]} metric={metric} scope="ALL" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td><td><DetailCount value={data.up[k]} metric={metric} scope="UP" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td><td><DetailCount value={data.dn[k]} metric={metric} scope="DN" fromDate={data.fromDate} toDate={data.toDate} onOpen={onOpen}/></td></tr>)}</tbody></table>
}
function DetailDivisionTable({title,rows,onOpen,fromDate,toDate,metric,showStation=false,groupByName=false}){
 return <section className="excel-detail-section"><h4>{title}</h4><table className="excel-detail-table division-report-table"><thead><tr><th className="sn-col"></th><th>{showStation?'Station':''}</th><th>TOTAL</th><th>UP</th><th>DN</th></tr></thead><tbody>{rows.length?rows.map((r,i)=><tr key={`${title}-${r.name}-${i}`}><td className="sn-col">{i+1}</td><td>{showStation?r.station||r.name:r.name}</td><td><DetailCount value={r.total} metric={metric} scope="ALL" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td><td><DetailCount value={r.up} metric={metric} scope="UP" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td><td><DetailCount value={r.dn} metric={metric} scope="DN" fromDate={fromDate} toDate={toDate} groupValue={groupByName?r.name:''} onOpen={onOpen}/></td></tr>):<tr><td className="sn-col">—</td><td>No data</td><td>0</td><td>0</td><td>0</td></tr>}</tbody></table></section>
}
function DetailRecordsTable({title,rows,includeDate=false}){
 const headers=includeDate?['S.N','Date','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark']:['S.N','Train No','Loco No','SHED','KAVACH MAKE','TLC Remark','Shed Remark','OEM Remark'];
 return <section className={`excel-detail-section records-report-section ${includeDate?'records-with-date':''}`}><h4>{title}</h4><div className="detail-record-scroll"><table className="excel-detail-table record-report-table"><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.length?rows.map(r=><tr key={`${title}-${r.sn}`}><td>{r.sn}</td>{includeDate&&<td>{r.date||''}</td>}<td>{r.trainNo||''}</td><td>{r.locoNo||''}</td><td>{r.shed||''}</td><td>{r.kavachMake||''}</td><td>{r.tlcRemark||''}</td><td>{r.shedRemark||''}</td><td>{r.oemRemark||''}</td></tr>):<tr><td>1</td><td colSpan={headers.length-1}>No data</td></tr>}</tbody></table></div></section>
}
function NonKavachAttachDetailTable({rows=[],title='NON KAVACH LOCO ATTACH BY LAST ATTACH DIVISON IN LINK'}){
 const headers=['S.No.','Date','Train No','Loco No','SHED','Loco Attach Div','TLC Remark'];
 return <section className="excel-detail-section records-report-section non-kavach-detail-section"><h4>{title}</h4><div className="detail-record-scroll"><table className="excel-detail-table record-report-table non-kavach-attach-table"><thead><tr>{headers.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.length?rows.map(r=><tr key={`non-kavach-${r.sn}`}><td>{r.sn}</td><td>{r.date||''}</td><td>{r.trainNo||''}</td><td>{r.locoNo||''}</td><td>{r.shed||''}</td><td>{r.locoAttachDiv||''}</td><td className="remark-cell">{r.tlcRemark||''}</td></tr>):<tr><td>1</td><td colSpan={6}>No data</td></tr>}</tbody></table></div></section>
}
function ActiveSessions({setMsg}){
 const [sessions,setSessions]=useState([]),[now,setNow]=useState(Date.now()),[loading,setLoading]=useState(true),[busy,setBusy]=useState('');
 const load=async()=>{try{const{data}=await api.get('/admin/sessions');setSessions(data.sessions||[])}catch(e){setMsg(e.response?.data?.message||'Unable to load active sessions.')}finally{setLoading(false)}};
 useEffect(()=>{load();const poll=setInterval(load,10000),clock=setInterval(()=>setNow(Date.now()),1000);return()=>{clearInterval(poll);clearInterval(clock)}},[]);
 const fmt=ms=>{const sec=Math.max(0,Math.ceil(ms/1000));return `${String(Math.floor(sec/60)).padStart(2,'0')}:${String(sec%60).padStart(2,'0')}`};
 async function terminate(session){if(!window.confirm(`Log out ${session.username} (${session.role.toUpperCase()}) from this session?`))return;setBusy(session.id);try{await api.post(`/admin/sessions/${encodeURIComponent(session.id)}/logout`);setSessions(v=>v.filter(s=>s.id!==session.id));setMsg(`Session for ${session.username} terminated.`)}catch(e){setMsg(e.response?.data?.message||'Unable to terminate session.');load()}finally{setBusy('')}}
 return <section className="panel active-sessions-panel"><div className="panel-intro compact"><div className="intro-icon"><UsersRound/></div><div><span className="eyebrow">ADMINISTRATION</span><h3>Active Sessions</h3><p>Monitor authenticated SHED and OEM sessions and terminate an individual session when required.</p></div></div><div className="table-shell"><table className="active-sessions-table"><thead><tr><th>Username</th><th>Role</th><th>Login Time</th><th>Last Activity</th><th>Time Remaining</th><th>Action</th></tr></thead><tbody>{loading?<tr><td colSpan="6" className="empty-cell">Loading active sessions…</td></tr>:sessions.length?sessions.map(s=>{const remaining=300000-(now-Date.parse(s.lastActivityAt));return <tr key={s.id}><td>{s.username}</td><td><span className={`dir-pill ${s.role}`}>{s.role.toUpperCase()}</span></td><td>{new Date(s.loginAt).toLocaleString()}</td><td>{new Date(s.lastActivityAt).toLocaleString()}</td><td><strong className={remaining<60000?'session-low':''}>{fmt(remaining)}</strong></td><td><button className="session-logout-btn" disabled={busy===s.id} onClick={()=>terminate(s)}>{busy===s.id?'Logging out…':'Logout'}</button></td></tr>}) : <tr><td colSpan="6" className="empty-cell">No active SHED or OEM sessions.</td></tr>}</tbody></table></div></section>
}

function SummaryStatCard({title,count,description,tone='blue'}){
 return <article className={`about-stat-card tone-${tone}`}><div className="about-stat-title" title={title}>{title}</div><div className="about-stat-count">{count}</div><div className="about-stat-description">{description}</div></article>;
}
function AboutUs({anchorTimestamp}){
 const anchorRef=useRef(anchorTimestamp||new Date().toISOString());
 const [summary,setSummary]=useState(null),[summaryLoading,setSummaryLoading]=useState(true),[summaryError,setSummaryError]=useState('');
 useEffect(()=>{let alive=true,inFlight=false;const load=async()=>{if(inFlight)return;inFlight=true;try{const{data}=await api.get('/public/about-summary',{params:{anchor:anchorRef.current}});if(alive){setSummary(data);setSummaryError('')}}catch(e){if(alive)setSummaryError(e.response?.data?.message||'Unable to load summary cards.')}finally{inFlight=false;if(alive)setSummaryLoading(false)}};load();const timer=setInterval(load,60000);return()=>{alive=false;clearInterval(timer)}},[]);
 return <section className="about-page">
  <div className="about-hero">
   <div className="about-hero-icon"><Info/></div>
   <div>
    <span className="eyebrow">ABOUT THE SYSTEM</span>
    <h2>Traction Operation - Kavach Monitoring System</h2>
    <p>A public information page for the Traction Operation workflow of Kota Division, West Central Railway (WCR).</p>
   </div>
  </div>
  <section className="about-summary-section" aria-label="Last 30 days summary">
   <div className="about-summary-heading"><span className="eyebrow">OPERATIONAL SNAPSHOT</span><h3>Last 30 Days Summary</h3><p>{summary?`${summary.fromDate} to ${summary.toDate}`:'Loading reporting period…'} · Values refresh automatically</p></div>
   {summaryError&&<div className="alert"><XCircle/>{summaryError}</div>}
   {(()=>{const cards=[{key:'total-trains',title:'Total Trains',count:summaryLoading?'—':summary?.totalTrains??0,description:'Saved operational records'},{key:'kavach-working',title:'KAVACH WORKING',count:summaryLoading?'—':summary?.kavachWorking??0,description:'Reason is NA or blank'},...(summary?.reasons||[]).map((r,i)=>({key:`reason-${i}-${r.label}`,title:r.label,count:summaryLoading?'—':r.count,description:'Reported reason'}))];const duration=Math.max(28,cards.length*8);return <div className="about-summary-marquee" style={{'--snapshot-duration':`${duration}s`}} onTouchStart={e=>e.currentTarget.classList.add('touch-paused')} onTouchEnd={e=>e.currentTarget.classList.remove('touch-paused')} onTouchCancel={e=>e.currentTarget.classList.remove('touch-paused')}><div className="about-summary-track">{[0,1].map(copy=><div className="about-summary-group" key={copy} aria-hidden={copy===1?true:undefined}>{cards.map((card,i)=><SummaryStatCard key={`${copy}-${card.key}`} title={card.title} count={card.count} description={card.description} tone={['blue','mint','lavender','yellow','coral','teal','rose','sky','peach'][i%9]} />)}</div>)}</div></div>})()}
  </section>
  <div className="about-grid">
   <section className="about-card about-card-wide">
    <div className="about-card-head"><Workflow/><div><span className="eyebrow">WHAT THIS WEBSITE DOES</span><h3>One connected workflow for Kavach monitoring</h3></div></div>
    <p>This website brings together the daily traction roster, Kavach-related loco information, saved operational records, summaries and detailed reports in one workflow. Authorized users enter or update operational information from the Dashboard, while the public-facing reports provide a clear view of the resulting saved data.</p>
    <div className="about-points">
      <div><CheckCircle2/><span>Uses saved operational records as the source for summaries and reports.</span></div>
      <div><CheckCircle2/><span>Supports date-wise and DIR-wise reporting with detailed record views.</span></div>
      <div><CheckCircle2/><span>Provides Excel/PDF outputs for operational review and record keeping.</span></div>
    </div>
   </section>
   <section className="about-card">
    <div className="about-card-head"><ShieldCheck/><div><span className="eyebrow">PUBLIC ACCESS</span><h3>Reports without operational editing</h3></div></div>
    <p>The Summary, Summary &amp; Detail and Saved Data / Export pages are available through the public interface. Operational updates remain part of the authorized workflow.</p>
   </section>
  </div>
  <section className="about-card workflow-card">
   <div className="about-card-head"><ArrowRight/><div><span className="eyebrow">WORKFLOW</span><h3>How the system works</h3></div></div>
   <div className="workflow-steps">
    <div className="workflow-step"><span>01</span><Database/><h4>Source Data</h4><p>Authorized users maintain the required railway source workbooks.</p></div>
    <div className="workflow-step"><span>02</span><TrainFront/><h4>Daily Dashboard</h4><p>Daily train and loco records are reviewed and completed.</p></div>
    <div className="workflow-step"><span>03</span><FileCheck2/><h4>Kavach Lookup &amp; Save</h4><p>Loco details are matched and the completed record is saved or updated.</p></div>
    <div className="workflow-step"><span>04</span><BarChart3/><h4>Summary &amp; Reports</h4><p>Saved records feed date-wise summaries, detailed reports and exports.</p></div>
   </div>
  </section>
  <div className="about-grid role-workflow-grid">
   <section className="about-card">
    <div className="about-card-head"><UsersRound/><div><span className="eyebrow">SHED USER WORKFLOW</span><h3>How Shed users work</h3></div></div>
    <div className="about-points">
      <div><CheckCircle2/><span>Sign in with the assigned Shed account and open the Shed Remark workflow.</span></div>
      <div><CheckCircle2/><span>Review records requiring Shed action, add or update the Shed Remark, and save the change.</span></div>
      <div><CheckCircle2/><span>After the remark is saved, close the record when the Shed-side action is complete.</span></div>
    </div>
   </section>
   <section className="about-card">
    <div className="about-card-head"><ShieldCheck/><div><span className="eyebrow">OEM USER WORKFLOW</span><h3>How OEM users work</h3></div></div>
    <div className="about-points">
      <div><CheckCircle2/><span>Sign in with the assigned OEM account to view records requiring OEM attention.</span></div>
      <div><CheckCircle2/><span>Enter or update the OEM remark for the relevant record and save it.</span></div>
      <div><CheckCircle2/><span>Close the record after the OEM-side remark has been completed.</span></div>
    </div>
   </section>
  </div>
  <section className="about-card contact-card">
   <div className="about-card-head"><UsersRound/><div><span className="eyebrow">TRACTION OPERATION</span><h3>Kota Division (WCR)</h3></div></div>
   <div className="contact-grid">
    <div className="contact-person"><div className="contact-avatar">KM</div><div><h4>KIRTESH MEENA</h4><p>Senior Divisional Electrical Engineer</p><a href="tel:9001017302"><Phone/> CUG: 9001017302</a></div></div>
    <div className="contact-person"><div className="contact-avatar">VK</div><div><h4>VIVEK KHARE</h4><p>Divisional Electrical Engineer</p><a href="tel:9001017305"><Phone/> CUG: 9001017305</a></div></div>
   </div>
   <div className="contact-office"><Mail/><div><strong>TRACTION OPERATION · KOTA DIVISION (WCR)</strong><a href="mailto:srdeetrokota@gmail.com">srdeetrokota@gmail.com</a></div></div>
  </section>
  <div className="about-note"><ShieldCheck/><span>Traction Operation · Kota Division · West Central Railway</span></div>
 </section>
}

function KavachMakeBreakdown({rows=[],fromDate,toDate,onOpen}){
 const columns=[
  ['totalTrains','TOTAL TRAINS','totalTrains'],
  ['working','KAVACH WORKING','working'],
  ['defective','KAVACH DEFECTIVE','defective'],
  ['doNotStart','DO NOT START REMARK/ STICKER','doNotStart'],
  ['isolation','KAVACH ISOLATION PREVIOUS REMARKS','isolation'],
  ['other','OTHER REASON','other'],
  ['total','TOTAL','total']
 ];
 return <section className="excel-detail-section make-summary-section"><h4>KAVACH MAKE SUMMARY BREAK UP</h4><div className="make-summary-scroll"><table className="excel-detail-table make-summary-table"><thead><tr><th>LOCO KAVACH MAKE</th>{columns.map(([key,label])=><th key={key}>{label}</th>)}<th>% WORKING</th></tr></thead><tbody>{rows.length?rows.map(r=><tr key={`make-${r.name}`}><td className="make-name-cell">{r.name}</td>{columns.map(([key,label,metric])=><td key={key}><DetailCount value={r[key]} metric={metric} scope="ALL" fromDate={fromDate} toDate={toDate} onOpen={payload=>onOpen({...payload,makeValue:r.name})}/></td>)}<td>{r.percentWorking}%</td></tr>):<tr><td colSpan={columns.length+2}>No Kavach Make data for the selected dates.</td></tr>}</tbody></table></div></section>
}
function RemarksPending({setMsg,user}){
  const params=useMemo(()=>new URLSearchParams(window.location.search),[]);
  const defaultRole=user?.role==='oem'?'oem':'shed';
  const [role,setRole]=useState(params.get('role')||defaultRole),[fromDate,setFromDate]=useState(params.get('fromDate')||''),[toDate,setToDate]=useState(params.get('toDate')||''),[data,setData]=useState(null),[loading,setLoading]=useState(false),[popup,setPopup]=useState(null);
  const canShed=!user||['admin','admin_limited'].includes(user?.role)||user?.role==='shed',canOem=!user||['admin','admin_limited'].includes(user?.role)||user?.role==='oem';
  useEffect(()=>{if(!canShed&&role==='shed')setRole('oem');if(!canOem&&role==='oem')setRole('shed')},[canShed,canOem,role]);
  async function filter(){
    if(!fromDate||!toDate)return setMsg('Select From Date and To Date before filtering.');
    setLoading(true);setMsg('');
    try{
      const{data:x}=await api.get('/remarks-pending',{params:{role,fromDate,toDate}});
      setData(x);navigatePage('remarksPending',{role,fromDate,toDate},true);
    }catch(e){setData(null);setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}
  }
  function openPopup(groupValue,status){
    setPopup({role,fromDate:data.fromDate,toDate:data.toDate,groupValue,status});
  }
  const isShed=role==='shed';
  const label=isShed?'SHED NAME':'OEM NAME';
  return <section className="panel remarks-pending-panel">
    <div className="panel-intro compact"><div className="intro-icon"><ClipboardList/></div><div><span className="eyebrow">REMARK STATUS REPORT</span><h3>Remarks Pending</h3><p>Track updated and pending remarks from the existing Dashboard database.</p></div></div>
    <div className="filterbar remarks-pending-filterbar">
      <div className="field"><label><UsersRound/> Remark Pending For</label><select className="top-select" value={role} onChange={e=>{setRole(e.target.value);setData(null);setPopup(null)}}>{canShed&&<option value="shed">SHED</option>}{canOem&&<option value="oem">OEM</option>}</select></div>
      <DatePicker label="From Date" value={fromDate} onChange={setFromDate}/><DatePicker label="To Date" value={toDate} onChange={setToDate}/>
      <button className="secondary-btn" onClick={filter} disabled={loading||!fromDate||!toDate}><Search/>{loading?'Fetching...':'Filter'}</button>
      {isShed&&data&&<button className="primary-btn" onClick={async()=>{try{const r=await api.get('/remarks-pending/export',{params:{role,fromDate:data.fromDate,toDate:data.toDate,status:'all',kind:'report'},responseType:'blob'});downloadBlob(r,'Report_Kavach_Shed_Remark_Pending.xlsx')}catch(e){setMsg(await blobError(e,'Report download failed.'))}}}><FileSpreadsheet/>Download Report</button>}
    </div>
    {data&&isShed&&<><div className="summary-meta"><strong>{data.rows.reduce((s,r)=>s+r.total,0)}</strong> eligible SHED record(s) from <strong>{data.fromDate}</strong> to <strong>{data.toDate}</strong>.</div>
      <div className="remarks-pending-table-wrap shed-remarks-scroll"><table className="summary-table remarks-pending-table shed-remarks-report-table"><thead><tr><th>S.No.</th><th>SHED</th><th>RLY</th><th>SHED_NAME</th><th>DIVISION</th><th>Remark Updated</th><th>Remark Pending</th><th>Total</th></tr></thead><tbody>{data.rows.length?data.rows.map(r=><tr key={r.shed}><td>{r.sno}</td><td>{r.shed}</td><td>{r.rly}</td><td>{r.shedName}</td><td>{r.division}</td><td><button type="button" className="count-link" disabled={!r.updated} onClick={()=>openPopup(r.shed,'updated')}>{r.updated}</button></td><td><button type="button" className="count-link" disabled={!r.pending} onClick={()=>openPopup(r.shed,'pending')}>{r.pending}</button></td><td><button type="button" className="count-link" disabled={!r.total} onClick={()=>openPopup(r.shed,'all')}>{r.total}</button></td></tr>):<tr><td colSpan="8" className="empty-cell">No Shed records found for the selected dates.</td></tr>}<tr className="remarks-pending-total-row"><td></td><td>TOTAL</td><td></td><td></td><td></td><td>{data.totals?.updated??data.rows.reduce((s,r)=>s+r.updated,0)}</td><td>{data.totals?.pending??data.rows.reduce((s,r)=>s+r.pending,0)}</td><td>{data.totals?.total??data.rows.reduce((s,r)=>s+r.total,0)}</td></tr></tbody></table></div></>}
    {data&&!isShed&&<><div className="summary-meta"><strong>{data.total}</strong> pending OEM remark(s) from <strong>{data.fromDate}</strong> to <strong>{data.toDate}</strong>.</div><div className="remarks-pending-table-wrap"><table className="summary-table remarks-pending-table"><thead><tr><th>OEM NAME</th><th>REMARKS PENDING</th></tr></thead><tbody>{data.rows.length?data.rows.map(r=><tr key={r.name}><td>{r.name}</td><td><button type="button" className="count-link" onClick={()=>setPopup({role:'oem',fromDate:data.fromDate,toDate:data.toDate,groupValue:r.name,status:'pending'})}>{r.count}</button></td></tr>):<tr><td colSpan="2" className="empty-cell">No pending OEM remarks found for the selected dates.</td></tr>}</tbody></table></div></>}
    {!data&&!loading&&<div className="detail-placeholder">Select the report type, choose a date range and click <strong>Filter</strong> to view the report.</div>}
    <RemarksRecordsModal open={Boolean(popup)} onClose={()=>setPopup(null)} {...(popup||{})}/>
  </section>
}
function SummaryDetail({setMsg}){
 const params=new URLSearchParams(window.location.search); const [fromDate,setFromDate]=useState(params.get('fromDate')||''),[toDate,setToDate]=useState(params.get('toDate')||''),[data,setData]=useState(null),[loading,setLoading]=useState(false),[popup,setPopup]=useState(null);
 async function filter(){if(!fromDate||!toDate)return setMsg('Select From Date and To Date before filtering.');setLoading(true);setMsg('');try{const{data:x}=await api.get('/summary-detail',{params:{fromDate,toDate}});setData(x);navigatePage('detail',{fromDate,toDate})}catch(e){setData(null);setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 async function downloadPdf(){if(!data)return setMsg('Filter the report first.');try{const r=await api.get('/summary-detail/pdf',{params:{fromDate,toDate},responseType:'blob'});downloadBlobFile(r,'Summary_Detail.pdf','application/pdf')}catch(e){setMsg(await blobError(e,'PDF download failed.'))}}
 async function downloadExcel(){if(!data)return setMsg('Filter the report first.');try{const r=await api.get('/summary-detail/excel',{params:{fromDate,toDate},responseType:'blob'});downloadBlob(r,`Summary_Detail_${fromDate}_to_${toDate}.xlsx`)}catch(e){setMsg(await blobError(e,'Excel download failed.'))}}
 return <section className="panel detail-panel"><div className="panel-intro compact no-print"><div className="intro-icon"><FileSpreadsheet/></div><div><span className="eyebrow">DATE-WISE DETAIL REPORT</span><h3>Summary &amp; Detail</h3><p>Generate the report from existing database records for the selected date range.</p></div></div><div className="filterbar detail-filterbar no-print"><DatePicker label="From Date" value={fromDate} onChange={setFromDate}/><DatePicker label="To Date" value={toDate} onChange={setToDate}/><button className="secondary-btn" onClick={filter} disabled={loading||!fromDate||!toDate}><Search/>{loading?'Fetching...':'Filter'}</button>{data&&<><button className="primary-btn" type="button" onClick={downloadPdf}><FileDown/>Download PDF</button><button className="secondary-btn" type="button" onClick={downloadExcel}><FileSpreadsheet/>Download Excel</button><button className="secondary-btn" type="button" onClick={()=>window.print()}><Printer/>Print A4</button></>}</div>{data&&<div className="summary-detail-a4"><div className="a4-report-title">Summary &amp; Detail - Date From&nbsp; {data.fromDate} &nbsp;To&nbsp; {data.toDate}</div><DetailMetricTable data={data} onOpen={setPopup}/><KavachMakeBreakdown rows={data.kavachMakeBreakdown} fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailDivisionTable title="KAVACH DEFECTIVE LOCO" rows={data.defectiveByShed} metric="defectiveByShed" groupByName fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><DetailDivisionTable title="NON KAVACH LOCO" rows={data.nonKavachPlainByStation} metric="nonKavachPlain" groupByName fromDate={data.fromDate} toDate={data.toDate} onOpen={setPopup}/><NonKavachAttachDetailTable rows={data.nonKavachPlainRows} title="NON KAVACH LOCO ATTACH BY LAST ATTACH DIVISON IN LINK"/><DetailRecordsTable title="KAVACH DEFECTIVE LOCO DETAILS" rows={data.defectiveRows} includeDate/><DetailRecordsTable title="OTHER REASON" rows={data.otherRows} includeDate/></div>}{!data&&!loading&&<div className="detail-placeholder no-print">Select a date range and click <strong>Filter</strong> to fetch the report.</div>}<RecordsModal open={Boolean(popup)} onClose={()=>setPopup(null)} {...(popup||{})}/></section>
}
function Summary({setMsg}){
 const params=useMemo(()=>new URLSearchParams(window.location.search),[]),[fromDate,setFromDate]=useState(params.get('fromDate')||''),[toDate,setToDate]=useState(params.get('toDate')||''),[dir,setDir]=useState(params.get('dir')||'UP'),[data,setData]=useState(),[loading,setLoading]=useState(false),[hasFiltered,setHasFiltered]=useState(false),[popup,setPopup]=useState(null); useEffect(()=>{if(fromDate&&toDate&&dir)filter()},[]);
 async function filter(){if(!fromDate||!toDate||!dir)return setMsg('Select From Date, To Date and DIR before filtering.');setLoading(true);setMsg('');try{const{data:x}=await api.get('/summary',{params:{fromDate,toDate,dir}});setData(x);setHasFiltered(true);navigatePage('summary',{fromDate,toDate,dir})}catch(e){setData();setMsg(e.response?.data?.message||e.message)}finally{setLoading(false)}}
 async function download(){if(!data)return setMsg('Filter the summary first.');try{const r=await api.get('/summary/export',{params:{fromDate,toDate,dir},responseType:'blob'});downloadBlob(r,'Traction_Operation_Summary.xlsx')}catch(e){setMsg(await blobError(e,'Summary export failed.'))}}
 const cell=(value,r,metric,strong=false)=> <SummaryCount value={value} date={r.date} endDate={r.date} dir={r.dir} metric={metric} strong={strong} onOpen={setPopup}/>;
 return <section className="panel summary-panel"><div className="panel-intro compact"><div className="intro-icon"><BarChart3/></div><div><span className="eyebrow">DATE-WISE KAVACH SUMMARY</span><h3>Traction Operation Summary</h3><p>Generate the summary from saved database records using Date + DIR. The reference workbook is not used for calculations.</p></div></div><div className="filterbar summary-filterbar"><DatePicker label="From Date" value={fromDate} onChange={setFromDate}/><DatePicker label="To Date" value={toDate} onChange={setToDate}/><div className="field"><label><Filter/> DIR</label><select className="top-select" value={dir} onChange={e=>setDir(e.target.value)}><option value="ALL">All</option><option value="UP">UP</option><option value="DN">DN</option></select></div><button className="secondary-btn" onClick={filter} disabled={loading||!fromDate||!toDate}><Search/>{loading?'Generating...':'Filter'}</button><button className="primary-btn" onClick={download} disabled={!data}><FileSpreadsheet/>Download Summary</button></div>
 {hasFiltered&&data&&<><div className="summary-meta"><strong>{data.totalRecords}</strong> saved record(s) included for <strong>{data.dir}</strong> from <strong>{data.fromDate}</strong> to <strong>{data.toDate}</strong>.</div><div className="summary-table-wrap"><table className="summary-table"><thead><tr><th>DAY</th><th>DATE</th><th>DIR</th><th>TOTAL TRAINS LOCO IN KAVACH SECTION</th><th>NO. OF LOCO KAVACH FITTED</th><th>NO. OF LOCO KAVACH WORKING</th><th>KAVACH DEFECTIVE</th><th>KERNEX MAKE APPROVAL PENDING</th><th>DO NOT START REMARK/ STICKER</th><th>KAVACH ISOLATION PREVIOUS REMARKS</th><th>OTHER REASON</th><th>TOTAL</th><th>% WORKING</th></tr></thead><tbody>{data.rows.map(r=><tr key={`${r.date}-${r.dir}`}><td>{r.day}</td><td>{r.date}</td><td><span className="dir-pill">{r.dir}</span></td><td>{cell(r.totalTrains,r,'totalTrains')}</td><td>{cell(r.fitted,r,'fitted')}</td><td>{cell(r.working,r,'working')}</td><td>{cell(r.defective,r,'defective')}</td><td>{cell(r.kernex,r,'kernex')}</td><td>{cell(r.doNotStart,r,'doNotStart')}</td><td>{cell(r.isolation,r,'isolation')}</td><td>{cell(r.other,r,'other')}</td><td>{cell(r.total,r,'total',true)}</td><td><strong>{r.percentWorking}%</strong></td></tr>)}{data.totalRow&&<tr className="summary-total-row"><td><strong>TOTAL</strong></td><td></td><td><span className="dir-pill">{data.totalRow.dir}</span></td><td>{cell(data.totalRow.totalTrains,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'totalTrains',true)}</td><td>{cell(data.totalRow.fitted,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'fitted',true)}</td><td>{cell(data.totalRow.working,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'working',true)}</td><td>{cell(data.totalRow.defective,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'defective',true)}</td><td>{cell(data.totalRow.kernex,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'kernex',true)}</td><td>{cell(data.totalRow.doNotStart,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'doNotStart',true)}</td><td>{cell(data.totalRow.isolation,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'isolation',true)}</td><td>{cell(data.totalRow.other,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'other',true)}</td><td>{cell(data.totalRow.total,{date:data.fromDate,endDate:data.toDate,dir:data.dir},'total',true)}</td><td><strong>{data.totalRow.percentWorking}%</strong></td></tr>}</tbody></table></div>{data.rows.length===0&&<div className="empty-panel">No summary dates found.</div>}</>}
 <RecordsModal open={Boolean(popup)} onClose={()=>setPopup(null)} allowExcelDownload {...(popup||{})}/></section>
}

function downloadBlob(response,fallback){const cd=response.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||fallback;const url=URL.createObjectURL(new Blob([response.data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}
function downloadBlobFile(response,fallback,type){const cd=response.headers['content-disposition']||'';const m=cd.match(/filename="?([^";]+)"?/i);const filename=m?.[1]||fallback;const url=URL.createObjectURL(new Blob([response.data],{type}));const a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)}
async function blobError(error,fallback){if(error.response?.data instanceof Blob){try{const text=await error.response.data.text();const json=JSON.parse(text);return json.message||fallback}catch{return fallback}}return error.response?.data?.message||error.message||fallback}

function Upload({status,refresh,setMsg}){
 const[tf,setTf]=useState(),[kf,setKf]=useState(),[pw,setPw]=useState(''),[busy,setBusy]=useState(false),[mf,setMf]=useState(),[masterPw,setMasterPw]=useState(''),[masterBusy,setMasterBusy]=useState(false);
 const initialized=!!status?.initialized, masterExists=!!status?.master;
 async function submit(){if(!tf&&!kf)return setMsg('Select at least one Excel file.');if(!pw)return setMsg('Enter the upload password.');setBusy(true);try{const{data:a}=await api.post('/upload/authenticate',{password:pw}),f=new FormData();if(tf)f.append('trainsFile',tf);if(kf)f.append('kavachFile',kf);if(!initialized){if(!tf||!kf){setMsg('For first setup, select both source files.');setBusy(false);return}await api.post('/upload/initial',f,{headers:{Authorization:'Bearer '+a.token}})}else await api.post('/upload/update',f,{headers:{Authorization:'Bearer '+a.token}});setMsg('Source data updated successfully. Dropdown validation lists were refreshed from the workbook.');setPw('');setTf();setKf();await refresh()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setBusy(false)}}
 async function download(key){setMsg('');try{const r=await api.get(`/upload/download/${key}`,{responseType:'blob'});downloadBlob(r,key==='trains'?'TRAINS_PER_DAY.xlsx':'Kavach_Loco_Details.xlsx')}catch(e){setMsg(await blobError(e,'Source download failed.'))}}
 async function uploadMaster(){if(!mf)return setMsg('Select Master_Data_Shed_Rly_Division.xlsx first.');if(masterExists&&!masterPw)return setMsg('Enter the password to replace the existing Master Data file.');setMasterBusy(true);setMsg('');try{const f=new FormData();f.append('masterFile',mf);if(masterPw)f.append('password',masterPw);await api.post('/master-data/upload',f);setMsg(masterExists?'Master Data replaced successfully.':'Master Data uploaded successfully.');setMf();setMasterPw('');await refresh()}catch(e){setMsg(e.response?.data?.message||e.message)}finally{setMasterBusy(false)}}
 async function downloadMaster(){try{const r=await api.get('/master-data/download',{responseType:'blob'});downloadBlob(r,'Master_Data_Shed_Rly_Division.xlsx')}catch(e){setMsg(await blobError(e,'Master Data download failed.'))}}
 return <section className="panel">
  <div className="panel-intro compact"><div className="intro-icon"><Database/></div><div><span className="eyebrow">SOURCE DATA MANAGEMENT</span><h3>Upload &amp; maintain source workbooks</h3><p>Operational source files use the existing upload security. Master Data has a password-free first upload and password-protected replacements.</p></div></div>
  <div className="source-grid"><SourceCard title="TRAINS_PER_DAY.xlsx" file={tf} set={setTf} status={status?.train} onDownload={()=>download('trains')}/><SourceCard title="Kavach_Loco_Details.xlsx" file={kf} set={setKf} status={status?.kavach} onDownload={()=>download('kavach')}/></div>
  <section className="master-data-card">
   <div className="source-card-top"><FileSpreadsheet/><span>{status?.master?.count??0} mappings</span></div>
   <h4>Master_Data_Shed_Rly_Division.xlsx</h4>
   <p>{masterExists?`Stored: ${status.master.fileName}`:'Not uploaded yet — first upload does not require a password.'}</p>{status?.master?.columns?.length>0&&<div className="source-column-list"><strong>{status.master.columns.length} detected columns</strong><div>{status.master.columns.map((c,i)=><span key={`${c.key}-${i}`}>{c.label||c.key}</span>)}</div></div>}
   <label className="picker"><UploadCloud/><span>{mf?mf.name:'Choose Master Data Excel file'}</span><input type="file" accept=".xlsx,.xls,.xlsm" onChange={e=>setMf(e.target.files?.[0])}/></label>
   {masterExists&&<div className="field master-password-field"><label><LockKeyhole/> Password for replacement</label><input type="password" value={masterPw} onChange={e=>setMasterPw(e.target.value)} placeholder="Required only when replacing existing Master Data"/></div>}
   <div className="master-actions"><button className="primary-btn" onClick={uploadMaster} disabled={masterBusy||!mf}><UploadCloud/>{masterBusy?'Uploading...':masterExists?'Replace Master Data':'Upload Master Data'}</button><button className="download-source" onClick={downloadMaster} disabled={!masterExists}><FileDown/>Download stored file</button></div>
  </section>
  <section className="security"><div className="security-head"><div className="security-icon"><LockKeyhole/></div><div><span className="eyebrow">UPLOAD SECURITY</span><h4>{initialized?'Password required for source updates':'Password required for initial source upload'}</h4></div></div><div className="security-form"><input type="password" value={pw} onChange={e=>setPw(e.target.value)} placeholder="Enter upload password"/><button className="primary-btn" onClick={submit} disabled={busy}><LockKeyhole/>{busy?'Processing...':initialized?'Authenticate & Update':'Authenticate & Upload'}</button></div><div className="secure"><CheckCircle2/> Source workbook uploads remain password-protected. Master Data first upload is separate and password-free.</div></section>
 </section>
}
function SourceCard({title,file,set,status,onDownload}){return <section className="source-card"><div className="source-card-top"><FileSpreadsheet/><span>{status?.count??0} rows</span></div><h4>{title}</h4><p>{status?.fileName?`Stored: ${status.fileName}`:'Not uploaded yet'}</p>{status?.columns?.length>0&&<div className="source-column-list"><strong>{status.columns.length} detected columns</strong><div>{status.columns.map((c,i)=><span key={`${c.key}-${i}`}>{c.label||c.key}</span>)}</div></div>}<label className="picker"><UploadCloud/><span>{file?file.name:'Choose .xlsx / .xls / .xlsm file'}</span><input type="file" accept=".xlsx,.xls,.xlsm" onChange={e=>set(e.target.files?.[0])}/></label><button className="download-source" onClick={onDownload} disabled={!status}><FileDown/>Download stored file</button></section>}

createRoot(document.getElementById('root')).render(<App/>);
