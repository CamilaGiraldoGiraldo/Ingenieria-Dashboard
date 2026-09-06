import React,{useMemo,useState} from 'react';
import {createRoot} from 'react-dom/client';
import * as XLSX from 'xlsx';
import {BarChart,Bar,XAxis,YAxis,Tooltip,ResponsiveContainer,PieChart,Pie,Cell,LineChart,Line,CartesianGrid,Legend,AreaChart,Area} from 'recharts';
import {LayoutDashboard,FolderKanban,Users,Server,LifeBuoy,Clock3,Upload,Search,RefreshCw,AlertTriangle,CheckCircle2,TrendingUp,Database,Activity,CalendarDays,BriefcaseBusiness,ChevronRight,Target,Timer,HardHat} from 'lucide-react';
import initial from '../unified.json';
import './styles.css';

const COLORS=['#2563eb','#7c3aed','#0891b2','#059669','#f59e0b','#ef4444','#64748b','#db2777'];
const clean=v=>String(v??'').trim();
const money=n=>new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',maximumFractionDigits:0}).format(Number(n)||0);
const num=n=>new Intl.NumberFormat('es-CO',{maximumFractionDigits:1}).format(Number(n)||0);
const pct=n=>`${num(n)}%`;
const uniq=(arr)=>[...new Set(arr.map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));

const normalizeKey=v=>clean(v).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
const aliases={
  Fecha:['fecha','date','dia'], Colaborador:['colaborador','responsable','ingeniero','empleado','nombrecolaborador'],
  Area:['area'], Proyecto:['proyecto','nombreproyecto','nombredelproyecto','project'], Cliente:['cliente','customer'],
  Actividad:['actividad','tarea','descripcion','description'], Estado:['estado','status'], Prioridad:['prioridad'],
  Horas:['horas','horasreales','horareal','horasregistradas','tiemporeal'],
  'Horas Estimadas':['horasestimadas','horasestimada','estimado','tiempoestimado'],
  'Hora inicio':['horainicio','horadeinicio','inicio','start','starttime'],
  'Hora fin':['horafin','horadefin','fin','end','endtime'],
  'Tipo de actividad':['tipodeactividad','tipoactividad','tipo'],
  'Fecha de compromiso':['fechadecompromiso','fechacompromiso','compromiso','deadline'],
  Reproceso:['reproceso','reprocesos']
};
const canonicalizeRow=row=>{
  const out={...row}; const keys=Object.keys(row);
  Object.entries(aliases).forEach(([canonical,names])=>{
    if(clean(out[canonical])) return;
    const k=keys.find(x=>names.includes(normalizeKey(x)));
    if(k) out[canonical]=row[k];
  });
  if(!out.Colaborador && out.Responsable) out.Colaborador=out.Responsable;
  if(!out.Proyecto && out['Nombre del Proyecto']) out.Proyecto=out['Nombre del Proyecto'];
  return out;
};
const normDate=v=>{
  if(v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,'0')}-${String(v.getDate()).padStart(2,'0')}`;
  if(v==null||v==='') return '';
  if(typeof v==='number'){
    const d=new Date(Math.round((v-25569)*86400*1000)); if(!isNaN(d)) return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,'0')}-${String(d.getUTCDate()).padStart(2,'0')}`;
  }
  const s=clean(v).trim();

  // Excel puede devolver las fechas como "18/08/2026", "18/08/2026 00:00:00"
  // o como "2026-08-18 00:00:00". Normalizamos todos esos formatos a YYYY-MM-DD
  // para que el cálculo de días laborables y, por tanto, la capacidad no quede en 0.
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s|T|$)/);
  if(m)return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;

  m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s|T|$)/);
  if(m)return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;

  return s;
};
const hours=v=>{
  if(v==null||v==='')return 0;
  if(v instanceof Date) return v.getHours()+v.getMinutes()/60+v.getSeconds()/3600;
  if(typeof v==='number') return v<1 ? v*24 : v;
  const s=clean(v).toLowerCase();
  const ampm=s.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)$/);
  if(ampm){let h=+ampm[1],m=+(ampm[2]||0);if(ampm[3]==='pm'&&h<12)h+=12;if(ampm[3]==='am'&&h===12)h=0;return h+m/60;}
  const tm=s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if(tm)return +tm[1]+(+tm[2])/60+(+(tm[3]||0))/3600;
  const n=Number(s.replace(/\s/g,'').replace(',','.'));return Number.isFinite(n)?n:0;
};
const actualHours=r=>{
  const ini=hours(r['Hora inicio']), fin=hours(r['Hora fin']);
  if(clean(r['Hora inicio'])&&clean(r['Hora fin'])&&(fin||ini)) {let d=fin-ini;if(d<0)d+=24;return d;}
  return hours(r.Horas);
};
const numVal=v=>{
  if(v==null||v==='')return 0;
  if(typeof v==='number')return Number.isFinite(v)?v:0;
  let s=clean(v).replace(/[$\s]/g,'');
  // Soporta 4.5, 4,5, 1.234,56 y 1,234.56
  if(s.includes(',')&&s.includes('.')){
    if(s.lastIndexOf(',')>s.lastIndexOf('.')) s=s.replace(/\./g,'').replace(',','.');
    else s=s.replace(/,/g,'');
  }else if(s.includes(',')){
    s=s.replace(',','.');
  }
  const n=Number(s);
  return Number.isFinite(n)?n:0;
};
const HOLIDAYS_2026=new Set(['2026-01-01','2026-01-12','2026-03-23','2026-04-02','2026-04-03','2026-05-01','2026-05-18','2026-06-08','2026-06-15','2026-06-29','2026-07-13','2026-07-20','2026-08-07','2026-08-17','2026-10-12','2026-11-02','2026-11-16','2026-12-08','2026-12-25']);
const workingDays=(a,b)=>{
  if(!a||!b)return 0;let d=new Date(a+'T00:00:00'),e=new Date(b+'T00:00:00'),n=0;
  while(d<=e){const iso=d.toISOString().slice(0,10);const day=d.getDay();if(day!==0&&day!==6&&!HOLIDAYS_2026.has(iso))n++;d.setDate(d.getDate()+1)}return n
};
const isProductive=r=>{const type=clean(r['Tipo de actividad']).toLowerCase(),project=clean(r.Proyecto).toLowerCase();return !!clean(r.Actividad)&&actualHours(r)>0&&!['administrativo','almuerzo','reunión','reunion'].includes(type)&&project!=='almuerzo'};
function App(){
 const [business,setBusiness]=useState(initial.business);const [activities,setActivities]=useState(initial.activities);const [tab,setTab]=useState('Resumen');const [notice,setNotice]=useState('');
 const [filters,setFilters]=useState({q:'',client:'Todos',area:'Todos',person:'Todos',status:'Todos',start:'',end:''});
 const sheets=Object.keys(business);
 const activityDates=useMemo(()=>activities.map(r=>normDate(r.Fecha)).filter(Boolean).sort(),[activities]);
 const start=filters.start||activityDates[0]||'';const end=filters.end||activityDates.at(-1)||'';
 const filteredActivities=useMemo(()=>activities.filter(r=>{const d=normDate(r.Fecha);if(filters.start&&d<filters.start)return false;if(filters.end&&d>filters.end)return false;if(filters.area!=='Todos'&&clean(r['Área'])!==filters.area)return false;if(filters.person!=='Todos'&&clean(r.Colaborador)!==filters.person)return false;if(filters.status!=='Todos'&&clean(r.Estado)!==filters.status)return false;if(filters.q&&!Object.values(r).join(' ').toLowerCase().includes(filters.q.toLowerCase()))return false;return true}),[activities,filters]);
 const population=useMemo(()=>{let p=activities;if(filters.area!=='Todos')p=p.filter(r=>clean(r['Área'])===filters.area);if(filters.person!=='Todos')p=p.filter(r=>clean(r.Colaborador)===filters.person);return uniq(p.map(r=>r.Colaborador))},[activities,filters.area,filters.person]);
 const capPer=workingDays(start,end)*7;const capacity=population.length*capPer;
 const actual=filteredActivities.reduce((a,r)=>a+actualHours(r),0);
 const productive=filteredActivities.filter(isProductive).reduce((a,r)=>a+actualHours(r),0);
 const estimated=filteredActivities.reduce((a,r)=>a+numVal(r['Horas Estimadas']),0);
 const finished=filteredActivities.filter(r=>['finalizada','finalizado','completada','completado','terminada','terminado'].includes(clean(r.Estado).toLowerCase())).length;
 const engKpis={total:filteredActivities.length,actual,productive,estimated,capacity,load:capacity?estimated/capacity*100:0,utilization:capacity?productive/capacity*100:0,completion:filteredActivities.length?finished/filteredActivities.length*100:0,avgAdvance:filteredActivities.length?filteredActivities.reduce((a,r)=>a+numVal(r['% Avance']),0)/filteredActivities.length:0};
 const totals=useMemo(()=>{const c=business.CONSOLIDADO||[];const h=business['BOLSA DE HORAS']||[];const s=business.SERVICIOS||[];const m=business['MESA DE AYUDA']||[];return {projects:Math.max(c.length-1,0),value:c.slice(1).reduce((a,r)=>a+numVal(r[4]),0),billed:c.slice(1).reduce((a,r)=>a+numVal(r[5]),0),services:Math.max(s.length-1,0),hours:h.slice(1).reduce((a,r)=>a+numVal(r[3]),0),used:h.slice(1).reduce((a,r)=>a+numVal(r[4]),0),help:Math.max(m.length-1,0)}} , [business]);
 const projectRows=useMemo(()=>{const c=business.CONSOLIDADO||[];return c.slice(1).filter(r=>r.some(Boolean)).map(r=>({cliente:clean(r[0]),proyecto:clean(r[1]),pc:clean(r[2]),valor:numVal(r[4]),facturado:numVal(r[5]),avance:numVal(r[7])*100,obs:clean(r[8])})).filter(r=>(filters.client==='Todos'||r.cliente===filters.client)&&(!filters.q||Object.values(r).join(' ').toLowerCase().includes(filters.q.toLowerCase())))},[business,filters.client,filters.q]);
 const clientRows=useMemo(()=>{const map=new Map();projectRows.forEach(r=>{if(!map.has(r.cliente))map.set(r.cliente,{name:r.cliente,projects:0,value:0,billed:0});const x=map.get(r.cliente);x.projects++;x.value+=r.valor;x.billed+=r.facturado});return [...map.values()].sort((a,b)=>b.value-a.value)},[projectRows]);
 const serviceRows=useMemo(()=>{const s=business.SERVICIOS||[];return s.slice(1).filter(r=>r.some(Boolean)).map(r=>({entity:clean(r[0]),product:clean(r[1]),availability:clean(r[2]),monthly:numVal(r[3]),cycle:clean(r[4]),duration:clean(r[5]),start:clean(r[6]),renew:clean(r[7])}))},[business]);
 const bolsaRows=useMemo(()=>{const s=business['BOLSA DE HORAS']||[];return s.slice(1).filter(r=>r.some(Boolean)).map(r=>{const contracted=numVal(r[3]),used=numVal(r[4]);return {client:clean(r[0]),type:clean(r[1]),services:clean(r[2]),contracted,used,remaining:Math.max(contracted-used,0),status:clean(r[8])}})},[business]);
 const helpRows=useMemo(()=>{const s=business['MESA DE AYUDA']||[];return s.slice(1).filter(r=>r.some(Boolean)).map(r=>({client:clean(r[0]),type:clean(r[1]),services:clean(r[2]),start:clean(r[4]),end:clean(r[5]),status:clean(r[6])}))},[business]);
 const personData=useMemo(()=>population.map(p=>{const rs=filteredActivities.filter(r=>clean(r.Colaborador)===p);return {name:p,productive:rs.filter(isProductive).reduce((a,r)=>a+actualHours(r),0),estimated:rs.reduce((a,r)=>a+numVal(r['Horas Estimadas']),0),capacity:capPer}}),[population,filteredActivities,capPer]);
 const statusData=useMemo(()=>{const m={};filteredActivities.forEach(r=>m[clean(r.Estado)||'Sin estado']=(m[clean(r.Estado)||'Sin estado']||0)+1);return Object.entries(m).map(([name,value])=>({name,value}))},[filteredActivities]);
 const trendData=useMemo(()=>{const m={};filteredActivities.forEach(r=>{const d=normDate(r.Fecha);if(d)m[d]=(m[d]||0)+actualHours(r)});return Object.entries(m).sort().map(([date,value])=>({date:date.slice(5),hours:Number(value.toFixed(1))}))},[filteredActivities]);
 async function importExcel(file,type){
   if(!file)return;
   try{
     const wb=XLSX.read(await file.arrayBuffer(),{cellDates:true,raw:false});
     if(type==='activities'){
       const preferred=['Base_Actividades','Seguimiento','Seguimiento Diario','Base'];
       const sheet=preferred.find(n=>wb.SheetNames.includes(n))||wb.SheetNames.find(n=>{
         const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:null,raw:false});
         const keys=Object.keys(rows[0]||{}).map(normalizeKey);
         return keys.some(k=>k==='actividad'||k==='tarea') && keys.some(k=>k==='fecha');
       });
       if(!sheet)throw new Error('No se encontró una hoja de seguimiento con Fecha y Actividad');
       const arr=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:null,raw:false})
         .map(canonicalizeRow).filter(r=>Object.values(r).some(v=>clean(v)));
       if(!arr.length)throw new Error('El Excel no contiene registros');
       setActivities(arr);
       setNotice(`Seguimiento KPI cargado: ${file.name} · ${arr.length} registros · ${sheet}`);
     }else{
       const sheet=wb.SheetNames.includes('CONSOLIDADO')?'CONSOLIDADO':wb.SheetNames[0];
       const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheet],{defval:null,raw:false});
       const out={}; out[sheet]=rows;
       setBusiness(out);
       setNotice(`Bitácora comercial cargada: ${file.name} · ${rows.length} registros`);
     }
     setTimeout(()=>setNotice(''),4500);
   }catch(e){setNotice(`No se pudo calcular el Excel: ${e.message||'formato no reconocido'}`)}
 }
 const nav=[['Resumen',LayoutDashboard],['KPIs Ingeniería',Target],['Proyectos',FolderKanban],['Clientes',Users],['Servicios',Server],['Bolsa de horas',Clock3],['Mesa de ayuda',LifeBuoy],['Bitácora',Activity]];
 return <div className="app"><aside><div className="brand"><div className="logo">C</div><div><b>CORE IP</b><span>Centro de control</span></div></div><nav>{nav.map(([n,I])=><button key={n} className={tab===n?'active':''} onClick={()=>setTab(n)}><I size={17}/>{n}</button>)}</nav><div className="source"><Database size={17}/><div><b>Fuentes conectadas</b><span>{sheets.length} hojas comerciales · {activities.length} actividades</span></div></div></aside>
 <main><header><div><p className="eyebrow">CENTRO DE OPERACIONES · CORE IP</p><h1>{tab}</h1><p className="sub">Una sola vista para operación, proyectos, clientes, servicios y desempeño de ingeniería.</p></div><div className="uploads"><label className="upload"><Upload size={16}/> Bitácora comercial<input type="file" accept=".xlsx,.xls" onChange={e=>importExcel(e.target.files?.[0],'business')}/></label><label className="upload secondary"><Upload size={16}/> Seguimiento KPI<input type="file" accept=".xlsx,.xls" onChange={e=>importExcel(e.target.files?.[0],'activities')}/></label></div></header>{notice&&<div className="toast"><CheckCircle2 size={17}/>{notice}</div>}
 {tab==='Resumen'&&<Executive totals={totals} kpi={engKpis} clients={clientRows} serviceRows={serviceRows} personData={personData} statusData={statusData} trendData={trendData} setTab={setTab}/>} 
 {tab==='KPIs Ingeniería'&&<KPIDashboard kpi={engKpis} personData={personData} statusData={statusData} trendData={trendData} filters={filters} setFilters={setFilters} start={start} end={end} population={population} activities={activities}/>} 
 {tab==='Proyectos'&&<Projects rows={projectRows} clients={uniq(projectRows.map(x=>x.cliente))} filters={filters} setFilters={setFilters}/>} 
 {tab==='Clientes'&&<Clients data={clientRows}/>} {tab==='Servicios'&&<Services data={serviceRows}/>} {tab==='Bolsa de horas'&&<Bolsa data={bolsaRows}/>} {tab==='Mesa de ayuda'&&<Help data={helpRows}/>} {tab==='Bitácora'&&<Bitacora rows={filteredActivities} filters={filters} setFilters={setFilters} start={start} end={end} population={population}/>} 
 </main></div>
}
function KPI({icon:I,label,value,detail,tone='blue'}){return <div className="kpi"><div className="kpiTop"><span>{label}</span><span className={'ico '+tone}><I size={17}/></span></div><strong>{value}</strong><small>{detail}</small></div>}
function Filters({filters,setFilters,activities}){const upd=(k,v)=>setFilters(f=>({...f,[k]:v}));const dates=activities.map(r=>normDate(r.Fecha)).filter(Boolean).sort();return <div className="filters"><div className="search"><Search size={17}/><input placeholder="Buscar actividad, proyecto, cliente..." value={filters.q} onChange={e=>upd('q',e.target.value)}/></div><select value={filters.area} onChange={e=>upd('area',e.target.value)}><option>Todos</option>{uniq(activities.map(r=>r['Área'])).map(x=><option key={x}>{x}</option>)}</select><select value={filters.person} onChange={e=>upd('person',e.target.value)}><option>Todos</option>{uniq(activities.map(r=>r.Colaborador)).map(x=><option key={x}>{x}</option>)}</select><select value={filters.status} onChange={e=>upd('status',e.target.value)}><option>Todos</option>{uniq(activities.map(r=>r.Estado)).map(x=><option key={x}>{x}</option>)}</select><input type="date" min={dates[0]} max={dates.at(-1)} value={filters.start} onChange={e=>upd('start',e.target.value)}/><input type="date" min={dates[0]} max={dates.at(-1)} value={filters.end} onChange={e=>upd('end',e.target.value)}/><button className="ghost" onClick={()=>setFilters({q:'',client:'Todos',area:'Todos',person:'Todos',status:'Todos',start:'',end:''})}><RefreshCw size={15}/> Limpiar</button></div>}
function Executive({totals,kpi,clients,serviceRows,personData,statusData,trendData,setTab}){const top=clients.slice(0,8).map(x=>({name:x.name.length>14?x.name.slice(0,14)+'…':x.name,value:x.value}));return <><section className="hero"><div><span className="badge">CONTROL EJECUTIVO</span><h2>Operación + Ingeniería en una sola pantalla</h2><p>Los datos financieros y comerciales del Excel se complementan con los KPIs de carga, productividad, capacidad y utilización.</p></div><div className="heroStats"><div><b>{totals.projects}</b><span>proyectos</span></div><div><b>{totals.help}</b><span>contratos soporte</span></div><div><b>{num(totals.hours)}</b><span>h de bolsa</span></div></div></section><div className="sectionTitle"><div><h2>KPIs del negocio</h2><p>Panorama comercial y operativo.</p></div></div><section className="grid kpis"><KPI icon={FolderKanban} label="Proyectos" value={totals.projects} detail={money(totals.value)+' contratados'} /><KPI icon={TrendingUp} label="Facturación" value={money(totals.billed)} detail={totals.value?`${pct(totals.billed/totals.value*100)} del valor contratado`:'N/D'} tone="purple"/><KPI icon={Server} label="Servicios" value={totals.services} detail="Registros del portafolio" tone="cyan"/><KPI icon={Clock3} label="Bolsa" value={`${num(totals.used)} h`} detail={`${num(totals.hours)} h contratadas`} tone="orange"/></section><div className="sectionTitle"><div><h2>KPIs de Ingeniería</h2><p>Capacidad calculada a 7 h por día laboral y por colaborador.</p></div><button className="ghost" onClick={()=>setTab('KPIs Ingeniería')}>Abrir análisis completo <ChevronRight size={15}/></button></div><section className="grid kpis"><KPI icon={Activity} label="Actividades" value={kpi.total} detail={`${num(kpi.actual)} h registradas`} /><KPI icon={Timer} label="Carga estimada" value={pct(kpi.load)} detail={`${num(kpi.estimated)} h / ${num(kpi.capacity)} h capacidad`} tone={kpi.load>100?'red':kpi.load>=80?'orange':'blue'}/><KPI icon={Target} label="Utilización" value={pct(kpi.utilization)} detail={`${num(kpi.productive)} h productivas`} tone={kpi.utilization>=80?'green':'cyan'}/><KPI icon={CheckCircle2} label="Finalización" value={pct(kpi.completion)} detail={`${filteredText(kpi.total,kpi.completion)} actividades`} tone="green"/></section><section className="chartGrid"><Panel title="Valor por cliente" subtitle="Concentración comercial"><ResponsiveContainer width="100%" height={280}><BarChart data={top}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="name" tick={{fontSize:10}}/><YAxis tickFormatter={v=>`${Math.round(v/1e6)}M`} tick={{fontSize:10}}/><Tooltip formatter={v=>money(v)}/><Bar dataKey="value" fill="#2563eb" radius={[6,6,0,0]}/></BarChart></ResponsiveContainer></Panel><Panel title="Estado de actividades" subtitle="Distribución del trabajo registrado"><ResponsiveContainer width="100%" height={280}><PieChart><Pie data={statusData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} label>{statusData.map((_,i)=><Cell key={i} fill={COLORS[i%COLORS.length]}/>)}</Pie><Tooltip/></PieChart></ResponsiveContainer></Panel><Panel title="Horas por día" subtitle="Evolución del esfuerzo registrado"><ResponsiveContainer width="100%" height={280}><AreaChart data={trendData}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="date"/><YAxis/><Tooltip/><Area type="monotone" dataKey="hours" fill="#dbeafe" stroke="#2563eb"/></AreaChart></ResponsiveContainer></Panel></section><section className="panel"><div className="panelHead"><div><h2>Alertas de gestión</h2><p>Prioriza lo que requiere intervención.</p></div></div><div className="alertGrid"><Alert icon={AlertTriangle} title="Capacidad" text={kpi.load>100?'La carga estimada supera la capacidad.':'La carga está dentro de la capacidad disponible.'} tone={kpi.load>100?'danger':'ok'}/><Alert icon={Clock3} title="Bolsa de horas" text={totals.hours-totals.used<10?'Revisar saldos bajos de bolsa.':'Consumo de bolsa bajo control.'} tone={totals.hours-totals.used<10?'warn':'ok'}/><Alert icon={BriefcaseBusiness} title="Facturación" text={totals.value&&totals.billed<totals.value?'Hay valor contratado aún no facturado.':'Facturación alineada al valor registrado.'} tone="warn"/></div></section></>}
function filteredText(total,completion){return `${Math.round(total*completion/100)} de ${total}`}
function Alert({icon:I,title,text,tone}){return <div className={'alert '+tone}><I size={19}/><div><b>{title}</b><p>{text}</p></div></div>}
function Panel({title,subtitle,children}){return <div className="panel"><div className="panelHead"><div><h2>{title}</h2><p>{subtitle}</p></div></div>{children}</div>}
function KPIDashboard({kpi,personData,statusData,trendData,filters,setFilters,start,end,population,activities}){return <><Filters filters={filters} setFilters={setFilters} activities={activities}/><section className="infoBanner"><CalendarDays size={19}/><div><b>Capacidad real del periodo</b><span>{start||'N/D'} → {end||'N/D'} · {workingDays(start,end)} días laborables · 7 h/persona/día · {population.length} colaboradores</span></div><strong>{num(kpi.capacity)} h</strong></section><section className="grid kpis six"><KPI icon={Activity} label="Actividades" value={kpi.total} detail={`${num(kpi.actual)} h registradas`}/><KPI icon={Timer} label="Horas productivas" value={`${num(kpi.productive)} h`} detail="Excluye almuerzo/administrativo" tone="cyan"/><KPI icon={BriefcaseBusiness} label="Carga estimada" value={pct(kpi.load)} detail={`${num(kpi.estimated)} h estimadas`} tone={kpi.load>100?'red':kpi.load>=80?'orange':'blue'}/><KPI icon={Target} label="Utilización" value={pct(kpi.utilization)} detail={`${num(kpi.productive)} / ${num(kpi.capacity)} h`} tone={kpi.utilization>=80?'green':'cyan'}/><KPI icon={CheckCircle2} label="Finalización" value={pct(kpi.completion)} detail="Actividades finalizadas" tone="green"/><KPI icon={TrendingUp} label="Avance promedio" value={pct(kpi.avgAdvance)} detail="% Avance registrado" tone="purple"/></section><section className="chartGrid"><Panel title="Capacidad por colaborador" subtitle="Asignado vs capacidad"><ResponsiveContainer width="100%" height={300}><BarChart data={personData}><CartesianGrid strokeDasharray="3 3" vertical={false}/><XAxis dataKey="name" tick={{fontSize:9}}/><YAxis/><Tooltip/><Legend/><Bar dataKey="capacity" name="Capacidad" fill="#cbd5e1"/><Bar dataKey="estimated" name="Estimado" fill="#7c3aed" radius={[5,5,0,0]}/></BarChart></ResponsiveContainer></Panel><Panel title="Horas productivas" subtitle="Por colaborador"><ResponsiveContainer width="100%" height={300}><BarChart data={personData}><XAxis dataKey="name" tick={{fontSize:9}}/><YAxis/><Tooltip/><Bar dataKey="productive" fill="#0891b2" radius={[5,5,0,0]}/></BarChart></ResponsiveContainer></Panel><Panel title="Tendencia diaria" subtitle="Horas registradas"><ResponsiveContainer width="100%" height={300}><LineChart data={trendData}><CartesianGrid strokeDasharray="3 3"/><XAxis dataKey="date"/><YAxis/><Tooltip/><Line type="monotone" dataKey="hours" stroke="#2563eb" strokeWidth={3}/></LineChart></ResponsiveContainer></Panel></section><Panel title="Semáforo de capacidad" subtitle="Más de 100% = sobrecarga · 80–100% = atención · menos de 80% = normal"><div className="capacityGrid">{personData.map(x=>{const load=x.capacity?x.estimated/x.capacity*100:0;return <div className="capacityCard" key={x.name}><b>{x.name}</b><div><span>Capacidad</span><strong>{num(x.capacity)} h</strong></div><div><span>Asignado</span><strong>{num(x.estimated)} h</strong></div><div><span>Disponible</span><strong>{num(Math.max(x.capacity-x.estimated,0))} h</strong></div><div className="loadBar"><span className={load>100?'over':''} style={{width:`${Math.min(load,130)/1.3}%`}}/></div><small>{pct(load)} de carga</small></div>})}</div></Panel></>}
function Projects({rows,clients,filters,setFilters}){return <><div className="toolbar"><div className="search"><Search size={17}/><input placeholder="Buscar proyecto, cliente, observación..." value={filters.q} onChange={e=>setFilters(f=>({...f,q:e.target.value}))}/></div><select value={filters.client} onChange={e=>setFilters(f=>({...f,client:e.target.value}))}><option>Todos</option>{clients.map(c=><option key={c}>{c}</option>)}</select><button className="ghost" onClick={()=>setFilters(f=>({...f,q:'',client:'Todos'}))}><RefreshCw size={15}/> Limpiar</button></div><Table title={`${rows.length} proyectos`} subtitle="Consolidado financiero y operativo." heads={['Cliente','Proyecto','Valor','Facturado','Avance','Observación']} rows={rows.map(x=><><td><b>{x.cliente}</b></td><td>{x.proyecto}</td><td>{money(x.valor)}</td><td>{money(x.facturado)}</td><td><Progress v={x.avance}/></td><td>{x.obs||'—'}</td></>)}/></>}
function Clients({data}){return <Table title="Clientes" subtitle="Ranking por valor contratado." heads={['Cliente','Proyectos','Valor','Facturado','Avance facturación']} rows={data.map(x=><><td><b>{x.name}</b></td><td>{x.projects}</td><td>{money(x.value)}</td><td>{money(x.billed)}</td><td><Progress v={x.value?x.billed/x.value*100:0}/></td></>)}/>} 
function Services({data}){return <Table title="Servicios" subtitle="Portafolio, facturación y renovaciones." heads={['Entidad','Producto','Disponibilidad','Mensual','Ciclo','Duración','Renovación']} rows={data.map(x=><><td><b>{x.entity}</b></td><td>{x.product}</td><td>{x.availability||'—'}</td><td>{money(x.monthly)}</td><td>{x.cycle||'—'}</td><td>{x.duration||'—'}</td><td>{x.renew||'—'}</td></>)}/>} 
function Bolsa({data}){return <Table title="Bolsa de horas" subtitle="Consumo y saldo por cliente." heads={['Cliente','Servicio','Contratadas','Utilizadas','Restantes','Estado']} rows={data.map(x=><><td><b>{x.client}</b></td><td>{x.services}</td><td>{num(x.contracted)} h</td><td>{num(x.used)} h</td><td><span className={'tag '+(x.remaining<=2?'warn':'ok')}>{num(x.remaining)} h</span></td><td>{x.status||'—'}</td></>)}/>} 
function Help({data}){return <Table title="Mesa de ayuda" subtitle="Contratos y cobertura de soporte." heads={['Cliente','Contrato','Servicios','Inicio','Fin','Estado']} rows={data.map(x=><><td><b>{x.client}</b></td><td>{x.type||'—'}</td><td>{x.services||'—'}</td><td>{x.start||'—'}</td><td>{x.end||'—'}</td><td><span className="tag ok">{x.status||'Sin estado'}</span></td></>)}/>} 
function Bitacora({rows,filters,setFilters,start,end,population}){return <><Filters filters={filters} setFilters={setFilters} activities={rows}/><section className="infoBanner"><Activity size={18}/><div><b>Bitácora filtrada</b><span>{start} → {end} · {population.length} colaboradores en población</span></div><strong>{rows.length} registros</strong></section><Table title="Detalle de actividades" subtitle="Fuente: Base_Actividades del seguimiento diario." heads={['Fecha','Colaborador','Área','Proyecto','Actividad','Tipo','Horas','Estimadas','Estado','Avance']} rows={rows.map((r,i)=><><td>{normDate(r.Fecha)}</td><td><b>{clean(r.Colaborador)}</b></td><td>{clean(r['Área'])}</td><td>{clean(r.Proyecto)}</td><td className="wide">{clean(r.Actividad)}</td><td>{clean(r['Tipo de actividad'])}</td><td>{num(actualHours(r))} h</td><td>{num(numVal(r['Horas Estimadas']))} h</td><td>{clean(r.Estado)}</td><td>{num(Number(r['% Avance'])||0)}%</td></>)}/></>}
function Table({title,subtitle,heads,rows}){return <section className="panel tablePanel"><div className="panelHead"><div><h2>{title}</h2><p>{subtitle}</p></div></div><div className="tableWrap"><table><thead><tr>{heads.map(h=><th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((r,i)=><tr key={i}>{r}</tr>)}</tbody></table></div></section>}
function Progress({v}){const n=Math.max(0,Math.min(100,Number(v)||0));return <div className="progCell"><div className="progress"><span style={{width:n+'%'}}/></div><small>{num(n)}%</small></div>}
createRoot(document.getElementById('root')).render(<App/>);
