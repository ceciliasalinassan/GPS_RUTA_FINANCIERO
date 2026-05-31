
import React,{Component,useEffect,useMemo,useState}from"react";
import{createRoot}from"react-dom/client";
import{AlertTriangle,Bell,CalendarDays,CheckCircle,Clock,CreditCard,Edit,Eye,FileText,Lock,LogOut,Mail,Paperclip,Plus,Search,ShieldCheck,Trash2,TrendingDown,TrendingUp,UploadCloud,User,Users,Wallet,Save,Building2,MessageCircle,Bot,Sparkles,Download,Upload,HardDrive}from"lucide-react";
import{ComposedChart,Bar,Line,XAxis,YAxis,Tooltip,ResponsiveContainer,CartesianGrid,PieChart,Pie,Cell,Legend,RadarChart,PolarGrid,PolarAngleAxis,PolarRadiusAxis,Radar}from"recharts";
import * as XLSX from "xlsx";
import"./style.css";

const KEY="gpsruta_financiero_pro_v2";
const SESSION="gpsruta_login";
const PASS="1234";
const SENDER_EMAIL="gpsruta007@outlook.com";
const incomeCats=["Venta de GPS","Instalación de GPS","Servicio mensual","Pago de factura","Abono cliente","Otro ingreso"];
const expenseCats=["Pago instalador","Pago IVA","Pago contadora","Cotizaciones","Sueldos","Pago internet","Pago plataforma","Compra de equipos","Otro egreso"];
const seed={clients:[
{id:1,nombre:"Transportes del Sur SpA",rut:"76.543.210-9",giro:"Transporte de carga",telefono:"56912345678",email:"contacto@delsur.cl",direccion:"Santiago",contacto:"Juan Pérez"},
{id:2,nombre:"Constructora Andes Ltda.",rut:"77.555.333-1",giro:"Construcción",telefono:"56998765432",email:"pagos@andes.cl",direccion:"San Carlos",contacto:"María Torres"}],
invoices:[
{id:1,clienteId:1,factura:"FAC-2026-001",emision:"2026-05-01",vencimiento:"2026-05-18",monto:1250000,estado:"Pendiente",detalle:"Servicio GPS mensual"},
{id:2,clienteId:2,factura:"FAC-2026-002",emision:"2026-04-20",vencimiento:"2026-05-10",monto:2850000,estado:"Vencida",detalle:"Instalación y monitoreo"},
{id:3,clienteId:1,factura:"FAC-2026-003",emision:"2026-05-04",vencimiento:"2026-05-26",monto:950000,estado:"Pagada",detalle:"Mantención plataforma"}],
incomes:[{id:1,fecha:"2026-05-10",categoria:"Pago de factura",descripcion:"Pago FAC-2026-003",monto:950000,facturaId:3}],
expenses:[{id:1,fecha:"2026-05-11",categoria:"Pago plataforma",descripcion:"Servidor",monto:180000},{id:2,fecha:"2026-05-13",categoria:"Pago contadora",descripcion:"Honorarios contables",monto:75000}],
debts:[{id:1,fecha:"2026-05-15",proveedor:"Proveedor GPS",emailProveedor:"proveedor@gps.cl",categoria:"Compra de equipos",descripcion:"Equipos GPS por pagar",monto:1200000,vencimiento:"2026-05-30",estado:"Pendiente"}],
attachments:{},tasks:[]};

const money=v=>new Intl.NumberFormat("es-CL",{style:"currency",currency:"CLP",maximumFractionDigits:0}).format(+v||0);
const today=()=>new Date().toISOString().slice(0,10);
const mk=d=>(d||today()).slice(0,7);
const ml=k=>{let[a,b]=k.split("-");return `${b}/${a}`};
function load(){try{let s=JSON.parse(localStorage.getItem(KEY));if(!s)return seed;return{...seed,...s,clients:(s.clients||seed.clients).map(c=>({giro:"",...c})),incomes:s.incomes||seed.incomes,expenses:s.expenses||seed.expenses,debts:s.debts||seed.debts}}catch{return seed}}
function days(d){let n=new Date();n.setHours(0,0,0,0);return Math.ceil((new Date(d+"T00:00:00")-n)/86400000)}
function ist(i){if(i.estado==="Pagada")return{l:"Pagada",c:"ok",I:CheckCircle};let d=days(i.vencimiento);if(i.estado==="Vencida"||d<0)return{l:"Vencida",c:"bad",I:AlertTriangle};if(d<=3&&d>=0)return{l:"Por vencer",c:"warn",I:Clock};return{l:"Pendiente",c:"soft",I:FileText}}
function isVencida15(i){
  const d=days(i.vencimiento);
  return ist(i).l==="Vencida" && d>=-15 && d<0;
}
function isPorVencer15(i){
  const d=days(i.vencimiento);
  return ist(i).l!=="Pagada" && d>=0 && d<=15;
}
function reminderText(i){return `Estimado cliente, se recuerda su Factura ${i.factura} por la suma de ${money(i.monto)}. Saludos Cordiales GpsRuta`}
function invoiceSendText(i){return `Estimado cliente, se adjunta Factura ${i.factura} por la suma de ${money(i.monto)}. Saludos Cordiales. GpsRuta.`}
function waReminder(i,c){return `https://wa.me/${c?.telefono||""}?text=${encodeURIComponent(reminderText(i))}`}
function emailReminder(i,c){let s=`Recordatorio factura ${i.factura}`;let b=reminderText(i)+`\n\nCorreo cobranza: ${SENDER_EMAIL}`;return `mailto:${c?.email||""}?cc=${encodeURIComponent(SENDER_EMAIL)}&subject=${encodeURIComponent(s)}&body=${encodeURIComponent(b)}`}
function waInvoice(i,c){return `https://wa.me/${c?.telefono||""}?text=${encodeURIComponent(invoiceSendText(i))}`}
function emailInvoice(i,c){let s=`Envío factura ${i.factura}`;let b=invoiceSendText(i)+`\n\nNota: adjuntar factura antes de enviar.\nCorreo cobranza: ${SENDER_EMAIL}`;return `mailto:${c?.email||""}?cc=${encodeURIComponent(SENDER_EMAIL)}&subject=${encodeURIComponent(s)}&body=${encodeURIComponent(b)}`}
function wa(i,c){return waReminder(i,c)}
function em(i,c){return emailReminder(i,c)}
function debtEmail(d){
  const subject = `Consulta / pago de factura pendiente - ${d.proveedor}`;
  const body = `Estimados,\n\nJunto con saludar, informamos que tenemos registrada una factura/deuda pendiente por pagar.\n\nProveedor: ${d.proveedor}\nDetalle: ${d.descripcion}\nMonto: ${money(d.monto)}\nFecha de vencimiento: ${d.vencimiento}\n\nFavor confirmar datos de pago o estado de la factura.\n\nSaludos cordiales.\nGPSRUTA.\n\nCorreo cobranza: ${SENDER_EMAIL}`;
  return `mailto:${d.emailProveedor||""}?cc=${encodeURIComponent(SENDER_EMAIL)}&subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}
function W(){return <svg viewBox="0 0 32 32" width="18" height="18"><path fill="currentColor" d="M16.04 3C8.86 3 3 8.84 3 16.02c0 2.3.6 4.54 1.75 6.51L3 29l6.64-1.7a12.95 12.95 0 0 0 6.4 1.64h.01C23.22 28.94 29 23.1 29 15.92 29 8.8 23.18 3 16.04 3Zm7.56 18.45c-.32.9-1.86 1.7-2.6 1.8-.67.1-1.52.14-2.45-.15-.56-.18-1.28-.42-2.2-.82-3.87-1.68-6.4-5.6-6.6-5.86-.2-.26-1.58-2.1-1.58-4s1-2.84 1.35-3.23c.36-.4.78-.5 1.04-.5h.75c.24.01.57-.09.9.69.32.78 1.1 2.69 1.2 2.89.1.2.16.43.03.69-.13.26-.2.42-.4.64-.2.23-.42.5-.6.67-.2.2-.4.42-.17.82.23.4 1.02 1.68 2.2 2.72 1.51 1.35 2.78 1.77 3.18 1.97.4.2.63.17.86-.1.23-.26 1-1.16 1.26-1.56.26-.4.53-.33.9-.2.36.13 2.3 1.08 2.7 1.28.4.2.66.3.76.46.1.16.1.95-.22 1.85Z"/></svg>}
function Logo(){return <div className="logo"><div className="pin">⌖</div><div><h1><span>GPS</span><b>ruta</b><small>.cl</small></h1><p>SEGUIMIENTO Y SEGURIDAD</p></div></div>}
function Login({onLogin}){const[p,setP]=useState(""),[e,setE]=useState("");return <div className="loginPage"><form className="loginCard" onSubmit={x=>{x.preventDefault();if(p===PASS){sessionStorage.setItem(SESSION,"1");onLogin()}else setE("Clave incorrecta. Clave demo: 1234")}}><Logo/><h2>Ingreso Seguro</h2><p>Sistema financiero y cobranza</p><div className="loginInput"><Lock size={18}/><input type="password" value={p} onChange={x=>setP(x.target.value)} placeholder="Clave de acceso"/></div>{e&&<div className="error">{e}</div>}<button className="primary full"><ShieldCheck size={18}/>Ingresar</button></form></div>}
function K({t,v,s,icon:Icon,tone="green"}){return <div className="card kpi"><div className={`kpiIcon ${tone}`}><Icon size={32}/></div><div><small>{t}</small><h3>{v}</h3><p>{s}</p></div></div>}
function Fields({obj,set,fields}){return <div className="formGrid">{fields.map(f=><input key={f} value={obj[f]||""} onChange={e=>set({...obj,[f]:e.target.value})} placeholder={f.toUpperCase()} type={["fecha","emision","vencimiento"].includes(f)?"date":f==="monto"?"number":"text"}/>)}</div>}
function InvTable({items,client,edit,del,data={},attachFile=()=>{},canSendInvoice=()=>true}){return <div className="tableWrap"><table><thead><tr><th>Factura</th><th>Cliente</th><th>Vence</th><th>Mes/Año</th><th>Monto</th><th>Estado</th><th>Acciones</th></tr></thead><tbody>{items.map(i=>{let c=client(i.clienteId),s=ist(i),Icon=s.I;return <tr key={i.id}><td><b>{i.factura}</b></td><td>{c?.nombre}</td><td>{i.vencimiento}</td><td>{ml(mk(i.vencimiento))}</td><td>{money(i.monto)}</td><td><span className={`status ${s.c}`}><Icon size={14}/>{s.l}</span></td><td><div className="actions"><label className="icon attachMini" title="Adjuntar factura"><Paperclip size={17}/><input type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.xls,.xlsx" onChange={e=>attachFile(i.id,e.target.files?.[0])}/></label><a className="icon whatsapp" title="Recordatorio WhatsApp" href={waReminder(i,c)} target="_blank"><W/></a><a className="icon mail" title="Recordatorio correo manual" href={emailReminder(i,c)}><Mail size={17}/></a><button className="icon autoMailIcon" title="Recordatorio automático Outlook" onClick={()=>sendAutoReminder(i,c)}>AUTO</button><button className={`icon autoInvoiceIcon ${!data.attachments?.[i.id]?"disabled":""}`} title="Enviar factura automática Outlook" onClick={()=>sendAutoInvoice(i,c)}>PDF</button><a className={`icon invoiceSend ${!data.attachments?.[i.id]?"disabled":""}`} title="Enviar factura WhatsApp" href={data.attachments?.[i.id]?waInvoice(i,c):"#"} onClick={e=>{if(!canSendInvoice(i.id))e.preventDefault()}} target="_blank">FAC</a><button className="icon edit" onClick={()=>edit(i)}><Edit size={17}/></button><button className="icon trash" onClick={()=>del(i.id)}><Trash2 size={17}/></button>{data.attachments?.[i.id]&&<span className="attachedOk">Adjunta</span>}</div></td></tr>})}</tbody></table></div>}

class ErrorBoundary extends Component{
  constructor(props){
    super(props);
    this.state={hasError:false,error:null};
  }
  static getDerivedStateFromError(error){
    return {hasError:true,error};
  }
  componentDidCatch(error,info){
    console.error("GPSRUTA Error:",error,info);
  }
  render(){
    if(this.state.hasError){
      return <div className="fatalError">
        <div>
          <h1>GPSRUTA</h1>
          <h2>Se detectó un error visual del sistema</h2>
          <p>La información guardada no se ha perdido. Recarga la página o importa tu último respaldo.</p>
          <button onClick={()=>location.reload()}>Recargar sistema</button>
        </div>
      </div>;
    }
    return this.props.children;
  }
}

function App(){const[logged,setLogged]=useState(()=>sessionStorage.getItem(SESSION)==="1"),[data,setData]=useState(load),[tab,setTab]=useState("dashboard"),[clock,setClock]=useState(new Date()),[search,setSearch]=useState(""),[taskText,setTaskText]=useState(""),[historyQuickFilter,setHistoryQuickFilter]=useState(""),[historyMonthFilter,setHistoryMonthFilter]=useState("Todos"),[historyStatusFilter,setHistoryStatusFilter]=useState("Todos"),[alertSearch,setAlertSearch]=useState(""),[reminderStatusFilter,setReminderStatusFilter]=useState("Todas"),[saved,setSaved]=useState("Sin cambios"),[selectedInvoiceId,setSelectedInvoiceId]=useState(null),[selectedMonth,setSelectedMonth]=useState(mk(today())),[invoiceFolderMonth,setInvoiceFolderMonth]=useState("Todas"),[invoiceStatusFilter,setInvoiceStatusFilter]=useState("Vencida"),[invoicePage,setInvoicePage]=useState(1),[chatOpen,setChatOpen]=useState(true),[chatInput,setChatInput]=useState(""),[emailSending,setEmailSending]=useState(false),[chatMessages,setChatMessages]=useState([{role:"ia",text:"Hola, soy la LUXURY GPSRUTA. Pregúntame: ¿quién debe más?, ¿cuál es la factura más antigua?, ¿facturas vencidas?, ¿clientes premium?, ¿resumen del mes?"}]);
const[clientForm,setClientForm]=useState({nombre:"",rut:"",giro:"",telefono:"569",email:"",direccion:"",contacto:""}),[invoiceForm,setInvoiceForm]=useState({clienteId:"",factura:"",emision:today(),vencimiento:today(),monto:"",estado:"Pendiente",detalle:""}),[incomeForm,setIncomeForm]=useState({fecha:today(),categoria:"Pago de factura",descripcion:"",monto:"",facturaId:""}),[expenseForm,setExpenseForm]=useState({fecha:today(),categoria:"Pago instalador",descripcion:"",monto:"",debtId:"",numeroFacturaPago:""}),[debtForm,setDebtForm]=useState({fecha:today(),proveedor:"",emailProveedor:"",categoria:"Compra de equipos",descripcion:"",monto:"",vencimiento:today(),estado:"Pendiente"}),[editingClient,setEditingClient]=useState(null),[editingInvoice,setEditingInvoice]=useState(null);
useEffect(()=>{localStorage.setItem(KEY,JSON.stringify(data));setSaved(new Date().toLocaleTimeString("es-CL"))},[data]);
useEffect(()=>{let t=setInterval(()=>setClock(new Date()),1000);return()=>clearInterval(t)},[]);
useEffect(()=>{setInvoicePage(1)},[invoiceFolderMonth,invoiceStatusFilter,search]);
const client=id=>data.clients.find(c=>+c.id===+id);
const months=useMemo(()=>{let s=new Set([selectedMonth,mk(today())]);data.invoices.forEach(i=>s.add(mk(i.vencimiento)));data.incomes.forEach(i=>s.add(mk(i.fecha)));data.expenses.forEach(e=>s.add(mk(e.fecha)));data.debts.forEach(d=>s.add(mk(d.vencimiento)));return [...s].sort().reverse()},[data,selectedMonth]);
const historyMonths=useMemo(()=>{
 const set=new Set(["Todos"]);
 data.incomes.forEach(i=>set.add(mk(i.fecha)));
 data.expenses.forEach(e=>set.add(mk(e.fecha)));
 data.debts.forEach(d=>set.add(mk(d.vencimiento||d.fecha)));
 return [...set].sort().reverse();
},[data.incomes,data.expenses,data.debts]);
const fm=useMemo(()=>({invoices:data.invoices.filter(i=>mk(i.vencimiento)===selectedMonth),incomes:data.incomes.filter(i=>mk(i.fecha)===selectedMonth),expenses:data.expenses.filter(e=>mk(e.fecha)===selectedMonth),debts:data.debts.filter(d=>mk(d.vencimiento)===selectedMonth)}),[data,selectedMonth]);
const stats=useMemo(()=>{let ingresos=fm.incomes.reduce((s,x)=>s+ +x.monto,0),egresos=fm.expenses.reduce((s,x)=>s+ +x.monto,0),deudas=fm.debts.filter(d=>d.estado!=="Pagada").reduce((s,x)=>s+ +x.monto,0),pend=fm.invoices.filter(i=>ist(i).l!=="Pagada").reduce((s,x)=>s+ +x.monto,0),venc=fm.invoices.filter(i=>ist(i).l==="Vencida");return{ingresos,egresos,deudas,pend,venc,saldo:ingresos-egresos}},[fm]);
const filteredClients=data.clients.filter(c=>`${c.nombre} ${c.rut} ${c.email} ${c.giro}`.toLowerCase().includes(search.toLowerCase()));
const filteredInvoices=data.invoices.filter(i=>`${i.factura} ${client(i.clienteId)?.nombre||""}`.toLowerCase().includes(search.toLowerCase()));
const invoiceMonths=useMemo(()=>{
 const set=new Set(["Todas",mk(today())]);
 data.invoices.forEach(i=>set.add(mk(i.vencimiento||i.emision||today())));
 return [...set].sort().reverse();
},[data.invoices]);
const filteredInvoicesByFolder=useMemo(()=>{
 return data.invoices
  .filter(i=>invoiceFolderMonth==="Todas"||mk(i.vencimiento||i.emision||today())===invoiceFolderMonth)
  .filter(i=>`${i.factura} ${client(i.clienteId)?.nombre||""} ${client(i.clienteId)?.rut||""}`.toLowerCase().includes(search.toLowerCase()))
  .filter(i=>invoiceStatusFilter==="Todas"||ist(i).l===invoiceStatusFilter)
  .sort((a,b)=>String(a.factura||"").localeCompare(String(b.factura||""),undefined,{numeric:true}));
},[data.invoices,invoiceFolderMonth,invoiceStatusFilter,search]);
const invoicePageSize=50;
const invoiceTotalPages=Math.max(1,Math.ceil(filteredInvoicesByFolder.length/invoicePageSize));
const invoicePageSafe=Math.min(invoicePage,invoiceTotalPages);
const paginatedInvoices=filteredInvoicesByFolder.slice((invoicePageSafe-1)*invoicePageSize,invoicePageSafe*invoicePageSize);
const alertInvoices=data.invoices.filter(i=>isVencida15(i)||isPorVencer15(i)).filter(i=>`${i.factura} ${client(i.clienteId)?.nombre||""}`.toLowerCase().includes(alertSearch.toLowerCase())).filter(i=>reminderStatusFilter==="Todas"||ist(i).l===reminderStatusFilter);
const selectedInvoice=selectedInvoiceId?data.invoices.find(i=>i.id===selectedInvoiceId):alertInvoices[0], selectedClient=selectedInvoice?client(selectedInvoice.clienteId):null;
const monthly=months.slice().reverse().map(m=>({mes:ml(m),ingresos:data.incomes.filter(i=>mk(i.fecha)===m).reduce((s,x)=>s+ +x.monto,0),egresos:data.expenses.filter(e=>mk(e.fecha)===m).reduce((s,x)=>s+ +x.monto,0),deudas:data.debts.filter(d=>mk(d.vencimiento)===m&&d.estado!=="Pagada").reduce((s,x)=>s+ +x.monto,0),vencidas:data.invoices.filter(i=>mk(i.vencimiento)===m&&ist(i).l==="Vencida").reduce((s,x)=>s+ +x.monto,0),saldo:data.incomes.filter(i=>mk(i.fecha)===m).reduce((s,x)=>s+ +x.monto,0)-data.expenses.filter(e=>mk(e.fecha)===m).reduce((s,x)=>s+ +x.monto,0)}));
const pie=[{name:"Pagadas",value:fm.invoices.filter(i=>ist(i).l==="Pagada").length},{name:"Pendientes",value:fm.invoices.filter(i=>ist(i).l==="Pendiente").length},{name:"Por vencer",value:fm.invoices.filter(i=>ist(i).l==="Por vencer").length},{name:"Vencidas",value:fm.invoices.filter(i=>ist(i).l==="Vencida").length}];

const aiData=useMemo(()=>{
 const vencidas=data.invoices.filter(i=>isVencida15(i));
 const porVencer=data.invoices.filter(i=>isPorVencer15(i));
 const riesgoCliente=data.clients.map(c=>{
  const inv=data.invoices.filter(i=>+i.clienteId===+c.id);
  const venc=inv.filter(i=>ist(i).l==="Vencida").length;
  const pend=inv.filter(i=>ist(i).l!=="Pagada").length;
  const montoVenc=inv.filter(i=>ist(i).l==="Vencida").reduce((s,x)=>s+ +x.monto,0);
  const premium=venc===0&&pend===0;
  const riesgo=premium?"PREMIUM":venc>=4?"ALTO":venc===3?"MEDIO":venc===2?"BAJO":"AL DÍA";
  return {...c,riesgo,premium,vencidas:venc,pendientes:pend,montoVencido:montoVenc};
 }).sort((a,b)=>b.montoVencido-a.montoVencido);
 const sugerencias=[
  vencidas.length>0?`Enviar recordatorio a ${vencidas.length} factura(s) vencida(s) dentro de los últimos 15 días.`:"No hay facturas vencidas dentro de los últimos 15 días.",
  porVencer.length>0?`Programar aviso preventivo para ${porVencer.length} factura(s) por vencer dentro de 15 días.`:"No hay facturas próximas a vencer dentro de 15 días.",
  riesgoCliente.filter(c=>c.riesgo==="ALTO").length>0?`Revisar clientes de riesgo ALTO: 4 o más facturas vencidas.`:`Clientes PREMIUM al día detectados: ${riesgoCliente.filter(c=>c.riesgo==="PREMIUM").length}`
 ];
 const pendientes2=riesgoCliente.filter(c=>c.pendientes===2);
 const pendientes3=riesgoCliente.filter(c=>c.pendientes===3);
 const pendientes4=riesgoCliente.filter(c=>c.pendientes>=4);
 return {vencidas,porVencer,riesgoCliente,sugerencias,pendientes2,pendientes3,pendientes4};
},[data]);

const radarData=useMemo(()=>{
 const maxBase=Math.max(stats.ingresos,stats.egresos,stats.deudas,stats.pend,1);
 const premiumCount=aiData.riesgoCliente.filter(c=>c.riesgo==="PREMIUM").length;
 const altoCount=aiData.riesgoCliente.filter(c=>c.riesgo==="ALTO").length;
 return [
  {area:"Ingresos",valor:Math.min(100,Math.round((stats.ingresos/maxBase)*100))},
  {area:"Egresos",valor:Math.min(100,Math.round((stats.egresos/maxBase)*100))},
  {area:"Deudas",valor:Math.min(100,Math.round((stats.deudas/maxBase)*100))},
  {area:"Por cobrar",valor:Math.min(100,Math.round((stats.pend/maxBase)*100))},
  {area:"Premium",valor:Math.min(100,premiumCount*20)},
  {area:"Riesgo",valor:Math.min(100,altoCount*25)}
 ];
},[stats,aiData]);

const liveActivity=useMemo(()=>{
 const acts=[];
 data.incomes.slice(0,3).forEach(i=>acts.push({icon:"💰",text:`Ingreso registrado: ${money(i.monto)} · ${i.categoria}`,date:i.fecha,type:"ok"}));
 data.expenses.slice(0,3).forEach(e=>acts.push({icon:"📤",text:`Egreso registrado: ${money(e.monto)} · ${e.categoria}`,date:e.fecha,type:"bad"}));
 aiData.vencidas.slice(0,3).forEach(f=>acts.push({icon:"⚠️",text:`Factura vencida: ${f.factura} · ${client(f.clienteId)?.nombre||""}`,date:f.vencimiento,type:"warn"}));
 aiData.riesgoCliente.filter(c=>c.riesgo==="PREMIUM").slice(0,2).forEach(c=>acts.push({icon:"⭐",text:`Cliente PREMIUM al día: ${c.nombre}`,date:today(),type:"premium"}));
 return acts.sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,8);
},[data,aiData]);


const vencimientosPanel=useMemo(()=>{
 const vencidas=data.invoices.filter(i=>ist(i).l==="Vencida");
 const porVencer15=data.invoices.filter(i=>ist(i).l!=="Pagada"&&days(i.vencimiento)>=0&&days(i.vencimiento)<=15);
 const totalPorCobrar=data.invoices.filter(i=>ist(i).l!=="Pagada").reduce((s,i)=>s+(+i.monto||0),0);
 const clientes=data.clients.map(c=>{
  const inv=data.invoices.filter(i=>+i.clienteId===+c.id);
  const pendientes=inv.filter(i=>ist(i).l!=="Pagada").length;
  const venc=inv.filter(i=>ist(i).l==="Vencida").length;
  const monto=inv.filter(i=>ist(i).l!=="Pagada").reduce((s,i)=>s+(+i.monto||0),0);
  return {...c,pendientes,vencidas:venc,montoPendiente:monto};
 }).filter(c=>c.pendientes>=2).sort((a,b)=>b.montoPendiente-a.montoPendiente);
 return{vencidas,porVencer15,totalPorCobrar,clientes2:clientes.filter(c=>c.pendientes===2),clientes3:clientes.filter(c=>c.pendientes===3),clientes4:clientes.filter(c=>c.pendientes>=4),ranking:clientes.slice(0,8)};
},[data]);

function aiMessage(i){return `Estimado cliente, se recuerda su Factura ${i.factura} por la suma de ${money(i.monto)}. Saludos Cordiales GpsRuta`}

function saveClient(){
if(!clientForm.nombre||!clientForm.rut){alert("Complete nombre/razón social y RUT del cliente.");return;}
const cleanRut=(rut)=>String(rut||"").replace(/\./g,"").replace(/-/g,"").replace(/\s/g,"").toLowerCase();
const duplicatedRut=data.clients.some(c=>cleanRut(c.rut)===cleanRut(clientForm.rut)&&c.id!==editingClient);
if(duplicatedRut){alert("ERROR: Ya existe un cliente registrado con este RUT. No se puede repetir.");return;}if(editingClient){setData({...data,clients:data.clients.map(c=>c.id===editingClient?{...clientForm,id:editingClient}:c)});setEditingClient(null)}else setData({...data,clients:[{...clientForm,id:Date.now()},...data.clients]});setClientForm({nombre:"",rut:"",giro:"",telefono:"569",email:"",direccion:"",contacto:""})}
function editClient(c){setEditingClient(c.id);setClientForm(c);setTab("clientes")}
function saveInvoice(){
if(!invoiceForm.clienteId||!invoiceForm.factura||!invoiceForm.monto){alert("Complete cliente, número de factura y monto.");return;}
const duplicated=data.invoices.some(i=>String(i.factura).trim().toLowerCase()===String(invoiceForm.factura).trim().toLowerCase()&&i.id!==editingInvoice);
if(duplicated&&!confirm("Esta factura ya existe. ¿Desea guardarla igualmente?"))return;
let p={...invoiceForm,id:editingInvoice||Date.now(),clienteId:+invoiceForm.clienteId,monto:+invoiceForm.monto};setData({...data,invoices:editingInvoice?data.invoices.map(i=>i.id===editingInvoice?p:i):[p,...data.invoices]});setEditingInvoice(null);setInvoiceFolderMonth(mk(p.vencimiento||p.emision));setInvoiceForm({clienteId:"",factura:"",emision:today(),vencimiento:today(),monto:"",estado:"Pendiente",detalle:""})}
function editInvoice(i){setEditingInvoice(i.id);setInvoiceForm({...i,clienteId:String(i.clienteId)});setTab("facturas")}
function saveIncome(){if(!incomeForm.categoria||!incomeForm.monto)return;let inv=data.invoices,desc=incomeForm.descripcion;if(incomeForm.facturaId){let f=data.invoices.find(i=>+i.id===+incomeForm.facturaId);inv=data.invoices.map(i=>+i.id===+incomeForm.facturaId?{...i,estado:"Pagada"}:i);if(!desc&&f)desc=`Pago ${f.factura}`}setData({...data,invoices:inv,incomes:[{...incomeForm,id:Date.now(),monto:+incomeForm.monto,descripcion:desc},...data.incomes]});setIncomeForm({fecha:today(),categoria:"Pago de factura",descripcion:"",monto:"",facturaId:""})}
function saveExpense(){
if(!expenseForm.categoria||!expenseForm.monto)return;
let debts=data.debts;
let descripcion=expenseForm.descripcion;
if(expenseForm.debtId){
 const deuda=data.debts.find(d=>+d.id===+expenseForm.debtId);
 debts=data.debts.map(d=>+d.id===+expenseForm.debtId?{...d,estado:"Pagada"}:d);
 if(!descripcion&&deuda)descripcion=`Pago factura/deuda ${deuda.proveedor} - ${deuda.descripcion}`;
}
setData({...data,debts,expenses:[{...expenseForm,id:Date.now(),monto:+expenseForm.monto,descripcion},...data.expenses]});
setExpenseForm({fecha:today(),categoria:"Pago instalador",descripcion:"",monto:"",debtId:"",numeroFacturaPago:""});
}
function saveDebt(){if(!debtForm.proveedor||!debtForm.monto)return;setData({...data,debts:[{...debtForm,estado:"Pendiente",id:Date.now(),monto:+debtForm.monto},...data.debts]});setDebtForm({fecha:today(),proveedor:"",emailProveedor:"",categoria:"Compra de equipos",descripcion:"",monto:"",vencimiento:today(),estado:"Pendiente"})}
function attachFile(id,file){if(!file)return;let r=new FileReader();r.onload=()=>setData({...data,attachments:{...(data.attachments||{}),[id]:{name:file.name,size:file.size,type:file.type,dataUrl:r.result,attachedAt:new Date().toLocaleString("es-CL")}}});r.readAsDataURL(file)}
function canSendInvoice(id){
  if(!data.attachments?.[id]){
    alert("Debe adjuntar la factura antes de usar Enviar factura.");
    return false;
  }
  return true;
}


function manualSave(){
  localStorage.setItem(KEY, JSON.stringify(data));
  setSaved(new Date().toLocaleTimeString("es-CL"));
  alert("Información guardada correctamente.");
}
function exportBackup(){
  const blob = new Blob([JSON.stringify(data,null,2)], {type:"application/json"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "respaldo_gpsruta_"+new Date().toISOString().slice(0,10)+".json";
  a.click();
  URL.revokeObjectURL(url);
}
function importBackup(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try{
      const imported = JSON.parse(reader.result);
      setData({...seed, ...imported});
      alert("Respaldo importado correctamente.");
    }catch(e){
      alert("El archivo no es válido.");
    }
  };
  reader.readAsText(file);
}
function deleteMonthHistory(kind){
  if(!confirm("¿Eliminar todos los registros de "+ml(selectedMonth)+"? Esta acción no se puede deshacer.")) return;
  if(kind==="debts") setData({...data, debts:data.debts.filter(d=>mk(d.vencimiento)!==selectedMonth)});
  if(kind==="incomes") setData({...data, incomes:data.incomes.filter(i=>mk(i.fecha)!==selectedMonth)});
  if(kind==="expenses") setData({...data, expenses:data.expenses.filter(e=>mk(e.fecha)!==selectedMonth)});
  if(kind==="invoices") setData({...data, invoices:data.invoices.filter(i=>mk(i.vencimiento)!==selectedMonth)});
}
function deleteAllHistory(kind){
  if(!confirm("¿Eliminar TODO el historial de esta sección? Esta acción no se puede deshacer.")) return;
  if(kind==="debts") setData({...data, debts:[]});
  if(kind==="incomes") setData({...data, incomes:[]});
  if(kind==="expenses") setData({...data, expenses:[]});
  if(kind==="invoices") setData({...data, invoices:[]});
}

function importClientsExcel(file){
  if(!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try{
      const workbook = XLSX.read(new Uint8Array(e.target.result), {type:"array"});
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {defval:""});
      if(!rows.length){
        alert("El Excel no tiene datos.");
        return;
      }
      const normalize = (row, keys) => {
        for(const k of keys){
          const found = Object.keys(row).find(x => x.toLowerCase().trim() === k.toLowerCase());
          if(found) return String(row[found] || "").trim();
        }
        return "";
      };
      const nuevos = rows.map((r,idx)=>({
        id: Date.now()+idx,
        nombre: normalize(r,["nombre","cliente","razon social","razón social","RAZON SOCIAL","RAZÓN SOCIAL","empresa"]),
        rut: normalize(r,["rut","r.u.t","rut cliente"]),
        giro: normalize(r,["giro","actividad","actividad economica","actividad económica"]),
        telefono: normalize(r,["telefono","teléfono","celular","whatsapp"]),
        email: normalize(r,["email","correo","correo electronico","correo electrónico"]),
        direccion: normalize(r,["direccion","dirección","domicilio"]),
        contacto: normalize(r,["contacto","persona contacto","encargado"])
      })).filter(c=>c.nombre || c.rut);
      if(!nuevos.length){
        alert("No se encontraron clientes válidos. Revisa que el Excel tenga columnas: nombre, rut, giro, telefono, email, direccion, contacto.");
        return;
      }
      setData({...data, clients:[...nuevos, ...data.clients]});
      alert("Clientes cargados correctamente: "+nuevos.length);
    }catch(err){
      alert("No se pudo leer el Excel. Revisa el formato del archivo.");
    }
  };
  reader.readAsArrayBuffer(file);
}

function exportFinancialReport(){
  const clienteNombre = (id) => client(id)?.nombre || "";
  const clienteRut = (id) => client(id)?.rut || "";
  const workbook = XLSX.utils.book_new();

  const resumen = [{
    "Mes/Año": ml(selectedMonth),
    "Ingresos": stats.ingresos,
    "Egresos": stats.egresos,
    "Saldo Neto": stats.saldo,
    "Deudas por pagar": stats.deudas,
    "Facturas pendientes por cobrar": stats.pend,
    "Facturas vencidas": stats.venc.length,
    "Fecha descarga": new Date().toLocaleString("es-CL")
  }];

  const pagos = data.incomes
    .filter(i => i.categoria === "Pago de factura" || i.facturaId)
    .map(i => {
      const inv = data.invoices.find(f => Number(f.id) === Number(i.facturaId));
      return {
        "Fecha Pago": i.fecha,
        "Mes/Año": ml(mk(i.fecha)),
        "Factura": inv?.factura || i.descripcion || "",
        "Cliente": inv ? clienteNombre(inv.clienteId) : "",
        "RUT": inv ? clienteRut(inv.clienteId) : "",
        "Categoría": i.categoria,
        "Descripción": i.descripcion,
        "Monto": Number(i.monto || 0)
      };
    });

  const ingresos = data.incomes.map(i => ({
    "Fecha": i.fecha,
    "Mes/Año": ml(mk(i.fecha)),
    "Categoría": i.categoria,
    "Descripción": i.descripcion,
    "Factura Asociada": data.invoices.find(f => Number(f.id) === Number(i.facturaId))?.factura || "",
    "Monto": Number(i.monto || 0)
  }));

  const egresos = data.expenses.map(e => ({
    "Fecha": e.fecha,
    "Mes/Año": ml(mk(e.fecha)),
    "Categoría": e.categoria,
    "Descripción": e.descripcion,
    "Monto": Number(e.monto || 0)
  }));

  const deudas = data.debts
    .filter(d => d.estado !== "Pagada")
    .map(d => ({
      "Fecha Registro": d.fecha,
      "Proveedor": d.proveedor,
      "Email Proveedor": d.emailProveedor || "",
      "Categoría": d.categoria,
      "Descripción": d.descripcion,
      "Vencimiento": d.vencimiento,
      "Mes/Año": ml(mk(d.vencimiento)),
      "Estado": d.estado,
      "Monto": Number(d.monto || 0)
    }));

  const facturasPorPagar = data.debts.filter(d => d.estado !== "Pagada").map(d => ({
    "Proveedor": d.proveedor,
    "Categoría": d.categoria,
    "Descripción": d.descripcion,
    "Vencimiento": d.vencimiento,
    "Mes/Año": ml(mk(d.vencimiento)),
    "Estado": d.estado,
    "Monto": Number(d.monto || 0)
  }));

  const facturasPorCobrar = data.invoices
    .filter(f => ["Vencida","Por vencer"].includes(ist(f).l))
    .map(f => ({
      "Factura": f.factura,
      "Cliente": clienteNombre(f.clienteId),
      "RUT Cliente": clienteRut(f.clienteId),
      "Emisión": f.emision,
      "Vencimiento": f.vencimiento,
      "Mes/Año": ml(mk(f.vencimiento)),
      "Estado": ist(f).l,
      "Detalle": f.detalle,
      "Monto": Number(f.monto || 0)
    }));

  const clientes = data.clients.map(c => ({
    "Nombre / Razón Social": c.nombre,
    "RUT": c.rut,
    "Giro": c.giro,
    "Teléfono": c.telefono,
    "Email": c.email,
    "Dirección": c.direccion,
    "Contacto": c.contacto
  }));

  const sheets = [
    ["Resumen", resumen],
    ["Pagos", pagos],
    ["Ingresos", ingresos],
    ["Egresos", egresos],
    ["Deudas", deudas],
    ["Facturas por pagar", facturasPorPagar],
    ["Facturas por cobrar", facturasPorCobrar],
    ["Clientes", clientes]
  ];

  sheets.forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}]);
    XLSX.utils.book_append_sheet(workbook, ws, name);
  });

  XLSX.writeFile(workbook, "informe_gpsruta_"+selectedMonth+".xlsx");
}

function importInvoicesExcel(file){
 if(!file)return;
 const reader=new FileReader();
 reader.onload=(e)=>{
  try{
   const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
   const sh=wb.Sheets[wb.SheetNames[0]];
   const rows=XLSX.utils.sheet_to_json(sh,{defval:""});
   const norm=(row,keys)=>{for(const k of keys){const f=Object.keys(row).find(x=>x.toLowerCase().trim()===k.toLowerCase());if(f)return String(row[f]||"").trim()}return ""};
   const findClientId=(rut,nombre)=>{
    const rc=String(rut||"").replace(/\./g,"").replace(/-/g,"").toLowerCase();
    let c=data.clients.find(cli=>String(cli.rut||"").replace(/\./g,"").replace(/-/g,"").toLowerCase()===rc);
    if(c)return c.id;
    c=data.clients.find(cli=>String(cli.nombre||"").toLowerCase().trim()===String(nombre||"").toLowerCase().trim());
    return c?.id||"";
   };
   const nuevas=rows.map((r,idx)=>{
    const rut=norm(r,["rut cliente","rut","RUT","RUT CLIENTE"]);
    const nombre=norm(r,["cliente","nombre cliente","razon social","razón social","RAZON SOCIAL"]);
    const monto=Number(String(norm(r,["monto","valor","total"])).replace(/\$/g,"").replace(/\./g,"").replace(/,/g,"."))||0;
    return {id:Date.now()+idx,clienteId:findClientId(rut,nombre),factura:norm(r,["factura","n factura","n° factura","numero factura","número factura","folio"]),emision:norm(r,["emision","emisión","fecha emision","fecha emisión"])||today(),vencimiento:norm(r,["vencimiento","fecha vencimiento","vence"])||today(),monto,estado:norm(r,["estado"])||"Pendiente",detalle:norm(r,["detalle","descripcion","descripción","glosa"])||"Factura cargada masivamente"};
   }).filter(f=>f.factura&&f.clienteId&&f.monto);
   if(!nuevas.length){alert("No se encontraron facturas válidas. Revisa columnas: factura, rut cliente o cliente, vencimiento, monto.");return}
   setData({...data,invoices:[...nuevas,...data.invoices]});
   setInvoiceFolderMonth(mk(nuevas[0].vencimiento||nuevas[0].emision));
   alert("Facturas cargadas correctamente: "+nuevas.length+" en carpeta "+ml(mk(nuevas[0].vencimiento||nuevas[0].emision)));
  }catch(err){alert("No se pudo leer el Excel. Revisa el formato del archivo.")}
 };
 reader.readAsArrayBuffer(file);
}

function askDashboardAI(){
  const q=chatInput.trim();
  if(!q)return;
  const lower=q.toLowerCase();
  let answer="";
  const vencidas=data.invoices.filter(i=>isVencida15(i));
  const porVencer=data.invoices.filter(i=>isPorVencer15(i));
  const deudasPendientes=data.debts.filter(d=>d.estado!=="Pagada");
  const topMorosos=aiData.riesgoCliente.filter(c=>c.vencidas>0).slice(0,5);
  const facturasPendientes=data.invoices
    .filter(i=>ist(i).l!=="Pagada")
    .sort((a,b)=>new Date(a.vencimiento||a.emision)-new Date(b.vencimiento||b.emision));
  const diasVencida=(i)=>Math.max(0,Math.abs(days(i.vencimiento)));

  if(lower.includes("factura más antigua")||lower.includes("factura mas antigua")){
    const f=facturasPendientes[0];
    answer=f?`La factura más antigua pendiente es:\n• Factura: ${f.factura}\n• Cliente: ${client(f.clienteId)?.nombre||""}\n• Emisión: ${f.emision}\n• Vencimiento: ${f.vencimiento}\n• Monto: ${money(f.monto)}\n• Antigüedad: ${diasVencida(f)} día(s) desde vencimiento.`:"No hay facturas pendientes.";
  }
  else if(lower.includes("10 facturas más antiguas")||lower.includes("10 facturas mas antiguas")||lower.includes("facturas más antiguas")||lower.includes("facturas mas antiguas")){
    answer=facturasPendientes.length?"Las facturas pendientes más antiguas son:\n"+facturasPendientes.slice(0,10).map((f,idx)=>`${idx+1}. ${f.factura} · ${client(f.clienteId)?.nombre||""} · vence ${f.vencimiento} · ${diasVencida(f)} día(s) · ${money(f.monto)}`).join("\n"):"No hay facturas pendientes.";
  }
  else if(lower.includes("deuda más antigua")||lower.includes("deuda mas antigua")||lower.includes("cliente tiene la deuda más antigua")||lower.includes("cliente tiene la deuda mas antigua")){
    const f=facturasPendientes[0];
    answer=f?`El cliente con deuda más antigua es:\n• Cliente: ${client(f.clienteId)?.nombre||""}\n• Factura: ${f.factura}\n• Vencimiento: ${f.vencimiento}\n• Días vencida: ${diasVencida(f)}\n• Monto: ${money(f.monto)}`:"No hay deuda pendiente registrada.";
  }
  else if(lower.includes("90 días")||lower.includes("90 dias")||lower.includes("más de 90")||lower.includes("mas de 90")){
    const mayores90=facturasPendientes.filter(f=>diasVencida(f)>90);
    const total=mayores90.reduce((s,f)=>s+(+f.monto||0),0);
    answer=`Facturas con más de 90 días:\n• Cantidad: ${mayores90.length}\n• Total: ${money(total)}${mayores90.length?"\n\nDetalle:\n"+mayores90.slice(0,10).map(f=>`• ${f.factura} · ${client(f.clienteId)?.nombre||""} · ${diasVencida(f)} día(s) · ${money(f.monto)}`).join("\n"):""}`;
  }
  else if(lower.includes("cobrar primero")||lower.includes("prioridad de cobro")||lower.includes("prioridad cobranza")){
    answer=facturasPendientes.length?"Prioridad de cobranza recomendada:\n"+facturasPendientes.slice(0,10).map((f,idx)=>`${idx+1}. ${f.factura} · ${client(f.clienteId)?.nombre||""} · ${diasVencida(f)} día(s) · ${money(f.monto)}`).join("\n"):"No hay facturas pendientes para priorizar.";
  }
  else if(lower.includes("pendiente")||lower.includes("pendientes")){ const lista=aiData.riesgoCliente.filter(c=>c.pendientes>=2).slice(0,10); answer=lista.length?"Clientes con 2 o más facturas pendientes:\n"+lista.map(c=>`• ${c.nombre}: ${c.pendientes} pendiente(s), riesgo ${c.riesgo}`).join("\n"):"No hay clientes con 2 o más facturas pendientes."; }
  else if(lower.includes("moroso")||lower.includes("riesgo")||lower.includes("debe más")||lower.includes("deuda cliente")){
    answer=topMorosos.length
      ? "Clientes con mayor riesgo:\\n"+topMorosos.map(c=>`• ${c.nombre}: ${c.vencidas} vencida(s), ${money(c.montoVencido)}, riesgo ${c.riesgo}`).join("\\n")
      : "No detecto clientes morosos en este momento.";
  }else if(lower.includes("vencida")||lower.includes("vencidas")){
    answer=vencidas.length
      ? "Vencidas últimos 15 días:\\n"+vencidas.slice(0,8).map(i=>`• ${i.factura} · ${client(i.clienteId)?.nombre||""} · ${money(i.monto)} · venció ${i.vencimiento}`).join("\\n")
      : "No hay facturas vencidas.";
  }else if(lower.includes("por vencer")||lower.includes("vencer")){
    answer=porVencer.length
      ? "Por vencer próximos 15 días:\\n"+porVencer.slice(0,8).map(i=>`• ${i.factura} · ${client(i.clienteId)?.nombre||""} · ${money(i.monto)} · vence ${i.vencimiento}`).join("\\n")
      : "No hay facturas por vencer.";
  }else if(lower.includes("ingreso")||lower.includes("ingresos")){
    answer=`Ingresos del mes ${ml(selectedMonth)}: ${money(stats.ingresos)}. Registros: ${fm.incomes.length}.`;
  }else if(lower.includes("egreso")||lower.includes("egresos")){
    answer=`Egresos del mes ${ml(selectedMonth)}: ${money(stats.egresos)}. Registros: ${fm.expenses.length}.`;
  }else if(lower.includes("deuda")||lower.includes("deudas")){
    answer=deudasPendientes.length
      ? "Deudas/facturas por pagar pendientes:\\n"+deudasPendientes.slice(0,8).map(d=>`• ${d.proveedor}: ${money(d.monto)} · vence ${d.vencimiento}`).join("\\n")
      : "No hay deudas pendientes por pagar.";
  }else if(lower.includes("premium")){
    const premium=aiData.riesgoCliente.filter(c=>c.riesgo==="PREMIUM");
    answer=premium.length
      ? "Clientes PREMIUM al día:\\n"+premium.slice(0,8).map(c=>`• ${c.nombre} · RUT ${c.rut}`).join("\\n")
      : "Aún no hay clientes PREMIUM detectados.";
  }else{
    answer=`Resumen ${ml(selectedMonth)}:\\n• Ingresos: ${money(stats.ingresos)}\\n• Egresos: ${money(stats.egresos)}\\n• Saldo: ${money(stats.saldo)}\\n• Deudas por pagar: ${money(stats.deudas)}\\n• Facturas pendientes: ${money(stats.pend)}\\n• Facturas vencidas: ${stats.venc.length}`;
  }
  setChatMessages([...chatMessages,{role:"user",text:q},{role:"ia",text:answer}]);
  setChatInput("");
}

async function sendAutomaticEmail(payload){
  try{
    setEmailSending(true);
    const resp = await fetch("/api/send-email",{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify(payload)
    });
    const result = await resp.json();
    if(!resp.ok || !result.ok) throw new Error(result.error || "No se pudo enviar el correo.");
    alert("Correo enviado correctamente desde Outlook.");
  }catch(err){
    alert("Error al enviar correo automático: "+(err.message || err));
  }finally{
    setEmailSending(false);
  }
}
function sendAutoReminder(i,c){
  if(!c?.email){alert("El cliente no tiene correo registrado.");return;}
  sendAutomaticEmail({
    to:c.email,
    subject:`Recordatorio factura ${i.factura}`,
    body:reminderText(i),
    attachments:[]
  });
}
function sendAutoInvoice(i,c){
  if(!c?.email){alert("El cliente no tiene correo registrado.");return;}
  const attached=data.attachments?.[i.id];
  if(!attached){alert("Debe adjuntar la factura antes de enviar factura automática.");return;}
  sendAutomaticEmail({
    to:c.email,
    subject:`Envío factura ${i.factura}`,
    body:invoiceSendText(i),
    attachments:[attached]
  });
}
function sendAutoDebt(d){
  if(!d?.emailProveedor){alert("La deuda/proveedor no tiene correo registrado.");return;}
  sendAutomaticEmail({
    to:d.emailProveedor,
    subject:`Consulta / pago de factura pendiente - ${d.proveedor}`,
    body:`Estimados,\n\nJunto con saludar, informamos que tenemos registrada una factura/deuda pendiente por pagar.\n\nProveedor: ${d.proveedor}\nDetalle: ${d.descripcion}\nMonto: ${money(d.monto)}\nFecha de vencimiento: ${d.vencimiento}\n\nFavor confirmar datos de pago o estado de la factura.\n\nSaludos cordiales.\nGPSRUTA.`,
    attachments:[]
  });
}

function addTask(){
 if(!taskText.trim())return;
 setData({...data,tasks:[{id:Date.now(),text:taskText.trim(),done:false,createdAt:new Date().toLocaleString("es-CL",{hour12:false})},...(data.tasks||[])]});
 setTaskText("");
}
function toggleTask(id){setData({...data,tasks:(data.tasks||[]).map(t=>t.id===id?{...t,done:!t.done}:t)})}
function deleteTask(id){setData({...data,tasks:(data.tasks||[]).filter(t=>t.id!==id)})}
if(!logged)return <Login onLogin={()=>setLogged(true)}/>;
return <div className="app"><aside><Logo/><div className="admin"><User size={24}/><div><b>Administrador</b><p>admin@gpsruta.cl</p></div></div><nav>{[["dashboard","Dashboard",Eye],["clientes","Clientes",Users],["facturas","Facturas por cobrar",FileText],["deudas","Deudas / Facturas por pagar",CreditCard],["ingresos","Ingresos",TrendingUp],["egresos","Egresos",TrendingDown],["alertas","Cobros / Recordatorios",Bell]].map(([v,l,I])=><button key={v} onClick={()=>setTab(v)} className={tab===v?"active":""}><I size={20}/>{l}</button>)}</nav><div className="autosave"><CheckCircle size={20}/><div><b>Guardado automático activo</b><p>Último guardado: {saved}</p></div></div><button className="logout" onClick={()=>{sessionStorage.removeItem(SESSION);setLogged(false)}}><LogOut size={19}/>Cerrar sesión</button></aside><main><header><div className="search"><Search size={17}/><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar cliente, factura o giro..."/></div><div className="chips"><select value={selectedMonth} onChange={e=>setSelectedMonth(e.target.value)} className="monthSelect">{months.map(m=><option key={m} value={m}>{ml(m)}</option>)}</select><span><CalendarDays size={17}/>{clock.toLocaleDateString("es-CL")}</span><span><Clock size={17}/>{clock.toLocaleTimeString("es-CL",{hour12:false,hour:"2-digit",minute:"2-digit",second:"2-digit"})}</span><span className="green"><Save size={17}/>Guardado automático</span></div></header>
<section className="backupToolbar">
  <button className="backupBtn saveManual" onClick={manualSave}><HardDrive size={17}/>Guardar ahora</button>
  <button className="backupBtn" onClick={exportBackup}><Download size={17}/>Respaldar</button>
  <label className="backupBtn importBtn"><Upload size={17}/>Importar respaldo<input type="file" accept=".json" onChange={e=>importBackup(e.target.files?.[0])}/></label><button className="backupBtn reportBtn" onClick={exportFinancialReport}>📊 Descargar informe Excel</button><span className="emailConfigured">📧 Correo cobranza: {SENDER_EMAIL}</span><span className={`emailStatus ${emailSending?"sending":""}`}>{emailSending?"📤 Enviando correo...":"✅ Outlook automático listo"}</span>
</section>
<section className="kpis"><K t="Ingresos del mes" v={money(stats.ingresos)} s={ml(selectedMonth)} icon={TrendingUp}/><K t="Egresos del mes" v={money(stats.egresos)} s={ml(selectedMonth)} icon={TrendingDown} tone="red"/><K t="Deudas por pagar" v={money(stats.deudas)} s="Según vencimiento" icon={CreditCard} tone="gold"/><K t="Facturas pendientes" v={money(stats.pend)} s="Por cobrar" icon={FileText} tone="blue"/><K t="Facturas vencidas" v={stats.venc.length} s="Cantidad mensual" icon={AlertTriangle} tone="red"/></section>
{tab==="dashboard"&&<section className="gridDash dashboard3d"><div className="card chartPro wide holoCard mainHolo"><div className="holoBadge">HOLOGRAPHIC FINANCE PANEL</div><h2>Resumen financiero mensual 3D</h2><div className="chart compact hologramChart"><ResponsiveContainer><ComposedChart data={monthly} margin={{top:12,right:18,left:0,bottom:0}}><CartesianGrid stroke="rgba(255,255,255,.07)" vertical={false}/><XAxis dataKey="mes" stroke="#ccc" tick={{fontSize:12}}/><YAxis stroke="#ccc" tickFormatter={v=>`${Math.round(v/1000000)}M`} tick={{fontSize:12}}/><Tooltip formatter={v=>money(v)} contentStyle={{background:"#050505",border:"1px solid #FFD43B",borderRadius:"12px"}}/><Legend wrapperStyle={{fontSize:12}}/><Bar dataKey="ingresos" name="Ingresos" fill="#7CFC00" radius={[8,8,0,0]} barSize={20}/><Bar dataKey="egresos" name="Egresos" fill="#ff3131" radius={[8,8,0,0]} barSize={20}/><Bar dataKey="deudas" name="Deudas" fill="#FFD43B" radius={[8,8,0,0]} barSize={20}/><Line type="monotone" dataKey="saldo" name="Saldo neto" stroke="#00a3ff" strokeWidth={4} dot={{r:4,fill:"#00a3ff"}} activeDot={{r:7}}/></ComposedChart></ResponsiveContainer></div></div><div className="card chartPro estadoFacturasCard holoCard sideHolo"><div className="holoBadge small">STATUS SCAN</div><h2>Estado facturas del mes</h2><div className="chart donut hologramDonut"><ResponsiveContainer><PieChart><Pie data={pie} dataKey="value" nameKey="name" innerRadius={42} outerRadius={72} label={({name,value})=>`${name}: ${value}`} label={({value})=>value}>{pie.map((_,i)=><Cell key={i} fill={["#7CFC00","#FFD43B","#ff9500","#ff3131"][i]}/>)}</Pie><Tooltip contentStyle={{background:"#050505",border:"1px solid #00a3ff"}}/></PieChart></ResponsiveContainer></div><div className="estadoFacturasNumeros">{pie.map((p,i)=><div className={`estadoItem estado${i}`} key={p.name}><span></span><b>{p.name}</b><strong>{p.value}</strong></div>)}</div></div><div className="card wide holoCard summaryHolo"><div className="holoBadge small">MONTHLY CORE</div><h2>Resumen del mes seleccionado</h2><div className="summaryGrid">{[[money(stats.ingresos),"Ingresos",""],[money(stats.egresos),"Egresos",""],[money(stats.saldo),"Saldo neto",stats.saldo<0?"negativeBalance":""],[money(stats.deudas),"Deudas por pagar",""],[money(stats.pend),"Facturas por cobrar",""],[stats.venc.length,"Facturas vencidas",""]].map(([a,b,cls])=><div key={b} className={cls}><b>{a}</b><span>{b}</span></div>)}</div></div></section>}
<div className="card wide holoCard aiPanel"><div className="holoBadge small">LUXURY GPSRUTA</div><h2>LUXURY GPSRUTA</h2><div className="aiGrid"><div><b>{aiData.vencidas.length}</b><span>Vencidas últimos 15 días</span></div><div><b>{aiData.porVencer.length}</b><span>Por vencer próximos 15 días</span></div><div><b>{aiData.riesgoCliente.filter(c=>c.riesgo==="ALTO").length}</b><span>Clientes riesgo alto</span></div><div><b>{aiData.riesgoCliente.filter(c=>c.riesgo==="PREMIUM").length}</b><span>Clientes PREMIUM al día</span></div><div><b>{aiData.pendientes2?.length||0}</b><span>Clientes con 2 pendientes</span></div><div><b>{aiData.pendientes3?.length||0}</b><span>Clientes con 3 pendientes</span></div><div><b>{aiData.pendientes4?.length||0}</b><span>Clientes con 4+ pendientes</span></div></div><div className="aiColumns"><div><h3>Sugerencias automáticas</h3>{aiData.sugerencias.map((s,i)=><p key={i}>🤖 {s}</p>)}</div><div><h3>Clientes por riesgo</h3>{aiData.riesgoCliente.slice(0,5).map(c=><p key={c.id}><b className={`risk ${String(c.riesgo).toLowerCase().replace(" ","-")}`}>{c.riesgo}</b> {c.nombre} · {money(c.montoVencido)}</p>)}<h3>Clientes con pendientes</h3>{aiData.riesgoCliente.filter(c=>c.pendientes>=2).slice(0,8).map(c=><p key={`p-${c.id}`}><b className="risk medio">{c.pendientes}</b> {c.nombre} · facturas pendientes</p>)}</div></div>{aiData.vencidas[0]&&<div className="aiMessage"><b>Mensaje sugerido:</b><p>{aiMessage(aiData.vencidas[0])}</p></div>}</div>
<div className={`aiChatWidget ${chatOpen?"open":"closed"}`}>
  <button className="aiChatToggle" onClick={()=>setChatOpen(!chatOpen)}><Bot size={20}/>{chatOpen?"Cerrar IA":"Abrir IA"}</button>
  {chatOpen&&<div className="aiChatBox">
    <div className="aiChatHeader"><Sparkles size={18}/><div><b>LUXURY GPSRUTA</b><small>Inteligencia Artificial Financiera PRO</small></div></div>
    <div className="aiChatMessages">
      {chatMessages.map((m,i)=><div key={i} className={`aiMsg ${m.role}`}><pre>{m.text}</pre></div>)}
    </div>
    <div className="aiChatInput">
      <input value={chatInput} onChange={e=>setChatInput(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")askDashboardAI()}} placeholder="Ej: ¿cuál es la factura más antigua? / ¿quién debe más?"/>
      <button onClick={askDashboardAI}><MessageCircle size={17}/>Enviar</button>
    </div>
  </div>}
</div>

{tab==="dashboard"&&<section className="controlCenterPanel">
  <div className="card holoCard radarPanel">
    <div className="holoBadge small">RADAR IA</div>
    <h2>Radar financiero inteligente</h2>
    <div className="radarBox">
      <ResponsiveContainer width="100%" height="100%">
        <RadarChart data={radarData}>
          <PolarGrid stroke="rgba(0,255,210,.22)"/>
          <PolarAngleAxis dataKey="area" tick={{fill:"#eaffff",fontSize:11}}/>
          <PolarRadiusAxis angle={30} domain={[0,100]} tick={false} axisLine={false}/>
          <Radar name="GPSRUTA" dataKey="valor" stroke="#00ffd2" fill="#00ffd2" fillOpacity={0.28} strokeWidth={3}/>
          <Tooltip contentStyle={{background:"#050505",border:"1px solid #00ffd2",borderRadius:"12px"}}/>
        </RadarChart>
      </ResponsiveContainer>
    </div>
  </div>
  <div className="card holoCard alertCenter">
    <div className="holoBadge small">CENTRO DE ALERTAS</div>
    <h2>Alertas inteligentes</h2>
    <div className="alertGridPro">
      <div className="alertBox red"><b>{aiData.vencidas.length}</b><span>Vencidas 15 días</span></div>
      <div className="alertBox yellow"><b>{aiData.porVencer.length}</b><span>Por vencer 15 días</span></div>
      <div className="alertBox green"><b>{aiData.riesgoCliente.filter(c=>c.riesgo==="PREMIUM").length}</b><span>Premium</span></div>
      <div className="alertBox blue"><b>{aiData.riesgoCliente.filter(c=>c.riesgo==="ALTO").length}</b><span>Riesgo alto</span></div>
    </div>
  </div>
  <div className="card holoCard livePanel">
    <div className="holoBadge small">LIVE FEED</div>
    <h2>Actividad en tiempo real</h2>
    <div className="liveFeed">
      {liveActivity.length?liveActivity.map((a,i)=><div className={`liveItem ${a.type}`} key={i}><strong>{a.icon}</strong><div><b>{a.text}</b><span>{a.date}</span></div></div>):<p>No hay actividad reciente.</p>}
    </div>
  </div>
</section>}

{tab==="dashboard"&&<section className="panelVencimientosPro">
  <div className="card holoCard vencimientosMain">
    <div className="holoBadge small">CONTROL VENCIMIENTOS</div>
    <h2>Panel de Control de Vencimientos</h2>
    <div className="vencimientoKpis">
      <div className="vKpi red"><b>{vencimientosPanel.vencidas.length}</b><span>Facturas vencidas</span></div>
      <div className="vKpi yellow"><b>{vencimientosPanel.porVencer15.length}</b><span>Por vencer 15 días</span></div>
      <div className="vKpi green"><b>{money(vencimientosPanel.totalPorCobrar)}</b><span>Total por cobrar</span></div>
      <div className="vKpi blue"><b>{vencimientosPanel.clientes4.length}</b><span>Clientes 4+ pendientes</span></div>
    </div>
    <div className="semaforoPanel">
      <div><span className="dot red"></span><b>Alto:</b> 4+ facturas pendientes</div>
      <div><span className="dot yellow"></span><b>Medio:</b> 3 facturas pendientes</div>
      <div><span className="dot green"></span><b>Bajo:</b> 2 facturas pendientes</div>
    </div>
  </div>
  <div className="card holoCard rankingClientesCriticos">
    <div className="holoBadge small">RANKING CRÍTICO</div>
    <h2>Clientes más críticos</h2>
    <div className="rankingList">
      {vencimientosPanel.ranking.length?vencimientosPanel.ranking.map(c=><div className="rankingItem" key={c.id}><div><b>{c.nombre}</b><span>{c.pendientes} pendientes · {c.vencidas} vencidas</span></div><strong>{money(c.montoPendiente)}</strong><div className="rankingActions"><a href={`https://wa.me/${c.telefono||""}?text=${encodeURIComponent("Estimado cliente, se recuerda que mantiene facturas pendientes por pagar. Saludos Cordiales GpsRuta")}`} target="_blank"><W/></a><a href={`mailto:${c.email||""}?subject=${encodeURIComponent("Recordatorio de facturas pendientes")}&body=${encodeURIComponent("Estimado cliente, se recuerda que mantiene facturas pendientes por pagar. Saludos Cordiales GpsRuta")}`}><Mail size={15}/></a></div></div>):<p>No hay clientes críticos.</p>}
    </div>
  </div>
  <div className="card holoCard tareasPanel">
    <div className="holoBadge small">TAREAS</div>
    <h2>Block de notas / tareas por hacer</h2>
    <div className="taskInput"><input value={taskText} onChange={e=>setTaskText(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")addTask()}} placeholder="Agregar tarea o nota..."/><button onClick={addTask}>Agregar</button></div>
    <div className="taskList">{(data.tasks||[]).length?(data.tasks||[]).slice(0,12).map(t=><div className={`taskItem ${t.done?"done":""}`} key={t.id}><button onClick={()=>toggleTask(t.id)}>{t.done?"✅":"⬜"}</button><div><b>{t.text}</b><span>{t.createdAt}</span></div><button className="deleteTask" onClick={()=>deleteTask(t.id)}>✕</button></div>):<p>No hay tareas pendientes.</p>}</div>
  </div>
</section>}
{tab==="clientes"&&<section className="two"><div className="card clientFormSticky"><h2>{editingClient?"Editar cliente":"Nuevo cliente"}</h2>
<div className="excelImportBox">
  <label className="excelBtn">📥 Cargar clientes desde Excel
    <input type="file" accept=".xlsx,.xls,.csv" onChange={e=>importClientsExcel(e.target.files?.[0])}/>
  </label>
  <small>Columnas aceptadas: nombre, RAZON SOCIAL, rut, giro, telefono, email, direccion, contacto.</small>
</div>
<Fields obj={clientForm} set={setClientForm} fields={["nombre","rut","giro","telefono","email","direccion","contacto"]}/><button className="primary" onClick={saveClient}><Plus size={17}/>Guardar cliente</button></div><div className="cards">{filteredClients.map(c=><div className="card client compactClient" key={c.id}><div className="clientMiniInfo"><b>{c.nombre}</b><span>RUT: {c.rut}</span></div><div className="actions clientMiniActions"><button className="icon edit" onClick={()=>editClient(c)}><Edit size={16}/></button><button className="icon trash" onClick={()=>setData({...data,clients:data.clients.filter(x=>x.id!==c.id),invoices:data.invoices.filter(i=>+i.clienteId!==+c.id)})}><Trash2 size={16}/></button></div></div>)}</div></section>}
{tab==="facturas"&&<section className="two"><div className="card invoiceFormSticky"><h2>{editingInvoice?"Editar factura":"Nueva factura por cobrar"}</h2>
<div className="excelImportBox"><label className="excelBtn">📥 Cargar facturas desde Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={e=>importInvoicesExcel(e.target.files?.[0])}/></label><small>Columnas: factura, rut cliente o cliente, emision, vencimiento, monto, estado, detalle.</small></div><select value={invoiceForm.clienteId} onChange={e=>setInvoiceForm({...invoiceForm,clienteId:e.target.value})}><option value="">Seleccionar cliente</option>{data.clients.map(c=><option key={c.id} value={c.id}>{c.nombre}</option>)}</select><Fields obj={invoiceForm} set={setInvoiceForm} fields={["factura","emision","vencimiento","monto","detalle"]}/><select value={invoiceForm.estado} onChange={e=>setInvoiceForm({...invoiceForm,estado:e.target.value})}><option>Pendiente</option><option>Vencida</option><option>Pagada</option></select><button className="primary" onClick={saveInvoice}><Plus size={17}/>Guardar factura</button></div><div className="card invoiceListCard">
<div className="historyHead"><h2>Listado de facturas por cobrar</h2><small className="requiredAttach">Organizado por carpeta mensual. Recordatorio no requiere adjunto.</small></div>
<div className="invoiceFolderBar">
  <div><label>Carpeta mes/año</label><select value={invoiceFolderMonth} onChange={e=>setInvoiceFolderMonth(e.target.value)}>{invoiceMonths.map(m=><option key={m} value={m}>{m==="Todas"?"TODAS":ml(m)}</option>)}</select></div>
  <div><label>Estado</label><select value={invoiceStatusFilter} onChange={e=>setInvoiceStatusFilter(e.target.value)}><option>Vencida</option><option>Por vencer</option><option>Todas</option></select></div>
  <div className="folderStats"><b>{filteredInvoicesByFolder.length}</b><span>facturas en carpeta</span></div>
</div>
<InvTable items={paginatedInvoices} client={client} edit={editInvoice} del={id=>setData({...data,invoices:data.invoices.filter(i=>i.id!==id)})} data={data} attachFile={attachFile} canSendInvoice={canSendInvoice}/>
<div className="paginationBar">
  <button onClick={()=>setInvoicePage(Math.max(1,invoicePageSafe-1))} disabled={invoicePageSafe<=1}>⬅ Anterior</button>
  <span>Mostrando {(invoicePageSafe-1)*invoicePageSize+1} - {Math.min(invoicePageSafe*invoicePageSize,filteredInvoicesByFolder.length)} de {filteredInvoicesByFolder.length} · Página {invoicePageSafe} / {invoiceTotalPages}</span>
  <button onClick={()=>setInvoicePage(Math.min(invoiceTotalPages,invoicePageSafe+1))} disabled={invoicePageSafe>=invoiceTotalPages}>Siguiente ➡</button>
</div>
</div></section>}
{tab==="deudas"&&<section className="two"><div className="card historyFormSticky"><h2>Deudas / Facturas por pagar</h2><Fields obj={debtForm} set={setDebtForm} fields={["fecha","proveedor","emailProveedor","descripcion","monto","vencimiento"]}/><select value={debtForm.categoria} onChange={e=>setDebtForm({...debtForm,categoria:e.target.value})}>{expenseCats.map(c=><option key={c}>{c}</option>)}</select><select value={debtForm.estado} onChange={e=>setDebtForm({...debtForm,estado:e.target.value})}><option>Pendiente</option></select><button className="primary gold" onClick={saveDebt}><Plus size={17}/>Guardar deuda</button></div><div className="card compactHistoryCard"><div className="historyHead"><h2>Historial de deudas por mes/año</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("debts")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("debts")}>Eliminar todo</button></div></div><HistoryFilters historyQuickFilter={historyQuickFilter} setHistoryQuickFilter={setHistoryQuickFilter} historyMonthFilter={historyMonthFilter} setHistoryMonthFilter={setHistoryMonthFilter} historyStatusFilter={historyStatusFilter} setHistoryStatusFilter={setHistoryStatusFilter} historyMonths={historyMonths} showStatus={true}/><Table rows={data.debts} type="debts" setData={setData} data={data} quickFilter={historyQuickFilter} monthFilter={historyMonthFilter} statusFilter={historyStatusFilter}/></div></section>}
{tab==="ingresos"&&<section className="two"><div className="card historyFormSticky"><h2>Ingresar ingreso</h2><Fields obj={incomeForm} set={setIncomeForm} fields={["fecha","descripcion","monto"]}/><select value={incomeForm.categoria} onChange={e=>setIncomeForm({...incomeForm,categoria:e.target.value})}>{incomeCats.map(c=><option key={c}>{c}</option>)}</select><select value={incomeForm.facturaId} onChange={e=>{let inv=data.invoices.find(i=>+i.id===+e.target.value);setIncomeForm({...incomeForm,facturaId:e.target.value,monto:inv?inv.monto:incomeForm.monto,descripcion:inv?`Pago ${inv.factura}`:incomeForm.descripcion})}}><option value="">Sin asociar factura</option>{data.invoices.filter(i=>ist(i).l!=="Pagada").map(i=><option key={i.id} value={i.id}>{i.factura} · {client(i.clienteId)?.nombre} · {money(i.monto)}</option>)}</select><button className="primary" onClick={saveIncome}><Plus size={17}/>Guardar ingreso</button></div><div className="card compactHistoryCard"><div className="historyHead"><h2>Historial de ingresos por fecha</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("incomes")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("incomes")}>Eliminar todo</button></div></div><HistoryFilters historyQuickFilter={historyQuickFilter} setHistoryQuickFilter={setHistoryQuickFilter} historyMonthFilter={historyMonthFilter} setHistoryMonthFilter={setHistoryMonthFilter} historyStatusFilter={historyStatusFilter} setHistoryStatusFilter={setHistoryStatusFilter} historyMonths={historyMonths}/><Table rows={data.incomes} type="incomes" setData={setData} data={data} quickFilter={historyQuickFilter} monthFilter={historyMonthFilter}/></div></section>}
{tab==="egresos"&&<section className="two"><div className="card historyFormSticky"><h2>Ingresar egreso</h2><Fields obj={expenseForm} set={setExpenseForm} fields={["fecha","descripcion","monto","numeroFacturaPago"]}/><select value={expenseForm.categoria} onChange={e=>setExpenseForm({...expenseForm,categoria:e.target.value})}>{expenseCats.map(c=><option key={c}>{c}</option>)}</select>
<select value={expenseForm.debtId} onChange={e=>{let d=data.debts.find(x=>+x.id===+e.target.value);setExpenseForm({...expenseForm,debtId:e.target.value,monto:d?d.monto:expenseForm.monto,descripcion:d?`Pago factura/deuda ${d.proveedor} - ${d.descripcion}`:expenseForm.descripcion,categoria:d?d.categoria:expenseForm.categoria})}}>
<option value="">Sin vincular factura por pagar</option>
{data.debts.filter(d=>d.estado!=="Pagada").map(d=><option key={d.id} value={d.id}>{d.proveedor} · {d.descripcion} · {money(d.monto)} · vence {d.vencimiento}</option>)}
</select>
<small className="hint">Al vincular una factura/deuda por pagar y guardar el egreso, quedará marcada como Pagada.</small>
<button className="primary gold" onClick={saveExpense}><Plus size={17}/>Guardar egreso</button></div><div className="card compactHistoryCard"><div className="historyHead"><h2>Historial de egresos por fecha</h2><div><button className="smallDanger" onClick={()=>deleteMonthHistory("expenses")}>Eliminar mes/año</button><button className="smallDanger ghost" onClick={()=>deleteAllHistory("expenses")}>Eliminar todo</button></div></div><HistoryFilters historyQuickFilter={historyQuickFilter} setHistoryQuickFilter={setHistoryQuickFilter} historyMonthFilter={historyMonthFilter} setHistoryMonthFilter={setHistoryMonthFilter} historyStatusFilter={historyStatusFilter} setHistoryStatusFilter={setHistoryStatusFilter} historyMonths={historyMonths}/><Table rows={data.expenses} type="expenses" setData={setData} data={data} quickFilter={historyQuickFilter} monthFilter={historyMonthFilter}/></div></section>}
{tab==="alertas"&&<section className="alerts"><div className="card reminders compactReminders"><h2><Bell size={20}/>Cobros / Recordatorios</h2>
<div className="reminderTools">
  <div className="search inner"><Search size={17}/><input value={alertSearch} onChange={e=>setAlertSearch(e.target.value)} placeholder="Buscar factura o cliente..."/></div>
  <select value={reminderStatusFilter} onChange={e=>setReminderStatusFilter(e.target.value)}>
    <option>Todas</option>
    <option>Vencida</option>
    <option>Por vencer</option>
  </select>
  <div className="reminderCount"><b>{alertInvoices.length}</b><span>alertas</span></div>
</div>
<div className="reminderList compactList">{alertInvoices.map(inv=>{let c=client(inv.clienteId),s=ist(inv),Icon=s.I;return <button key={inv.id} className={`reminder compactReminder ${selectedInvoice?.id===inv.id?"selected":""}`} onClick={()=>setSelectedInvoiceId(inv.id)}><Icon className={s.c} size={20}/><div><b>{inv.factura}</b><p>{c?.nombre}</p></div><strong>{money(inv.monto)}</strong><small>{inv.vencimiento}</small></button>})}</div></div><div className="card attach stickyCobroAction"><h2><Paperclip size={20}/>Adjuntar factura</h2>{selectedInvoice?<><div className="selected"><b>{selectedInvoice.factura}</b><p>{selectedClient?.nombre} · {money(selectedInvoice.monto)}</p><small className="requiredAttach">Para usar Enviar factura, primero debe adjuntar la factura. El recordatorio no requiere adjunto.</small></div><label className="drop"><UploadCloud size={32}/><b>Buscar factura en mi PC</b><small>PDF, JPG, PNG, DOCX, XLSX</small><input type="file" onChange={e=>attachFile(selectedInvoice.id,e.target.files?.[0])}/></label>{data.attachments?.[selectedInvoice.id]&&<div className="fileBox"><FileText size={22}/><div><b>{data.attachments[selectedInvoice.id].name}</b><p>{(data.attachments[selectedInvoice.id].size/1024/1024).toFixed(2)} MB</p></div></div>}<div className="actions big"><div className="sendGroup"><b>Recordatorio</b><a className="send whatsapp" href={waReminder(selectedInvoice,selectedClient)} target="_blank"><W/>WhatsApp</a><a className="send mail" href={emailReminder(selectedInvoice,selectedClient)}><Mail size={18}/>Correo manual</a><button className="send autoMail" onClick={()=>sendAutoReminder(selectedInvoice,selectedClient)}>Enviar automático</button></div><div className="sendGroup"><b>Enviar factura</b><a className={`send whatsapp ${!data.attachments?.[selectedInvoice.id]?"disabled":""}`} href={data.attachments?.[selectedInvoice.id]?waInvoice(selectedInvoice,selectedClient):"#"} onClick={(e)=>{if(!canSendInvoice(selectedInvoice.id))e.preventDefault()}} target="_blank"><W/>WhatsApp</a><a className={`send mail ${!data.attachments?.[selectedInvoice.id]?"disabled":""}`} href={data.attachments?.[selectedInvoice.id]?emailInvoice(selectedInvoice,selectedClient):"#"} onClick={(e)=>{if(!canSendInvoice(selectedInvoice.id))e.preventDefault()}}><Mail size={18}/>Correo manual</a><button className={`send autoMail ${!data.attachments?.[selectedInvoice.id]?"disabled":""}`} onClick={()=>sendAutoInvoice(selectedInvoice,selectedClient)}>Enviar automático</button></div></div></>:<p>No hay facturas por cobrar.</p>}</div></section>}
</main></div>}
function HistoryFilters({historyQuickFilter,setHistoryQuickFilter,historyMonthFilter,setHistoryMonthFilter,historyStatusFilter,setHistoryStatusFilter,historyMonths,showStatus=false}){
 return <div className="historyFilters"><input value={historyQuickFilter} onChange={e=>setHistoryQuickFilter(e.target.value)} placeholder="Buscar en historial..."/><select value={historyMonthFilter} onChange={e=>setHistoryMonthFilter(e.target.value)}>{historyMonths.map(m=><option key={m} value={m}>{m==="Todos"?"Todos los meses":ml(m)}</option>)}</select>{showStatus&&<select value={historyStatusFilter} onChange={e=>setHistoryStatusFilter(e.target.value)}><option>Todos</option><option>Pendiente</option><option>Pagada</option><option>Vencida</option></select>}<button onClick={()=>{setHistoryQuickFilter("");setHistoryMonthFilter("Todos");setHistoryStatusFilter("Todos")}}>Limpiar</button></div>}

function Table({rows,type,setData,data,quickFilter="",monthFilter="Todos",statusFilter="Todos"}){
 let sorted=[...rows]
  .filter(r=>{
    const text=`${r.fecha||""} ${r.vencimiento||""} ${r.categoria||""} ${r.descripcion||""} ${r.proveedor||""} ${r.emailProveedor||""} ${r.estado||""}`.toLowerCase();
    return text.includes(String(quickFilter||"").toLowerCase());
  })
  .filter(r=>monthFilter==="Todos"||mk(r.fecha||r.vencimiento)===monthFilter||mk(r.vencimiento||r.fecha)===monthFilter)
  .filter(r=>statusFilter==="Todos"||String(r.estado||"").toLowerCase()===String(statusFilter).toLowerCase())
  .sort((a,b)=>(b.fecha||b.vencimiento).localeCompare(a.fecha||a.vencimiento));
 let total=sorted.reduce((s,r)=>s+(+r.monto||0),0);
 return <div className="historyBox"><div className="historyMiniStats"><b>{sorted.length}</b><span>registros</span><strong>{money(total)}</strong></div><div className="tableWrap historyScroll"><table><thead><tr><th>Fecha</th><th>Mes/Año</th><th>Categoría</th><th>Descripción</th><th>Monto</th><th>Acciones</th></tr></thead><tbody>{sorted.map(r=><tr key={r.id}><td>{r.fecha||r.vencimiento}</td><td>{ml(mk(r.fecha||r.vencimiento))}</td><td>{r.categoria}{r.proveedor&&<small>{r.proveedor}</small>}{r.emailProveedor&&<small>{r.emailProveedor}</small>}</td><td>{r.descripcion}{r.estado&&<small>Estado: {r.estado}</small>}</td><td>{money(r.monto)}</td><td>{type==="debts"&&<a className="icon mail" href={debtEmail(r)} title="Correo manual deuda"><Mail size={17}/></a>}{type==="debts"&&typeof sendAutoDebt==="function"&&<button className="icon autoMailIcon" onClick={()=>sendAutoDebt(r)} title="Enviar deuda automático">AUTO</button>}<button className="icon trash" onClick={()=>setData({...data,[type]:(data[type]||[]).filter(x=>x.id!==r.id)})}><Trash2 size={17}/></button></td></tr>)}</tbody></table></div></div>}

createRoot(document.getElementById("root")).render(<ErrorBoundary><App/></ErrorBoundary>);
