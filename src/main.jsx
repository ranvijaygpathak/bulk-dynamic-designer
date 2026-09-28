import React,{useEffect,useMemo,useRef,useState} from "react";
import {createRoot} from "react-dom/client";
import {Canvas,FabricText,Rect,Circle,Line,FabricImage} from "fabric";
import * as XLSX from "xlsx";
import {jsPDF} from "jspdf";
import JSZip from "jszip";
import {saveAs} from "file-saver";
import QRCode from "qrcode";
import * as pdfjsLib from "pdfjs-dist";
import pdfWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;
import {Type,Image as ImageIcon,Square,Circle as CircleIcon,Minus,QrCode,Trash2,Copy,Lock,Unlock,ChevronUp,ChevronDown,Upload,Download,Save,FolderOpen,Layers,FileSpreadsheet,Link2,RefreshCw,ChevronLeft,ChevronRight} from "lucide-react";
import "./styles.css";

const UNICODE_FONTS=[
{name:"Mangal",family:"Mangal",note:"Windows Hindi"},
{name:"Nirmala UI",family:"Nirmala UI",note:"Windows Unicode"},
{name:"Noto Sans Devanagari",family:"Noto Sans Devanagari",note:"Devanagari"},
{name:"Noto Serif Devanagari",family:"Noto Serif Devanagari",note:"Devanagari"},
{name:"Hind",family:"Hind",note:"Devanagari"},
{name:"Mukta",family:"Mukta",note:"Devanagari"},
{name:"Kokila",family:"Kokila",note:"Windows Devanagari"},
{name:"Aparajita",family:"Aparajita",note:"Windows Devanagari"},
{name:"Arial Unicode MS",family:"Arial Unicode MS",note:"Unicode"},
{name:"Arial",family:"Arial",note:"Latin"},
{name:"Georgia",family:"Georgia",note:"Latin"},
{name:"Times New Roman",family:"Times New Roman",note:"Latin"},
{name:"Verdana",family:"Verdana",note:"Latin"}];
async function ensureFontLoaded(family,size=32){try{if(document.fonts?.load)await document.fonts.load(`${size}px "${family}"`)}catch(e){}}
const uid=()=>Math.random().toString(36).slice(2,10);
const presets={
"A4 Portrait":[794,1123],"A4 Landscape":[1123,794],"A3 Portrait":[1123,1587],
"A3 Landscape":[1587,1123],"A5 Portrait":[559,794],"A5 Landscape":[794,559],
"Letter Portrait":[612,792],"Letter Landscape":[792,612],"Custom":[1000,700]
};
const tokenRe=/\{\{(.*?)\}\}/g;

function App(){
 const el=useRef(null), cRef=useRef(null),[selected,setSelected]=useState(null);
 const [tab,setTab]=useState("design"),[paper,setPaper]=useState("A4 Landscape"),[zoom,setZoom]=useState(.75);
 const [name,setName]=useState("My Certificate"),[bg,setBg]=useState("#ffffff"),[status,setStatus]=useState("Ready");
 const [rows,setRows]=useState([]),[cols,setCols]=useState([]),[idx,setIdx]=useState(0),[map,setMap]=useState({});
 const [objects,setObjects]=useState([]),[fontStatus,setFontStatus]=useState("Unicode fonts available");
 const [file,setFile]=useState(null);
 const [width,height]=presets[paper];

 useEffect(()=>{
   const c=new Canvas(el.current,{width,height,backgroundColor:bg,preserveObjectStacking:true});
   cRef.current=c;
   c.on("selection:created",e=>setSelected(e.selected?.[0]||e.target));
   c.on("selection:updated",e=>setSelected(e.selected?.[0]||e.target));
   c.on("selection:cleared",()=>setSelected(null));
   c.on("object:modified",sync);c.on("object:added",sync);c.on("object:removed",sync);
   starter(c,width,height);
   return()=>c.dispose();
 },[]);
 useEffect(()=>{let c=cRef.current;if(!c)return;c.setDimensions({width:width*zoom,height:height*zoom});c.setZoom(zoom);c.renderAll()},[paper,zoom]);
 useEffect(()=>{let c=cRef.current;if(c){c.backgroundColor=bg;c.renderAll()}},[bg]);

 function starter(c,w,h){
   c.clear(); c.backgroundColor=bg;
   c.add(new Rect({left:12,top:12,width:w-24,height:h-24,fill:"transparent",stroke:"#183153",strokeWidth:5,selectable:false,evented:false,customId:uid()}));
   c.add(new FabricText("CERTIFICATE OF COMPLETION",{left:w/2,top:h*.13,originX:"center",originY:"center",fontSize:42,fontWeight:"bold",fill:"#183153",customId:uid(),templateText:"CERTIFICATE OF COMPLETION"}));
   c.add(new FabricText("{{name}}",{left:w/2,top:h*.40,originX:"center",originY:"center",fontSize:52,fontWeight:"bold",fill:"#111827",customId:uid(),dynamicField:"name",templateText:"{{name}}"}));
   c.add(new FabricText("Son/Daughter of {{father_name}}",{left:w/2,top:h*.51,originX:"center",originY:"center",fontSize:27,fill:"#374151",customId:uid(),templateText:"Son/Daughter of {{father_name}}"}));
   c.add(new FabricText("Course: {{course}}",{left:w/2,top:h*.61,originX:"center",originY:"center",fontSize:30,fill:"#374151",customId:uid(),templateText:"Course: {{course}}"}));
   c.renderAll();sync();
 }
 function sync(){let c=cRef.current;if(!c)return;setObjects(c.getObjects().map(o=>({id:o.customId,type:o.type,field:o.dynamicField||"",text:o.text||"",locked:!!o.locked})))}

 async function importExistingDesign(file){
   if(!file)return;
   setStatus("Importing existing design...");
   try{
     let dataUrl;
     if(file.type==="application/pdf" || file.name.toLowerCase().endsWith(".pdf")){
       const bytes=await file.arrayBuffer();
       const pdf=await pdfjsLib.getDocument({data:bytes}).promise;
       const page=await pdf.getPage(1);
       const viewport=page.getViewport({scale:2});
       const cv=document.createElement("canvas");
       cv.width=viewport.width; cv.height=viewport.height;
       await page.render({canvasContext:cv.getContext("2d"),viewport}).promise;
       dataUrl=cv.toDataURL("image/png");
     }else{
       dataUrl=await new Promise((resolve,reject)=>{
         const r=new FileReader(); r.onload=()=>resolve(r.result); r.onerror=reject; r.readAsDataURL(file);
       });
     }
     const img=await FabricImage.fromURL(dataUrl);
     const c=cRef.current;
     c.getObjects().filter(o=>o.isImportedBackground).forEach(o=>c.remove(o));
     const scale=Math.max(width/img.width,height/img.height);
     img.set({left:width/2,top:height/2,originX:"center",originY:"center",
       scaleX:scale,scaleY:scale,selectable:false,evented:false,
       customId:uid(),isImportedBackground:true});
     c.insertAt(0,img); c.renderAll(); sync(); 
     setStatus("Existing design imported. Add dynamic fields on top.");
     setTab("design");
   }catch(e){console.error(e);setStatus("Could not import the design.");}
 }
 async function changeFont(family){
   if(!selected || selected.type!=="text")return;
   await ensureFontLoaded(family,selected.fontSize||30);
   selected.set({fontFamily:family});
   cRef.current?.renderAll(); sync(); setFontStatus(`Font: ${family}`);
 }

 function addText(field=""){let c=cRef.current,t=new FabricText(field?`{{${field}}}`:"New Text",{left:width/2,top:height*.25,originX:"center",originY:"center",fontSize:30,fontFamily:"Noto Sans Devanagari",fill:"#111827",customId:uid(),dynamicField:field,templateText:field?`{{${field}}}`:"New Text"});c.add(t);c.setActiveObject(t);c.renderAll();setSelected(t)}
 function addShape(kind){let c=cRef.current,o=kind==="circle"?new Circle({left:width/2-60,top:height*.3,radius:60,fill:"transparent",stroke:"#183153",strokeWidth:3,customId:uid()}):new Rect({left:width/2-100,top:height*.3,width:200,height:120,fill:"transparent",stroke:"#183153",strokeWidth:3,rx:10,ry:10,customId:uid()});c.add(o);c.setActiveObject(o);c.renderAll();setSelected(o)}
 function addLine(){let c=cRef.current,l=new Line([width*.35,height*.4,width*.65,height*.4],{stroke:"#183153",strokeWidth:3,customId:uid()});c.add(l);c.setActiveObject(l);c.renderAll();setSelected(l)}
 async function addImage(f){if(!f)return;let img=await FabricImage.fromURL(URL.createObjectURL(f));let s=Math.min(width*.35/img.width,height*.3/img.height,1);img.set({left:width/2,top:height*.28,originX:"center",originY:"center",scaleX:s,scaleY:s,customId:uid()});cRef.current.add(img);cRef.current.setActiveObject(img);cRef.current.renderAll();setSelected(img)}
 async function addQR(){let data=rows[idx]?.certificate_no||rows[idx]?.Certificate_No||"{{certificate_no}}";let u=await QRCode.toDataURL(String(data),{width:180,margin:1});let img=await FabricImage.fromURL(u);img.set({left:width*.84,top:height*.83,originX:"center",originY:"center",customId:uid(),isQR:true,dynamicField:"qr_code"});cRef.current.add(img);cRef.current.setActiveObject(img);cRef.current.renderAll();setSelected(img)}
 function update(p,v){if(!selected)return;selected.set({[p]:v});cRef.current.renderAll();sync()}
 function bind(v){if(!selected||selected.type!=="text")return;selected.dynamicField=v;selected.templateText=`{{${v}}}`;selected.set({text:`{{${v}}}`});cRef.current.renderAll();sync()}
 function remove(){if(selected){cRef.current.remove(selected);setSelected(null);cRef.current.renderAll()}}
 function duplicate(){if(!selected)return;selected.clone().then(x=>{x.set({left:(x.left||0)+25,top:(x.top||0)+25,customId:uid()});cRef.current.add(x);cRef.current.setActiveObject(x);cRef.current.renderAll()})}
 function layer(d){if(!selected)return;d==="up"?cRef.current.bringObjectForward(selected):cRef.current.sendObjectBackwards(selected);cRef.current.renderAll();sync()}
 function lock(){if(!selected)return;selected.locked=!selected.locked;selected.set({selectable:!selected.locked,evented:!selected.locked});cRef.current.renderAll();sync()}
 function importExcel(f){if(!f)return;let r=new FileReader();r.onload=e=>{let wb=XLSX.read(e.target.result,{type:"array"}),ws=wb.Sheets[wb.SheetNames[0]],data=XLSX.utils.sheet_to_json(ws,{defval:""});if(!data.length)return setStatus("No data found.");setRows(data);setCols(Object.keys(data[0]));setIdx(0);setStatus(`${data.length} records imported.`);setTab("data")};r.readAsArrayBuffer(f)}
 const fields=useMemo(()=>{let s=new Set();cRef.current?.getObjects().forEach(o=>{if(o.dynamicField)s.add(o.dynamicField);(o.text||"").match(tokenRe)?.forEach(x=>s.add(x.slice(2,-2).trim()))});return [...s]},[objects,rows]);
 function renderRow(row){let c=cRef.current;if(!c)return;c.getObjects().forEach(o=>{if(o.type==="text"&&!o.isBackground){let t=o.templateText??o.text??"";o.templateText=t;o.set({text:t.replace(tokenRe,(_,k)=>row[map[k.trim()]||k.trim()]??"")})}});c.renderAll()}
 useEffect(()=>{if(rows.length)renderRow(rows[idx])},[idx]);
 function reset(){cRef.current?.getObjects().forEach(o=>{if(o.type==="text"&&o.templateText)o.set({text:o.templateText})});cRef.current?.renderAll()}
 function save(){let data={name,paper,bg,width,height,objects:cRef.current.toJSON(["customId","dynamicField","templateText","locked","isQR","isImportedBackground"]).objects};saveAs(new Blob([JSON.stringify(data,null,2)],{type:"application/json"}),`${name.replace(/\s+/g,"_")}.json`);setStatus("Template saved.")}
 function load(f){if(!f)return;let r=new FileReader();r.onload=async e=>{let d=JSON.parse(e.target.result);setName(d.name||"Template");setPaper(d.paper||"A4 Landscape");setBg(d.bg||"#fff");let c=cRef.current;c.clear();await c.loadFromJSON(d);c.renderAll();sync();setStatus("Template loaded.")};r.readAsText(f)}
 async function generate(kind){if(!rows.length)return setStatus("Import Excel data first.");let c=cRef.current,zip=new JSZip(),old=zoom;c.setZoom(1);c.setDimensions({width,height});for(let i=0;i<rows.length;i++){renderRow(rows[i]);let png=c.toDataURL({format:"png"}),base=png.split(",")[1],n=String(rows[i].name||rows[i].Name||`record_${i+1}`).replace(/[^\w\u0900-\u097F-]+/g,"_");if(kind==="png")zip.file(`${i+1}_${n}.png`,base,{base64:true});else{let pdf=new jsPDF({orientation:width>=height?"landscape":"portrait",unit:"px",format:[width,height]});pdf.addImage(png,"PNG",0,0,width,height);zip.file(`${i+1}_${n}.pdf`,pdf.output("arraybuffer"))}}reset();c.setZoom(old);c.setDimensions({width:width*old,height:height*old});c.renderAll();saveAs(await zip.generateAsync({type:"blob"}),`${name.replace(/\s+/g,"_")}_${kind}.zip`);setStatus(`${rows.length} ${kind.toUpperCase()} files generated.`)}

 return <div className="app">
  <header><div className="brand"><b>BD</b><span><strong>Bulk Dynamic Designer</strong><small>Certificate • Poster • Banner</small></span></div><input value={name} onChange={e=>setName(e.target.value)}/><button onClick={save}><Save/> Save</button><button onClick={()=>document.getElementById("load").click()}><FolderOpen/> Load</button><input hidden id="load" type="file" accept=".json" onChange={e=>load(e.target.files[0])}/></header>
  <nav>{[["design","Design",Layers],["data","Excel Data",FileSpreadsheet],["mapping","Field Mapping",Link2],["generate","Bulk Generate",Download]].map(([k,n,I])=><button className={tab===k?"active":""} onClick={()=>setTab(k)} key={k}><I/>{n}{k==="data"&&rows.length?` (${rows.length})`:""}</button>)}</nav>
  <main>
   <aside className="left"><h3>Elements</h3><button onClick={()=>addText()}><Type/> Text</button>
          <button onClick={()=>document.getElementById("existingDesignInput").click()}><Upload/> Import Existing Design</button>
          <input hidden id="existingDesignInput" type="file" accept=".png,.jpg,.jpeg,.pdf,image/png,image/jpeg,application/pdf" onChange={e=>importExistingDesign(e.target.files[0])}/><button onClick={()=>document.getElementById("image").click()}><ImageIcon/> Image</button><input hidden id="image" type="file" accept="image/*" onChange={e=>addImage(e.target.files[0])}/><button onClick={addQR}><QrCode/> Dynamic QR</button><button onClick={()=>addShape("rect")}><Square/> Rectangle</button><button onClick={()=>addShape("circle")}><CircleIcon/> Circle</button><button onClick={addLine}><Minus/> Line</button><hr/><h3>Paper</h3><select value={paper} onChange={e=>{setPaper(e.target.value);setStatus(`${e.target.value} selected`)}}>{Object.keys(presets).map(x=><option key={x}>{x}</option>)}</select><label>Background <input type="color" value={bg} onChange={e=>setBg(e.target.value)}/></label><div className="unicodeBox"><b>Unicode / Hindi / Bhojpuri</b><span>Devanagari Unicode works in text objects and Excel-driven fields.</span><small>{fontStatus}</small></div><hr/><h3>Selected Object</h3>{selected?<><label>Dynamic field<select value={selected.dynamicField||""} onChange={e=>bind(e.target.value)}><option value="">Static</option>{cols.map(c=><option key={c}>{c}</option>)}</select></label>{selected.type==="text"&&<><label>Font size<input type="number" value={selected.fontSize||20} onChange={e=>update("fontSize",+e.target.value)}/></label><label>Font<select value={selected.fontFamily||"Noto Sans Devanagari"} onChange={e=>changeFont(e.target.value)}>{UNICODE_FONTS.map(f=><option key={f.family} value={f.family}>{f.name} — {f.note}</option>)}</select></label><label>Color<input type="color" value={typeof selected.fill==="string"?selected.fill:"#111"} onChange={e=>update("fill",e.target.value)}/></label></>}<div className="two"><button onClick={duplicate}><Copy/> Copy</button><button onClick={lock}>{selected.locked?<Unlock/>:<Lock/>}{selected.locked?"Unlock":"Lock"}</button></div><div className="two"><button onClick={()=>layer("up")}><ChevronUp/> Forward</button><button onClick={()=>layer("down")}><ChevronDown/> Back</button></div><button className="danger" onClick={remove}><Trash2/> Delete</button></>:<p>Select an object on the canvas.</p>}</aside>
   <section className="work"><div className="toolbar"><span>{status}</span><span>Zoom <input type="range" min=".4" max="1.1" step=".05" value={zoom} onChange={e=>setZoom(+e.target.value)}/></span></div><div className="canvas"><canvas ref={el}/></div>{rows.length>0&&<div className="records"><button disabled={!idx} onClick={()=>setIdx(idx-1)}><ChevronLeft/></button><b>Record {idx+1} / {rows.length}</b><button disabled={idx===rows.length-1} onClick={()=>setIdx(idx+1)}><ChevronRight/></button><span>{rows[idx]?.name||rows[idx]?.Name||"Preview"}</span></div>}</section>
   <aside className="right">{tab==="design"&&<><h3>Dynamic Fields</h3>{fields.map(f=><div className="chip" key={f}>{"{{"}{f}{"}}"}</div>)}<h3>Layers</h3>{objects.slice().reverse().map(o=><div className="layer" key={o.id}>{o.type}{o.field?` → ${o.field}`:""}{o.locked?" 🔒":""}</div>)}</>}{tab==="data"&&<><h3>Import Excel / CSV</h3><button className="primary" onClick={()=>document.getElementById("excel").click()}><Upload/> Upload Excel</button><input hidden id="excel" type="file" accept=".xlsx,.xls,.csv" onChange={e=>importExcel(e.target.files[0])}/>{rows.length>0&&<p><b>{rows.length}</b> records • <b>{cols.length}</b> columns</p>}<div className="table"><table><thead><tr>{cols.map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{rows.slice(0,15).map((r,i)=><tr key={i}>{cols.map(c=><td key={c}>{String(r[c])}</td>)}</tr>)}</tbody></table></div></>}{tab==="mapping"&&<><h3>Field Mapping</h3>{fields.map(f=><label key={f}>{"{{"}{f}{"}}"}<select value={map[f]||f} onChange={e=>setMap({...map,[f]:e.target.value})}><option value="">-- select --</option>{cols.map(c=><option key={c}>{c}</option>)}</select></label>)}<button className="primary" onClick={()=>setStatus("Mapping saved.")}><Save/> Save Mapping</button></>}{tab==="generate"&&<><h3>Bulk Generate</h3><div className="card"><b>{rows.length} records ready</b><button className="primary" onClick={()=>generate("pdf")}><Download/> PDF ZIP</button><button onClick={()=>generate("png")}><Download/> PNG ZIP</button></div><button onClick={()=>{reset();setStatus("Preview reset.")}}><RefreshCw/> Reset</button></>}</aside>
  </main>
 </div>
}
createRoot(document.getElementById("root")).render(<App/>);
