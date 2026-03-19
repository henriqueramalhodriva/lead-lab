import { useState, useRef } from "react";
import * as XLSX from "xlsx";

const DEFAULT_TEMPLATE_COLS = [
  "RAIZ CNPJ","CNPJ","RAZÃO SOCIAL","NOME FANTASIA",
  "CNAE FISCAL","CNAE SUBCLASSE","CNAE SECUNDÁRIO VAREJO",
  "TPV ESTIMADO","FATURAMENTO","FUNCIONÁRIOS","FILIAIS",
  "ENDEREÇO","BAIRRO","CIDADE","ESTADO","CEP","DOMINIO",
  "PLATAFORMA","ACESSOS","TICKET MÉDIO","INSTAGRAM",
  "SEGUIDORES","EMAIL - PERSONAS","LINKEDIN",
  "E-MAIL EMPRESAS","TELEFONE","TELEFONE WHATSAPP"
];
const GENERATED_COLS = ["E-MAIL EMPRESAS","TELEFONE","TELEFONE WHATSAPP","DOMINIO","EMAIL - PERSONAS","LINKEDIN","PLATAFORMA","ACESSOS","TICKET MÉDIO"];
const BG="#f0f2f5",CARD="#fff",DARK="#1a2340",MID="#4a5568",LIGHT="#718096",ORANGE="#e8520a",BORDER="#e2e8f0",ACCENT_BG="#f7f8fa";

// ─── CNAE ────────────────────────────────────────────────────────────────────
const CNAE_TABLE={"4711-3/02":{taxa:0.032,ticket:200},"4711-3/01":{taxa:0.030,ticket:150},"4791-1/00":{taxa:0.030,ticket:150},"4771-7/02":{taxa:0.029,ticket:110},"4729-6/02":{taxa:0.029,ticket:35},"4712-1/00":{taxa:0.028,ticket:80},"4771-7/01":{taxa:0.028,ticket:95},"4789-0/06":{taxa:0.028,ticket:120},"4721-1/03":{taxa:0.028,ticket:65},"4792-9/00":{taxa:0.028,ticket:80},"4773-3/00":{taxa:0.027,ticket:120},"4789-0/05":{taxa:0.027,ticket:85},"4721-1/02":{taxa:0.027,ticket:50},"4724-5/00":{taxa:0.027,ticket:55},"4729-6/99":{taxa:0.027,ticket:90},"4789-0/04":{taxa:0.027,ticket:85},"4721-1/01":{taxa:0.026,ticket:45},"4721-1/04":{taxa:0.026,ticket:35},"4723-7/00":{taxa:0.026,ticket:65},"4771-7/03":{taxa:0.025,ticket:85},"4722-9/01":{taxa:0.025,ticket:70},"4729-6/03":{taxa:0.025,ticket:110},"4789-0/08":{taxa:0.025,ticket:55},"4722-9/02":{taxa:0.024,ticket:85},"4789-0/01":{taxa:0.007,ticket:65},"4763-6/03":{taxa:0.006,ticket:15000},"4774-1/00":{taxa:0.006,ticket:450},"4783-1/02":{taxa:0.006,ticket:320},"4783-1/01":{taxa:0.005,ticket:1850},"4785-7/99":{taxa:0.005,ticket:180},"4763-6/02":{taxa:0.005,ticket:2800},"4789-0/09":{taxa:0.005,ticket:2800},"4789-0/03":{taxa:0.004,ticket:2500},"4789-0/99":{taxa:0.020,ticket:95},"4644-3/01":{taxa:0.018,ticket:850},"4644-3/02":{taxa:0.015,ticket:650},"4669-9/99":{taxa:0.008,ticket:2500},"4663-0/00":{taxa:0.006,ticket:8500},"4511-1/02":{taxa:0.003,ticket:45000},"4541-2/03":{taxa:0.003,ticket:15000},"4541-2/04":{taxa:0.003,ticket:8500},"4511-1/01":{taxa:0.002,ticket:85000},"4541-2/01":{taxa:0.002,ticket:15000},"4763-6/05":{taxa:0.014,ticket:120},"4781-4/00":{taxa:0.013,ticket:165},"4755-5/01":{taxa:0.012,ticket:95},"4782-2/01":{taxa:0.012,ticket:145},"4782-2/03":{taxa:0.012,ticket:180},"4799-5/99":{taxa:0.012,ticket:150},"4755-5/03":{taxa:0.011,ticket:220},"4782-2/02":{taxa:0.011,ticket:195},"4763-6/04":{taxa:0.011,ticket:165},"4756-3/00":{taxa:0.010,ticket:650},"4754-7/01":{taxa:0.010,ticket:850},"4754-7/02":{taxa:0.010,ticket:950},"4772-5/00":{taxa:0.022,ticket:95},"4789-0/02":{taxa:0.022,ticket:45},"4731-8/00":{taxa:0.022,ticket:150},"4761-0/03":{taxa:0.021,ticket:50},"4762-8/00":{taxa:0.021,ticket:50},"4761-0/01":{taxa:0.020,ticket:65},"4784-9/00":{taxa:0.020,ticket:85},"4761-0/02":{taxa:0.019,ticket:35},"4729-6/01":{taxa:0.018,ticket:45},"4751-2/01":{taxa:0.018,ticket:450},"4763-6/01":{taxa:0.018,ticket:120},"4799-5/01":{taxa:0.018,ticket:35},"4751-2/02":{taxa:0.017,ticket:280},"4753-9/00":{taxa:0.017,ticket:620},"4754-7/03":{taxa:0.017,ticket:45},"4754-7/04":{taxa:0.017,ticket:75},"4752-1/00":{taxa:0.016,ticket:850},"4789-0/07":{taxa:0.016,ticket:75},"4755-5/02":{taxa:0.015,ticket:180},"4732-6/00":{taxa:0.009,ticket:95},"4753-9/01":{taxa:0.008,ticket:1650},"4759-8/01":{taxa:0.008,ticket:1200},"4741-5/00":{taxa:0.008,ticket:165},"4530-7/01":{taxa:0.008,ticket:380},"4530-7/03":{taxa:0.008,ticket:380},"4541-2/05":{taxa:0.008,ticket:220},"4742-3/00":{taxa:0.007,ticket:180},"4744-0/01":{taxa:0.007,ticket:185},"4744-0/03":{taxa:0.007,ticket:195},"4744-0/05":{taxa:0.007,ticket:320},"4744-0/99":{taxa:0.007,ticket:420},"4530-7/02":{taxa:0.007,ticket:650},"4530-7/04":{taxa:0.007,ticket:220},"4530-7/05":{taxa:0.007,ticket:650},"4541-2/02":{taxa:0.007,ticket:280},"4743-1/00":{taxa:0.006,ticket:250},"4744-0/02":{taxa:0.006,ticket:420},"4744-0/04":{taxa:0.006,ticket:280},"4530-7/06":{taxa:0.006,ticket:450}};

function parseFat(f){if(!f)return null;const s=String(f).toUpperCase().replace(/\s/g,"");const m={K:1e3,M:1e6,B:1e9};const en=str=>{const x=str.match(/([\d.,]+)([KMB]?)/);return x?parseFloat(x[1].replace(",","."))*( m[x[2]]||1):null;};const p=s.split(/A|-|–/).map(v=>v.trim()).filter(Boolean);const n=p.map(en).filter(v=>v!==null);return n.length?Math.max(...n):null;}
function parseSeg(v){if(!v)return 0;const s=String(v).toUpperCase().replace(/\s/g,"");const m={K:1e3,M:1e6};const x=s.match(/([\d.,]+)([KM]?)/);return x?parseFloat(x[1].replace(",","."))*( m[x[2]]||1):0;}
function tSc(v){if(v>=30000)return 10;if(v>=20000)return 8;if(v>=15000)return 6;if(v>=8000)return 4;return 0;}
function fSc(v){if(v>=500000)return 10;if(v>=200000)return 8;if(v>=100000)return 6;if(v>=30000)return 4;if(v>=10000)return 2;return 0;}
function oPct(f){if(!f)return 0.25;if(f>=1e9)return 0.50;if(f>=3e8)return 0.42;if(f>=1e8)return 0.38;if(f>=5e7)return 0.35;if(f>=2e7)return 0.32;if(f>=1e7)return 0.30;if(f>=4.8e6)return 0.27;if(f>=1e6)return 0.22;return 0.20;}
function calcTier(m){if(m>=5e6)return"Tier IV";if(m>=1e6)return"Tier III";if(m>=5e5)return"Tier II";if(m>=3e5)return"Tier I";return"Micro";}
function calcTPV(row){
  const ft=parseFat(row["FATURAMENTO"]);const ck=String(row["CNAE SUBCLASSE"]||row["CNAE FISCAL"]||"").trim();const cd=CNAE_TABLE[ck]||{};const tb=cd.taxa||0.012;
  const vis=parseFloat(String(row["ACESSOS"]||"0").replace(/[^\d.]/g,""))||0;const sT=tSc(vis);const hT=vis>=8000;
  const seg=parseSeg(row["SEGUIDORES"]);const sS=fSc(seg);const tE=hT?vis:seg*0.02;
  let tk=parseFloat(String(row["TICKET MÉDIO"]||"").replace(/[^\d.,]/g,"").replace(",","."))||0;
  if(tk<30||tk>15000)tk=cd.ticket||150;if(tk<30)tk=150;
  const sd=(sT+sS)/2;const aj=((sd-5)/10)*0.15;const tf=Math.min(0.04,Math.max(0.002,tb*(1+aj)));
  let fa=null;
  if(hT||tE>=8000){const bu=tk*tE*tf*12;if(ft){const p=Math.min(0.9,Math.max(0.05,oPct(ft)*(1+aj*0.15)));fa=Math.min(bu,ft*p);}else fa=bu;}
  else if(ft){const p=Math.min(0.9,Math.max(0.05,oPct(ft)*(1+aj*0.15)));fa=ft*p;}
  if(!fa||fa<=0)return{tpv:"",tier:""};const tier=calcTier(fa/12);return{tpv:tier,tier};
}

function expandRows(rows){
  const out=[];
  rows.forEach(row=>{
    const ph=(row["TELEFONE"]||"").split(",").map(v=>v.trim()).filter(Boolean);
    const wh=(row["TELEFONE WHATSAPP"]||"").split(",").map(v=>v.trim()).filter(Boolean);
    const em=(row["E-MAIL EMPRESAS"]||"").split(",").map(v=>v.trim()).filter(Boolean);

    // Personas: pair email+linkedin by index to keep them together
    const peEmails=(row["EMAIL - PERSONAS"]||"").split(",").map(v=>v.trim());
    const liLinks=(row["LINKEDIN"]||"").split(",").map(v=>v.trim());
    const pairCount=Math.max(peEmails.filter(Boolean).length, liLinks.filter(Boolean).length);
    const pairs=Array.from({length:pairCount},(_,i)=>({email:peEmails[i]||"",linkedin:liLinks[i]||""})).filter(p=>p.email||p.linkedin);

    const base={...row,"TELEFONE":"","TELEFONE WHATSAPP":"","E-MAIL EMPRESAS":"","EMAIL - PERSONAS":"","LINKEDIN":""};
    const add=ov=>out.push({...base,...ov});

    ph.forEach(v=>add({"TELEFONE":v}));
    wh.forEach(v=>add({"TELEFONE WHATSAPP":v}));
    em.forEach(v=>add({"E-MAIL EMPRESAS":v}));
    // Each persona pair gets its own line with both email and linkedin
    pairs.forEach(p=>add({"EMAIL - PERSONAS":p.email,"LINKEDIN":p.linkedin}));

    if(!ph.length&&!wh.length&&!em.length&&!pairs.length)out.push({...row});
  });
  return out;
}

// ─── UI helpers ───────────────────────────────────────────────────────────────
function Tag({label,color}){return <span style={{display:"inline-block",padding:"3px 10px",borderRadius:20,background:color+"14",color,fontSize:11,fontWeight:600,border:"1px solid "+color+"30"}}>{label}</span>;}
function Card({title,accent,children}){return(<div style={{background:CARD,borderRadius:12,padding:20,marginBottom:14,border:"1px solid "+BORDER,boxShadow:"0 1px 3px rgba(0,0,0,0.04)"}}>{title&&<div style={{fontWeight:700,fontSize:14,marginBottom:14,color:accent||DARK,display:"flex",alignItems:"center",gap:8}}><span style={{width:3,height:16,borderRadius:2,background:accent||DARK,display:"inline-block"}}/>{title}</div>}{children}</div>);}
function Bar({pct,color}){return <div style={{background:"#edf2f7",borderRadius:999,height:5}}><div style={{background:color,borderRadius:999,height:5,width:pct+"%"}}/></div>;}
function StatCard({label,value,pct,color}){return(<div style={{background:ACCENT_BG,borderRadius:10,padding:"14px 16px",border:"1px solid "+BORDER}}><div style={{fontSize:24,fontWeight:800,color:color||DARK,marginBottom:2}}>{value}</div><div style={{fontSize:11,color:LIGHT,fontWeight:600,marginBottom:pct?4:0}}>{label}</div>{pct!==undefined&&<div style={{fontSize:11,color,fontWeight:700}}>{pct}%</div>}</div>);}
function SelField({lbl,val,setter,opts,sampleRow,accent}){return(<div style={{flex:1,minWidth:130}}><div style={{fontSize:11,color:LIGHT,fontWeight:700,marginBottom:4,textTransform:"uppercase",letterSpacing:.5}}>{lbl}</div><select value={val} onChange={e=>setter(e.target.value)} style={{width:"100%",padding:"7px 10px",borderRadius:8,border:"1px solid "+(val?accent||BORDER:BORDER),fontSize:12,background:val&&accent?accent+"10":CARD,color:DARK}}><option value="">-- ignorar --</option>{opts.map(h=><option key={h} value={h}>{h}{sampleRow&&sampleRow[h]?" ("+String(sampleRow[h]).slice(0,14)+")":""}</option>)}</select></div>);}
function ColMapper({headers,sampleRow,mapping,onChange,manualCols}){return(<div style={{display:"flex",flexDirection:"column",gap:6}}>{(manualCols||[]).map(tc=>(<div key={tc} style={{display:"flex",alignItems:"center",gap:8}}><span style={{width:190,fontSize:12,color:MID,fontWeight:600,flexShrink:0}}>{tc}</span><select value={mapping[tc]||""} onChange={e=>onChange({...mapping,[tc]:e.target.value})} style={{flex:1,padding:"5px 8px",borderRadius:7,border:"1px solid "+(mapping[tc]?"#4f46e5":BORDER),fontSize:12,color:DARK,background:mapping[tc]?"#f0f0fe":CARD}}><option value="">-- ignorar --</option>{headers.map(h=><option key={h} value={h}>{h}{sampleRow&&sampleRow[h]?" ("+String(sampleRow[h]).slice(0,20)+")":""}</option>)}</select></div>))}</div>);}

// ─── Multi-sheet toggle helper ────────────────────────────────────────────────
function SheetToggle({sheetNames,selected,onToggle,color}){
  return(<div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:14}}>
    {sheetNames.map(n=>(
      <button key={n} onClick={()=>onToggle(n)} style={{padding:"5px 14px",borderRadius:20,cursor:"pointer",fontSize:12,border:"1.5px solid "+(selected.includes(n)?color:BORDER),background:selected.includes(n)?color:CARD,fontWeight:selected.includes(n)?700:400,color:selected.includes(n)?"#fff":MID,display:"flex",alignItems:"center",gap:5}}>
        {selected.includes(n)&&<svg width="10" height="10" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="#fff" strokeWidth="2" fill="none"/></svg>}
        {n}
      </button>
    ))}
  </div>);
}

// ─── Multi-sheet contact picker ───────────────────────────────────────────────
function MultiSheetContactPicker({title,accent,sheetNames,allRows,selectedSheets,setSelectedSheets,colConfigs,setColConfigs,fields}){
  const toggle=sk=>setSelectedSheets(prev=>prev.includes(sk)?prev.filter(s=>s!==sk):[...prev,sk]);
  const upd=(sk,field,val)=>setColConfigs(prev=>({...prev,[sk]:{...(prev[sk]||{}),[field]:val}}));
  return(<Card title={title} accent={accent}>
    <p style={{fontSize:13,color:MID,margin:"0 0 12px"}}>Selecione uma ou mais abas e configure as colunas de cada uma.</p>
    <SheetToggle sheetNames={sheetNames} selected={selectedSheets} onToggle={toggle} color={accent}/>
    {selectedSheets.map(sk=>{
      const h=allRows[sk]&&allRows[sk].length>0?Object.keys(allRows[sk][0]):[];
      const sr=allRows[sk]?allRows[sk][0]:null;
      const cfg=colConfigs[sk]||{};
      return(<div key={sk} style={{background:ACCENT_BG,borderRadius:10,padding:"12px 14px",marginBottom:10,border:"1px solid "+BORDER}}>
        <div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10,display:"flex",alignItems:"center",gap:6}}>
          <span style={{width:8,height:8,borderRadius:"50%",background:accent,display:"inline-block"}}/>
          {sk}
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {fields.map(f=>(
            <div key={f.key} style={{flex:1,minWidth:130}}>
              <div style={{fontSize:11,color:f.optional?LIGHT:accent,fontWeight:700,marginBottom:4,textTransform:"uppercase",letterSpacing:.5}}>{f.label}{f.optional?" (opcional)":""}</div>
              <select value={cfg[f.key]||""} onChange={e=>upd(sk,f.key,e.target.value)} style={{width:"100%",padding:"6px 8px",borderRadius:7,border:"1px solid "+(cfg[f.key]?accent:BORDER),fontSize:12,background:cfg[f.key]?accent+"10":CARD,color:DARK}}>
                <option value="">-- ignorar --</option>
                {h.map(c=><option key={c} value={c}>{c}{sr&&sr[c]?" ("+String(sr[c]).slice(0,14)+")":""}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>);
    })}
  </Card>);
}

// ─── Enrichment Picker ────────────────────────────────────────────────────────
function EnrichmentPicker({sheetNames,allRows,config,onConfirm,templateCols}){
  const [sheet,setSheet]=useState(config?config.sheet:"");
  const [keyType,setKeyType]=useState(config?config.keyType:"cnpj");
  const [keyCol,setKeyCol]=useState(config?config.keyCol:"");
  const [colMap,setColMap]=useState(config?config.colMap:{});
  const headers=sheet&&allRows[sheet]&&allRows[sheet].length>0?Object.keys(allRows[sheet][0]):[];
  const sampleRow=allRows[sheet]?allRows[sheet][0]:null;
  const ready=sheet&&keyCol&&Object.keys(colMap).some(k=>colMap[k]);
  const KEY_TYPES=[{id:"cnpj",label:"CNPJ",desc:"CNPJ completo"},{id:"raiz",label:"Raiz CNPJ",desc:"8 primeiros dígitos"},{id:"domain",label:"Domínio",desc:"Domínio do site"}];
  return(<>
    <div style={{marginBottom:12}}>
      <div style={{fontSize:11,color:LIGHT,fontWeight:700,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Aba</div>
      <SheetToggle sheetNames={sheetNames} selected={sheet?[sheet]:[]} onToggle={n=>{setSheet(n===sheet?"":n);setKeyCol("");setColMap({});}} color="#d97706"/>
    </div>
    {headers.length>0&&(<>
      <div style={{marginBottom:12}}>
        <div style={{fontSize:11,color:LIGHT,fontWeight:700,marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Chave de cruzamento</div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
          {KEY_TYPES.map(kt=>(
            <div key={kt.id} onClick={()=>setKeyType(kt.id)} style={{padding:"8px 14px",borderRadius:9,border:"2px solid "+(keyType===kt.id?"#d97706":BORDER),background:keyType===kt.id?"#fffbeb":CARD,cursor:"pointer"}}>
              <div style={{fontSize:12,fontWeight:700,color:keyType===kt.id?"#92400e":DARK}}>{kt.label}</div>
              <div style={{fontSize:11,color:LIGHT}}>{kt.desc}</div>
            </div>
          ))}
        </div>
        <SelField lbl={"Coluna de "+KEY_TYPES.find(k=>k.id===keyType).label} val={keyCol} setter={setKeyCol} opts={headers} sampleRow={sampleRow} accent="#d97706"/>
      </div>
      <div style={{marginBottom:14}}>
        <div style={{fontSize:11,color:LIGHT,fontWeight:700,marginBottom:8,textTransform:"uppercase",letterSpacing:.5}}>Mapear colunas externas para o template</div>
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {templateCols.filter(tc=>!["E-MAIL EMPRESAS","TELEFONE","TELEFONE WHATSAPP","EMAIL - PERSONAS","LINKEDIN"].includes(tc)).map(tc=>(
            <div key={tc} style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{width:180,fontSize:12,color:MID,fontWeight:600,flexShrink:0}}>{tc}</span>
              <span style={{fontSize:11,color:LIGHT,marginRight:4}}>←</span>
              <select value={colMap[tc]||""} onChange={e=>setColMap({...colMap,[tc]:e.target.value})} style={{flex:1,padding:"5px 8px",borderRadius:7,border:"1px solid "+(colMap[tc]?"#d97706":BORDER),fontSize:12,color:DARK,background:colMap[tc]?"#fffbeb":CARD}}>
                <option value="">-- não mapear --</option>
                {headers.map(h=><option key={h} value={h}>{h}{sampleRow&&sampleRow[h]?" ("+String(sampleRow[h]).slice(0,18)+")":""}</option>)}
              </select>
            </div>
          ))}
        </div>
      </div>
      <button onClick={()=>{if(ready)onConfirm({sheet,keyType,keyCol,colMap,rows:allRows[sheet]});}} disabled={!ready} style={{padding:"7px 20px",borderRadius:7,border:"none",fontWeight:700,fontSize:12,background:ready?"#d97706":"#e2e8f0",color:ready?"#fff":LIGHT,cursor:ready?"pointer":"default"}}>
        {ready?"Confirmar":"Selecione a aba e mapeie ao menos uma coluna"}
      </button>
      {config&&<span style={{marginLeft:10,fontSize:12,color:"#38a169",fontWeight:700}}>✓ Configurado</span>}
    </>)}
  </>);
}

// ─── Template Builder ─────────────────────────────────────────────────────────
function TemplateBuilder({sheetHeaders,templateCols,setTemplateCols,onNext,onBack}){
  const [dragIdx,setDragIdx]=useState(null);const [newCol,setNewCol]=useState("");
  const allAvailable=[...new Set([...sheetHeaders,...DEFAULT_TEMPLATE_COLS])];
  const isSel=c=>templateCols.includes(c);
  const toggle=c=>{if(isSel(c))setTemplateCols(templateCols.filter(x=>x!==c));else setTemplateCols([...templateCols,c]);};
  const addCustom=()=>{const n=newCol.trim();if(!n||isSel(n))return;setTemplateCols([...templateCols,n]);setNewCol("");};
  const moveUp=i=>{if(i===0)return;const a=[...templateCols];[a[i-1],a[i]]=[a[i],a[i-1]];setTemplateCols(a);};
  const moveDown=i=>{if(i===templateCols.length-1)return;const a=[...templateCols];[a[i+1],a[i]]=[a[i],a[i+1]];setTemplateCols(a);};
  const remove=i=>setTemplateCols(templateCols.filter((_,idx)=>idx!==i));
  const onDragStart=i=>setDragIdx(i);
  const onDragOver=(e,i)=>{e.preventDefault();if(dragIdx===null||dragIdx===i)return;const a=[...templateCols];const it=a.splice(dragIdx,1)[0];a.splice(i,0,it);setTemplateCols(a);setDragIdx(i);};
  const onDragEnd=()=>setDragIdx(null);
  const groups={"Da planilha principal":allAvailable.filter(c=>sheetHeaders.includes(c)&&!GENERATED_COLS.includes(c)),"Geradas pelo app":allAvailable.filter(c=>GENERATED_COLS.includes(c)),"Outras (template padrão)":allAvailable.filter(c=>!sheetHeaders.includes(c)&&!GENERATED_COLS.includes(c))};
  return(<div>
    <Card title="Monte o template do arquivo final" accent="#4f46e5">
      <p style={{fontSize:13,color:MID,marginBottom:16}}>Selecione as colunas e arraste para reordenar.</p>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
        <div>
          <div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10,textTransform:"uppercase",letterSpacing:.5}}>Colunas disponíveis</div>
          <div style={{maxHeight:380,overflowY:"auto",border:"1px solid "+BORDER,borderRadius:10,background:ACCENT_BG,padding:10,display:"flex",flexDirection:"column",gap:3}}>
            {Object.entries(groups).map(([grp,cols])=>cols.length===0?null:(
              <div key={grp}>
                <div style={{fontSize:10,fontWeight:700,color:LIGHT,textTransform:"uppercase",letterSpacing:.5,margin:"6px 0 4px"}}>{grp}</div>
                {cols.map(col=>(
                  <div key={col} onClick={()=>toggle(col)} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",borderRadius:7,cursor:"pointer",background:isSel(col)?"#ebf8ff":CARD,border:"1px solid "+(isSel(col)?"#90cdf4":BORDER),marginBottom:2}}>
                    <div style={{width:16,height:16,borderRadius:4,border:"2px solid "+(isSel(col)?"#3182ce":BORDER),background:isSel(col)?"#3182ce":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                      {isSel(col)&&<svg width="10" height="10" viewBox="0 0 12 12"><polyline points="2,6 5,9 10,3" stroke="#fff" strokeWidth="2" fill="none"/></svg>}
                    </div>
                    <span style={{fontSize:12,color:DARK,fontWeight:isSel(col)?600:400,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{col}</span>
                    {GENERATED_COLS.includes(col)&&<span style={{fontSize:10,background:"#ebf8ff",color:"#2b6cb0",padding:"1px 6px",borderRadius:10,fontWeight:600,flexShrink:0}}>app</span>}
                  </div>
                ))}
              </div>
            ))}
          </div>
          <div style={{marginTop:10,display:"flex",gap:6}}>
            <input value={newCol} onChange={e=>setNewCol(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addCustom()} placeholder="Nova coluna personalizada..." style={{flex:1,padding:"7px 10px",borderRadius:8,border:"1px solid "+BORDER,fontSize:12,color:DARK,background:CARD,outline:"none"}}/>
            <button onClick={addCustom} style={{padding:"7px 14px",borderRadius:8,border:"none",background:ORANGE,color:"#fff",fontSize:12,fontWeight:700,cursor:"pointer"}}>+</button>
          </div>
        </div>
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
            <div style={{fontSize:12,fontWeight:700,color:DARK,textTransform:"uppercase",letterSpacing:.5}}>Ordem final ({templateCols.length})</div>
            <button onClick={()=>setTemplateCols([])} style={{fontSize:11,color:"#e53e3e",background:"none",border:"none",cursor:"pointer",fontWeight:600}}>Limpar</button>
          </div>
          <div style={{maxHeight:380,overflowY:"auto",border:"1px solid "+BORDER,borderRadius:10,background:ACCENT_BG,padding:10,display:"flex",flexDirection:"column",gap:4}}>
            {templateCols.length===0?<div style={{textAlign:"center",padding:"40px 0",color:LIGHT,fontSize:13}}>Nenhuma coluna selecionada</div>:
              templateCols.map((col,i)=>(
                <div key={col+i} draggable onDragStart={()=>onDragStart(i)} onDragOver={e=>onDragOver(e,i)} onDragEnd={onDragEnd}
                  style={{display:"flex",alignItems:"center",gap:6,padding:"6px 10px",borderRadius:7,background:dragIdx===i?"#ebf8ff":CARD,border:"1px solid "+(dragIdx===i?"#90cdf4":BORDER),cursor:"grab",userSelect:"none"}}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke={LIGHT} strokeWidth="2" style={{flexShrink:0}}><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                  <span style={{fontSize:11,color:LIGHT,fontWeight:600,width:18,flexShrink:0}}>{i+1}</span>
                  <span style={{fontSize:12,color:DARK,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{col}</span>
                  {GENERATED_COLS.includes(col)&&<span style={{fontSize:10,background:"#ebf8ff",color:"#2b6cb0",padding:"1px 6px",borderRadius:10,fontWeight:600,flexShrink:0}}>app</span>}
                  <div style={{display:"flex",gap:2,flexShrink:0}}>
                    <button onClick={()=>moveUp(i)} disabled={i===0} style={{padding:"2px 5px",borderRadius:4,border:"1px solid "+BORDER,background:CARD,cursor:i===0?"default":"pointer",color:i===0?BORDER:MID,fontSize:10}}>↑</button>
                    <button onClick={()=>moveDown(i)} disabled={i===templateCols.length-1} style={{padding:"2px 5px",borderRadius:4,border:"1px solid "+BORDER,background:CARD,cursor:i===templateCols.length-1?"default":"pointer",color:i===templateCols.length-1?BORDER:MID,fontSize:10}}>↓</button>
                    <button onClick={()=>remove(i)} style={{padding:"2px 5px",borderRadius:4,border:"1px solid #fed7d7",background:"#fff5f5",cursor:"pointer",color:"#e53e3e",fontSize:10}}>✕</button>
                  </div>
                </div>
              ))
            }
          </div>
          <div style={{marginTop:10,display:"flex",gap:6}}>
            <button onClick={()=>setTemplateCols([...DEFAULT_TEMPLATE_COLS])} style={{flex:1,padding:"7px 10px",borderRadius:8,border:"1px solid "+BORDER,background:CARD,fontSize:12,fontWeight:600,color:MID,cursor:"pointer"}}>Template padrão</button>
            <button onClick={()=>setTemplateCols(sheetHeaders)} style={{flex:1,padding:"7px 10px",borderRadius:8,border:"1px solid "+BORDER,background:CARD,fontSize:12,fontWeight:600,color:MID,cursor:"pointer"}}>Todas da planilha</button>
          </div>
        </div>
      </div>
    </Card>
    <div style={{display:"flex",gap:10}}>
      <button onClick={onBack} style={{padding:"11px 22px",borderRadius:9,border:"1px solid "+BORDER,background:CARD,cursor:"pointer",fontSize:14,fontWeight:600,color:MID}}>Voltar</button>
      <button onClick={onNext} disabled={templateCols.length===0} style={{flex:1,padding:"11px",borderRadius:9,border:"none",fontWeight:700,fontSize:14,background:templateCols.length>0?DARK:"#e2e8f0",color:templateCols.length>0?"#fff":LIGHT,cursor:templateCols.length>0?"pointer":"default"}}>Configurar fontes →</button>
    </div>
  </div>);
}

// ─── TPV Modal ────────────────────────────────────────────────────────────────
function TPVModal({result,onClose,onDownload}){
  const [tpvResult,setTpvResult]=useState(null);
  const TC={"Tier IV":"#e53e3e","Tier III":"#dd6b20","Tier II":"#38a169","Tier I":"#3182ce","Micro":"#805ad5"};
  const run=()=>{const tiers={"Tier IV":0,"Tier III":0,"Tier II":0,"Tier I":0,"Micro":0};const rows=result.map(row=>{const{tier}=calcTPV(row);if(tier)tiers[tier]=(tiers[tier]||0)+1;return{...row,"TPV ESTIMADO":tier,"TIER":tier};});setTpvResult({rows,tiers});};
  return(<div style={{position:"fixed",inset:0,background:"rgba(26,35,64,0.4)",zIndex:200,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
    <div style={{background:CARD,borderRadius:14,width:"100%",maxWidth:660,maxHeight:"90vh",overflowY:"auto",padding:28,boxShadow:"0 20px 60px rgba(0,0,0,0.15)"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:20}}>
        <div><div style={{fontWeight:800,fontSize:18,color:DARK}}>Calcular Tier Estimado</div><div style={{fontSize:13,color:LIGHT,marginTop:2}}>Metodologia de 9 camadas · tabela CNAE embutida</div></div>
        <button onClick={onClose} style={{background:ACCENT_BG,border:"1px solid "+BORDER,borderRadius:8,width:32,height:32,cursor:"pointer",color:MID,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
      </div>
      {!tpvResult?(<>
        <div style={{background:ACCENT_BG,borderRadius:10,padding:"14px 16px",marginBottom:20,fontSize:12,color:MID,border:"1px solid "+BORDER}}>
          <div style={{fontWeight:700,marginBottom:6,color:DARK}}>Hierarquia do Ticket Médio</div>
          <div>1. Coluna TICKET MÉDIO da lista (R$30–R$15k)</div>
          <div>2. Ticket médio da tabela CNAE por segmento</div>
          <div>3. Padrão: R$ 150</div>
        </div>
        <button onClick={run} style={{width:"100%",padding:"13px",borderRadius:9,border:"none",background:DARK,color:"#fff",fontWeight:700,fontSize:15,cursor:"pointer"}}>Calcular para {result.length} CNPJs</button>
      </>):(<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:8,marginBottom:20}}>
          {Object.entries(TC).map(([t,c])=>(<div key={t} style={{background:ACCENT_BG,borderRadius:10,padding:"12px 8px",border:"1px solid "+BORDER,textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:c}}>{tpvResult.tiers[t]||0}</div><div style={{fontSize:11,color:LIGHT,fontWeight:600}}>{t}</div></div>))}
        </div>
        <div style={{overflowX:"auto",marginBottom:20}}>
          <table style={{borderCollapse:"collapse",fontSize:11,width:"100%"}}>
            <thead><tr style={{background:ACCENT_BG}}>{["CNPJ","RAZÃO SOCIAL","CNAE SUBCLASSE","TICKET MÉDIO","TIER"].map(h=>(<th key={h} style={{padding:"8px 10px",border:"1px solid "+BORDER,textAlign:"left",fontWeight:700,color:h==="TIER"?ORANGE:MID,fontSize:10,whiteSpace:"nowrap"}}>{h}</th>))}</tr></thead>
            <tbody>{tpvResult.rows.slice(0,10).map((row,i)=>(<tr key={i} style={{background:i%2===0?CARD:ACCENT_BG}}>{["CNPJ","RAZÃO SOCIAL","CNAE SUBCLASSE","TICKET MÉDIO","TIER"].map(tc=>(<td key={tc} style={{padding:"7px 10px",border:"1px solid "+BORDER,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:tc==="TIER"?ORANGE:DARK,fontWeight:tc==="TIER"?700:400}}>{row[tc]||""}</td>))}</tr>))}</tbody>
          </table>
        </div>
        <button onClick={()=>onDownload(tpvResult.rows)} style={{width:"100%",padding:"12px",borderRadius:9,border:"none",background:DARK,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer"}}>Baixar .xlsx com Tier</button>
      </>)}
    </div>
  </div>);
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App(){
  const [uploadedFiles,setUploadedFiles]=useState([]);
  const [allRows,setAllRows]=useState({});
  const [sheetNames,setSheetNames]=useState([]);
  const [sheetHeaders,setSheetHeaders]=useState([]);
  const [fileName,setFileName]=useState("");
  const [step,setStep]=useState(1);
  const [templateCols,setTemplateCols]=useState([...DEFAULT_TEMPLATE_COLS]);

  // Cadastral
  const [cadSheets,setCadSheets]=useState([]);
  const [cadMapping,setCadMapping]=useState({});

  // Contact sheets (from main files)
  const [emailSheets,setEmailSheets]=useState([]);
  const [phoneSheets,setPhoneSheets]=useState([]);
  const [emailColConfigs,setEmailColConfigs]=useState({});
  const [phoneColConfigs,setPhoneColConfigs]=useState({});

  // Personas/LinkedIn — external file
  const [personasFile,setPersonasFile]=useState(null);
  const [personasSheets,setPersonasSheets]=useState([]);
  const [personasColConfigs,setPersonasColConfigs]=useState({});
  const personasFileRef=useRef();

  // Enrichments
  const [enrichments,setEnrichments]=useState([]);

  const [result,setResult]=useState([]);
  const [metrics,setMetrics]=useState(null);
  const [showTPV,setShowTPV]=useState(false);
  const [expandMode,setExpandMode]=useState(false);
  const fileRef=useRef();

  // Derived
  const cadHeaders=[...new Set(cadSheets.flatMap(sk=>{const rows=allRows[sk]||[];return rows.length>0?Object.keys(rows[0]):[];}))];
  const cadSample=cadSheets.length>0?(allRows[cadSheets[0]]||[])[0]:null;
  const manualCols=templateCols.filter(c=>!GENERATED_COLS.includes(c));


  const readFile=(file,cb)=>{const r=new FileReader();r.onload=e=>{const wb=XLSX.read(e.target.result,{type:"binary"});const p={};wb.SheetNames.forEach(n=>{p[n]=XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""});});cb(wb.SheetNames,p,file.name);};r.readAsBinaryString(file);};

  const rebuildMerge=(all)=>{
    const mergedRows={};const mergedSheets=[];
    all.forEach(f=>{f.sheetNames.forEach(sn=>{const key=all.length>1?f.name.replace(/\.[^/.]+$/,"")+" · "+sn:sn;mergedRows[key]=f.allRows[sn];mergedSheets.push(key);});});
    const allHeaders=[...new Set(all.flatMap(f=>f.sheetNames.flatMap(sn=>{const rows=f.allRows[sn]||[];return rows.length>0?Object.keys(rows[0]):[]})))];
    return{mergedRows,mergedSheets,allHeaders};
  };

  const handleFile=e=>{
    const files=Array.from(e.target.files);if(!files.length)return;
    let pending=files.length;const newLoaded=[];
    files.forEach(file=>{readFile(file,(ns,p,n)=>{newLoaded.push({name:n,allRows:p,sheetNames:ns});pending--;if(pending===0){
      setUploadedFiles(prev=>{
        const all=[...prev,...newLoaded];
        const{mergedRows,mergedSheets,allHeaders}=rebuildMerge(all);
        setAllRows(mergedRows);setSheetNames(mergedSheets);setSheetHeaders(allHeaders);
        setFileName(all.length===1?all[0].name:all.length+" arquivos carregados");
        setCadSheets([]);setCadMapping({});setEmailSheets([]);setPhoneSheets([]);setPersonasSheets([]);setEmailColConfigs({});setPhoneColConfigs({});setPersonasColConfigs({});setResult([]);setMetrics(null);
        return all;
      });
    }});});
    e.target.value="";
  };

  const removeFile=(i)=>{
    const next=uploadedFiles.filter((_,idx)=>idx!==i);
    setUploadedFiles(next);
    if(next.length===0){setAllRows({});setSheetNames([]);setSheetHeaders([]);setFileName("");}
    else{const{mergedRows,mergedSheets,allHeaders}=rebuildMerge(next);setAllRows(mergedRows);setSheetNames(mergedSheets);setSheetHeaders(allHeaders);setFileName(next.length===1?next[0].name:next.length+" arquivos");}
  };

  const normRaiz=v=>v.replace(/\D/g,"").padStart(8,"0").slice(0,8);
  const normPlat=v=>{if(!v)return"";let s=String(v).trim().replace(/^\{|\}$/g,"").trim();const l=s.toLowerCase();if(l.includes("cart func")||l.includes("virtual store redirect"))return"Desenvolvimento Próprio";return s;};

  const handleProcess=()=>{
    const emailMap={},phoneMap={},whatsMap={},personasMap={},sc={},cc={};

    emailSheets.forEach(sk=>{const cfg=emailColConfigs[sk]||{};if(!cfg.cnpjCol||!cfg.dataCol)return;(allRows[sk]||[]).forEach(row=>{const cnpj=String(row[cfg.cnpjCol]||"").trim();if(!cnpj)return;const val=String(row[cfg.dataCol]||"").trim();if(!val)return;if(!emailMap[cnpj])emailMap[cnpj]=[];val.split(",").map(v=>v.trim()).filter(Boolean).forEach(v=>{if(!emailMap[cnpj].includes(v))emailMap[cnpj].push(v);});});});

    phoneSheets.forEach(sk=>{const cfg=phoneColConfigs[sk]||{};if(!cfg.cnpjCol||!cfg.dataCol)return;(allRows[sk]||[]).forEach(row=>{const cnpj=String(row[cfg.cnpjCol]||"").trim();if(!cnpj)return;const val=String(row[cfg.dataCol]||"").trim();if(!val)return;const wa=cfg.waCol?String(row[cfg.waCol]||"").trim().toLowerCase():"";const isWa=wa==="sim";const map=cfg.waCol?(isWa?whatsMap:phoneMap):phoneMap;if(!map[cnpj])map[cnpj]=[];val.split(",").map(v=>v.trim()).filter(Boolean).forEach(v=>{if(!map[cnpj].includes(v))map[cnpj].push(v);});if(cfg.statusCol){const s=String(row[cfg.statusCol]||"").trim();if(s)sc[s]=(sc[s]||0)+1;}if(cfg.catCol){const c=String(row[cfg.catCol]||"").trim();if(c)cc[c]=(cc[c]||0)+1;}});});

    personasSheets.forEach(sk=>{
      const rows = personasFile ? (personasFile.rows[sk]||[]) : [];
      const cfg=personasColConfigs[sk]||{};if(!cfg.cnpjCol)return;
      rows.forEach(row=>{
        const cnpj=String(row[cfg.cnpjCol]||"").trim();if(!cnpj)return;
        if(!personasMap[cnpj])personasMap[cnpj]={pairs:[]};
        const email=cfg.emailCol?String(row[cfg.emailCol]||"").trim():"";
        const linkedin=cfg.linkedinCol?String(row[cfg.linkedinCol]||"").trim():"";
        if(email||linkedin) personasMap[cnpj].pairs.push({email,linkedin});
      });
    });

    const enrichMaps=enrichments.filter(e=>e.config).map(e=>{const{keyType,keyCol,colMap,rows}=e.config;const map={};rows.forEach(row=>{let key=String(row[keyCol]||"").trim();if(keyType==="raiz")key=key.replace(/\D/g,"").padStart(8,"0").slice(0,8);else if(keyType==="domain")key=key.toLowerCase();if(!key)return;map[key]={};Object.entries(colMap).forEach(([tc,sc2])=>{if(sc2)map[key][tc]=String(row[sc2]||"").trim();});});return{map,keyType};});

    const cnpjKey=cadMapping["CNPJ"]||cadMapping["RAIZ CNPJ"];
    const cadByCnpj={};
    cadSheets.forEach(sk=>{(allRows[sk]||[]).forEach(row=>{const c=String(row[cnpjKey]||"").trim();if(c&&!cadByCnpj[c])cadByCnpj[c]=row;});});

    const allCnpjs=[...new Set([...Object.keys(cadByCnpj),...Object.keys(emailMap),...Object.keys(phoneMap),...Object.keys(whatsMap),...Object.keys(personasMap)])].sort();

    const rows=allCnpjs.map(cnpj=>{
      const cad=cadByCnpj[cnpj]||{};const raiz=normRaiz(cnpj);const out={};
      templateCols.forEach(tc=>{
        if(tc==="E-MAIL EMPRESAS")out[tc]=(emailMap[cnpj]||[]).join(", ");
        else if(tc==="TELEFONE")out[tc]=(phoneMap[cnpj]||[]).join(", ");
        else if(tc==="TELEFONE WHATSAPP")out[tc]=(whatsMap[cnpj]||[]).join(", ");
        else if(tc==="EMAIL - PERSONAS")out[tc]=personasMap[cnpj]?personasMap[cnpj].pairs.map(p=>p.email).filter(Boolean).join(", "):(cadMapping[tc]?String(cad[cadMapping[tc]]||""):"");
        else if(tc==="LINKEDIN")out[tc]=personasMap[cnpj]?personasMap[cnpj].pairs.map(p=>p.linkedin).filter(Boolean).join(", "):(cadMapping[tc]?String(cad[cadMapping[tc]]||""):"");
        else if(tc==="PLATAFORMA")out[tc]=normPlat(cadMapping[tc]?String(cad[cadMapping[tc]]||""):"");
        else out[tc]=cadMapping[tc]?String(cad[cadMapping[tc]]||""):"";
      });
      // Apply enrichments
      enrichMaps.forEach(({map,keyType})=>{
        const lk=keyType==="raiz"?raiz:keyType==="domain"?(out["DOMINIO"]||"").toLowerCase():cnpj;
        const ed=map[lk];if(!ed)return;
        Object.entries(ed).forEach(([tc,val])=>{if(val)out[tc]=tc==="PLATAFORMA"?normPlat(val):val;});
      });
      return out;
    });

    const tc=rows.length,we=rows.filter(r=>r["E-MAIL EMPRESAS"]).length,wp=rows.filter(r=>r["TELEFONE"]).length,ww=rows.filter(r=>r["TELEFONE WHATSAPP"]).length,wd=rows.filter(r=>r["DOMINIO"]).length,wpe=rows.filter(r=>r["EMAIL - PERSONAS"]).length,wli=rows.filter(r=>r["LINKEDIN"]).length;
    const tpr=phoneSheets.reduce((a,sk)=>a+(allRows[sk]||[]).length,0);
    const ivE=e=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e.trim());
    const ae=rows.flatMap(r=>r["E-MAIL EMPRESAS"]?r["E-MAIL EMPRESAS"].split(",").map(v=>v.trim()):[]);
    const ve=ae.filter(ivE);
    const stC={},plC={},cnC={};
    rows.forEach(r=>{if(r["ESTADO"])stC[r["ESTADO"]]=(stC[r["ESTADO"]]||0)+1;if(r["PLATAFORMA"])plC[r["PLATAFORMA"]]=(plC[r["PLATAFORMA"]]||0)+1;if(r["CNAE SUBCLASSE"])cnC[r["CNAE SUBCLASSE"]]=(cnC[r["CNAE SUBCLASSE"]]||0)+1;});
    setMetrics({totalCnpjs:tc,withEmail:we,withPhone:wp,withWhats:ww,withDomain:wd,withPersonas:wpe,withLinkedin:wli,totalPhoneRows:tpr,statusCount:sc,catCount:cc,totalEmails:ae.length,validEmails:ve.length,topStates:Object.entries(stC).sort((a,b)=>b[1]-a[1]).slice(0,5),topPlatforms:Object.entries(plC).sort((a,b)=>b[1]-a[1]).slice(0,5),topCnaes:Object.entries(cnC).sort((a,b)=>b[1]-a[1]).slice(0,5)});
    setResult(rows);setStep(4);
  };

  const doDownload=(rows,filename)=>{
    const finalRows=expandMode?expandRows(rows):rows;
    const cols=finalRows[0]&&finalRows[0]["TIER"]?[...templateCols,"TPV ESTIMADO","TIER"]:templateCols;
    const ws=XLSX.utils.json_to_sheet(finalRows,{header:cols});ws["!cols"]=cols.map(()=>({wch:25}));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Resultado");XLSX.writeFile(wb,filename||"planilha_formatada.xlsx");
  };

  const reset=()=>{setUploadedFiles([]);setAllRows({});setSheetNames([]);setSheetHeaders([]);setFileName("");setCadSheets([]);setCadMapping({});setEnrichments([]);setEmailSheets([]);setPhoneSheets([]);setPersonasFile(null);setPersonasSheets([]);setEmailColConfigs({});setPhoneColConfigs({});setPersonasColConfigs({});setResult([]);setMetrics(null);setShowTPV(false);setExpandMode(false);setTemplateCols([...DEFAULT_TEMPLATE_COLS]);setStep(1);if(fileRef.current)fileRef.current.value="";if(personasFileRef.current)personasFileRef.current.value="";};

  const addEnrichment=()=>setEnrichments(prev=>[...prev,{id:Date.now(),file:null,config:null}]);
  const removeEnrichment=id=>setEnrichments(prev=>prev.filter(e=>e.id!==id));
  const setEnrichFile=(id,fd)=>setEnrichments(prev=>prev.map(e=>e.id===id?{...e,file:fd,config:null}:e));
  const setEnrichConfig=(id,cfg)=>setEnrichments(prev=>prev.map(e=>e.id===id?{...e,config:cfg}:e));
  const expandCount=result.length>0?expandRows(result).length:0;

  return(
    <div style={{fontFamily:"Inter,system-ui,sans-serif",minHeight:"100vh",background:BG,paddingBottom:56}}>
      {showTPV&&<TPVModal result={result} onClose={()=>setShowTPV(false)} onDownload={rows=>{doDownload(rows,"planilha_com_tier.xlsx");setShowTPV(false);}}/>}

      {/* Nav */}
      <div style={{background:CARD,borderBottom:"1px solid "+BORDER,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:56,boxShadow:"0 1px 3px rgba(0,0,0,0.04)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <svg width="22" height="22" viewBox="0 0 32 32"><polygon points="4,28 16,4 28,28 16,20" fill={ORANGE}/></svg>
          <span style={{fontWeight:700,fontSize:15,color:DARK}}>LeadLab</span>
        </div>
      </div>

      <div style={{maxWidth:900,margin:"0 auto",padding:"32px 16px"}}>
        {/* Steps */}
        <div style={{display:"flex",gap:4,marginBottom:28}}>
          {["Upload","Template","Configurar","Resultado"].map((s,i)=>(<div key={i} style={{flex:1,textAlign:"center",padding:"9px 4px",borderRadius:8,fontSize:12,fontWeight:600,background:step===i+1?DARK:step>i+1?"#ebf8ff":CARD,color:step===i+1?"#fff":step>i+1?"#2b6cb0":LIGHT,border:"1px solid "+(step===i+1?DARK:step>i+1?"#bee3f8":BORDER)}}>{step>i+1?"✓ ":((i+1)+". ")}{s}</div>))}
        </div>

        {/* Step 1 — Upload */}
        {step===1&&(<Card><div style={{padding:"32px 24px"}}>
          <div style={{textAlign:"center",marginBottom:28}}>
            <div style={{width:72,height:72,borderRadius:16,background:ORANGE+"12",margin:"0 auto 16px",display:"flex",alignItems:"center",justifyContent:"center"}}><svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke={ORANGE} strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg></div>
            <h2 style={{margin:"0 0 6px",fontSize:20,fontWeight:800,color:DARK}}>Selecione sua(s) planilha(s)</h2>
            <p style={{color:LIGHT,fontSize:13,margin:0}}>Adicione um ou mais arquivos. Eles serão concatenados antes de processar.</p>
          </div>
          {uploadedFiles.length>0&&(
            <div style={{marginBottom:16,display:"flex",flexDirection:"column",gap:6}}>
              {uploadedFiles.map((f,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderRadius:9,background:ACCENT_BG,border:"1px solid "+BORDER}}>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#38a169" strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                  <div style={{flex:1}}>
                    <div style={{fontSize:13,fontWeight:600,color:DARK}}>{f.name}</div>
                    <div style={{fontSize:11,color:LIGHT}}>{f.sheetNames.length} aba(s) · {Object.values(f.allRows)[0]?.length||0} linhas</div>
                  </div>
                  <button onClick={()=>removeFile(i)} style={{fontSize:11,color:"#e53e3e",background:"none",border:"none",cursor:"pointer",fontWeight:700,padding:"4px 8px"}}>Remover</button>
                </div>
              ))}
              <div style={{fontSize:12,color:"#38a169",fontWeight:600,padding:"4px 0"}}>✓ {uploadedFiles.reduce((a,f)=>a+Object.values(f.allRows).reduce((b,r)=>b+r.length,0),0)} linhas no total</div>
            </div>
          )}
          <div style={{display:"flex",gap:10}}>
            <div style={{flex:1}}>
              <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}} id="fi"/>
              <label htmlFor="fi" style={{display:"block",textAlign:"center",padding:"11px",borderRadius:9,border:"2px dashed "+BORDER,cursor:"pointer",fontSize:13,fontWeight:600,color:MID,background:ACCENT_BG}}>
                + {uploadedFiles.length===0?"Escolher arquivo":"Adicionar outro arquivo"}
              </label>
            </div>
            {uploadedFiles.length>0&&(<button onClick={()=>setStep(2)} style={{padding:"11px 28px",borderRadius:9,border:"none",background:DARK,color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>Continuar →</button>)}
          </div>
        </div></Card>)}

        {/* Step 2 — Template */}
        {step===2&&<TemplateBuilder sheetHeaders={sheetHeaders} templateCols={templateCols} setTemplateCols={setTemplateCols} onNext={()=>setStep(3)} onBack={()=>setStep(1)}/>}

        {/* Step 3 — Configurar */}
        {step===3&&(<>
          <div style={{marginBottom:16}}>
            <div style={{fontSize:13,color:MID,marginBottom:uploadedFiles.length>1?8:0}}><b style={{color:DARK}}>{fileName}</b> · {sheetNames.length} aba(s) · <b style={{color:"#4f46e5"}}>{templateCols.length} colunas no template</b></div>
            {uploadedFiles.length>1&&(<div style={{display:"flex",flexWrap:"wrap",gap:6}}>{uploadedFiles.map((f,i)=>(<span key={i} style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 10px",borderRadius:20,background:ACCENT_BG,border:"1px solid "+BORDER,fontSize:11,color:MID}}><svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke={ORANGE} strokeWidth="2"><path d="M14 2H6a2 2 0 0 0-2 2v16h12V8z"/><polyline points="14 2 14 8 20 8"/></svg>{f.name}</span>))}</div>)}
          </div>

          <Card title="Dados Cadastrais" accent="#4f46e5">
            <p style={{fontSize:13,color:MID,margin:"0 0 12px"}}>Selecione uma ou mais abas com dados cadastrais. Em caso de CNPJ duplicado, a primeira aba tem prioridade.</p>
            <SheetToggle sheetNames={sheetNames} selected={cadSheets} onToggle={sk=>{setCadSheets(prev=>prev.includes(sk)?prev.filter(s=>s!==sk):[...prev,sk]);setCadMapping({});}} color="#4f46e5"/>
            {cadSheets.length>1&&(<div style={{background:"#eff6ff",borderRadius:8,padding:"8px 12px",marginBottom:12,fontSize:12,color:"#1e40af",border:"1px solid #bfdbfe"}}>{cadSheets.length} abas selecionadas — dados serão combinados pelo CNPJ.</div>)}
            {cadHeaders.length>0&&<ColMapper headers={cadHeaders} sampleRow={cadSample} mapping={cadMapping} onChange={setCadMapping} manualCols={manualCols}/>}
          </Card>

          {/* Enrichments */}
          <Card title="Planilhas de enriquecimento (opcional)" accent="#d97706">
            <p style={{fontSize:13,color:MID,margin:"0 0 14px"}}>Adicione planilhas externas para enriquecer os dados. Escolha a chave de cruzamento e mapeie as colunas livremente.</p>
            {enrichments.map((enr,idx)=>(
              <div key={enr.id} style={{background:ACCENT_BG,borderRadius:10,padding:16,marginBottom:12,border:"1px solid "+BORDER}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <span style={{fontWeight:700,fontSize:13,color:DARK}}>Planilha {idx+1}{enr.file?" · "+enr.file.name:""}</span>
                  <button onClick={()=>removeEnrichment(enr.id)} style={{fontSize:11,color:"#e53e3e",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Remover</button>
                </div>
                {!enr.file?(
                  <>
                    <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} id={"ef-"+enr.id}
                      onChange={e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{const wb=XLSX.read(ev.target.result,{type:"binary"});const p={};wb.SheetNames.forEach(n=>{p[n]=XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""});});setEnrichFile(enr.id,{rows:p,sheetNames:wb.SheetNames,name:f.name});};r.readAsBinaryString(f);}}/>
                    <label htmlFor={"ef-"+enr.id} style={{background:"#d97706",color:"#fff",padding:"7px 18px",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>Escolher arquivo</label>
                  </>
                ):(
                  <EnrichmentPicker sheetNames={enr.file.sheetNames} allRows={enr.file.rows} config={enr.config} templateCols={templateCols} onConfirm={cfg=>setEnrichConfig(enr.id,cfg)}/>
                )}
              </div>
            ))}
            <button onClick={addEnrichment} style={{padding:"8px 18px",borderRadius:8,border:"2px dashed #d97706",background:"transparent",color:"#d97706",fontSize:13,fontWeight:700,cursor:"pointer",width:"100%"}}>+ Adicionar planilha de enriquecimento</button>
          </Card>

          <MultiSheetContactPicker title="E-mails de Empresas" accent="#3182ce" sheetNames={sheetNames} allRows={allRows} selectedSheets={emailSheets} setSelectedSheets={setEmailSheets} colConfigs={emailColConfigs} setColConfigs={setEmailColConfigs} fields={[{key:"cnpjCol",label:"CNPJ"},{key:"dataCol",label:"Coluna de E-mail"}]}/>
          <MultiSheetContactPicker title="Telefones" accent="#38a169" sheetNames={sheetNames} allRows={allRows} selectedSheets={phoneSheets} setSelectedSheets={setPhoneSheets} colConfigs={phoneColConfigs} setColConfigs={setPhoneColConfigs} fields={[{key:"cnpjCol",label:"CNPJ"},{key:"dataCol",label:"Coluna de Telefone"},{key:"waCol",label:"WhatsApp (Sim/Não)",optional:true},{key:"statusCol",label:"Status",optional:true},{key:"catCol",label:"Categoria",optional:true}]}/>

          {/* Personas/LinkedIn — external file */}
          <Card title="E-mail Personas e LinkedIn (planilha externa)" accent="#805ad5">
            {!personasFile?(
              <>
                <p style={{fontSize:13,color:MID,margin:"0 0 12px"}}>Suba a planilha com e-mails de personas e LinkedIn. Chave: CNPJ.</p>
                <input ref={personasFileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} id="personasFile"
                  onChange={e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=ev=>{const wb=XLSX.read(ev.target.result,{type:"binary"});const p={};wb.SheetNames.forEach(n=>{p[n]=XLSX.utils.sheet_to_json(wb.Sheets[n],{defval:""});});setPersonasFile({rows:p,sheetNames:wb.SheetNames,name:f.name});setPersonasSheets([]);setPersonasColConfigs({});};r.readAsBinaryString(f);e.target.value="";}}/>
                <label htmlFor="personasFile" style={{background:"#805ad5",color:"#fff",padding:"7px 18px",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>Escolher arquivo</label>
              </>
            ):(
              <>
                <p style={{fontSize:13,color:MID,margin:"0 0 12px"}}>
                  Arquivo: <b style={{color:DARK}}>{personasFile.name}</b>
                  <button onClick={()=>{setPersonasFile(null);setPersonasSheets([]);setPersonasColConfigs({});if(personasFileRef.current)personasFileRef.current.value="";}} style={{marginLeft:8,fontSize:11,color:"#e53e3e",background:"none",border:"none",cursor:"pointer",fontWeight:700}}>Remover</button>
                </p>
                <MultiSheetContactPicker title="" accent="#805ad5" sheetNames={personasFile.sheetNames} allRows={personasFile.rows} selectedSheets={personasSheets} setSelectedSheets={setPersonasSheets} colConfigs={personasColConfigs} setColConfigs={setPersonasColConfigs} fields={[{key:"cnpjCol",label:"CNPJ"},{key:"emailCol",label:"E-mail Persona",optional:true},{key:"linkedinCol",label:"LinkedIn",optional:true}]}/>
              </>
            )}
          </Card>

          {(emailSheets.length>0||phoneSheets.length>0||personasFile)&&(
            <div style={{background:"#f0fff4",border:"1px solid #c6f6d5",borderRadius:10,padding:"12px 16px",marginBottom:16,fontSize:12,color:"#276749"}}>
              {emailSheets.length>0&&<div>✓ E-mails: {emailSheets.length} aba(s)</div>}
              {phoneSheets.length>0&&<div>✓ Telefones: {phoneSheets.length} aba(s)</div>}
              {personasFile&&personasSheets.length>0&&<div>✓ Personas/LinkedIn: {personasSheets.length} aba(s) de <b>{personasFile.name}</b></div>}
            </div>
          )}

          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setStep(2)} style={{padding:"11px 22px",borderRadius:9,border:"1px solid "+BORDER,background:CARD,cursor:"pointer",fontSize:14,fontWeight:600,color:MID}}>← Template</button>
            <button onClick={handleProcess} style={{flex:1,padding:"11px",borderRadius:9,border:"none",fontWeight:700,fontSize:14,background:DARK,color:"#fff",cursor:"pointer"}}>Processar lista →</button>
          </div>
        </>)}

        {/* Step 4 — Resultado */}
        {step===4&&(<>
          <div style={{background:CARD,borderRadius:12,padding:"16px 20px",marginBottom:20,display:"flex",gap:14,alignItems:"center",border:"1px solid #c6f6d5",boxShadow:"0 1px 3px rgba(0,0,0,0.04)"}}>
            <div style={{width:44,height:44,borderRadius:10,background:"#f0fff4",display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>✅</div>
            <div><div style={{fontWeight:800,color:DARK,fontSize:15}}>Processamento concluído!</div><div style={{fontSize:13,color:LIGHT}}>{result.length} CNPJs · {templateCols.length} colunas no template</div></div>
          </div>

          {/* Créditos */}
          {(()=>{
            const tT=result.reduce((a,r)=>a+(r["TELEFONE"]?r["TELEFONE"].split(",").filter(v=>v.trim()).length:0)+(r["TELEFONE WHATSAPP"]?r["TELEFONE WHATSAPP"].split(",").filter(v=>v.trim()).length:0),0);
            const tEE=result.reduce((a,r)=>a+(r["E-MAIL EMPRESAS"]?r["E-MAIL EMPRESAS"].split(",").filter(v=>v.trim()).length:0),0);
            const tEP=result.reduce((a,r)=>a+(r["EMAIL - PERSONAS"]?r["EMAIL - PERSONAS"].split(",").filter(v=>v.trim()).length:0),0);
            const tLI=result.reduce((a,r)=>a+(r["LINKEDIN"]?r["LINKEDIN"].split(",").filter(v=>v.trim()).length:0),0);
            const tDC=result.filter(r=>r["RAZÃO SOCIAL"]||r["CNPJ"]).length;
            const cT=tT*6,cEE=tEE,cEP=tEP,cLI=tLI,cDC=tDC,total=cT+cEE+cEP+cLI+cDC;
            const items=[{label:"Dados cadastrais",qty:tDC,cred:cDC,unit:1,color:"#4f46e5",icon:"🏢"},{label:"Telefones",qty:tT,cred:cT,unit:6,color:"#38a169",icon:"📞"},{label:"E-mails empresa",qty:tEE,cred:cEE,unit:1,color:"#3182ce",icon:"✉️"},{label:"E-mails persona",qty:tEP,cred:cEP,unit:1,color:"#805ad5",icon:"👤"},{label:"LinkedIn",qty:tLI,cred:cLI,unit:1,color:"#0077b5",icon:"💼"}];
            return(<Card title="Relatório de Créditos" accent={ORANGE}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",background:"#fff8f0",borderRadius:10,padding:"16px 20px",marginBottom:16,border:"1px solid #fbd38d"}}>
                <div><div style={{fontSize:11,color:"#c05621",fontWeight:700,textTransform:"uppercase",letterSpacing:.5,marginBottom:2}}>Total de Créditos Estimados</div><div style={{fontSize:34,fontWeight:800,color:"#c05621"}}>{total.toLocaleString("pt-BR")}</div><div style={{fontSize:12,color:"#9c4221"}}>{result.length} empresas · {tT} telefones · {tEE+tEP} e-mails</div></div>
                <div style={{fontSize:44}}>🪙</div>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {items.map(item=>(<div key={item.label} style={{display:"flex",alignItems:"center",gap:12,background:ACCENT_BG,borderRadius:9,padding:"10px 14px",border:"1px solid "+BORDER}}>
                  <span style={{fontSize:18,width:28,textAlign:"center"}}>{item.icon}</span>
                  <div style={{flex:1}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><span style={{fontSize:13,fontWeight:700,color:DARK}}>{item.label}</span><span style={{fontSize:13,fontWeight:800,color:item.color}}>{item.cred.toLocaleString("pt-BR")} créditos</span></div>
                    <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:11,color:LIGHT}}>{item.qty.toLocaleString("pt-BR")} registros × {item.unit} crédito{item.unit>1?"s":""}</span><span style={{fontSize:11,color:LIGHT}}>{total>0?Math.round(item.cred/total*100):0}% do total</span></div>
                    <Bar pct={total>0?Math.round(item.cred/total*100):0} color={item.color}/>
                  </div>
                </div>))}
              </div>
            </Card>);
          })()}

          {/* Qualidade */}
          {metrics&&(<Card title="Relatório de Qualidade" accent={ORANGE}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(100px,1fr))",gap:10,marginBottom:18}}>
              {[{label:"CNPJs",value:metrics.totalCnpjs,color:"#4f46e5"},{label:"Com E-mail",value:metrics.withEmail,pct:Math.round(metrics.withEmail/metrics.totalCnpjs*100),color:"#3182ce"},{label:"Com Telefone",value:metrics.withPhone,pct:Math.round(metrics.withPhone/metrics.totalCnpjs*100),color:"#38a169"},{label:"Com WhatsApp",value:metrics.withWhats,pct:Math.round(metrics.withWhats/metrics.totalCnpjs*100),color:"#48bb78"},{label:"Com Domínio",value:metrics.withDomain,pct:Math.round(metrics.withDomain/metrics.totalCnpjs*100),color:ORANGE},{label:"Com Personas",value:metrics.withPersonas,pct:Math.round(metrics.withPersonas/metrics.totalCnpjs*100),color:"#805ad5"},{label:"Com LinkedIn",value:metrics.withLinkedin,pct:Math.round(metrics.withLinkedin/metrics.totalCnpjs*100),color:"#0077b5"}].map(c=>(<StatCard key={c.label} label={c.label} value={c.value} pct={c.pct} color={c.color}/>))}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10,marginBottom:14}}>
              {Object.keys(metrics.statusCount).length>0&&(<div style={{background:ACCENT_BG,borderRadius:10,padding:"14px",border:"1px solid "+BORDER}}><div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10}}>Status dos Telefones</div>{Object.entries(metrics.statusCount).sort((a,b)=>b[1]-a[1]).map(([s,count])=>{const pct=metrics.totalPhoneRows>0?Math.round(count/metrics.totalPhoneRows*100):0;const col=s.toLowerCase().includes("v")&&!s.toLowerCase().includes("inv")?"#38a169":"#e53e3e";return(<div key={s} style={{marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}><span style={{color:MID,fontWeight:600}}>{s}</span><span style={{color:col,fontWeight:700}}>{count} ({pct}%)</span></div><Bar pct={pct} color={col}/></div>);})}</div>)}
              {Object.keys(metrics.catCount).length>0&&(<div style={{background:ACCENT_BG,borderRadius:10,padding:"14px",border:"1px solid "+BORDER}}><div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10}}>Categoria dos Telefones</div>{Object.entries(metrics.catCount).sort((a,b)=>b[1]-a[1]).map(([c,count],idx)=>{const pct=metrics.totalPhoneRows>0?Math.round(count/metrics.totalPhoneRows*100):0;const cols=["#4f46e5","#3182ce","#38a169",ORANGE,"#e53e3e","#805ad5"];const col=cols[idx%cols.length];return(<div key={c} style={{marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}><span style={{color:MID,fontWeight:600}}>{c}</span><span style={{color:col,fontWeight:700}}>{count} ({pct}%)</span></div><Bar pct={pct} color={col}/></div>);})}</div>)}
              <div style={{background:ACCENT_BG,borderRadius:10,padding:"14px",border:"1px solid "+BORDER}}><div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10}}>E-mails válidos</div><div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:6}}><span style={{color:LIGHT}}>{metrics.validEmails} / {metrics.totalEmails}</span><span style={{fontWeight:700,color:"#3182ce"}}>{metrics.totalEmails>0?Math.round(metrics.validEmails/metrics.totalEmails*100):0}%</span></div><Bar pct={metrics.totalEmails>0?Math.round(metrics.validEmails/metrics.totalEmails*100):0} color="#3182ce"/></div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[{title:"Top Estados",data:metrics.topStates,color:"#4f46e5"},{title:"Top Plataformas",data:metrics.topPlatforms,color:ORANGE}].map(block=>(<div key={block.title} style={{background:ACCENT_BG,borderRadius:10,padding:"14px",border:"1px solid "+BORDER}}><div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10}}>{block.title}</div>{block.data.length===0?<div style={{fontSize:12,color:LIGHT}}>Sem dados</div>:block.data.map(([name,count])=>{const pct=Math.round(count/metrics.totalCnpjs*100);return(<div key={name} style={{marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}><span style={{color:MID,fontWeight:600,maxWidth:"70%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{name}</span><span style={{color:block.color,fontWeight:700}}>{count} ({pct}%)</span></div><Bar pct={pct} color={block.color}/></div>);})}</div>))}
              <div style={{background:ACCENT_BG,borderRadius:10,padding:"14px",border:"1px solid "+BORDER,gridColumn:"1 / -1"}}><div style={{fontSize:12,fontWeight:700,color:DARK,marginBottom:10}}>Top CNAEs</div>{metrics.topCnaes.length===0?<div style={{fontSize:12,color:LIGHT}}>Sem dados</div>:metrics.topCnaes.map(([name,count])=>{const pct=Math.round(count/metrics.totalCnpjs*100);return(<div key={name} style={{marginBottom:8}}><div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:3}}><span style={{color:MID,fontWeight:600,maxWidth:"80%",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{name}</span><span style={{color:"#0ea5e9",fontWeight:700}}>{count} ({pct}%)</span></div><Bar pct={pct} color="#0ea5e9"/></div>);})}  </div>
            </div>
          </Card>)}

          {/* Prévia */}
          <Card title={"Prévia (5 primeiros) · "+templateCols.length+" colunas"}>
            <div style={{overflowX:"auto"}}><table style={{borderCollapse:"collapse",fontSize:11,minWidth:600}}><thead><tr style={{background:ACCENT_BG}}>{templateCols.map(h=>(<th key={h} style={{padding:"8px 10px",background:GENERATED_COLS.includes(h)?ORANGE+"10":ACCENT_BG,border:"1px solid "+BORDER,textAlign:"left",whiteSpace:"nowrap",fontWeight:700,color:GENERATED_COLS.includes(h)?ORANGE:MID,fontSize:10}}>{h}</th>))}</tr></thead><tbody>{result.slice(0,5).map((row,i)=>(<tr key={i} style={{background:i%2===0?CARD:ACCENT_BG}}>{templateCols.map(tc=>(<td key={tc} style={{padding:"7px 10px",border:"1px solid "+BORDER,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:GENERATED_COLS.includes(tc)?ORANGE:DARK}}>{row[tc]||""}</td>))}</tr>))}</tbody></table></div>
          </Card>

          {/* Formato de exportação */}
          <Card title="Formato de exportação" accent="#4f46e5">
            <p style={{fontSize:13,color:MID,margin:"0 0 14px"}}>Como deseja que os contatos apareçam no arquivo final?</p>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[{mode:false,title:"Agrupado por CNPJ",desc:"Todos os contatos ficam na mesma linha, separados por vírgula.",example:"62 9999-0001, 62 9999-0002"},{mode:true,title:"Um contato por linha",desc:"Cada contato gera uma linha repetindo os dados cadastrais.",example:"62 9999-0001\n62 9999-0002"}].map(opt=>(
                <div key={String(opt.mode)} onClick={()=>setExpandMode(opt.mode)} style={{padding:16,borderRadius:10,border:"2px solid "+(expandMode===opt.mode?"#4f46e5":BORDER),background:expandMode===opt.mode?"#f0f0fe":CARD,cursor:"pointer"}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                    <div style={{width:16,height:16,borderRadius:"50%",border:"2px solid "+(expandMode===opt.mode?"#4f46e5":BORDER),background:expandMode===opt.mode?"#4f46e5":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{expandMode===opt.mode&&<div style={{width:6,height:6,borderRadius:"50%",background:"#fff"}}/>}</div>
                    <span style={{fontWeight:700,fontSize:13,color:DARK}}>{opt.title}</span>
                  </div>
                  <p style={{fontSize:12,color:MID,margin:"0 0 8px"}}>{opt.desc}</p>
                  <pre style={{fontSize:11,color:"#4f46e5",background:"#f0f0fe",padding:"6px 10px",borderRadius:6,margin:0,fontFamily:"monospace"}}>{opt.example}</pre>
                </div>
              ))}
            </div>
            {expandMode&&(<div style={{marginTop:12,background:"#fffbeb",borderRadius:8,padding:"10px 14px",fontSize:12,color:"#92400e",border:"1px solid #fde68a"}}>No modo expandido o arquivo terá aproximadamente <b>{expandCount} linhas</b> (vs {result.length} no modo agrupado).</div>)}
          </Card>

          <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
            <button onClick={()=>setStep(3)} style={{padding:"11px 20px",borderRadius:9,border:"1px solid "+BORDER,background:CARD,cursor:"pointer",fontSize:13,fontWeight:600,color:MID}}>Reconfigurar</button>
            <button onClick={()=>setStep(2)} style={{padding:"11px 20px",borderRadius:9,border:"1px solid "+BORDER,background:CARD,cursor:"pointer",fontSize:13,fontWeight:600,color:MID}}>Editar template</button>
            <button onClick={reset} style={{padding:"11px 20px",borderRadius:9,border:"1px solid "+BORDER,background:CARD,cursor:"pointer",fontSize:13,fontWeight:600,color:MID}}>Nova planilha</button>
            <button onClick={()=>setShowTPV(true)} style={{flex:1,padding:"11px 20px",borderRadius:9,border:"none",background:ORANGE,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700}}>Calcular Tier</button>
            <button onClick={()=>doDownload(result)} style={{flex:1,padding:"11px 20px",borderRadius:9,border:"none",background:DARK,color:"#fff",cursor:"pointer",fontSize:14,fontWeight:700}}>Baixar .xlsx</button>
          </div>
        </>)}
      </div>
    </div>
  );
}
