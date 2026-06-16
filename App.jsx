import { useState, useEffect, useCallback, useRef } from 'react'
import { getSections, TRANSLATIONS, ALL_FIELD_IDS, ACADEMIC_YEARS, SEMESTERS } from './sections.js'
import { REALISATIONS_SECTIONS } from './realisations_data.js'
import { saveSubmission, loadSubmissions, clearSubmissions, loadLocal } from './storage.js'
import { exportToExcel } from './export.js'
import { generateBilanExcel } from './bilan_export.js'
import { sendAcknowledgement, sendViaAppsScript } from './email_service.js'

const C = { navy:'#0D1B2A',blue:'#1A56DB',teal:'#047481',green:'#057A55',violet:'#5521B5',orange:'#B45309',red:'#BE123C',gold:'#FBBF24',g1:'#F9FAFB',g3:'#E5E7EB',gt:'#6B7280',gd:'#111928' }
const ADMIN_CODE = (typeof import.meta !== 'undefined' && import.meta.env?.VITE_ADMIN_CODE) || 'GSMI2025'
const SCRIPT_URL = (typeof import.meta !== 'undefined' && import.meta.env?.VITE_APPS_SCRIPT_URL) || null

// ── Toast ─────────────────────────────────────────────────────────────────
function Toast({ t, bottom=24 }) {
  return <div style={{position:'fixed',bottom,left:'50%',transform:'translateX(-50%)',
    background:t.type==='error'?C.red:t.type==='info'?C.blue:C.gd,
    color:'#fff',padding:'11px 22px',borderRadius:10,fontSize:13,zIndex:9999,
    whiteSpace:'nowrap',boxShadow:'0 4px 20px rgba(0,0,0,.2)'}}>{t.msg}</div>
}

// ── DOI Verifier ──────────────────────────────────────────────────────────
function DoiField({ value, onChange }) {
  const [dois, setDois] = useState(value ? value.split('\n').filter(Boolean) : [''])
  const [st, setSt] = useState({})
  const [chk, setChk] = useState({})

  const upd = (i, v) => {
    const n=[...dois]; n[i]=v; setDois(n); onChange(n.filter(Boolean).join('\n'))
  }
  const rem = i => {
    const n=dois.filter((_,j)=>j!==i); const nn=n.length?n:['']
    setDois(nn); onChange(nn.filter(Boolean).join('\n'))
    setSt(p=>{const q={...p};delete q[i];return q})
  }

  const verify = async (i, doi) => {
    if (!doi) return
    const clean = doi.trim().replace(/^https?:\/\/doi\.org\//i,'')
    if (!/^10\.\d{4,}\//.test(clean)) {
      setSt(p=>({...p,[i]:{ok:false,msg:'❌ Format invalide — attendu : 10.XXXX/XXXXX'}})); return
    }
    setChk(p=>({...p,[i]:true}))
    setSt(p=>({...p,[i]:{ok:null,msg:'🔍 Vérification via doi.org...'}}))
    try {
      const r = await fetch(`https://doi.org/api/handles/${encodeURIComponent(clean)}`)
      const data = await r.json()
      const ok = data.responseCode === 1
      setSt(p=>({...p,[i]:ok
        ?{ok:true,msg:'✅ DOI valide et résolu — publication vérifiée'}
        :{ok:false,msg:'⚠️ DOI introuvable sur doi.org — vérifier la valeur exacte'}
      }))
    } catch { setSt(p=>({...p,[i]:{ok:null,msg:'🌐 Vérification hors-ligne — format correct'}})) }
    setChk(p=>({...p,[i]:false}))
  }

  return (
    <div>
      {dois.map((doi,i)=>(
        <div key={i} style={{marginBottom:8}}>
          <div style={{display:'flex',gap:6}}>
            <input type="text" value={doi} onChange={e=>upd(i,e.target.value)}
              onBlur={e=>verify(i,e.target.value)}
              placeholder="10.XXXX/XXXXX"
              style={{flex:1,padding:'9px 12px',border:`1.5px solid ${st[i]?.ok===false?C.red:st[i]?.ok===true?C.green:C.g3}`,borderRadius:8,fontSize:13,fontFamily:'monospace',outline:'none',color:C.gd,background:'#fff'}}/>
            <button onClick={()=>verify(i,doi)} disabled={chk[i]}
              style={{padding:'9px 12px',background:C.blue,color:'#fff',border:'none',borderRadius:8,fontSize:12,cursor:'pointer',whiteSpace:'nowrap'}}>
              {chk[i]?'…':'🔍 Vérifier'}
            </button>
            {dois.length>1&&<button onClick={()=>rem(i)} style={{padding:'9px 10px',background:'transparent',color:C.gt,border:`1px solid ${C.g3}`,borderRadius:8,cursor:'pointer'}}>✕</button>}
          </div>
          {st[i]&&<p style={{fontSize:12,margin:'3px 0 0',color:st[i].ok===true?C.green:st[i].ok===false?C.red:C.orange}}>{st[i].msg}</p>}
        </div>
      ))}
      <button onClick={()=>setDois([...dois,''])}
        style={{marginTop:4,width:'100%',padding:'7px',background:'transparent',color:C.blue,border:`1px dashed ${C.blue}`,borderRadius:8,fontSize:13,cursor:'pointer'}}>
        + Ajouter un DOI
      </button>
    </div>
  )
}

// ── Simple Field ──────────────────────────────────────────────────────────
function Field({f, form, onChange, errors}) {
  const val = form[f.id]??'', err = errors[f.id]
  const [focused, setFoc] = useState(false)
  const base = {
    width:'100%',padding:'10px 12px',outline:'none',fontFamily:'inherit',boxSizing:'border-box',
    border:`1.5px solid ${err?C.red:focused?C.blue:C.g3}`,borderRadius:8,fontSize:14,color:C.gd,
    background:'#fff',transition:'border-color .15s',boxShadow:focused?`0 0 0 3px ${C.blue}18`:'none',
  }
  const ch = e => { onChange(f.id, e.target.value); }

  if (f.type==='doi') return (
    <div style={{marginBottom:20}}>
      <label style={{display:'block',fontSize:13,fontWeight:600,color:C.gd,marginBottom:5}}>{f.label}</label>
      {f.hint&&<p style={{fontSize:12,color:C.gt,margin:'0 0 6px'}}>{f.hint}</p>}
      <DoiField value={val} onChange={v=>onChange(f.id,v)}/>
      {err&&<p style={{fontSize:12,color:C.red,margin:'4px 0 0'}}>⚠ {err}</p>}
    </div>
  )

  return (
    <div style={{marginBottom:18}}>
      <label style={{display:'block',fontSize:13,fontWeight:600,color:C.gd,marginBottom:5}}>
        {f.label}{f.required&&<span style={{color:C.red,marginLeft:3}}>*</span>}
      </label>
      {f.hint&&<p style={{fontSize:12,color:C.gt,margin:'0 0 6px',lineHeight:1.4}}>{f.hint}</p>}
      {f.type==='textarea'
        ?<textarea value={val} onChange={ch} placeholder={f.placeholder||''} onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)} style={{...base,minHeight:76,resize:'vertical',lineHeight:1.5}}/>
        :f.type==='select'
        ?<select value={val} onChange={ch} onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)} style={{...base,cursor:'pointer'}}>
           <option value="">— Sélectionner —</option>
           {f.options.map(o=><option key={o} value={o}>{o}</option>)}
         </select>
        :<input type={f.type==='number'?'number':'text'} value={val} onChange={ch}
           placeholder={f.placeholder||''} min={f.min} max={f.max}
           onFocus={()=>setFoc(true)} onBlur={()=>setFoc(false)} style={base}/>
      }
      {err&&<p style={{fontSize:12,color:C.red,margin:'4px 0 0'}}>⚠ {err}</p>}
    </div>
  )
}

// ── Realisations Table ────────────────────────────────────────────────────
function RealisationsTable({ section, rows, onChange }) {
  const addRow = () => onChange([...rows, {}])
  const delRow = i => onChange(rows.filter((_,j)=>j!==i))
  const setCell = (i,id,v) => { const n=[...rows]; n[i]={...n[i],[id]:v}; onChange(n) }

  return (
    <div>
      {/* Header */}
      <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:12}}>
        <div>
          <h4 style={{margin:0,fontSize:15,fontWeight:700,color:section.color}}>{section.icon} {section.title}</h4>
          <p style={{margin:'3px 0 0',fontSize:12,color:C.gt}}>{section.description}</p>
        </div>
        <button onClick={addRow} style={{padding:'8px 14px',background:section.color,color:'#fff',border:'none',borderRadius:8,fontSize:13,cursor:'pointer',whiteSpace:'nowrap'}}>
          + Ajouter une ligne
        </button>
      </div>

      {rows.length===0
        ?<div style={{border:`1.5px dashed ${C.g3}`,borderRadius:10,padding:'28px 20px',textAlign:'center',color:C.gt,fontSize:14}}>
           Cliquer "Ajouter une ligne" pour saisir vos réalisations
         </div>
        :<div style={{overflowX:'auto',borderRadius:10,border:`1px solid ${C.g3}`}}>
           <table style={{borderCollapse:'collapse',fontSize:12,minWidth:rows.length?'max-content':'100%',width:'100%'}}>
             <thead>
               <tr style={{background:section.color}}>
                 {section.table.cols.map(c=>(
                   <th key={c.id} style={{padding:'8px 10px',color:'#fff',fontWeight:600,whiteSpace:'nowrap',minWidth:c.width,textAlign:'left'}}>
                     {c.label}{c.required&&<span style={{color:C.gold,marginLeft:2}}>*</span>}
                   </th>
                 ))}
                 <th style={{padding:'8px 10px',color:'#fff',fontWeight:600,width:36}}>✕</th>
               </tr>
             </thead>
             <tbody>
               {rows.map((row,i)=>(
                 <tr key={i} style={{borderBottom:`0.5px solid ${C.g3}`,background:i%2===0?'#fff':C.g1}}>
                   {section.table.cols.map(c=>(
                     <td key={c.id} style={{padding:'5px 8px',minWidth:c.width}}>
                       {c.type==='doi'
                         ?<DoiField value={row[c.id]||''} onChange={v=>setCell(i,c.id,v)}/>
                         :c.type==='select'
                         ?<select value={row[c.id]||''} onChange={e=>setCell(i,c.id,e.target.value)}
                            style={{width:'100%',padding:'6px 8px',border:`1px solid ${C.g3}`,borderRadius:6,fontSize:12,color:C.gd,background:'#fff',outline:'none'}}>
                            <option value="">—</option>
                            {c.options.map(o=><option key={o} value={o}>{o}</option>)}
                          </select>
                         :<input type={c.type==='number'?'number':'text'}
                            value={row[c.id]||''}
                            onChange={e=>setCell(i,c.id,c.type==='number'?+e.target.value:e.target.value)}
                            min={c.min} max={c.max}
                            style={{width:'100%',padding:'6px 8px',border:`1px solid ${C.g3}`,borderRadius:6,fontSize:12,color:C.gd,background:'transparent',outline:'none',boxSizing:'border-box'}}/>
                       }
                     </td>
                   ))}
                   <td style={{padding:'5px 8px',textAlign:'center'}}>
                     <button onClick={()=>delRow(i)} style={{background:'transparent',border:'none',color:C.red,cursor:'pointer',fontSize:14,padding:'2px 6px'}}>✕</button>
                   </td>
                 </tr>
               ))}
             </tbody>
           </table>
         </div>
      }

      {/* Aggregates */}
      {rows.length>0&&(
        <div style={{marginTop:12,display:'flex',flexWrap:'wrap',gap:8}}>
          {section.table.aggregates.map(agg=>(
            <div key={agg.label} style={{background:section.color+'18',border:`1px solid ${section.color}30`,borderRadius:8,padding:'6px 12px'}}>
              <span style={{fontSize:11,color:C.gt}}>{agg.label}: </span>
              <strong style={{fontSize:13,color:section.color}}>{agg.fn(rows)}</strong>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

// ══════════════════════════════════════════════════════════════════════════
// MAIN APP
// ══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [lang,setLang]         = useState('fr')
  const [view,setView]         = useState('home')
  const [mode,setMode]         = useState(null)
  const [step,setStep]         = useState(0)
  const [form,setForm]         = useState({})
  const [errors,setErrors]     = useState({})
  const [realisations,setReal] = useState({})  // { real_production:[], real_projets:[], ... }
  const [subs,setSubs]         = useState([])
  const [loading,setLoading]   = useState(false)
  const [adminCode,setAdminCode]= useState('')
  const [adminOk,setAdminOk]   = useState(false)
  const [adminSearch,setAdminSearch]= useState('')
  const [adminTab,setAdminTab] = useState('kpi')
  const [toast,setToast]       = useState(null)

  const T = TRANSLATIONS[lang]

  useEffect(()=>{loadSubmissions().then(setSubs)},[])

  function showToast(msg,type='success'){setToast({msg,type});setTimeout(()=>setToast(null),4000)}

  const sections = mode ? getSections(mode, lang) : []

  const handleChange = useCallback((id,v)=>{
    setForm(p=>({...p,[id]:v}))
    setErrors(p=>{const n={...p};delete n[id];return n})
  },[])

  function validate(stepIdx) {
    const sec = sections[stepIdx]; if(!sec) return true
    const errs={}
    sec.fields.forEach(f=>{
      const val=form[f.id]??''
      if(f.required&&(!val||val==='')) { errs[f.id]=lang==='fr'?'Champ requis':'Required field'; return }
      if(f.validate){const m=f.validate(val,form);if(m)errs[f.id]=m}
      if(f.type==='number'&&val!==''&&isNaN(+val)) errs[f.id]=lang==='fr'?'Valeur numérique requise':'Number required'
    })
    setErrors(errs)
    return Object.keys(errs).length===0
  }

  function startForm(selectedMode){
    setMode(selectedMode); setForm({mode:selectedMode}); setStep(0)
    setErrors({}); setReal({}); setView('form'); window.scrollTo(0,0)
  }

  function next(){
    if(!validate(step)){showToast(lang==='fr'?'Corriger les erreurs':'Fix errors','error');return}
    if(step<sections.length-1){setStep(s=>s+1);window.scrollTo(0,0)}
    else handleSubmit()
  }
  function prev(){setStep(s=>Math.max(0,s-1));setErrors({});window.scrollTo(0,0)}

  async function handleSubmit(){
    if(!validate(step))return
    setLoading(true)
    try {
      const payload = { ...form, mode, realisations }
      const all = await saveSubmission(payload)
      setSubs(all)
      // Accusé de réception
      if(form.email&&form.nom) {
        await sendAcknowledgement(form.email, form.nom, mode, form.annee_academique||'')
          .catch(e=>console.warn('Email ACK failed:',e))
        if(SCRIPT_URL) {
          const {sendViaAppsScript}=await import('./email_service.js')
          await sendViaAppsScript(SCRIPT_URL, form.email, form.nom, mode, form.annee_academique||'')
            .catch(()=>{})
        }
        showToast(`✅ Soumission enregistrée — accusé envoyé à ${form.email}`)
      }
      setView('thanks')
    } catch(e) { showToast(lang==='fr'?'Erreur — réessayez':'Error — retry','error') }
    setLoading(false)
  }

  // Bilan export — cherche les 3 modes pour ce prof + année
  function handleBilanExport(profEmail, annee) {
    const all = loadLocal()
    const match = (m) => all.find(s=>s.email===profEmail&&s.annee_academique===annee&&s.mode===m)
    const prev  = match('prevision')
    const rev   = match('revision_s1')
    const bilan = match('bilan_annuel')
    const real  = (bilan||prev||rev)?.realisations || {}
    generateBilanExcel(null, prev, rev, bilan||{}, real)
    showToast('Bilan Excel généré — téléchargement en cours')
  }

  const sec = sections[step]

  // ── HOME ────────────────────────────────────────────────────────────────
  if(view==='home') return (
    <div style={{minHeight:'100vh',background:C.g1,fontFamily:'system-ui,-apple-system,sans-serif'}}>
      <style>{`*{box-sizing:border-box}button{transition:opacity .15s,transform .1s}button:hover{opacity:.87}button:active{transform:scale(.97)}`}</style>

      {/* Lang toggle */}
      <div style={{position:'fixed',top:14,right:14,zIndex:100}}>
        <div style={{display:'flex',background:'#fff',border:`1px solid ${C.g3}`,borderRadius:8,overflow:'hidden'}}>
          {['fr','en'].map(l=>(
            <button key={l} onClick={()=>setLang(l)}
              style={{padding:'7px 13px',border:'none',background:lang===l?C.navy:'transparent',color:lang===l?'#fff':C.gt,fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
              {l==='fr'?'🇫🇷 FR':'🇬🇧 EN'}
            </button>
          ))}
        </div>
      </div>

      {/* Hero */}
      <div style={{background:C.navy,padding:'52px 24px 44px',textAlign:'center'}}>
        <p style={{color:C.gold,fontSize:11,letterSpacing:'.14em',textTransform:'uppercase',margin:'0 0 12px',fontWeight:600}}>
          Green & Sustainable Mining Institute · UM6P
        </p>
        <h1 style={{color:'#fff',fontSize:32,fontWeight:700,margin:'0 0 10px',lineHeight:1.15}}>
          {lang==='fr'?'Carnet du Chercheur':'Researcher Notebook'}
        </h1>
        <p style={{color:'#8899BB',fontSize:14,margin:'0 0 10px',maxWidth:500,marginLeft:'auto',marginRight:'auto',lineHeight:1.6}}>
          {lang==='fr'
            ?'Gestion des prévisions, révision semestrielle et bilan annuel — GSMI / UM6P'
            :'Forecast management, mid-year revision and annual review — GSMI / UM6P'}
        </p>

        {/* Logo placeholder */}
        <div style={{display:'inline-flex',alignItems:'center',gap:8,background:'rgba(255,255,255,.07)',borderRadius:8,padding:'8px 16px',marginBottom:8}}>
          <span style={{fontSize:11,color:'rgba(255,255,255,.4)',fontStyle:'italic'}}>
            {lang==='fr'?'[ Intégrez ici vos logos UM6P & GSMI ]':'[ Insert your UM6P & GSMI logos here ]'}
          </span>
        </div>
      </div>

      {/* Objectif + Axes */}
      <div style={{maxWidth:700,margin:'0 auto',padding:'28px 18px 0'}}>
        <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,padding:'20px 22px',marginBottom:18}}>
          <h3 style={{margin:'0 0 10px',fontSize:14,fontWeight:700,color:C.navy}}>
            {lang==='fr'?'🎯 Objectif':'🎯 Objective'}
          </h3>
          <p style={{margin:'0 0 14px',fontSize:13,color:C.gt,lineHeight:1.65}}>
            {lang==='fr'
              ?'Chaque professeur GSMI saisit ses prévisions, les révise à mi-année (S1), puis présente son bilan annuel à l\'Academic Meeting. Les 3 modes génèrent automatiquement un bilan comparatif Excel avec calcul des écarts.'
              :'Each GSMI professor enters forecasts, revises at mid-year (S1), then presents an annual review at the Academic Meeting. All 3 modes automatically generate a comparative Excel review with variance calculations.'}
          </p>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(170px,1fr))',gap:8}}>
            {[
              {icon:'🗺️',label:'Geology & Exploration',color:C.teal},
              {icon:'⚙️',label:'Mine & Mineral Processing',color:C.green},
              {icon:'🌍',label:'Sustainable Mining & Environment',color:C.violet},
            ].map(ax=>(
              <div key={ax.label} style={{padding:'9px 12px',borderRadius:8,background:ax.color+'12',borderLeft:`3px solid ${ax.color}`,display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontSize:18}}>{ax.icon}</span>
                <span style={{fontSize:11,fontWeight:600,color:ax.color,lineHeight:1.3}}>{ax.label}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Logique d'écarts */}
        <div style={{background:'#FFF8E6',border:`1px solid ${C.gold}40`,borderRadius:10,padding:'14px 18px',marginBottom:18}}>
          <h4 style={{margin:'0 0 8px',fontSize:13,fontWeight:700,color:C.orange}}>
            📐 {lang==='fr'?'Logique de calcul des écarts':'Variance calculation logic'}
          </h4>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
            <div style={{background:'#fff',borderRadius:8,padding:'10px 12px',borderLeft:`3px solid ${C.blue}`}}>
              <p style={{margin:0,fontSize:12,fontWeight:600,color:C.blue}}>Écart 1 — Réalisé vs Prévu</p>
              <p style={{margin:'4px 0 0',fontSize:11,color:C.gt}}>C – A : mesure l'atteinte des objectifs initiaux</p>
            </div>
            <div style={{background:'#fff',borderRadius:8,padding:'10px 12px',borderLeft:`3px solid ${C.teal}`}}>
              <p style={{margin:0,fontSize:12,fontWeight:600,color:C.teal}}>Écart 2 — Réalisé vs Révisé S1</p>
              <p style={{margin:'4px 0 0',fontSize:11,color:C.gt}}>C – B : mesure l'atteinte après ajustement mi-année</p>
            </div>
          </div>
        </div>

        {/* 3 modes */}
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(185px,1fr))',gap:12,marginBottom:18}}>
          {[
            {mode:'prevision',   icon:'🎯',color:C.blue,   title:T.mode_prevision, hint:T.mode_hint_prev},
            {mode:'revision_s1', icon:'🔄',color:C.teal,   title:T.mode_revision,  hint:T.mode_hint_rev},
            {mode:'bilan_annuel',icon:'📊',color:C.violet, title:T.mode_bilan,     hint:T.mode_hint_bilan},
          ].map(m=>(
            <button key={m.mode} onClick={()=>startForm(m.mode)}
              style={{background:'#fff',border:`1.5px solid ${C.g3}`,borderRadius:12,padding:'18px 16px',cursor:'pointer',textAlign:'left',fontFamily:'inherit',borderTop:`4px solid ${m.color}`,transition:'box-shadow .2s'}}>
              <div style={{fontSize:26,marginBottom:10}}>{m.icon}</div>
              <p style={{margin:'0 0 5px',fontWeight:700,fontSize:13,color:C.gd,lineHeight:1.2}}>{m.title}</p>
              <p style={{margin:0,fontSize:11,color:C.gt,lineHeight:1.4}}>{m.hint}</p>
            </button>
          ))}
        </div>

        {/* Realisations mode */}
        <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,marginBottom:18,overflow:'hidden'}}>
          <button onClick={()=>startForm('realisations')}
            style={{width:'100%',padding:'16px 20px',background:'none',border:'none',cursor:'pointer',fontFamily:'inherit',display:'flex',alignItems:'center',gap:14,textAlign:'left'}}>
            <span style={{fontSize:28}}>📋</span>
            <div>
              <p style={{margin:0,fontWeight:700,fontSize:14,color:C.gd}}>{lang==='fr'?'④ Saisie des réalisations détaillées':'④ Detailed achievements entry'}</p>
              <p style={{margin:'3px 0 0',fontSize:12,color:C.gt}}>
                {lang==='fr'?'Publications Scopus ligne par ligne, projets, rayonnement, formation, prestations — avec vérification DOI':'Scopus publications line by line, projects, outreach, training, consulting — with DOI verification'}
              </p>
            </div>
            <span style={{marginLeft:'auto',color:C.gt,fontSize:20}}>→</span>
          </button>
        </div>

        <div style={{textAlign:'center',paddingBottom:40}}>
          <button onClick={()=>setView('admin')}
            style={{background:'transparent',color:C.gt,border:`0.5px solid ${C.g3}`,borderRadius:8,padding:'10px 20px',fontSize:13,cursor:'pointer',fontFamily:'inherit'}}>
            🔒 {lang==='fr'?'Accès Direction GSMI':'GSMI Management Access'}
          </button>
        </div>
      </div>
      {toast&&<Toast t={toast}/>}
    </div>
  )

  // ── FORM (prévision / révision / bilan) ────────────────────────────────
  if(view==='form'&&mode!=='realisations'&&sec) return (
    <div style={{minHeight:'100vh',background:C.g1,fontFamily:'system-ui,-apple-system,sans-serif'}}>
      <style>{`*{box-sizing:border-box}button:hover{opacity:.87}button:active{transform:scale(.97)}`}</style>

      <div style={{background:sec.color,padding:'18px 20px 0',position:'sticky',top:0,zIndex:10}}>
        <div style={{maxWidth:620,margin:'0 auto'}}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:12}}>
            <button onClick={()=>setView('home')} style={{background:'transparent',border:'none',color:'rgba(255,255,255,.7)',fontSize:13,cursor:'pointer',padding:0,fontFamily:'inherit'}}>
              ← {lang==='fr'?'Accueil':'Home'}
            </button>
            <span style={{color:'rgba(255,255,255,.6)',fontSize:12}}>
              {T[`mode_${mode==='prevision'?'prevision':mode==='revision_s1'?'revision':'bilan'}`]} · {step+1}/{sections.length}
            </span>
          </div>
          <div style={{display:'flex',alignItems:'center',gap:10,marginBottom:14}}>
            <span style={{fontSize:22}}>{sec.icon}</span>
            <div>
              <p style={{color:'rgba(255,255,255,.5)',fontSize:10,margin:'0 0 1px',textTransform:'uppercase',letterSpacing:'.08em'}}>Section {step+1}/{sections.length}</p>
              <h2 style={{color:'#fff',fontSize:17,fontWeight:700,margin:0}}>{sec.title}</h2>
            </div>
          </div>
          <div style={{display:'flex',gap:3}}>
            {sections.map((_,i)=>(
              <div key={i} style={{flex:1,height:3,borderRadius:'3px 3px 0 0',
                background:i<step?'#4ADE80':i===step?'#fff':'rgba(255,255,255,.2)',transition:'background .3s'}}/>
            ))}
          </div>
        </div>
      </div>

      <div style={{maxWidth:620,margin:'0 auto',padding:'24px 20px 120px'}}>
        {/* Accusé de réception notice */}
        {step===0&&<div style={{background:'#EFF6FF',border:`1px solid ${C.blue}30`,borderRadius:8,padding:'10px 14px',marginBottom:20,borderLeft:`3px solid ${C.blue}`}}>
          <p style={{margin:0,fontSize:12,color:'#1e40af'}}>
            📧 {lang==='fr'?'Un accusé de réception sera envoyé à votre adresse email après soumission.':'A confirmation email will be sent to your address after submission.'}
          </p>
        </div>}
        {sec.hint&&<div style={{background:sec.color+'10',borderRadius:8,padding:'10px 14px',marginBottom:20,borderLeft:`3px solid ${sec.color}`}}>
          <p style={{margin:0,fontSize:13,color:sec.color,lineHeight:1.5}}>{sec.hint}</p>
        </div>}
        {sec.fields.map(f=><Field key={f.id} f={f} form={form} onChange={handleChange} errors={errors}/>)}
      </div>

      <div style={{position:'fixed',bottom:0,left:0,right:0,background:'#fff',borderTop:`0.5px solid ${C.g3}`,padding:'13px 20px',display:'flex',justifyContent:'space-between',alignItems:'center',boxSizing:'border-box',zIndex:10}}>
        <button onClick={prev} disabled={step===0}
          style={{background:'transparent',border:`1px solid ${C.g3}`,borderRadius:8,padding:'10px 18px',fontSize:14,cursor:step===0?'not-allowed':'pointer',color:step===0?C.gt:C.gd,opacity:step===0?.35:1,fontFamily:'inherit'}}>
          ← {lang==='fr'?'Précédent':'Back'}
        </button>
        <div style={{display:'flex',gap:6}}>
          {sections.map((_,i)=>(<div key={i} style={{width:8,height:8,borderRadius:'50%',background:i===step?sec.color:i<step?C.green:C.g3,transition:'background .25s',transform:i===step?'scale(1.3)':'scale(1)'}}/>))}
        </div>
        <button onClick={next} disabled={loading}
          style={{background:sec.color,color:'#fff',border:'none',borderRadius:8,padding:'10px 22px',fontSize:14,fontWeight:600,cursor:loading?'wait':'pointer',minWidth:110,fontFamily:'inherit'}}>
          {loading?'…':step===sections.length-1?(lang==='fr'?'Soumettre ✓':'Submit ✓'):(lang==='fr'?'Suivant →':'Next →')}
        </button>
      </div>
      {toast&&<Toast t={toast} bottom={90}/>}
    </div>
  )

  // ── REALISATIONS FORM ──────────────────────────────────────────────────
  if(view==='form'&&mode==='realisations') return (
    <div style={{minHeight:'100vh',background:C.g1,fontFamily:'system-ui,-apple-system,sans-serif'}}>
      <style>{`*{box-sizing:border-box}button:hover{opacity:.87}button:active{transform:scale(.97)}`}</style>

      <div style={{background:C.navy,padding:'18px 20px',position:'sticky',top:0,zIndex:10}}>
        <div style={{maxWidth:1100,margin:'0 auto',display:'flex',alignItems:'center',justifyContent:'space-between',gap:12}}>
          <div style={{display:'flex',alignItems:'center',gap:12}}>
            <button onClick={()=>setView('home')} style={{background:'transparent',border:'none',color:'rgba(255,255,255,.7)',fontSize:13,cursor:'pointer',padding:0,fontFamily:'inherit'}}>
              ← {lang==='fr'?'Accueil':'Home'}
            </button>
            <h2 style={{color:'#fff',fontSize:16,fontWeight:700,margin:0}}>
              📋 {lang==='fr'?'Saisie des réalisations détaillées':'Detailed achievements entry'}
            </h2>
          </div>
          <button onClick={handleSubmit} disabled={loading}
            style={{background:C.green,color:'#fff',border:'none',borderRadius:8,padding:'9px 20px',fontSize:13,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
            {loading?'…':(lang==='fr'?'✓ Enregistrer':'✓ Save')}
          </button>
        </div>
      </div>

      <div style={{maxWidth:1100,margin:'0 auto',padding:'24px 18px 40px'}}>
        {/* Identification rapide */}
        <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,padding:'18px 20px',marginBottom:24}}>
          <h3 style={{margin:'0 0 14px',fontSize:14,fontWeight:700,color:C.navy}}>👤 Identification</h3>
          <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(180px,1fr))',gap:12}}>
            {[
              {id:'nom',label:'Nom complet',type:'text',required:true,placeholder:'Prénom NOM'},
              {id:'email',label:'Email UM6P',type:'email',required:true,placeholder:'prenom.nom@um6p.ma'},
              {id:'annee_academique',label:'Année académique',type:'select',required:true,options:ACADEMIC_YEARS},
            ].map(f=><Field key={f.id} f={f} form={form} onChange={handleChange} errors={errors}/>)}
          </div>
        </div>

        {/* Realisations sections */}
        {REALISATIONS_SECTIONS.map(sec=>(
          <div key={sec.id} style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,padding:'20px 20px',marginBottom:18}}>
            <RealisationsTable
              section={sec}
              rows={realisations[sec.id]||[]}
              onChange={rows=>setReal(p=>({...p,[sec.id]:rows}))}
            />
          </div>
        ))}

        <div style={{textAlign:'right',marginTop:8}}>
          <button onClick={handleSubmit} disabled={loading}
            style={{background:C.navy,color:'#fff',border:'none',borderRadius:10,padding:'13px 28px',fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
            {loading?'…':(lang==='fr'?'✓ Enregistrer les réalisations':'✓ Save achievements')}
          </button>
        </div>
      </div>
      {toast&&<Toast t={toast}/>}
    </div>
  )

  // ── THANKS ─────────────────────────────────────────────────────────────
  if(view==='thanks') return (
    <div style={{minHeight:'100vh',background:C.g1,display:'flex',alignItems:'center',justifyContent:'center',fontFamily:'system-ui,sans-serif',padding:20}}>
      <div style={{textAlign:'center',maxWidth:440}}>
        <div style={{width:72,height:72,borderRadius:'50%',background:'#D1FAE5',display:'flex',alignItems:'center',justifyContent:'center',fontSize:34,margin:'0 auto 20px'}}>✓</div>
        <h2 style={{color:C.gd,fontSize:22,fontWeight:700,margin:'0 0 10px'}}>{lang==='fr'?'Soumission enregistrée !':'Submission saved!'}</h2>
        <p style={{color:C.gt,fontSize:14,lineHeight:1.65,margin:'0 0 14px'}}>
          {lang==='fr'?'Vos données ont été transmises à la Direction GSMI.':'Your data has been sent to GSMI management.'}
        </p>
        {form.email&&<div style={{background:'#EFF6FF',borderRadius:8,padding:'12px 16px',borderLeft:`3px solid ${C.blue}`,textAlign:'left',marginBottom:18}}>
          <p style={{margin:0,fontSize:13,color:'#1e40af',lineHeight:1.55}}>
            📧 {lang==='fr'?`Un accusé de réception a été envoyé à ${form.email}`:`Confirmation sent to ${form.email}`}
          </p>
        </div>}
        {mode==='bilan_annuel'&&<div style={{marginBottom:18}}>
          <button onClick={()=>handleBilanExport(form.email, form.annee_academique)}
            style={{width:'100%',background:C.green,color:'#fff',border:'none',borderRadius:10,padding:'13px',fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:'inherit',marginBottom:8}}>
            📊 {lang==='fr'?'Générer mon Bilan Annuel Excel':'Generate my Annual Excel Review'}
          </button>
          <p style={{margin:0,fontSize:11,color:C.gt}}>{lang==='fr'?'Compile : Prévu / Révisé S1 / Réalisé + 2 types d\'écarts + réalisations détaillées':'Compiles: Forecast / S1 Revision / Achieved + 2 variance types + detailed achievements'}</p>
        </div>}
        <button onClick={()=>{setView('home');setMode(null)}}
          style={{background:C.blue,color:'#fff',border:'none',borderRadius:10,padding:'12px 26px',fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
          ← {lang==='fr'?'Retour à l\'accueil':'Back to home'}
        </button>
      </div>
    </div>
  )

  // ── ADMIN ───────────────────────────────────────────────────────────────
  if(view==='admin') {
    if(!adminOk) return (
      <div style={{minHeight:'100vh',background:C.g1,display:'flex',alignItems:'center',justifyContent:'center',fontFamily:'system-ui,sans-serif'}}>
        <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:14,padding:'32px 28px',width:'100%',maxWidth:340}}>
          <div style={{width:52,height:52,borderRadius:12,background:'#EFF6FF',display:'flex',alignItems:'center',justifyContent:'center',fontSize:26,marginBottom:18}}>🔒</div>
          <h2 style={{margin:'0 0 6px',fontSize:17,fontWeight:700,color:C.gd}}>{lang==='fr'?'Accès Direction GSMI':'GSMI Management'}</h2>
          <p style={{margin:'0 0 18px',fontSize:13,color:C.gt}}>{lang==='fr'?'Code administrateur':'Admin code'}</p>
          <input type="password" value={adminCode} onChange={e=>setAdminCode(e.target.value)}
            placeholder={lang==='fr'?'Code d\'accès...':'Access code...'}
            style={{width:'100%',padding:'11px 14px',border:`1.5px solid ${C.g3}`,borderRadius:8,fontSize:14,boxSizing:'border-box',marginBottom:12,fontFamily:'inherit',outline:'none'}}
            onKeyDown={e=>e.key==='Enter'&&(adminCode===ADMIN_CODE?setAdminOk(true):showToast(lang==='fr'?'Code incorrect':'Wrong code','error'))}/>
          <button onClick={()=>adminCode===ADMIN_CODE?setAdminOk(true):showToast(lang==='fr'?'Code incorrect':'Wrong code','error')}
            style={{width:'100%',background:C.navy,color:'#fff',border:'none',borderRadius:8,padding:'12px',fontSize:14,fontWeight:600,cursor:'pointer',marginBottom:10,fontFamily:'inherit'}}>
            {lang==='fr'?'Accéder':'Access'}
          </button>
          <button onClick={()=>setView('home')}
            style={{width:'100%',background:'transparent',color:C.gt,border:'none',fontSize:13,cursor:'pointer',padding:6,fontFamily:'inherit'}}>
            ← {lang==='fr'?'Retour':'Back'}
          </button>
        </div>
        {toast&&<Toast t={toast}/>}
      </div>
    )

    // Build prof map
    const profMap = {}
    subs.forEach(s=>{
      const k=s.email||s.nom||'?'
      if(!profMap[k]) profMap[k]={nom:s.nom,email:s.email,grade:s.grade,axe:s.axe_recherche,annee:s.annee_academique}
      profMap[k][s.mode]=s
    })
    const profs = Object.values(profMap)

    // Filter by search
    const filteredProfs = profs.filter(p=>
      !adminSearch || (p.nom||'').toLowerCase().includes(adminSearch.toLowerCase()) || (p.email||'').toLowerCase().includes(adminSearch.toLowerCase())
    )

    // KPI by Axe
    const axes = ['Geology & Exploration','Mine & Mineral Processing (MMP)','Sustainable Mining & Environment (SME)']
    const kpiByAxe = axes.map(axe=>{
      const axeSubs = subs.filter(s=>s.axe_recherche===axe)
      const prev = axeSubs.filter(s=>s.mode==='prevision')
      const bilan = axeSubs.filter(s=>s.mode==='bilan_annuel')
      const sum = (arr,k)=>arr.reduce((a,s)=>a+(+s[k]||0),0)
      return {
        axe,
        profs: new Set(axeSubs.map(s=>s.email||s.nom)).size,
        pub_prev: sum(prev,'prev_pub_total'), pub_real: sum(bilan,'pub_acceptees'),
        proj_prev: sum(prev,'prev_projets_obt'), proj_real: sum(bilan,'projets_obtenus'),
        budget_prev: sum(prev,'prev_budget'), budget_real: sum(bilan,'budget_mad'),
        h_init_real: sum(bilan,'h_initiale'), doct_real: sum(bilan,'doctorants'),
      }
    })

    return (
      <div style={{minHeight:'100vh',background:C.g1,fontFamily:'system-ui,-apple-system,sans-serif'}}>
        <style>{`button:hover{opacity:.87}button:active{transform:scale(.97)}`}</style>

        {/* Header */}
        <div style={{background:C.navy,padding:'18px 22px',display:'flex',alignItems:'center',justifyContent:'space-between',flexWrap:'wrap',gap:12,position:'sticky',top:0,zIndex:10}}>
          <div>
            <p style={{color:C.gold,fontSize:11,letterSpacing:'.1em',margin:'0 0 3px',textTransform:'uppercase',fontWeight:600}}>GSMI — Direction</p>
            <h1 style={{color:'#fff',fontSize:17,fontWeight:700,margin:0}}>{lang==='fr'?'Tableau de bord consolidé':'Consolidated Dashboard'}</h1>
          </div>
          <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
            <button onClick={()=>exportToExcel(subs)}
              style={{background:C.green,color:'#fff',border:'none',borderRadius:8,padding:'8px 14px',fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}}>
              ⬇ Excel global
            </button>
            <button onClick={()=>{setAdminOk(false);setAdminCode('');setView('home')}}
              style={{background:'transparent',color:'#8899BB',border:'1.5px solid #2D3F55',borderRadius:8,padding:'8px 12px',fontSize:12,cursor:'pointer',fontFamily:'inherit'}}>
              {lang==='fr'?'Déconnexion':'Sign out'}
            </button>
          </div>
        </div>

        {/* Tabs */}
        <div style={{background:'#fff',borderBottom:`1px solid ${C.g3}`,padding:'0 22px',display:'flex',gap:0}}>
          {[
            {id:'kpi',label:lang==='fr'?'📊 KPI par Axe':'📊 KPI by Axis'},
            {id:'profs',label:lang==='fr'?'👥 Suivi professeurs':'👥 Professor tracking'},
            {id:'bilan',label:lang==='fr'?'📋 Bilan individuel':'📋 Individual review'},
          ].map(tab=>(
            <button key={tab.id} onClick={()=>setAdminTab(tab.id)}
              style={{padding:'13px 18px',background:'none',border:'none',borderBottom:adminTab===tab.id?`3px solid ${C.blue}`:'3px solid transparent',fontWeight:adminTab===tab.id?700:400,color:adminTab===tab.id?C.blue:C.gt,cursor:'pointer',fontSize:13,fontFamily:'inherit',whiteSpace:'nowrap'}}>
              {tab.label}
            </button>
          ))}
        </div>

        <div style={{maxWidth:1080,margin:'0 auto',padding:'20px 18px'}}>

          {/* ── TAB 1: KPI par Axe ── */}
          {adminTab==='kpi'&&(
            <div>
              {kpiByAxe.map((ax,i)=>{
                const axColors=[C.teal,C.green,C.violet]
                const col=axColors[i]
                return (
                  <div key={ax.axe} style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,marginBottom:16,overflow:'hidden'}}>
                    <div style={{background:col,padding:'12px 18px',display:'flex',alignItems:'center',justifyContent:'space-between'}}>
                      <h3 style={{color:'#fff',margin:0,fontSize:14,fontWeight:700}}>{ax.axe}</h3>
                      <span style={{background:'rgba(255,255,255,.2)',color:'#fff',borderRadius:6,padding:'3px 10px',fontSize:12}}>
                        {ax.profs} professeur{ax.profs>1?'s':''}
                      </span>
                    </div>
                    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(160px,1fr))',gap:0}}>
                      {[
                        {label:'Publications',prev:ax.pub_prev,real:ax.pub_real},
                        {label:'Projets obtenus',prev:ax.proj_prev,real:ax.proj_real},
                        {label:'Budget (MAD)',prev:ax.budget_prev.toLocaleString('fr-MA'),real:ax.budget_real.toLocaleString('fr-MA'),noEcart:true},
                        {label:'H. Formation initiale',prev:'—',real:ax.h_init_real,noEcart:true},
                        {label:'Doctorants',prev:'—',real:ax.doct_real,noEcart:true},
                      ].map((kpi,j)=>{
                        const ecart = (!kpi.noEcart&&kpi.prev&&kpi.real) ? (+kpi.real-(+kpi.prev)) : null
                        return (
                          <div key={kpi.label} style={{padding:'14px 16px',borderRight:j<4?`0.5px solid ${C.g3}`:'none',borderTop:`0.5px solid ${C.g3}`}}>
                            <p style={{margin:'0 0 8px',fontSize:11,color:C.gt}}>{kpi.label}</p>
                            <div style={{display:'flex',alignItems:'baseline',gap:8}}>
                              <span style={{fontSize:20,fontWeight:700,color:C.gd}}>{kpi.real}</span>
                              {kpi.prev!=='—'&&<span style={{fontSize:11,color:C.gt}}>/ {kpi.prev} prév.</span>}
                            </div>
                            {ecart!==null&&<p style={{margin:'4px 0 0',fontSize:11,color:ecart>=0?C.green:C.red,fontWeight:600}}>
                              {ecart>=0?'+':''}{ecart}
                            </p>}
                          </div>
                        )
                      })}
                    </div>
                  </div>
                )
              })}
            </div>
          )}

          {/* ── TAB 2: Suivi professeurs ── */}
          {adminTab==='profs'&&(
            <div>
              <div style={{marginBottom:14}}>
                <input value={adminSearch} onChange={e=>setAdminSearch(e.target.value)}
                  placeholder={lang==='fr'?'Rechercher un professeur...':'Search professor...'}
                  style={{width:'100%',maxWidth:360,padding:'10px 14px',border:`1.5px solid ${C.g3}`,borderRadius:8,fontSize:14,outline:'none',fontFamily:'inherit',color:C.gd,boxSizing:'border-box'}}/>
              </div>
              <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,overflow:'auto'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:12,minWidth:700}}>
                  <thead>
                    <tr style={{background:C.navy}}>
                      {['Nom','Axe','Année','🎯 Prévu','🔄 Révisé','📊 Bilan','📋 Réalisations','Bilan Excel','Complétude'].map(h=>(
                        <th key={h} style={{padding:'10px 11px',color:'#fff',fontWeight:600,textAlign:'left',fontSize:11,whiteSpace:'nowrap'}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredProfs.length===0
                      ?<tr><td colSpan={9} style={{padding:'32px',textAlign:'center',color:C.gt}}>
                         {adminSearch?'Aucun résultat':'Aucune soumission'}
                       </td></tr>
                      :filteredProfs.map((p,i)=>{
                        const done=[p.prevision,p.revision_s1,p.bilan_annuel,p.realisations].filter(Boolean).length
                        const pct=Math.round(done/4*100)
                        const pc=pct===100?C.green:pct>=50?C.orange:C.red
                        return (
                          <tr key={i} style={{borderBottom:`0.5px solid ${C.g3}`,background:i%2===0?'#fff':C.g1}}>
                            <td style={{padding:'9px 11px',fontWeight:600,color:C.gd}}>{p.nom||'—'}<br/><span style={{fontSize:10,color:C.gt,fontWeight:400}}>{p.email||''}</span></td>
                            <td style={{padding:'9px 11px',color:C.gt,fontSize:10}}>{(p.axe||'').replace('Mine & Mineral Processing (MMP)','MMP').replace('Sustainable Mining & Environment (SME)','SME')}</td>
                            <td style={{padding:'9px 11px',color:C.gt}}>{p.annee||'—'}</td>
                            {['prevision','revision_s1','bilan_annuel','realisations'].map(m=>(
                              <td key={m} style={{padding:'9px 11px',textAlign:'center'}}>
                                {p[m]?<span style={{color:C.green,fontSize:15}}>✅</span>:<span style={{color:C.g3,fontSize:15}}>⏳</span>}
                              </td>
                            ))}
                            <td style={{padding:'9px 11px'}}>
                              <button onClick={()=>handleBilanExport(p.email,p.annee)}
                                style={{padding:'5px 10px',background:C.navy,color:'#fff',border:'none',borderRadius:6,fontSize:11,cursor:'pointer',fontFamily:'inherit',whiteSpace:'nowrap'}}>
                                ⬇ Excel
                              </button>
                            </td>
                            <td style={{padding:'9px 11px'}}>
                              <div style={{display:'flex',alignItems:'center',gap:6}}>
                                <div style={{flex:1,height:5,background:C.g3,borderRadius:3}}>
                                  <div style={{width:`${pct}%`,height:'100%',background:pc,borderRadius:3}}/>
                                </div>
                                <span style={{fontSize:10,color:pc,fontWeight:700,minWidth:28}}>{pct}%</span>
                              </div>
                            </td>
                          </tr>
                        )
                      })}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ── TAB 3: Bilan individuel ── */}
          {adminTab==='bilan'&&(
            <div>
              <div style={{background:'#fff',border:`0.5px solid ${C.g3}`,borderRadius:12,padding:'20px 22px',marginBottom:16}}>
                <h3 style={{margin:'0 0 14px',fontSize:14,fontWeight:700,color:C.navy}}>
                  {lang==='fr'?'Générer le bilan d\'un professeur':'Generate professor review'}
                </h3>
                <div style={{display:'flex',gap:12,flexWrap:'wrap'}}>
                  <input value={adminSearch} onChange={e=>setAdminSearch(e.target.value)}
                    placeholder={lang==='fr'?'Nom ou email du professeur...':'Professor name or email...'}
                    style={{flex:1,minWidth:200,padding:'10px 14px',border:`1.5px solid ${C.g3}`,borderRadius:8,fontSize:14,outline:'none',fontFamily:'inherit',color:C.gd}}/>
                </div>
                {adminSearch&&(
                  <div style={{marginTop:12}}>
                    {filteredProfs.map(p=>(
                      <div key={p.email} style={{display:'flex',alignItems:'center',justifyContent:'space-between',padding:'12px 14px',background:C.g1,borderRadius:8,marginBottom:8}}>
                        <div>
                          <p style={{margin:0,fontWeight:600,fontSize:13,color:C.gd}}>{p.nom} <span style={{fontSize:11,color:C.gt,fontWeight:400}}>— {p.email}</span></p>
                          <p style={{margin:'3px 0 0',fontSize:11,color:C.gt}}>{p.axe} · {p.annee} · Modes : {[p.prevision&&'Prévu',p.revision_s1&&'Révisé',p.bilan_annuel&&'Bilan',p.realisations&&'Réal.'].filter(Boolean).join(', ')||'Aucun'}</p>
                        </div>
                        <button onClick={()=>handleBilanExport(p.email,p.annee)}
                          style={{padding:'9px 16px',background:C.navy,color:'#fff',border:'none',borderRadius:8,fontSize:13,cursor:'pointer',fontFamily:'inherit',whiteSpace:'nowrap'}}>
                          ⬇ Générer Bilan Excel
                        </button>
                      </div>
                    ))}
                    {filteredProfs.length===0&&<p style={{color:C.gt,fontSize:13,margin:0}}>Aucun professeur trouvé</p>}
                  </div>
                )}
              </div>
              <div style={{background:'#EFF6FF',borderRadius:10,padding:'14px 18px',borderLeft:`3px solid ${C.blue}`}}>
                <p style={{margin:0,fontSize:13,color:'#1e40af',lineHeight:1.6}}>
                  <strong>{lang==='fr'?'Contenu du bilan Excel :':'Excel review content:'}</strong><br/>
                  {lang==='fr'
                    ?'Onglet 1 : Dashboard KPI (Prévu / Révisé S1 / Réalisé + Écart C-A + Écart C-B + Statut)\nOnglet 2 : Publications Scopus détaillées\nOnglet 3 : Projets de recherche\nOnglet 4 : Formation & Encadrement\nOnglet 5 : Prestations de service\nOnglet 6 : Note justificative (commentaires + signatures)'
                    :'Tab 1: KPI Dashboard (Forecast/S1 Revision/Achieved + Variance C-A + Variance C-B + Status)\nTab 2: Detailed Scopus publications\nTab 3: Research projects\nTab 4: Training & supervision\nTab 5: Service contracts\nTab 6: Justification note (comments + signatures)'}
                </p>
              </div>
            </div>
          )}

        </div>
        {toast&&<Toast t={toast}/>}
      </div>
    )
  }

  return null
}
