import * as XLSX from 'xlsx'
import { ALL_FIELD_IDS } from './sections.js'

function s(bold, bg = 'FFFFFF', fg = '111928', align = 'left') {
  return {
    font: { name: 'Calibri', bold, sz: 10, color: { rgb: fg } },
    fill: { patternType: 'solid', fgColor: { rgb: bg } },
    alignment: { horizontal: align, vertical: 'center', wrapText: true },
    border: { top: { style: 'thin', color: { rgb: 'E5E7EB' } }, bottom: { style: 'thin', color: { rgb: 'E5E7EB' } }, left: { style: 'thin', color: { rgb: 'E5E7EB' } }, right: { style: 'thin', color: { rgb: 'E5E7EB' } } },
  }
}

export function exportToExcel(submissions) {
  const wb = XLSX.utils.book_new()
  const date = new Date().toISOString().slice(0, 10)

  // Filter by mode
  const prev   = submissions.filter(s => s.mode === 'prevision')
  const rev    = submissions.filter(s => s.mode === 'revision_s1')
  const bilan  = submissions.filter(s => s.mode === 'bilan_annuel')
  const all    = submissions

  // ── Sheet 1: Dashboard KPI ──────────────────────────────────────────
  function sum(arr, key) { return arr.reduce((a, s) => a + (+s[key] || 0), 0) }

  const kpiData = [
    ['GSMI — Dashboard consolidé', '', '', '', ''],
    [`Exporté le ${new Date().toLocaleDateString('fr-MA')}`, '', '', '', ''],
    [''],
    ['Indicateur', 'Prévisions', 'Révisions S1', 'Réalisé (Bilan)', 'Écart Prév→Réalisé'],
    ['Réponses reçues', prev.length, rev.length, bilan.length, ''],
    ['Publications acceptées', sum(prev,'prev_pub_total'), sum(rev,'pub_acceptees'), sum(bilan,'pub_acceptees'), sum(bilan,'pub_acceptees')-sum(prev,'prev_pub_total')],
    ['Publications Q1/Q2', sum(prev,'prev_pub_q1q2'), sum(rev,'pub_q1q2'), sum(bilan,'pub_q1q2'), sum(bilan,'pub_q1q2')-sum(prev,'prev_pub_q1q2')],
    ['Citations totales', sum(prev,'prev_citations'), sum(rev,'citations'), sum(bilan,'citations'), sum(bilan,'citations')-sum(prev,'prev_citations')],
    ['Projets obtenus', sum(prev,'prev_projets_obt'), sum(rev,'projets_obtenus'), sum(bilan,'projets_obtenus'), sum(bilan,'projets_obtenus')-sum(prev,'prev_projets_obt')],
    ['Budget obtenu (MAD)', sum(prev,'prev_budget'), sum(rev,'budget_mad'), sum(bilan,'budget_mad'), sum(bilan,'budget_mad')-sum(prev,'prev_budget')],
    ['Conférences internationales', sum(prev,'prev_conf_int'), sum(rev,'conferences_int'), sum(bilan,'conferences_int'), ''],
    ['Brevets déposés', sum(prev,'prev_brevets'), sum(rev,'brevets_deposes'), sum(bilan,'brevets_deposes'), ''],
    ['H. Formation initiale', sum(prev,'prev_h_init'), sum(rev,'h_initiale'), sum(bilan,'h_initiale'), sum(bilan,'h_initiale')-sum(prev,'prev_h_init')],
    ['H. Formation exécutive', sum(prev,'prev_h_exec'), sum(rev,'h_executive'), sum(bilan,'h_executive'), ''],
    ['H. Formation doctorale', sum(prev,'prev_h_doct'), sum(rev,'h_doctorale'), sum(bilan,'h_doctorale'), ''],
    ['Doctorants encadrés', sum(prev,'prev_doctorants'), sum(rev,'doctorants'), sum(bilan,'doctorants'), ''],
    ['Prestations de service', sum(prev,'prev_prestations'), sum(rev,'nb_presta'), sum(bilan,'nb_presta'), ''],
    ['Revenus prestations (MAD)', sum(prev,'prev_revenus'), sum(rev,'revenus_mad'), sum(bilan,'revenus_mad'), ''],
  ]

  const ws1 = XLSX.utils.aoa_to_sheet(kpiData)
  ws1['A1'].s = s(true, '0D1B2A', 'FFFFFF', 'left')
  ws1['A4'].s = s(true, '0D1B2A', 'FFFFFF', 'center')
  ;['B4','C4','D4','E4'].forEach(c => { if (ws1[c]) ws1[c].s = s(true, '0D1B2A', 'FFFFFF', 'center') })
  for (let r = 4; r < kpiData.length; r++) {
    ['A','B','C','D','E'].forEach(c => {
      const addr = `${c}${r+1}`
      if (!ws1[addr]) ws1[addr] = { v: '', t: 's' }
      const bg = r % 2 === 0 ? 'F9FAFB' : 'FFFFFF'
      ws1[addr].s = s(r === 4, bg, '111928', c === 'A' ? 'left' : 'center')
    })
  }
  ws1['!cols'] = [{ wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }]
  XLSX.utils.book_append_sheet(wb, ws1, '📊 Dashboard KPI')

  // ── Sheet 2: Données brutes ─────────────────────────────────────────
  const headers = ['Mode', 'Horodatage', ...ALL_FIELD_IDS]
  const rows = all.map(s => [s.mode || '', s.timestamp || '', ...ALL_FIELD_IDS.map(id => s[id] ?? '')])
  const ws2 = XLSX.utils.aoa_to_sheet([headers, ...rows])
  const r2 = XLSX.utils.decode_range(ws2['!ref'])
  for (let c = r2.s.c; c <= r2.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c })
    if (ws2[addr]) ws2[addr].s = s(true, '0D1B2A', 'FFFFFF', 'center')
  }
  ws2['!cols'] = Array(headers.length).fill({ wch: 18 })
  XLSX.utils.book_append_sheet(wb, ws2, '📋 Données brutes')

  // ── Sheet 3: Suivi réponses ─────────────────────────────────────────
  const suiviH = ['Nom', 'Email', 'Grade', 'Axe', 'Année acad.', 'Prévision', 'Révision S1', 'Bilan annuel', 'Statut global']
  const profs = {}
  all.forEach(s => {
    const k = s.email || s.nom
    if (!profs[k]) profs[k] = { nom: s.nom, email: s.email, grade: s.grade, axe: s.axe_recherche, annee: s.annee_academique }
    profs[k][s.mode] = '✅ ' + new Date(s.timestamp).toLocaleDateString('fr-MA')
  })
  const suiviRows = Object.values(profs).map(p => [
    p.nom, p.email, p.grade, p.axe, p.annee,
    p.prevision || '⏳ En attente',
    p.revision_s1 || '⏳ En attente',
    p.bilan_annuel || '⏳ En attente',
    p.prevision && p.revision_s1 && p.bilan_annuel ? '✅ Complet' : '🔄 Incomplet',
  ])
  const ws3 = XLSX.utils.aoa_to_sheet([suiviH, ...suiviRows])
  const r3 = XLSX.utils.decode_range(ws3['!ref'])
  for (let c = r3.s.c; c <= r3.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c })
    if (ws3[addr]) ws3[addr].s = s(true, '057A55', 'FFFFFF', 'center')
  }
  ws3['!cols'] = [{ wch: 22 }, { wch: 24 }, { wch: 20 }, { wch: 24 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 14 }]
  XLSX.utils.book_append_sheet(wb, ws3, '✅ Suivi réponses')

  XLSX.writeFile(wb, `GSMI_Consolidation_${date}.xlsx`)
}
