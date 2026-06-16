import * as XLSX from 'xlsx'

const NAV = '#0D1B2A', BLUE = '#1A56DB', TEAL = '#047481', GREEN = '#057A55'
const VIOLET = '#5521B5', ORANGE = '#B45309', GOLD = '#FBBF24'

function cs(bold, bg='FFFFFF', fg='111928', align='left', sz=10) {
  return {
    font:{ name:'Calibri', bold, sz, color:{rgb:fg} },
    fill:{ patternType:'solid', fgColor:{rgb:bg.replace('#','')} },
    alignment:{ horizontal:align, vertical:'center', wrapText:true },
    border:{ top:{style:'thin',color:{rgb:'E5E7EB'}}, bottom:{style:'thin',color:{rgb:'E5E7EB'}}, left:{style:'thin',color:{rgb:'E5E7EB'}}, right:{style:'thin',color:{rgb:'E5E7EB'}} },
  }
}

function ecartColor(ecart) {
  if (ecart === null || ecart === '') return ''
  return +ecart >= 0 ? GREEN.replace('#','') : 'BE123C'
}

export function generateBilanExcel(prof, prevData, revData, bilanData, realisations) {
  const wb = XLSX.utils.book_new()
  const annee = prevData?.annee_academique || bilanData?.annee_academique || '—'
  const nomProf = prevData?.nom || bilanData?.nom || '—'
  const axe = prevData?.axe_recherche || bilanData?.axe_recherche || '—'
  const date = new Date().toLocaleDateString('fr-MA')

  // ══ SHEET 1: BILAN KPI (Prévu / Révisé S1 / Réalisé / Écart1 / Écart2) ══
  const kpiRows = [
    ['BILAN ANNUEL — CARNET DU CHERCHEUR', '', '', '', '', '', ''],
    [`Professeur : ${nomProf}`, `Année : ${annee}`, `Axe : ${axe}`, `Généré le : ${date}`, '', '', ''],
    [''],
    ['Indicateur KPI', 'Prévu (A)', 'Révisé S1 (B)', 'Réalisé (C)', 'Écart C-A (vs Prévision)', 'Écart C-B (vs Révision S1)', 'Statut'],
    // Production
    ['🔬 PRODUCTION SCIENTIFIQUE', '', '', '', '', '', ''],
    ['Publications totales', prevData?.prev_pub_total||0, revData?.pub_soumises||0, realisations?.real_production?.length||bilanData?.pub_acceptees||0,
     (realisations?.real_production?.length||0)-(prevData?.prev_pub_total||0),
     (realisations?.real_production?.length||0)-(revData?.pub_soumises||0), ''],
    ['Publications Q1/Q2', prevData?.prev_pub_q1q2||0, revData?.pub_q1q2||0, bilanData?.pub_q1q2||0,
     (bilanData?.pub_q1q2||0)-(prevData?.prev_pub_q1q2||0), (bilanData?.pub_q1q2||0)-(revData?.pub_q1q2||0), ''],
    ['Citations', prevData?.prev_citations||0, revData?.citations||0, bilanData?.citations||0,
     (bilanData?.citations||0)-(prevData?.prev_citations||0), (bilanData?.citations||0)-(revData?.citations||0), ''],
    // Projets
    ['💰 PROJETS DE RECHERCHE', '', '', '', '', '', ''],
    ['Projets obtenus', prevData?.prev_projets_obt||0, revData?.projets_obtenus||0, bilanData?.projets_obtenus||0,
     (bilanData?.projets_obtenus||0)-(prevData?.prev_projets_obt||0), (bilanData?.projets_obtenus||0)-(revData?.projets_obtenus||0), ''],
    ['Budget obtenu (MAD)', prevData?.prev_budget||0, revData?.budget_mad||0, bilanData?.budget_mad||0,
     (bilanData?.budget_mad||0)-(prevData?.prev_budget||0), (bilanData?.budget_mad||0)-(revData?.budget_mad||0), ''],
    // Formation
    ['🎓 FORMATION & ENCADREMENT', '', '', '', '', '', ''],
    ['H. Formation initiale', prevData?.prev_h_init||0, revData?.h_initiale||0, bilanData?.h_initiale||0,
     (bilanData?.h_initiale||0)-(prevData?.prev_h_init||0), (bilanData?.h_initiale||0)-(revData?.h_initiale||0), ''],
    ['H. Formation exécutive', prevData?.prev_h_exec||0, revData?.h_executive||0, bilanData?.h_executive||0,
     (bilanData?.h_executive||0)-(prevData?.prev_h_exec||0), (bilanData?.h_executive||0)-(revData?.h_executive||0), ''],
    ['H. Formation doctorale', prevData?.prev_h_doct||0, revData?.h_doctorale||0, bilanData?.h_doctorale||0,
     (bilanData?.h_doctorale||0)-(prevData?.prev_h_doct||0), (bilanData?.h_doctorale||0)-(revData?.h_doctorale||0), ''],
    ['Doctorants encadrés', prevData?.prev_doctorants||0, revData?.doctorants||0, bilanData?.doctorants||0,
     (bilanData?.doctorants||0)-(prevData?.prev_doctorants||0), (bilanData?.doctorants||0)-(revData?.doctorants||0), ''],
    // Prestations
    ['💼 PRESTATIONS DE SERVICE', '', '', '', '', '', ''],
    ['Nb. prestations', prevData?.prev_prestations||0, revData?.nb_presta||0, bilanData?.nb_presta||0,
     (bilanData?.nb_presta||0)-(prevData?.prev_prestations||0), (bilanData?.nb_presta||0)-(revData?.nb_presta||0), ''],
    ['Revenus (MAD)', prevData?.prev_revenus||0, revData?.revenus_mad||0, bilanData?.revenus_mad||0,
     (bilanData?.revenus_mad||0)-(prevData?.prev_revenus||0), (bilanData?.revenus_mad||0)-(revData?.revenus_mad||0), ''],
    // Rayonnement
    ['🌍 RAYONNEMENT', '', '', '', '', '', ''],
    ['Conférences internationales', prevData?.prev_conf_int||0, revData?.conferences_int||0, bilanData?.conferences_int||0,
     (bilanData?.conferences_int||0)-(prevData?.prev_conf_int||0), (bilanData?.conferences_int||0)-(revData?.conferences_int||0), ''],
    ['Brevets déposés', prevData?.prev_brevets||0, revData?.brevets_deposes||0, bilanData?.brevets_deposes||0,
     (bilanData?.brevets_deposes||0)-(prevData?.prev_brevets||0), (bilanData?.brevets_deposes||0)-(revData?.brevets_deposes||0), ''],
  ]

  // Add statut column
  kpiRows.forEach((row, i) => {
    if (i < 4 || !row[3] || row[0].startsWith('🔬') || row[0].startsWith('💰') || row[0].startsWith('🎓') || row[0].startsWith('💼') || row[0].startsWith('🌍')) return
    const prev = +row[1], real = +row[3]
    if (!prev) { row[6] = 'N/A'; return }
    const pct = Math.round((real / prev) * 100)
    row[6] = real >= prev ? '✅ Atteint' : real >= prev * 0.75 ? '🟡 Partiel' : '🔴 Non atteint'
  })

  const ws1 = XLSX.utils.aoa_to_sheet(kpiRows)

  // Styles
  const sectionRows = [4, 8, 11, 15, 17, 19]
  const secColors = { 4:'047481', 8:'057A55', 11:'5521B5', 15:'B45309', 17:'B45309', 19:'5521B5' }

  for (let r = 0; r < kpiRows.length; r++) {
    const rowData = kpiRows[r]
    for (let c = 0; c < 7; c++) {
      const addr = XLSX.utils.encode_cell({r, c})
      if (!ws1[addr]) ws1[addr] = {v:'',t:'s'}

      if (r === 0) { ws1[addr].s = cs(true, '0D1B2A', 'FFFFFF', 'left', 14); continue }
      if (r === 1) { ws1[addr].s = cs(false, '1B2A3B', 'FBBF24', 'left', 9); continue }
      if (r === 3) { ws1[addr].s = cs(true, '0D1B2A', 'FFFFFF', 'center', 10); continue }

      if (sectionRows.includes(r)) {
        ws1[addr].s = cs(true, secColors[r] || '374151', 'FFFFFF', 'left', 10); continue
      }

      const bg = r % 2 === 0 ? 'F9FAFB' : 'FFFFFF'
      if (c === 0) { ws1[addr].s = cs(false, bg, '111928', 'left'); continue }
      if (c === 1) { ws1[addr].s = cs(true, bg, '1A56DB', 'center'); continue }
      if (c === 2) { ws1[addr].s = cs(true, bg, '047481', 'center'); continue }
      if (c === 3) { ws1[addr].s = cs(true, bg, '5521B5', 'center'); continue }
      if (c === 4) {
        const v = +rowData[4]
        ws1[addr].s = cs(true, bg, isNaN(v) ? '374151' : v >= 0 ? '057A55' : 'BE123C', 'center')
        continue
      }
      if (c === 5) {
        const v = +rowData[5]
        ws1[addr].s = cs(true, bg, isNaN(v) ? '374151' : v >= 0 ? '047481' : 'C27803', 'center')
        continue
      }
      if (c === 6) {
        const v = String(rowData[6]||'')
        const fg = v.includes('✅') ? '057A55' : v.includes('🟡') ? 'B45309' : v.includes('🔴') ? 'BE123C' : '374151'
        ws1[addr].s = cs(true, bg, fg, 'center')
        continue
      }
    }
  }

  ws1['!cols'] = [{wch:30},{wch:14},{wch:14},{wch:14},{wch:20},{wch:20},{wch:16}]
  ws1['!merges'] = [
    {s:{r:0,c:0},e:{r:0,c:6}},
    {s:{r:1,c:0},e:{r:1,c:3}},
  ]
  XLSX.utils.book_append_sheet(wb, ws1, '📊 Bilan KPI')

  // ══ SHEET 2: Publications détaillées ══
  if (realisations?.real_production?.length) {
    const pubH = ['Auteur(s)','Titre','Année','Semestre','Source','Volume/Pages','Citations','Type doc','Statut publication','DOI','Open Access','Affiliations','Commentaire']
    const pubRows = realisations.real_production.map(r => [
      r.authors||'', r.title||'', r.year||'', r.semester||'',
      r.source||'', r.vol_pages||'', r.citations||0,
      r.doc_type||'', r.stage||'', r.doi||'', r.open_access||'', r.affiliations||'', r.comment||''
    ])
    const ws2 = XLSX.utils.aoa_to_sheet([pubH, ...pubRows])
    const r2 = XLSX.utils.decode_range(ws2['!ref'])
    for (let c = r2.s.c; c <= r2.e.c; c++) {
      const a = XLSX.utils.encode_cell({r:0,c})
      if (ws2[a]) ws2[a].s = cs(true,'047481','FFFFFF','center')
    }
    ws2['!cols'] = [{wch:22},{wch:36},{wch:8},{wch:8},{wch:24},{wch:16},{wch:10},{wch:18},{wch:18},{wch:26},{wch:14},{wch:20},{wch:18}]
    XLSX.utils.book_append_sheet(wb, ws2, '🔬 Publications')
  }

  // ══ SHEET 3: Projets détaillés ══
  if (realisations?.real_projets?.length) {
    const projH = ['Année','Type','Intitulé','Rôle','Statut','Financeur','Budget (MAD)','Part UM6P (%)','Début','Fin','Multi-inst.','Commentaire']
    const projRows = realisations.real_projets.map(r => [r.annee||'',r.type||'',r.intitule||'',r.role||'',r.statut||'',r.financeur||'',r.budget||0,r.part_um6p||'',r.debut||'',r.fin||'',r.multi_inst||'',r.comment||''])
    const ws3 = XLSX.utils.aoa_to_sheet([projH,...projRows])
    const r3 = XLSX.utils.decode_range(ws3['!ref'])
    for (let c = r3.s.c; c <= r3.e.c; c++) { const a=XLSX.utils.encode_cell({r:0,c}); if(ws3[a]) ws3[a].s=cs(true,'057A55','FFFFFF','center') }
    ws3['!cols'] = Array(projH.length).fill({wch:18})
    XLSX.utils.book_append_sheet(wb, ws3, '💰 Projets')
  }

  // ══ SHEET 4: Formation ══
  if (realisations?.real_formation?.length) {
    const formH = ['Semestre','Type formation','Activité','Filière','H. Prévues','H. Réalisées','Commentaire']
    const formRows = realisations.real_formation.map(r => [r.semester||'',r.type_form||'',r.activite||'',r.filiere||'',r.h_prev||0,r.h_real||0,r.comment||''])
    const ws4 = XLSX.utils.aoa_to_sheet([formH,...formRows])
    const r4 = XLSX.utils.decode_range(ws4['!ref'])
    for (let c = r4.s.c; c <= r4.e.c; c++) { const a=XLSX.utils.encode_cell({r:0,c}); if(ws4[a]) ws4[a].s=cs(true,'5521B5','FFFFFF','center') }
    ws4['!cols'] = [{wch:10},{wch:20},{wch:22},{wch:22},{wch:12},{wch:14},{wch:20}]
    XLSX.utils.book_append_sheet(wb, ws4, '🎓 Formation')
  }

  // ══ SHEET 5: Prestations ══
  if (realisations?.real_prestations?.length) {
    const presH = ['Année','Intitulé','Type','Client','Rôle','Jours','Tarif/j (MAD)','Montant (MAD)','Statut','Début','Fin','Impact','Commentaire']
    const presRows = realisations.real_prestations.map(r=>[r.annee||'',r.intitule||'',r.type||'',r.client||'',r.role||'',r.jours||0,r.tarif||0,r.montant||0,r.statut||'',r.debut||'',r.fin||'',r.impact||'',r.comment||''])
    const ws5 = XLSX.utils.aoa_to_sheet([presH,...presRows])
    const r5 = XLSX.utils.decode_range(ws5['!ref'])
    for (let c = r5.s.c; c <= r5.e.c; c++) { const a=XLSX.utils.encode_cell({r:0,c}); if(ws5[a]) ws5[a].s=cs(true,'B45309','FFFFFF','center') }
    ws5['!cols'] = Array(presH.length).fill({wch:16})
    XLSX.utils.book_append_sheet(wb, ws5, '💼 Prestations')
  }

  // ══ SHEET 6: Note justificative ══
  const noteRows = [
    ['NOTE DE BILAN ANNUEL — CARNET DU CHERCHEUR GSMI',''],
    [''],
    ['Professeur :', nomProf],
    ['Axe de recherche :', axe],
    ['Année académique :', annee],
    ['Grade :', prevData?.grade || bilanData?.grade || '—'],
    ['Date de génération :', date],
    [''],
    ['OBJECTIFS ANNUELS (Prévisions initiales) :', prevData?.objectifs_texte || '—'],
    [''],
    ['FAITS MARQUANTS :', bilanData?.faits_marq || '—'],
    [''],
    ['JUSTIFICATION DES ÉCARTS :', bilanData?.justif_ecarts || '—'],
    [''],
    ['ACTIONS CORRECTIVES :', bilanData?.actions_corr || '—'],
    [''],
    ['BESOINS DE SUPPORT :', bilanData?.besoins || '—'],
    [''],
    ['PERSPECTIVES :', bilanData?.perspectives || '—'],
    [''],
    ['Statut objectifs :', bilanData?.statut_obj || '—'],
    [''],
    ['Signature du Professeur : ___________________', 'Date : ___________'],
    ['Visa Direction GSMI : ___________________', 'Date : ___________'],
  ]
  const ws6 = XLSX.utils.aoa_to_sheet(noteRows)
  ws6['A1'].s = cs(true,'0D1B2A','FFFFFF','left',14)
  noteRows.forEach((_,i) => {
    const a=`A${i+1}`, b=`B${i+1}`
    if (ws6[a]?.s) return
    if (ws6[a]) ws6[a].s = cs(ws6[a]?.v && String(ws6[a].v).endsWith(':'), 'F9FAFB','111928','left')
    if (ws6[b]) ws6[b].s = cs(false,'FFFFFF','374151','left')
  })
  ws6['!cols'] = [{wch:34},{wch:60}]
  XLSX.utils.book_append_sheet(wb, ws6, '📝 Note justificative')

  const fn = `GSMI_Bilan_${nomProf.replace(/\s/g,'_')}_${annee.replace('/','_')}.xlsx`
  XLSX.writeFile(wb, fn)
  return fn
}
