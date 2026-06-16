/**
 * Service d'accusé de réception par email
 * 
 * Utilise EmailJS (gratuit jusqu'à 200 emails/mois) ou un Google Apps Script
 * 
 * Configuration EmailJS :
 * 1. Créer un compte sur emailjs.com
 * 2. Ajouter un service email (Gmail, Outlook…)
 * 3. Créer un template avec ces variables : {{to_email}}, {{prof_name}}, {{mode}}, {{annee}}, {{timestamp}}
 * 4. Copier Service ID, Template ID, Public Key dans .env
 */

const EMAILJS_SERVICE_ID  = import.meta.env.VITE_EMAILJS_SERVICE_ID  || null
const EMAILJS_TEMPLATE_ID = import.meta.env.VITE_EMAILJS_TEMPLATE_ID || null
const EMAILJS_PUBLIC_KEY  = import.meta.env.VITE_EMAILJS_PUBLIC_KEY  || null

const MODE_LABELS = {
  prevision:    'Saisie des prévisions annuelles',
  revision_s1:  'Révision semestrielle S1',
  bilan_annuel: 'Bilan annuel',
  realisations: 'Réalisations détaillées',
}

export async function sendAcknowledgement(email, name, mode, annee) {
  if (!email || !EMAILJS_SERVICE_ID) {
    console.log(`[EmailJS non configuré] Accusé de réception pour : ${email} — Mode: ${mode}`)
    return { ok: true, simulated: true }
  }

  try {
    // Charger EmailJS dynamiquement
    const emailjs = await import('https://cdn.jsdelivr.net/npm/@emailjs/browser@4/dist/email.min.js').catch(() => null)
    if (!emailjs) throw new Error('EmailJS not loaded')

    const result = await emailjs.default.send(
      EMAILJS_SERVICE_ID,
      EMAILJS_TEMPLATE_ID,
      {
        to_email:   email,
        prof_name:  name,
        mode_label: MODE_LABELS[mode] || mode,
        annee,
        timestamp:  new Date().toLocaleString('fr-MA'),
        reply_to:   'direction-recherche@gsmi.um6p.ma',
      },
      EMAILJS_PUBLIC_KEY
    )
    return { ok: true, result }
  } catch (err) {
    console.warn('Email send failed:', err)
    return { ok: false, error: err.message }
  }
}

// Alternative : Google Apps Script webhook
export async function sendViaAppsScript(scriptUrl, email, name, mode, annee) {
  if (!scriptUrl) return { ok: false }
  try {
    await fetch(scriptUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'send_ack',
        to: email,
        name,
        mode: MODE_LABELS[mode] || mode,
        annee,
        timestamp: new Date().toISOString(),
      }),
    })
    return { ok: true }
  } catch (e) {
    return { ok: false, error: e.message }
  }
}
