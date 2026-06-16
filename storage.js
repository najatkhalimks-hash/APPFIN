const LOCAL_KEY = 'gsmi_v3'

function getScriptUrl() {
  try { return import.meta.env.VITE_APPS_SCRIPT_URL || null } catch { return null }
}

export async function saveSubmission(data) {
  const all = loadLocal()
  const entry = { ...data, timestamp: new Date().toISOString() }
  // Si même prof + même année + même mode → remplacer
  const idx = all.findIndex(s =>
    s.email === entry.email &&
    s.annee_academique === entry.annee_academique &&
    s.mode === entry.mode
  )
  if (idx >= 0) all[idx] = entry
  else all.push(entry)

  try { localStorage.setItem(LOCAL_KEY, JSON.stringify(all)) } catch {}

  const url = getScriptUrl()
  if (url) {
    try {
      await fetch(url, {
        method: 'POST', mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(entry),
      })
    } catch {}
  }
  return all
}

export async function loadSubmissions() {
  const url = getScriptUrl()
  if (url) {
    try {
      const r = await fetch(url)
      const d = await r.json()
      if (d.rows?.length > 1) {
        const [headers, ...rows] = d.rows
        return rows.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])))
      }
    } catch {}
  }
  return loadLocal()
}

export function loadLocal() {
  try { return JSON.parse(localStorage.getItem(LOCAL_KEY) || '[]') } catch { return [] }
}

export function clearSubmissions() {
  try { localStorage.removeItem(LOCAL_KEY) } catch {}
}
