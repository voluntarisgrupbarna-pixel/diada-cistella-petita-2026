// Vercel Edge Function — proxy per l'endpoint de llista d'inscrits
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx8OTqxo40s-BVl6gZO3i-KqgDDyBfHAslMQ2HkBuLxtHd9NfKFzKirL-iAR_oRB3j8/exec';

export default async function handler(req, res) {
  try {
    const response = await fetch(APPS_SCRIPT_URL + '?action=list&callback=x', {
      redirect: 'follow',
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    const text = await response.text();
    const json = text.replace(/^x\(/, '').replace(/\)$/, '');
    const data = JSON.parse(json);

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 's-maxage=30, stale-while-revalidate');
    res.setHeader('Content-Type', 'application/json');
    res.status(200).json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
