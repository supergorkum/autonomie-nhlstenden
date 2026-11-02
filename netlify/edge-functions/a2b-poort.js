// Slot van A.I. 2 Business voor deze demo. Vervangt het Netlify-sitewachtwoord.
// Wie op de portfolio (https://edgar-van-gorkum.netlify.app) is ingelogd en daar
// op de demo klikt, krijgt een doorgangspas mee in de link (?a2b=...). Deze
// functie laat de portfolio die pas controleren, zet hem als koekje (12 uur, of
// korter voor een gast) en haalt hem uit de adresbalk. Zonder geldige pas gaat
// de bezoeker naar de portfolio om in te loggen.
// Bron: portfolio-ai2business/demo-slot/a2b-poort.js. Pas DEMO per site aan.
// Let op: de serverfuncties onder /.netlify/functions/ vallen hier buiten.

const PORTFOLIO = 'https://edgar-van-gorkum.netlify.app';
const DEMO = 'autonomie';
const uitkomsten = new Map();

async function geldig(pas, modus) {
  const sleutel = `${modus}:${pas}`;
  const bekend = uitkomsten.get(sleutel);
  if (bekend && bekend.tot > Date.now()) return bekend.ok;
  let ok = false;
  try {
    const r = await fetch(`${PORTFOLIO}/api/pas-check`, {
      method: 'POST',
      body: new URLSearchParams({ pas, demo: DEMO, modus }),
    });
    ok = r.ok && (await r.json()).ok === true;
  } catch {
    ok = false;
  }
  if (uitkomsten.size > 500) uitkomsten.clear();
  uitkomsten.set(sleutel, { ok, tot: Date.now() + (ok ? 5 * 60 * 1000 : 20 * 1000) });
  return ok;
}

export default async (request, context) => {
  const url = new URL(request.url);

  const nieuw = url.searchParams.get('a2b');
  if (nieuw && (await geldig(nieuw, 'link'))) {
    url.searchParams.delete('a2b');
    return new Response(null, {
      status: 302,
      headers: {
        Location: url.toString(),
        'Set-Cookie': `a2b=${nieuw}; Path=/; Max-Age=43200; HttpOnly; Secure; SameSite=Lax`,
        'Cache-Control': 'no-store',
      },
    });
  }

  const koekje = context.cookies.get('a2b');
  if (koekje && (await geldig(koekje, 'sessie'))) return;

  const pagina = request.method === 'GET' && (request.headers.get('accept') || '').includes('text/html');
  if (!pagina) {
    return new Response('Geen toegang. Log in via de portfolio.', { status: 401, headers: { 'Cache-Control': 'no-store' } });
  }
  url.searchParams.delete('a2b');
  const naar = `${PORTFOLIO}/api/doorgang?naar=${DEMO}&pad=${encodeURIComponent(url.pathname + url.search)}`;
  return new Response(null, {
    status: 302,
    headers: { Location: naar, 'Cache-Control': 'no-store', 'Set-Cookie': 'a2b=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax' },
  });
};

export const config = { path: '/*' };
