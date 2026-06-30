// WeRoCon Lab — TrackCourier.io tracking proxy
// Supabase Edge Function. Keeps the TrackCourier.io API key server-side
// (never exposed in the public GitHub repo / browser) and adds CORS
// headers so the static site can call it.
//
// Deploy with:
//   supabase functions deploy track-shipment
//
// Set the secret once (NOT committed to git):
//   supabase secrets set TRACKCOURIER_API_KEY=tc_live_xxxxxxxxxxxx
//
// TrackCourier.io docs: https://api.trackcourier.io/  (api.trackcourier.io/docs)
//   GET /v1/track?courier={slug}&tracking_number={number}
//   Auth header: X-API-Key: <key>
//
// Courier slugs used by this app (must match js/tracking.js COURIERS list):
//   Delhivery -> delhivery | BlueDart -> bluedart | DTDC -> dtdc
//   India Post -> indiapost | Ecom Express -> ecomexpress
//   FedEx -> fedex | Xpressbees -> xpressbees

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS'
};

const COURIER_SLUG_MAP = {
  'delhivery': 'delhivery',
  'bluedart': 'bluedart',
  'dtdc': 'dtdc',
  'india post': 'indiapost',
  'indiapost': 'indiapost',
  'speed post': 'speedpost',
  'speedpost': 'speedpost',
  'ecom express': 'ecomexpress',
  'ecomexpress': 'ecomexpress',
  'fedex': 'fedex',
  'xpressbees': 'xpressbees',
  'dhl': 'dhl'
};

function toCourierSlug(courierName) {
  const key = (courierName || '').toLowerCase().trim();
  return COURIER_SLUG_MAP[key] || key.replace(/\s+/g, '');
}

// Maps TrackCourier.io's status strings down to the app's 5-step
// pipeline. The exact field/value names returned by the API weren't
// confirmed against a live response while building this (their Swagger
// docs page is JS-rendered and couldn't be statically inspected) —
// every raw response is stored in shipment_tracking.raw_status_payload,
// so if statuses come through unmapped (falling to 'Order placed' by
// default), check that column in Supabase and extend the mapping below.
function mapStatus(raw) {
  const s = JSON.stringify(raw || {}).toUpperCase();
  if (s.includes('DELIVERED')) return 'Delivered';
  if (s.includes('OUT FOR DELIVERY') || s.includes('OUT_FOR_DELIVERY')) return 'Out for delivery';
  if (s.includes('IN TRANSIT') || s.includes('IN_TRANSIT') || s.includes('TRANSIT') || s.includes('SHIPPED')) return 'In transit';
  if (s.includes('PICKED') || s.includes('PICKUP') || s.includes('PICKED_UP')) return 'Picked up';
  if (s.includes('PLACED') || s.includes('BOOKED') || s.includes('MANIFEST')) return 'Order placed';
  return 'Order placed';
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: CORS_HEADERS });
  }

  try {
    const url = new URL(req.url);
    const trackingNumber = url.searchParams.get('awb');
    const courierName = url.searchParams.get('courier');
    if (!trackingNumber) {
      return new Response(JSON.stringify({ error: 'Missing awb (tracking number) parameter' }), {
        status: 400,
        headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
      });
    }

    const apiKey = Deno.env.get('TRACKCOURIER_API_KEY');
    if (!apiKey) {
      return new Response(JSON.stringify({ error: 'TRACKCOURIER_API_KEY secret not set. Run: supabase secrets set TRACKCOURIER_API_KEY=tc_live_...' }), {
        status: 500,
        headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
      });
    }

    const slug = toCourierSlug(courierName);
    const trackUrl = `https://api.trackcourier.io/v1/track?courier=${encodeURIComponent(slug)}&tracking_number=${encodeURIComponent(trackingNumber)}`;

    const trackRes = await fetch(trackUrl, {
      headers: { 'X-API-Key': apiKey }
    });

    if (!trackRes.ok) {
      const bodyText = await trackRes.text();
      return new Response(JSON.stringify({ error: 'TrackCourier.io call failed', status: trackRes.status, body: bodyText }), {
        status: 502,
        headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
      });
    }

    const raw = await trackRes.json();
    const mapped = mapStatus(raw);

    return new Response(JSON.stringify({
      current_status: mapped,
      raw
    }), {
      headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
    });

  } catch (err) {
    return new Response(JSON.stringify({ error: String(err) }), {
      status: 500,
      headers: { ...CORS_HEADERS, 'Content-Type': 'application/json' }
    });
  }
});
