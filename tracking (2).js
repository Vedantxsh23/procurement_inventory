/* WeRoCon Lab — Procurement & Inventory System
   Shipment tracking.

   Right now: manual status updates (free, works on static GitHub Pages).
   When you get a Shiprocket API key, paste it into CONFIG below and
   set CONFIG.enabled = true. fetchLiveStatus() will then call Shiprocket's
   tracking endpoint instead of relying on manual updates.

   Shiprocket tracking docs: https://apidocs.shiprocket.in/
   (Requires a Shiprocket account; the API needs a small server-side proxy
   in production because Shiprocket does not send CORS headers for
   browser-only calls — a free Cloudflare Worker or Vercel function works.
   Ask Claude to wire this up once you have the key.)
*/

const Tracking = (() => {
  const CONFIG = {
    enabled: false,        // flip to true once you have a working proxy + API key
    apiKey: '',            // Shiprocket API token (do NOT hardcode in a public repo — use env/proxy)
    proxyUrl: ''            // your CORS proxy endpoint, e.g. https://your-worker.workers.dev/track
  };

  const STEPS = ['Order placed', 'Picked up', 'In transit', 'Out for delivery', 'Delivered'];
  const COURIERS = ['Delhivery', 'BlueDart', 'DTDC', 'FedEx', 'Ecom Express', 'India Post', 'Xpressbees', 'Shiprocket', 'Other'];

  function newTracking({ courier, trackingId, status = 'Order placed', expected = '' }) {
    return {
      courier, trackingId, status, expected,
      history: [{ status, at: new Date().toISOString(), source: 'manual' }],
      autoTracked: false
    };
  }

  function updateStatus(tracking, status) {
    tracking.status = status;
    tracking.history = tracking.history || [];
    tracking.history.push({ status, at: new Date().toISOString(), source: 'manual' });
    return tracking;
  }

  // Stub for live tracking - returns null (not connected) until CONFIG.enabled = true
  async function fetchLiveStatus(trackingId, courier) {
    if (!CONFIG.enabled) return null;
    try {
      const res = await fetch(`${CONFIG.proxyUrl}?awb=${encodeURIComponent(trackingId)}&courier=${encodeURIComponent(courier)}`, {
        headers: { 'Authorization': `Bearer ${CONFIG.apiKey}` }
      });
      if (!res.ok) throw new Error('Tracking API error: ' + res.status);
      const data = await res.json();
      // Map Shiprocket's status string to our STEPS vocabulary - adjust once
      // you see real response shape from your account.
      return data.current_status || null;
    } catch (e) {
      console.error('Live tracking fetch failed', e);
      return null;
    }
  }

  function stepIndex(status) {
    return STEPS.indexOf(status);
  }

  return { CONFIG, STEPS, COURIERS, newTracking, updateStatus, fetchLiveStatus, stepIndex };
})();
