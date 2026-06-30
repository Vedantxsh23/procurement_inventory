/* WeRoCon Lab — Procurement & Inventory System
   Shipment tracking constants and helpers.

   Right now: manual status updates always work (free, no setup).
   Once the track-shipment Supabase Edge Function is deployed with your
   TrackCourier.io API key (see README.md), the "Sync now" button on the
   tracking tab pulls live status — that call lives in db.js
   (DB.syncTrackingLive), which hits the Edge Function. This file just
   holds the courier list shown in the UI dropdown and the 5-step status
   pipeline used throughout the app.

   TrackCourier.io docs: https://api.trackcourier.io/
   (171+ couriers, strong India coverage — BlueDart, Delhivery, DTDC,
   India Post, Ecom Express, etc. One API key, courier passed as a slug.)
*/

const Tracking = (() => {
  const STEPS = ['Order placed', 'Picked up', 'In transit', 'Out for delivery', 'Delivered'];

  // Must match COURIER_SLUG_MAP in supabase/functions/track-shipment/index.ts
  // for auto-sync to work. Add new couriers in both places.
  const COURIERS = ['Delhivery', 'BlueDart', 'DTDC', 'India Post', 'Speed Post', 'Ecom Express', 'FedEx', 'Xpressbees', 'DHL', 'Other'];

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

  function stepIndex(status) {
    return STEPS.indexOf(status);
  }

  return { STEPS, COURIERS, newTracking, updateStatus, stepIndex };
})();
