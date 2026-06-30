/* WeRoCon Lab — Procurement & Inventory System
   Configuration.

   SUPABASE_URL and SUPABASE_ANON_KEY are safe to commit to a public
   GitHub repo. They are NOT secret — access control is enforced by
   Row Level Security policies in the database (see supabase_schema.sql),
   not by hiding these values. Never put your Supabase service_role key,
   DB password, or TrackCourier.io API key here.
*/

const APP_CONFIG = {
  SUPABASE_URL: 'https://lzmcblzqoakxfmgexyar.supabase.co',
  SUPABASE_ANON_KEY: 'sb_publishable_Dr-VUuWdf3Yx6PA6iCobhw_cnt0XbEf',
  STORAGE_BUCKET: 'procurement-files',

  // Set once you've deployed the track-shipment Edge Function (see
  // supabase/functions/track-shipment/index.ts and README.md). This
  // function proxies calls to TrackCourier.io (api.trackcourier.io),
  // keeping your tc_live_... API key off the public GitHub repo.
  // Format: https://<project-ref>.supabase.co/functions/v1/track-shipment
  TRACKING_FUNCTION_URL: 'https://lzmcblzqoakxfmgexyar.supabase.co/functions/v1/track-shipment',

  LAB_NAME: 'WeRoCon Lab',
  PROJECT_TITLE_DEFAULT: 'Wearable Robotics Control'
};

const sb = window.supabase.createClient(APP_CONFIG.SUPABASE_URL, APP_CONFIG.SUPABASE_ANON_KEY);
