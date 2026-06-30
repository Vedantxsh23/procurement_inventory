/* WeRoCon Lab — Procurement & Inventory System
   Data layer — backed by Supabase Postgres. Real permanent delete
   (no localStorage re-seeding), shared across every device/browser,
   with realtime updates so Dr. Saurav's screen reflects changes live.
*/

const DB = (() => {

  // ---------- components ----------
  async function listComponents() {
    const { data, error } = await sb.from('components').select('*').order('created_at', { ascending: false });
    if (error) { console.error(error); return []; }
    return data;
  }

  async function addComponent(comp) {
    const { data, error } = await sb.from('components').insert({
      name: comp.name,
      category: comp.category || 'Other',
      item_type: comp.itemType || 'Non-Recurring',
      qty: Number(comp.qty) || 1,
      unit_price: Number(comp.unitPrice) || 0,
      vendor: comp.vendor || '',
      gem_status: comp.gemStatus || 'Not checked',
      gem_search_ref: comp.gemSearchRef || '',
      invoice_no: comp.invoiceNo || '',
      payment_status: comp.paymentStatus || 'Pending',
      payment_method: comp.paymentMethod || 'Not selected',
      remarks: comp.remarks || '',
      pi_name: comp.piName || '',
      project_title: comp.projectTitle || '',
      project_no: comp.projectNo || '',
      created_by: comp.createdBy || ''
    }).select().single();
    if (error) { console.error(error); throw error; }
    return data;
  }

  async function updateComponent(id, patch) {
    // map camelCase -> snake_case for the known fields we allow editing
    const map = {
      name: 'name', category: 'category', itemType: 'item_type', qty: 'qty',
      unitPrice: 'unit_price', vendor: 'vendor', gemStatus: 'gem_status',
      gemSearchRef: 'gem_search_ref', invoiceNo: 'invoice_no',
      paymentStatus: 'payment_status', paymentMethod: 'payment_method',
      remarks: 'remarks', piName: 'pi_name', projectTitle: 'project_title',
      projectNo: 'project_no'
    };
    const payload = {};
    Object.keys(patch).forEach(k => { if (map[k]) payload[map[k]] = patch[k]; });
    const { data, error } = await sb.from('components').update(payload).eq('id', id).select().single();
    if (error) { console.error(error); throw error; }
    return data;
  }

  // Real permanent delete — cascades to files/tracking/history via FK ON DELETE CASCADE
  async function deleteComponent(id) {
    const { error } = await sb.from('components').delete().eq('id', id);
    if (error) { console.error(error); throw error; }
    return true;
  }

  // ---------- files ----------
  async function uploadFile(componentId, fileType, file, uploadedBy) {
    const path = `${componentId}/${fileType}_${Date.now()}_${file.name}`;
    const { error: upErr } = await sb.storage.from(APP_CONFIG.STORAGE_BUCKET).upload(path, file, { upsert: true });
    if (upErr) { console.error(upErr); throw upErr; }

    const { data, error } = await sb.from('component_files').insert({
      component_id: componentId,
      file_type: fileType,
      file_name: file.name,
      storage_path: path,
      uploaded_by: uploadedBy || ''
    }).select().single();
    if (error) { console.error(error); throw error; }
    return data;
  }

  function getFileUrl(storagePath) {
    const { data } = sb.storage.from(APP_CONFIG.STORAGE_BUCKET).getPublicUrl(storagePath);
    return data.publicUrl;
  }

  async function listFiles(componentId) {
    const { data, error } = await sb.from('component_files').select('*').eq('component_id', componentId).order('uploaded_at', { ascending: false });
    if (error) { console.error(error); return []; }
    return data;
  }

  async function deleteFile(fileId, storagePath) {
    await sb.storage.from(APP_CONFIG.STORAGE_BUCKET).remove([storagePath]);
    const { error } = await sb.from('component_files').delete().eq('id', fileId);
    if (error) { console.error(error); throw error; }
    return true;
  }

  // ---------- tracking ----------
  async function getTracking(componentId) {
    const { data, error } = await sb.from('shipment_tracking').select('*').eq('component_id', componentId).maybeSingle();
    if (error) { console.error(error); return null; }
    return data;
  }

  async function listAllTracking() {
    const { data, error } = await sb.from('shipment_tracking').select('*, components(name, id)');
    if (error) { console.error(error); return []; }
    return data;
  }

  async function upsertTracking(componentId, { courier, trackingId, status, expected }) {
    const { data, error } = await sb.from('shipment_tracking').upsert({
      component_id: componentId,
      courier, tracking_id: trackingId, status,
      expected_delivery: expected || null
    }, { onConflict: 'component_id' }).select().single();
    if (error) { console.error(error); throw error; }

    await sb.from('tracking_history').insert({ component_id: componentId, status, source: 'manual' });
    return data;
  }

  async function syncTrackingLive(componentId, trackingId, courier) {
    if (!APP_CONFIG.TRACKING_FUNCTION_URL) return { error: 'TRACKING_FUNCTION_URL not set in config.js' };
    try {
      const res = await fetch(`${APP_CONFIG.TRACKING_FUNCTION_URL}?awb=${encodeURIComponent(trackingId)}&courier=${encodeURIComponent(courier)}`, {
        headers: { 'Authorization': `Bearer ${APP_CONFIG.SUPABASE_ANON_KEY}` }
      });
      const result = await res.json().catch(() => ({}));
      if (!res.ok) {
        console.error('Edge Function error', res.status, result);
        return { error: result.error || `Edge Function returned HTTP ${res.status}`, status: res.status, detail: result };
      }
      if (result.current_status) {
        await sb.from('shipment_tracking').update({
          status: result.current_status,
          last_synced_at: new Date().toISOString(),
          auto_tracked: true,
          raw_status_payload: result.raw || null
        }).eq('component_id', componentId);
        await sb.from('tracking_history').insert({ component_id: componentId, status: result.current_status, source: 'trackcourier' });
      }
      return result;
    } catch (e) {
      console.error('Live tracking sync failed (network/fetch error)', e);
      return { error: 'Network error reaching Edge Function: ' + e.message };
    }
  }

  // ---------- document log ----------
  async function logDocument(docType, refNo, componentIds, generatedBy) {
    const { error } = await sb.from('document_log').insert({
      doc_type: docType, ref_no: refNo, component_ids: componentIds, generated_by: generatedBy || ''
    });
    if (error) console.error(error);
  }

  // ---------- realtime subscriptions ----------
  function subscribeToChanges(onChange) {
    const channel = sb.channel('components-and-tracking')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'components' }, onChange)
      .on('postgres_changes', { event: '*', schema: 'public', table: 'shipment_tracking' }, onChange)
      .subscribe();
    return () => sb.removeChannel(channel);
  }

  return {
    listComponents, addComponent, updateComponent, deleteComponent,
    uploadFile, getFileUrl, listFiles, deleteFile,
    getTracking, listAllTracking, upsertTracking, syncTrackingLive,
    logDocument, subscribeToChanges
  };
})();
