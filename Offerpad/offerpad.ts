/// <reference types="switch-scripting" />
// @ts-nocheck
// Enfocus Switch Script: Offerpad letter + postcard hydrate -> OMS order
//
// Flow:
//   1. Build a FES hydrate job for the LETTER template and one for the POSTCARD
//      template, POSTing to https://api.oneilprint.com/api/fes/hydrate/jobs/create
//      with the FES API key as the x-api-key header. Each response returns a
//      data.url that is the hydrated artwork.
//   2. Submit ONE OMS order to https://api.oneilprint.com/api/external/create-any
//      (x-api-key = ONeil API key) containing two orderItems:
//        - product_sku "offerpad_letter"   with artworkSingle = letter hydrate url
//        - product_sku "offerpad_postcard" with artworkSingle = postcard hydrate url
//      shipped to the recipient's property address.
//
// Job private data consumed (read by key, case-insensitive / underscore-tolerant):
//   first_name, last_name, property_street_address,
//   property_city, property_state, property_zip_code

// ---------------------------------------------------------------------------
// Defaults for required create-any fields that are NOT user-input parameters.
// Adjust these constants if Offerpad mailers should use a different shipping
// method or order-id convention.
// ---------------------------------------------------------------------------
var SHIPPING_METHOD = 'USPS_BULK';            // required by create-any; change to your mailing method code
var LETTER_PRODUCT_SKU = 'offerpad_letter';
var POSTCARD_PRODUCT_SKU = 'offerpad_postcard';
var ITEM_QUANTITY = 1;                    // units per order item

// FES template field names (left) sourced from private data keys (right).
var FES_FIELD_MAP = [
  { fieldName: 'First Name',     keys: ['first_name', 'firstname', 'first name'] },
  { fieldName: 'Last Name',      keys: ['last_name', 'lastname', 'last name'] },
  { fieldName: 'Street Address', keys: ['property_street_address', 'street_address', 'street address', 'address1', 'address'] },
  { fieldName: 'City',           keys: ['property_city', 'city'] },
  { fieldName: 'State',          keys: ['property_state', 'state'] },
  { fieldName: 'Zip',            keys: ['property_zip_code', 'property_zip', 'zip_code', 'zip', 'postal_code'] },
  { fieldName: 'Webhook ID',     keys: ['webhook_id'] }
];

function getScriptTimeout() {
  return 300; // seconds
}

function getSettingsDefinition() {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<settings>' +
      '<setting name="fes_api_key" displayName="FES API Key" type="string" required="Yes">' +
        '<description>x-api-key sent to the FES hydrate endpoint (create hydrate jobs)</description>' +
      '</setting>' +
      '<setting name="oneil_api_key" displayName="ONeil API Key" type="string" required="Yes">' +
        '<description>x-api-key sent to the OMS create-any endpoint</description>' +
      '</setting>' +
      '<setting name="letter_template_id" displayName="Letter Template ID" type="string" required="Yes">' +
        '<description>FES template id used to hydrate the offerpad_letter artwork</description>' +
      '</setting>' +
      '<setting name="postcard_template_id" displayName="Postcard Template ID" type="string" required="Yes">' +
        '<description>FES template id used to hydrate the offerpad_postcard artwork</description>' +
      '</setting>' +
    '</settings>'
  );
}

function getDefaultSettings() {
  return {
    fes_api_key: '',
    oneil_api_key: '',
    letter_template_id: '',
    postcard_template_id: ''
  };
}

/* -------- logging helpers -------- */
function logInfo(s, job, m){ try{ job.log(LogLevel.Info, m); }catch(_){ try{s.log(3,m);}catch(__){} } }
function logWarn(s, job, m){ try{ job.log(LogLevel.Warning, m); }catch(_){ try{s.log(2,m);}catch(__){} } }
function logErr(s, job, m){ try{ job.log(LogLevel.Error, m); }catch(_){ try{s.log(1,m);}catch(__){} } }

/* -------- settings helpers -------- */
async function tryGetProp(f, name){
  try{ if(f && typeof f.getPropertyValue==='function'){ var v=f.getPropertyValue(name); if(v && typeof v.then==='function') v=await v; if(v!=null) return String(v);} }catch(_){}
  try{ if(f && typeof f.getPropertyStringValue==='function'){ var v2=f.getPropertyStringValue(name); if(v2 && typeof v2.then==='function') v2=await v2; if(v2!=null) return String(v2);} }catch(_){}
  return '';
}
async function getVal(f, names){
  for (var i=0;i<names.length;i++){ var v = await tryGetProp(f, names[i]); if(v){ return v; } }
  return '';
}

/* -------- private data helpers -------- */
// Read a private data value, trying each candidate key until one is non-empty.
async function getPrivateDataValue(s, job, keys){
  for (var i=0;i<keys.length;i++){
    try {
      var v = job.getPrivateData(keys[i]);
      if (v && typeof v.then === 'function') v = await v;
      if (v != null && String(v) !== '') return String(v);
    } catch(_){}
  }
  return '';
}

// Build the record from job private data: { first_name: ..., ... }
async function readRecordFromPrivateData(s, job){
  var record = {};
  for (var i=0;i<FES_FIELD_MAP.length;i++){
    var spec = FES_FIELD_MAP[i];
    record[spec.keys[0]] = await getPrivateDataValue(s, job, spec.keys);
  }
  return record;
}

/* -------- HTTP helpers -------- */
function httpsPostJson(hostname, requestPath, headers, payload, timeoutMs, label, s, job){
  var https = require('https');
  return new Promise(function(resolve){
    var settled = false;
    function finish(result){
      if (settled) return;
      settled = true;
      resolve(result);
    }

    var allHeaders = { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payload) };
    for (var h in headers){ if (Object.prototype.hasOwnProperty.call(headers, h)) allHeaders[h] = headers[h]; }

    var options = { hostname: hostname, port: 443, path: requestPath, method: 'POST', headers: allHeaders };

    logInfo(s, job, 'POST ' + label + ' (' + payload.length + ' bytes)...');

    var req = https.request(options, function(res){
      var body = '';
      res.on('data', function(c){ body += c; });
      res.on('end', function(){
        logInfo(s, job, label + ' response: ' + res.statusCode + ', length: ' + body.length);
        req.setTimeout(0);
        finish({ statusCode: res.statusCode, body: body });
      });
      res.on('error', function(e){
        logErr(s, job, label + ' response error: ' + e.message);
        finish({ statusCode: 0, body: 'Error: ' + e.message });
      });
    });
    req.on('error', function(e){
      if (settled) return;
      logErr(s, job, label + ' request error: ' + e.message);
      finish({ statusCode: 0, body: 'Error: ' + e.message });
    });
    req.setTimeout(timeoutMs, function(){
      if (settled) return;
      logErr(s, job, label + ' timeout after ' + timeoutMs + ' ms');
      finish({ statusCode: 0, body: 'Error: timeout' });
      req.destroy(new Error('timeout'));
    });
    req.write(payload);
    req.end();
  });
}

// Create a FES hydrate job and return its hydrated artwork url.
async function createHydrateJob(templateId, record, fesApiKey, label, s, job){
  var data = [];
  for (var i=0;i<FES_FIELD_MAP.length;i++){
    var spec = FES_FIELD_MAP[i];
    data.push({ Field_name: spec.fieldName, Value: record[spec.keys[0]] || '' });
  }
  var payload = JSON.stringify({ Template: templateId, Data: data });

  var result = await httpsPostJson(
    'api.oneilprint.com',
    '/api/fes/hydrate/jobs/create',
    { 'x-api-key': fesApiKey },
    payload,
    60000,
    'FES hydrate (' + label + ')',
    s, job
  );

  if (result.statusCode < 200 || result.statusCode >= 300){
    throw new Error('FES hydrate (' + label + ') failed (' + result.statusCode + '): ' + String(result.body).substring(0, 300));
  }

  var resp;
  try { resp = JSON.parse(result.body); }
  catch (e) { throw new Error('FES hydrate (' + label + ') returned non-JSON: ' + String(result.body).substring(0, 200)); }

  var url = resp && resp.data ? resp.data.url : null;
  if (!url){
    throw new Error('FES hydrate (' + label + ') response missing data.url: ' + String(result.body).substring(0, 300));
  }
  logInfo(s, job, 'Hydrated ' + label + ' url: ' + url);
  return url;
}

function sanitizeIdPart(v){ return String(v || '').replace(/[^A-Za-z0-9]+/g, '').substring(0, 24); }

/* -------- send job forward -------- */
async function sendForward(s, job){
  if (!job || typeof job.sendToSingle !== 'function') {
    throw new Error('job.sendToSingle is unavailable');
  }

  var p = '';
  try {
    p = job.get(AccessLevel.ReadOnly);
    if (p && typeof p.then === 'function') p = await p;
  } catch (_) {}

  try {
    var direct = p ? job.sendToSingle(p) : job.sendToSingle();
    if (direct && typeof direct.then === 'function') await direct;
    return;
  } catch (directErr) {
    logWarn(
      s,
      job,
      'sendToSingle with current path failed; retrying default send: ' +
        (directErr && directErr.message ? directErr.message : directErr)
    );
  }

  try {
    var fallback = job.sendToSingle();
    if (fallback && typeof fallback.then === 'function') await fallback;
  } catch (e) {
    throw new Error('sendToSingle failed: ' + (e && e.message ? e.message : e));
  }
}

async function failJob(job, message){
  var failed = job.fail(message);
  if (failed && typeof failed.then === 'function') await failed;
}

/* -------- entry -------- */
async function jobArrived(s, f, job){
  logInfo(s, job, '=== Offerpad hydrate + OMS order script starting ===');

  try {
    var fesApiKey = (await getVal(f, ['fes_api_key'])).trim();
    var oneilApiKey = (await getVal(f, ['oneil_api_key'])).trim();
    var letterTemplateRaw = (await getVal(f, ['letter_template_id'])).trim();
    var postcardTemplateRaw = (await getVal(f, ['postcard_template_id'])).trim();

    if (!fesApiKey){ await failJob(job, 'Missing FES API Key setting'); return; }
    if (!oneilApiKey){ await failJob(job, 'Missing ONeil API Key setting'); return; }

    var letterTemplateId = parseInt(letterTemplateRaw, 10);
    var postcardTemplateId = parseInt(postcardTemplateRaw, 10);
    if (isNaN(letterTemplateId)){ await failJob(job, 'Letter Template ID must be an integer: ' + letterTemplateRaw); return; }
    if (isNaN(postcardTemplateId)){ await failJob(job, 'Postcard Template ID must be an integer: ' + postcardTemplateRaw); return; }

    // ---- Read recipient variables from private data ----
    var record = await readRecordFromPrivateData(s, job);

    var firstName = record[FES_FIELD_MAP[0].keys[0]];
    var lastName = record[FES_FIELD_MAP[1].keys[0]];
    var streetAddress = record[FES_FIELD_MAP[2].keys[0]];
    var city = record[FES_FIELD_MAP[3].keys[0]];
    var state = record[FES_FIELD_MAP[4].keys[0]];
    var zip = record[FES_FIELD_MAP[5].keys[0]];
    var webhookId = record[FES_FIELD_MAP[6].keys[0]];

    if (!streetAddress || !city || !state || !zip){
      await failJob(job, 'Missing required address private data (property_street_address/property_city/property_state/property_zip_code)');
      return;
    }

    // ---- Step 1: hydrate both templates ----
    var letterUrl = await createHydrateJob(letterTemplateId, record, fesApiKey, 'letter', s, job);
    var postcardUrl = await createHydrateJob(postcardTemplateId, record, fesApiKey, 'postcard', s, job);

    // ---- Step 2: one OMS order with two items ----
    // var externalOrderID = sanitizeIdPart(lastName) + '-' + sanitizeIdPart(zip) + '-' + Date.now();
    var externalOrderID = webhookId ? String(webhookId).substring(0, 50) : ('offerpad-' + Date.now());
    var attention = (firstName + ' ' + lastName).trim();

    var order = {
      externalOrderID: externalOrderID,
      shippingMethod: SHIPPING_METHOD,
      status: 'New',
      shippingAddresses: [
        {
          attention: attention || 'Resident',
          address1: streetAddress,
          address2: '',
          city: city,
          state: state,
          zip: zip
        }
      ],
      orderItems: [
        {
          productSKU: LETTER_PRODUCT_SKU,
          name: 'Offerpad Letter',
          quantity: ITEM_QUANTITY,
          artworkSingle: letterUrl,
          artworkLocal: ''
        },
        {
          productSKU: POSTCARD_PRODUCT_SKU,
          name: 'Offerpad Postcard',
          quantity: ITEM_QUANTITY,
          artworkSingle: postcardUrl,
          artworkLocal: ''
        }
      ]
    };

    var orderPayload = JSON.stringify(order);
    var omsResult = await httpsPostJson(
      'api.oneilprint.com',
      '/api/external/create-any',
      { 'x-api-key': oneilApiKey },
      orderPayload,
      60000,
      'OMS create-any',
      s, job
    );

    var omsOk = omsResult.statusCode >= 200 && omsResult.statusCode < 300;

    // ---- Persist results into private data ----
    try {
      await job.setPrivateData('OfferpadExternalOrderID', externalOrderID);
      await job.setPrivateData('OfferpadLetterUrl', letterUrl);
      await job.setPrivateData('OfferpadPostcardUrl', postcardUrl);
      await job.setPrivateData('OfferpadOmsResponse', omsResult.body);
    } catch(_){}

    if (!omsOk){
      await failJob(job, 'OMS create-any failed (' + omsResult.statusCode + '): ' + String(omsResult.body).substring(0, 400));
      return;
    }

    logInfo(s, job, 'OMS order created (' + externalOrderID + ')');
    logInfo(s, job, 'Sending job to output...');
    await sendForward(s, job);
    logInfo(s, job, '=== Offerpad script completed ===');
    return;
  } catch (e) {
    logErr(s, job, 'Unhandled error: ' + (e && e.message ? e.message : e));
    logErr(s, job, 'Stack: ' + (e && e.stack ? e.stack : 'no stack'));
    await failJob(job, 'Offerpad script error: ' + (e && e.message ? e.message : e));
  }
}
