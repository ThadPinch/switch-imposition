/// <reference types="switch-scripting" />
// @ts-nocheck
// OpenAI Responses API for Enfocus Switch 2023
// - Optionally fetches OMS Products/Kits catalogs
// - Optionally includes file in request
// - Stores GPT response in private data

/**
 * Tell Switch to allow longer script execution (5 minutes)
 * This is needed for API calls that may take time
 */
function getScriptTimeout() {
  return 300; // 5 minutes in seconds
}

function getSettingsDefinition() {
    return (
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<settings>' +
        '<setting name="openaiapikey" displayName="OpenAI API Key" type="string" required="Yes" password="Yes"/>' +
        '<setting name="model" displayName="Model" type="enum" required="Yes">' +
          '<choices>' +
            '<choice value="gpt-5.4"/>' +
          '</choices>' +
        '</setting>' +
        '<setting name="outputkey" displayName="Output Private Data Key" type="string" required="Yes">' +
          '<description>Private data key to store the OpenAI response</description>' +
        '</setting>' +
        '<setting name="openaisysmessage" displayName="System Message" type="multiline" required="No"/>' +
        '<setting name="openaiusermessage" displayName="User Message" type="multiline" required="Yes"/>' +
        '<setting name="appendomscatalog" displayName="Append OMS Catalog" type="enum" required="Yes">' +
          '<choices><choice value="No"/><choice value="Yes"/></choices>' +
          '<description>If Yes, fetches Products and Kits from OMS API</description>' +
          '<setting name="oneilapikey" displayName="ONeil API Key (Read)" type="string" required="No" password="Yes">' +
            '<description>API key for reading from api.oneilprint.com</description>' +
          '</setting>' +
        '</setting>' +
        '<setting name="openaiincludefile" displayName="Include File" type="enum" required="Yes">' +
          '<choices><choice value="No"/><choice value="Yes"/></choices>' +
          '<description>If Yes, includes the incoming file in the OpenAI request</description>' +
          '<setting name="filenamepattern" displayName="File Name Pattern" type="string" required="No">' +
            '<description>Only process files matching this pattern (e.g., "*.pdf", "*PO-*")</description>' +
          '</setting>' +
        '</setting>' +
        '<setting name="sendapicalltooms" displayName="Send APICall to OMS" type="enum" required="Yes">' +
          '<choices><choice value="No"/><choice value="Yes"/></choices>' +
          '<description>If Yes, POSTs the APICall to OMS create-any endpoint</description>' +
          '<setting name="oneilapikey2" displayName="ONeil API Key (Write)" type="string" required="No" password="Yes">' +
            '<description>API key for writing to api.oneilprint.com</description>' +
          '</setting>' +
        '</setting>' +
      '</settings>'
    );
  }
  
  function getDefaultSettings() {
    return {
      openaiapikey: '',
      model: 'gpt-5.4',
      outputkey: 'GPTResponse',
      openaisysmessage: '',
      openaiusermessage: '',
      appendomscatalog: 'No',
      oneilapikey: '',
      openaiincludefile: 'Yes',
      filenamepattern: '',
      sendapicalltooms: 'No',
      oneilapikey2: ''
    };
  }
  
  /* -------- helpers -------- */
  function asBool(v){ if(typeof v==='boolean')return v; if(v==null)return false; var s=String(v).trim().toLowerCase(); return s==='yes'||s==='true'||s==='1'; }
  function logInfo(s,job,m){ try{job.log(LogLevel.Info,m);}catch(_){try{s.log(3,m);}catch(_){}} }
  function logErr(s,job,m){ try{job.log(LogLevel.Error,m);}catch(_){try{s.log(1,m);}catch(_){}} }
  function logWarn(s,job,m){ try{job.log(LogLevel.Warning,m);}catch(_){try{s.log(2,m);}catch(_){}} }
  
  function globToRegex(pattern) {
    if (!pattern || !pattern.trim()) return null;
    var escaped = pattern
      .replace(/[.+^${}()|[\]\\]/g, '\\$&')
      .replace(/\*/g, '.*')
      .replace(/\?/g, '.');
    return new RegExp('^' + escaped + '$', 'i');
  }
  
  function matchesPattern(filename, pattern) {
    if (!pattern || !pattern.trim()) return true;
    var regex = globToRegex(pattern);
    if (!regex) return true;
    return regex.test(filename);
  }
  
  function getMimeType(filename) {
    var ext = (filename || '').split('.').pop().toLowerCase();
    var mimeTypes = {
      'pdf': 'application/pdf',
      'txt': 'text/plain',
      'csv': 'text/csv',
      'jpg': 'image/jpeg',
      'jpeg': 'image/jpeg',
      'png': 'image/png',
      'json': 'application/json',
      'xml': 'application/xml'
    };
    return mimeTypes[ext] || 'application/octet-stream';
  }

  function getPdfArrayNumber(arr, index, PDFNumberClass) {
    try {
      var n = arr.lookupMaybe(index, PDFNumberClass);
      if (n && typeof n.asNumber === 'function') return n.asNumber();
    } catch(_) {}
    try {
      var raw = arr.get(index);
      if (raw && typeof raw.asNumber === 'function') return raw.asNumber();
      var num = Number(raw);
      if (!isNaN(num)) return num;
    } catch(_) {}
    return NaN;
  }

  function decodePdfValue(valueObj) {
    if (!valueObj) return '';
    try { if (typeof valueObj.decodeText === 'function') return String(valueObj.decodeText() || ''); } catch(_) {}
    try { if (typeof valueObj.asString === 'function') return String(valueObj.asString() || ''); } catch(_) {}
    try { return String(valueObj || ''); } catch(_) {}
    return '';
  }

  async function extractPdfLinksInVisualOrder(filePath, s, job) {
    var fs = require('fs');
    var pdfLib;
    try {
      pdfLib = require('pdf-lib');
    } catch(e) {
      logWarn(s, job, 'pdf-lib is unavailable; skipping PDF link extraction: ' + e.message);
      return [];
    }

    var PDFDocument = pdfLib.PDFDocument;
    var PDFName = pdfLib.PDFName;
    var PDFArray = pdfLib.PDFArray;
    var PDFDict = pdfLib.PDFDict;
    var PDFString = pdfLib.PDFString;
    var PDFHexString = pdfLib.PDFHexString;
    var PDFNumber = pdfLib.PDFNumber;

    var bytes;
    try {
      bytes = fs.readFileSync(filePath);
    } catch(e) {
      logWarn(s, job, 'Cannot read PDF for link extraction: ' + e.message);
      return [];
    }

    var doc;
    try {
      doc = await PDFDocument.load(bytes);
    } catch(e) {
      logWarn(s, job, 'Cannot parse PDF for link extraction: ' + e.message);
      return [];
    }

    var pages = doc.getPages();
    var linkEntries = [];

    for (var p = 0; p < pages.length; p++) {
      var page = pages[p];
      var annots = null;
      try {
        annots = page && page.node && typeof page.node.lookupMaybe === 'function'
          ? page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
          : null;
      } catch(_) {}
      if (!annots || typeof annots.size !== 'function') continue;

      for (var i = 0; i < annots.size(); i++) {
        var annot = null;
        try {
          annot = annots.lookupMaybe(i, PDFDict);
          if (!annot) {
            var annotRef = annots.get(i);
            annot = doc.context.lookup(annotRef, PDFDict);
          }
        } catch(_) {}
        if (!annot) continue;

        var subtype = null;
        try { subtype = annot.lookupMaybe(PDFName.of('Subtype'), PDFName); } catch(_) {}
        if (!subtype || String(subtype) !== '/Link') continue;

        var action = null;
        try { action = annot.lookupMaybe(PDFName.of('A'), PDFDict); } catch(_) {}
        if (!action) continue;

        var uriObj = null;
        try { uriObj = action.lookupMaybe(PDFName.of('URI'), PDFString, PDFHexString); } catch(_) {}
        if (!uriObj) {
          try { uriObj = action.lookupMaybe(PDFName.of('URI')); } catch(_) {}
        }
        if (!uriObj) continue;

        var uri = decodePdfValue(uriObj).trim();
        if (!uri) continue;

        // If URI came from PDF string formatting, normalize wrappers.
        if (uri.charAt(0) === '(' && uri.charAt(uri.length - 1) === ')') {
          uri = uri.substring(1, uri.length - 1).trim();
        }
        if (!uri) continue;

        var label = '';
        try {
          var labelObj = annot.lookupMaybe(PDFName.of('Contents'), PDFString, PDFHexString);
          if (labelObj) label = decodePdfValue(labelObj).trim();
        } catch(_) {}
        if (!label) {
          try {
            var altLabelObj = annot.lookupMaybe(PDFName.of('TU'), PDFString, PDFHexString);
            if (altLabelObj) label = decodePdfValue(altLabelObj).trim();
          } catch(_) {}
        }
        if (label.charAt(0) === '(' && label.charAt(label.length - 1) === ')') {
          label = label.substring(1, label.length - 1).trim();
        }

        var rect = null;
        try { rect = annot.lookupMaybe(PDFName.of('Rect'), PDFArray); } catch(_) {}

        var left = null;
        var top = null;
        if (rect && typeof rect.size === 'function' && rect.size() >= 4) {
          var x1 = getPdfArrayNumber(rect, 0, PDFNumber);
          var y1 = getPdfArrayNumber(rect, 1, PDFNumber);
          var x2 = getPdfArrayNumber(rect, 2, PDFNumber);
          var y2 = getPdfArrayNumber(rect, 3, PDFNumber);
          if (!isNaN(x1) && !isNaN(x2)) left = Math.min(x1, x2);
          if (!isNaN(y1) && !isNaN(y2)) top = Math.max(y1, y2);
        }

        linkEntries.push({
          page: p + 1,
          annotIndex: i,
          left: left,
          top: top,
          label: label,
          url: uri
        });
      }
    }

    linkEntries.sort(function(a, b) {
      if (a.page !== b.page) return a.page - b.page;
      if (a.top != null && b.top != null) {
        var topDiff = b.top - a.top;
        if (Math.abs(topDiff) > 0.1) return topDiff;
      } else if (a.top != null || b.top != null) {
        return a.top != null ? -1 : 1;
      }
      if (a.left != null && b.left != null) {
        var leftDiff = a.left - b.left;
        if (Math.abs(leftDiff) > 0.1) return leftDiff;
      } else if (a.left != null || b.left != null) {
        return a.left != null ? -1 : 1;
      }
      return a.annotIndex - b.annotIndex;
    });

    return linkEntries;
  }
  
  async function tryGetProp(f,name){
    try{ if(f&&typeof f.getPropertyValue==='function'){ var v=f.getPropertyValue(name); if(v&&typeof v.then==='function') v=await v; if(v!=null) return String(v);} }catch(_){}
    try{ if(f&&typeof f.getPropertyStringValue==='function'){ var v2=f.getPropertyStringValue(name); if(v2&&typeof v2.then==='function') v2=await v2; if(v2!=null) return String(v2);} }catch(_){}
    return '';
  }
  async function getVal(f, names){ for (var i=0;i<names.length;i++){ var v = await tryGetProp(f, names[i]); if(v){ return v; } } return ''; }
  
  async function getIncomingPath(job){
    try{ if(typeof AccessLevel!=='undefined' && job && typeof job.get==='function'){ var p=job.get(AccessLevel.ReadOnly); if(p&&typeof p.then==='function') p=await p; return p; } }catch(_){}
    if(job && typeof job.getPath==='function'){ var p2=job.getPath(); if(p2&&typeof p2.then==='function') p2=await p2; return p2; }
    throw new Error('Cannot resolve job path');
  }

  async function sendJobForward(s, job, fallbackPath){
    if (!job || typeof job.sendToSingle !== 'function') {
      throw new Error('job.sendToSingle is unavailable');
    }

    // Prefer no-arg send; this is the most stable path in Switch script runtimes.
    try {
      var direct = job.sendToSingle();
      if (direct && typeof direct.then === 'function') await direct;
      return;
    } catch (directErr) {
      logWarn(
        s,
        job,
        'sendToSingle() without path failed; retrying with explicit path: ' +
          ((directErr && directErr.message) ? directErr.message : String(directErr || 'unknown error'))
      );
    }

    var sendPath = '';
    try { sendPath = await getIncomingPath(job); } catch(_) {}
    if (!sendPath && fallbackPath) sendPath = fallbackPath;
    if (!sendPath) throw new Error('Cannot resolve output path for sendToSingle');

    var explicit = job.sendToSingle(sendPath);
    if (explicit && typeof explicit.then === 'function') await explicit;
  }
  
  /**
   * Make HTTPS GET request to OMS API
   */
  function fetchOMSData(endpoint, apiKey, s, job) {
    var https = require('https');
    
    return new Promise(function(resolve) {
      var options = {
        hostname: 'api.oneilprint.com',
        port: 443,
        path: '/api/external/' + endpoint,
        method: 'GET',
        headers: {
          'x-api-key': apiKey
        }
      };
      
      logInfo(s, job, 'Fetching OMS ' + endpoint + ' (HTTPS)...');
      
      var req = https.request(options, function(res) {
        var body = '';
        res.on('data', function(c) { body += c; });
        res.on('end', function() {
          logInfo(s, job, 'OMS ' + endpoint + ' response: ' + res.statusCode + ', length: ' + body.length);
          if (res.statusCode === 200) {
            resolve(body);
          } else {
            logWarn(s, job, 'OMS ' + endpoint + ' non-200: ' + body.substring(0, 200));
            resolve(null);
          }
        });
        res.on('error', function(e) {
          logErr(s, job, 'OMS ' + endpoint + ' response error: ' + e.message);
          resolve(null);
        });
      });
      
      req.on('error', function(e) {
        logErr(s, job, 'OMS ' + endpoint + ' request error: ' + e.message);
        resolve(null);
      });
      
      req.setTimeout(30000, function() {
        logErr(s, job, 'OMS ' + endpoint + ' timeout');
        req.destroy();
        resolve(null);
      });
      
      req.end();
    });
  }
  
  /**
   * POST APICall to OMS create-any endpoint
   */
  function postToOMS(apiCallJson, apiKey, s, job) {
    var https = require('https');
    
    return new Promise(function(resolve) {
      var payload = apiCallJson;
      
      var options = {
        hostname: 'api.oneilprint.com',
        port: 443,
        path: '/api/external/create-any',
        method: 'POST',
        headers: {
          'x-api-key': apiKey,
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(payload)
        }
      };
      
      logInfo(s, job, 'POSTing to OMS create-any (' + payload.length + ' bytes)...');
      
      var req = https.request(options, function(res) {
        var body = '';
        res.on('data', function(c) { body += c; });
        res.on('end', function() {
          logInfo(s, job, 'OMS create-any response: ' + res.statusCode + ', length: ' + body.length);
          resolve({ statusCode: res.statusCode, body: body });
        });
        res.on('error', function(e) {
          logErr(s, job, 'OMS create-any response error: ' + e.message);
          resolve({ statusCode: 0, body: 'Error: ' + e.message });
        });
      });
      
      req.on('error', function(e) {
        logErr(s, job, 'OMS create-any request error: ' + e.message);
        resolve({ statusCode: 0, body: 'Error: ' + e.message });
      });
      
      req.setTimeout(60000, function() {
        logErr(s, job, 'OMS create-any timeout');
        req.destroy();
        resolve({ statusCode: 0, body: 'Error: timeout' });
      });
      
      req.write(payload);
      req.end();
    });
  }

  /**
   * Make HTTPS POST request to OpenAI
   */
  function callOpenAI(payload, apiKey, s, job) {
    var https = require('https');
    
    return new Promise(function(resolve, reject) {
      var settled = false;
      function resolveOnce(value) {
        if (settled) return;
        settled = true;
        resolve(value);
      }
      function rejectOnce(err) {
        if (settled) return;
        settled = true;
        reject(err);
      }

      var options = {
        hostname: 'api.openai.com',
        port: 443,
        path: '/v1/responses',
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + apiKey
        }
      };
      
      logInfo(s, job, 'Calling OpenAI (' + payload.length + ' bytes)...');
      
      var req = https.request(options, function(res) {
        var body = '';
        res.on('data', function(c) { body += c; });
        res.on('end', function() {
          logInfo(s, job, 'OpenAI response: ' + res.statusCode + ', length: ' + body.length);
          resolveOnce({ statusCode: res.statusCode, body: body });
        });
        res.on('error', function(e) {
          logErr(s, job, 'OpenAI response error: ' + e.message);
          rejectOnce(e);
        });
      });
      
      req.on('error', function(e) {
        logErr(s, job, 'OpenAI request error: ' + e.message);
        rejectOnce(e);
      });
      
      // Keep request timeout below script timeout (5 minutes).
      var requestTimeoutMs = 270000;
      req.setTimeout(requestTimeoutMs, function() {
        logErr(s, job, 'OpenAI timeout after ' + requestTimeoutMs + ' ms');
        req.destroy(new Error('timeout'));
      });
      
      req.write(payload);
      req.end();
    });
  }
  
  /* -------- entry -------- */
  async function jobArrived(s, f, job){
    var fs = require('fs');
    var path = require('path');
  
    logInfo(s, job, '=== OpenAI Script Starting ===');
    
    try {
      // Get settings
      var openaiApiKey     = await getVal(f, ['openaiapikey']);
      var model            = (await getVal(f, ['model'])) || 'gpt-5.4';
      var outputKey        = (await getVal(f, ['outputkey'])) || 'GPTResponse';
      var systemMsg        = await getVal(f, ['openaisysmessage']);
      var userMsg          = await getVal(f, ['openaiusermessage']);
      var appendOmsCatalog = asBool(await getVal(f, ['appendomscatalog']));
      var oneilApiKey      = await getVal(f, ['oneilapikey']);
      var includeFile      = asBool(await getVal(f, ['openaiincludefile']));
      var filePattern      = await getVal(f, ['filenamepattern']);
      var sendApiCallToOms = asBool(await getVal(f, ['sendapicalltooms']));
      var oneilApiKey2     = await getVal(f, ['oneilapikey2']);
    
      if(!openaiApiKey){ logErr(s,job,'Missing OpenAI API key'); await job.fail('Missing OpenAI API key'); return; }
      if(!userMsg){ logErr(s,job,'Missing user message'); await job.fail('Missing message'); return; }
    
      var inPath = await getIncomingPath(job);
      var fname = path.basename(inPath) || 'document';
      
      logInfo(s, job, 'Input path: ' + inPath);
      logInfo(s, job, 'Include file: ' + includeFile + ', Append OMS: ' + appendOmsCatalog + ', Send to OMS: ' + sendApiCallToOms);
      
      // Check if this is a directory - if so, find the actual file
      var stats;
      try {
        stats = fs.statSync(inPath);
      } catch(e) {
        logErr(s, job, 'Cannot stat path: ' + inPath);
        await job.fail('Cannot access file');
        return;
      }
      
      var actualFilePath = inPath;
      if (stats.isDirectory()) {
        logInfo(s, job, 'Input is directory, searching for PDF...');
        var files = fs.readdirSync(inPath);
        var foundFile = null;
        
        for (var i = 0; i < files.length; i++) {
          var filePath = path.join(inPath, files[i]);
          try {
            var fileStat = fs.statSync(filePath);
            if (fileStat.isFile()) {
              var ext = path.extname(files[i]).toLowerCase();
              if (ext === '.pdf') {
                foundFile = filePath;
                break;
              } else if (!foundFile) {
                foundFile = filePath;
              }
            }
          } catch(e) {}
        }
        
        if (foundFile) {
          actualFilePath = foundFile;
          fname = path.basename(foundFile);
          logInfo(s, job, 'Found file: ' + fname);
        } else {
          logErr(s, job, 'No file found in directory');
          await job.fail('No file found');
          return;
        }
      }
      
      // Check file pattern (only if including file)
      if (includeFile && filePattern && filePattern.trim()) {
        if (!matchesPattern(fname, filePattern)) {
          logInfo(s, job, 'File "' + fname + '" does not match pattern "' + filePattern + '", passing through.');
          await sendJobForward(s, job, inPath);
          return;
        }
        logInfo(s, job, 'File "' + fname + '" matches pattern.');
      }
      
      // Fetch OMS catalogs if enabled
      var catalogSection = '';
      if (appendOmsCatalog) {
        if (!oneilApiKey) {
          logWarn(s, job, 'Append OMS Catalog is Yes but no ONeil API key provided');
        } else {
          logInfo(s, job, 'Fetching OMS catalogs...');
          
          var productsData = await fetchOMSData('products', oneilApiKey, s, job);
          var kitsData = await fetchOMSData('kits', oneilApiKey, s, job);
          var shippingOptionsData = await fetchOMSData('shipping-options', oneilApiKey, s, job);
          
          if (productsData || kitsData || shippingOptionsData) {
            var catalogParts = [];
            if (productsData) {
              catalogParts.push('Products=' + encodeURIComponent(productsData));
              logInfo(s, job, 'Products catalog: ' + productsData.length + ' chars');
            }
            if (kitsData) {
              catalogParts.push('Kits=' + encodeURIComponent(kitsData));
              logInfo(s, job, 'Kits catalog: ' + kitsData.length + ' chars');
            }
            if (shippingOptionsData) {
              catalogParts.push('ShippingOptions=' + encodeURIComponent(shippingOptionsData));
              logInfo(s, job, 'Shipping options: ' + shippingOptionsData.length + ' chars');
            }
            catalogSection = '\n\n<catalog_data encoding="url">\n' + catalogParts.join('\n') + '\n</catalog_data>\n\nNOTE: Catalog data is URL-encoded. Decode to get JSON. Use ONLY shipping methods from ShippingOptions - the shippingMethod field must match one of the available options exactly.';
          }
        }
      }
      
      // Extract PDF links and append to the user message so artwork links are explicit.
      var pdfLinksSection = '';
      var extForLinks = path.extname(fname).toLowerCase();
      if (extForLinks === '.pdf') {
        var orderedPdfLinks = await extractPdfLinksInVisualOrder(actualFilePath, s, job);
        if (orderedPdfLinks.length > 0) {
          var fileLinkOnly = [];
          for (var fl = 0; fl < orderedPdfLinks.length; fl++) {
            var labelText = String(orderedPdfLinks[fl].label || '');
            if (/file\s*link/i.test(labelText)) fileLinkOnly.push(orderedPdfLinks[fl]);
          }

          var linkLines = [];
          for (var li = 0; li < orderedPdfLinks.length; li++) {
            var entry = orderedPdfLinks[li];
            linkLines.push((li + 1) + '. ' + entry.url);
          }

          var fileLinkLines = [];
          for (var fi = 0; fi < fileLinkOnly.length; fi++) {
            fileLinkLines.push((fi + 1) + '. ' + fileLinkOnly[fi].url);
          }

          pdfLinksSection =
            '\n\n<pdf_links_in_order>' +
            '\nThese are PDF hyperlink URLs in visual order (page, top-to-bottom, left-to-right).' +
            '\nFor purchase orders, the row-level artwork links are usually the "File Link" entries.' +
            '\n' + linkLines.join('\n') +
            '\n</pdf_links_in_order>' +
            (fileLinkLines.length > 0
              ? '\n\n<pdf_file_links_in_order>' +
                '\nThese are links whose annotation label explicitly contains "File Link".' +
                '\n' + fileLinkLines.join('\n') +
                '\n</pdf_file_links_in_order>'
              : '\n\n<pdf_file_links_in_order>\nNo explicit "File Link" labels were found in annotation metadata.\n</pdf_file_links_in_order>');
          logInfo(
            s,
            job,
            'Extracted ' + orderedPdfLinks.length + ' PDF hyperlink(s) from ' + fname +
            (fileLinkOnly.length > 0 ? ('; File Link-labeled URLs: ' + fileLinkOnly.length) : '')
          );
        } else {
          logWarn(s, job, 'No PDF hyperlink annotations found in ' + fname);
        }
      }

      // Build full user message
      var fullUserMsg = userMsg + catalogSection + pdfLinksSection;
      
      // Build user content
      var userContent = [];
      
      // Add file if enabled
      if (includeFile) {
        try {
          var buf = fs.readFileSync(actualFilePath);
          var b64 = buf.toString('base64');
          var mimeType = getMimeType(fname);
          logInfo(s, job, 'Including file: ' + fname + ' (' + buf.length + ' bytes)');
          userContent.push({
            type: 'input_file',
            filename: fname,
            file_data: 'data:' + mimeType + ';base64,' + b64
          });
        } catch(e) {
          logErr(s, job, 'Failed to read file: ' + e.message);
          await job.fail('File read error');
          return;
        }
      }
      
      // Add user message
      userContent.push({
        type: 'input_text',
        text: fullUserMsg
      });
      
      // Build OpenAI request
      var input = [];
      if (String(systemMsg).trim()) {
        input.push({ role: 'system', content: [{ type: 'input_text', text: systemMsg }] });
      }
      input.push({ role: 'user', content: userContent });
      
      var payload = JSON.stringify({ model: model, input: input });
      
      // Call OpenAI
      var response = await callOpenAI(payload, openaiApiKey, s, job);
      
      var parsed;
      try {
        parsed = JSON.parse(response.body);
      } catch(e) {
        logErr(s, job, 'Failed to parse OpenAI response: ' + response.body.substring(0, 500));
        await job.fail('Invalid response from OpenAI');
        return;
      }
      
      if (parsed.error) {
        logErr(s, job, 'OpenAI error: ' + JSON.stringify(parsed.error));
        await job.fail('OpenAI: ' + (parsed.error.message || 'Error'));
        return;
      }
      
      // Extract response text
      var responseText = response.body;
      if (parsed.output && Array.isArray(parsed.output)) {
        for (var i = 0; i < parsed.output.length; i++) {
          var item = parsed.output[i];
          if (item.type === 'message' && item.content && Array.isArray(item.content)) {
            for (var j = 0; j < item.content.length; j++) {
              if (item.content[j].type === 'output_text') {
                responseText = item.content[j].text;
                break;
              }
            }
          }
        }
      }
      
      logInfo(s, job, 'Response text length: ' + responseText.length);
      logInfo(s, job, 'Response preview: ' + responseText.substring(0, 200));
      
      // Parse the response JSON and extract APICall, alerts, Confirmation
      var apiCallStr = '';
      var alertSubject = '';
      var alertMessage = '';
      var prepressAlertSubject = '';
      var prepressAlertMessage = '';
      var csrAlertSubject = '';
      var csrAlertMessage = '';
      var confirmationMessage = '';
      var omsResponse = '';
      var omsSuccess = false;
      
      try {
        // Clean up response - remove markdown code blocks if present
        var cleanedResponse = responseText.trim();
        if (cleanedResponse.startsWith('```json')) {
          cleanedResponse = cleanedResponse.replace(/^```json\s*/, '').replace(/\s*```$/, '');
        } else if (cleanedResponse.startsWith('```')) {
          cleanedResponse = cleanedResponse.replace(/^```\s*/, '').replace(/\s*```$/, '');
        }
        
        var responseObj = JSON.parse(cleanedResponse);
        
        // Extract APICall - stringify it
        if (responseObj.APICall) {
          apiCallStr = JSON.stringify(responseObj.APICall);
          logInfo(s, job, 'Extracted APICall: ' + apiCallStr.substring(0, 100) + '...');
        }
        
        // Extract Alert fields if present
        if (responseObj.Alert && responseObj.Alert !== null) {
          alertSubject = responseObj.Alert.Subject || '';
          alertMessage = responseObj.Alert.Message || '';
          logInfo(s, job, 'Extracted Alert - Subject: ' + alertSubject);
        } else {
          logInfo(s, job, 'No Alert in response (null)');
        }

        if (responseObj.PrepressAlert && responseObj.PrepressAlert !== null) {
          prepressAlertSubject = responseObj.PrepressAlert.Subject || '';
          prepressAlertMessage = responseObj.PrepressAlert.Message || '';
          logInfo(s, job, 'Extracted PrepressAlert - Subject: ' + prepressAlertSubject);
        } else {
          logInfo(s, job, 'No PrepressAlert in response (null)');
        }

        if (responseObj.CSRAlert && responseObj.CSRAlert !== null) {
          csrAlertSubject = responseObj.CSRAlert.Subject || '';
          csrAlertMessage = responseObj.CSRAlert.Message || '';
          logInfo(s, job, 'Extracted CSRAlert - Subject: ' + csrAlertSubject);
        } else {
          logInfo(s, job, 'No CSRAlert in response (null)');
        }
        
        if (responseObj.Confirmation != null) {
          confirmationMessage =
            typeof responseObj.Confirmation === 'string'
              ? responseObj.Confirmation
              : String(responseObj.Confirmation);
          logInfo(s, job, 'Extracted Confirmation: ' + confirmationMessage.substring(0, 120));
        } else {
          logInfo(s, job, 'No Confirmation in response');
        }
      } catch(parseErr) {
        logErr(s, job, 'Failed to parse response JSON: ' + parseErr.message);
        logErr(s, job, 'Raw response: ' + responseText.substring(0, 500));
        // Store raw response as fallback
        apiCallStr = responseText;
      }
      
      // POST APICall to OMS if enabled
      if (sendApiCallToOms && apiCallStr) {
        if (!oneilApiKey2) {
          logWarn(s, job, 'Send APICall to OMS is Yes but no ONeil API Key (Write) provided');
        } else {
          logInfo(s, job, 'Sending APICall to OMS...');
          var omsResult = await postToOMS(apiCallStr, oneilApiKey2, s, job);
          omsResponse = omsResult.body;
          omsSuccess = (omsResult.statusCode >= 200 && omsResult.statusCode < 300);
          
          if (omsSuccess) {
            logInfo(s, job, 'OMS create-any success: ' + omsResponse.substring(0, 200));
          } else {
            logErr(s, job, 'OMS create-any failed (' + omsResult.statusCode + '): ' + omsResponse.substring(0, 500));
          }
        }
      }
      
      // Store in private data
      logInfo(s, job, 'Storing response in private data...');
      await job.setPrivateData(outputKey, apiCallStr);
      await job.setPrivateData('AlertSubject', alertSubject);
      await job.setPrivateData('AlertMessage', alertMessage);
      await job.setPrivateData('PrepressAlertSubject', prepressAlertSubject);
      await job.setPrivateData('PrepressAlertMessage', prepressAlertMessage);
      await job.setPrivateData('CSRAlertSubject', csrAlertSubject);
      await job.setPrivateData('CSRAlertMessage', csrAlertMessage);
      await job.setPrivateData('Confirmation', confirmationMessage);
      await job.setPrivateData('OMSResponse', omsResponse);
      await job.setPrivateData('OMSSuccess', omsSuccess ? 'true' : 'false');
      logInfo(
        s,
        job,
        'Stored: ' +
          outputKey +
          ', AlertSubject, AlertMessage, PrepressAlertSubject, PrepressAlertMessage, CSRAlertSubject, CSRAlertMessage, Confirmation, OMSResponse, OMSSuccess'
      );
      
      // Send job forward - get fresh path from job
      logInfo(s, job, 'Sending job to output...');
      logInfo(s, job, '=== Job completed successfully; forwarding ===');
      await sendJobForward(s, job, inPath);
      return;
      
    } catch(e) {
      logErr(s, job, 'Unhandled error: ' + e.message);
      logErr(s, job, 'Stack: ' + (e.stack || 'no stack'));
      await job.fail('Error: ' + e.message);
    }
  }
