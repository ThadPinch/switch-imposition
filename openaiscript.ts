/// <reference types="switch-scripting" />
// @ts-nocheck
// OpenAI Responses API for Enfocus Switch 2023
// - Sends matching files + private data as file attachments to OpenAI
// - Passes original files through unchanged
// - Stores GPT response in a named dataset

function getSettingsDefinition() {
    return (
      '<?xml version="1.0" encoding="UTF-8"?>' +
      '<settings>' +
        '<setting name="model" displayName="Model" type="enum" required="Yes">' +
          '<choices>' +
            '<choice value="gpt-5-mini"/>' +
            '<choice value="gpt-5"/>' +
          '</choices>' +
        '</setting>' +
        '<setting name="openaiapikey" displayName="API Key" type="string" required="Yes" password="Yes"/>' +
        '<setting name="openaisysmessage" displayName="System Message" type="multiline" required="No"/>' +
        '<setting name="openaiusermessage" displayName="User Message" type="multiline" required="Yes"/>' +
        '<setting name="filenamepattern" displayName="File Name Pattern" type="string" required="No">' +
          '<description>Only process files matching this pattern (e.g., "*.pdf", "*PO-*"). Leave blank to process all files.</description>' +
        '</setting>' +
        '<setting name="includeprivatedata" displayName="Include Private Data Keys" type="string" required="No">' +
          '<description>Comma-separated list of private data keys to include as JSON files (e.g., "Products,Kits")</description>' +
        '</setting>' +
        '<setting name="datasetname" displayName="Output Dataset Name" type="string" required="Yes">' +
          '<description>Name of the dataset to store the OpenAI response</description>' +
        '</setting>' +
      '</settings>'
    );
  }
  
  function getDefaultSettings() {
    return {
      model: 'gpt-5-mini',
      openaiapikey: '',
      openaisysmessage: '',
      openaiusermessage: '',
      filenamepattern: '',
      includeprivatedata: 'Products,Kits',
      datasetname: 'GPTResponse'
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
  
  async function tryGetProp(f,name){
    try{ if(f&&typeof f.getPropertyValue==='function'){ var v=f.getPropertyValue(name); if(v&&typeof v.then==='function') v=await v; if(v!=null) return String(v);} }catch(_){}
    try{ if(f&&typeof f.getPropertyStringValue==='function'){ var v2=f.getPropertyStringValue(name); if(v2&&typeof v2.then==='function') v2=await v2; if(v2!=null) return String(v2);} }catch(_){}
    return '';
  }
  async function getVal(f, names, redact){ for (var i=0;i<names.length;i++){ var v = await tryGetProp(f, names[i]); if(v){ if(!redact){} return v; } } return ''; }
  
  async function getIncomingPath(job){
    try{ if(typeof AccessLevel!=='undefined' && job && typeof job.get==='function'){ var p=job.get(AccessLevel.ReadOnly); if(p&&typeof p.then==='function') p=await p; return p; } }catch(_){}
    if(job && typeof job.getPath==='function'){ var p2=job.getPath(); if(p2&&typeof p2.then==='function') p2=await p2; return p2; }
    throw new Error('Cannot resolve job path');
  }
  
  /**
   * Get private data value and clean it up
   */
  async function getPrivateDataValue(job, key, s) {
    try {
      var value = await job.getPrivateData(key);
      if (value === null || value === undefined) {
        logWarn(s, job, 'Private data "' + key + '" is null/undefined');
        return null;
      }
      
      // Convert to string if needed
      var strValue = String(value);
      
      // Clean up: remove BOM, trim whitespace
      strValue = strValue.replace(/^\uFEFF/, '').trim();
      
      // If it's double-encoded (string wrapped in quotes), unwrap it
      if (strValue.startsWith('"') && strValue.endsWith('"')) {
        try {
          var unwrapped = JSON.parse(strValue);
          if (typeof unwrapped === 'string') {
            strValue = unwrapped;
          }
        } catch(e) {
          // Not double-encoded, use as-is
        }
      }
      
      logInfo(s, job, 'Got private data "' + key + '" (' + strValue.length + ' chars)');
      return strValue;
    } catch(e) {
      logErr(s, job, 'Failed to get private data "' + key + '": ' + e.message);
      return null;
    }
  }
  
  /**
   * Write output dataset
   */
  async function writeOutputDataset(job, datasetName, content, s) {
    var fs = require('fs');
    var path = require('path');
    
    // Try createDataset
    try {
      if (typeof job.createDataset === 'function') {
        var dataset = await job.createDataset(datasetName);
        if (dataset) {
          var datasetPath = typeof dataset.getPath === 'function' ? await dataset.getPath() : String(dataset);
          if (datasetPath) {
            fs.writeFileSync(datasetPath, content, 'utf8');
            logInfo(s, job, 'Wrote dataset: ' + datasetPath);
            return true;
          }
        }
      }
    } catch(e) {
      logWarn(s, job, 'createDataset failed: ' + e);
    }
    
    // Try createPathWithName
    try {
      if (typeof job.createPathWithName === 'function') {
        var outPath = await job.createPathWithName(datasetName + '.json', false);
        if (outPath) {
          fs.writeFileSync(outPath, content, 'utf8');
          logInfo(s, job, 'Wrote file: ' + outPath);
          return true;
        }
      }
    } catch(e) {
      logWarn(s, job, 'createPathWithName failed: ' + e);
    }
    
    // Fallback to temp
    try {
      var os = require('os');
      var tempPath = path.join(os.tmpdir(), 'switch_' + datasetName + '_' + Date.now() + '.json');
      fs.writeFileSync(tempPath, content, 'utf8');
      logInfo(s, job, 'Wrote to temp: ' + tempPath);
      return true;
    } catch(e) {
      logErr(s, job, 'All write methods failed');
      return false;
    }
  }
  
  /* -------- entry -------- */
  async function jobArrived(s, f, job){
    var fs = require('fs');
    var https = require('https');
    var path = require('path');
  
    var model            = (await getVal(f, ['model'], false)) || 'gpt-5-mini';
    var apiKey           = await getVal(f, ['openaiapikey','api key','apikey'], true);
    var systemMsg        = await getVal(f, ['openaisysmessage','system message'], false);
    var userMsg          = await getVal(f, ['openaiusermessage','message','prompt'], false);
    var filePattern      = await getVal(f, ['filenamepattern','file name pattern'], false);
    var includePrivateData = await getVal(f, ['includeprivatedata','include private data'], false);
    var outputDataset    = (await getVal(f, ['datasetname','dataset name'], false)) || 'GPTResponse';
  
    if(!apiKey){ logErr(s,job,'Missing API key'); await job.fail('Missing API key'); return; }
    if(!userMsg){ logErr(s,job,'Missing user message'); await job.fail('Missing message'); return; }
  
    var inPath = await getIncomingPath(job);
    var fname = path.basename(inPath) || 'document';
    
    // Check if this is a directory (job folder) - if so, we need to find the actual file
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
      // Find PDF or matching file in directory
      logInfo(s, job, 'Input is directory, searching for files...');
      var files = fs.readdirSync(inPath);
      var foundFile = null;
      
      for (var i = 0; i < files.length; i++) {
        var filePath = path.join(inPath, files[i]);
        var fileStat = fs.statSync(filePath);
        if (fileStat.isFile()) {
          var ext = path.extname(files[i]).toLowerCase();
          // Prefer PDF, but take first file if no pattern
          if (ext === '.pdf') {
            foundFile = filePath;
            break;
          } else if (!foundFile && matchesPattern(files[i], filePattern)) {
            foundFile = filePath;
          }
        }
      }
      
      if (foundFile) {
        actualFilePath = foundFile;
        fname = path.basename(foundFile);
        logInfo(s, job, 'Found file in directory: ' + fname);
      } else {
        logErr(s, job, 'No matching file found in directory');
        await job.fail('No file found in job folder');
        return;
      }
    }
    
    // Check file pattern
    if (filePattern && filePattern.trim()) {
      if (!matchesPattern(fname, filePattern)) {
        logInfo(s, job, 'File "' + fname + '" does not match pattern, passing through.');
        await job.sendToSingle(inPath);
        return;
      }
      logInfo(s, job, 'File "' + fname + '" matches pattern, processing.');
    }
    
    // Build content array for user message
    var userContent = [];
    
    // 1. Add the incoming file (PDF)
    try {
      var buf = fs.readFileSync(actualFilePath);
      var b64 = buf.toString('base64');
      var mimeType = getMimeType(fname);
      logInfo(s, job, 'Adding file: ' + fname + ' (' + mimeType + ', ' + buf.length + ' bytes)');
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
    
    // 2. Add private data as JSON files
    if (includePrivateData && includePrivateData.trim()) {
      var privateDataKeys = includePrivateData.split(',').map(function(k) { return k.trim(); }).filter(function(k) { return k; });
      
      logInfo(s, job, 'Including private data keys: ' + privateDataKeys.join(', '));
      
      for (var i = 0; i < privateDataKeys.length; i++) {
        var pdKey = privateDataKeys[i];
        var pdValue = await getPrivateDataValue(job, pdKey, s);
        
        if (pdValue) {
          var pdFilename = pdKey + '.json';
          var pdB64 = Buffer.from(pdValue, 'utf8').toString('base64');
          logInfo(s, job, 'Adding private data as file: ' + pdFilename + ' (' + pdValue.length + ' chars)');
          userContent.push({
            type: 'input_file',
            filename: pdFilename,
            file_data: 'data:application/json;base64,' + pdB64
          });
        } else {
          logWarn(s, job, 'Skipping private data "' + pdKey + '" - no value');
        }
      }
    }
    
    // 3. Add the user message text
    userContent.push({
      type: 'input_text',
      text: userMsg
    });
    
    // Build input array
    var input = [];
    if (String(systemMsg).trim()) {
      input.push({ role: 'system', content: [{ type: 'input_text', text: systemMsg }] });
    }
    input.push({ role: 'user', content: userContent });
  
    var payload = JSON.stringify({ model, input });
    logInfo(s, job, 'Sending to OpenAI (' + payload.length + ' bytes, ' + userContent.length + ' content items)...');
  
    var options = {
      hostname: 'api.openai.com',
      port: 443,
      path: '/v1/responses',
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey }
    };
  
    return new Promise(function(resolve, reject) {
      var req = https.request(options, function(res) {
        var body = '';
        res.on('data', function(c) { body += c; });
        res.on('end', async function() {
          try {
            logInfo(s, job, 'OpenAI status: ' + res.statusCode);
            
            var response;
            try {
              response = JSON.parse(body);
            } catch(parseErr) {
              logErr(s, job, 'Parse error: ' + body.substring(0, 500));
              await job.fail('Invalid JSON from OpenAI');
              reject(parseErr);
              return;
            }
            
            if (response.error) {
              logErr(s, job, 'OpenAI error: ' + JSON.stringify(response.error));
              await job.fail('OpenAI: ' + (response.error.message || 'Error'));
              reject(new Error(response.error.message));
              return;
            }
            
            // Extract response content
            var content = body;
            if (response.output && Array.isArray(response.output)) {
              for (var i = 0; i < response.output.length; i++) {
                var item = response.output[i];
                if (item.type === 'message' && item.content && Array.isArray(item.content)) {
                  for (var j = 0; j < item.content.length; j++) {
                    if (item.content[j].type === 'output_text') {
                      content = item.content[j].text;
                      break;
                    }
                  }
                }
              }
            }
            
            logInfo(s, job, 'Response length: ' + content.length);
            
            // Write output
            var writeSuccess = await writeOutputDataset(job, outputDataset, content, s);
            if (!writeSuccess) {
              logErr(s, job, 'Failed to write output');
              await job.fail('Write failed');
              reject(new Error('Write failed'));
              return;
            }
            
            // Send original job forward (use original inPath, not actualFilePath)
            await job.sendToSingle(inPath);
            logInfo(s, job, 'Job completed successfully');
            resolve();
          } catch(e) {
            logErr(s, job, 'Error: ' + e.message);
            await job.fail('Error: ' + e.message);
            reject(e);
          }
        });
      });
  
      req.on('error', async function(e) {
        logErr(s, job, 'HTTPS error: ' + e.message);
        await job.fail('Request failed: ' + e.message);
        reject(e);
      });
      
      req.setTimeout(180000, async function() {
        try { req.destroy(new Error('timeout')); } catch(_) {}
        logErr(s, job, 'Timeout');
        await job.fail('Timeout');
        reject(new Error('timeout'));
      });
  
      req.write(payload);
      req.end();
    });
  }