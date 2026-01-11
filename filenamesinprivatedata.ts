/// <reference types="switch-scripting" />
// @ts-nocheck
// Enfocus Switch Script: Collect filenames in job -> semicolon-delimited -> store in Private Data
// - Reads all files in the job folder (or the parent folder if the job is a single file)
// - Excludes files matching an optional glob pattern (supports * wildcard; ? supported too)
// - Stores the resulting "file1.pdf;file2.pdf;..." into Private Data under a user-defined key

function getScriptTimeout() {
	return 300; // seconds
  }
  
  function getSettingsDefinition() {
	return (
	  '<?xml version="1.0" encoding="UTF-8"?>' +
	  '<settings>' +
		'<setting name="privatedatakey" displayName="Private Data Key" type="string" required="Yes">' +
		  '<description>Private data key where the semicolon-delimited filename list will be stored</description>' +
		'</setting>' +
		'<setting name="excludefilepattern" displayName="Exclude File Pattern" type="string" required="No">' +
		  '<description>Optional glob pattern(s) to EXCLUDE (e.g. "*PO-*"). Use * as wildcard. You can provide multiple patterns separated by ; or ,</description>' +
		'</setting>' +
	  '</settings>'
	);
  }
  
  function getDefaultSettings() {
	return {
	  privatedatakey: 'FileNameList',
	  excludefilepattern: ''
	};
  }
  
  /* -------- helpers -------- */
  function logInfo(s, job, m){ try{ job.log(LogLevel.Info, m); }catch(_){ try{s.log(3,m);}catch(__){} } }
  function logWarn(s, job, m){ try{ job.log(LogLevel.Warning, m); }catch(_){ try{s.log(2,m);}catch(__){} } }
  function logErr(s, job, m){ try{ job.log(LogLevel.Error, m); }catch(_){ try{s.log(1,m);}catch(__){} } }
  
  function globToRegex(pattern) {
	if (!pattern || !String(pattern).trim()) return null;
	var escaped = String(pattern)
	  .trim()
	  .replace(/[.+^${}()|[\]\\]/g, '\\$&') // escape regex chars
	  .replace(/\*/g, '.*')                // * -> any
	  .replace(/\?/g, '.');                // ? -> single
	return new RegExp('^' + escaped + '$', 'i');
  }
  
  function matchesAnyPattern(filename, patternListStr) {
	if (!patternListStr || !String(patternListStr).trim()) return false;
  
	// Allow multiple patterns separated by ; or ,
	var parts = String(patternListStr)
	  .split(/[;,]+/g)
	  .map(function(p){ return p.trim(); })
	  .filter(function(p){ return p.length > 0; });
  
	for (var i = 0; i < parts.length; i++) {
	  var rx = globToRegex(parts[i]);
	  if (rx && rx.test(filename)) return true;
	}
	return false;
  }
  
  async function tryGetProp(f, name){
	try{
	  if (f && typeof f.getPropertyValue === 'function'){
		var v = f.getPropertyValue(name);
		if (v && typeof v.then === 'function') v = await v;
		if (v != null) return String(v);
	  }
	}catch(_){}
	try{
	  if (f && typeof f.getPropertyStringValue === 'function'){
		var v2 = f.getPropertyStringValue(name);
		if (v2 && typeof v2.then === 'function') v2 = await v2;
		if (v2 != null) return String(v2);
	  }
	}catch(_){}
	return '';
  }
  
  async function getVal(f, names){
	for (var i=0;i<names.length;i++){
	  var v = await tryGetProp(f, names[i]);
	  if (v) return v;
	}
	return '';
  }
  
  async function getIncomingPath(job){
	try{
	  if (typeof AccessLevel !== 'undefined' && job && typeof job.get === 'function'){
		var p = job.get(AccessLevel.ReadOnly);
		if (p && typeof p.then === 'function') p = await p;
		return p;
	  }
	}catch(_){}
	if (job && typeof job.getPath === 'function'){
	  var p2 = job.getPath();
	  if (p2 && typeof p2.then === 'function') p2 = await p2;
	  return p2;
	}
	throw new Error('Cannot resolve job path');
  }
  
  /* -------- entry -------- */
  async function jobArrived(s, f, job){
	var fs = require('fs');
	var path = require('path');
  
	logInfo(s, job, '=== Filename List Script Starting ===');
  
	try {
	  var privateDataKey = (await getVal(f, ['privatedatakey'])) || 'FileNameList';
	  var excludePattern = await getVal(f, ['excludefilepattern']);
  
	  if (!privateDataKey || !String(privateDataKey).trim()){
		logErr(s, job, 'Missing privatedatakey setting');
		await job.fail('Missing Private Data Key setting');
		return;
	  }
  
	  var inPath = await getIncomingPath(job);
	  logInfo(s, job, 'Incoming path: ' + inPath);
  
	  // If job is a file, use its parent folder to find "all files within the job"
	  // If job is a folder, use it directly
	  var scanDir = inPath;
	  var st;
	  try {
		st = fs.statSync(inPath);
	  } catch (e) {
		logErr(s, job, 'Cannot stat path: ' + e.message);
		await job.fail('Cannot access job path');
		return;
	  }
  
	  if (st.isFile()) {
		scanDir = path.dirname(inPath);
		logInfo(s, job, 'Job is a file; scanning parent folder: ' + scanDir);
	  } else if (st.isDirectory()) {
		logInfo(s, job, 'Job is a folder; scanning folder for files.');
	  } else {
		logWarn(s, job, 'Job path is neither file nor folder; storing empty list.');
		await job.setPrivateData(privateDataKey, '');
		job.sendToSingle(await job.get(AccessLevel.ReadOnly));
		return;
	  }
  
	  var entries;
	  try {
		entries = fs.readdirSync(scanDir);
	  } catch (e2) {
		logErr(s, job, 'Cannot read directory: ' + e2.message);
		await job.fail('Cannot read job folder');
		return;
	  }
  
	  var names = [];
	  for (var i = 0; i < entries.length; i++) {
		var full = path.join(scanDir, entries[i]);
		var est;
		try { est = fs.statSync(full); } catch (_) { continue; }
		if (!est.isFile()) continue;
  
		var base = path.basename(entries[i]);
  
		// Exclude by pattern(s)
		if (matchesAnyPattern(base, excludePattern)) {
		  logInfo(s, job, 'Excluding by pattern: ' + base);
		  continue;
		}
  
		names.push(base);
	  }
  
	  // Sort for stability (optional)
	  names.sort(function(a,b){ return a.localeCompare(b); });
  
	  var list = names.join(';');
	  logInfo(s, job, 'Collected ' + names.length + ' filenames.');
	  logInfo(s, job, 'Storing into private data key "' + privateDataKey + '": ' + (list.length > 200 ? (list.substring(0,200) + '...') : list));
  
	  await job.setPrivateData(privateDataKey, list);
  
	  // Send along
	  var outPath = await job.get(AccessLevel.ReadOnly);
	  job.sendToSingle(outPath);
  
	  logInfo(s, job, '=== Filename List Script Completed ===');
  
	} catch (e) {
	  logErr(s, job, 'Unhandled error: ' + e.message);
	  logErr(s, job, 'Stack: ' + (e.stack || 'no stack'));
	  job.fail('Error: ' + e.message);
	}
  }
  