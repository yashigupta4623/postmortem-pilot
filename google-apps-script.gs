/**
 * Postmortem Pilot — AI "Publish" backend.
 * Deploy as a Google Apps Script Web App. When the tool sends your raw incident
 * narration, this uses a free Gemini model to write a properly-segregated RCA,
 * creates a Google Doc, and emails you the link. Runs as YOUR Google account,
 * so your API key is never exposed in the website.
 *
 * ── SETUP (one time, ~5 min) ──────────────────────────────────────────────
 * 1. Get a FREE Gemini API key:  https://aistudio.google.com/apikey  → Create key → copy it.
 * 2. Go to  https://script.google.com  → New project.
 * 3. Delete the sample code, paste ALL of this file.
 * 4. Paste your key into GEMINI_API_KEY below, click 💾 Save.
 * 5. Deploy → New deployment → ⚙ → Web app
 *       Execute as: Me   |   Who has access: Anyone   → Deploy
 *       → Authorize access → allow Google Docs + Gmail + external-fetch permissions.
 * 6. Copy the "Web app" URL (ends in /exec).
 * 7. In the tool's "Publish setup" box, paste that URL + your email. Done.
 *    Say "publish" → an AI-written Google Doc lands in your inbox.
 * ──────────────────────────────────────────────────────────────────────────
 */

var GEMINI_API_KEY = '';                 // ← paste your free key from aistudio.google.com/apikey
var GEMINI_MODEL   = 'gemini-2.0-flash'; // if this errors, try 'gemini-2.5-flash' or 'gemini-1.5-flash'

// Open the /exec URL in your browser to run this self-test. It tells you exactly
// what works and what fails — doc creation, email, and whether the Gemini key is set.
function doGet(e) {
  try {
    var sample = {
      title: 'Postmortem Pilot — TEST',
      sev: 'SEV2',
      email: Session.getEffectiveUser().getEmail(),
      raw: "I can't SSH into the EC2 instance, the connection times out. " +
           "Root cause is a security group that removed the port 22 rule. " +
           "I re-added the inbound rule and it works again. " +
           "We should add a guard so port 22 changes require review."
    };
    var usedGemini = !!GEMINI_API_KEY, gemOk = true, gemErr = '';
    var s;
    try { s = (GEMINI_API_KEY ? structureWithGemini(sample) : fallbackStructure(sample)); }
    catch (err2) { gemOk = false; gemErr = String(err2); s = fallbackStructure(sample); }
    var doc = buildDoc(s);
    MailApp.sendEmail({ to: sample.email, subject: 'Postmortem Pilot — TEST',
      htmlBody: 'Test doc: <a href="' + doc.getUrl() + '">' + doc.getUrl() + '</a>' });
    return HtmlService.createHtmlOutput(
      '<h2>✅ SUCCESS — email sent</h2>' +
      '<p>Google Doc created: <a href="' + doc.getUrl() + '">' + doc.getUrl() + '</a></p>' +
      '<p>Emailed to: <b>' + sample.email + '</b></p>' +
      '<p>Gemini AI structuring: <b>' +
        (!usedGemini ? 'OFF — key is empty (using built-in fallback, email still works)' :
         gemOk ? 'ON ✅' : 'FAILED, used fallback ⚠️ — ' + gemErr) + '</b></p>');
  } catch (err) {
    return HtmlService.createHtmlOutput('<h2>❌ ERROR</h2><pre>' + String(err) + '</pre>' +
      '<p>Most likely a bad Gemini API key or model name. Try GEMINI_MODEL = "gemini-2.5-flash" or "gemini-1.5-flash", or empty GEMINI_API_KEY to use the fallback.</p>');
  }
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    // AI structuring is best-effort — if Gemini errors, we STILL build the doc and email it.
    var s;
    try { s = (GEMINI_API_KEY ? structureWithGemini(d) : fallbackStructure(d)); }
    catch (gemErr) { s = fallbackStructure(d); }
    var doc = buildDoc(s);
    var url = doc.getUrl();
    var to = d.email || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: 'Postmortem: ' + s.title,
      htmlBody: 'Your incident postmortem is ready.<br><br>' +
                '<a href="' + url + '">📄 Open the Google Doc</a><br><br>— Postmortem Pilot'
    });
    return json({ ok: true, url: url });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Ask Gemini to segregate the raw narration into proper RCA sections. */
function structureWithGemini(d) {
  var transcript = d.raw || (d.timeline || []).map(function (t) { return t.text; }).join('\n');
  var prompt =
    'You are a senior SRE writing an incident postmortem from a raw spoken transcript. ' +
    'Read it and sort every statement into the correct section. ' +
    'Respond with STRICT JSON only, no markdown, with these keys: ' +
    'title (a short incident title), severity (SEV1/SEV2/SEV3), tldr (one sentence), ' +
    'impact (array of short bullet strings), timeline (array of short strings in order), ' +
    'rootCause (array), fix (array), prevention (array). ' +
    'Be concise and professional. If a section has nothing, use an empty array.\n\n' +
    'Known title: ' + (d.title || '(none)') + '\nSeverity: ' + (d.sev || 'SEV2') +
    '\n\nTRANSCRIPT:\n' + transcript;

  var resp = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + GEMINI_API_KEY,
    {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
      })
    });

  var body = JSON.parse(resp.getContentText());
  if (!body.candidates || !body.candidates[0]) throw new Error('Gemini: ' + resp.getContentText());
  var s = JSON.parse(body.candidates[0].content.parts[0].text);
  s.title    = s.title    || d.title || 'Incident Postmortem';
  s.severity = s.severity || d.sev   || 'SEV2';
  return s;
}

/** Used when no Gemini key is set — keeps the tool's own classification. */
function fallbackStructure(d) {
  return {
    title: d.title || 'Incident Postmortem',
    severity: d.sev || 'SEV2',
    tldr: d.tldr || '',
    impact: d.impact || [],
    timeline: (d.timeline || []).map(function (t) { return t.ts + ' — ' + t.text; }),
    rootCause: d.rootCause || [],
    fix: d.fix || [],
    prevention: d.prevention || []
  };
}

function buildDoc(s) {
  var doc = DocumentApp.create(s.title || 'Incident Postmortem');
  var b = doc.getBody();
  b.appendParagraph(s.title).setHeading(DocumentApp.ParagraphHeading.TITLE);
  b.appendParagraph([s.severity, 'Status: Resolved'].filter(Boolean).join('    |    '))
    .editAsText().setForegroundColor('#5f6368');

  function sec(name, items) {
    b.appendParagraph(name).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (items && items.length) items.forEach(function (t) { b.appendListItem(String(t)).setGlyphType(DocumentApp.GlyphType.BULLET); });
    else b.appendParagraph('—').editAsText().setForegroundColor('#9aa0a6');
  }

  b.appendParagraph('TL;DR').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  b.appendParagraph(s.tldr || '');
  sec('Impact', s.impact);
  b.appendParagraph('Timeline').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  (s.timeline || []).forEach(function (t) { b.appendParagraph(String(t)); });
  sec('Root cause', s.rootCause);
  sec('Fix / mitigation', s.fix);
  sec('Prevention / action items', s.prevention);
  doc.saveAndClose();
  return doc;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
