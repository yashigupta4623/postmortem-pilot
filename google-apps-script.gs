/**
 * Postmortem Pilot — "Publish" backend.
 * Deploy this as a Google Apps Script Web App. It creates a Google Doc from the
 * incident data the tool sends, then emails you the link. Runs as YOUR Google
 * account, so no API keys are ever exposed in the website.
 *
 * SETUP (one time, ~2 min):
 *  1. Go to https://script.google.com  →  New project
 *  2. Delete the sample code, paste ALL of this file, click 💾 Save
 *  3. Deploy  →  New deployment  →  gear ⚙  →  Web app
 *       - Description: Postmortem Pilot
 *       - Execute as:  Me
 *       - Who has access:  Anyone
 *     → Deploy  → Authorize access  → allow the Google Docs + Gmail permissions
 *  4. Copy the "Web app" URL (it ends in /exec)
 *  5. Paste that URL into the tool's "Publish setup" box, add your email → done.
 *     Now say "publish" and a Google Doc lands in your inbox.
 */
function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    var title = d.title || 'Incident Postmortem';

    var doc = DocumentApp.create(title);
    var b = doc.getBody();
    b.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    var meta = [d.sev, d.window, 'Status: Resolved'].filter(Boolean).join('    |    ');
    b.appendParagraph(meta).editAsText().setForegroundColor('#5f6368');

    function heading(t){ b.appendParagraph(t).setHeading(DocumentApp.ParagraphHeading.HEADING2); }
    function bullets(items){
      if (items && items.length) items.forEach(function(t){ b.appendListItem(t).setGlyphType(DocumentApp.GlyphType.BULLET); });
      else b.appendParagraph('—').editAsText().setForegroundColor('#9aa0a6');
    }

    heading('TL;DR');
    b.appendParagraph(d.tldr || '');
    heading('Impact');            bullets(d.impact);
    heading('Timeline');
    (d.timeline || []).forEach(function(t){ b.appendParagraph(t.ts + '   —   ' + t.text); });
    heading('Root cause');        bullets(d.rootCause);
    heading('Fix / mitigation');  bullets(d.fix);
    heading('Prevention / action items'); bullets(d.prevention);

    doc.saveAndClose();
    var url = doc.getUrl();

    var to = d.email || Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: 'Postmortem: ' + title,
      htmlBody: 'Your incident postmortem is ready.<br><br>' +
                '<a href="' + url + '">📄 Open the Google Doc</a><br><br>— Postmortem Pilot'
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true, url: url }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
