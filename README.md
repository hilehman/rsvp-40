# RSVP – 40th anniversary evening

Minimal RSVP landing page (Hebrew, RTL) hosted on GitHub Pages.

- `index.html` – the page. Submissions are POSTed to a Google Apps Script web app (`SCRIPT_URL`).
- `Code.gs` – the Apps Script backend (copy of the deployed script). It writes each response to a
  Google Sheet and keeps a live summary tab. Re-submitting with the same first + last name updates
  the existing row.
