# Frontend (React + Office.js)

This will host the Outlook add-in UI. Use Microsoft's React quickstart for Outlook add-ins to scaffold:
- https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart-react

After scaffolding, ensure:
- API base URL points to `http://127.0.0.1:8000`
- Add a query form that calls `POST /query` and renders the answer + sources
- Keep `manifest.xml` updated for sideloading
