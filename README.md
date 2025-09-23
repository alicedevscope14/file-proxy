# Files Proxy Function

This Azure Function acts as a secure proxy for downloading files from SharePoint (via Microsoft Graph API).  
Instead of exposing direct SharePoint links, CRM (or other apps) can generate links that point to this function.  
The function authenticates users with Azure AD (EasyAuth) and retrieves files on their behalf, returning them as downloads.

---

## Features
- Protects direct SharePoint file links.
- Uses **Azure AD EasyAuth** for user authentication.
- Supports **Microsoft Graph Sites.Selected** permissions (least privilege).
- Returns files either by:
  - **302 Redirect** to Graph’s `downloadUrl` (efficient).
  - **Stream** (hides the origin, but more resource-intensive).
- Extensible: external API can be integrated for fine-grained access validation (multi-customer scenarios).

---

## Prerequisites
- Azure subscription.
- Function App deployed with **Node.js 18 LTS** (or higher).
- App Registration in Azure AD with:
  - **Application permissions**: `Sites.Selected` (Microsoft Graph).
  - (Optional for Dataverse integration: `Dynamics CRM → Delegated → user_impersonation`).
- Admin consent granted.
qa