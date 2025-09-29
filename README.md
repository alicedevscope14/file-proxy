# README — File Proxy for Expenses (Azure Function + Graph + Dataverse)

This README documents the **technical** part of the solution that addresses access permissions to expense attachments through an **Azure Function** acting as a **proxy** between CRM (Dataverse) and SharePoint. It also includes the required **Microsoft Graph requests** to obtain the `listId` and to resolve the file from the list `ItemId`.

---

## Prerequisites
- Azure subscription.  
- Function App deployed with **Node.js 18 LTS** (or higher).  
- App Registration in Azure AD with:  
  - **Application permissions**: `Files.Read.All` (Microsoft Graph > Application)  and `user_impersonation` (Dynamics CRM > Delegated).  
- Admin consent granted.  

---

## 1) Overview (how it works)

1. The mobile/canvas app uploads the file to **SharePoint**.  
2. The Flow moves/copies the file to the **_ExpensesSent_ document library** and returns the **ItemId** (integer from the list).  
3. The record is created in **CRM (Dataverse)** and the Flow updates the item in SharePoint with the **`CrmExpenseId`** (CRM record GUID).  
4. In CRM, the “attachment link” field does **not** store the SharePoint link, but rather the link to the **Azure Function**:  
https://files-proxy-function-fjeyged6eugbcbfz.westeurope-01.azurewebsites.net/api/file-proxy?mode=spitem&siteHost=devscope365.sharepoint.com&sitePath=sites/time-expenses-report&listId=7cdef9b6-7680-4a94-be8c-7bc778d96160&spItemId=144

5. When someone clicks the link, the **Function**:  
- requires login (EasyAuth),  
- resolves the `driveItem.id`/`driveId` from the `spItemId`,  
- reads the `CrmExpenseId` in the item,  
- validates in **Dataverse** (impersonation) if the user has **Read** permission on the record,  
- if **YES** → streams the file; if **NO** → returns **403**.  

> **Proxy** = the Function is the **intermediary**: the user does not access SharePoint directly; the Function validates and delivers the file.

---

## 2) Requirements and permissions

### 2.1 Azure Function App
- **Authentication (EasyAuth)**: Provider **Microsoft**, **Require authentication = On**, **Return 401** to anonymous users.  
- **App Settings** (examples):  
- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` *(or `USE_MANAGED_IDENTITY=true`)*  
- `DATAVERSE_ORG_URL` → `https://devscope365.crm4.dynamics.com`  
- `CRM_ENTITY_LOGICAL_NAME` → e.g., `dev_expenses`  
- Default SharePoint values (optional, for shorter URL):  
 - `SP_SITE_HOST` → `devscope365.sharepoint.com`  
 - `SP_SITE_PATH` → `sites/time-expenses-report`  
 - `SP_LIST_ID` → `7cdef9b6-7680-4a94-be8c-7bc778d96160` *(ExpensesSent)*  

### 2.2 App Registration (Service Principal)
- **Graph (application permissions)**: `Files.Read.All` (Microsoft Graph > Application) and `user_impersonation` (Dynamics CRM > Delegated).  
- **Dataverse**: create an **Application User** with a **Security Role** that includes:  
- **Read** permission on the expenses table (`dev_expense(s)`),  
- **Act on Behalf of Another User** (impersonation).  
- **Grant site access** when using `Files.Read.All`.  


