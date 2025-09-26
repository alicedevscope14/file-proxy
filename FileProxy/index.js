import { ClientSecretCredential, ManagedIdentityCredential } from "@azure/identity";

// ===== Auth helpers =====
function getCredential() {
  return process.env.USE_MANAGED_IDENTITY === "true"
    ? new ManagedIdentityCredential()
    : new ClientSecretCredential(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);
}
async function getToken(cred, scope) {
  const t = await cred.getToken(scope);
  return t.token;
}
function getUserFromEasyAuth(req) {
  const b64 = req.headers["x-ms-client-principal"];
  if (!b64) {
    const err = new Error("EasyAuth: pedido sem identidade (verifica Require authentication = On).");
    err.status = 401;
    throw err;
  }
  const principal = JSON.parse(Buffer.from(b64, "base64").toString("utf8"));
  const claims = Object.fromEntries(principal.claims.map(c => [c.typ, c.val]));
  const oid = claims["http://schemas.microsoft.com/identity/claims/objectidentifier"];
  const upn = claims["http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn"];
  if (!oid) {
    const err = new Error("EasyAuth: sem OID do utilizador nas claims.");
    err.status = 401;
    throw err;
  }
  return { oid, upn };
}

// ===== HTTP helpers =====
async function httpJson(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const txt = await r.text().catch(() => "");
    const err = new Error(`HTTP ${r.status} at ${url}: ${txt}`);
    err.status = r.status;
    throw err;
  }
  return r.json();
}

async function httpRaw(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const txt = await r.text().catch(() => "");
    const err = new Error(`HTTP ${r.status} at ${url}: ${txt}`);
    err.status = r.status;
    throw err;
  }
  return r;
}

export default async function (context, req) {
  try {
    const q = req.query || {};
    const cred = getCredential();
    const graphToken = await getToken(cred, "https://graph.microsoft.com/.default");

    let driveId, driveItemId, fileName = "file", contentType = "application/octet-stream";
    let crmExpenseId = null;

    // ===== A) Receber por ItemId (SharePoint) =====
    if (q.mode === "spitem") {
      const siteHost = q.siteHost || process.env.SP_SITE_HOST;
      const sitePath = q.sitePath || process.env.SP_SITE_PATH;
      const spItemId = q.spItemId;
      const listId   = q.listId || process.env.SP_LIST_ID;

      if (!siteHost || !sitePath || !spItemId) {
        context.res = { status: 400, body: "Faltam parâmetros: siteHost/sitePath (ou app settings) e spItemId." };
        return;
      }

      // 1) siteId
      const site = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteHost)}:/${encodeURIComponent(sitePath)}?select=id`,
        graphToken
      );
      const siteId = site.id;

      // 2) listId
      let effListId = listId;
      if (!effListId) {
        if (!q.listTitle) {
          context.res = { status: 400, body: "Falta listId (ou SP_LIST_ID) ou listTitle." };
          return;
        }
        const lists = await httpJson(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName&$filter=displayName eq '${encodeURIComponent(q.listTitle)}'`,
          graphToken
        );
        if (!lists.value?.length) { context.res = { status: 404, body: "Lista não encontrada." }; return; }
        effListId = lists.value[0].id;
      }

      // 3) driveItem (para id/driveId/mime)
      const li = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${effListId}/items/${encodeURIComponent(spItemId)}?$expand=driveItem`,
        graphToken
      );
      // 4) fields (para CrmExpenseId)
      const fields = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${effListId}/items/${encodeURIComponent(spItemId)}/fields`,
        graphToken
      );

      crmExpenseId = fields?.CrmExpenseId || null;
      driveItemId  = li?.driveItem?.id;
      driveId      = li?.driveItem?.parentReference?.driveId;
      fileName     = li?.driveItem?.name || fileName;
      contentType  = li?.driveItem?.file?.mimeType || contentType;

      if (!driveItemId || !driveId) {
        context.res = { status: 404, body: "driveItem não encontrado para esse ItemId." };
        return;
      }
    }
    // ===== B) Receber por driveId + itemId (Graph) =====
    else {
      driveId = q.driveId;
      driveItemId = q.itemId;
      if (!driveId || !driveItemId) {
        context.res = { status: 400, body: "Parâmetros obrigatórios: mode=spitem&spItemId=... OU driveId & itemId." };
        return;
      }
      const meta = await httpJson(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}?select=name,file`,
        graphToken
      );
      fileName    = meta?.name || fileName;
      contentType = meta?.file?.mimeType || contentType;

      // fields → CrmExpenseId
      const fields = await httpJson(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}/listItem/fields`,
        graphToken
      );
      crmExpenseId = fields?.CrmExpenseId || null;
    }

    // ===== Validação com CRM (Dataverse) =====
    if (!crmExpenseId) {
      context.res = { status: 403, body: "Ficheiro sem CrmExpenseId associado." };
      return;
    }

    // 1) Utilizador autenticado (EasyAuth)
    const user = getUserFromEasyAuth(req);

    // 2) Token app-only para o Dataverse
    const dvOrg = process.env.DATAVERSE_ORG_URL; // ex: https://org.crm4.dynamics.com
    const dvToken = await getToken(cred, `${dvOrg}/.default`);

    // 3) Mapeia OID -> systemuserid
    //    (precisa que o utilizador exista e esteja Enabled no Dataverse)
    const usersUrl = `${dvOrg}/api/data/v9.2/systemusers`
      + `?$select=systemuserid,azureactivedirectoryobjectid,domainname,isdisabled`
      + `&$filter=azureactivedirectoryobjectid eq ${user.oid}`;
    const usersRes = await fetch(usersUrl, {
      headers: { Authorization: `Bearer ${dvToken}`, Accept: "application/json", "OData-Version": "4.0" }
    });
    if (usersRes.status !== 200) {
      context.log.warn(`Dataverse systemusers lookup falhou status=${usersRes.status}`);
      context.res = { status: 403, body: "Não tens acesso (utilizador não mapeado no CRM)." };
      return;
    }
    const users = await usersRes.json();
    const systemUserId = users?.value?.[0]?.systemuserid;
    const isDisabled = users?.value?.[0]?.isdisabled;
    if (!systemUserId || isDisabled) {
      context.log.warn(`Sem systemuser válido para OID=${user.oid} (disabled=${isDisabled})`);
      context.res = { status: 403, body: "Não tens acesso (utilizador não existe/está desativado no CRM)." };
      return;
    }

    // 4) Ler o registo como esse utilizador (impersonação estável com MSCRMCallerID)
    const entityLogicalName = process.env.CRM_ENTITY_LOGICAL_NAME || "dev_expenses"; // <- ajusta
    const dvUrl = `${dvOrg}/api/data/v9.2/${entityLogicalName}(${crmExpenseId})`;
    const dvRes = await fetch(dvUrl, {
      headers: {
        Authorization: `Bearer ${dvToken}`,
        Accept: "application/json",
        "OData-Version": "4.0",
        // Impersonação por Dataverse UserId
        MSCRMCallerID: systemUserId
      }
    });

    if (dvRes.status !== 200) {
      const body = await dvRes.text().catch(() => "");
      context.log.warn(`DENY crmExpenseId=${crmExpenseId} sysUser=${systemUserId} status=${dvRes.status} body=${body}`);
      context.res = { status: 403, body: "Não tens acesso a esta despesa no CRM." };
      return;
    }

    // ===== Conteúdo do ficheiro =====
    const content = await httpRaw(
      `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}/content`,
      graphToken
    );
    const buffer = Buffer.from(await content.arrayBuffer());
    const disposition = (process.env.DEFAULT_DISPOSITION || "inline") + `; filename="${fileName}"`;

    context.res = {
      status: 200,
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": disposition,
        "Cache-Control": "private, max-age=60"
      },
      body: buffer
    };
  } catch (err) {
    context.log.error(err);
    const status = err.status || 500;
    const msg = status === 401 ? "Não autenticado"
              : status === 403 ? "Sem autorização"
              : "Erro interno";
    context.res = { status, body: msg };
  }
}