import { ClientSecretCredential, ManagedIdentityCredential } from "@azure/identity";

function getCredential() {
  return process.env.USE_MANAGED_IDENTITY === "true"
    ? new ManagedIdentityCredential()
    : new ClientSecretCredential(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);
}
async function getToken(cred, scope) {
  const t = await cred.getToken(scope);
  return t.token;
}
async function httpJson(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw Object.assign(new Error(await r.text()), { status: r.status, url });
  return r.json();
}
async function httpRaw(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) throw Object.assign(new Error(await r.text()), { status: r.status, url });
  return r;
}

export default async function (context, req) {
  try {
    const q = req.query || {};
    const cred = getCredential();
    const graphToken = await getToken(cred, "https://graph.microsoft.com/.default");

    let driveId, driveItemId, fileName = "file", contentType = "application/octet-stream";
    let crmExpenseId = null; // para futura validação

    // ==============================
    // A) Receber por ItemId (SharePoint)
    // ==============================
    if (q.mode === "spitem") {
      const siteHost = q.siteHost || process.env.SP_SITE_HOST;
      const sitePath = q.sitePath || process.env.SP_SITE_PATH;
      const spItemId = q.spItemId;
      if (!siteHost || !sitePath || !spItemId) {
        context.res = { status: 400, body: "Faltam parâmetros: siteHost/sitePath (ou definir em app settings) e spItemId." };
        return;
      }

      // 1) siteId
      const site = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteHost)}:/${encodeURIComponent(sitePath)}?select=id`,
        graphToken
      );
      const siteId = site.id;

      // 2) listId
      let listId = q.listId || process.env.SP_LIST_ID;
      if (!listId) {
        if (!q.listTitle) {
          context.res = { status: 400, body: "Falta listId (ou SP_LIST_ID) ou listTitle." };
          return;
        }
        const lists = await httpJson(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName&$filter=displayName eq '${encodeURIComponent(q.listTitle)}'`,
          graphToken
        );
        if (!lists.value?.length) { context.res = { status: 404, body: "Lista não encontrada." }; return; }
        listId = lists.value[0].id;
      }

      // 3) obter driveItem + fields a partir do ItemId
      const li = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(spItemId)}?$expand=driveItem`,
        graphToken
      );
      // (se precisares de metadados/CrmExpenseId:)
      const fields = await httpJson(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(spItemId)}/fields`,
        graphToken
      );
      crmExpenseId = fields?.CrmExpenseId || null;

      driveItemId = li?.driveItem?.id;
      driveId = li?.driveItem?.parentReference?.driveId;
      fileName = li?.driveItem?.name || fileName;
      contentType = li?.driveItem?.file?.mimeType || contentType;

      if (!driveItemId || !driveId) {
        context.res = { status: 404, body: "driveItem não encontrado para esse ItemId." };
        return;
      }
    }
    // ==============================
    // B) Receber por driveId + itemId (Graph)
    // ==============================
    else {
      driveId = q.driveId;
      driveItemId = q.itemId;
      if (!driveId || !driveItemId) {
        context.res = { status: 400, body: "Parâmetros obrigatórios: mode=spitem&spItemId=... (com site/list) OU driveId & itemId." };
        return;
      }
      // Metadados mínimos (nome/mime)
      const meta = await httpJson(
        `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}?select=name,file`,
        graphToken
      );
      fileName = meta?.name || fileName;
      contentType = meta?.file?.mimeType || contentType;
      // (se fores usar validação por CrmExpenseId quando vier por Graph IDs:)
      // const fields = await httpJson(
      //   `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}/listItem/fields`,
      //   graphToken
      // );
      // crmExpenseId = fields?.CrmExpenseId || null;
    }

    // ==============================
    // (Opcional) Validação CRM — ativa depois
    // ==============================
    // if (!crmExpenseId) return context.res = { status: 403, body: "Ficheiro sem CrmExpenseId associado." };
    // const user = getUserFromEasyAuth(req); // ler OID/UPN do cabeçalho X-MS-CLIENT-PRINCIPAL (EasyAuth)
    // const dvToken = await getToken(cred, `${process.env.DATAVERSE_ORG_URL}/.default`);
    // const dvRes = await fetch(`${process.env.DATAVERSE_ORG_URL}/api/data/v9.2/new_expenses(${crmExpenseId})`, {
    //   headers: { Authorization: `Bearer ${dvToken}`, "OData-Version": "4.0", Accept: "application/json", CallerObjectId: user.oid }
    // });
    // if (dvRes.status !== 200) return context.res = { status: 403, body: "Não tens acesso a esta despesa no CRM." };

    // ==============================
    // Download do conteúdo e resposta
    // ==============================
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
    context.res = { status: err.status || 500, body: err.status ? `Erro ${err.status}` : "Erro interno" };
  }
}
