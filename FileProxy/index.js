import { ClientSecretCredential, ManagedIdentityCredential } from "@azure/identity";

/**
 * Helper: token do Graph
 */
async function getGraphToken() {
  const scope = "https://graph.microsoft.com/.default";
  const useMI = process.env.USE_MANAGED_IDENTITY === "true";

  if (useMI) {
    const cred = new ManagedIdentityCredential();
    const token = await cred.getToken(scope);
    return token.token;
  } else {
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
    const token = await cred.getToken(scope);
    return token.token;
  }
}

/**
 * Helper: GET JSON do Graph
 */
async function graphGetJson(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) return null;
  return await r.json();
}

export default async function (context, req) {
  try {
    const { driveId, itemId } = req.query || {};

    if (!driveId || !itemId) {
      context.res = { status: 400, body: "Parâmetros obrigatórios: driveId, itemId." };
      return;
    }

    const token = await getGraphToken();

    // 1) Metadados (nome, mime type)
    const metaUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
    const meta = await graphGetJson(metaUrl, token);
    if (!meta) {
      context.res = { status: 404, body: "Ficheiro não encontrado." };
      return;
    }

    const fileName = meta?.name || "file";
    const contentType = meta?.file?.mimeType || "application/octet-stream";
    const disposition = (process.env.DEFAULT_DISPOSITION || "inline") + `; filename="${fileName}"`;

    // 2) Conteúdo
    const contentUrl = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`;
    const fileRes = await fetch(contentUrl, { headers: { Authorization: `Bearer ${token}` } });

    if (!fileRes.ok) {
      context.res = { status: fileRes.status, body: "Conteúdo não disponível." };
      return;
    }

    // SIMPLES: buffer em memória (ok para PDFs/imagens típicas)
    const buffer = Buffer.from(await fileRes.arrayBuffer());
    context.res = {
      status: 200,
      headers: {
        "Content-Type": contentType,
        "Content-Disposition": disposition,
        "Cache-Control": "private, max-age=60"
      },
      body: buffer
    };

    // Se quiseres evitar buffer em memória (ficheiros muito grandes):
    // Usa body como stream e isRaw: true (depende do host). Mantive simples aqui.
  } catch (err) {
    context.log.error(err);
    context.res = { status: 500, body: "Erro interno ao processar o ficheiro." };
  }
}
