import { ClientSecretCredential, ManagedIdentityCredential } from "@azure/identity";

// ===== Auth helpers =====
function getCredential() {
  return process.env.USE_MANAGED_IDENTITY === "true"
    ? new ManagedIdentityCredential()
    : new ClientSecretCredential(
        process.env.TENANT_ID, 
        process.env.CLIENT_ID, 
        process.env.CLIENT_SECRET
      );
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
    
    // === ROTA DE DIAGNÓSTICO ===
    if (q.mode === "debug") {
      const cred = getCredential();
      const dvOrg = process.env.DATAVERSE_ORG_URL;
      
      if (!dvOrg) {
        context.res = { 
          status: 200, 
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ error: "DATAVERSE_ORG_URL não configurado" }, null, 2)
        };
        return;
      }
      
      try {
        // Obtém info do utilizador
        const user = getUserFromEasyAuth(req);
        
        // Token para Dataverse
        const dvToken = await getToken(cred, `${dvOrg}/.default`);
        
        // Procura o utilizador de várias formas
        const results = {};
        
        // 1. Por OID
        const oidUrl = `${dvOrg}/api/data/v9.2/systemusers?$select=systemuserid,fullname,domainname,azureactivedirectoryobjectid&$filter=azureactivedirectoryobjectid eq '${user.oid}'`;
        const oidRes = await fetch(oidUrl, {
          headers: { 
            Authorization: `Bearer ${dvToken}`, 
            Accept: "application/json", 
            "OData-Version": "4.0"
          }
        });
        results.byOid = {
          status: oidRes.status,
          found: oidRes.status === 200 ? (await oidRes.json()).value : []
        };
        
        // 2. Por UPN/domainname
        if (user.upn) {
          const upnUrl = `${dvOrg}/api/data/v9.2/systemusers?$select=systemuserid,fullname,domainname,azureactivedirectoryobjectid&$filter=domainname eq '${user.upn}'`;
          const upnRes = await fetch(upnUrl, {
            headers: { 
              Authorization: `Bearer ${dvToken}`, 
              Accept: "application/json", 
              "OData-Version": "4.0"
            }
          });
          results.byUpn = {
            status: upnRes.status,
            found: upnRes.status === 200 ? (await upnRes.json()).value : []
          };
        }
        
        // 3. Lista todos os utilizadores ativos (limitado a 10)
        const allUrl = `${dvOrg}/api/data/v9.2/systemusers?$select=systemuserid,fullname,domainname,azureactivedirectoryobjectid,isdisabled&$filter=isdisabled eq false&$top=10`;
        const allRes = await fetch(allUrl, {
          headers: { 
            Authorization: `Bearer ${dvToken}`, 
            Accept: "application/json", 
            "OData-Version": "4.0"
          }
        });
        
        let allUsersData = null;
        let errorDetails = null;
        
        if (allRes.status !== 200) {
          errorDetails = await allRes.text().catch(() => "Unable to read error");
        } else {
          allUsersData = await allRes.json();
        }
        
        results.sampleUsers = {
          status: allRes.status,
          errorDetails: errorDetails,
          users: allUsersData ? allUsersData.value.map(u => ({
            name: u.fullname,
            email: u.domainname,
            hasOid: !!u.azureactivedirectoryobjectid,
            oid: u.azureactivedirectoryobjectid
          })) : []
        };
        
        // 4. Testa acesso à entidade de despesas
        const testExpenseUrl = `${dvOrg}/api/data/v9.2/${process.env.CRM_ENTITY_LOGICAL_NAME || "dev_expense"}?$top=1`;
        const expenseRes = await fetch(testExpenseUrl, {
          headers: { 
            Authorization: `Bearer ${dvToken}`, 
            Accept: "application/json", 
            "OData-Version": "4.0"
          }
        });
        
        results.expenseEntityAccess = {
          status: expenseRes.status,
          entityName: process.env.CRM_ENTITY_LOGICAL_NAME || "dev_expense",
          hasAccess: expenseRes.status === 200,
          error: expenseRes.status !== 200 ? await expenseRes.text().catch(() => "") : null
        };
        
        context.res = {
          status: 200,
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            debug: true,
            currentUser: {
              oid: user.oid,
              upn: user.upn
            },
            searchResults: results,
            environment: {
              dataverseUrl: dvOrg,
              entityName: process.env.CRM_ENTITY_LOGICAL_NAME || "dev_expenses"
            }
          }, null, 2)
        };
        return;
        
      } catch (err) {
        context.res = {
          status: 200,
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            debug: true,
            error: err.message,
            stack: err.stack
          }, null, 2)
        };
        return;
      }
    }
    // === FIM DA ROTA DE DIAGNÓSTICO ===
    
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
        if (!lists.value?.length) { 
          context.res = { status: 404, body: "Lista não encontrada." }; 
          return; 
        }
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

      // DEBUG: Log todos os campos disponíveis
      context.log(`Campos disponíveis no item SharePoint ${spItemId}:`, Object.keys(fields));
      context.log(`Valor de CrmExpenseId:`, fields?.CrmExpenseId);

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
      try {
        const fields = await httpJson(
          `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}/listItem/fields`,
          graphToken
        );
        
        // DEBUG: Log dos campos
        context.log(`Campos disponíveis no driveItem:`, Object.keys(fields));
        context.log(`Valor de CrmExpenseId:`, fields?.CrmExpenseId);
        
        crmExpenseId = fields?.CrmExpenseId || null;
      } catch (err) {
        context.log.warn(`Não foi possível obter campos do listItem: ${err.message}`);
      }
    }

    // ===== Validação com CRM (Dataverse) =====
    if (!crmExpenseId) {
      context.res = { 
        status: 403, 
        body: "Ficheiro sem CrmExpenseId associado. Verifique se o campo existe e está preenchido no SharePoint." 
      };
      return;
    }

    // 1) Utilizador autenticado (EasyAuth)
    const user = getUserFromEasyAuth(req);
    context.log(`Utilizador autenticado - OID: ${user.oid}, UPN: ${user.upn}`);

    // 2) Token app-only para o Dataverse
    const dvOrg = process.env.DATAVERSE_ORG_URL; // ex: https://org.crm4.dynamics.com
    if (!dvOrg) {
      context.res = { status: 500, body: "DATAVERSE_ORG_URL não configurado." };
      return;
    }
    
    const dvToken = await getToken(cred, `${dvOrg}/.default`);

    // 3) Mapeia OID -> systemuserid 
    // NOTA: Tenta primeiro por OID, depois por UPN se falhar
    let usersUrl = `${dvOrg}/api/data/v9.2/systemusers`
      + `?$select=systemuserid,azureactivedirectoryobjectid,domainname,fullname,internalemailaddress,isdisabled`
      + `&$filter=azureactivedirectoryobjectid eq '${user.oid}'`;
    
    context.log(`Procurando utilizador no Dataverse por OID: ${user.oid}`);
    
    let usersRes = await fetch(usersUrl, {
      headers: { 
        Authorization: `Bearer ${dvToken}`, 
        Accept: "application/json", 
        "OData-Version": "4.0",
        "Prefer": "odata.include-annotations=*"
      }
    });
    
    // Se não encontrou por OID, tenta por UPN
    if (usersRes.status !== 200 || (await usersRes.clone().json()).value?.length === 0) {
      if (user.upn) {
        context.log(`Não encontrado por OID, tentando por UPN: ${user.upn}`);
        
        // Tenta por domainname (UPN)
        usersUrl = `${dvOrg}/api/data/v9.2/systemusers`
          + `?$select=systemuserid,azureactivedirectoryobjectid,domainname,fullname,internalemailaddress,isdisabled`
          + `&$filter=domainname eq '${user.upn}' or internalemailaddress eq '${user.upn}'`;
        
        usersRes = await fetch(usersUrl, {
          headers: { 
            Authorization: `Bearer ${dvToken}`, 
            Accept: "application/json", 
            "OData-Version": "4.0",
            "Prefer": "odata.include-annotations=*"
          }
        });
      }
    }
    
    if (usersRes.status !== 200) {
      const errorText = await usersRes.text().catch(() => "");
      context.log.error(`Dataverse systemusers lookup falhou status=${usersRes.status}, erro: ${errorText}`);
      
      // Log adicional para debug
      context.log.error(`Debug - OID tentado: ${user.oid}`);
      context.log.error(`Debug - UPN tentado: ${user.upn}`);
      context.log.error(`Debug - URL completa: ${usersUrl}`);
      
      context.res = { 
        status: 403, 
        body: `Não tens acesso (utilizador não mapeado no CRM). Status: ${usersRes.status}. OID: ${user.oid?.substring(0,8)}...` 
      };
      return;
    }
    
    const users = await usersRes.json();
    context.log(`Utilizadores encontrados: ${users?.value?.length || 0}`);
    
    if (users?.value?.length > 0) {
      const user = users.value[0];
      context.log(`Utilizador CRM encontrado:`);
      context.log(`  - Nome: ${user.fullname}`);
      context.log(`  - SystemUserId: ${user.systemuserid}`);
      context.log(`  - Email/Domain: ${user.domainname || user.internalemailaddress}`);
      context.log(`  - Azure OID no CRM: ${user.azureactivedirectoryobjectid || 'NÃO PREENCHIDO'}`);
    }
    
    const systemUserId = users?.value?.[0]?.systemuserid;
    const isDisabled = users?.value?.[0]?.isdisabled;
    
    if (!systemUserId) {
      context.log.error(`Sem systemuser para OID=${user.oid}`);
      context.res = { 
        status: 403, 
        body: "Não tens acesso (utilizador não existe no CRM). Confirma que o teu utilizador está sincronizado." 
      };
      return;
    }
    
    if (isDisabled) {
      context.log.warn(`Utilizador desativado: OID=${user.oid}`);
      context.res = { 
        status: 403, 
        body: "Não tens acesso (utilizador está desativado no CRM)." 
      };
      return;
    }

    // 4) Ler o registo como esse utilizador (impersonação com MSCRMCallerID)
    const entityLogicalName = process.env.CRM_ENTITY_LOGICAL_NAME || "dev_expenses";
    
    // Formatar o GUID corretamente (sem chavetas se tiver)
    const formattedCrmExpenseId = crmExpenseId.replace(/[{}]/g, '');
    
    const dvUrl = `${dvOrg}/api/data/v9.2/${entityLogicalName}(${formattedCrmExpenseId})`;
    context.log(`Verificando acesso à despesa: ${dvUrl} como utilizador ${systemUserId}`);
    
    const dvRes = await fetch(dvUrl, {
      headers: {
        Authorization: `Bearer ${dvToken}`,
        Accept: "application/json",
        "OData-Version": "4.0",
        "Prefer": "odata.include-annotations=*",
        // Impersonação por Dataverse UserId
        "MSCRMCallerID": systemUserId
      }
    });

    if (dvRes.status === 404) {
      context.log.warn(`Despesa não encontrada: ${formattedCrmExpenseId}`);
      context.res = { 
        status: 404, 
        body: `Despesa ${formattedCrmExpenseId} não encontrada no CRM.` 
      };
      return;
    }

    if (dvRes.status === 403 || dvRes.status === 401) {
      const body = await dvRes.text().catch(() => "");
      context.log.warn(`DENY acesso negado - ExpenseId=${formattedCrmExpenseId} User=${systemUserId} Status=${dvRes.status} Body=${body}`);
      context.res = { 
        status: 403, 
        body: "Não tens permissões para aceder a esta despesa no CRM." 
      };
      return;
    }

    if (dvRes.status !== 200) {
      const body = await dvRes.text().catch(() => "");
      context.log.error(`Erro inesperado - ExpenseId=${formattedCrmExpenseId} Status=${dvRes.status} Body=${body}`);
      context.res = { 
        status: 500, 
        body: `Erro ao verificar acesso no CRM (Status: ${dvRes.status})` 
      };
      return;
    }

    // Sucesso - utilizador tem acesso
    const expenseData = await dvRes.json();
    context.log(`ALLOW acesso permitido - Despesa: ${expenseData?.dev_name || formattedCrmExpenseId}`);

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
        "Cache-Control": "private, max-age=60",
        "X-CRM-ExpenseId": formattedCrmExpenseId // Para debug
      },
      body: buffer
    };
    
  } catch (err) {
    context.log.error(`Erro geral: ${err.message}`, err.stack);
    const status = err.status || 500;
    const msg = status === 401 ? "Não autenticado"
              : status === 403 ? "Sem autorização"
              : `Erro interno: ${err.message}`;
    context.res = { status, body: msg };
  }
}