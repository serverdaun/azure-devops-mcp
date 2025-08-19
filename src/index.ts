#!/usr/bin/env node

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as azdev from "azure-devops-node-api";
import { AccessToken, AzureCliCredential, ChainedTokenCredential, DefaultAzureCredential, TokenCredential, OnBehalfOfCredential } from "@azure/identity";
import yargs from "yargs";
import { hideBin } from "yargs/helpers";

import { configurePrompts } from "./prompts.js";
import { configureAllTools } from "./tools.js";
import { UserAgentComposer } from "./useragent.js";
import { packageVersion } from "./version.js";

// Parse command line arguments using yargs
const argv = yargs(hideBin(process.argv))
  .scriptName("mcp-server-azuredevops")
  .usage("Usage: $0 <organization> [options]")
  .version(packageVersion)
  .command("$0 <organization>", "Azure DevOps MCP Server", (yargs) => {
    yargs.positional("organization", {
      describe: "Azure DevOps organization name",
      type: "string",
    });
  })
  .option("tenant", {
    alias: "t",
    describe: "Azure tenant ID (optional, required for multi-tenant scenarios)",
    type: "string",
  })
  .help()
  .parseSync();

export const orgName = (argv.organization as string) || (process.env.ADO_ORG as string);
const tenantId = argv.tenant || process.env.AZURE_AD_TENANT_ID;
if (!orgName) {
  throw new Error("Azure DevOps organization is required. Provide as positional arg or set ADO_ORG env var.");
}
const orgUrl = "https://dev.azure.com/" + orgName;

async function getAzureDevOpsToken(): Promise<AccessToken> {
  const scope = "499b84ac-1321-427f-aa17-267ca6975798/.default";
  const authMode = (process.env.ADO_AUTH || "").toLowerCase();

  // On-Behalf-Of flow using the user's bearer (assertion)
  if (authMode === "obo") {
    const userAssertion = process.env.MCP_USER_ASSERTION || process.env.ADO_USER_ASSERTION;
    const envTenant = process.env.AZURE_AD_TENANT_ID || process.env.AZURE_TENANT_ID || tenantId;
    const clientId = process.env.AZURE_AD_CLIENT_ID || process.env.AZURE_CLIENT_ID;
    const clientSecret = process.env.AZURE_AD_CLIENT_SECRET || process.env.AZURE_CLIENT_SECRET;

    console.error("[ADO][OBO] Starting OBO token acquisition", {
      hasUserAssertion: userAssertion,
      hasTenant: envTenant,
      hasClientId: clientId,
      hasClientSecret: clientSecret,
      orgUrl,
    });

    if (!userAssertion) {
      throw new Error("ADO_AUTH=obo is set but MCP_USER_ASSERTION was not provided.");
    }
    if (!envTenant || !clientId || !clientSecret) {
      throw new Error("ADO_AUTH=obo requires AZURE_(AD_)TENANT_ID, AZURE_(AD_)CLIENT_ID and AZURE_(AD_)CLIENT_SECRET to be set.");
    }

    const obo = new OnBehalfOfCredential({ tenantId: envTenant, clientId, clientSecret, userAssertionToken: userAssertion });
    const token = await obo.getToken(scope);
    console.error("[ADO][OBO] Token acquired", { expiresOn: token?.expiresOnTimestamp });
    if (!token) {
      throw new Error("On-behalf-of flow failed to acquire Azure DevOps token.");
    }
    return token;
  }

  // Fallback: DefaultAzureCredential/CLI chain
  if (process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS) {
    process.env.AZURE_TOKEN_CREDENTIALS = process.env.ADO_MCP_AZURE_TOKEN_CREDENTIALS;
  } else {
    process.env.AZURE_TOKEN_CREDENTIALS = "dev";
  }
  let credential: TokenCredential = new DefaultAzureCredential(); // CodeQL [SM05138] resolved by explicitly setting AZURE_TOKEN_CREDENTIALS
  if (tenantId) {
    // Use Azure CLI credential if tenantId is provided for multi-tenant scenarios
    const azureCliCredential = new AzureCliCredential({ tenantId });
    credential = new ChainedTokenCredential(azureCliCredential, credential);
  }

  const token = await credential.getToken(scope);
  if (!token) {
    throw new Error("Failed to obtain Azure DevOps token. Ensure you have Azure CLI logged in or another token source setup correctly.");
  }
  return token;
}

function getAzureDevOpsClient(userAgentComposer: UserAgentComposer): () => Promise<azdev.WebApi> {
  return async () => {
    const token = await getAzureDevOpsToken();
    const authMode = (process.env.ADO_AUTH || "").toLowerCase();
    let connection: azdev.WebApi;
    if (authMode === "obo" && typeof (azdev.WebApi as any).createWithBearerToken === "function") {
      connection = (azdev.WebApi as any).createWithBearerToken(orgUrl, token.token, undefined, {
        productName: "AzureDevOps.MCP",
        productVersion: packageVersion,
        userAgent: userAgentComposer.userAgent,
      });
    } else {
      const authHandler = azdev.getBearerHandler(token.token);
      connection = new azdev.WebApi(orgUrl, authHandler, undefined, {
        productName: "AzureDevOps.MCP",
        productVersion: packageVersion,
        userAgent: userAgentComposer.userAgent,
      });
    }
    console.error("[ADO] WebApi client initialized", { orgUrl, authMode });
    return connection;
  };
}

async function main() {
  const server = new McpServer({
    name: "Azure DevOps MCP Server",
    version: packageVersion,
  });

  const userAgentComposer = new UserAgentComposer(packageVersion);
  server.server.oninitialized = () => {
    userAgentComposer.appendMcpClientInfo(server.server.getClientVersion());
  };

  configurePrompts(server);

  configureAllTools(server, getAzureDevOpsToken, getAzureDevOpsClient(userAgentComposer), () => userAgentComposer.userAgent);

  const transport = new StdioServerTransport();
  console.error("[ADO MCP] Connecting stdio server...");
  await server.connect(transport);
  console.error("[ADO MCP] Server connected.");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});
