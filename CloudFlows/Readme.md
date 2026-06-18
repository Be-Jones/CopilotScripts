# Cloud Flows

## Block Agents not in Allowlist

`BlockAgentsnotinAllowlist` is a manually triggered cloud flow that blocks Microsoft 365 Copilot agents unless their package IDs are included in an allowlist. It uses the Microsoft Graph Copilot Package Management API through the Power Automate Web Contents connector.

The flow gets Copilot agents from Microsoft Graph, removes allowlisted and already-blocked agents from the target set, attempts to block the remaining agents, and reports which agents were blocked, which failed because they are in Draft status, and which failed for another reason.

### Prerequisites

- The HTTP with Entra connector connection must be created by an AI Administrator or Microsoft 365 Administrator.
- The connection must be created with https://graph.microsoft.com and the URL and resource URI.
- The tenant must have the Microsoft Agent 365 licensing required for the Copilot Package Management API.
- The HTTP with Entra connector must be configured following the documenation here: https://learn.microsoft.com/en-us/connectors/webcontentsv2/#authorize-the-connector-to-act-on-behalf-of-a-signed-in-user with access to `CopilotPackages.ReadWrite.All` 

### Important Notes

- The flow uses Microsoft Graph beta endpoints. Beta API behavior can change.
- The current list call reads the response from the first Graph request. If Graph returns an `@odata.nextLink`, pagination handling should be added so every page of agents is processed.
- The flow skips already-blocked agents before attempting POST calls, so the blocked report only represents agents blocked during this run.
- Draft detection is performed after a failed block attempt by calling the package detail endpoint and checking `elementDetails` for `PublishedStatus":"Draft`.