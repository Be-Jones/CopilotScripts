# AgentScripts

Automation assets for administering Microsoft 365 Copilot agents and related agent-management workflows.

## Scripts

| Item | Description |
| --- | --- |
| [BlockAgents](Scripts/BlockAgents/Block-AllCopilotAgents.ps1) | PowerShell 7 script blocks agents that are not in the supplied allowlist via a beta Graph API endpoint. It skips agents that are already blocked and reports blocked, protected, draft/unpublished, and failed agents. |

## Cloud Flows
**Note:** Cloud Flows can be imported individually or the entire solution/zip file can be imported.
| Item | Description |
| --- | --- |
| [BlockAgentsnotinAllowlist](CloudFlows/modernflows/55397225-786a-f111-a826-000d3a3ba81e/BlockAgentsnotinAllowlist-55397225-786A-F111-A826-000D3A3BA81E.json) | Power Automate cloud flow that blocks Microsoft 365 Copilot agents unless their package IDs are in an allowlist. It uses an HTTP with Entra connection to Microsoft Graph, skips already-blocked agents, and reports blocked, Draft, and failed agents. |

See [Scripts/BlockAgents/README.md](Scripts/BlockAgents/README.md) for script usage details and [CloudFlows/Readme.md](CloudFlows/Readme.md) for cloud flow prerequisites and notes.
