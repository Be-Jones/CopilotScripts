# AgentScripts

Automation assets for administering Microsoft 365 Copilot agents and related agent-management workflows.

## Scripts

| Item | Description |
| --- | --- |
| [BlockAgents](Scripts/BlockAgents/Block-AllCopilotAgents.ps1) | PowerShell 7 script blocks agents that are not in the supplied allowlist via a beta Graph API endpoint. It skips agents that are already blocked and reports blocked, protected, draft/unpublished, and failed agents. |
|[ListAgentsWithAttachments](Scripts/ListAgentsWithAttachments/Get-AgentKnowledgeSources.ps1) | PowerShell 7 script that retrieves agents from the M365 Admin Center inventory and can be filtered to show only Agents that have knowledge sources that were uploaded and attached directly to the agent. This is meant to help identify agents that may be sharing data unknowingly, as the data in those attached files is automatically made available to all users of the agent. Agent 365 license is required. See [Scripts/ListAgentsWithAttachments/README.md](Scripts/ListAgentsWithAttachments/README.md) for details| 

## Cloud Flows
**Note:** Cloud Flows can be imported individually or the entire _/CloudFlows/AgentManagementScripts_version.zip_ file can be imported directly in make.powerautomate.com.
| Item | Description |
| --- | --- |
| [BlockAgentsnotinAllowlist](CloudFlows/modernflows/55397225-786a-f111-a826-000d3a3ba81e/BlockAgentsnotinAllowlist-55397225-786A-F111-A826-000D3A3BA81E.json) | Power Automate cloud flow that blocks Microsoft 365 Copilot agents unless their package IDs are in an allowlist. It uses an HTTP with Entra connection to Microsoft Graph, skips already-blocked agents, and reports blocked, Draft, and failed agents. |

See [Scripts/BlockAgents/README.md](Scripts/BlockAgents/README.md) for script usage details and [CloudFlows/Readme.md](CloudFlows/Readme.md) for cloud flow prerequisites and notes.
