# CopilotScripts
Automation scripts for the administration of Copilot features and skills

## Copilot Agent Blocking Workflow

1. Generate an allowlist-ready JSON file for review:

	```powershell
	.\Powershell_Scripts\Get-ListAllAgents.ps1 -OutputPath .\Powershell_Scripts\allowedAgents.generated.json
	```

2. Review the generated JSON and keep only the agents that must not be blocked.

3. Run the blocking script with the reviewed allowlist:

	```powershell
	.\Powershell_Scripts\Block-AllCopilotAgents.ps1 -AllowList .\Powershell_Scripts\allowedAgents.json
	```
