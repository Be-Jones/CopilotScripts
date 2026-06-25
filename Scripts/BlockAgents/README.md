# CopilotScripts
Automation scripts for the administration of Copilot features and skills

## Copilot Agent Blocking Workflow

1. From https://admin.microsoft365.com/#/agents/all, click the Export option to open a .csv file listing all agents present. From that list, build your allowedAgents.json file following the structure in the file. Use the Title ID field for the id field in the allowedAgents.json

2. Run the blocking script with the reviewed allowlist:

	```powershell
	.\Scripts\BlockAgents\Block-AllCopilotAgents.ps1 -AllowList .\Scripts\BlockAgents\allowedAgents.json
	```

	By default, the script excludes agents where `platform` is `Microsoft 365 Copilot Agent Builder`.
	To include Microsoft 365 Copilot Agent Builder agents in block attempts, add the `-BlockAgentBuilderAgents` switch:

	```powershell
	.\Scripts\BlockAgents\Block-AllCopilotAgents.ps1 -AllowList .\Scripts\BlockAgents\allowedAgents.json -BlockAgentBuilderAgents
	```
