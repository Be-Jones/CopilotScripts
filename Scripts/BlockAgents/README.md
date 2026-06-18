# CopilotScripts
Automation scripts for the administration of Copilot features and skills

## Copilot Agent Blocking Workflow

1. From https://admin.microsoft365.com/#/agents/all, click the Export option to open a .csv file listing all agents present. From that list, build your allowedAgents.json file following the structure in the file. Use the Title ID field for the id field in the allowedAgents.json

2. Run the blocking script with the reviewed allowlist:

	```powershell
	.\Powershell_Scripts\Block-AllCopilotAgents.ps1 -AllowList .\Powershell_Scripts\allowedAgents.json
	```
