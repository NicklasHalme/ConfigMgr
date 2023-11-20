# OS Deployment Computer Management
With this tool you can add computers to MDT Database. You can add one by one, or use a CSV file to add many computers at once.
You can also remove computers from AD, ConfigMgr and MDT databases. You can remove one by one, or use a CSV file to add many computers at once.

Requirements:

For AD functions:
  Remote Server Administration Tools (RSAT) for Windows 10 
  or 
  Active Directory module for Windows PowerShell installed on Windows Server

For SCCM functions:
  Right to remove Resource from the ConfigMgr database


For MDT functions:  
  Rights to add and remove data from MDT database.
  MDTDM.psm1 Module from Michael Niehaus, https://techcommunity.microsoft.com/t5/windows-blog-archive/manipulating-the-microsoft-deployment-toolkit-database-using/ba-p/706876
   
