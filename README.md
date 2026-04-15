# UCLobby Microsoft 365 PowerShell Module
Please use the PowerShell Gallery to install this module:
<br/>
<br/>PowerShell Gallery – UcLobbyTeams
<br/>https://www.powershellgallery.com/packages/UcLobbyMicrosoft365/
<br/>
<br/>Available cmdlets:

<br/>Get-UcM365TenantId 
<br/>Export-UcM365LicenseAssignment

<br/>More info:
<br/>https://uclobby.com/uclobby-microsoft-365-powershell-module/

<br/>Change Log:
<br/>1.1.0 - 2026/04/15
<ul>
  <li>Get-UcM365TenantId
  <br/>Major: Only returns the Tenant ID, previously was returning the OnMicrosoft domains (Multi GEO). This is due a recent Microsoft 365 change.
  </li>
  <li>Get-UcM365Domains
  <br/>Major: Removed cmdlet due recent changes on Microsoft 365.
</li>
</ul>
<br/>1.0.1 - 2025/11/03
<ul>
  <li>Export-UcM365LicenseAssignment
  <br/>Patch: Issue when using EntraAuth to authenticate to Microsoft Graph.
  <br/>Patch: SkipServicePlan was ignored when using DuplicateServicePlansOnly.
  <br/>Patch: Performance improvement by reducing the ammount of writes to file.
  </li>
</ul>
<br/>1.0.0 - 2025/07/23
<ul>
  <li>Major: Initial Release uploaded to PowerShell Gallery</li>
</ul>