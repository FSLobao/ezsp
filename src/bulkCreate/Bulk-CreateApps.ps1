<#
.SYNOPSIS
    Bulk create Azure AD applications for SharePoint access

.DESCRIPTION
    Reads a JSON file with app configurations, creates Azure AD app registrations
    with Sites.Selected permission, and outputs results with credentials.

.PARAMETER InputPath
    Path to input JSON file with array of app configurations

.PARAMETER OutputPath
    Path to output JSON file (default: input_file-output.json)

.PARAMETER SkipLogin
    Skip authentication (assume already logged in)

.EXAMPLE
    .\Bulk-CreateApps.ps1 -InputPath "apps.json"
    
.EXAMPLE
    .\Bulk-CreateApps.ps1 -InputPath "apps.json" -OutputPath "results.json" -SkipLogin

.NOTES
    Requires Microsoft.Graph PowerShell module (Install-Module Microsoft.Graph)
    Requires Tenant Admin or Privileged Role Administrator permissions
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory = $false)]
    [switch]$SkipLogin
)

# ==================== CONFIGURATION ====================

$GraphScopes = @("Application.ReadWrite.All", "Directory.Read.All", "AppRoleAssignment.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All", "Sites.FullControl.All")
$MicrosoftGraphAppId = "00000003-0000-0000-c000-000000000000"

# ==================== HELPER FUNCTIONS ====================

function Write-Status {
    param([string]$Message, [ValidateSet("Info", "Success", "Warning", "Error")]$Level = "Info")
    $emoji = @{
        "Info"    = "ℹ️ "
        "Success" = "✅ "
        "Warning" = "⚠️  "
        "Error"   = "❌ "
    }
    Write-Host "$($emoji[$Level])$Message" -ForegroundColor @{
        "Info"    = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error"   = "Red"
    }[$Level]
}

function Test-JsonFile {
    param([string]$Path)
    
    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Status "Input file not found: $Path" "Error"
        return $false
    }
    
    try {
        $content = Get-Content -Path $Path -Raw | ConvertFrom-Json
        if ($content -isnot [System.Collections.Generic.List`1[System.Object]] -and -not ($content -is [array])) {
            Write-Status "JSON must be an array of app configurations" "Error"
            return $false
        }
        return $true
    }
    catch {
        Write-Status "Invalid JSON: $_" "Error"
        return $false
    }
}

function Connect-ToGraph {
    Write-Status "Logging in to Microsoft Graph..."
    
    try {
        Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop | Out-Null
        Write-Status "Logged in successfully" "Success"
        return $true
    }
    catch {
        Write-Status "Login failed: $_" "Error"
        return $false
    }
}

function Get-TenantId {
    try {
        $org = Get-MgOrganization -ErrorAction Stop
        return $org.Id
    }
    catch {
        Write-Status "Failed to get tenant ID: $_" "Error"
        return $null
    }
}

function Get-SiteAccessSettings {
    param([hashtable]$Config, [string]$AuthType)

    $siteId = $Config["site_id"]
    if ([string]::IsNullOrWhiteSpace($siteId)) {
        return @{ SiteId = $null; SiteRole = $null; ErrorMessage = "Missing required field: 'site_id' for all applications" }
    }

    $accessType = $Config["access_type"]
    if ([string]::IsNullOrWhiteSpace($accessType)) {
        return @{ SiteId = $null; SiteRole = $null; ErrorMessage = "Missing required field: 'access_type' for all applications" }
    }

    $normalizedAccessType = $accessType.ToString().ToLowerInvariant()
    $siteRole = switch ($normalizedAccessType) {
        "leitura" { "read" }
        "escrita" { "write" }
        "read" { "read" }
        "write" { "write" }
        default { $null }
    }

    if ($null -eq $siteRole) {
        return @{ SiteId = $null; SiteRole = $null; ErrorMessage = "Invalid 'access_type': expected 'leitura' or 'escrita'" }
    }

    return @{ SiteId = $siteId.ToString().Trim(); SiteRole = $siteRole; ErrorMessage = $null }
}

function Get-SitesSelectedPermission {
    param(
        [string]$AuthType,
        [object]$GraphServicePrincipal
    )

    if ($AuthType -eq "delegated") {
        $permission = $GraphServicePrincipal.Oauth2PermissionScopes | Where-Object { $_.Value -eq "Sites.Selected" } | Select-Object -First 1
        return @{ Id = $permission.Id; Type = "Scope" }
    }

    $permission = $GraphServicePrincipal.AppRoles | Where-Object { $_.Value -eq "Sites.Selected" -and $_.AllowedMemberTypes -contains "Application" } | Select-Object -First 1
    return @{ Id = $permission.Id; Type = "Role" }
}

function Get-OrCreateServicePrincipal {
    param([string]$AppId)

    $servicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$AppId'" | Select-Object -First 1
    if ($null -eq $servicePrincipal) {
        $servicePrincipal = New-MgServicePrincipal -AppId $AppId -ErrorAction Stop
    }

    return $servicePrincipal
}

function Grant-GraphAdminConsent {
    param(
        [string]$AuthType,
        [object]$ServicePrincipal,
        [object]$GraphServicePrincipal,
        [hashtable]$Permission
    )

    if ($AuthType -eq "delegated") {
        $existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($ServicePrincipal.Id)' and resourceId eq '$($GraphServicePrincipal.Id)' and consentType eq 'AllPrincipals'" | Select-Object -First 1
        if ($null -eq $existingGrant) {
            New-MgOauth2PermissionGrant -ClientId $ServicePrincipal.Id -ConsentType "AllPrincipals" -ResourceId $GraphServicePrincipal.Id -Scope "Sites.Selected" -ErrorAction Stop | Out-Null
            return
        }

        if ($existingGrant.Scope -notmatch '(^| )Sites\.Selected($| )') {
            $updatedScope = (($existingGrant.Scope, "Sites.Selected") -join " ").Trim()
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id -Scope $updatedScope -ErrorAction Stop | Out-Null
        }
        return
    }

    $existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id | Where-Object {
        $_.ResourceId -eq $GraphServicePrincipal.Id -and $_.AppRoleId -eq $Permission.Id
    } | Select-Object -First 1
    if ($null -eq $existingAssignment) {
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -PrincipalId $ServicePrincipal.Id -AppRoleId $Permission.Id -ResourceId $GraphServicePrincipal.Id -ErrorAction Stop | Out-Null
    }
}

function Grant-SiteAccess {
    param(
        [string]$SiteId,
        [string]$ClientId,
        [string]$DisplayName,
        [string]$SiteRole
    )

    $body = @{
        roles = @($SiteRole)
        grantedToIdentities = @(
            @{
                application = @{
                    id = $ClientId
                    displayName = $DisplayName
                }
            }
        )
    } | ConvertTo-Json -Depth 5

    try {
        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions" -Body $body -ContentType "application/json" | Out-Null
        return $true
    }
    catch {
        Write-Status "Could not grant site access '$SiteRole': $_" "Warning"
        return $false
    }
}

function New-AzureAdApp {
    param(
        [hashtable]$Config
    )
    
    $result = @{}
    $result.PSObject.Properties | ForEach-Object { $result.Remove($_.Name) }
    
    # Copy input to output
    $Config.GetEnumerator() | ForEach-Object { $result[$_.Key] = $_.Value }
    
    $displayName = $Config["display_name"] ?? $Config["name"] ?? "App"
    $signInAudience = $Config["sign_in_audience"] ?? "AzureADMyOrg"
    $authType = ($Config["auth_type"] ?? "app_only").ToString().ToLowerInvariant()
    $redirectUri = $Config["redirect_uri"]

    if ($authType -notin @("app_only", "delegated")) {
        $result["status"] = "error"
        $result["error_message"] = "Invalid 'auth_type': expected 'app_only' or 'delegated'"
        Write-Status "Invalid authentication type for '$displayName': use 'app_only' or 'delegated'" "Error"
        return $result
    }

    if ($authType -eq "delegated" -and [string]::IsNullOrWhiteSpace($redirectUri)) {
        $result["status"] = "error"
        $result["error_message"] = "Missing required field: 'redirect_uri' for delegated applications"
        Write-Status "Missing redirect URI for delegated application '$displayName'" "Error"
        return $result
    }

    $siteAccessSettings = Get-SiteAccessSettings -Config $Config -AuthType $authType
    if ($null -ne $siteAccessSettings["ErrorMessage"]) {
        $result["status"] = "error"
        $result["error_message"] = $siteAccessSettings["ErrorMessage"]
        Write-Status "Invalid site access settings for '$displayName': $($siteAccessSettings["ErrorMessage"])" "Error"
        return $result
    }
    $siteId = $siteAccessSettings["SiteId"]
    $siteRole = $siteAccessSettings["SiteRole"]
    
    Write-Host "  Creating app registration: $displayName"
    
    # Create app
    try {
        $appParams = @{
            DisplayName    = $displayName
            SignInAudience = $signInAudience
            ErrorAction    = "Stop"
        }
        if ($authType -eq "delegated") {
            $appParams["DefaultRedirectUri"] = $redirectUri
            $appParams["Web"] = @{ RedirectUris = @($redirectUri) }
        }

        $app = New-MgApplication @appParams
        
        $result["app_object_id"] = $app.Id
        $result["client_id"] = $app.AppId
        Write-Host "    ✓ App created"
        if ($authType -eq "delegated") {
            Write-Host "    ✓ Redirect URI configured: $redirectUri"
        }
    }
    catch {
        $result["status"] = "error"
        $result["error_message"] = "Failed to create app: $_"
        Write-Status "Failed to create app: $_" "Error"
        return $result
    }
    
    # Create client secret
    $secretDisplayName = $Config["secret_display_name"] ?? "$displayName-secret"
    $secretExpirationDate = $Config["secret_expiration_date"]
    
    # Validate and parse expiration date (format: dd/mm/aaaa)
    if ([string]::IsNullOrEmpty($secretExpirationDate)) {
        $result["status"] = "error"
        $result["error_message"] = "Missing required field: 'secret_expiration_date' (format: dd/mm/aaaa)"
        Write-Status "Missing expiration date for '$displayName'" "Error"
        return $result
    }
    
    try {
        $endDateTime = [DateTime]::ParseExact($secretExpirationDate, "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        $result["status"] = "error"
        $result["error_message"] = "Invalid date format for 'secret_expiration_date': expected dd/mm/aaaa (e.g., 11/05/2028)"
        Write-Status "Invalid date format for '$displayName': expected dd/mm/aaaa" "Error"
        return $result
    }
    
    # Validate maximum secret expiration (730 days from today)
    $maxExpirationDate = (Get-Date).AddDays(730)
    if ($endDateTime -gt $maxExpirationDate) {
        $result["status"] = "error"
        $result["error_message"] = "Secret expiration date cannot exceed 730 days from today. Requested: $(($endDateTime - (Get-Date)).Days) days, max allowed: 730 days"
        Write-Status "Invalid expiration date for '$displayName': max 730 days from today (until $(Get-Date $maxExpirationDate -Format 'dd/MM/yyyy'))" "Error"
        return $result
    }
    
    try {
        $secretParams = @{
            DisplayName = $secretDisplayName
            EndDateTime = $endDateTime
        }
        
        $secret = Add-MgApplicationPassword `
            -ApplicationId $app.Id `
            -PasswordCredential $secretParams `
            -ErrorAction Stop
        
        $result["client_secret"] = $secret.SecretText
        Write-Host "    ✓ Secret created (expires: $secretExpirationDate)"
    }
    catch {
        $result["status"] = "error"
        $result["error_message"] = "Failed to create secret: $_"
        Write-Status "Failed to create secret: $_" "Error"
        return $result
    }

    # Add Sites.Selected permission
    try {
        $graphServicePrincipal = Get-MgServicePrincipal `
            -Filter "AppId eq '$MicrosoftGraphAppId'" `
            -ErrorAction Stop
        $sitesSelectedPermission = Get-SitesSelectedPermission -AuthType $authType -GraphServicePrincipal $graphServicePrincipal
        if ($null -eq $sitesSelectedPermission["Id"]) {
            throw "Sites.Selected permission not found for auth_type '$authType'"
        }
        
        Update-MgApplication `
            -ApplicationId $app.Id `
            -RequiredResourceAccess @(
                @{
                    ResourceAppId  = $MicrosoftGraphAppId
                    ResourceAccess = @(
                        @{
                            Id   = $sitesSelectedPermission["Id"]
                            Type = $sitesSelectedPermission["Type"]
                        }
                    )
                }
            ) `
            -ErrorAction Stop
        
        Write-Host "    ✓ Sites.Selected permission added"
    }
    catch {
        $result["status"] = "warning"
        $result["error_message"] = "App created but permission could not be added: $_"
        Write-Status "Could not add permission: $_" "Warning"
        return $result
    }
    
    # Grant admin consent
    try {
        $spn = Get-OrCreateServicePrincipal -AppId $app.AppId
        if ($null -eq $graphServicePrincipal) {
            $graphServicePrincipal = Get-MgServicePrincipal -Filter "AppId eq '$MicrosoftGraphAppId'" -ErrorAction Stop
        }
        if ($null -eq $sitesSelectedPermission) {
            $sitesSelectedPermission = Get-SitesSelectedPermission -AuthType $authType -GraphServicePrincipal $graphServicePrincipal
        }

        Grant-GraphAdminConsent -AuthType $authType -ServicePrincipal $spn -GraphServicePrincipal $graphServicePrincipal -Permission $sitesSelectedPermission
        
        Write-Host "    ✓ Admin consent granted"
    }
    catch {
        $result["status"] = "warning"
        $result["error_message"] = "Permission added but admin consent could not be granted: $_"
        Write-Status "Could not grant admin consent: $_" "Warning"
        return $result
    }

    if (-not (Grant-SiteAccess -SiteId $siteId -ClientId $app.AppId -DisplayName $displayName -SiteRole $siteRole)) {
        $result["status"] = "warning"
        $result["error_message"] = "Admin consent granted but site access '$siteRole' could not be granted"
        return $result
    }
    Write-Host "    ✓ Site access granted ($siteRole)"
    
    $result["status"] = "success"
    $result["created_at"] = (Get-Date -Format "o")
    if ($authType -eq "delegated") {
        $result["auth_note"] = "Delegated application created with site-restricted Sites.Selected access. Runtime access is restricted to the granted site and intersected with the signed-in user's permissions."
    }
    
    return $result
}

# ==================== MAIN SCRIPT ====================

# Validate input
if (-not (Test-JsonFile -Path $InputPath)) {
    exit 1
}

# Set output path
if ([string]::IsNullOrEmpty($OutputPath)) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $directory = Split-Path -Path $InputPath -Parent
    $OutputPath = Join-Path -Path $directory -ChildPath "$baseName-output.json"
}

# Login
if (-not $SkipLogin) {
    if (-not (Connect-ToGraph)) {
        exit 1
    }
}

Write-Host ""

# Get tenant ID
$tenantId = Get-TenantId
if ($null -eq $tenantId) {
    exit 1
}

# Read input
try {
    $appsConfig = Get-Content -Path $InputPath -Raw | ConvertFrom-Json
    if ($appsConfig -isnot [array]) {
        $appsConfig = @($appsConfig)
    }
}
catch {
    Write-Status "Failed to read JSON: $_" "Error"
    exit 1
}

Write-Status "Creating $($appsConfig.Count) application(s)..." "Info"
Write-Host ""

# Create apps
$results = @()
$successCount = 0

for ($i = 0; $i -lt $appsConfig.Count; $i++) {
    $appConfig = $appsConfig[$i]
    $displayName = $appConfig.display_name ?? $appConfig.name ?? "App"
    
    Write-Host "[$($i + 1)/$($appsConfig.Count)] $displayName"
    
    $result = New-AzureAdApp -Config $appConfig
    $result["tenant_id"] = $tenantId
    
    $results += $result
    
    if ($result.status -eq "success") {
        Write-Status "Created successfully" "Success"
        $successCount++
    }
    elseif ($result.status -eq "warning") {
        Write-Status $result.error_message "Warning"
    }
    else {
        Write-Status $result.error_message "Error"
    }
    
    Write-Host ""
}

# Save results
try {
    $results | ConvertTo-Json -Depth 10 | Set-Content -Path $OutputPath
    Write-Status "Results saved to: $OutputPath" "Success"
}
catch {
    Write-Status "Failed to save results: $_" "Error"
    exit 1
}

# Summary
Write-Host ""
Write-Status "Summary: $successCount/$($appsConfig.Count) apps created successfully" "Info"

if ($successCount -eq $appsConfig.Count) {
    Write-Host ""
    Write-Host "📋 Next steps:"
    Write-Host "  1. Review the output JSON: $OutputPath"
    Write-Host "  2. Configure your .env file with AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID"
    Write-Host "  3. Add apps to SharePoint (requires SharePoint admin)"
    Write-Host "  4. Run examples from examples/ folder"
}
