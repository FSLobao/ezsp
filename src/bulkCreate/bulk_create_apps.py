#!/usr/bin/env python3
"""
bulk_create_apps.py — Bulk creation of Azure AD applications for Microsoft Graph SharePoint access.

This script reads a JSON file containing a list of application configurations,
creates Azure AD app registrations with required permissions and secrets,
and outputs an updated JSON file with the generated credentials.

Usage:
    python -m msgraphtest.bulk_create_apps input.json [--output output.json] [--method cli|powershell]

Input JSON format:
    [
        {
            "name": "MyApp-SharePoint",
            "display_name": "My SharePoint Application",
            "auth_type": "app_only",
            "sign_in_audience": "AzureADMyOrg",
            "redirect_uri": "http://localhost:8000",
                "site_id": "contoso.sharepoint.com,site-guid,web-guid",
                "access_type": "leitura",
            "secret_display_name": "MyApp-secret",
            "secret_expiration_date": "11/05/2028"
        },
        ...
    ]

Output JSON includes all input fields plus:
    - "client_id": The Azure AD app client ID
    - "tenant_id": The Azure AD tenant ID (added to first app only, shared across all)
    - "client_secret": The client secret (generated, valid for specified years)
    - "created_at": ISO timestamp when the app was created
    - "status": "success" or "error"
    - "error_message": If status is "error"
"""

import argparse
import json
import subprocess
import sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Optional


MICROSOFT_GRAPH_APP_ID = "00000003-0000-0000-c000-000000000000"
SITES_SELECTED_PERMISSION_NAME = "Sites.Selected"
SITES_SELECTED_PERMISSION_KIND = {
    "app_only": "Role",
    "delegated": "Scope",
}


def _get_auth_settings(
    app_config: dict[str, Any],
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    auth_type = str(app_config.get("auth_type", "app_only")).strip().lower()
    if auth_type not in {"app_only", "delegated"}:
        return None, None, "Invalid 'auth_type': expected 'app_only' or 'delegated'"

    redirect_uri = app_config.get("redirect_uri")
    if auth_type == "delegated":
        if not isinstance(redirect_uri, str) or not redirect_uri.strip():
            return (
                None,
                None,
                "Missing required field: 'redirect_uri' for delegated applications",
            )
        redirect_uri = redirect_uri.strip()

    return auth_type, redirect_uri, None


def _get_site_access_settings(
    app_config: dict[str, Any],
) -> tuple[Optional[str], Optional[str], Optional[str]]:
    site_id = app_config.get("site_id")
    if not isinstance(site_id, str) or not site_id.strip():
        return None, None, "Missing required field: 'site_id' for all applications"

    access_type = app_config.get("access_type")
    if not isinstance(access_type, str) or not access_type.strip():
        return (
            None,
            None,
            "Missing required field: 'access_type' for all applications",
        )

    role_map = {
        "leitura": "read",
        "escrita": "write",
        "read": "read",
        "write": "write",
    }
    site_role = role_map.get(access_type.strip().lower())
    if site_role is None:
        return None, None, "Invalid 'access_type': expected 'leitura' or 'escrita'"

    return site_id.strip(), site_role, None


class AppCreator:
    """Base class for creating Azure AD applications."""

    def __init__(self, method: str = "cli"):
        self.method = method
        self.tenant_id: Optional[str] = None

    def login(self) -> bool:
        """Authenticate with Azure. Returns True if successful."""
        raise NotImplementedError

    def create_app(self, app_config: dict[str, Any]) -> dict[str, Any]:
        """Create an Azure AD app registration. Returns config dict with credentials."""
        raise NotImplementedError

    def add_sites_selected_permission(self, app_object_id: str, auth_type: str) -> bool:
        """Add Sites.Selected permission to app. Returns True if successful."""
        raise NotImplementedError

    def grant_admin_consent(self, app_object_id: str, auth_type: str) -> bool:
        """Grant admin consent for the app. Returns True if successful."""
        raise NotImplementedError

    def grant_site_access(
        self,
        site_id: str,
        client_id: str,
        display_name: str,
        site_role: str,
    ) -> bool:
        """Grant site-specific SharePoint access for the app. Returns True if successful."""
        raise NotImplementedError


class AzureCliAppCreator(AppCreator):
    """Creates apps using Azure CLI."""

    def __init__(self, method: str = "cli"):
        super().__init__(method)
        self._sites_selected_permission_ids: dict[str, str] = {}

    def login(self) -> bool:
        """Authenticate with Azure CLI."""
        try:
            result = subprocess.run(
                ["az", "login"],
                capture_output=True,
                text=True,
                timeout=60,
            )
            return result.returncode == 0
        except FileNotFoundError:
            print(
                "❌ Azure CLI not found. Install with: winget install Microsoft.AzureCLI"
            )
            return False
        except subprocess.TimeoutExpired:
            print("❌ Azure CLI login timed out")
            return False

    def _run_command(self, command: list[str], error_msg: str = "") -> tuple[bool, str]:
        """Run az command and return (success, output)."""
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=30,
            )
            if result.returncode == 0:
                return True, result.stdout.strip()
            else:
                error = result.stderr.strip() or result.stdout.strip()
                print(f"  ❌ {error_msg}: {error}")
                return False, error
        except subprocess.TimeoutExpired:
            print(f"  ❌ {error_msg}: Command timed out")
            return False, "timeout"

    def _get_sites_selected_permission_id(self, auth_type: str) -> Optional[str]:
        permission_id = self._sites_selected_permission_ids.get(auth_type)
        if permission_id:
            return permission_id

        if auth_type == "delegated":
            query = "oauth2PermissionScopes[?value=='Sites.Selected'] | [0].id"
        else:
            query = "appRoles[?value=='Sites.Selected' && contains(allowedMemberTypes, 'Application')] | [0].id"

        success, output = self._run_command(
            [
                "az",
                "ad",
                "sp",
                "show",
                "--id",
                MICROSOFT_GRAPH_APP_ID,
                "--query",
                query,
                "-o",
                "tsv",
            ],
            f"Failed to resolve {SITES_SELECTED_PERMISSION_NAME} permission ID for '{auth_type}'",
        )
        if not success or not output:
            return None

        permission_id = output.strip()
        self._sites_selected_permission_ids[auth_type] = permission_id
        return permission_id

    def create_app(self, app_config: dict[str, Any]) -> dict[str, Any]:
        """Create Azure AD app via Azure CLI."""
        result = app_config.copy()

        # Get tenant ID if not cached
        if not self.tenant_id:
            success, tenant = self._run_command(
                ["az", "account", "show", "--query", "tenantId", "-o", "tsv"],
                "Failed to get tenant ID",
            )
            if not success:
                result["status"] = "error"
                result["error_message"] = tenant
                return result
            self.tenant_id = tenant
        result["tenant_id"] = self.tenant_id

        # Create app registration
        display_name = app_config.get("display_name", app_config.get("name", "App"))
        sign_in_audience = app_config.get("sign_in_audience", "AzureADMyOrg")
        auth_type, redirect_uri, auth_error = _get_auth_settings(app_config)
        if auth_error:
            result["status"] = "error"
            result["error_message"] = auth_error
            print(
                f"❌ Invalid authentication settings for '{display_name}': {auth_error}"
            )
            return result
        assert auth_type is not None
        site_id, site_role, site_error = _get_site_access_settings(app_config)
        if site_error:
            result["status"] = "error"
            result["error_message"] = site_error
            print(f"❌ Invalid site access settings for '{display_name}': {site_error}")
            return result

        create_command = [
            "az",
            "ad",
            "app",
            "create",
            "--display-name",
            display_name,
            "--sign-in-audience",
            sign_in_audience,
        ]
        if auth_type == "delegated" and redirect_uri:
            create_command.extend(["--web-redirect-uris", redirect_uri])
        create_command.extend(["--query", "id", "-o", "tsv"])

        success, app_obj_id = self._run_command(
            create_command,
            f"Failed to create app '{display_name}'",
        )
        if not success:
            result["status"] = "error"
            result["error_message"] = f"App creation failed: {app_obj_id}"
            return result

        result["app_object_id"] = app_obj_id

        # Get client ID
        success, client_id = self._run_command(
            [
                "az",
                "ad",
                "app",
                "show",
                "--id",
                app_obj_id,
                "--query",
                "appId",
                "-o",
                "tsv",
            ],
            f"Failed to get client ID for app '{display_name}'",
        )
        if not success:
            result["status"] = "error"
            result["error_message"] = f"Failed to get client ID: {client_id}"
            return result

        result["client_id"] = client_id

        # Create client secret
        secret_display_name = app_config.get(
            "secret_display_name", f"{display_name}-secret"
        )
        secret_expiration_date = app_config.get("secret_expiration_date")

        # Validate and parse expiration date (format: dd/mm/aaaa)
        if not secret_expiration_date:
            result["status"] = "error"
            result["error_message"] = (
                "Missing required field: 'secret_expiration_date' (format: dd/mm/aaaa)"
            )
            print(f"❌ Missing expiration date for '{display_name}'")
            return result

        try:
            end_date_time = datetime.strptime(secret_expiration_date, "%d/%m/%Y")
        except ValueError:
            result["status"] = "error"
            result["error_message"] = (
                "Invalid date format for 'secret_expiration_date': expected dd/mm/aaaa (e.g., 11/05/2028)"
            )
            print(f"❌ Invalid date format for '{display_name}': expected dd/mm/aaaa")
            return result

        # Validate maximum secret expiration (730 days from today)
        max_expiration_date = datetime.now() + timedelta(days=730)
        if end_date_time > max_expiration_date:
            days_diff = (end_date_time - datetime.now()).days
            result["status"] = "error"
            result["error_message"] = (
                f"Secret expiration date cannot exceed 730 days from today. Requested: {days_diff} days, max allowed: 730 days"
            )
            print(
                f"❌ Invalid expiration date for '{display_name}': max 730 days from today"
            )
            return result

        secret_end_date = end_date_time.strftime("%Y-%m-%dT23:59:59")
        success, secret = self._run_command(
            [
                "az",
                "ad",
                "app",
                "credential",
                "reset",
                "--id",
                app_obj_id,
                "--display-name",
                secret_display_name,
                "--end-date",
                secret_end_date,
                "--query",
                "password",
                "-o",
                "tsv",
            ],
            f"Failed to create secret for app '{display_name}'",
        )
        if not success:
            result["status"] = "error"
            result["error_message"] = f"Secret creation failed: {secret}"
            return result

        result["client_secret"] = secret

        # Add Sites.Selected permission
        success = self.add_sites_selected_permission(app_obj_id, auth_type)
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                "App created but Sites.Selected permission could not be added"
            )
            return result

        # Grant admin consent
        success = self.grant_admin_consent(app_obj_id, auth_type)
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                "Sites.Selected permission added but admin consent could not be granted"
            )
            return result

        assert site_id is not None and site_role is not None
        success = self.grant_site_access(
            site_id=site_id,
            client_id=client_id,
            display_name=display_name,
            site_role=site_role,
        )
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                f"Admin consent granted but site access '{site_role}' could not be granted"
            )
            return result

        result["status"] = "success"
        result["created_at"] = datetime.now(timezone.utc).isoformat()
        if auth_type == "delegated":
            result["auth_note"] = (
                "Delegated application created with site-restricted Sites.Selected access. Runtime access is restricted to the granted site and intersected with the signed-in user's permissions."
            )
        return result

    def add_sites_selected_permission(self, app_object_id: str, auth_type: str) -> bool:
        """Add Sites.Selected permission."""
        permission_id = self._get_sites_selected_permission_id(auth_type)
        if not permission_id:
            return False

        success, _ = self._run_command(
            [
                "az",
                "ad",
                "app",
                "permission",
                "add",
                "--id",
                app_object_id,
                "--api",
                MICROSOFT_GRAPH_APP_ID,
                "--api-permissions",
                f"{permission_id}={SITES_SELECTED_PERMISSION_KIND[auth_type]}",
            ],
            "Failed to add Sites.Selected permission",
        )
        return success

    def grant_admin_consent(self, app_object_id: str, auth_type: str) -> bool:
        """Grant admin consent."""
        success, _ = self._run_command(
            [
                "az",
                "ad",
                "app",
                "permission",
                "admin-consent",
                "--id",
                app_object_id,
            ],
            f"Failed to grant admin consent for {auth_type} permissions",
        )
        return success

    def grant_site_access(
        self,
        site_id: str,
        client_id: str,
        display_name: str,
        site_role: str,
    ) -> bool:
        """Grant site-specific SharePoint access using Microsoft Graph."""
        body = json.dumps(
            {
                "roles": [site_role],
                "grantedToIdentities": [
                    {
                        "application": {
                            "id": client_id,
                            "displayName": display_name,
                        }
                    }
                ],
            }
        )
        success, _ = self._run_command(
            [
                "az",
                "rest",
                "--method",
                "POST",
                "--uri",
                f"https://graph.microsoft.com/v1.0/sites/{site_id}/permissions",
                "--headers",
                "Content-Type=application/json",
                "--body",
                body,
            ],
            f"Failed to grant site access '{site_role}' for app '{display_name}'",
        )
        return success


class PowerShellAppCreator(AppCreator):
    """Creates apps using Microsoft Graph PowerShell SDK."""

    def login(self) -> bool:
        """Authenticate with Microsoft Graph PowerShell."""
        try:
            script = """
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.Read.All", "AppRoleAssignment.ReadWrite.All", "DelegatedPermissionGrant.ReadWrite.All", "Sites.FullControl.All" -NoWelcome
$null
"""
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=60,
            )
            if "Error" in result.stderr:
                print(f"❌ PowerShell login failed: {result.stderr}")
                return False
            return True
        except FileNotFoundError:
            print("❌ PowerShell not found")
            return False
        except subprocess.TimeoutExpired:
            print("❌ PowerShell login timed out")
            return False

    def create_app(self, app_config: dict[str, Any]) -> dict[str, Any]:
        """Create Azure AD app via PowerShell."""
        result = app_config.copy()

        # Get tenant ID if not cached
        if not self.tenant_id:
            script = "Get-MgOrganization | Select-Object -ExpandProperty Id"
            try:
                proc = subprocess.run(
                    ["powershell", "-NoProfile", "-Command", script],
                    capture_output=True,
                    text=True,
                    timeout=30,
                )
                if proc.returncode == 0:
                    self.tenant_id = proc.stdout.strip()
                else:
                    result["status"] = "error"
                    result["error_message"] = "Failed to get tenant ID"
                    return result
            except Exception as e:
                result["status"] = "error"
                result["error_message"] = f"Failed to get tenant ID: {str(e)}"
                return result

        result["tenant_id"] = self.tenant_id

        # Create app registration
        display_name = app_config.get("display_name", app_config.get("name", "App"))
        sign_in_audience = app_config.get("sign_in_audience", "AzureADMyOrg")
        auth_type, redirect_uri, auth_error = _get_auth_settings(app_config)
        if auth_error:
            result["status"] = "error"
            result["error_message"] = auth_error
            print(
                f"❌ Invalid authentication settings for '{display_name}': {auth_error}"
            )
            return result
        assert auth_type is not None
        site_id, site_role, site_error = _get_site_access_settings(app_config)
        if site_error:
            result["status"] = "error"
            result["error_message"] = site_error
            print(f"❌ Invalid site access settings for '{display_name}': {site_error}")
            return result

        if auth_type == "delegated" and redirect_uri:
            script = f"""
$app = New-MgApplication -DisplayName "{display_name}" -SignInAudience "{sign_in_audience}" -DefaultRedirectUri "{redirect_uri}" -Web @{{ RedirectUris = @("{redirect_uri}") }}
$app.Id
"""
        else:
            script = f"""
$app = New-MgApplication -DisplayName "{display_name}" -SignInAudience "{sign_in_audience}"
$app.Id
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if proc.returncode != 0:
                result["status"] = "error"
                result["error_message"] = f"Failed to create app: {proc.stderr}"
                return result
            app_object_id = proc.stdout.strip()
            result["app_object_id"] = app_object_id
        except Exception as e:
            result["status"] = "error"
            result["error_message"] = f"Failed to create app: {str(e)}"
            return result

        # Get client ID
        script = f"""
Get-MgApplication -ApplicationId "{app_object_id}" | Select-Object -ExpandProperty AppId
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if proc.returncode != 0:
                result["status"] = "error"
                result["error_message"] = f"Failed to get client ID: {proc.stderr}"
                return result
            client_id = proc.stdout.strip()
            result["client_id"] = client_id
        except Exception as e:
            result["status"] = "error"
            result["error_message"] = f"Failed to get client ID: {str(e)}"
            return result

        # Create client secret
        secret_display_name = app_config.get(
            "secret_display_name", f"{display_name}-secret"
        )
        secret_expiration_date = app_config.get("secret_expiration_date")

        # Validate and parse expiration date (format: dd/mm/aaaa)
        if not secret_expiration_date:
            result["status"] = "error"
            result["error_message"] = (
                "Missing required field: 'secret_expiration_date' (format: dd/mm/aaaa)"
            )
            print(f"❌ Missing expiration date for '{display_name}'")
            return result

        try:
            end_date_time = datetime.strptime(secret_expiration_date, "%d/%m/%Y")
        except ValueError:
            result["status"] = "error"
            result["error_message"] = (
                "Invalid date format for 'secret_expiration_date': expected dd/mm/aaaa (e.g., 11/05/2028)"
            )
            print(f"❌ Invalid date format for '{display_name}': expected dd/mm/aaaa")
            return result

        # Validate maximum secret expiration (730 days from today)
        max_expiration_date = datetime.now() + timedelta(days=730)
        if end_date_time > max_expiration_date:
            days_diff = (end_date_time - datetime.now()).days
            result["status"] = "error"
            result["error_message"] = (
                f"Secret expiration date cannot exceed 730 days from today. Requested: {days_diff} days, max allowed: 730 days"
            )
            print(
                f"❌ Invalid expiration date for '{display_name}': max 730 days from today"
            )
            return result

        script = f"""
$secret = Add-MgApplicationPassword -ApplicationId "{app_object_id}" -PasswordCredential @{{
    DisplayName = "{secret_display_name}"
    EndDateTime = [DateTime]::ParseExact("{secret_expiration_date}", "dd/MM/yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
}}
$secret.SecretText
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            if proc.returncode != 0:
                result["status"] = "error"
                result["error_message"] = f"Failed to create secret: {proc.stderr}"
                return result
            client_secret = proc.stdout.strip()
            result["client_secret"] = client_secret
        except Exception as e:
            result["status"] = "error"
            result["error_message"] = f"Failed to create secret: {str(e)}"
            return result

        # Add Sites.Selected permission
        success = self.add_sites_selected_permission(app_object_id, auth_type)
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                "App created but Sites.Selected permission could not be added"
            )
            return result

        # Grant admin consent
        success = self.grant_admin_consent(app_object_id, auth_type)
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                "Sites.Selected permission added but admin consent could not be granted"
            )
            return result

        assert site_id is not None and site_role is not None
        success = self.grant_site_access(
            site_id=site_id,
            client_id=client_id,
            display_name=display_name,
            site_role=site_role,
        )
        if not success:
            result["status"] = "warning"
            result["error_message"] = (
                f"Admin consent granted but site access '{site_role}' could not be granted"
            )
            return result

        result["status"] = "success"
        result["created_at"] = datetime.now(timezone.utc).isoformat()
        if auth_type == "delegated":
            result["auth_note"] = (
                "Delegated application created with site-restricted Sites.Selected access. Runtime access is restricted to the granted site and intersected with the signed-in user's permissions."
            )
        return result

    def add_sites_selected_permission(self, app_object_id: str, auth_type: str) -> bool:
        """Add Sites.Selected permission via PowerShell."""
        script = f"""
$app = Get-MgApplication -ApplicationId "{app_object_id}"
$graphSP = Get-MgServicePrincipal -Filter "AppId eq '{MICROSOFT_GRAPH_APP_ID}'"
$permission = if ("{auth_type}" -eq "delegated") {{
    $graphSP.Oauth2PermissionScopes | Where-Object {{ $_.Value -eq "{SITES_SELECTED_PERMISSION_NAME}" }} | Select-Object -First 1
}} else {{
    $graphSP.AppRoles | Where-Object {{ $_.Value -eq "{SITES_SELECTED_PERMISSION_NAME}" -and $_.AllowedMemberTypes -contains "Application" }} | Select-Object -First 1
}}
$permissionType = if ("{auth_type}" -eq "delegated") {{ "Scope" }} else {{ "Role" }}

Update-MgApplication -ApplicationId "{app_object_id}" `
    -RequiredResourceAccess @(
        @{{
            ResourceAppId = "{MICROSOFT_GRAPH_APP_ID}"
            ResourceAccess = @(
                @{{ Id = $permission.Id; Type = $permissionType }}
            )
        }}
    )
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            return proc.returncode == 0
        except Exception:
            return False

    def grant_admin_consent(self, app_object_id: str, auth_type: str) -> bool:
        """Grant admin consent via PowerShell."""
        script = f"""
$app = Get-MgApplication -ApplicationId "{app_object_id}"
$graphSP = Get-MgServicePrincipal -Filter "AppId eq '{MICROSOFT_GRAPH_APP_ID}'"
$spn = Get-MgServicePrincipal -Filter "AppId eq '$($app.AppId)'" | Select-Object -First 1
if (-not $spn) {{
    $spn = New-MgServicePrincipal -AppId $app.AppId
}}

if ("{auth_type}" -eq "delegated") {{
    $existingGrant = Get-MgOauth2PermissionGrant -Filter "clientId eq '$($spn.Id)' and resourceId eq '$($graphSP.Id)' and consentType eq 'AllPrincipals'" | Select-Object -First 1
    if ($existingGrant) {{
        if ($existingGrant.Scope -notmatch '(^| )Sites\\.Selected($| )') {{
            $updatedScope = (($existingGrant.Scope, "Sites.Selected") -join " ").Trim()
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $existingGrant.Id -Scope $updatedScope | Out-Null
        }}
    }} else {{
        New-MgOauth2PermissionGrant -ClientId $spn.Id -ConsentType "AllPrincipals" -ResourceId $graphSP.Id -Scope "Sites.Selected" | Out-Null
    }}
}} else {{
    $sitesSelectedRole = $graphSP.AppRoles | Where-Object {{ $_.Value -eq "{SITES_SELECTED_PERMISSION_NAME}" -and $_.AllowedMemberTypes -contains "Application" }} | Select-Object -First 1
    $existingAssignment = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spn.Id | Where-Object {{ $_.ResourceId -eq $graphSP.Id -and $_.AppRoleId -eq $sitesSelectedRole.Id }} | Select-Object -First 1
    if (-not $existingAssignment) {{
        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spn.Id -PrincipalId $spn.Id -AppRoleId $sitesSelectedRole.Id -ResourceId $graphSP.Id | Out-Null
    }}
}}
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            return proc.returncode == 0
        except Exception:
            return False

    def grant_site_access(
        self,
        site_id: str,
        client_id: str,
        display_name: str,
        site_role: str,
    ) -> bool:
        """Grant site-specific SharePoint access via Graph PowerShell."""
        body = json.dumps(
            {
                "roles": [site_role],
                "grantedToIdentities": [
                    {
                        "application": {
                            "id": client_id,
                            "displayName": display_name,
                        }
                    }
                ],
            }
        )
        script = f"""
$body = @'
{body}
'@
Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/{site_id}/permissions" -Body $body -ContentType "application/json"
"""
        try:
            proc = subprocess.run(
                ["powershell", "-NoProfile", "-Command", script],
                capture_output=True,
                text=True,
                timeout=30,
            )
            return proc.returncode == 0
        except Exception:
            return False


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description="Bulk create Azure AD applications for SharePoint access",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m msgraphtest.bulk_create_apps apps.json
  python -m msgraphtest.bulk_create_apps apps.json --output apps-created.json --method powershell
        """,
    )
    parser.add_argument("input", help="Input JSON file with app configurations")
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="Output JSON file (default: input_file-output.json)",
    )
    parser.add_argument(
        "--method",
        "-m",
        choices=["cli", "powershell"],
        default="cli",
        help="Azure authentication method (default: cli)",
    )
    parser.add_argument(
        "--skip-login",
        action="store_true",
        help="Skip login (assume already authenticated)",
    )

    args = parser.parse_args()

    # Read input JSON
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"❌ Input file not found: {input_path}")
        sys.exit(1)

    try:
        with open(input_path) as f:
            apps_config = json.load(f)
    except json.JSONDecodeError as e:
        print(f"❌ Invalid JSON in input file: {e}")
        sys.exit(1)

    if not isinstance(apps_config, list):
        print("❌ Input JSON must be an array of app configurations")
        sys.exit(1)

    # Create app creator
    creator = AzureCliAppCreator() if args.method == "cli" else PowerShellAppCreator()

    # Login
    if not args.skip_login:
        print(f"🔐 Logging in to Azure ({args.method})...")
        if not creator.login():
            sys.exit(1)
        print("✅ Logged in\n")

    # Create apps
    print(f"📱 Creating {len(apps_config)} application(s)...\n")
    results = []

    for i, app_config in enumerate(apps_config, 1):
        app_name = app_config.get("display_name", app_config.get("name", f"App #{i}"))
        print(f"[{i}/{len(apps_config)}] Creating: {app_name}")

        result = creator.create_app(app_config)
        results.append(result)

        if result.get("status") == "success":
            print("  ✅ Created successfully")
            print(f"     Client ID: {result.get('client_id')}")
            if i == 1:
                print(f"     Tenant ID: {result.get('tenant_id')}")
        elif result.get("status") == "warning":
            print(f"  ⚠️  Created with warnings: {result.get('error_message')}")
        else:
            print(f"  ❌ Failed: {result.get('error_message')}")
        print()

    # Write output JSON
    output_path = args.output or str(
        input_path.parent / f"{input_path.stem}-output.json"
    )
    try:
        with open(output_path, "w") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"✅ Results saved to: {output_path}")
    except Exception as e:
        print(f"❌ Failed to write output file: {e}")
        sys.exit(1)

    # Summary
    success_count = sum(1 for r in results if r.get("status") == "success")
    print(f"\n📊 Summary: {success_count}/{len(results)} apps created successfully")


if __name__ == "__main__":
    main()
