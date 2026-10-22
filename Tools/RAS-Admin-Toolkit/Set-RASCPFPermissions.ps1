# --- CONFIGURATION ---
# Adjust the path to your Parallels RAS custom provider script directory
$TargetDir = "C:\CPF_Scripts"
# --------------------

# Check if the folder exists, if not, create it
if (-not (Test-Path -Path $TargetDir)) {
    New-Item -Path $TargetDir -ItemType Directory | Out-Null
    Write-Host "Folder created: $TargetDir" -ForegroundColor Cyan
}

# 1. Fetch the current Access Control List (ACL) of the folder
$Acl = Get-Acl -Path $TargetDir

# 2. Disable inheritance and convert inherited rules to explicit rules ($true, $false)
# This prevents permission changes on the C:\ drive from leaking into this folder
$Acl.SetAccessRuleProtection($true, $false)

# Define inheritance and propagation flags so files and subfolders inherit these rights
$InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
$PropagationFlags = [System.Security.AccessControl.PropagationFlags]::None
$AccessType        = [System.Security.AccessControl.AccessControlType]::Allow

# 3. Define the core entities and grant them FullControl
$Accounts = @(
    "NT AUTHORITY\SYSTEM",
    "BUILTIN\Administrators",
    "$env:USERDOMAIN\$env:USERNAME" # The current logged-in user (including domain prefix)
)

# Clean up existing explicit rules for these targets to avoid duplicates
foreach ($Account in $Accounts) {
    $Acl.Access | Where-Object { $_.IdentityReference -eq $Account } | ForEach-Object {
        $Acl.RemoveAccessRule($_) | Out-Null
    }
    
    # Create and apply the FullControl rule
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($Account, "FullControl", $InheritanceFlags, $PropagationFlags, $AccessType)
    $Acl.SetAccessRule($AccessRule)
}

# 4. Remove standard 'Users' and 'Authenticated Users' to secure the custom provider scripts
$Acl.Access | Where-Object { $_.IdentityReference -eq "BUILTIN\Users" -or $_.IdentityReference -eq "NT AUTHORITY\Authenticated Users" } | ForEach-Object {
    $Acl.RemoveAccessRuleAll($_) | Out-Null
}

# 5. Apply the modified ACL back to the directory
Set-Acl -Path $TargetDir -AclObject $Acl

Write-Host "[SUCCESS] Permissions successfully set for SYSTEM, Administrators, and $env:USERDOMAIN\$env:USERNAME!" -ForegroundColor Green
