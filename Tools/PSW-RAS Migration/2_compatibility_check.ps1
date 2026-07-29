

# If the Workspace domains have different AD / LDAP servers specified, it will require a multi-tenant broker implementation in RAS.

$Domains = (Get-Content -Path "data/domains.json" | ConvertFrom-Json).results

# If there are multiple Workspace domains pointing to different AD / LDAP servers, 
# this would need a multi-tenant broker configuration in RAS (unless the domains are trusted).


# Some customers randomized the order of their AD / LDAP servers.
$Domains | ForEach-Object {

    $Servers = $_.server -Split ","
    $Servers = $Servers | Sort-Object
    $_.server = $Servers -Join ","

}

if(($Domains.server | Get-Unique).length -gt 1) {

    throw "The Workspace domains point to different AD / LDAP servers. This migration toolbox does currently not support this scenario."

}

# To-do: Check for context policy labels. Only show a warning that these will not (yet) be migrated.
