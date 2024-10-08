#requires -Version 7.4
Set-StrictMode -Version 3

function safecount($theobj)
{
    # it shouldn't be this difficult to count an object in strict mode :(
    ($null -eq $theobj) ? 0 : @($theobj).Count
}

$activity = "Getting firewall rules"
Write-Progress -Activity $activity
$rules = Get-NetFirewallRule | ForEach-Object {
    [PSCustomObject]@{
        Rule = $_
        AppFilter = $null
    }
}
Write-Progress -Activity $activity -Completed

$activity = "Locating apps associated with firewall rules"
Write-Progress -Activity $activity
$progress=0
$rules | ForEach-Object {
    Write-Progress -Activity $activity -Status $_.Rule.DisplayName -PercentComplete (100 * $progress / $rules.Count)
    $_.AppFilter = ($_.Rule | Get-NetFirewallApplicationFilter) 
    $progress++
}
Write-Progress -Activity $activity -Completed

# ghost rules are the rules where the app no longer exists:
$ghostrules = $rules | Where-Object { 
    $_.AppFilter.Program -ne $null -and 
    $_.AppFilter.Program -ne 'Any' -and 
    $_.AppFilter.Program -ne 'System' -and 
    !(Test-Path -PathType Leaf -Path ([environment]::ExpandEnvironmentVariables($_.AppFilter.Program)))  
}

$selectedRules = $ghostrules | ForEach-Object { 
    [PSCustomObject]@{
        Name = $_.Rule.DisplayName
        Direction = $_.Rule.Direction
        Action = $_.Rule.Action
        Program = $_.AppFilter.Program
        Rule = $_.Rule
    }
} | Out-GridView -Title "Found $(safecount $ghostrules) ghost program firewall rules, please select which rules to remove" -OutputMode Multiple | ForEach-Object { $_.Rule }

if((safecount $selectedRules) -gt 0)
{
    if((Read-Host -Prompt "Do you really want to delete these $(safecount $selectedRules) rules, type yes to confirm") -eq 'yes')
    {
        $selectedRules | Remove-NetFirewallRule
    }
}
