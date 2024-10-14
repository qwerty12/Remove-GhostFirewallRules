#requires -Version 7.4
Set-StrictMode -Version 3

function safecount($theobj)
{
    # it shouldn't be this difficult to count an object in strict mode :(
    ($null -eq $theobj) ? 0 : @($theobj).Count
}

Write-Host "Locating ghost application firewall rules..."

$ghostrules = Get-NetFirewallApplicationFilter | Where-Object { 
    $_.Program -ne $null -and
    $_.Program -ne 'Any' -and
    $_.Program -ne 'System' -and
    !(Test-Path -PathType Leaf -Path ([environment]::ExpandEnvironmentVariables($_.Program)))
}

if((safecount $ghostrules) -gt 0)
{
    Write-Host "Found $(safecount $ghostrules) ghost program firewall rules..."

    $rules = Get-NetFirewallRule

    $selectedRules = $ghostrules | Foreach-Object { 
        # find the full firewall rule that matches this firewall application filter
        $matchingRule =  ($rules | Where-Object InstanceID -ieq $_.InstanceID | Select-Object -First 1)
        if($null -ne $matchingRule){
            [PSCustomObject]@{
                Name = $matchingRule.DisplayName
                Direction = $matchingRule.Direction
                Action = $matchingRule.Action
                Program= $_.Program
                Rule = $matchingRule
            }
        }   
    } | Out-GridView -Title "Found $(safecount $ghostrules) ghost program firewall rules, please select which rules to remove" -OutputMode Multiple | ForEach-Object { $_.Rule }

    if((safecount $selectedRules) -gt 0)
    {
        if((Read-Host -Prompt "Do you really want to delete these $(safecount $selectedRules) rules, type yes to confirm") -eq 'yes')
        {
            $selectedRules | Remove-NetFirewallRule
        }
    }
}
else 
{
    Write-Host "no ghost firewall rules located."
}

Write-Host "Ready."
