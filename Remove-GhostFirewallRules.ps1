#requires -RunAsAdministrator
Set-StrictMode -Version 3

function safecount($theobj)
{
    # it shouldn't be this difficult to count an object in strict mode :(
    if ($null -eq $theobj) {
        return 0
    }
    return @($theobj).Count
}

Write-Host "Locating ghost application firewall rules..."

$ghostrules = Get-NetFirewallApplicationFilter | Where-Object { 
    $_.Program -ne $null -and
    $_.Program -ne 'Any' -and
    $_.Program -ne 'System' -and
    !(Test-Path -PathType Leaf -Path ([environment]::ExpandEnvironmentVariables($_.Program)))
}

$count = safecount $ghostrules
if($count -gt 0)
{
    $rules = Get-NetFirewallRule
    $selectedRules = [PSCustomObject[]]::new($count)

    $index = 0
    foreach ($ghost in $ghostrules) {
        # find the full firewall rule that matches this firewall application filter
        $matchingRule = $rules | Where-Object InstanceID -ieq $ghost.InstanceID | Select-Object -First 1
        if($null -ne $matchingRule){
            if ($null -ne $matchingRule.Group -and $matchingRule.Group.StartsWith("@FirewallAPI.dll,")) {
                continue
            }

            $selectedRules[$index++] = [PSCustomObject]@{
                Name      = $matchingRule.DisplayName
                Direction = $matchingRule.Direction
                Action    = $matchingRule.Action
                Program   = $ghost.Program
                Rule      = $matchingRule
            }
        }
    }

    if($index -gt 0)
    {
        $msg = "Found $index ghost program firewall rules"
        Write-Host "$msg for possible removal"
        $selectedRules = $selectedRules | Out-GridView -Title "$msg, please select which to remove" -OutputMode Multiple | ForEach-Object { $_.Rule }

        if((safecount $selectedRules) -gt 0)
        {
            $selectedRules | ForEach-Object {
                Write-Host "Removing firewall rule: $($_.DisplayName)"
                $_ | Remove-NetFirewallRule
            }
        }
        return
    }
}

Write-Host "no ghost firewall rules located."
