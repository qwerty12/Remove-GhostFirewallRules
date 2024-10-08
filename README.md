# Remove-GhostFirewallRules
Interactive powershell script to remove firewall rules for programs that are no longer installed.

When you have a windows system that survived several years of upgrades and usage, you will probably have some useless windows firewall rules for programs that are no longer present on the system (e.g. I found that Steam is particularly messy about removing firewall rules for uninstalled games).

If you run this script (powershell 7.3+) it will find those ghost firewall rules and present a dialog where you can multi-select the rules to delete. For deletion an elevated powershell is required, but if you just want to have a look a non-elevated shell will suffice.





