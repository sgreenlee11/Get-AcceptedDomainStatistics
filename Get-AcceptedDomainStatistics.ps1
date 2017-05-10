<# 
.SYNOPSIS 
   Function to report on the usage of accepted domains in an Exchange environment. 
.DESCRIPTION 
   This function provides the ability to report on the usage of accepted domains across all Exchange recipient types. The script provides a count of addresses per 
   accepted domain. By default the script will only look at the primary addresses for recipients. This function can be run from the Exchange Management Shell, or via a
   Remote Powershell session, as long as the remote session is imported with Import-PSSession
   
 
   To execute the script, you will have to dot-source it first after which you can call the function: "Get-AcceptedDomainStatistics" 
 
.Parameter Domain 
   Generates report based on proxy addresses instead of primary addresses only. An object with multiple proxy addresses in a single domain will increment the count for that 
   domain for each address. This Paramater is mandatory 
 
.Parameter IncludeProxy 
   Generates report based on proxy addresses instead of primary addresses only. An object with multiple proxy addresses in a single domain will increment the count for that 
   domain for each address 
 
.Parameter MBXOnly 
   Generates a report for mailbox recipients only. 
 
.EXAMPLE 
   By default the function will only look at the primary SMTP addresses for recipients. DomainName is the only required parameter. To retrieve the number of objects 
   with a primary address in a given domain the below command can be used: 
 
   Get-AcceptedDomainStatistics -Domain Domain.com 
 
.EXAMPLE 
    This function now supports pipeline input as well. Input can be the Get-AcceptedDomain cmdlet (This may not work in Exchange 2007/2010 using remote Powershell), or 
    source. Pipeline info also may not work if the accepted domain name and name do not match. The third example shows a workaround for this (This workaround will take 
    longer to run, and will consume more memory, it is not recommended for environments with over 10k recipients): 
 
   Get-AcceptedDomain | Get-AcceptedDomainStatistics 
   Get-Content Domains.txt | Get-AcceptedDomainStatistics 
   Get-AcceptedDomain | Select-Object -ExpandProperty DomainName | select Domain | Get-AcceptedDomainStatistics 
 
.EXAMPLE 
   By default the function will only look at the primary SMTP addresses for recipients. The following command will give you the number of primary SMTP addresses per 
   Accepted Domain for all recipient types and output it to GridView: 
 
   Get-AcceptedDomainStatistics -Domain Domain.com | Out-GridView
.EXAMPLE 
   There is also an option to look at all proxy addresses as well. In this mode if a user has multiple proxy addresses in a given domain, the count will incremente for  
   each address. The following command can be used to look at the proxy addresses for all recipient types and output to a CSV file: 
 
   Get-AcceptedDomainStatistics -IncludeProxy | Export-CSV AcceptedDomains.csv -NoTypeInformation 
 
#> 
function Get-AcceptedDomainStatistics { 
    [CmdletBinding()]
    Param 
    ( 
        [Parameter(Mandatory = $true,  
            ValueFromPipeline = $true, 
            ValueFromPipelineByPropertyName = $true)] 
        [string[]] 
        $DomainName, 
        [Parameter(Mandatory = $false)] 
        [switch] 
        $IncludeProxy, 
        [Parameter(Mandatory = $false)] 
        [Switch] 
        $MBXOnly 
    ) 
    #Using Begin,Process,End to support pipeline. Structure is a bit weird since the pipeline input is just the Domain Name, but achieves desired results
    Begin {
        $DomainCount = New-Object System.Collections.ArrayList

    }
    Process {
        #Gather recipient information 
        foreach ($d in $DomainName) { 
            $domobj = New-Object PSObject 
            $domobj | Add-Member NoteProperty -Name Domain -Value $d 
            $domobj | Add-Member NoteProperty -Name Count -Value 0
            [void]$DomainCount.Add($domobj)
        }
    }
    # Loop through recipients and count total primay addresses in each domain
    End {
        #Gather recipient information
        if ($MBXOnly) { 
            $ExRecipients = Get-Mailbox -Resultsize Unlimited | Select-Object emailaddresses, primarysmtpaddress 
        } 
        else { 
            $ExRecipients = Get-Recipient -ResultSize Unlimited | Select-Object emailaddresses, primarysmtpaddress 
        }   
        Foreach ($r in $ExRecipients) { 
            if ($IncludeProxy) { 
                Foreach ($a in ($r.EmailAddresses | Where-Object {$_ -like "smtp*"})) { 
                    $adddomain = $a -split "@"
                    ($DomainCount | Where-Object {$_.Domain -eq $adddomain[1]}).Count ++
                } 
            }  
            else { 
                $adddomain = $r.PrimarySmtpAddress -Split "@"
                if ($DomainCount.Domain -contains $adddomain[1]) {
                    ($DomainCount | Where-Object {$_.Domain -eq $adddomain[1]}).Count ++
                }
 
            } 
        } 

        return $DomainCount
    }
}