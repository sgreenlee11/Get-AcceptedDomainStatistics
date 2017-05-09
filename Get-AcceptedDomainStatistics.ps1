<# 
.SYNOPSIS 
   Function to report on the usage of accepted domains in an Exchange environment. 
.DESCRIPTION 
   This function provides the ability to report on the usage of accepted domains across all Exchange recipient types. The script provides a count of addresses per 
   accepted domain. By default the script will only look at the primary addresses for recipients. This function can be run from the Exchange Management Shell, or a standar 
   PowerShell window after adding the Exchange Snap-In: Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.E2010". DomainName is a required parameter, and can take pipeline 
   input. 
 
   To execute the script, you will have to dot-source it first after which you can call the cmdlet: "Get-AcceptedDomainStatistics" 
 
.Parameter Domain 
   Generates report based on proxy addresses instead of primary addresses only. An object with multiple proxy addresses in a single domain will increment the count for that 
   domain for each address. This Paramater is mandatory 
 
.Parameter IncludeProxy 
   Generates report based on proxy addresses instead of primary addresses only. An object with multiple proxy addresses in a single domain will increment the count for that 
   domain for each address 
 
.Parameter OutCSV 
   Generatees a CSV report in the same directory the function is run from. The parameter needs to be given a CSV file name. 
 
.Parameter OutGrid 
   Generates a PowerShell grid report of accepted domain statistics. 
 
.Parameter MBXOnly 
   Generates a report for mailbox recipients only. 
 
.EXAMPLE 
   By default the function will only look at the primary SMTP addresses for recipients. DomainName is the only required parameter. To retrieve the number of objects 
   with a primary address in a given domain the below command can be used: 
 
   Get-AcceptedDomainStatistics -Domain Domain.com 
 
.EXAMPLE 
    This function now supports pipeline input as well. Input can be the Get-AcceptedDomain cmdlet (This may not work in Exchange 2007/2010 using remote Powershell), or 
    source. Pipeline info also may not work if theh accepted domain name and name do not match. The third example shows a workaround for this (This workaround will take 
    longer to run, and will consume more memory, it is not recommended for environments with over 10k recipients): 
 
   Get-AcceptedDomain | Get-AcceptedDomainStatistics 
   Get-Content Domains.txt | Get-AccpetedDomainStatistics 
   Get-AcceptedDomain | Select-Object -ExpandProperty DomainName | select Domain | Get-AcceptedDomainStatistics 
 
.EXAMPLE 
   By default the function will only look at the primary SMTP addresses for recipients. The following command will give you the number of primary SMTP addresses per 
   Accepted Domain for all recipient types and output it to GridView: 
 
   Get-AcceptedDomainStatistics -Domain Domain.com -OutGrid 
.EXAMPLE 
   There is also an option to look at all proxy addresses as well. In this mode if a user has multiple proxy addresses in a given domain, the count will incremente for  
   each address. The following command can be used to look at the proxy addresses for all recipient types and output to a CSV file: 
 
   Get-AcceptedDomainStatistics -IncludeProxy -OutCSV AcceptedDomains.csv 
 
#> 
function Get-AcceptedDomainStatistics 
{ 
    [CmdletBinding()] 
    [OutputType([int])] 
    Param 
    ( 
        [Parameter(Mandatory=$true,  
        ValueFromPipeline=$true, 
        ValueFromPipelineByPropertyName=$true)] 
        [string[]] 
        $DomainName, 
        [Parameter(Mandatory=$false)] 
        [switch] 
        $IncludeProxy, 
        [Parameter(Mandatory=$false)] 
        [String] 
        $OutCSV, 
        [Parameter(Mandatory=$false)] 
        [Switch] 
        $OutGrid, 
        [Parameter(Mandatory=$false)] 
        [Switch] 
        $MBXOnly 
    ) 
 
    Begin 
    { 
        $DomainCount = @() 
        #Gather recipient information 
        if($MBXOnly) 
        { 
            $ExRecipients = Get-Mailbox -Resultsize Unlimited | select emailaddresses,primarysmtpaddress 
        } 
        else 
        { 
            $ExRecipients = Get-Recipient -ResultSize Unlimited | select emailaddresses,primarysmtpaddress 
        } 
 
    } 
    Process 
    {    
        foreach($d in $DomainName) 
        { 
            $domobj = New-Object PSObject 
            $domobj | Add-Member NoteProperty -Name Domain -Value $d 
            $domobj | Add-Member NoteProperty -Name Count -Value "0" 
            $DomainCount = $DomainCount += $domobj 
        } 
        # Loop through recipients and count total primay addresses in each domain 
 
    Foreach($r in $ExRecipients) 
    { 
        if($IncludeProxy) 
        { 
            Foreach($d in $DomainName) 
            { 
                if($r.EmailAddresses -match $d) 
                { 
                    $dadd = $DomainCount | ?{$_.Domain -eq $d} 
                    $dadd.count = [int]$dadd.Count + 1 
                } 
            } 
        } 
        else 
        { 
            Foreach($d in $DomainName) 
            { 
                if($r.PrimarySmtpAddress.Domain -match $d) 
                { 
                    $dadd = $DomainCount | ?{$_.Domain -eq $d} 
                    $dadd.count = [int]$dadd.Count + 1 
                } 
 
            } 
        } 
    } 
 
    } 
    End 
    { 
    if($OutGrid) 
    { 
        $DomainCount | Out-GridView 
    } 
    if($OutCSV) 
    { 
        $DomainCount | Export-CSV -NoTypeInformation -Path $OutCSV 
    } 
    else 
    { 
        return $DomainCount 
    } 
    } 
}