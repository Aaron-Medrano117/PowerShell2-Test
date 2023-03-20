## Add Domains
$domains = import-csv $path

$domains | foreach {New-MsolDomain -name $_.domain}

#### Get DNS verification

$VerificationRecords =@()
$unverifiedDomains = get-msoldomain -status unverified
foreach ($lookup in $unverifiedDomains)
{
    #$VerificationValue = Get-MsolDomainVerificationDns -DomainName $lookup.name
    #$DNS = $VerificationValue.label -split '.'
    #$TXT = "MS=" + $VerificationValue.label

    $VerificationSetting = Get-MsolDomainVerificationDns -DomainName $lookup.name
    $TXT = "MS=" + $VerificationSetting.label

    Write-Host "$txt"

    $tmp = "" | Select Domain, Verification_Record_Type, Verification_Record_Value
    $tmp.Domain = $lookup.name
    $tmp.Verification_Record_Type = "TXT"
    $tmp.Verification_Record_Value = $TXT
    $VerificationRecords += $tmp
}

## Verify Domains

$domains = import-csv $path

$domains | foreach {Confirm-MsolDomain -domainname $_.domain}


$unverifiedDomains | foreach {Confirm-MsolDomain -domainname $_.name}