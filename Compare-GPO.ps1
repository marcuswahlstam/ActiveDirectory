$baseOUProd = "dn to prod ou"
$baseOUDev = "dn to dev ou"

$OUsInBaseOUProd = Get-ADOrganizationalUnit -Filter * -SearchBase $baseOUProd -SearchScope Subtree | where {$_.DistinguishedName -notlike "*VDI*" -and $_.DistinguishedName -notlike "*GPOtest*"}
$OUsInBaseOUDev = Get-ADOrganizationalUnit -Filter * -SearchBase $baseOUDev -SearchScope Subtree | where {$_.DistinguishedName -notlike "*VDI*" -and $_.DistinguishedName -notlike "*GPOtest*"}


$allProdGPO = @()
$allTestGPO = @()

foreach ($prodOU in $OUsInBaseOUProd)
{
    $linkedGPOsProd = $prodOU | Select-object -ExpandProperty LinkedGroupPolicyObjects

    foreach ($prodGPO in $linkedGPOsProd)
    {
        # Hämta namn och GUID på aktuell GPO
        $gpoInfoProd = Get-GPO -Guid $($prodGPO.Substring(4,36))
        $gpoNameProd = $gpoInfoProd.DisplayName
        $gpoGUIDProd = $gpoInfoProd.Id
        $allProdGPO += $gpoNameProd
    }
}

foreach ($testOU in $OUsInBaseOUTest)
{
    $linkedGPOsTest = $testOU | Select-object -ExpandProperty LinkedGroupPolicyObjects

    foreach ($testGPO in $linkedGPOsTest)
    {
        # Hämta namn och GUID på aktuell GPO
        $gpoInfoTest = Get-GPO -Guid $($testGPO.Substring(4,36))
        $gpoNameTest = $gpoInfoTest.DisplayName
        $gpoGUIDTest = $gpoInfoTest.Id
        $allTestGPO += $gpoNameTest -replace " - TEST",""
    }
}

$allProdGPO = $allProdGPO | sort -Unique
$allTestGPO = $allTestGPO | sort -Unique

$existsOnlyInProd = Compare-Object $allProdGPO $allTestGPO | where {$_.sideindicator -eq "<="} | sort inputobject

$existsOnlyInTest = Compare-Object $allProdGPO $allTestGPO | where {$_.sideindicator -eq "=>"} | sort inputobject

Write-Host "`nEXISTS ONLY IN PROD OU`n----------------------" -ForegroundColor Green
$existsOnlyInProd.inputobject

Write-Host "`n`nEXISTS ONLY IN TEST OU`n----------------------" -ForegroundColor Green
$existsOnlyInTest.inputobject
