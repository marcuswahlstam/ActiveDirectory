$timeStamp = Get-Date
$timeString = $timeStamp.ToString("yyyy-MM-dd_HH.mm")
$baseOUTest = "dn to ou"
$baseOUProd = "dn to ou"
$copyPath = "F:\Temp\GPOCopy_ProdToTest_$timeString"
$copyLogCSV = Join-Path -Path $copyPath -ChildPath "CopiedGPOs.csv"
$errorLogCSV = Join-Path -Path $copyPath -ChildPath "NotCopiedGPOs.csv"
$copyLog = Join-Path -Path $copyPath -ChildPath "CopiedGPOs.log"
#$errorLog = Join-Path -Path $copyPath -ChildPath "NotCopiedGPOs.log"

# Skapa mapp för export av Prod-GPOer
if (!(Test-Path $copyPath)){New-Item -Path $copyPath -ItemType Directory -Force | Out-Null}

# Skriv in header i CSV
"ProdGPOName;ProdGPOGuid;ProdOU;TestGPOName;TestGPOGUID;TestOU" | Out-File $copyLogCSV
"GPOName;GPOGuid;OU" | Out-File $errorLogCSV

# Hämta lista på OUn i Prod och Test
$OUsInBaseOUProd = Get-ADOrganizationalUnit -Filter * -SearchBase $baseOUProd -SearchScope Subtree | where {$_.DistinguishedName -notlike "*VDI*" -and $_.DistinguishedName -notlike "*GPOtest*"} | sort name
$OUsInBaseOUTest = Get-ADOrganizationalUnit -Filter * -SearchBase $baseOUTest -SearchScope Subtree | where {$_.DistinguishedName -notlike "*VDI*" -and $_.DistinguishedName -notlike "*GPOtest*"}

# Foreach-loop för OUs i Prod
foreach ($prodOU in $OUsInBaseOUProd)
{
    Remove-Variable linkedGPOsProd -ErrorAction Ignore

    # Hämta GPOs i aktuellt OU som loopen hanterar
    $linkedGPOsProd = $prodOU | Select-object -ExpandProperty LinkedGroupPolicyObjects

    # Hämta OU-namn
    $ouNameProd = $prodOU.Name

    # Sätt variabel för namn på Test-OU, baserat på Prod-OU. Byter ut "Prod" mot "Test" i namnet.
    $ouNameTest = $ouNameProd -replace "Prod","Test"
    $ouDNTest = (Get-ADOrganizationalUnit -Filter * | where {$_.Name -eq "$ouNameTest"}).DistinguishedName

    if (!$ouDNTest)
    {
        Write-Host "ERROR: Test OU `"$ouNameTest`" does not exist, skipping GPOs from OU $ouNameProd" -ForegroundColor Red
        "ERROR: Test OU `"$ouNameTest`" does not exist, skipping GPOs from OU $ouNameProd" | Out-File $copyLog -Append
        continue
    }

    # Länkade test-GPOer i aktuellt Test-OU
    $testOU = Get-ADOrganizationalUnit -Filter "Name -eq `"$ouNameTest`""
    $linkedGPOsTest = $testOU | Select-object -ExpandProperty LinkedGroupPolicyObjects
    if (!$linkedGPOsTest){$linkedGPOsTest = "cn={00000000-AAAA-0000-BBBB-000000000000}"}


    Write-Host "Working on OU `"$ouNameProd`"" -ForegroundColor Green
    "Working on OU `"$ouNameProd`"" | Out-File $copyLog -Append
    
    # Om OUt har länkade GPOs, gå vidare
    if ($linkedGPOsProd -ne $null)
    {
        # Skapa mapp för aktuellt OU
        $copyPathOU = Join-Path -Path $copyPath -ChildPath $ouNameProd
        if (!(Test-Path $copyPathOU)){New-Item -Path $copyPathOU -ItemType Directory -Force | Out-Null}

        # Foreach-loop för varje länkad GPO i aktuellt OU
        foreach ($prodGPO in $linkedGPOsProd)
        {
            # Hämta namn och GUID på aktuell GPO
            $gpoInfoProd = Get-GPO -Guid $($prodGPO.Substring(4,36))
            $gpoNameProd = $gpoInfoProd.DisplayName
            $gpoGUIDProd = $gpoInfoProd.Id

            Write-Host "- Working on GPO `"$gpoNameProd`"" -ForegroundColor Green
            "- Working on GPO `"$gpoNameProd`"" | Out-File $copyLog -Append

            # Lägg på " - TEST" på Prod-namnet för att skapa namnet på Test-GPO
            $gpoNameTest = $gpoNameProd + " - TEST"
            $gpoNameTestAlt = $gpoNameProd + "-TEST"

            # Skapa mapp för aktuell GPO dit export av aktuell GPO kommer göras.
            $copyPathGPO = Join-Path -Path $copyPathOU -ChildPath $gpoNameProd
            if (!(Test-Path $copyPathGPO)){New-Item -Path $copyPathGPO -ItemType Directory | Out-Null}

            # Kolla om GPO-länk är aktiv
            [xml]$GPOReport = (Get-GPOReport -Name "$gpoNameProd" -ReportType xml)
            $gpoProdEnabled = ($GPOReport.GPO.LinksTo | where {$_.SOMPath -like "*$ouNameProd"}).Enabled
            
            # Gör backup av aktuell GPO
            try
            {
                "-- Backing up $gpoNameProd to $copyPathGPO" | Out-File $copyLog -Append
                Write-Host "-- Backing up $gpoNameProd" -ForegroundColor Green
                Backup-GPO -Guid $gpoGUIDProd -Path $copyPathGPO | Out-Null
            }
            catch
            {
                "ERROR Backing up $gpoNameProd" | Out-File $copyLog -Append
                Write-Host "ERROR Backing up $gpoNameProd" -ForegroundColor Red
            }

            # Om en GPO hittas som heter $gpoNameTest, gör import av ovan export/backup
            if (Get-GPO -Name $gpoNameTest -ErrorAction Ignore)
            {
                $gpoInfoTest = Get-GPO -Name $gpoNameTest
                $gpoGUIDTest = $gpoInfoTest.Id
                try
                {
                    "-- Importing $gpoNameProd to $gpoNameTest" | Out-File $copyLog -Append
                    Write-Host "-- Importing backup of `"$gpoNameProd`" to `"$gpoNameTest`"" -ForegroundColor Green
                    Import-GPO -Path $copyPathGPO -BackupGpoName $gpoNameProd -TargetName $gpoNameTest | Out-Null

                    # Skriv info till loggen
                    "$gpoNameProd;$gpoGUIDProd;$prodOU;$gpoNameTest;$gpoGUIDTest;$ouNameTest" | Out-File $copyLogCSV -Append
                    
                    # Kolla länkning
                    if ($($linkedGPOsTest).substring(4,36) -contains $gpoGUIDTest)
                    {
                        "-- Link for `"$gpoNameTest`" to `"$ouNameTest`" exists" | Out-File $copyLog -Append
                        Write-Host "-- Link for `"$gpoNameTest`" to `"$ouNameTest`" exists" -ForegroundColor Green
                        
                        [xml]$GPOReportTest = (Get-GPOReport -Name "$gpoNameTest" -ReportType xml)
                        $gpoTestEnabled = ($GPOReportTest.GPO.LinksTo | where {$_.SOMPath -like "*$ouNameTest"}).Enabled

                        if ($gpoProdEnabled -eq "true" -and $gpoTestEnabled -eq "false")
                        {
                            Write-Host "-- Setting link to enabled" -ForegroundColor Green
                            "-- Setting link to enabled" | Out-File $copyLog -Append
                            Set-GPLink -Name $gpoNameTest -Target $ouNameTest -LinkEnabled Yes | Out-Null
                        }
                        elseif ($gpoProdEnabled -eq "false" -and $gpoTestEnabled -eq "true")
                        {
                            Write-Host "-- Setting link to disabled" -ForegroundColor Green
                            "-- Setting link to disabled" | Out-File $copyLog -Append
                            Set-GPLink -Name $gpoNameTest -Target $ouNameTest -LinkEnabled No | Out-Null
                        }
                        else
                        {
                            Write-Host "-- Link state OK" -ForegroundColor Green
                            "-- Link state OK" | Out-File $copyLog -Append
                        }
                    }                   
                    else
                    {
                        "-- Link to $ouNameTest DOES NOT exists for $gpoNameTest" | Out-File $copyLog -Append
                        Write-Host "-- Link to $ouNameTest DOES NOT exists for $gpoNameTest" -ForegroundColor Yellow
                        if ($gpoProdEnabled -eq "true")
                        {
                            Write-Host "-- Creating link for `"$gpoNameTest`" for `"$ouNameTest`" as enabled" -ForegroundColor Green
                            "-- Creating link for `"$gpoNameTest`" for `"$ouNameTest`" as enabled" | Out-File $copyLog -Append
                            New-GPLink -Name $gpoNameTest -Target $ouDNTest -LinkEnabled Yes | Out-Null
                        }
                        elseif ($gpoProdEnabled -eq "false")
                        {
                            Write-Host "-- Creating link for `"$gpoNameTest`" for `"$ouNameTest`" as disabled" -ForegroundColor Green
                            "-- Creating link for `"$gpoNameTest`" for `"$ouNameTest`" as disabled" | Out-File $copyLog -Append
                            New-GPLink -Name $gpoNameTest -Target $ouDNTest -LinkEnabled No | Out-Null
                        }
                    }
                }
                catch
                {
                    "$gpoNameProd;$gpoGUIDProd;$prodOU;$gpoNameTest;$gpoGUIDTest;$ouNameTest" | Out-File $errorLogCSV -Append
                    "-- ERROR Importing $copyPathGPO to $gpoNameTest" | Out-File $copyLog -Append
                    Write-Host "-- ERROR Importing $gpoNameTest" -ForegroundColor Red
                }
            }
            else
            {
                "-- $gpoNameTest does not exists, creating it" | Out-File $copyLog -Append
                Write-Host "-- $gpoNameTest does not exists, creating it" -ForegroundColor Yellow
                New-GPO -Name $gpoNameTest | Out-Null

                "-- Importing settings from $gpoNameProd to $gpoNameTest" | Out-File $copyLog -Append
                Write-Host "-- Importing settings from $gpoNameProd to $gpoNameTest" -ForegroundColor Yellow
                Import-GPO -Path $copyPathGPO -BackupGpoName $gpoNameProd -TargetName $gpoNameTest | Out-Null
                
                if ($gpoProdEnabled -eq "true")
                {
                    Write-Host "-- Creating link for $gpoNameTest to $ouNameTest as Enabled" -ForegroundColor Green
                    "-- Creating link for $gpoNameTest to $ouNameTest as Enabled" | Out-File $copyLog -Append
                    New-GPLink -Name $gpoNameTest -Target $ouDNTest -LinkEnabled Yes | Out-Null
                }
                elseif ($gpoProdEnabled -eq "false")
                {
                    Write-Host "-- Creating link for $gpoNameTest to $ouNameTest as Disabled" -ForegroundColor Green
                    "-- Creating link for $gpoNameTest to $ouNameTest as Disabled" | Out-File $copyLog -Append
                    New-GPLink -Name $gpoNameTest -Target $ouDNTest -LinkEnabled No | Out-Null
                }
            }
            #Write-Host "Sleeping 3 sec" -ForegroundColor Yellow
            #Start-Sleep 3
        }
    }
    else
    {
        "No GPOs linked to OU $prodOU" | Out-File $copyLog -Append
        "No GPOs linked to OU $prodOU"
    }
}
