<#
.Synopsis
    This script should be run every day after school ends to import .csv and/or .xlsx lists with students that should be removed or added to privatist group
    If there's a list for today, those students should be removed
    If there's a list for tomorrow, those students should be added
#>
[CmdletBinding(SupportsShouldProcess = $True)]
param(
    [Parameter()]
    [switch]$TestRun,

    [Parameter()]
    [DateTime]$EmulatedDate
)

######################################
# Configuration
######################################

$config = (Get-Content -Path "config.json" -Encoding UTF8) | ConvertFrom-Json

# folder path to the .csv and/or .xlsx files
$Path = $config.Path

# students will be added/removed from this active directory group
$group = $config.ad.group

# AD server to use
$adServer = $config.ad.server

# name of organization
$orgName = $config.orgName

######################################
# Functions
######################################

Function Start-FileParse
{
    [CmdletBinding(SupportsShouldProcess = $True)]
    param(
        [Parameter(Mandatory = $True)]
        [ValidateNotNull()]
        $FileContent,

        [Parameter(Mandatory = $True)]
        [ValidateSet("Add", "Remove")]
        [string]$Action,

        [Parameter(Mandatory = $True)]
        [string]$Class
    )

    # current person count
    [int]$CurrentIndex = 1

    # total person count
    [int]$IndexCount = $FileContent.Count
    if ($IndexCount -eq 0)
    {
        $IndexCount = 1
    }
    $Global:TotalIndexCount += $IndexCount

    # failed person count
    [int]$Failed = 0

    if ($IndexCount -gt 0)
    {
        Write-Log -Message "##########################################################################"
    }

    # check every Personid againt Active Directory
    foreach ($Person in $FileContent)
    {
        # create a new object to contain Person info
        $Obj = New-Object PSObject

        [string]$Id = $Person.Personid
        [string]$Name = $Person."Fullstendig navn"

        Write-Host "[$CurrentIndex / $IndexCount] -- [$Name] -- " -ForegroundColor Green -NoNewline
        Write-Log -Message "[$CurrentIndex / $IndexCount] -- [$Id - '$Name']"

        # check if Personid is given
        if ($Id -eq $null -or $Id -eq "")
        {
            Write-Host "[Personid er tom! Dette bør sjekkes!]" -ForegroundColor Red
            Write-Log -Message "[Personid er tom! Dette bør sjekkes!]" -Level ERROR

            $Obj | Add-Member NoteProperty Id("failure")
            $Obj | Add-Member NoteProperty Navn($Name)
            $Obj | Add-Member NoteProperty Type("Ikke sjekket: Personid er tom! Dette bør sjekkes!")
            $Obj | Add-Member NoteProperty Klasse($Class)

            $Global:ParsedStudents.Add($obj)

            $Global:TotalInvalidCount += 1
            $CurrentIndex += 1

            Write-Log -Message "##########################################################################"

            continue;
        }

        Write-Host "[$Id] -- " -ForegroundColor Green -NoNewline

        # check that Personid has correct number of digits
        if ($Id.get_Length() -ne 11)
        {
            if ($Id.get_Length() -eq 10)
            {
                # Personid is probably missing a leading 0
                Write-Host "[Ufullstendig Personid. Legger til 0] -- " -ForegroundColor Yellow -NoNewline
                Write-Log -Message "[Ufullstendig Personid. Legger til 0]" -Level WARNING
                $Id = "0$Id"
            }
            else
            {
                # Personid has invalid number of digits. Skip
                Write-Host "[Ugyldig Personid. Skipper '$Name']" -ForegroundColor Red
                Write-Log -Message "[Ugyldig Personid. Skipper '$Name']" -Level ERROR
                $Obj | Add-Member NoteProperty Id("failure$($Id)")
                $Obj | Add-Member NoteProperty Navn($Name)
                $Obj | Add-Member NoteProperty Type("Ikke sjekket: Ugyldig Personid! Dette bør sjekkes!")
                $Obj | Add-Member NoteProperty Klasse($Class)

                $Global:ParsedStudents.Add($obj)

                $Global:TotalInvalidCount += 1
                $CurrentIndex += 1

                Write-Log -Message "##########################################################################"

                continue;
            }
        }

        Write-Log -Message "[Using $Id]"

        # check if Personid exists in skole.top.no
        if ($config.ad.enabledUsersOnly)
        {
            $User = Get-ADUser -Server $adServer -Filter { employeeNumber -eq $Id -and Enabled -eq $True } -SearchBase $config.ad.searchBase -Properties DistinguishedName,DisplayName
        }
        else
        {
            $User = Get-ADUser -Server $adServer -Filter { employeeNumber -eq $Id } -SearchBase $config.ad.searchBase -Properties DistinguishedName,DisplayName
        }

        if ($User)
        {
            Write-Host "[Funnet] -- " -ForegroundColor Green -NoNewline
            Write-Log -Message "[Bruker funnet i AD]"

            $DN = $User.DistinguishedName
            $DisplayName = $User.DisplayName

            if ($Action -eq "Add")
            {
                try
                {
                    if (!$TestRun)
                    {
                        Add-ADGroupMember -Server $adServer -Identity $Group -Members $DN -Confirm:$false -ErrorAction Stop
                    }
                    else
                    {
                        Write-Host "[TestRun enabled. No add is made] -- " -ForegroundColor Cyan -NoNewline
                    }
                    Write-Host "['$DisplayName' lagt til]" -ForegroundColor Green
                    Write-Log -Message "['$DisplayName' lagt til i '$Group']"
                    $Obj | Add-Member NoteProperty Id("success$($Id.Substring(0, 6))******")
                    $Obj | Add-Member NoteProperty Navn($DisplayName)
                    $Obj | Add-Member NoteProperty Type("Lagt til i gruppe")
                    $Obj | Add-Member NoteProperty Klasse($Class)

                    $Global:TotalAddedCount += 1
                }
                catch
                {
                    Write-Host "['$DisplayName' ble ikke lagt til]" -ForegroundColor Red
                    Write-Log -Message "['$DisplayName' ikke lagt til i '$Group': $_]" -Level ERROR
                    $Obj | Add-Member NoteProperty Id("failure$($Id.Substring(0, 6))******")
                    $Obj | Add-Member NoteProperty Navn($DisplayName)
                    $Obj | Add-Member NoteProperty Type("Feilet ved innmelding i gruppe")
                    $Obj | Add-Member NoteProperty Klasse($Class)

                    New-ArcheoMessage @archeoSplat -MessageType "Notification" -Description "'$DisplayName' ble ikke lagt til i '$Group'" -Status "Error" -MetaData @{ "ErrorMessage" = $_.ToString() }

                    $Global:TotalFailedCount += 1
                }

            }
            elseif ($Action -eq "Remove")
            {
                try
                {
                    if (!$TestRun)
                    {
                        Remove-ADGroupMember -Server $adServer -Identity $Group -Members $DN -Confirm:$false -ErrorAction Stop
                    }
                    else
                    {
                        Write-Host "[TestRun enabled. No remove is made] -- " -ForegroundColor Cyan -NoNewline
                    }
                    Write-Host "['$DisplayName' fjernet]" -ForegroundColor Green
                    Write-Log -Message "['$DisplayName' fjernet fra '$Group']"
                    $Obj | Add-Member NoteProperty Id("success$($Id.Substring(0, 6))******")
                    $Obj | Add-Member NoteProperty Navn($DisplayName)
                    $Obj | Add-Member NoteProperty Type("Fjernet fra gruppe")
                    $Obj | Add-Member NoteProperty Klasse($Class)

                    $Global:TotalRemovedCount += 1
                }
                catch
                {
                    Write-Host "['$DisplayName' ble ikke fjernet]" -ForegroundColor Red
                    Write-Log -Message "['$DisplayName' ikke fjernet fra '$Group': $_]" -Level ERROR
                    $Obj | Add-Member NoteProperty Id("failure$($Id.Substring(0, 6))******")
                    $Obj | Add-Member NoteProperty Navn($DisplayName)
                    $Obj | Add-Member NoteProperty Type("Feilet ved fjerning fra gruppe")
                    $Obj | Add-Member NoteProperty Klasse($Class)

                    New-ArcheoMessage @archeoSplat -MessageType "Notification" -Description "'$DisplayName' ble ikke fjernet fra '$Group'" -Status "Error" -MetaData @{ "ErrorMessage" = $_.ToString() }

                    $Global:TotalFailedCount += 1
                }
            }
        }
        else
        {
            Write-Host "[Ikke funnet]" -ForegroundColor Yellow
            Write-Log -Message "Bruker eksister ikke i AD" -Level WARNING
            $Obj | Add-Member NoteProperty Id("$($Id.Substring(0, 6))******")
            $Obj | Add-Member NoteProperty Navn($Name)
            $Obj | Add-Member NoteProperty Type("Bruker er ikke elev i $orgName")
            $Obj | Add-Member NoteProperty Klasse($Class)

            $Global:TotalFailedCount += 1
        }

        $CurrentIndex += 1

        $Global:ParsedStudents.Add($obj)

        Write-Log -Message "##########################################################################"
    }
}

Function sendmail([string] $body)
{
    $SmtpClient = New-Object System.Net.Mail.SmtpClient
    $MailMessage = New-Object System.Net.Mail.MailMessage

    $SmtpClient.Host = $config.smtp.server
    $MailMessage.From = (New-Object System.Net.Mail.MailAddress -ArgumentList $config.smtp.fromAddress,$config.smtp.fromDisplayName)

    $config.smtp.bcc | % { $MailMessage.Bcc.Add((New-Object System.Net.Mail.MailAddress -ArgumentList $_.address,$_.displayName)) }

    if (!$TestRun)
    {
        $config.smtp.to | % { $MailMessage.To.Add((New-Object System.Net.Mail.MailAddress -ArgumentList $_.address,$_.displayName)) }
        $config.smtp.cc | % { $MailMessage.CC.Add((New-Object System.Net.Mail.MailAddress -ArgumentList $_.address,$_.displayName)) }
    }
    else
    {
        $testRun = "TestRun enabled. Receivers in 'To' and 'Cc' will not receive this testmail!"
        $testRunTo = "'To' recipients would have been: '$(($config.smtp.to | % { "$($_.displayName) ($($_.address))" }) -join "', '")'"
        $testRunCc = "'Cc' recipients would have been: '$(($config.smtp.cc | % { "$($_.displayName) ($($_.address))" }) -join "', '")'"
        Write-Host "`n$testRun" -ForegroundColor Cyan
        Write-Log "$testRun" -Level WARNING
        Write-Host "$testRunTo" -ForegroundColor Cyan
        Write-Log "$testRunTo" -Level WARNING
        Write-Host "$testRunCc" -ForegroundColor Cyan
        Write-Log "$testRunCc" -Level WARNING
    }
    
    $MailMessage.Subject = "Privatisteksamen - $Today / $Tomorrow"
    $MailMessage.Body = $body
    $MailMessage.IsBodyHtml = $True

    try
    {
        $SmtpClient.Send($MailMessage)
        $MailLog = ""

        foreach ($Recipient in $MailMessage.To)
        {
            if ($MailLog -eq "")
            {
                $MailLog = "Mailen er sendt til $Recipient"
            }
            else
            {
                $MailLog = "$MailLog,$Recipient"
            }
        }
        
        if ($MailMessage.CC.Count -gt 0)
        {
            $MailLog = "$MailLog "

            foreach ($Recipient in $MailMessage.CC)
            {
                if ($MailLog.EndsWith(" "))
                {
                    $MailLog = "$($MailLog)med kopi til $Recipient"
                }
                else
                {
                    $MailLog = "$MailLog,$Recipient"
                }
            }
        }

        if ($MailMessage.Bcc.Count -gt 0)
        {
            $MailLog = "$MailLog "

            foreach ($Recipient in $MailMessage.Bcc)
            {
                if ($MailLog.EndsWith(" "))
                {
                    $MailLog = "$($MailLog)med blindkopi til $Recipient"
                }
                else
                {
                    $MailLog = "$MailLog,$Recipient"
                }
            }
        }

        if (!$MailLog.StartsWith("Mailen er sendt"))
        {
            $MailLog = "Mailen er sendt$MailLog"
        }

        Write-Host "$MailLog" -ForegroundColor Green
        Write-Log -Message $MailLog
    }
    catch
    {
        Write-Host "Feilet ved sending av mail: $_" -ForegroundColor Red
        Write-Log -Message "Feilet ved sending av mail: $_" -Level ERROR
    }
}

######################################
# variables used in this script
######################################

# date variable used troughout this script
if ($EmulatedDate)
{
    $Now = $EmulatedDate
}
else
{
    $Now = Get-Date
}

# students will be added/removed from this active directory group
Write-Host "Using Active Directory group '$Group'" -ForegroundColor Green

# files with this date in its name will be parsed as remove. This date will also be used as log file name
$Today = $Now.ToShortDateString()
Write-Host "$Today -- Files with this date will be parsed as remove" -ForegroundColor Green

# files with this date in its name will be parsed as add
$Tomorrow = $Now.AddDays(1).ToShortDateString()
Write-Host "$Tomorrow -- Files with this date will be parsed as add" -ForegroundColor Green

# get files from $Path to remove from Privatister
$RemoveFiles = Get-ChildItem -Path "$Path\*" -Recurse -Include "*.xlsx","*.csv" | Where { $_.BaseName.Contains($Today) }

# get files from $Path to add as Privatister
$AddFiles = Get-ChildItem -Path "$Path\*" -Recurse -Include "*.xlsx","*.csv" | Where { $_.BaseName.Contains($Tomorrow) }

if ($RemoveFiles -or $AddFiles)
{
    # create log folder
    $LogSubfolder = ""
    if ($Now.Month -ge 5 -and $Now.Month -le 6) # summer exam (lastYear_year_vår)
    {
        $LogSubfolder = "$(($Now.Year - 1))_$($Now.Year)_vår"
    }
    elseif ($Now.Month -ge 11 -and $Now.Month -le 12) # winter exam (year_nextYear_jul)
    {
        $LogSubfolder = "$($Now.Year)_$(($Now.Year + 1))_jul"
    }
    else # Privatist exam outside of regular exam times
    {
        if ($Now.Month -ge 8) # fall-winter exam
        {
            $LogSubfolder = "$($Now.Year)_$(($Now.Year + 1))_$($Now.Month)"
        }
        elseif ($Now.Month -ge 1 -and $Now.Month -le 4) # winter-spring exam
        {
            $LogSubfolder = "$(($Now.Year - 1))_$($Now.Year)_$($Now.Month)"
        }
    }

    if ($TestRun)
    {
        Add-LogTarget -Name CMTrace -Configuration @{ Path = "$LogSubfolder\$($today)_TestRun.log" }
    }
    else
    {
        Add-LogTarget -Name CMTrace -Configuration @{ Path = "$LogSubfolder\$today.log" }
    }
}

# go through csv files from $Path to remove from Privatister
if ($RemoveFiles)
{
    Write-Host "$($RemoveFiles.Count) files will be parsed for removal" -ForegroundColor Green
    Write-Log -Message "##########################################################################"
    Write-Log -Message "##########################################################################"
    Write-Log -Message "##########################################################################"
    Write-Log -Message "$($RemoveFiles.Count) files will be parsed for removal"
}
else
{
    Write-Host "0 files will be parsed for removal" -ForegroundColor Green
    #Write-Log -Message "0 files will be parsed for removal"
}

# go through csv files from $Path to add as Privatister
if ($AddFiles)
{
    Write-Host "$($AddFiles.Count) files will be parsed for adding" -ForegroundColor Green
    if (!$RemoveFiles)
    {
        Write-Log -Message "##########################################################################"
        Write-Log -Message "##########################################################################"
        Write-Log -Message "##########################################################################"
    }
    Write-Log -Message "$($AddFiles.Count) files will be parsed for adding"
}
else
{
    Write-Host "0 files will be parsed for adding" -ForegroundColor Green
    #Write-Log -Message "0 files will be parsed for adding"
}

# object to contain parsed students; added, removed and failed
$Global:ParsedStudents = [System.Collections.Generic.List[PSObject]]::new()
[int]$Global:TotalIndexCount = 0
[int]$Global:TotalInvalidCount = 0
[int]$Global:TotalAddedCount = 0
[int]$Global:TotalRemovedCount = 0
[int]$Global:TotalFailedCount = 0

############################
# parsing of csv files
############################

# go through all Privatister to remove
foreach ($File in $RemoveFiles)
{
    Write-Host "`r`nRemove privatister from: '$($File.Name)'`r`n" -ForegroundColor Green
    Write-Log -Message "Remove privatister from: '$($File.Name)'"

    if ($File.Extension -eq ".csv")
    {
        # changing encoding from ANSI to UTF-8
        $FileContent = Get-Content -Path $File.FullName
        $FileContent | Set-Content -Encoding UTF8 -Path $File.FullName

        $fileContent = Import-Csv -Path $File.FullName -Delimiter ';' -Encoding UTF8
    }
    elseif ($File.Extension -eq ".xlsx")
    {
        $fileContent = Import-Excel -Path $File.FullName -HeaderName "Personid","Fullstendig navn"
    }

    Start-FileParse -FileContent $fileContent -Action Remove -Class $File.BaseName
}

# go through all Privatister to remove
foreach ($File in $AddFiles)
{
    Write-Host "`r`nAdd privatister from: '$($File.Name)'`r`n" -ForegroundColor Green
    Write-Log -Message "Add privatister from: '$($File.Name)'"

    if ($File.Extension -eq ".csv")
    {
        # changing encoding from ANSI to UTF-8
        $FileContent = Get-Content -Path $File.FullName
        $FileContent | Set-Content -Encoding UTF8 -Path $File.FullName

        $fileContent = Import-Csv -Path $File.FullName -Delimiter ';' -Encoding UTF8
    }
    elseif ($File.Extension -eq ".xlsx")
    {
        $fileContent = Import-Excel -Path $File.FullName -HeaderName "Personid","Fullstendig navn"
    }
    
    Start-FileParse -FileContent $fileContent -Action Add -Class $File.BaseName
}

###############################
# sending mail rapports 
###############################

if ($Global:ParsedStudents.Count -gt 0)
{
    $header = "<meta http-equiv=`"content-type`" content=`"text/html;charset=utf-8`">`r`n<title>Privatisteksamen - $Today / $Tomorrow</title>`r`n<style>`r`ntd, th { border: 1px solid black; }`r`nth { border-bottom: 3px solid black; }`r`n.success { background: #00FF00; }`r`n.failure { background: #FF0000; }`r`n</style>"
    [string]$message = "<h1>Privatisteksamen</h1>`r`n<h4>`r`n<b>Antall kandidater sjekket</b>: $($Global:TotalIndexCount)<br />`r`n<b>Antall kandidater med ugyldig personnummer</b>: $($Global:TotalInvalidCount)<br />`r`n<b>Antall kandidater sperret</b>: $($Global:TotalAddedCount)<br />`r`n<b>Antall kandidater åpnet</b>: $($Global:TotalRemovedCount)<br />`r`n<b>Antall kandidater ikke elev i $orgName</b>: $($Global:TotalFailedCount)<br />`r`n</h4>`r`n$($Global:ParsedStudents | ConvertTo-Html -Fragment | Out-String)"
    $htmlMessage = (ConvertTo-Html -Body $message -Head $header | Out-String) -replace "\s<table>\s+</table>\s"

    # add success and failure classes
    $htmlMessage = $htmlMessage.Replace("<tr><td>success", "<tr class=`"success`"><td>").Replace("<tr><td>failure", "<tr class=`"failure`"><td>")

    if ($TestRun)
    {
        $htmlMessageFilePath = "$(Get-LogDir)\$($LogSubfolder)\$($Today)_TestRun_Message.html"
        [System.IO.File]::WriteAllText($htmlMessageFilePath, $htmlMessage)
    }
    else
    {
        $htmlMessageFilePath = "$(Get-LogDir)\$($LogSubfolder)\$($Today)_Message.html"
        [System.IO.File]::WriteAllText($htmlMessageFilePath, $htmlMessage)
    }
    Write-Log -Message "HTML file: '$htmlMessageFilePath'"
    sendmail $htmlMessage

    Write-Host "`nAntall kandidater: $($Global:TotalIndexCount)" -ForegroundColor Green
    Write-Log -Message "Antall kandidater: $($Global:TotalIndexCount)"

    Write-Host "Kandidater med ugyldig personnummer: $($Global:TotalInvalidCount)" -ForegroundColor Red
    Write-Log -Message "Kandidater med ugyldig personnummer: $($Global:TotalInvalidCount)" -Level ERROR

    Write-Host "Kandidater sperret: $($Global:TotalAddedCount)"
    Write-Log -Message "Kandidater sperret: $($Global:TotalAddedCount)"

    Write-Host "Kandidater åpnet: $($Global:TotalRemovedCount)"
    Write-Log -Message "Kandidater åpnet: $($Global:TotalRemovedCount)"

    Write-Host "Kandidater ikke elev i $($orgName): $($Global:TotalFailedCount)" -ForegroundColor Yellow
    Write-Log -Message "Kandidater ikke elev i $($orgName): $($Global:TotalFailedCount)" -Level WARNING
}
else
{
    #Write-Log -Message "No work to be done..."
}
