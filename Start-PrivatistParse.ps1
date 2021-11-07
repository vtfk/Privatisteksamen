<#
.Synopsis
    This script should be run every day after school ends to import .csv and/or .xlsx (.xls files can be used, but will be converted to .xlsx before used!) lists with students that should be removed or added to privatist group
    Students with Eksamensdato for today will be removed from Active Directory group
    Students with Eksamensdato for tomorrow will be added to Active Directory group
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

Function Start-StudentParse
{
    [CmdletBinding(SupportsShouldProcess = $True)]
    param(
        [Parameter(Mandatory = $True)]
        [string]$Class,
    
        [Parameter(Mandatory = $True)]
        [ValidateNotNull()]
        [string]$SSN,

        [Parameter(Mandatory = $True)]
        [ValidateSet("Add", "Remove")]
        [string]$Action
    )


    # check every Fødselsnummer againt Active Directory
    Write-Log -Message "##########################################################################"

    # create a new object to contain Person info
    $Obj = New-Object PSObject

    Write-Host "[$CurrentIndex / $IndexCount] -- [$SSN] -- " -ForegroundColor Green -NoNewline
    Write-Log -Message "[$CurrentIndex / $IndexCount] -- [$SSN]"

    # check if Fødselsnummer is given
    if ($SSN -eq $null -or $SSN -eq "")
    {
        Write-Host "[Fødselsnummer er tom! Dette bør sjekkes!]" -ForegroundColor Red
        Write-Log -Message "[Fødselsnummer er tom! Dette bør sjekkes!]" -Level ERROR

        $Obj | Add-Member NoteProperty Id("failure")
        $Obj | Add-Member NoteProperty Navn("")
        $Obj | Add-Member NoteProperty Type("Ikke sjekket: Fødselsnummer er tom! Dette bør sjekkes!")
        $Obj | Add-Member NoteProperty Eksamensparti($Class)

        $Global:ParsedStudents.Add($obj)

        $Global:TotalInvalidCount += 1

        Write-Log -Message "##########################################################################"

        return
    }

    # check that Fødselsnummer has correct number of digits
    if ($SSN.get_Length() -ne 11)
    {
        if ($SSN.get_Length() -eq 10)
        {
            # Fødselsnummer is probably missing a leading 0
            Write-Host "[Ufullstendig Fødselsnummer. Legger til 0] -- " -ForegroundColor Yellow -NoNewline
            Write-Log -Message "[Ufullstendig Fødselsnummer. Legger til 0]" -Level WARNING
            $SSN = "0$SSN"
        }
        else
        {
            # Fødselsnummer has invalid number of digits. Skip
            Write-Host "[Ugyldig Fødselsnummer. Skipper '$Name']" -ForegroundColor Red
            Write-Log -Message "[Ugyldig Fødselsnummer. Skipper '$Name']" -Level ERROR
            $Obj | Add-Member NoteProperty Id("failure$($SSN)")
            $Obj | Add-Member NoteProperty Navn("")
            $Obj | Add-Member NoteProperty Type("Ikke sjekket: Ugyldig Fødselsnummer! Dette bør sjekkes!")
            $Obj | Add-Member NoteProperty Eksamensparti($Class)

            $Global:ParsedStudents.Add($obj)

            $Global:TotalInvalidCount += 1

            Write-Log -Message "##########################################################################"

            return
        }
    }

    Write-Log -Message "[Using $SSN]"

    # check if Fødselsnummer exists in skole.top.no
    if ($config.ad.enabledUsersOnly)
    {
        $User = Get-ADUser -Server $adServer -Filter { employeeNumber -eq $SSN -and Enabled -eq $True } -SearchBase $config.ad.searchBase -Properties DistinguishedName,DisplayName
    }
    else
    {
        $User = Get-ADUser -Server $adServer -Filter { employeeNumber -eq $SSN } -SearchBase $config.ad.searchBase -Properties DistinguishedName,DisplayName
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
                $Obj | Add-Member NoteProperty Id("success$($SSN.Substring(0, 6))******")
                $Obj | Add-Member NoteProperty Navn($DisplayName)
                $Obj | Add-Member NoteProperty Type("Lagt til i gruppe")
                $Obj | Add-Member NoteProperty Eksamensparti($Class)

                $Global:TotalAddedCount += 1
            }
            catch
            {
                Write-Host "['$DisplayName' ble ikke lagt til]" -ForegroundColor Red
                Write-Log -Message "['$DisplayName' ikke lagt til i '$Group': $_]" -Level ERROR
                $Obj | Add-Member NoteProperty Id("failure$($SSN.Substring(0, 6))******")
                $Obj | Add-Member NoteProperty Navn($DisplayName)
                $Obj | Add-Member NoteProperty Type("Feilet ved innmelding i gruppe")
                $Obj | Add-Member NoteProperty Eksamensparti($Class)

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
                $Obj | Add-Member NoteProperty Id("success$($SSN.Substring(0, 6))******")
                $Obj | Add-Member NoteProperty Navn($DisplayName)
                $Obj | Add-Member NoteProperty Type("Fjernet fra gruppe")
                $Obj | Add-Member NoteProperty Eksamensparti($Class)

                $Global:TotalRemovedCount += 1
            }
            catch
            {
                Write-Host "['$DisplayName' ble ikke fjernet]" -ForegroundColor Red
                Write-Log -Message "['$DisplayName' ikke fjernet fra '$Group': $_]" -Level ERROR
                $Obj | Add-Member NoteProperty Id("failure$($SSN.Substring(0, 6))******")
                $Obj | Add-Member NoteProperty Navn($DisplayName)
                $Obj | Add-Member NoteProperty Type("Feilet ved fjerning fra gruppe")
                $Obj | Add-Member NoteProperty Eksamensparti($Class)

                $Global:TotalFailedCount += 1
            }
        }
    }
    else
    {
        Write-Host "[Ikke funnet]" -ForegroundColor Yellow
        Write-Log -Message "Bruker eksister ikke i AD" -Level WARNING
        $Obj | Add-Member NoteProperty Id("$($SSN.Substring(0, 6))******")
        $Obj | Add-Member NoteProperty Navn($Name)
        $Obj | Add-Member NoteProperty Type("Bruker er ikke elev i $orgName")
        $Obj | Add-Member NoteProperty Eksamensparti($Class)

        $Global:TotalFailedCount += 1
    }

    $Global:ParsedStudents.Add($obj)
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
Write-Host "$Today -- Candidates with this date as 'Eksamensdato' will be parsed as remove" -ForegroundColor Green

# files with this date in its name will be parsed as add
$Tomorrow = $Now.AddDays(1).ToShortDateString()
Write-Host "$Tomorrow -- Candidates with this date as 'Eksamensdato' will be parsed as add" -ForegroundColor Green

# get files from $Path to parse
$files = Get-ChildItem -Path "$Path\*" -Recurse -Include "*.xlsx","*.csv","*.xls"

if (!$files -or ($files | Measure-Object | Select-Object -ExpandProperty Count) -eq 0) {
    Write-Host "No files found"
    return
}

# object to contain students to parse
$Global:StudentsToParse = [System.Collections.Generic.List[PSObject]]::new()
# object to contain parsed students; added, removed and failed
$Global:ParsedStudents = [System.Collections.Generic.List[PSObject]]::new()
[int]$Global:TotalIndexCount = 0
[int]$Global:TotalInvalidCount = 0
[int]$Global:TotalAddedCount = 0
[int]$Global:TotalRemovedCount = 0
[int]$Global:TotalFailedCount = 0

foreach ($file in $files) {
    if ($file.Extension -eq ".xls") {
        try {
            ConvertTo-ExcelXlsx -Path $file.FullName -ErrorAction Stop
            Write-Host "Converted 'xls' to 'xlsx' ($($file.FullName))" -ForegroundColor Green
            Remove-Item -Path $file.FullName -Force -Confirm:$False # TODO: Instead of removing it, move it to same folder as finished ones
            $filePath = "$($file.FullName)x"
        }
        catch {
            Write-Host "Failed to convert 'xls' to 'xlsx' ($($file.FullName)) : $_" -ForegroundColor Red
            continue
        }
    }
    else {
        $filePath = $file.FullName
    }

    $fileContent = Import-Excel -Path $filePath
    $fileContent | Where { $_.Eksamensdato -eq $Today } | ForEach-Object {
        $obj = New-Object PSObject @{
            Eksamensparti = $_.Eksamensparti
            Eksamensdato = $_.Eksamensdato
            Fødselsnummer = $_.Fødselsnummer
            Remove = $true
        }

        $Global:StudentsToParse.Add($obj)
    }
    $fileContent | Where { $_.Eksamensdato -eq $Tomorrow } | ForEach-Object {
        $obj = New-Object PSObject @{
            Eksamensparti = $_.Eksamensparti
            Eksamensdato = $_.Eksamensdato
            Fødselsnummer = $_.Fødselsnummer
            Add = $true
        }

        $Global:StudentsToParse.Add($obj)
    }
}

if ($Global:StudentsToParse.Count -eq 0)
{
    Write-Host "Files don't contain dates for today/tomorrow" -ForegroundColor Yellow
    return
}

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

# current person count
[int]$currentIndex = 1

$studentsToRemove = $Global:StudentsToParse | Where { $_.Remove -and $_.Remove -eq $True  }
Write-Host "`n$($studentsToRemove.Count) candidates will be parsed for removal" -ForegroundColor Green
Write-Log -Message "##########################################################################"
Write-Log -Message "##########################################################################"
Write-Log -Message "$($studentsToRemove.Count) candidates will be parsed for removal"

# total person count to remove
[int]$indexCount = $studentsToRemove | Measure-Object | Select-Object -ExpandProperty Count
$Global:TotalIndexCount += $indexCount
foreach ($student in $studentsToRemove) {
    Start-StudentParse -Class $student.Eksamensparti -SSN $student.Fødselsnummer -Action Remove
    $CurrentIndex += 1
}

# current person count
[int]$currentIndex = 1

$studentsToAdd = $Global:StudentsToParse | Where { $_.Add -and $_.Add -eq $True }
Write-Host "`n$($studentsToAdd.Count) candidates will be parsed for adding" -ForegroundColor Green
Write-Log -Message "##########################################################################"
Write-Log -Message "##########################################################################"
Write-Log -Message "$($studentsToAdd.Count) candidates will be parsed for adding"

# total person count to add
[int]$indexCount = $studentsToAdd | Measure-Object | Select-Object -ExpandProperty Count
$Global:TotalIndexCount += $indexCount
foreach ($student in $studentsToAdd) {
    Start-StudentParse -Class $student.Eksamensparti -SSN $student.Fødselsnummer -Action Add
    $CurrentIndex += 1
}

###############################
# sending mail reports 
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

# TODO: Check all files for dates in the future. If a file doesn't have anymore dates into the future, move the file to another directory, keeping the same directory syntax