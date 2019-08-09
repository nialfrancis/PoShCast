####

[CmdletBinding()] Param(
    [int]$ExtraWeeks=0
)

####

$store = "$env:userprofile\Music\iTunes\iTunes Media\Podcasts"
$today = Get-Date
$dow = [int]$today.DayOfWeek

$PhoneName = 'Your Phone Name'

$ShowData = @{
    'DeepnBumpy' = @{
        'Path'     = 'deep-n-bumpy'
        'Provider' = 'DeepRadioNetwork'
        'Day'      = 'Friday' }
    'JackinGarage' = @{
        'Path'     = 'the-jackin-garage'
        'Provider' = 'DeepRadioNetwork'
        'Day'      = 'Saturday' }
    'UrbanNightGrooves' = @{
        'Path'     = 'urban-night-grooves'
        'Provider' = 'DeepRadioNetwork'
        'Day'      = 'Saturday' }
}

####

Enum Day {
    Sunday = 0
    Saturday = 1
    Friday = 2
    Thursday = 3
    Wednesday = 4
    Tuesday = 5
    Monday = 6
}

Import-Module BitsTransfer

function DateCalcWeekly ($PrevDay) {
    $result = @()
    
    foreach ($i in 0..$ExtraWeeks) {
        $subtract = [int]([Day]::$PrevDay) + ( $i * 7 ) + $dow
        $result += ($today.AddDays(-$subtract)).ToString("ddMMyy")
    }
    
    return $result
}

function CheckExisting {
    [CmdletBinding()] Param(
        [Parameter(Mandatory=$True)][array]$dates,
        [Parameter(Mandatory=$True)][string]$Show
    )
    
    $toget = @{}
    $provlogloc = (Join-Path $store $ShowData[$Show].Provider) + '.log'
    $provlog = Get-Content $provlogloc -ea Ignore
    
    foreach ($date in $dates) {
        if ( ($Show + '-' + $date) -notin $provlog ) {
            $toget[$date] = $provlogloc
        }
    }
    
    return $toget
}

function GetNew {
    [CmdletBinding()] Param(
        [Parameter()][hashtable]$toget,
        [Parameter()][string]$Show
    )
    
    $provider = $ShowData[$Show].Provider
    
    foreach ($item in $toget.GetEnumerator()) {
        $args = "-Show $Show -Date " + $item.Name + " -Log `'" + $item.Value + "`'"
        Invoke-Expression "$provider $args"
    }
}

function Get-Show {
    [CmdletBinding()] Param(
        [string]$Show
    )
    
    $dates = DateCalcWeekly $ShowData[$Show].Day
    $toget = CheckExisting -dates $dates -Show $Show
    GetNew -toget $toget -Show $Show
}

function DeepRadioNetwork {
    [CmdletBinding()] Param(
        [Parameter()][string]$Show,
        [Parameter()][string]$Date,
        [Parameter()][string]$Log
    )
    $uri = "http://media.d3ep.com/dl.php?f=" + $ShowData[$Show].Path + '-' + $Date + '.mp3'
    $FileName = [IO.Path]::Combine($store,$Show,$Date) + '.mp3'
    
    try {
        Start-BitsTransfer -Source $uri -Destination $FileName -Description "Downloading $Show $Date" -ea Stop
        "$Show-$Date" | Out-File -FilePath $Log -Append
    } catch {
        Write-Host "Download failed from $uri"
    }
}

Get-Show 'DeepnBumpy'
Get-Show 'JackinGarage'
Get-Show 'UrbanNightGrooves'

####

$Shell = New-Object -ComObject Shell.Application
$ShellItem = $Shell.NameSpace(17).Self
$Phone = $ShellItem.GetFolder.Items() | ?{$_.Name -eq $PhoneName}
if (!($Phone)) { break }

foreach ($dir in $ShowData.GetEnumerator()) {
    $DirName = $dir.Name
    $files = Get-ChildItem "$store\$DirName"
    $SourceDir = $Shell.NameSpace("$store\$DirName")
    
    $TargetDir = $Phone.GetFolder.ParseName("Internal shared storage\Music\$DirName")
    if (!($TargetDir)) {
        $new = $Phone.GetFolder.ParseName("Internal shared storage\Music")
        $new.GetFolder.NewFolder($DirName)
        $TargetDir = $Phone.GetFolder.ParseName("Internal shared storage\Music\$DirName")
    }
    
    foreach ($sf in $files) {
        if (!($TargetDir.GetFolder.ParseName($sf.Name))) {
            $SourceFile = $SourceDir.ParseName($sf.Name)
            $TargetDir.GetFolder.CopyHere($SourceFile)
        }
    }
}