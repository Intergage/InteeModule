# Random Functions - Functions that don't really do much
function banner {
    param ([Parameter(Mandatory = $true)][string]$string)

    $length = $string.Length + 6
    $boarder = "#" * [int]$length

    Write-Host -Foreground "red" $boarder
    Write-Host -Foreground "red" "#" -NoNewline; Write-Host "  $($string)  " -NoNewline; Write-Host -Foreground "red" "#"
    Write-Host -Foreground "red" $boarder
}
function Get-Boobs {
    param([int]$x)

    if (!$x) {
        Write-Host -ForegroundColor White -BackgroundColor Black ' ( . Y . ) ' -NoNewline
    }
    else {
        foreach ($y in 1..$x) {
            Write-Host -ForegroundColor White -BackgroundColor Black ' ( . Y . ) ' -NoNewline
        }
    }
}
function Start-List {
    $list = @()
    $title = Read-Host -Prompt "List Title _> "
    
    Write-Host "`nWhen finished type 'done' to print list`n"
    
    $loopsy = $true

    while ($loopsy) {
        $item = Read-Host -Prompt "Item _> "
        
        if ($item.ToLower() -eq 'done') {
            clear
            banner $title
            Write-Host "----------"
            $itemNum = 1
            foreach ($stage in $list) {
                Write-Host $itemNum':', $stage
                $itemNum += 1
                
                $loopsy = $false
            } 
            Write-Host "----------`n"
        }
        
        else {$list += $item}
    }

}
function watch {
    param ([Parameter(Mandatory = $true)][string]$cmd)
    
    while ($true -eq $true) {
        & $cmd
        write-host "`t`t#########################"
        sleep 5
    }
}
function su () {
    Start-Process powershell -Verb runAs
}
Function Clean-String($Str) {
    Foreach ($Char in [Char[]]"!@#$%^&*(){}|\/?><,.][+=-_") {$str = $str.replace("$Char", '')}
    Return $str
}
function ConvertTo-AsciiImg () {
    #------------------------------------------------------------------------------ 
    # Copyright 2006 Adrian Milliner (ps1 at soapyfrog dot com) 
    # http://ps1.soapyfrog.com 
    # 
    # This work is licenced under the Creative Commons  
    # Attribution-NonCommercial-ShareAlike 2.5 License.  
    # To view a copy of this licence, visit  
    # http://creativecommons.org/licenses/by-nc-sa/2.5/  
    # or send a letter to  
    # Creative Commons, 559 Nathan Abbott Way, Stanford, California 94305, USA. 
    #------------------------------------------------------------------------------ 
 
    #------------------------------------------------------------------------------ 
    # This script loads the specified image and outputs an ascii version to the 
    # pipe, line by line. 
    # 
    param( 
        [string]$path = "C:\Users\qxcin\Dropbox\Pictures\download.jpg", 
        [int]$maxwidth, # default is width of console 
        [string]$palette = "ascii", # choose a palette, "ascii" or "shade" 
        [float]$ratio = 1.5        # 1.5 means char height is 1.5 x width 
    ) 
 
 
 
    #------------------------------------------------------------------------------ 
    # here we go 
 
    $palettes = @{ 
        "ascii" = " .,:;=|iI+hHOE#`$" 
        "shade" = " " + [char]0x2591 + [char]0x2592 + [char]0x2593 + [char]0x2588 
        "bw"    = " " + [char]0x2588 
    } 
    $c = $palettes[$palette] 
    if (-not $c) { 
        write-warning "palette should be one of:  $($palettes.keys.GetEnumerator())" 
        write-warning "defaulting to ascii" 
        $c = $palettes.ascii 
    } 
    [char[]]$charpalette = $c.ToCharArray() 
 
    # we need the drawing assembly 
    try {
        $dllpath = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
        [Reflection.Assembly]::LoadFrom($dllpath) | out-null 
        # load the image 
        $image = [Drawing.Image]::FromFile($path) 
        if ($maxwidth -le 0) { [int]$maxwidth = $host.ui.rawui.WindowSize.Width - 1} 
        [int]$imgwidth = $image.Width 
        [int]$maxheight = $image.Height / ($imgwidth / $maxwidth) / $ratio 
        $bitmap = new-object Drawing.Bitmap ($image, $maxwidth, $maxheight) 
        [int]$bwidth = $bitmap.Width; [int]$bheight = $bitmap.Height 
        # draw it! 
        $cplen = $charpalette.count 
        for ([int]$y = 0; $y -lt $bheight; $y++) { 
            $line = "" 
            for ([int]$x = 0; $x -lt $bwidth; $x++) { 
                $colour = $bitmap.GetPixel($x, $y) 
                $bright = $colour.GetBrightness() 
                [int]$offset = [Math]::Floor($bright * $cplen) 
                $ch = $charpalette[$offset] 
                if (-not $ch) { $ch = $charpalette[-1] } #overflow 
                $line += $ch 
            } 
            $line 
        } 
    }
    catch {}

}


# System Functions - Functions that return local system information
function Get-DirInfo {
    param([string]$dir)

    #Don't need Attributes as 'Mode' is the same. d--hs- = Directory, Hidden, System. da---- = Directory, Archive. etc
    Get-ChildItem -Force $dir | ft -AutoSize Mode, @{expression = {[math]::Round($_.Length / 1MB, 2)}; label = 'MB' }, IsReadOnly, LastAccessTime, name
}
function Get-Temperature {
    $temps = Get-WMIObject MSAcpi_ThermalZoneTemperature -Namespace "root/wmi"
    $temps | Select-Object -Property InstanceName, @{n = "Temps C"; e = {(($_.currenttemperature / 10 - 273.15))}}
}
function Stress-CPU {
    Write-Host "Stressing CPU..." -ForegroundColor Red
    $result = 1; foreach ($number in 1..2147483647) {$result = $result * $number / 5614984; [math]::Round(($result - ($number / $result)))}; 
}
function du($dir = ".") { 
    get-childitem $dir | 
        % { $f = $_ ; 
        get-childitem -r $_.FullName | 
            measure-object -property length -sum | 
            select @{Name = "Name"; Expression = {$f}}, Sum}
}
function Get-LastShutDown() {
    Get-EventLog -LogName System | where {$_.EventID -eq 1074} | ft Index, TimeGenerated, Source, InstanceId, EventID
}
function Get-PCInfo() {
    $info = Get-CimInstance Win32_OperatingSystem

    Write-Host "`nOS Version: `t$($info.Caption)" -ForegroundColor Red
    Write-Host "Build Number: `t$($info.BuildNumber)" -ForegroundColor Red
    Write-Host "Build Type: `t$($info.BuildType)" -ForegroundColor Red
    Write-Host "Install Date: `t$($info.InstallDate)" -ForegroundColor Red
    Write-Host "Service Pack: `t$($info.ServicePackMajorVersion)" -ForegroundColor Red
    Write-Host "OS Arch: `t$($info.OSArchitecture)" -ForegroundColor Red
    Write-Host "Boot Device: `t$($info.BootDevice)" -ForegroundColor Red
    Write-Host "Host Name: `t$($info.CSName)`n" -ForegroundColor Red
}
function Start-LogUsersOff () {
    $explorerprocesses = @(Get-WmiObject -Query "Select * FROM Win32_Process WHERE Name='Explorer.exe'" -ErrorAction SilentlyContinue)
    If ($explorerprocesses.Count -eq 0) {
        "No explorer process found / Nobody interactively logged on"
    }
    Else {
        ForEach ($i in $explorerprocesses) {
            $Username = $i.GetOwner().User
            logoff.exe $Username
            Write-Host "$Username logged out"
        }
    }
}

function Get-ProductInfo () {
    param ($targets = ".")
    $hklm = 2147483650
    $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
    $regValue = "DigitalProductId4"
    Foreach ($target in $targets) {
        $productKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$target\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm, $regPath, $regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B", "C", "D", "E", "F", "G", "H", "J", "K", "M", "P", "Q", "R", "T", "V", "W", "X", "Y", "2", "3", "4", "5", "6", "7", "8", "9"
        ## decrypt base24 encoded binary data to characters. 
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
            }
            $productKey = $charsArray[$k] + $productKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $productKey = "-" + $productKey
            }
        }
        $win32os = Get-WmiObject Win32_OperatingSystem -computer $target
        $obj = New-Object Object
        $obj | Add-Member Noteproperty Computer -value $target
        $obj | Add-Member Noteproperty Caption -value $win32os.Caption
        $obj | Add-Member Noteproperty CSDVersion -value $win32os.CSDVersion
        $obj | Add-Member Noteproperty OSArch -value $win32os.OSArchitecture
        $obj | Add-Member Noteproperty BuildNumber -value $win32os.BuildNumber
        $obj | Add-Member Noteproperty RegisteredTo -value $win32os.RegisteredUser
        $obj | Add-Member Noteproperty ProductID -value $win32os.SerialNumber
        $obj | Add-Member Noteproperty ProductKey -value $productkey
        $obj
    }
}

# Internet Functions - Functinos that require internet access
function Git-Push { 
    $files = Read-Host -Prompt 'File / Files to add_> '
    $commit = Read-host -Prompt 'Commit comment_> '
    $origin = Read-Host -Prompt 'Origin (Normally just origin) _> '

    git add $files
    git commit -m $commit
    git push -f $origin master

}
function Get-ExtIp {
    $url = "http://icanhazip.com"
    $WClient = New-Object System.Net.WebClient
    $IP = $WClient.DownloadString($url)
    banner $IP.Trim("`n")
}
Function WhoIs {
    # Stolen from somewhere...
    param (
        [Parameter(Mandatory = $True,
            HelpMessage = 'Please enter domain name (e.g. microsoft.com)')]
        [string]$domain
    )
    Write-Host "Connecting to Web Services URL..." -ForegroundColor Green
    try {
        #Retrieve the data from web service WSDL
        If ($whois = New-WebServiceProxy -uri "http://www.webservicex.net/whois.asmx/GetWhoIs") {Write-Host "Ok" -ForegroundColor Green}
        else {Write-Host "Error" -ForegroundColor Red}
        Write-Host "Gathering $domain data..." -ForegroundColor Green
        #Return the data
        (($whois.getwhois("=$domain")).Split("<<<")[0])
    }
    catch {
        Write-Host "Please enter valid domain name (e.g. microsoft.com)." -ForegroundColor Red
    }
}
function PingLog { 
    $destination = Read-Host "Specify destination host name or IP address "
    $logonly = "Y"
    $filelocation = $destination + "_ping_results.txt"

    Do {
        If ($logonly -notmatch '^[yn]$' ) { Write-Warning "Invalid Entry" }
        $logonly = Read-Host "Log to a file and Output to console? (Y/N) "
    } While ($logonly -notmatch '^[yn]$')

    Write-Host "Default ping vaules: -count 999999 -delay 2"

    if ($logonly -eq "y") { 
        write-host "Pinging host:"$destination". Log file location: "$filelocation -BackgroundColor Yellow -ForegroundColor Red
        test-connection $destination -count 999999 -delay 2 -Verbose | format-table @{n = 'TimeStamp'; e = {Get-Date}}, @{Expression = {$_.Address}; Label = 'Destination'}, IPV4aDDRESS, ResponseTime | tee-object -filepath $filelocation -Append 
    }

    else { 
        test-connection $destination -count 999999 -delay 2 -Verbose | format-table @{n = 'TimeStamp'; e = {Get-Date}}, @{Expression = {$_.Address}; Label = "Destination"}, IPV4aDDRESS, ResponseTime 
        
    }
}
function Test-Port($hostname, $port) {
    # This works no matter in which form we get $host - hostname or ip address
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($hostname) |
            select-object IPAddressToString -expandproperty  IPAddressToString
        if ($ip.GetType().Name -eq "Object[]") {
            #If we have several ip's for that address, let's take first one
            $ip = $ip[0]
        }
    }
    catch {
        Write-Host "Possibly $hostname is wrong hostname or IP"
        return
    }
    $t = New-Object Net.Sockets.TcpClient
    # We use Try\Catch to remove exception info from console if we can't connect
    try {
        $t.Connect($ip, $port)
    }
    catch {}

    if ($t.Connected) {
        $t.Close()
        $msg = "Port $port is operational"
    }
    else {
        $msg = "Port $port on $ip is closed, "
        $msg += "You may need to contact your IT team to open it. "
    }
    Write-Host $msg
}
function Check-Word () {
    # Using Oxford API to see if word is real or not
    param([string]$str)

    $data = Invoke-RestMethod https://od-api.oxforddictionaries.com/api/v1/entries/en/$str -Headers @{"app_id" = "c3e95f65"; "app_key" = "dc3939f9b3e8b0480fdd72730bfaaf39"}

    $data.results.lexicalEntries.entries.senses.definitions
}
function CheckSpelling () {
    Param([string]$str)
    $data = Invoke-RestMethod -Uri "https://api.cognitive.microsoft.com/bing/v5.0/spellcheck?text=$str&count=3&mkt=en-us" -Headers @{ "Ocp-Apim-Subscription-Key" = "5f4e2cbf3cd345548a3d79c0583f51cd" }

    $data.flaggedTokens.suggestions
}


# Work Functions - Functions made purely for AQQ 
function AQMarkup {
    # Supplier Price + $10  >  Markup Percent  >  Client price EX GST

    banner "AQQ Markup Script" 

    $supPrice = Read-Host -Prompt "`nSupplier Price _> " 
    $Markup = Read-Host -Prompt "Markup by _> "

    # Adding Shipping
    $supPrice = [float]$supPrice + 10

    # Converting $Markup to decimal
    [float]$Markup = [float]$Markup / 100

    # Getting Total (Supplier Price + 
    $Markup = [float]$supPrice * $Markup
    $Total = [float]$supPrice + [float]$Markup 

    Write-Host
    banner "Total price EX GST: $([math]::Round($Total,2))"


}
function Get-PhoneMsg {
    param ([string]$date)

    $folder = 'E:\PhoneCalls'
    cd $folder


    foreach ($x in (Get-ChildItem $folder -File).name) {
        if ($x.StartsWith($date)) {
            write-host  -foreground "red" "`n********************************************"
            write-host  -foreground "red" "**              $($x)                **"
            write-host  -foreground "red" "********************************************`n"

            cat $x
        }

    }
}
function Start-PhoneMsg {
    # Arguments if they want to be used
    param (
        [Parameter(Mandatory = $true)][string]$Shop,
        [Parameter(Mandatory = $true)][string]$Client,
        [Parameter(Mandatory = $true)][string]$Issue
    )

    cd $PSScriptRoot

    # Few random placeholders 
    $date = Get-Date 
    $spacer = '==========='
    $filename = (Get-Date -format "yyyy-MM-dd")


    #Get the actually message
    $TempFile = [System.IO.Path]::GetTempFileName()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    nano $TempFile
    $sw.Stop()

    # Log input
    $spacer + $date + $spacer | Add-Content "E:\PhoneCalls\$($filename)"
    "$($Client) called from $($Shop) regarding $($Issue)`n" | Add-Content "E:\PhoneCalls\$($filename)"
    Get-Content $TempFile | Add-Content "E:\PhoneCalls\$($filename)"
    "`nTime: $($sw.Elapsed)`n========== END ===========`n" | Add-Content "E:\PhoneCalls\$($filename)"

    # Deleting $TempFile
    Remove-Item $TempFile

}

function AQToner () {
    function roundNumber($int) {
        $newInt = [math]::Round($int, 2)
        $newInt
    }

    # Getting basic info from user
    $tonerYield = Read-Host -Prompt "Toner Page Yield _> "
    $tonersellPrice = Read-Host -Prompt "Sell price of toner _> "
    [datetime]$supDate = Read-Host -Prompt "Old Printer Supply Date (mm/dd/yyyy) _> "
    $totalprintPages = Read-Host -Prompt "Total Printed Pages _> "

    # Getting date span for life of printer
    $curDate = Get-Date
    $usedDays = New-TimeSpan -Start $supDate -End $curDate

    # Removing Weekends from .Days
    $usedworkingDays = ($usedDays.Days / 7) * 2
    $usedworkingDays = $usedDays.Days - $usedworkingDays

    # Average pages printed per day over lifetime of printer
    $avgpagesDay = $totalprintPages / $usedworkingDays

    # How many toners will be needed to handle the above average for a year
    $tonersYear = $tonerYield / $avgpagesDay
    $tonersYear = 365 / $tonersYear

    # Yearly cost of toners to support the above average
    $tonersCost = $tonersYear * $tonersellPrice

    # Rounding all numbers
    $tonersYear = roundNumber($tonersYear)
    $avgpagesDay = roundNumber($avgpagesDay)
    $usedworkingDays = roundNumber($usedworkingDays)
    $tonersCost = roundNumber($tonersCost)

    # Output
    Write-Host @"

Printer used for busniess days:                 $usedworkingDays
Printed average pages/day:                      $avgpagesDay
Toners/year:                                    $tonersYear

"@ -ForegroundColor Red

    Write-Host "This will cost you $ `b$tonersCost/year in toners." -ForegroundColor Red
}

# Profile Settings
function Prompt() {
    $(Get-Location).Path + " _> "
}
# Exporting all functions
Export-ModuleMember -Function *