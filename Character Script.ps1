#EQDeets - Character Parsing - v1.0.3 - 10/13/2022
#Written by Vanifac - Discord Vanifac#0123
Remove-Variable * -ErrorAction SilentlyContinue
$PSDefaultParameterValues = @{ '*:Encoding' = 'utf8' }
#=========Variables (Change these)=========
$EQLogDir = "E:\Games\P99\Logs"
$CheckDelay = 5 #Seconds
#$LogSplitThreshhold = 10

#=========Variables (Do Not Change)=========
$CharInfoDir = "$EQLogDir\Character Info"
$ActiveDir = "$CharInfoDir\Active Character"
$ActiveTxt = "$ActiveDir\0. Active Character.txt"
$SaveTxt = "$CharInfoDir\1. Character List.txt"
$OFS = " - "
$TotalLogs = "Total Deaths"
$EventLogs = "Death", "Level"

#=========Creating Text Files=========
New-Item -ItemType Directory -Path $CharInfoDir -ErrorAction 'silentlycontinue' | Out-Null
New-Item -ItemType Directory -Path $ActiveDir -ErrorAction 'silentlycontinue' | Out-Null
New-Item -ItemType File -Path $ActiveTxt -ErrorAction 'silentlycontinue' | Out-Null
New-Item -ItemType File -Path $SaveTxt -ErrorAction 'silentlycontinue' | Out-Null

Foreach ($Total in $TotalLogs) {
    $TotalTest = Test-Path "$CharInfoDir\$Total.txt"
    if ( -not $TotalTest ){0 | Out-File "$CharInfoDir\$Total.txt"}}

#=========Script=========
Clear-Host
Write-Host "====================================="
Do {
    Do {
        #===Stat Variables===
        $Stats = "Name", "Server", "Race", "Class", "Guild", "Zone", "Level", "Deaths", "Close Calls", "Killing Blows", "Crafts", "Failed Crafts"
        $Char = @("Soandso", "Orange", "Bunny Girl", "Senpai", "<Level Up>", "The IronGuard Confine", "1", 0, 0, 0, 0, 0)
        $NChar = @("Soandso", "Orange", "Bunny Girl", "Senpai", "<Level Up>", "The IronGuard Confine", "1", 0, 0, 0, 0, 0)
        $Nam = 0; $Ser = 1; $Rac = 2; $Cla = 3; $Gui = 4; $Zon = 5; $Lev = 6; $Dea = 7; $Clo = 8; $Kil = 9; $Cra = 10; $Fai = 11
        $CharCount = (($Char.Count) - 1)
        #===Cash Tracking===
        $CoinTrack = "Total Coin", "Looted Coin", "Vendor Coin", "Quest Coin", "Traded Coin"
        $Coin = @(0, 0, 0, 0, 0)
        $NCoin = @(0, 0, 0, 0, 0)
        $TCo = 0; $LCo = 1; $VCo = 2; $QCo = 3; $TrCo = 4
        #===Line Variables for tracking log lines===
        $StatLine = [int[]]::new($stats.count)
        $L = 0; Foreach ($Stat in $Stats) { $StatLine[$L] = 1; $L++ }
        $NStatLine = [int[]]::new($stats.count)
        $L = 0; Foreach ($Stat in $Stats) { $NStatLine[$L] = 1; $L++ }
        $CoinLine = [int[]]::new($CoinTrack.count)
        $L = 0; Foreach ($Track in $CoinTrack) { $CoinLine[$L] = 1; $L++ }
        $NCoinLine = [int[]]::new($CoinTrack.count)
        $L = 0; Foreach ($Track in $CoinTrack) { $NCoinLine[$L] = 1; $L++ }

        #---Selecting Log---
        Write-Host "Selecting Active Log.."; Start-Sleep -Milliseconds 500
        $ActiveLogP = Get-ChildItem $EQLogDir\eqlog_*1999*.txt | Sort-Object LastWriteTime | Select-Object -last 1
        $ActiveLogN = Get-Item $ActiveLogP | select-object -ExpandProperty Name
        $Tr, ($Char[$Nam]), $Serv = "$ActiveLogN" -split '_'; ($Char[$Ser]) = ($Serv.Substring(5)).TrimEnd('.txt')
        if ( $ActiveLogP -contains "project1999") { $Char[$Ser] = "Blue" }
        #$FirstLine =  (Get-Content $ActiveLogP | Measure-Object).Count

        #---Loading Saved Character---
        $NameSer = ($Char[$Nam, $Ser] -Join " - ")
        Write-Host "Loading Saved Character.."; Start-Sleep -Milliseconds 500
        $Loaded = (Select-String $SaveTxt -Pattern $NameSer -Simplematch).line

        #---Creating New Character---
        $WhoCheck = "] " + $Char[$Nam] + " ("
        if ($null -eq $Loaded) {
            Write-Host "No Save Found..`nCreating New Character Save.."; Start-Sleep -Seconds 1
            Do {
                $NewSave = Select-String $ActiveLogP -Pattern $WhoCheck -Simplematch | Select-Object -Last 1 | select-object -ExpandProperty Line
                if ($null -eq $NewSave) {
                    Write-Host "Waiting for a /Who message to populate character info for"$Char[0]", ensure logging is enabled in game (/log).." -ForegroundColor Red 
                    Start-Sleep -Seconds 5 
                }
            }
            While ($Null -eq $NewSave)

            $Race1 = $NewSave.LastIndexOf('(') + 1; $Ace = ($NewSave.LastIndexOf(')') - $Race1); $Char[$Rac] = $NewSave.Substring($Race1, $Ace).Trim()
            $Clas = $NewSave.LastIndexOf('[') + 3; $Lass = ($NewSave.LastIndexOf(']') - $Clas); $Char[$Cla] = $NewSave.Substring($Clas, $Lass).Trim()
            $Char[$Lev] = $NewSave.Substring($NewSave.LastIndexOf('[') + 1, 2).Trim()
            $WhoZone = Select-String $ActiveLogP -Pattern '(There are .* players in )' | Select-Object -Last 1 
            if ($null -eq $WhoZone) { $WhoZone = Select-String $ActiveLogP -Pattern 'There is 1 player in ' | Select-Object -Last 1 }
            $Char[$Zon] = ($WhoZone.Line.Substring($WhoZone.Line.LastIndexOf(" in ") + 4)).TrimEnd('.')
            if ($NewSave -like "*<*") { 
                $Guil = $NewSave.LastIndexOf('<') + 1; $Uild = ($NewSave.LastIndexOf('>') - $Guil); $Char[$Gui] = $NewSave.Substring($Guil, $Uild).Trim() 
            }
            Else { $Char[$Gui] = "Guildless" }

            Start-Sleep -seconds 1
            $Loaded = [String]$Char[0..5]
            Add-Content $SaveTxt $Loaded
        }
    
        #===Post-Load Variables===
        $Loaded | Out-File $ActiveTxt -Force
        $c = 0; $Loaded -Split " - " | Foreach-Object { $Char[$c] = $_; $c++ }
        
        #===Build File Structure===
        $NameDir = "$CharInfoDir\$NameSer"
        $DataDir = "$NameDir\Data"
        $Dirs = $CharInfoDir, $ActiveDir, $NameDir, $DataDir
        Foreach ($Dir in $Dirs) {
            $DirTest = Test-Path $Dir
            if (-Not $DirTest) {
                New-Item -ItemType Directory -Path $Dir | Out-Null
                Write-Host "Creating $Dir.."
                Start-Sleep -Seconds 1
                if ( $Dir -eq "$DataDir" ) { 
                    Write-Host "Creating StatLine.csv.."
                    $StatLine | Out-File "$DataDir\StatLine.csv"
                    Start-Sleep -Milliseconds 125
                    Write-Host "Creating CoinLine.csv.."
                    $CoinLine | Out-File "$DataDir\CoinLine.csv"
                    Start-Sleep -Milliseconds 125
                    Write-Host "Creating Copper.csv.."
                    $Coin | Out-File "$DataDir\Copper.csv"
                    Start-Sleep -Milliseconds 125
                }
            }
        }
        $n = 0; Foreach ($Stat in $Stats) {
            $StatTest = Test-Path "$NameDir\*$Stat.txt"
            if (-Not $StatTest) {
                Write-Host "Creating $Stat.."
                $Char[$n] | Out-File "$NameDir\$n. $Stat.txt" -Force
                Start-Sleep -Milliseconds 125
            }
            $n++
        }
        Foreach ( $log in $EventLogs ) {
            $LogTest = Test-Path "$NameDir\${log}Log.txt"
            if ( -Not $LogTest ) {
                Write-Host "Creating your $log Log.."
                New-Item -ItemType File -Path "$NameDir\${log}Log.txt" -Erroraction 'silentlycontinue' | Out-Null
            }
        }

        #---Load Stats/Lines---
        $Coin = (Get-Content "$DataDir\Copper.csv")
        $TotalDeaths = [Int](Get-Content "$CharInfoDir\Total Deaths.txt")
        $n = 6; Foreach ( $Stat in $Stats[6..($CharCount)]) {
            $Char[$n] = Get-Content "$NameDir\$n. $Stat.txt"
            $n++
        }

        #---Load Lines / Verify/Modify Length of Line Arrays---
        $Statline = Get-Content "$DataDir\Statline.csv"
        if ( $Stats.count -ne $Statline.count ) {
            $MissingLines = $Stats.count - $Statline.count
            for ($SL = 0; $SL -lt $MissingLines; $SL++) { 
                $Statline += 1
            }
            $Statline | Out-File "$DataDir\Statline.csv" -Force
        }
        $NStatline = Get-Content "$DataDir\Statline.csv"

        $Coinline = Get-Content "$DataDir\Coinline.csv"
        if ( $Coin.count -ne $Coinline.count ) {
            $MissingCoins = $Coin.count - $Coinline.count
            for ($SL = 0; $SL -lt $MissingCoins; $SL++) { 
                $Coinline += 1
            }
            $Coinline | Out-File "$DataDir\Statline.csv" -Force
        }
        $NCoinline = Get-Content "$DataDir\Coinline.csv"

        #---.csv File---
        $CSVTest = Test-Path "$NameDir\Character Info.csv"
        if ( -Not $CSVTest ) {
            New-Item -ItemType File -Path "$NameDir\Character Info.csv" -Erroraction 'silentlycontinue' | Out-Null
            $StatsCSV = '"' + ($Stats -Join '","') + '"'
            Add-Content -Path "$NameDir\Character Info.csv" $StatsCSV
            $CharCSV = '"' + ($Char -Join '","') + '"'
            Add-Content -Path "$NameDir\Character Info.csv" $CharCSV
        }
        $StatsCSV = (Select-String "$NameDir\Character Info.csv" -Pattern '("Name","Server")').Line
        if ( $StatsCSV -ne ('"' + ($Stats -Join '","') + '"') ) {
            (Get-Content "$NameDir\Character Info.csv").Replace($StatsCSV, ('"' + ($Stats -Join '","') + '"')) | Set-Content "$NameDir\Character Info.csv"
        }

        Write-Host $Char[$Nam]"is on"$Char[$Ser]"Server and is a Level"$Char[$Lev] $Char[$Rac] $Char[$Cla]
        $c = 0; $Char | Foreach-Object { $NChar[$c] = $_; $c++ }
        Start-Sleep -Seconds 3
        $CharCSV = Get-Content "$NameDir\Character Info.csv" | Select-Object -Last 1

        #---Setting Active Folder---
        $ActiveCSVTest = Test-Path "$ActiveDir\Character Info.csv"
        if ( -Not $ActiveCSVTest ) {
            Remove-Item "$ActiveDir\Character Info.csv" -Erroraction 'silentlycontinue'
            Start-Sleep -Seconds 5
        }
        Foreach ( $Stat in $Stats ) {
            Remove-Item "$ActiveDir\*. $Stat*"
            Copy-Item "$NameDir\*. $Stat*" $ActiveDir\ -Force
        }
        Copy-Item "$NameDir\Character Info.csv" $ActiveDir\
        
        #===Loop Parsing===
        Write-Host "Parsing.."
        Do {
            #---/Who Parse---
            $Who = Select-String $ActiveLogP -Pattern $WhoCheck -SimpleMatch | Select-Object -Last 1
            if ($null -ne $who) {
                $Clas = $Who.Line.LastIndexOf('[') + 3; $Lass = ($Who.Line.LastIndexOf(']') - $Clas); $NChar[$Cla] = $Who.Line.Substring($Clas, $Lass).Trim()
                #$NChar[$Lev] = $Who.Line.Substring($Who.Line.LastIndexOf('[') + 1, 2).Trim() - Ruins Ding parsing, shouldn't need it anyways.
                if ( $Who.Line -Like "*<*>" ) {
                    $Guil = $Who.Line.LastIndexOf('<') + 1; $Uild = ($Who.Line.LastIndexOf('>') - $Guil); $NChar[$Gui] = $Who.Line.Substring($Guil, $Uild).Trim()
                }
                Else { $NChar[$Gui] = "Guildless" }
            }
            
            #---Zoning Parse---
            $Zoning = Select-String $ActiveLogP -Pattern 'You have entered ' -SimpleMatch | Select-Object -Last 1
            $NChar[$Zon] = ($Zoning.Line.Substring($Zoning.Line.LastIndexOf('entered') + 8)).TrimEnd('.')
            
            #---Misc Parsing---
            #--(De)Ding Parse--
            $Dings = Select-String $ActiveLogP -Pattern "You have gained a level! Welcome to level " -SimpleMatch | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Lev]) }
            if ( $null -ne $Dings) {
                $Ding = $Dings | Select-Object -Last 1
                $NChar[$Lev] = ($Ding.Line.Substring($Ding.Line.LastIndexOf(' level ') + 7, 2).TrimEnd('!'))
                Write-Host "Ding! Welcome to level"$NChar[$Lev] -ForegroundColor Green
                If ( ([Int]$StatLine[$Lev]) -eq 1) {
                    Foreach ( $inst in $Dings ) {
                        Add-Content "$NameDir\LevelLog.txt" $Inst.line
                    }
                }
                Else { Add-Content "$NameDir\LevelLog.txt" (($Ding.line) + " - " + $NChar[$Zon]) }
                $NStatLine[$Lev] = $Ding.LineNumber + 1
            }

            #--Deaths--
            $Dying = Select-String $ActiveLogP -Pattern "You have been slain by " -SimpleMatch | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Dea]) }
            if ( $null -ne $Dying ) {
                [Int]$NChar[$Dea] += $Dying.count
                $NStatLine[$Dea] = ($Dying.LineNumber | Select-Object -last 1)
                If ( ([Int]($StatLine[$Dea])) -eq 1) {
                    Add-Content "$NameDir\DeathLog.txt" $Dying.line
                }Else { Add-Content "$NameDir\DeathLog.txt" (($Dying.line) + " - " + $NChar[$Zon]) }
                $Killedby = Select-String $ActiveLogP -Pattern "You have been slain by " -SimpleMatch | Select-Object -Last 1
                $Killer = $Killedby.Line.Substring($Killedby.Line.LastIndexOf(' by ') + 4 ).TrimEnd('!')
                Write-Host "You have died!"$Killer" probably really enjoyed that.. weirdo." -ForegroundColor Red
                #--Update Total deaths--
                [Int]$TotalDeaths += [Int]$Dying.Count
                $TotalDeaths | Out-File "$CharInfoDir\Total Deaths.txt"
            }

            #--Close Calls--
            #$CloseCall = Select-String $ActiveLogP -Pattern '(You have been knocked unconscious!)' -Context | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Kil]) }
            #if( $Closecall.count -gt $NChar[$Dea] ){
            #    $Shotta = Select-String $ActiveLogP -Pattern "You have been knocked unconscious!"  -SimpleMatch | Select-Object -Last 1
            #    Write-Host "Close call!"$Shotta" must have been reeeaaal tired of your rapping.." -ForegroundColor Red
            #    $NChar[$Dea] = [String]([Int]$NChar[$Dea] + [Int]$Dying.count)
            #}  

            #--Final Blows--
            $Kills = Select-String $ActiveLogP -Pattern '(You have slain a .*!)' | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Kil]) }
            if ( $null -ne $Kills ) {
                [Int]$NChar[$Kil] += $Kills.Count
                $NStatLine[$Kil] = ($Kills.LineNumber | Select-Object -Last 1)
            }

            #--Crafts--
            $Crafts = Select-String $ActiveLogP -Pattern '(You have fashioned the items together to create something new!)' | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Cra]) }
            if ( $null -ne $Crafts ) {
                [Int]$NChar[$Cra] += $Crafts.Count
                $NStatLine[$Cra] = ($Crafts.LineNumber | Select-Object -Last 1)
            }
 
            #--Failed Crafts--
            $Fails = Select-String $ActiveLogP -Pattern '(You lacked the skills to fashion the items together)' | Where-Object { $_.LineNumber -GT ([Int]$StatLine[$Fai]) }
            if ( $null -ne $Fails ) {
                [Int]$NChar[$Fai] += $Fails.Count
                $NStatLine[$Fai] = ($Fails.LineNumber | Select-Object -Last 1)
            }

            #---Coin---
            #--Looted Coin--
            $LootedCoin = Select-String $ActiveLogP -Pattern '(You receive .* from the corpse)' | Where-Object { $_.LineNumber -GT ([Int]$CoinLine[$LCo]) }
            if ( $null -ne $LootedCoin ) {
                ForEach ( $Corpse in $LootedCoin) {
                    if ($Corpse.Line -like "*copper*") {$NCoin[$LCo] += [Int]($Corpse.Line.Substring($Corpse.Line.IndexOf("copper") - 3, 3).Trim())}
                    if ($Corpse.Line -like "*silver*") {$NCoin[$LCo] += [Int]($Corpse.Line.Substring($Corpse.Line.IndexOf("silver") - 3, 3).Trim()) * 10}
                    if ($Corpse.Line -like "*gold*") {$NCoin[$LCo] += [Int]($Corpse.Line.Substring($Corpse.Line.IndexOf("gold") - 3, 3).Trim()) * 100}
                    if ($Corpse.Line -like "*platinum*") {$NCoin[$LCo] += [Int]($Corpse.Line.Substring($Corpse.Line.IndexOf("platinum") - 3, 3).Trim()) * 1000}
                    $CoinLine[$LCo] = [Int]$Corpse.LineNumber
                }
            }
            #---Vendored Coin---
            $VendorCoin = Select-String $ActiveLogP -Pattern '(You receive .* from .* for the )' | Where-Object { $_.LineNumber -GT ([Int]$CoinLine[$VCo]) }
            if ( $null -ne $VendorCoin ) {
                ForEach ( $Sale in $VendorCoin) {
                    if ($Sale.Line -like "*copper*") {$NCoin[$VCo] += [Int]($Sale.Line.Substring($Sale.Line.IndexOf("copper") - 3, 3).Trim())}
                    if ($Sale.Line -like "*silver*") {$NCoin[$VCo] += [Int]($Sale.Line.Substring($Sale.Line.IndexOf("silver") - 3, 3).Trim()) * 10}
                    if ($Sale.Line -like "*gold*") {$NCoin[$VCo] += [Int]($Sale.Line.Substring($Sale.Line.IndexOf("gold") - 3, 3).Trim()) * 100}
                    if ($Sale.Line -like "*platinum*") {$NCoin[$VCo] += [Int]($Sale.Line.Substring($Sale.Line.IndexOf("platinum") - 3, 3).Trim()) * 1000}
                    $CoinLine[$VCo] = [Int]$Sale.LineNumber
                }
            }
            
            #---Quest Coin---
            $QuestCoin = Select-String $ActiveLogP -Pattern '(You receive .* pieces.)' | Where-Object { $_.LineNumber -GT ([Int]$CoinLine[$QCo]) }
            if ( $null -ne $QuestCoin ) {
                ForEach ( $Quest in $QuestCoin) {
                    if ($Quest.Line -like "*copper*") {$NCoin[$QCo] += [Int]($Quest.Line.Substring($Quest.Line.IndexOf("copper") - 3, 3).Trim())}
                    if ($Quest.Line -like "*silver*") {$NCoin[$QCo] += [Int]($Quest.Line.Substring($Quest.Line.IndexOf("silver") - 3, 3).Trim()) * 10}
                    if ($Quest.Line -like "*gold*") {$NCoin[$QCo] += [Int]($Quest.Line.Substring($Quest.Line.IndexOf("gold") - 3, 3).Trim()) * 100}
                    if ($Quest.Line -like "*platinum*") {$NCoin[$QCo] += [Int]($Quest.Line.Substring($Quest.Line.IndexOf("platinum") - 3, 3).Trim()) * 1000}
                    $CoinLine[$QCo] = [Int]$Quest.LineNumber
                }
            }

            #---Trade Coin---
            #$TradeCoin

            #---Tell Parsing---
            $WikiSearch = Select-String $ActiveLogP -Pattern '(EQD.Wiki)' | Where-Object { $_.LineNumber -GT ([Int]$Statline[$Nam]) } | Select-Object -Last 3
                if ( $null -ne $WikiSearch ) {
                    Foreach ( $Search in $WikiSearch ){
                        If ( $Search.line -like "*EQD.Wiki is*" ) {
                            Write-host "Opening P1999 wiki.."
                            Start-Process https://wiki.project1999.com/ }
                        Elseif ( $Search.line -like "*EQD.Wiki:Zone *" ) {
                            Write-host "Opening P1999 wiki page for"$NChar[$Zon]".."
                            $URL = ($NChar[$Zon].replace( " ", "_" ))
                            Start-Process "https://wiki.project1999.com/$URL" }
                        Elseif ( $Search.line -like "*EQD.Wiki:Class *" ) {
                            Write-host "Opening P1999 wiki page for"$NChar[$Cla]".."
                            $URL = ($NChar[$Cla].replace( " ", "_" ))
                            Start-Process "https://wiki.project1999.com/$URL" }
                        Elseif ( $Search.line -like "*EQD.Wiki:*" ) {
                            $Que = ($Search.Line.LastIndexOf(':') +1); $ery = ($Search.Line.LastIndexOf(' is not') - $Que)
                            $Query = $Search.Line.Substring( $Que, $ery )
                            Write-host "Opening P1999 wiki page for $Query.."
                            Start-Process "https://wiki.project1999.com/index.php?title=Special%253ASearch&search=$Query&go=Go" }
                        $NStatLine[$Nam] = ($WikiSearch.LineNumber | Select-Object -Last 1)
                    }
                }

            #===Saving===
            if (([String]$Char + " - " + "0 - 0 - 0 - 0 - 0" + " - " + [String]$Statline) -ne ([String]$NChar + " - " + [String]$NCoin + " - " + [String]$NStatline)) {
                #---Updating Character Stats---
                if ([String]$Char -ne [String]$NChar) {
                    #---Updating Character List.txt---
                    (Get-Content -Path $SaveTxt).Replace($Loaded, [String]$NChar[0..5]) | Set-Content -Path $SaveTxt
                    #---Update Save .txts---
                    $c = 4; $NChar[4..$CharCount] | ForEach-Object {
                        $StatTxt = [String]("$NameDir\" + $c + ". " + $Stats[$c] + ".txt")
                        $Saved = [String](Get-Content -Path $StatTxt)
                        if ( [String]$Nchar[$c] -ne $Saved ) {
                            Write-Host $Stats[$c]":"$NChar[$c]
                            [String]$_ | Out-file $StatTxt -Force
                        }
                        $c++
                    }
                    #---Update Stat .csv---
                    $NCharCSV = '"' + ($NChar -Join '","') + '"'
                    $OldCSV = Get-Content "$NameDir\Character Info.csv" | Select-Object -Last 1
                    (Get-Content -Path "$NameDir\Character Info.csv").Replace($OldCSV, $NCharCSV) | Set-Content -Path "$NameDir\Character Info.csv"
                    (Get-Content -Path "$ActiveDir\Character Info.csv").Replace($OldCSV, $NCharCSV) | Set-Content -Path "$ActiveDir\Character Info.csv"
                    
                    #---Reloading Baseline Variables---
                    $Loaded = (Select-String $SaveTxt -Pattern $NameSer -Simplematch).line
                    $c = 0; $NChar | Foreach-Object { $Char[$c] = $_; $c++ }
                    $CharCSV = Get-Content "$NameDir\Character Info.csv" | Select-Object -Last 1
                }

                #---Updating Coin Count---
                if ( [String]$NCoin -ne "0 - 0 - 0 - 0 - 0" ) {
                    
                    $c = 0; $NCoin | Foreach-Object { [Int]$Coin[$c] += $_; $c++ }
                    $c = 0; $NCoin | Foreach-Object { [Int]$Coin[$TCo] += $_; $c++ }
                    $Coin | Out-file "$DataDir\Copper.csv"

                    #---Update Coin .csv---
                    $CoinTXT = @() ; $c = 0
                    Foreach ( $Source in $Coin) {
                        $Length = ([String]$Source).Length
                        $CP = ([String]$Source).Substring($Length - 1)
                        if ( $Length -ge 2) { $SP = ([String]$Source).Substring($Length - 2, 1) }Else { $SP = 0 }
                        if ( $Length -ge 3) { $GP = ([String]$Source).Substring($Length - 3, 1) }Else { $GP = 0 }
                        if ( $Length -ge 4) { $PP = ([String]$Source).Substring(0, $Length - 3) }Else { $PP = 0 }
                        $CoinTXT += @( $CoinTrack[$c] + ": $PP`PP $GP`GP $SP`SP $CP`CP")
                        $c++
                    }
                    $CoinTXT | Out-file "$NameDir\Cash.txt" -Force

                    $CoinCSV = @() ; $c = 0
                    Foreach ( $Source in $Coin) {
                        $Length = ([String]$Source).Length
                        $CP = ([String]$Source).Substring($Length - 1)
                        if ( $Length -ge 2) { $SP = ([String]$Source).Substring($Length - 2, 1) }Else { $SP = 0 }
                        if ( $Length -ge 3) { $GP = ([String]$Source).Substring($Length - 3, 1) }Else { $GP = 0 }
                        if ( $Length -ge 4) { $PP = ([String]$Source).Substring(0, $Length - 3) }Else { $PP = 0 }
                        #$CoinCSV += @( $CoinTrack[$c]+'":","$PP`PP","$GP`GP","$SP`SP","$CP`CP"')
                        $CoinCSV += @( $CoinTrack[$c] + ":, $PP`PP,$GP`GP,$SP`SP,$CP`CP")
                        $c++
                    }
                    $CoinCSV | Out-file "$NameDir\Cash.csv" -Force
                   
                    $Coinline | Out-file "$DataDir\Coinline.csv" -Force
                    #---Update Baseline Variables---

                    $NCoin = @(0, 0, 0, 0, 0)
                }
                
                #---Updating line tracking---
                if ( [String]$NStatline -ne [String]$Statline ){
                    $c = 0; $NStatline | Foreach-Object { $StatLine[$c] = $_; $c++ }
                    $NStatline | Out-file "$DataDir\Statline.csv"
                }

                #---Update Active Folder---
                Foreach ( $Stat in $Stats ) {
                    Copy-Item "$NameDir\*. $Stat*" $ActiveDir\ -Force
                }
            }

            #---Pause---
            for ($a = 0; $a -lt $checkdelay; $a++) { 
                Write-Host ("." * ($a / 3))`r -NoNewline
                Start-Sleep -Seconds 1
            }
            Write-Host "                      "`r -NoNewline

            #---End Checks---
            $EQStatus = Get-Process -Name eqgame -ErrorAction SilentlyContinue
            if ($null -eq $EQStatus) { 
                Write-host "Everquest has closed. Closing Script." -ForegroundColor Red
                Return
            }
            $CheckActive = Get-ChildItem $EQLogDir\eqlog_*1999*.txt | Sort-Object LastWriteTime | Select-Object -last 1 | select-object -ExpandProperty Name
        }While ( $ActiveLogN -eq $CheckActive )

        Write-Host $Char[0]"is no longer the active character. Updating character.."`n
        Start-Sleep -Seconds 15
    }While ( $null -ne $EQStatus )

    Write-host "Everquest has closed. Closing Script."
    Start-Sleep -Seconds 20

}While ( $null -ne $EQStatus )

Write-host "Everquest has closed. Closing Script." -ForegroundColor RedS