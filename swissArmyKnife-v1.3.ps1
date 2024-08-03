$logfile = "\\server1\archive\psLogs\SAKLogs.txt"
$reportDir = "\\server2\thehub\information technology\_IT 2017-07-01\Archive\Users"
$host.ui.RawUI.WindowTitle = "Swiss Army Knife Menu"
$runAt = (Get-Date).ToString("yyyy/MM/dd HH:mm")

#todo - add correct options, build out menu builder dealy
$mainMenuOptions = # set menu options here, number then text of the option in quotes.
@'
    8,      "menu text for option 8"
    9,      "option 9"
    10,     "10th emnu"
    102,    "actual actual last one "
'@

class menuOption {

    [int]     $Option
    [string]  $Text

    menuOption([int]$Option, [string]$Text) {
        $this.Init(@{Option = $Option; Text = $Text })
    }
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }
}


function setFilter ([string]$option,[string]$type) { # builds out the commands for search params
 
    $ioAfter = $null

    if ( $option -eq 'remove' ) {
        $ioBefore = (Get-Date)    #grabs todays date        
    }
    if ( $option -eq 'pre' ) { 
        $ioBefore = (Get-Date -year $arcYr -month 1 -day 1).AddDays(-1)    #previous year Dec 31
    }
    if ( $option -eq 'all' -or $option -eq '2nd' ) {
        $ioBefore = (Get-Date -year $arcYr -month 12 -day 31)    #archive year dec 31
    }
    if ( $option -eq '1st'-or $option -eq 'all' ) {  
        $ioAfter = (Get-Date -year $arcYr -month 1 -day 1)    #archive year Jan 1
    }
    if ( $option -eq '1st' ) {
        $ioBefore = (Get-Date -year $arcYr -month 6 -day 30)    #archive year June 30
    }
    if ( $option -eq '2nd' ) {
        $ioAfter = (Get-Date -year $arcYr -month 7 -day 1)    #archive year July 1
    }
    if ( $null -ne $ioAfter ) {
        $ioAfter = $ioAfter.ToString("MM/dd/yyyy")
    }
    $ioBefore = $ioBefore.ToString("MM/dd/yyyy")    
    
    switch ( $type ) {        
        {@("search" , "delete") -contains $_ }
        {
            #search doesnt care about sender or receiver
            if ( $null -ne $ioAfter ) {
                $sendBuilder = "Sent:$ioAfter..$ioBefore"
            }
            else {
                $sendBuilder = "Sent:<=$ioBefore"
            }  
            $filter = $sendBuilder
        }
        "archive" {
            $recdBuilder = "( Received -le '$ioBefore' )"
            $sendBuilder = "( Sent -le '$ioBefore' )"
            if ( $null -ne $ioAfter ) {
                $recdBuilder = "(( Received -ge '$ioAfter' ) -and $recdBuilder)"
                $sendBuilder = "(( Sent -ge '$ioAfter' ) -and $sendBuilder)"
            }    
            $filter = "$recdBuilder -or $sendBuilder"        
        }

    }
    return [string]$filter
}

function getExchangeCommand ([string]$exOption, [string]$exUser, [string]$exFilter, [string]$exYr) {
    $mailboxToBeArchived = "$exUser@domain.com"
            
    switch ($exOption) {
        "search"{
            $exCommand = "Search-Mailbox -Identity $mailboxToBeArchived -SearchQuery $exFilter -TargetMailbox administrator -TargetFolder SearchAndDeleteLog -LogOnly -LogLevel Full"
            }

        "delete"{
            $exCommand = "Search-Mailbox -Identity $mailboxToBeArchived -SearchQuery $exFilter -DeleteContent -Force"
        }

        "archive"{
            $arcName = $exUser + $exOption + $exYr
            $arcPath = "`"\\server1\archive\Users PST archive\$exYr\$u-$comment$exYr.pst`""
            $exCommand = "New-MailboxExportRequest -Name $arcName -Mailbox $mailboxToBeArchived -ContentFilter `"$exFilter`" -Filepath $arcPath"
        }
    }

    return [string]$exCommand
}

function frmBars{ # dumb formatting, but this keeps it standard
    Write-Host "`n
 <=========================================================><=========================================================>
     ##############################################################################################################
 <=========================================================><=========================================================>

    "
}

function scriptHeader {
    frmBars
    Write-host "`n
                                             WW&TB Swiss Army Knife Script
    
                                                                             .:^
                                                       ^                    /   |
                                          '`.        /;/                   /   /
                                          \  \      /;/                   /   /
                                           \\ \    /;/                   / ///
                                            \\ \  /;/                   / ///
                                             \  \/_/___________________/   /
                                            `/                         \  /
                                            {  o       BASCO         o  }'
                                             \_________________________/

                               `"The formatting on the ASCII art isn't even correct!`"
                                                `"What could go wrong??!`"
    "
    frmBars
}

function collectUsers{ #collects comma separated usernames, breaks them apart and stores them as an array of strings
    $userArray = [System.Collections.ArrayList]::new()
    [string]$inString = Read-Host "
            Please enter the user name(s) - comma separated if there are multiple" # collects username(s)
    if ( $inString.Contains(',') ) {
        $userArray = $inString.split(',').Trim()
    } # formats multiple usernames
    else {
        $userArray += $inString
    } # adds a single username to the array
    return $userArray
}

function menuTitle([string]$title) {
    Clear-Host
    frmBars
    Write-Host "  |    $title    |"
}

#Main menu loop
$menuOptions = New-Object System.Collections.Generic.List[menuOption]
$mainMenuOptions | ConvertFrom-Csv -Header 'Option', 'Text' `
| ForEach-Object { $menuOptions.Add( [menuOption]::new($_.Option, $_.Text) ) }

Do {
    Clear-Host
    scriptHeader
    $menuOption = read-host "
        Select an Option:
        1: Get locked out users list
        2: Unlock A User
        3: Get AD groups for a user
        4: Connect remote exchange session
        5: Search, Archive or Delete mailboxes
        6: Exchange status checks
        7: Disconnect exchange session

        98: Relaunch a new instance of this script
        99: Open ISE

        111 to exit
         
         Selection "

    switch ($menuOption) {
        1 {
            search-adaccount -lockedout        
        }
        2 {
            $poorSoul = Read-Host "          What is the username"
            Unlock-ADAccount -Identity $poorSoul
        }
        3 {
            $nameuser = read-host "          Username? "
            $reportPath = $reportDir + "\" + $nameuser + "_AD_Groups.txt"
            get-aduser $nameuser -Properties memberof | Select-Object -ExpandProperty memberof >> $reportPath
            write-host "     The file is $reportPath"
            Pause
        }
        4 { #    Connects remote PS session to exchange server
            $UserCredential = Get-Credential
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://mwebmail.domain.com -Authentication Kerberos -Credential $UserCredential
            Import-PSSession $Session -DisableNameChecking -AllowClobber
        }
        5 { #exchange search / archive menu
            Do {
                menuTitle("Archive Menu")

                $usersToArchive = collectUsers
                $sora = Read-Host "
                    Search - Queries the exchange server for the size of the user's mailbox
                    Archive - Exports all mailbox items to the archive PST for the 
                    Delete - Deletes all content - must be run after an archive to remove the items from the user's mailbox

                      Search, Archive, Delete?" # collects action to perform
                $comment = Read-Host "
                    Remove - performs the selected action against the whole mailbox. Year will be ignored
                    Pre - action taken agains everything before the specified year
                    1st / 2nd - applies actions to the respective half of the year
                    All - applies to the whole year

                    Remove, Pre, 1st, 2nd, all?" # collects which part of the year or the whole thing
                $arcYr = Read-Host "
                    Which year? If you entered 'Remove', it automatically assumes all user emails" # collects which year


                if ($null -eq $arcYr) { $comment="remove" } # assumes remove if year is skipped

                $SorAFilter = setFilter $comment $sora
                $runAt = (Get-Date).ToString("yyyy/MM/dd HH:mm")

                foreach ( $u in $usersToArchive ){
                    $command = getExchangeCommand $sora $u $SorAFilter $arcYr
                    "  $runAt - We are running the following command on mail: $command"  |Out-File -filePath $logfile -NoClobber -Append
               
                    Invoke-Expression $command
                }

                $menuOption = Read-Host "
                Run more commands (Y), or back to the main menu?"

            } while ( @("Y","y","yes") -contains $menuOption)
        }
        6 {
            Do {
                menuTitle("Exchange DB Checks")
                $exchDB = Read-Host "
                    Gets top 10 users from the selection.
                    Which database would you like to check (1,2,3,4,5, all)?"

                    if ( @(1,2,3,4,5) -contains $exchDB) {
                        $command = "get-mailbox -database `"database$exchDB`" | get-mailboxstatistics | Sort-Object TotalItemSize -Descending | select -first 10 | ft DisplayName, @{label=`"TotalItemSize(MB)`";expression={`$_.TotalItemSize.Value.ToMB()}}"
                    }
                    else {
                        $command = "Get-MailboxStatistics -server `"Mail`" | Sort-Object TotalItemSize -Descending | select -first 10 | ft DisplayName, @{label=`"TotalItemSize(MB)`";expression={`$_.TotalItemSize.Value.ToMB()}}"
                    }
                       
                "  $runAt - We are running the following command on mail: $command"  |Out-File -filePath $logfile -NoClobber -Append
                Invoke-Expression $command
                $exchDB = ""
                $menuOption = Read-Host "
                    Run more commands (Y), or back to the main menu?"

            } while (@("Y","y","yes") -contains $menuOption)

        }
        7 {
            if ($null -ne $Session){
                Remove-PSSession $Session
            }
        }

        98 {
            #grabs the currently running script and path. all the backticks ( ` ) make the various quotes literal w/i the $relaunch variable
            $relaunch = "`"& `'$PSCommandPath`'`"" 
            Invoke-Expression "cmd /c start powershell -command $relaunch"
        }

        99 {
            ise
        }

    }
} While ($menuOption -ne 111)
if ($null -ne $Session){
    Remove-PSSession $Session
}