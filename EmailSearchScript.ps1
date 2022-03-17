########################
# Email Search Script  #
# Author: DBemrose     #
# Build: POC-Dev       #
# Build V: 1.0.0       #
########################
$Script:Version = "1.0.0"
$Script:Build = "DEV"
########################

Function StartComplianceSearch{
    #Create a new compliance search based on the variables provided
    New-ComplianceSearch -Name $script:SearchName -ExchangeLocation $script:InboxChecks -ContentMatchQuery "recipients:$script:SearchAddress"
    #Start the compliance search
    Start-ComplianceSearch -Identity $script:SearchName

    Do{
    Write-Host "Checking Search Status..."
    $script:SearchStatus = Get-ComplianceSearch -Identity $script:SearchName | Select-Object -ExpandProperty Status #-Property Status
    Write-Host $script:SearchStatus
    start-sleep 2
    cls
    }   while ($script:SearchStatus -notcontains "Completed")
    Read-Host 'Press any key to continue...'
}

Function CheckSearch{
    $script:SearchResultItems = Get-ComplianceSearch -Identity $script:SearchName | Select-Object -ExpandProperty Items #-Property Items
    Write-Host "Total Items Found:" $script:SearchResultItems
    Write-Host "Generating Output Preview"
    New-ComplianceSearchAction -SearchName $script:SearchName -Preview
    $script:ComplianceSearchName = $script:SearchName+"_preview"
    Do{
        Write-Host "Checking Preview Status..."
        $script:ComplianceSearchStatus = Get-ComplianceSearchAction $script:ComplianceSearchName | Select-Object -ExpandProperty Status #-Property Status
        Write-Host $script:ComplianceSearchStatus
        start-sleep 2
        cls
        }   while ($script:ComplianceSearchStatus -notcontains "Completed")
}

Function ExchangeLogin{
    #This imports the ExchangeOnlineManagement Module. This is assumed that the module is already installed but can be easily obtained or an auto install function can be created if needed
    Import-Module -Name ExchangeOnlineManagement
    #initiates a connection with no details so the technician can login as whoever they need
    Connect-IPPSSession
    MainMenu    
}

Function MainMenu{
    Clear-Host
    #Puts the version and build within the output
    Write-Host "Email Searcher v $Script:Version($Script:Build)"
    Write-Host "Main Menu:"
    #Login type is not currently used but is planned to be replaced with the email that was logged in with
    Write-Host "1. Login To Exchange" -NoNewline; Write-Host "($Script:LoginType)" -ForegroundColor Red
    Write-Host "2. Email Search Options"
    Write-Host "3. Miscellaneous" -NoNewline; Write-Host " (" -NoNewline -f Gray; Write-Host "Coming Soon" -NoNewline -f Red; Write-Host ")" -f Gray
    Write-Host "Q. Quit"
    #Waits for the user to provide a menu option
    $Script:MMResult = Read-Host "Option"
    #Based on the users output it will go to a specific menu
    Switch($Script:MMResult){
        1{ExchangeLogin}
        #Some basic logic to make sure the user has logged into the exchange account and loaded the needed commands. If it is not they will be taken to the login menu instead.
        2{If($Script:Login = $null){ExchangeLogin}Else{SelectionMenu}}
        4{MainMenu}
        default{MainMenu}
        "Q"{Exit}
    }
}


Function SetSearchName{
    $script:SearchName = Read-Host 'Specify the search name'
}

Function SetSearchAddress{
    $script:SearchAddress = Read-Host 'Specify the search address'
}

Function SetInboxParams{
    $script:InboxChecks = Read-Host 'Specify the inboxes to check'
}

Function ShowPreview{
    (Get-ComplianceSearchAction $script:ComplianceSearchName | Select-Object -ExpandProperty Results) -split ","
    Read-Host 'Press any key to continue...'
}

Function SelectionMenu{
    Clear-Host
    Write-Host "Selection Menu"
    If ([string]::IsNullOrEmpty($Script:SearchName)){Write-Host "1. Select Search Name"}Else{Write-Host "1. Select Search Name" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:SearchName -f green -NoNewline; Write-Host ")" -f gray}
    If ([string]::IsNullOrEmpty($Script:SearchAddress)){Write-Host "2. Select Search Address"}Else{Write-Host "2. Select Search Address" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:SearchAddress -f green -NoNewline; Write-Host ")" -f gray}
    If ([string]::IsNullOrEmpty($Script:SearchAddress)){Write-Host "3. Select Inboxes to search"}Else{Write-Host "3. Select Inboxes to search" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $script:InboxChecks -f green -NoNewline; Write-Host ")" -f gray}
    Write-Host "4. Start Search"
    Write-Host "5. Check the search"
    Write-Host "6. Show Preview"
    Write-Host "7. Main Menu"
    Write-Host "Q. Quit"
    $script:SMResult = Read-Host "Option"
    Switch($script:SMResult){
        1{SetSearchName; SelectionMenu}
        2{SetSearchAddress; SelectionMenu}
        3{SetInboxParams; SelectionMenu}
        4{StartComplianceSearch; SelectionMenu}
        5{CheckSearch; SelectionMenu}
        6{ShowPreview ; SelectionMenu}
        7{MainMenu}
        default{SelectionMenu}
        "Q"{Exit}

    }
}
MainMenu