$global:HostVar = $Host
$Branch = "main"
$Version = "v0.1.6"
$Title = @"
 ######\  ######\  ######\ ########\                             
##  __##\ \_##  _|##  __##\\__##  __|                            
## /  \__|  ## |  ## /  \__|  ## |                               
## |####\   ## |  \######\    ## |                               
## |\_## |  ## |   \____##\   ## |                               
## |  ## |  ## |  ##\   ## |  ## |                               
\######  |######\ \######  |  ## |                               
 \______/ \______| \______/   \__|   - Gist Intune Script Trigger

Version: $Version ($Branch) by https://x.com/MrWyss 
Source: https://github.com/MrWyss-MSFT/gist

2026-02-06: [NEW] Updated visual style of the menu
"@

$GistCatalog = @(
    [ordered]@{
        Name        = "Get-Win32AppsOrder"
        Category    = "Intune"
        Url         = "https://gist.githubusercontent.com/MrWyss-MSFT/c52f0a4567cba24e9ac8a38416a05bac/raw/d0f5f0e7c03f1607c193b0b4976296d3d34f5b05/Get-Win32AppsOrder.ps1"
        Description = "Reads the IntuneManagementExtension.log and returns a list of Win32Apps"
        Author      = "MrWyss-MSFT"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "Get-CoMgmtWL"
        Category    = "CoMgmt"
        Url         = "https://gist.githubusercontent.com/MrWyss-MSFT/228ed9f3dd2b67077790a3bef98442d5/raw/51699848761c0625991ff1374f227f945e8f42cd/Get-CoMgmtWL.ps1"
        Description = "Returns if CoManagement is enabled and how the workloads are configured"
        Author      = "MrWyss-MSFT"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "Get-AutopilotAndESPProgress"
        Category    = "Autopilot"
        Url         = "https://gist.githubusercontent.com/MrWyss-MSFT/d511d655f55762233a1d442c24d584f6/raw/da47f0c6c31708ec7f83cd0d9188a7fe77116877/Get-AutopilotAndESPProgress.ps1"
        Description = "View the Autopilot and ESP Progress"
        Author      = "MrWyss-MSFT"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "Copy-AutoPilotHashToClipboard"
        Category    = "Autopilot"
        Url         = "https://gist.githubusercontent.com/MrWyss/b83733e8378139af6032381359e00803/raw/7f87d17406e7e120be25c6c35fe1fc154233f17e/Copy-AutoPilotHashToClipboard.ps1"
        Description = "Copy the Autopilot Hash to the clipboard"
        Author      = "MrWyss"
        Elevation   = $true
    }
    [ordered] @{
        Name        = "New-IntuneRegistryFavorites"
        Category    = "Intune"
        Url         = "https://gist.githubusercontent.com/MrWyss-MSFT/500b2270b0b23b2fdc9ddc78092c355d/raw/ed9ce7714b335cdb52c2de97be4842dcbd7db359/New-IntuneRegistryFavorites.ps1"
        Description = "Create Registry Favorites for Intune and Autopilot"
        Author      = "MrWyss-MSFT"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "BITS-Monitor"
        Category    = "Windows"
        Url         = "https://raw.githubusercontent.com/jonasatgit/scriptrepo/03e54bd2a07ee8205831ee9698a4c7bf21317f52/General/BITS-Monitor.ps1"
        Description = "Monitor BITS transfer jobs. Will refresh every five seconds"
        Author      = "Jonasatgit"
        Elevation   = $true
    }
    [ordered] @{
        Name        = "DO-Monitor"
        Category    = "Windows"
        Url         = "https://raw.githubusercontent.com/itwaman/myPSscripts/e56133ad74b4ec49e94ecfcfdc20dabe67aed410/DOStatusMonitor.ps1"
        Description = "Monitor Delivery Optimization jobs. Will refresh every two seconds"
        Author      = "itwaman"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "Get-PerfCounterList"
        Category    = "Windows"
        Url         = "https://raw.githubusercontent.com/jonasatgit/scriptrepo/03e54bd2a07ee8205831ee9698a4c7bf21317f52/General/Get-PerfCounterList.ps1"
        Description = "Script to list all available performance counters via GridView"
        Author      = "Jonasatgit"
        Elevation   = $false
    }
    [ordered] @{
        Name        = "Set-WindowsCorporateIdentifiers"
        Category    = "Intune"
        Url         = "https://gist.githubusercontent.com/MrWyss-MSFT/a9401456c32a2ccd3cd1f09b5f2d9a2c/raw/5a54f5825f5acbc8ab0b49d2741de314bc0b9b24/Set-WindowsCorporateIdentifiers.ps1"
        Description = "Upload the Windows Corporate Identifiers for the device to Intune"
        Author      = "MrWyss-MSFT"
        Elevation   = $false
    }    
    [ordered] @{
        Name        = "Show-SecureBootCerts"
        Category    = "Windows"
        Url         = "https://raw.githubusercontent.com/MrWyss-MSFT/SecureBootCerts/bdb72cb12927c674c95ac0a3538c2a9a4dafcdab/Show-SecureBootCerts.ps1"
        Description = "Retrieves Secure Boot certificates and related data"
        Author      = "MrWyss-MSFT"
        Elevation   = $true
    }
    #[ordered] @{
    #    Name        = ""
    #    Category    = ""
    #    Url         = ""
    #    Description = ""
    #    Author      = ""
    #    Elevation   = $false | $true
    #}
)

# Class to create a console menu
Class ConsoleMenu {
    [string]$Title
    [array]$Options
    [int]$MaxStringLength = 0
    [string[]]$ExcludeProperties
    [switch]$AddDevideLines
    [switch]$StopIfWrongWidth
    [bool]$cleared = $false
    [string]$selection = -1

    # Default constructor
    ConsoleMenu() { 
        $this.Init(@{}) 
    }
    # Convenience constructor from hashtable
    ConsoleMenu([hashtable]$Properties) { 
        $this.Init($Properties) 
    }
    # Common constructor for title and Options
    ConsoleMenu([string]$Title, [array]$Options) { 
        $this.Init(@{Title = $Title; Options = $Options }) 
    }
    # Shared initializer method
    [void] Init([hashtable]$Properties) {
        foreach ($Property in $Properties.Keys) {
            $this.$Property = $Properties.$Property
        }
    }
    # Helper function to split a long string into chunks of a given size
    [string[]] SplitLongString([string]$String, [int]$ChunkSize) {
        $regex = "(.{1,$ChunkSize})(?:\s|$)"
        $splitStrings = [regex]::Matches($string, $regex) | ForEach-Object { $_.Groups[1].Value }
        return $splitStrings
    }
    # Helper function to remove filtered properties from the output
    [array] SelectedProperties() {
        # exclude properties from output if they are in the $ExcludeProperties array and store the result in $selectedProperties
        if ($this.ExcludeProperties) {
            [array]$selectedProperties = $this.Options[0].Keys | ForEach-Object {
                if ($this.ExcludeProperties -notcontains $_) {
                    $_
                }
            }
        }
        else {
            # Nothing to exclude
            [array]$selectedProperties = $this.Options[0].Keys
        }
        return $selectedProperties
    }
    # Helper function to calculate the maximum length of each value and key in the hashtable and the title
    [hashtable] GetLengths() {
        # Calculate the maximum length of each value and key in case the key is longer than the value
        $lengths = @{ }
        foreach ($item in $this.Options) {
            foreach ($property in $item.Keys) {
                if ($this.SelectedProperties() -icontains $property) {
                    $valueLength = $item[$property].ToString().Length
                    $keyLength = $property.ToString().Length
                    # if key name is longer than value, use key length as value length
                    if ($keyLength -gt $valueLength) {
                        $valueLength = $keyLength
                    }
                    if ($lengths.ContainsKey($property)) {
                        if ($valueLength -gt $lengths[$property]) {
                            $lengths[$property] = $valueLength
                        }
                    }
                    else {
                        $lengths.Add($property, $valueLength)
                    }
                }
            }
        }
        return $lengths
    }
    # Helper function to calculate the maximum width of the title
    [int] GetTitleWidth() {
        # Calculate the maximum width of the title
        $titleWidth = ($this.Title -split "\r?\n" | Measure-Object -Maximum -Property Length).Maximum
        # Add some extra space for each added table character
        # for the outer characters plus a space
        $titleWidth += 3 
        return $titleWidth
    }
    # Helper function to calculate the maximum width of the table
    [int] GetMaxWidth([hashtable]$lengths, [array]$selectedProperties) {
        # Calculcate the maximum width using all selected properties
        $maxWidth = ($lengths.Values | Measure-Object -Sum).Sum
    
        # Add some extra space for each added table character
        # for the outer characters plus a space
        $maxWidth += 3 
        # 1 + for the "Nr" header and one for each selected property times three because of two spaces and one extra character
        $maxWidth = $maxWidth + (1 + $selectedProperties.count) * 3
        if ($this.GetTitleWidth() -gt $maxWidth) {
            $maxWidth = $this.GetTitleWidth()
        }
        return $maxWidth
    }
    # Helper function to limit the maximum length of each string in $lengths
    [hashtable] LimitStrings ([hashtable]$lengths, [array]$selectedProperties, [int]$MaxStringLength) {
        # if $MaxStringLength is set, we need to limit the maximum length of each string in $lengths
        foreach ($property in $selectedProperties) {
            if ($lengths[$property] -gt $MaxStringLength) {
                $lengths[$property] = $MaxStringLength
            }
        }
        return $lengths
    }
    # Helper function to build horizontal divider lines with proper junctions
    [string] BuildDividerLine([hashtable]$lengths, [array]$selectedProperties, [string]$LineType, [int]$maxWidth) {
        # LineType: "top", "header", "middle", "bottom", "titleEnd"
        $leftChar = switch ($LineType) {
            "top"      { [char]0x250C }  # ┌
            "titleEnd" { [char]0x251C }  # ├
            "header"   { [char]0x251C }  # ├
            "middle"   { [char]0x251C }  # ├
            "bottom"   { [char]0x2514 }  # └
        }
        $rightChar = switch ($LineType) {
            "top"      { [char]0x2510 }  # ┐
            "titleEnd" { [char]0x2524 }  # ┤
            "header"   { [char]0x2524 }  # ┤
            "middle"   { [char]0x2524 }  # ┤
            "bottom"   { [char]0x2518 }  # ┘
        }
        $junctionChar = switch ($LineType) {
            "top"      { [char]0x252C }  # ┬
            "titleEnd" { [char]0x252C }  # ┬ (connects title to columns)
            "header"   { [char]0x253C }  # ┼
            "middle"   { [char]0x253C }  # ┼
            "bottom"   { [char]0x2534 }  # ┴
        }
        
        $line = "$leftChar"
        # Nr column: width 6 (" Nr" + 3 spaces)
        $line += "$([char]0x2500)" * 6
        $line += "$junctionChar"
        
        # Add each property column
        foreach ($property in $selectedProperties) {
            if ($property -eq $selectedProperties[-1]) {
                # Last column: pad to maxWidth to match header/data row padding
                # Header does: remove separator, add ($maxWidth - length) + 1 spaces, add separator
                # So we need: ($maxWidth - $line.Length) + 1 dashes (no final junction)
                $remainingWidth = ($maxWidth - $line.Length) + 1
                $line += "$([char]0x2500)" * $remainingWidth
            } else {
                # Each column: 1 space + property content length
                $line += "$([char]0x2500)" * ($lengths[$property] + 1)
                $line += "$junctionChar"
            }
        }
        
        $line += "$rightChar"
        return $line
    }
    # Helper function to build simple divider line (no junctions, for title box)
    [string] BuildSimpleDividerLine([int]$width, [string]$LineType) {
        # LineType: "top", "bottom"
        $leftChar = if ($LineType -eq "top") { [char]0x250C } else { [char]0x251C }  # ┌ or ├
        $rightChar = if ($LineType -eq "top") { [char]0x2510 } else { [char]0x2524 }  # ┐ or ┤
        
        return "$leftChar" + "$([char]0x2500)" * $width + "$rightChar"
    }
    # Draws the menu
    [void] DrawMenu() {
        Clear-Host  # Clear screen for terminal resizing
        
        [array]$selectedProperties = $this.SelectedProperties()
        [hashtable] $lengths = $this.GetLengths()
        
        # Get current terminal width and calculate dynamic MaxStringLength
        $terminalWidth = $global:HostVar.UI.RawUI.BufferSize.Width
        if ($terminalWidth -le 0) { $terminalWidth = 120 }  # Fallback if can't detect
        
        # Calculate space used by fixed columns (Nr=6, Name=32, Category=10, Author=12, Elevation=15, borders/spaces=18)
        # Total fixed: ~93 characters. Remaining space goes to Description column
        $fixedColumnsWidth = 93
        $availableForDescription = $terminalWidth - $fixedColumnsWidth - 4  # 4 for some margin
        
        # Use dynamic description width, with min/max bounds
        $dynamicMaxLength = [Math]::Max(20, [Math]::Min($availableForDescription, 100))
        
        # Apply dynamic wrapping if MaxStringLength is set, otherwise use calculated value
        $effectiveMaxLength = if ($this.MaxStringLength -gt 0) { $dynamicMaxLength } else { 0 }
        
        if ($effectiveMaxLength -gt 0) {
            $lengths = $this.LimitStrings($lengths, $selectedProperties, $effectiveMaxLength)
        }
        $maxWidth = $this.GetMaxWidth($lengths, $selectedProperties)
   
        # create the menu with the $consoleMenu array
        $consoleMenu = @()
        $consoleMenu += $this.BuildSimpleDividerLine($maxWidth, "top")
    
        foreach ($titlePart in ($this.Title -split "\r?\n")) {
            $leftPad = 1  # Left-aligned with 1 space margin
            $rightPad = $maxWidth - $titlePart.Length - $leftPad
            $consoleMenu += "$([Char]0x2502)" + " " * $leftPad + $titlePart + " " * $rightPad + "$([Char]0x2502)"    
        }
        $consoleMenu += $this.BuildDividerLine($lengths, $selectedProperties, "titleEnd", $maxWidth)
        # now add the header using just the properties from $selectedProperties
        $header = "$([Char]0x2502)" + " Nr" + " " * (3) + "$([Char]0x2502)"
    
        foreach ($property in $selectedProperties) {
            $header += " " + $property + " " * ($lengths[$property] - $property.Length) + "$([Char]0x2502)"
            # Fix if title longer than all properties
            if ($property -eq $selectedProperties[-1]) {
                $header = $header.Substring(0, $header.Length - 1)
                $header += " " * (($maxWidth - $header.length) + 1) + "$([Char]0x2502)"
            }
        }
        $consoleMenu += $header
        $consoleMenu += $this.BuildDividerLine($lengths, $selectedProperties, "header", $maxWidth)
        # now add the items
        $i = 0
        foreach ($item in $this.Options) {
            $i++
            
            # First pass: split all properties that need wrapping
            $propertyValues = @{}
            $maxRows = 1
            foreach ($property in $selectedProperties) {
                if ($effectiveMaxLength -gt 0 -and ($item.$property).Length -gt $effectiveMaxLength) {
                    [array]$strList = $this.SplitLongString($item.$property, $effectiveMaxLength)
                    $propertyValues[$property] = $strList
                    if ($strList.Count -gt $maxRows) {
                        $maxRows = $strList.Count
                    }
                }
                else {
                    $propertyValues[$property] = @($item.$property.ToString())
                }
            }
            
            # Second pass: build each row with all columns
            for ($rowNum = 0; $rowNum -lt $maxRows; $rowNum++) {
                if ($rowNum -eq 0) {
                    $line = "$([Char]0x2502)" + " " + "$i" + " " * (5 - $i.ToString().Length) + "$([Char]0x2502)"
                }
                else {
                    $line = "$([Char]0x2502)" + "      " + "$([Char]0x2502)"
                }
                
                foreach ($property in $selectedProperties) {
                    if ($rowNum -lt $propertyValues[$property].Count) {
                        $value = $propertyValues[$property][$rowNum]
                    }
                    else {
                        $value = ""
                    }
                    $line += " " + $value + " " * ($lengths[$property] - $value.Length) + "$([Char]0x2502)"
                }
                
                # Fix if title longer than all properties
                $line = $line.Substring(0, $line.Length - 1)
                $line += " " * (($maxWidth - $line.length) + 1) + "$([Char]0x2502)"
                
                $consoleMenu += $line
            }

            # Add divider lines between items
            if ($this.AddDevideLines) {
                if ($i -lt $this.Options.Count) {
                    $consoleMenu += $this.BuildDividerLine($lengths, $selectedProperties, "middle", $maxWidth)
                }
            }
        }
        $consoleMenu += $this.BuildDividerLine($lengths, $selectedProperties, "bottom", $maxWidth)
        $consoleMenu += " "
      
        # test if the console window is wide enough to display the menu
        if (($global:HostVar.UI.RawUI.WindowSize.Width -lt $maxWidth) -or ($global:HostVar.UI.RawUI.BufferSize.Width -lt $maxWidth)) {
            if ($this.StopIfWrongWidth) {
                Write-Warning "Change your console window size to at least $maxWidth characters width and re-run the script"
                #Write-Warning "Or exclude some properties via '-ExcludeProperties' parameter of 'New-ConsoleMenu' cmdlet in the script"    
                break
            }
            else {
                foreach ($line in $consoleMenu) {
                    Write-Host $line
                }
                
                Write-Warning "Change your console window size to at least $maxWidth characters width and re-run the script"
                #Write-Warning "Or exclude some properties via '-ExcludeProperties' parameter of 'New-ConsoleMenu' cmdlet in the script"
            }
        }
        else {
            foreach ($line in $consoleMenu) {
                Write-Host $line
            }
        }
    }
    # Ask the user for input
    [hashtable] AskUser() {
        $SelectedObject = $null

        if ($this.cleared) {
            Write-Host "`"$($this.selection)`" is invalid. Use any of the shown numbers or type `"Q`" to quit" -ForegroundColor Yellow
        }

        $this.selection = Read-Host 'Please type the number of the script you want to run or type "Q" to quit'
        # $selection = 'Q'
        # test if the selection is q to quit
        if ($this.selection -imatch 'q') {
            break
        }
    
        # test if the selection is a number
        # test if selection is between 1 and the number of options
        if ($this.selection -match '^\d+$' -and $this.selection -ge 1 -and [int]$this.selection -le $this.Options.Count) {
            Clear-Host
            $SelectedObject = $this.Options[[int]$this.selection - 1]                
        }
        else {
            Clear-Host
            $this.cleared = $true
        }
        return $SelectedObject
    }
}

# Function to download and run the selected script
Function Invoke-Gist {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$ScriptObject,
        [switch]$NoConfirm
    )
    Write-Host "You selected:" -ForegroundColor Green
    Write-Host "  Gist Name: $($ScriptObject.Name)" -ForegroundColor Green
    Write-Host "  Description: $($ScriptObject.Description)" -ForegroundColor Green
    Write-Host "  Author: $($ScriptObject.Author)" -ForegroundColor DarkMagenta 
    Write-Host "  Elevation Required: $($ScriptObject.Elevation)" -ForegroundColor $(if ($($ScriptObject.Elevation)) { 'Red' } else { 'Green' })
    Write-Host "  URL: $($ScriptObject.Url)" -ForegroundColor Blue

    If ($NoConfirm) {
        $confirmRun = 0
    }
    else {
        $confirmRun = $Host.UI.PromptForChoice('Run the Script', 'Are you sure you want to proceed?', @('&Yes'; '&No'), 0)
    }
    if ($confirmRun -eq 0) {
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Encoding = [System.Text.Encoding]::UTF8
            Invoke-Expression ($wc.DownloadString($($ScriptObject.Url)))
        }
        catch {
            Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

$Menu = [ConsoleMenu]@{
    Title             = $Title
    Options           = $GistCatalog
    ExcludeProperties = "Url"  # Optimize for terminal width
    MaxStringLength   = 1  # Enable dynamic width calculation (will be overridden)
    AddDevideLines    = $true  # Visual separation between items
    #StopIfWrongWidth  = $true
}

# Check if the script is called with a script number
if ($($MyInvocation.MyCommand) -match 'gist\.ittips\.ch/(?:test/|dev/)?(\d+)') {
    [int]$paramScriptNumber = $($matches[1])
    Invoke-Gist -ScriptObject $GistCatalog[$paramScriptNumber - 1] -NoConfirm
}
# Show the menu and ask the user for input and run the selected script
else {
    do {
        $Menu.DrawMenu()
    }
    until ($SelectedGist = $Menu.AskUser())
    Invoke-Gist $SelectedGist
}
