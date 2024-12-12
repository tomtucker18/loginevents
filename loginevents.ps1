<#
.SYNOPSIS
    This script displays the login and logout activities of the current day.

.DESCRIPTION
    The script reads events from the OneApp_IGCC log and displays either all events of the day or just the first 
    event of the day and the longest lunch break. The lunch break is defined as the longest break between a "Lock" 
    and "Unlock" event that occurs between 11:00 AM and 2:00 PM.

.PARAMETER full
    When this flag is set, all events of the day are displayed. Without this flag, only the first event and the 
    longest lunch break are displayed.

.PARAMETER showIndex
    When this flag is set in combination with the -full parameter, the events are displayed with an index number 
    at the beginning of each line.

.PARAMETER deltaDays
    Specifies the number of days to go back or forward from the current day. The default is 0, which represents 
    the current day.

.EXAMPLE
    .\loginevents.ps1

    Displays only the first login event and the longest lunch break of the current day.

.EXAMPLE
    .\loginevents.ps1 -full

    Displays all login and logout events of the current day.

.EXAMPLE
    .\loginevents.ps1 -deltaDays -1

    Displays the worktime and lunch break of the previous day.

.NOTES
    Author: tomtucker18
    Date: 2024-10-24
#>

param (
    [switch]$full,
    [switch]$showIndex,
    [string]$deltaDays = 0
)

# Function to convert a TimeSpan object to a formatted string
function Convert-TimeSpanToString {
    param (
        [Parameter(Mandatory = $true)]
        [System.TimeSpan]$TimeSpan,
        
        [Parameter(Mandatory = $false)]
        [string]$Format = "HH:mm"
    )

    # Extract the components of the TimeSpan
    $hours = [math]::Floor($TimeSpan.TotalHours)
    $minutes = $TimeSpan.Minutes
    $seconds = $TimeSpan.Seconds    

    # Format the output based on the provided format
    switch ($Format) {
        "HH:mm:ss" {
            return "{0:00}:{1:00}:{2:00}" -f $hours, $minutes, $seconds
        }
        "HH:mm" {
            return "{0:00}:{1:00}" -f $hours, $minutes
        }
        "text" {
            return "{0}h{1}" -f $hours, $minutes
        }
        default {
            throw "Unsupported format. Please use 'HH:mm:ss', 'HH:mm' or 'text'."
        }
    }
}

Write-Host ""

if ($showIndex -and -not $full) {
    Write-Warning "The -showIndex parameter only takes effect in combination with the -full parameter."
    Write-Host ""
}

try {
    # Get the current date without the time component
    $today = (Get-Date).Date.AddDays($deltaDays)

    # Display the relevant date based on parameters
    if ($deltaDays -eq 0 -and $full) {
        Write-Host "Today's login events`n"
    }
    elseif ($deltaDays -ne 0) {
        Write-Host "Login events for $($today.Date.ToShortDateString())`n"
    }

    # Retrieve events from the OneApp_IGCC log
    $events = Get-WinEvent -LogName "OneApp_IGCC" | Where-Object {
        $_.TimeCreated -ge $today -and $_.TimeCreated -lt $today.AddDays(1) -and $_.ProviderName -eq "OneApp_IGCC_WinService"
    }

    # Filter events for specific messages
    $filteredEvents = $events | Where-Object {
        $_.Message -match "SessionLogon|SessionLock|SessionUnlock|SessionLogoff"
    }

    # Sort the filtered events by date ascending
    $sortedEvents = $filteredEvents | Sort-Object -Property TimeCreated

    # Determine the width of the index column
    $indexWidth = [Math]::Max(5, $sortedEvents.Count.ToString().Length)

    # Initialize variables
    $firstEvent = $null
    $lunchStart = $null
    $lunchEnd = $null
    $maxLunchDuration = [TimeSpan]::MinValue

    # Display index column in the header
    if ($showIndex -and $full) {
        Write-Host -NoNewline "Index".PadRight($indexWidth + 1)
    }
    Write-Host "Time  Action"
    # Display index column separator
    if ($showIndex -and $full) {
        Write-Host -NoNewline ("-" * ($indexWidth ))
        Write-Host -NoNewline " "
    }
    Write-Host "----- ------"

    for ($i = 0; $i -lt $sortedEvents.Count; $i++) {
        $currentEvent = $sortedEvents[$i]
        # Determine the action based on the event message
        $action = switch -Regex ($currentEvent.Message) {
            "SessionLogon" { "Login" }
            "SessionLock" { "Lock" }
            "SessionUnlock" { "Unlock" }
            "SessionLogoff" { "Logout" }
            default { $currentEvent.Message }
        }

        $time = $currentEvent.TimeCreated

        # Capture the first event
        if (-not $firstEvent) {
            $firstEvent = $currentEvent
            $firstEventTime = $time.ToString("HH:mm")
        }

        # Analyze lunch break
        if ($time.Hour -ge 11 -and $time.Hour -lt 14) {
            if ($action -eq "Lock") {
                $lunchStart = $currentEvent
            }
            elseif ($action -eq "Unlock" -and $lunchStart) {
                $lunchEnd = $currentEvent
                $lunchDuration = $lunchEnd.TimeCreated - $lunchStart.TimeCreated

                if ($lunchDuration -gt $maxLunchDuration) {
                    $maxLunchDuration = $lunchDuration
                    $maxLunchStart = $lunchStart
                    $maxLunchEnd = $lunchEnd
                }

                $lunchStart = $null
                $lunchEnd = $null
            }
        }

        # Display all events if the "full" flag is set
        if ($full) {
            # Select color based on action
            switch ($action) {
                "Lock" { $color = "Red" }
                "Logout" { $color = "Red" }
                "Login" { $color = "Green" }
                "Unlock" { $color = "Green" }
                default { $color = "White" }
            }

            # Output the line with the appropriate color for the time
            $eventTime = $time.ToString("HH:mm")
            if ($showIndex) {
                $paddedValue = $i.ToString().PadLeft($indexWidth - 1)
                Write-Host -NoNewline -ForegroundColor "White" "$paddedValue) "
            }
            Write-Host -NoNewline -ForegroundColor $color "$eventTime"
            Write-Host -NoNewline -ForegroundColor $color " "
            Write-Host "$action"
        }
    }

    # Display results if the "full" flag is not set
    if (-not $full) {
        # Display first event
        if ($firstEvent) {
            Write-Host -NoNewline -ForegroundColor "Green" "$firstEventTime"
            Write-Host -NoNewline -ForegroundColor "Green" " "
            Write-Host "Start"

            # Calculate total work time until now
            $now = Get-Date
            $totalWorkTime = $now - $firstEvent.TimeCreated
            
            # Display lunch break
            if ($maxLunchStart -and $maxLunchEnd -and $maxLunchDuration) {
                # Subtract lunch break from total work time
                $totalWorkTime = $totalWorkTime - $maxLunchDuration
                $lunchStartTime = $maxLunchStart.TimeCreated.ToString("HH:mm")
                $lunchEndTime = $maxLunchEnd.TimeCreated.ToString("HH:mm")
                
                # Display lunch break
                Write-Host -NoNewline -ForegroundColor "Red" "$lunchStartTime "
                Write-Host "Lunch"
                Write-Host -NoNewline -ForegroundColor "Green" "$lunchEndTime "
                Write-Host "End Lunch"
            }
            else {
                Write-Host "No lunch break found"
            }
            
            # Display total work time
            $sumMessage = "`nWork time until now: "
            if ($deltaDays -ne 0) {
                # Calculate work time for the whole day
                $lastEvent = $sortedEvents[-1]
                $lastEventTime = $lastEvent.TimeCreated.ToString("HH:mm")
                $totalWorkTime = $lastEvent.TimeCreated - $firstEvent.TimeCreated

                $sumMessage = "`nWork time for the day: "

                Write-Host -NoNewline -ForegroundColor "Red" "$lastEventTime "
                Write-Host "End Work"
            }
            $formattedTime = Convert-TimeSpanToString -TimeSpan $totalWorkTime -Format "text"
            Write-Host -NoNewline $sumMessage
            Write-Host -ForegroundColor "Blue" "$formattedTime"
        }
        else {
            Write-Host "No events found for today"
        }

    }
}
catch {
    Write-Error "Error retrieving events: $_"
}

Write-Host ""

# Prompt to keep window open
Read-Host -Prompt "Press Enter to exit"