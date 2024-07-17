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

.EXAMPLE
    .\loginevents.ps1

    Displays only the first login event and the longest lunch break of the current day.

.EXAMPLE
    .\loginevents.ps1 -full

    Displays all login and logout events of the current day.

.NOTES
    Author: tomtucker18
    Date: 2024-07-17
#>

param (
    [switch]$full
)

Write-Host ""

# Get the current date without the time component
$today = (Get-Date).Date

try {
    # Retrieve events from the OneApp_IGCC log
    $events = Get-WinEvent -LogName "OneApp_IGCC" | Where-Object {
        $_.TimeCreated -ge $today -and $_.ProviderName -eq "OneApp_IGCC_WinService"
    }

    # Filter events for specific messages
    $filteredEvents = $events | Where-Object {
        $_.Message -match "SessionLogon|SessionLock|SessionUnlock"
    }

    # Sort the filtered events by date ascending
    $sortedEvents = $filteredEvents | Sort-Object -Property TimeCreated

    # Initialize variables
    $firstEvent = $null
    $lunchStart = $null
    $lunchEnd = $null
    $maxLunchDuration = [TimeSpan]::MinValue

    Write-Host "Time  Action"
    Write-Host "----- ------"

    foreach ($event in $sortedEvents) {
        $action = switch -Regex ($event.Message) {
            "SessionLogon" { "Login" }
            "SessionLock" { "Lock" }
            "SessionUnlock" { "Unlock" }
            default { $event.Message }
        }

        $time = $event.TimeCreated

        # Capture the first event
        if (-not $firstEvent) {
            $firstEvent = $event
            $firstEventTime = $time.ToString("HH:mm")
        }

        # Analyze lunch break
        if ($time.Hour -ge 11 -and $time.Hour -lt 14) {
            if ($action -eq "Lock") {
                $lunchStart = $event
            }
            elseif ($action -eq "Unlock" -and $lunchStart) {
                $lunchEnd = $event
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
                "Login" { $color = "Green" }
                "Unlock" { $color = "Green" }
                default { $color = "White" }
            }

            # Output the line with the appropriate color for the time
            $eventTime = $time.ToString("HH:mm")
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
        }

        # Display lunch break
        if ($maxLunchStart -and $maxLunchEnd) {
            $lunchStartTime = $maxLunchStart.TimeCreated.ToString("HH:mm")
            $lunchEndTime = $maxLunchEnd.TimeCreated.ToString("HH:mm")
            Write-Host -NoNewline -ForegroundColor "Red" "$lunchStartTime"
            Write-Host -NoNewline -ForegroundColor "Red" " "
            Write-Host "Lunch"
            Write-Host -NoNewline -ForegroundColor "Green" "$lunchEndTime"
            Write-Host -NoNewline -ForegroundColor "Green" " "
            Write-Host "End Lunch"
        }
        else {
            Write-Host "No lunch break found"
        }
    }
}
catch {
    Write-Error "Error retrieving events: $_"
}

Write-Host ""

# Prompt to keep window open
Read-Host -Prompt "Press Enter to exit"
