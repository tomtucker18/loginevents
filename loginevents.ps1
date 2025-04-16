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

.PARAMETER keepOpen
    When this flag is set you will be prompted to press Enter before the script exits. This is useful when running the script
    from a shortcut or link, as it allows you to see the output before the window closes.

.PARAMETER deltaDays
    Specifies the number of days to go back or forward from the current day. The default is 0, which represents 
    the current day.

.PARAMETER from
    Provide an index to use as the first event of the day. This can be useful when the first event of the day is
    not a relevant login event.

    To determine the correct index, use the -showIndex and -full flag to display event indices.

.PARAMETER to
    Provide an index to use as the last event of the day. This can be useful when the last event of the day is
    not a relevant login event.

    To determine the correct index, use the -showIndex and -full flag to display event indices.

.EXAMPLE
    .\loginevents.ps1

    Displays only the first login event and the longest lunch break of the current day.

.EXAMPLE
    .\loginevents.ps1 -full

    Displays all login and logout events of the current day.

.EXAMPLE
    .\loginevents.ps1 -deltaDays -1

    Displays the worktime and lunch break of the previous day.

.EXAMPLE
    .\loginevents.ps1 -full -showIndex -from 5 -to 10

    Displays the events of the current day with an index number at the beginning of each line. Only events with
    indices between 5 and 10 are shown.

.NOTES
    Author: tomtucker18
    Date: 2024-01-23
#>

param (
    [switch]$full,
    [switch]$showIndex,
    [string]$deltaDays = 0,
    [switch]$keepOpen,
    [Nullable[int]]$from,
    [Nullable[int]]$to
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

# Start Windows Terminal loading animation
function StartLoadingAnimation {
    Write-Host -NoNewline ([char]27 + "]9;4;3;0" + [char]7)
}

# Stop Windows Terminal loading animation
function StopLoadingAnimation {
    # Stop Windows Terminal loading animation
    Write-Host -NoNewline ([char]27 + "]9;4;0;0" + [char]7)
}

# Helper function to validate index range
function ValidateIndexRange {
    param (
        [Nullable[int]]$from,
        [Nullable[int]]$to,
        [int]$maxCount
    )

    if ($from -ne $null -and $from -lt 0) {
        StopLoadingAnimation
        throw "Invalid index range. The 'from' index must be greater than or equal to 0."
    }

    if ($from -ne $null -and $from -ge $maxCount) {
        StopLoadingAnimation
        throw "Invalid index range. The 'from' index must be less than the total number of events."
    }

    if ($to -ne $null -and $to -ge $maxCount) {
        StopLoadingAnimation
        throw "Invalid index range. The 'to' index must be less than the total number of events."
    }

    if ($to -ne $null -and $to -lt 1) {
        StopLoadingAnimation
        throw "Invalid index range. The 'to' index must be greater than or equal to 1."
    }

    if ($from -ne $null -and $to -ne $null -and $from -gt $to) {
        StopLoadingAnimation
        throw "Invalid index range. The 'from' index must be less than or equal to the 'to' index."
    }
}

Write-Host ""

if ($showIndex -and -not $full) {
    Write-Warning "The -showIndex parameter only takes effect in combination with the -full parameter."
    Write-Host ""
}

try {
    StartLoadingAnimation

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

    ValidateIndexRange -from $from -to $to -maxCount $sortedEvents.Count

    # Filter events based on the provided index range
    if ($from -and $to) {
        $sortedEvents = $sortedEvents[$from..$to]
    }
    elseif ($from) {
        $sortedEvents = $sortedEvents[$from..($sortedEvents.Count - 1)]
    }
    elseif ($to) {
        $sortedEvents = $sortedEvents[0..$to]
    }

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
                $currentIndex = $i
                if ($from) {
                    $currentIndex += $from
                }
                $paddedValue = $currentIndex.ToString().PadLeft($indexWidth - 1)
                Write-Host -NoNewline -ForegroundColor "White" "$paddedValue) "
            }
            Write-Host -NoNewline -ForegroundColor $color "$eventTime"
            Write-Host -NoNewline -ForegroundColor $color " "
            Write-Host "$action"
        }
    }

    StopLoadingAnimation

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
            
            # Calculate work time for the whole day
            $lastEvent = $sortedEvents[-1]
            if ($deltaDays -ne 0) {
                $totalWorkTime = $lastEvent.TimeCreated - $firstEvent.TimeCreated
            }
            
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
                $sumMessage = "`nWork time for the day: "
                $lastEventTime = $lastEvent.TimeCreated.ToString("HH:mm")
                Write-Host -NoNewline -ForegroundColor "Red" "$lastEventTime "
                Write-Host "End Work"
            }
            $formattedTime = Convert-TimeSpanToString -TimeSpan $totalWorkTime -Format "text"
            Write-Host -NoNewline $sumMessage
            Write-Host -ForegroundColor "Blue" "$formattedTime"
        }
        else {
            $noEventsFoundMessage = "No events found for today";
            if ($deltaDays -ne 0) {
                $noEventsFoundMessage = "No events found for this day";
            }
            Write-Host $noEventsFoundMessage
        }

    }
}
catch {
    Write-Error "Error retrieving events: $_"
}

# Prompt to keep window open
if($keepOpen) {
    Write-Host ""
    Read-Host -Prompt "Press Enter to exit"
}