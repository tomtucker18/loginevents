# LoginEvents PowerShell Script

## Overview

The **LoginEvents PowerShell Script** helps analyze work time activities by reading events from the `OneApp_IGCC` event log. It provides an overview of daily activity, including the first login, the last logout, and the longest lunch break between "Lock" and "Unlock" events.

## Features

- Display all login/logout events of the day or just the most important ones (first login and longest lunch break).
- Customize the displayed time range using indices.
- View event details for any day by navigating backward or forward in time.

## Parameters

| Parameter    | Description                                                                                                |
| ------------ | ---------------------------------------------------------------------------------------------------------- |
| `-full`      | Displays all events of the day. Without this flag, only the first event and longest lunch break are shown. |
| `-showIndex` | Adds an index to each event. Requires `-full`.                                                             |
| `-deltaDays` | Specifies how many days to go back/forward from today. Default: `0` (today).                               |
| `-from`      | Specifies the starting index for events.                                                                   |
| `-to`        | Specifies the ending index for events.                                                                     |

## Usage Examples

1. **Default behavior (important events only):**

   ```powershell
   .\loginevents.ps1
   ```

   Displays the first login and the longest lunch break for the current day.

2. **View all events:**

   ```powershell
   .\loginevents.ps1 -full
   ```

3. **View events for the previous day:**

   ```powershell
   .\loginevents.ps1 -deltaDays -1
   ```

4. **View events with indices for troubleshooting:**

   ```powershell
   .\loginevents.ps1 -full -showIndex
   ```

5. **View events between specific indices:**

   ```powershell
   .\loginevents.ps1 -full -from 5 -to 10
   ```

## How It Works

1. **Event Filtering:** The script filters `OneApp_IGCC` logs for relevant events: `SessionLogon`, `SessionLogoff`, `SessionLock`, and `SessionUnlock`.
2. **Lunch Break Calculation:** The longest break between a `Lock` and `Unlock` event between 11:00 and 14:00 is identified as the lunch break.
3. **Work Time Calculation:** The total work time is computed as the time between the first login and the last logout, minus the lunch break.

## Notes

- Ensure that the `OneApp_IGCC` log and the `OneApp_IGCC_WinService` provider are configured on your system.
- The script uses indexed events for filtering. Use `-showIndex` to determine the correct indices for `-from` and `-to`.

## License

This script is licensed under the [MIT License](https://opensource.org/licenses/MIT). Feel free to contribute or modify as needed.
