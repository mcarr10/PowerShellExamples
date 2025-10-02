<#
.SYNOPSIS
    On-Call Rotation Scheduler

.DESCRIPTION
    This program generates a weekly on-call rotation schedule for a team,
    taking into account holidays, patching dates, and individual unavailability.

.NOTES
    Input files format:
    - team_members.txt: One team member name per line
    - holidays.txt: One date per line in YYYY-MM-DD format
    - unavailability.txt: Format: "Name,YYYY-MM-DD" (one entry per line)
    - patching.txt: One date per line in YYYY-MM-DD format (software patching dates)
    
    Compatible with PowerShell 5.1+
#>

function Read-TeamMembers {
    param (
        [string]$FileName
    )
    
    if (-not (Test-Path $FileName)) {
        Write-Error "Error: File '$FileName' not found."
        exit 1
    }
    
    $members = Get-Content $FileName | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_.Trim() }
    
    # Shuffle the team members list to randomize starting order
    $members = $members | Sort-Object { Get-Random }
    
    return $members
}

function Read-Holidays {
    param (
        [string]$FileName
    )
    
    if (-not (Test-Path $FileName)) {
        Write-Error "Error: File '$FileName' not found."
        exit 1
    }
    
    $holidays = @()
    Get-Content $FileName | ForEach-Object {
        $line = $_.Trim()
        if ($line -ne "") {
            try {
                $date = [DateTime]::ParseExact($line, "yyyy-MM-dd", $null)
                $holidays += $date.Date
            }
            catch {
                Write-Warning "Invalid date format '$line' in holidays file. Skipping."
            }
        }
    }
    
    return $holidays
}

function Read-Unavailability {
    param (
        [string]$FileName
    )
    
    if (-not (Test-Path $FileName)) {
        Write-Error "Error: File '$FileName' not found."
        exit 1
    }
    
    $unavailability = @{}
    Get-Content $FileName | ForEach-Object {
        $line = $_.Trim()
        if ($line -ne "") {
            $parts = $line -split ','
            if ($parts.Count -ne 2) {
                Write-Warning "Invalid format '$line' in unavailability file. Skipping."
                return
            }
            
            $name = $parts[0].Trim()
            try {
                $date = [DateTime]::ParseExact($parts[1].Trim(), "yyyy-MM-dd", $null)
                if (-not $unavailability.ContainsKey($name)) {
                    $unavailability[$name] = @()
                }
                $unavailability[$name] += $date.Date
            }
            catch {
                Write-Warning "Invalid date format in '$line'. Skipping."
            }
        }
    }
    
    return $unavailability
}

function Read-PatchingDates {
    param (
        [string]$FileName
    )
    
    if (-not (Test-Path $FileName)) {
        Write-Error "Error: File '$FileName' not found."
        exit 1
    }
    
    $patchingDates = @()
    Get-Content $FileName | ForEach-Object {
        $line = $_.Trim()
        if ($line -ne "") {
            try {
                $date = [DateTime]::ParseExact($line, "yyyy-MM-dd", $null)
                $patchingDates += $date.Date
            }
            catch {
                Write-Warning "Invalid date format '$line' in patching file. Skipping."
            }
        }
    }
    
    return $patchingDates
}

function Get-WeekStart {
    param (
        [DateTime]$Date
    )
    
    $dayOfWeek = [int]$Date.DayOfWeek
    # Monday = 1, so if it's Monday (1), subtract 1 day = 0
    # If it's Sunday (0), subtract -1 days = +1, but we want -6
    if ($dayOfWeek -eq 0) {
        $dayOfWeek = 7  # Treat Sunday as 7
    }
    
    return $Date.AddDays(-($dayOfWeek - 1))
}

function Test-IsAvailable {
    param (
        [string]$Member,
        [DateTime]$WeekStart,
        [hashtable]$Unavailability
    )
    
    if (-not $Unavailability.ContainsKey($Member)) {
        return $true
    }
    
    # Check all days in the week
    for ($dayOffset = 0; $dayOffset -lt 7; $dayOffset++) {
        $day = $WeekStart.AddDays($dayOffset).Date
        if ($Unavailability[$Member] -contains $day) {
            return $false
        }
    }
    
    return $true
}

function New-Schedule {
    param (
        [string[]]$TeamMembers,
        [DateTime]$StartDate,
        [int]$NumWeeks,
        [DateTime[]]$Holidays,
        [hashtable]$Unavailability,
        [DateTime[]]$PatchingDates
    )
    
    if ($TeamMembers.Count -eq 0) {
        Write-Error "Error: No team members provided."
        exit 1
    }
    
    $schedule = @()
    $rotationIndex = 0
    $holidayAssignments = @{}
    $patchingAssignments = @{}
    
    foreach ($member in $TeamMembers) {
        $holidayAssignments[$member] = 0
        $patchingAssignments[$member] = 0
    }
    
    $currentDate = Get-WeekStart -Date $StartDate
    
    for ($weekNum = 0; $weekNum -lt $NumWeeks; $weekNum++) {
        $weekStart = $currentDate
        $weekEnd = $weekStart.AddDays(6)
        
        # Check if this week contains a holiday
        $hasHoliday = $false
        for ($dayOffset = 0; $dayOffset -lt 7; $dayOffset++) {
            $day = $weekStart.AddDays($dayOffset).Date
            if ($Holidays -contains $day) {
                $hasHoliday = $true
                break
            }
        }
        
        # Check if this week contains a patching date
        $hasPatching = $false
        for ($dayOffset = 0; $dayOffset -lt 7; $dayOffset++) {
            $day = $weekStart.AddDays($dayOffset).Date
            if ($PatchingDates -contains $day) {
                $hasPatching = $true
                break
            }
        }
        
        # Find an available team member
        $attempts = 0
        $assigned = $false
        
        # Calculate minimum patching assignments
        $minPatching = ($patchingAssignments.Values | Measure-Object -Minimum).Minimum
        
        while ($attempts -lt $TeamMembers.Count) {
            $candidate = $TeamMembers[$rotationIndex % $TeamMembers.Count]
            
            # Check if available and hasn't exceeded holiday limit
            $canAssign = Test-IsAvailable -Member $candidate -WeekStart $weekStart -Unavailability $Unavailability
            
            if ($hasHoliday -and $holidayAssignments[$candidate] -ge 1) {
                $canAssign = $false
            }
            
            # For patching weeks, prioritize those with fewer patching assignments
            if ($hasPatching -and $patchingAssignments[$candidate] -gt $minPatching) {
                # Check if anyone with minimum patching is still available
                $hasMinAvailable = $false
                for ($testIdx = 0; $testIdx -lt $TeamMembers.Count; $testIdx++) {
                    $testCandidate = $TeamMembers[($rotationIndex + $testIdx) % $TeamMembers.Count]
                    $testAvailable = Test-IsAvailable -Member $testCandidate -WeekStart $weekStart -Unavailability $Unavailability
                    
                    if ($patchingAssignments[$testCandidate] -eq $minPatching -and 
                        $testAvailable -and 
                        (-not $hasHoliday -or $holidayAssignments[$testCandidate] -lt 1)) {
                        $hasMinAvailable = $true
                        break
                    }
                }
                
                if ($hasMinAvailable) {
                    $canAssign = $false
                }
            }
            
            if ($canAssign) {
                $schedule += [PSCustomObject]@{
                    WeekNum      = $weekNum + 1
                    WeekStart    = $weekStart.ToString("yyyy-MM-dd")
                    WeekEnd      = $weekEnd.ToString("yyyy-MM-dd")
                    AssignedTo   = $candidate
                    HasHoliday   = $hasHoliday
                    HasPatching  = $hasPatching
                }
                
                if ($hasHoliday) {
                    $holidayAssignments[$candidate]++
                }
                
                if ($hasPatching) {
                    $patchingAssignments[$candidate]++
                }
                
                $rotationIndex++
                $assigned = $true
                break
            }
            
            $rotationIndex++
            $attempts++
        }
        
        if (-not $assigned) {
            $schedule += [PSCustomObject]@{
                WeekNum      = $weekNum + 1
                WeekStart    = $weekStart.ToString("yyyy-MM-dd")
                WeekEnd      = $weekEnd.ToString("yyyy-MM-dd")
                AssignedTo   = "UNASSIGNED - No one available"
                HasHoliday   = $hasHoliday
                HasPatching  = $hasPatching
            }
        }
        
        $currentDate = $currentDate.AddDays(7)
    }
    
    return $schedule
}

function Show-Schedule {
    param (
        [PSCustomObject[]]$Schedule
    )
    
    Write-Host "`n" ("=" * 90)
    Write-Host "ON-CALL ROTATION SCHEDULE"
    Write-Host ("=" * 90)
    Write-Host ("{0,-6} {1,-12} {2,-12} {3,-25} {4,-10} {5,-10}" -f "Week", "Start Date", "End Date", "Assigned To", "Holiday", "Patching")
    Write-Host ("-" * 90)
    
    foreach ($entry in $Schedule) {
        $holidayMarker = if ($entry.HasHoliday) { "Yes" } else { "" }
        $patchingMarker = if ($entry.HasPatching) { "Yes" } else { "" }
        Write-Host ("{0,-6} {1,-12} {2,-12} {3,-25} {4,-10} {5,-10}" -f `
            $entry.WeekNum, $entry.WeekStart, $entry.WeekEnd, $entry.AssignedTo, $holidayMarker, $patchingMarker)
    }
    
    Write-Host ("=" * 90)
    
    # Print summary statistics
    Write-Host "`nSUMMARY:"
    Write-Host ("-" * 90)
    
    # Count assignments per person
    $assignments = @{}
    $holidayCounts = @{}
    $patchingCounts = @{}
    
    foreach ($entry in $Schedule) {
        if ($entry.AssignedTo -ne "UNASSIGNED - No one available") {
            $person = $entry.AssignedTo
            
            # PowerShell 5 compatible null handling
            if ($null -eq $assignments[$person]) {
                $assignments[$person] = 0
            }
            $assignments[$person] = $assignments[$person] + 1
            
            if ($entry.HasHoliday) {
                if ($null -eq $holidayCounts[$person]) {
                    $holidayCounts[$person] = 0
                }
                $holidayCounts[$person] = $holidayCounts[$person] + 1
            }
            
            if ($entry.HasPatching) {
                if ($null -eq $patchingCounts[$person]) {
                    $patchingCounts[$person] = 0
                }
                $patchingCounts[$person] = $patchingCounts[$person] + 1
            }
        }
    }
    
    Write-Host ("{0,-25} {1,-15} {2,-15} {3,-15}" -f "Team Member", "Total Weeks", "Holiday Weeks", "Patching Weeks")
    Write-Host ("-" * 90)
    
    foreach ($person in ($assignments.Keys | Sort-Object)) {
        $total = $assignments[$person]
        
        # PowerShell 5 compatible null handling
        $holidays = if ($null -eq $holidayCounts[$person]) { 0 } else { $holidayCounts[$person] }
        $patching = if ($null -eq $patchingCounts[$person]) { 0 } else { $patchingCounts[$person] }
        
        Write-Host ("{0,-25} {1,-15} {2,-15} {3,-15}" -f $person, $total, $holidays, $patching)
    }
    
    Write-Host ("=" * 90)
}

function Export-Schedule {
    param (
        [PSCustomObject[]]$Schedule,
        [string]$FileName
    )
    
    try {
        $csv = @()
        $csv += "Week,Start Date,End Date,Assigned To,Has Holiday,Has Patching"
        
        foreach ($entry in $Schedule) {
            $holiday = if ($entry.HasHoliday) { "Yes" } else { "No" }
            $patching = if ($entry.HasPatching) { "Yes" } else { "No" }
            $csv += "$($entry.WeekNum),$($entry.WeekStart),$($entry.WeekEnd),$($entry.AssignedTo),$holiday,$patching"
        }
        
        $csv | Out-File -FilePath $FileName -Encoding UTF8
        Write-Host "`nSchedule saved to '$FileName'"
    }
    catch {
        Write-Error "Error saving schedule: $_"
    }
}

# Main execution
Write-Host "On-Call Rotation Scheduler"
Write-Host ("-" * 40)

# Get input filenames
$teamFile = Read-Host "Enter team members file (default: team_members.txt)"
if ([string]::IsNullOrWhiteSpace($teamFile)) {
    $teamFile = "team_members.txt"
}

$holidaysFile = Read-Host "Enter holidays file (default: holidays.txt)"
if ([string]::IsNullOrWhiteSpace($holidaysFile)) {
    $holidaysFile = "holidays.txt"
}

$unavailFile = Read-Host "Enter unavailability file (default: unavailability.txt)"
if ([string]::IsNullOrWhiteSpace($unavailFile)) {
    $unavailFile = "unavailability.txt"
}

$patchingFile = Read-Host "Enter patching dates file (default: patching.txt)"
if ([string]::IsNullOrWhiteSpace($patchingFile)) {
    $patchingFile = "patching.txt"
}

# Get schedule parameters
$startDateStr = Read-Host "Enter start date (YYYY-MM-DD, default: today)"
if ([string]::IsNullOrWhiteSpace($startDateStr)) {
    $startDate = Get-Date
}
else {
    try {
        $startDate = [DateTime]::ParseExact($startDateStr, "yyyy-MM-dd", $null)
    }
    catch {
        Write-Warning "Invalid date format. Using today's date."
        $startDate = Get-Date
    }
}

$numWeeksStr = Read-Host "Enter number of weeks (default: 12)"
if ([string]::IsNullOrWhiteSpace($numWeeksStr)) {
    $numWeeks = 12
}
else {
    try {
        $numWeeks = [int]$numWeeksStr
    }
    catch {
        Write-Warning "Invalid number. Using 12 weeks."
        $numWeeks = 12
    }
}

# Read input files
Write-Host "`nReading input files..."
$teamMembers = Read-TeamMembers -FileName $teamFile
$holidays = Read-Holidays -FileName $holidaysFile
$unavailability = Read-Unavailability -FileName $unavailFile
$patchingDates = Read-PatchingDates -FileName $patchingFile

Write-Host "Loaded $($teamMembers.Count) team members"
Write-Host "Loaded $($holidays.Count) holidays"
Write-Host "Loaded unavailability for $($unavailability.Keys.Count) team members"
Write-Host "Loaded $($patchingDates.Count) patching dates"

# Generate schedule
Write-Host "`nGenerating schedule..."
$schedule = New-Schedule -TeamMembers $teamMembers -StartDate $startDate -NumWeeks $numWeeks `
    -Holidays $holidays -Unavailability $unavailability -PatchingDates $patchingDates

# Display schedule
Show-Schedule -Schedule $schedule

# Save to file
$outputFile = Read-Host "`nEnter output filename (default: schedule.csv)"
if ([string]::IsNullOrWhiteSpace($outputFile)) {
    $outputFile = "schedule.csv"
}

Export-Schedule -Schedule $schedule -FileName $outputFile