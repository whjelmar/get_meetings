# Define directories and maximum path length
$maxPathLength = 250  # Maximum length for the path

$markdownOutputDir = "C:\projects\tmp\meetings"
$peopleDir = "C:\projects\tmp\people"
$recurringDir = "C:\projects\tmp\recurring"
$logFilePath = "C:\projects\tmp\get_meetings.log"

# Define the template directory
$templateDir = "C:\projects\tmp\templates"

# Define template file paths using the template directory variable
$peopleTemplatePath = Join-Path $templateDir "template_people.md"
$meetingsTemplatePath = Join-Path $templateDir "template_meetings.md"
$recurringTemplatePath = Join-Path $templateDir "template_recurring.md"

function Write-Log {
    param (
        [string]$Message
    )
    Add-Content -Path $logFilePath -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
}

function ConvertTo-NormalizedName {
    param (
        [string]$Name
    )
    $trimmedName = $Name.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedName) -eq $false) {
        return $trimmedName
    }
    return $null
}

function ConvertTo-SafeFileName {
    param (
        [string]$FileName
    )
    $safeFileName = ($FileName -replace "[\\/:*?`"<>|@[\]]", "-").Trim()
    if ($safeFileName.Length -gt $maxPathLength) {
        $extension = ".md"
        $safeFileName = $safeFileName.Substring(0, $maxPathLength - $extension.Length) + $extension
    }
    return $safeFileName
}

function Generate-ContentFromTemplate {
    param (
        [string]$TemplatePath,
        [hashtable]$Replacements
    )
    $templateContent = Get-Content -Path $TemplatePath -Raw
    foreach ($key in $Replacements.Keys) {
        $templateContent = $templateContent -replace "{{\s*$key\s*}}", $Replacements[$key]
    }
    return $templateContent
}

function Update-AttendeeFile {
    param (
        [string]$FilePath,
        [string]$MeetingDetails
    )
    $replacements = @{
        "PersonName" = $MeetingDetails.PersonName;
        "MeetingList" = $MeetingDetails.MeetingList
    }
    $content = Generate-ContentFromTemplate -TemplatePath $peopleTemplatePath -Replacements $replacements
    Add-Content -Path $FilePath -Value $content
}

function Update-RecurringMeetingFile {
    param (
        [object]$Appointment,
        [string]$RecurringFilePath
    )
    if (-not (Test-Path -Path $RecurringFilePath)) {
        # File doesn't exist, use template to create it
        $replacements = @{
            "MeetingSubject" = $Appointment.Subject
            "MeetingFrequency" = "Define frequency here"
            "NextMeetingDate" = $Appointment.Start.ToString('yyyy-MM-dd')
            "PastMeetingsList" = ""
        }
        $content = Generate-ContentFromTemplate -TemplatePath $recurringTemplatePath -Replacements $replacements
    } else {
        # File exists, append a link to the new meeting
        $meetingDate = $Appointment.Start.ToString('yyyy-MM-dd')
        $content = "`r`n- [$meetingDate Meeting]($meetingDate - $($Appointment.Subject).md)"
    }
    Add-Content -Path $RecurringFilePath -Value $content
}

function Format-MeetingAsMarkdown {
    param (
        [object]$Appointment
    )
    $attendees = $Appointment.RequiredAttendees -split ";" | Sort-Object -Unique
    foreach ($attendee in $attendees) {
        $normalizedAttendeeName = ConvertTo-NormalizedName -Name $attendee
        if ($null -ne $normalizedAttendeeName) {
            $attendeeFileName = ConvertTo-SafeFileName -FileName "$normalizedAttendeeName.md"
            $attendeeFilePath = Join-Path $peopleDir $attendeeFileName
            $meetingDetails = "$($Appointment.Subject) on $($Appointment.Start.ToString('yyyy-MM-dd'))"
            Update-AttendeeFile -FilePath $attendeeFilePath -MeetingDetails $meetingDetails
        }
    }
    $attendeesList = $attendees -replace "^(.*)$", "- [ ] `$1"
    $replacements = @{
        "MeetingSubject" = $Appointment.Subject
        "MeetingStart" = $Appointment.Start.ToString('g')
        "MeetingEnd" = $Appointment.End.ToString('g')
        "MeetingLocation" = $Appointment.Location
        "MeetingOrganizer" = $Appointment.Organizer
        "AttendeesList" = $($attendeesList -join "`n")
    }
    return Generate-ContentFromTemplate -TemplatePath $meetingsTemplatePath -Replacements $replacements
}

try {
    Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" -ErrorAction Stop
    $outlook = New-Object -ComObject Outlook.Application -ErrorAction Stop
    Write-Log "Outlook application object created successfully."
} catch {
    Write-Log "Error loading Outlook assembly or creating application object: $_"
    exit
}

try {
    $namespace = $outlook.GetNameSpace("MAPI")
    $folder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar)
    $appointments = $folder.Items
    Write-Log "Accessed Calendar items successfully."
} catch {
    Write-Log "Error accessing Calendar items: $_"
    exit
}

try {
    if (-not (Test-Path -Path $markdownOutputDir)) {
        New-Item -Path $markdownOutputDir -ItemType Directory
        Write-Log "Created meetings output directory."
    }

    if (-not (Test-Path -Path $peopleDir)) {
        New-Item -Path $peopleDir -ItemType Directory
        Write-Log "Created people directory."
    }

    $appointments.Sort("[Start]")
    $appointments.IncludeRecurrences = $true
    $startRange = [DateTime]::Now
    $endRange = $startRange.AddDays(7) # Next 7 days
    $restriction = "[Start] >= `"$($startRange.ToString('g'))`" AND [End] <= `"$($endRange.ToString('g'))`""
    $appointments = $appointments.Restrict($restriction)

    foreach ($appointment in $appointments) {
        if ($appointment.IsRecurring) {
            $recurrencePattern = $appointment.GetRecurrencePattern()
            $recurringFileName = ConvertTo-SafeFileName -FileName ("Recurring - " + $appointment.Subject + ".md")
            $recurringFilePath = Join-Path $recurringDir $recurringFileName
            Update-RecurringMeetingFile -Appointment $appointment -RecurringFilePath $recurringFilePath
        }

    $markdownContent = Format-MeetingAsMarkdown -Appointment $appointment
    $sanitizedSubject = ConvertTo-SafeFileName -FileName ($appointment.Subject + ".md")
    $fileName = $appointment.Start.ToString('yyyy-MM-dd') + " - " + $sanitizedSubject
    $filePath = Join-Path $markdownOutputDir $fileName
    Set-Content -Path $filePath -Value $markdownContent
}
Write-Log "Meeting markdown files created."

} catch {
    Write-Log "Error processing appointments: $_"
}

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
Write-Log "COM object released."
