<#
.SYNOPSIS
Exports all event log data from one or more .evtx files into separate CSV files, sorted by time.

.DESCRIPTION
This script iterates through all .evtx files in a specified input directory (and its subdirectories),
parses the XML content of each event, and exports a CSV file for each log.
It calculates the Unix epoch time (in milliseconds) for each event, adds it as the first column,
and sorts the entire dataset by this numerical field. It processes ALL unique fields found in the XML
to ensure comprehensive data extraction, resulting in a column for every unique system and event data field.

OPTIMIZATION NOTE: To improve performance, the script now caches the XML representation of each event
in a single pass before starting the data extraction loops.

.PARAMETER indir
The directory path containing the input .evtx files. Files will be searched for recursively.

.PARAMETER outdir
The directory path where the resulting CSV files will be saved. Output files will be
named 'timeline-<original_filename_base>.csv'.

.EXAMPLE
.\evtx_to_csv_multidir.ps1 -indir C:\Logs\Raw -outdir C:\Logs\Processed
#>
param (
    [string]$indir = $null,
    [string]$outdir = $null
)

# --- Argument Validation and Help ---
if (!$indir -or !$outdir) {
    Write-Host "Usage: Bulk-Evtx-2-CSV.ps1 -indir <directory with .evtx files> -outdir <output directory for .csv files>"
    exit 1
}

# Ensure the input directory exists
if (-not (Test-Path $indir -PathType Container)) {
    Write-Error "Input directory '$indir' does not exist."
    exit 1
}

# Ensure the output directory exists, create if it doesn't
if (-not (Test-Path $outdir -PathType Container)) {
    Write-Host "Output directory '$outdir' does not exist. Creating it."
    New-Item -Path $outdir -ItemType Directory | Out-Null
}

$files = Get-ChildItem -Path $indir -Filter "*.evtx" -Recurse

if ($files.Count -eq 0) {
    Write-Host "No .evtx files found in $indir (recursively). Exiting."
    exit 0
}

Write-Host ("Found $($files.Count) .evtx file(s) to process.")

# Define the Unix Epoch start time (1970-01-01 00:00:00Z)
$UnixEpoch = (Get-Date "1970-01-01 00:00:00Z").ToUniversalTime()

# --- Main Processing Loop ---
foreach ($file in $files) {
    Write-Host "--------------------------------------------------------"
    Write-Host "Processing file: $($file.Name) (Path: $($file.DirectoryName))"
    
    $infile = $file.FullName
    $initem = $file
    $outfile = "timeline-$($initem.BaseName).csv"
    $output_filepath = Join-Path -Path $outdir -ChildPath $outfile

    Write-Host "Reading in the .evtx. (This might take a moment...)"
    try {
        # Use -ErrorAction Stop to handle potential errors during read
        $events = Get-WinEvent -Path $infile -ErrorAction Stop
    } catch {
        Write-Error "Failed to read event log file $($file.Name). Error: $($_.Exception.Message)"
        continue
    }

    # --- PERFORMANCE OPTIMIZATION: Cache XML Object in a single pass ---
    Write-Host "Caching XML for all $($events.Count) events..."
    $processed_events = $events | ForEach-Object {
        $event_xml_string = $_.ToXml()
        [PSCustomObject]@{
            EventObject = $_                  # Original event object (for TimeCreated, Message, etc.)
            XmlObject = [xml]$event_xml_string # Cached XML object (parsed once)
        }
    }
    
    # --- Field Discovery Pass ---
    echo "Finding unique fields."
    # Use a Hashtable for faster lookups during field discovery
    $field_set = @{}
    $fields = @()
    
    # Add essential default fields in desired order
    $fields += "EpochTime"
    $field_set["EpochTime"] = $true
    $fields += "TimeCreated"
    $field_set["TimeCreated"] = $true
    $fields += "Message"
    $field_set["Message"] = $true

    foreach ($pEvent in $processed_events) {
        $xml = $pEvent.XmlObject # Use the cached XML object
        
        # System nodes (ProviderName, EventID, Level, etc.)
        foreach ($s in $xml.Event.System.ChildNodes) {
            if ($s.Name -and (-not $field_set.ContainsKey($s.Name)) -and $s.Name -ne "Microsoft-Windows-Security-Auditing") {
                $fields += $s.Name
                $field_set[$s.Name] = $true
            }
        }
        
        # EventData nodes (Data fields)
        if ($xml.Event.EventData) {
            foreach ($d in $xml.Event.EventData.Data) {
                if ($d.Name -and (-not $field_set.ContainsKey($d.Name))) {
                    $fields += $d.Name
                    $field_set[$d.Name] = $true
                }
            }
        }
    }

    echo "Found $($fields.Count) unique fields."

    # --- Data Population Pass ---
    $lines = @()
    echo "Processing lines and populating data structure."
    foreach ($pEvent in $processed_events) {
        # hash of fields and their values in this event
        $line = @{}
        $event = $pEvent.EventObject
        $xml = $pEvent.XmlObject # Use the cached XML object
        
        # Calculate Epoch Time in Milliseconds
        $TimeCreatedUtc = $event.TimeCreated.ToUniversalTime()
        $EpochTimeMs = [long]($TimeCreatedUtc - $UnixEpoch).TotalMilliseconds
        
        $line.add("EpochTime", $EpochTimeMs) # Add EpochTime (1st column)
        
        # Add TimeCreated (2nd column) and clean Message (3rd column)
        $line.add("TimeCreated", $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss.fff"))
        $line.add("Message", ($event.Message -split '\n')[0].replace("`n","").replace("`r",""))
        
        $unlabled_fields = @()
        
        # System nodes
        foreach ($s in $xml.Event.System.ChildNodes) {
            if ($s.InnerText) {
                $line.Add($s.Name, $s.InnerText.replace("`n","\n").replace("`r","\n"))
            }
        }
        
        # EventData nodes
        if ($xml.Event.EventData) {
            foreach ($d in $xml.Event.EventData.Data) {
                # if the element has a name, then it is properly formatted and parse it
                if ($d.Name) {
                    $text = $d.InnerText
                    if ($text -eq $null) { $text = "" }
                    $text = $text.replace("`n","\n").replace("`r","\n")
                    
                    # Check for field name collision before adding
                    if (-not $line.ContainsKey($d.Name)) {
                        $line.Add($d.Name, $text)
                    }
                }
                # if the element does not have a name, it's poorly formatted
                elseif ($d) {
                    $text = $d.InnerXml
                    $text = $text.replace("`n","\n").replace("`r","\n")
                    
                    # Create a placeholder field name
                    $newfield = "unlabeled" + ([int]$unlabled_fields.count + 1)
                    $unlabled_fields += $newfield
                    
                    if (-not $line.ContainsKey($newfield)) {
                        $line.Add($newfield, $text)
                    }
                }
            }
        }
        
        $lines += New-Object PSObject -Property $line
        
        # add any new *unlabeled* field names that were added to $fields if needed
        foreach ($f in $unlabled_fields) {
            if (-not $field_set.ContainsKey($f)) {
                $fields += $f
                $field_set[$f] = $true
            }
        }
    }
    echo ("Processed $($lines.Count) events.")

    # --- CSV Writing ---
    echo "Writing output file to: $output_filepath (Sorted by EpochTime)"

    # Exporting the lines to CSV: Sort by EpochTime, then Select columns in the correct order
    try {
        $lines | Sort-Object -Property EpochTime | Select-Object -Property $fields | Export-Csv -Path $output_filepath -NoTypeInformation -Delimiter "," -Encoding UTF8 -Force
    } catch {
        Write-Error "Failed to write CSV file. Error: $($_.Exception.Message)"
    }
    
    Write-Host "Finished processing $($file.Name)."
}

Write-Host "--------------------------------------------------------"
Write-Host "All files processed successfully."