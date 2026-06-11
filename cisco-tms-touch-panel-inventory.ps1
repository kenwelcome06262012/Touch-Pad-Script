# --- SET TLS (safe for consistency) ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- LOAD FILE ---
$file = "$env:USERPROFILE\Desktop\TMS IPs.xlsx"
$data = Import-Excel $file

# --- CREDENTIALS ---
Write-Host "Enter PRIMARY credential"
$cred1 = Get-Credential

Write-Host "Enter SECONDARY credential (if needed)"
$cred2 = Get-Credential

# --- RESULTS ARRAY ---
$results = foreach ($row in $data) {

    $name = $row."System Name"
    $ip   = $row."IP Address"

    if (-not $ip) { continue }

    Write-Host "Checking $name ($ip)..."

    $status = "Unknown"
    $touchName = ""
    $touchIP = ""

    $success = $false

    foreach ($cred in @($cred1, $cred2)) {

        try {
            $url = "https://$ip/getxml?location=/Status/Peripherals/ConnectedDevice"

            $response = Invoke-WebRequest `
                -Uri $url `
                -Credential $cred `
                -SkipCertificateCheck `
                -TimeoutSec 5 `
                -ErrorAction Stop

            Write-Host "✅ Connected using $($cred.UserName)"

            [xml]$xml = $response.Content

            $devices = $xml.Status.Peripherals.ConnectedDevice

            if ($devices) {
                $touch = $devices | Where-Object { $_.Name -match "Touch|Navigator" }

                if ($touch) {
                    $touchName = $touch.Name
                    $touchIP   = $touch.NetworkAddress
                    $status    = "Success"

                    Write-Host "   ↳ Touch: $touchName ($touchIP)"
                }
                else {
                    $touchName = "None"
                    $status    = "No Touch Panel"

                    Write-Host "   ↳ No touch panel found"
                }
            }
            else {
                $status = "No Peripheral Data"
                Write-Host "   ↳ No peripheral data returned"
            }

            $success = $true
            break
        }
        catch {
            $errorMsg = $_.Exception.Message

            if ($errorMsg -match "401") {
                Write-Host "   ❌ Auth failed with $($cred.UserName)"
                $status = "Auth Failed"
                continue
            }
            elseif ($errorMsg -match "timed out|Unable to connect") {
                Write-Host "   ❌ Device unreachable"
                $status = "Offline"
                break
            }
            else {
                Write-Host "   ❌ Error: $errorMsg"
                $status = "Error"
                break
            }
        }
    }

    # If nothing worked
    if (-not $success -and $status -eq "Unknown") {
        $status = "Failed"
    }

    # OUTPUT OBJECT
    [PSCustomObject]@{
        SystemName = $name
        CodecIP    = $ip
        TouchName  = $touchName
        TouchIP    = $touchIP
        Status     = $status
    }
}

# --- EXPORT RESULTS ---
$output = "$env:USERPROFILE\Desktop\TMS_TouchPanels_FULL.csv"
$results | Export-Csv $output -NoTypeInformation

Write-Host "`n✅ Finished! File saved to:"
Write-Host $output

# Show results in console
$results