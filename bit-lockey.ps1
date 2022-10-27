$computername = Get-WmiObject -Class Win32_Bios | Select PSComputername
$computername = $computername.PSComputerName
$username = $env:UserName
$windows_key = (Get-WmiObject -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
$file_name = 'C:\Users\{0}\Downloads\Bitlocker_{1}_{2}.txt' -f $username, $username, $computername 
$bitkey = (Get-BitLockerVolume -MountPoint C).KeyProtector.RecoveryPassword

$HEADER = @" {Company Logo ASCII Art - Removed}
Performed by Tim Dang.
"@
$FOOTER = @"

--------------------------------------------------------------------------------------------------------------------------
-------------------------------------------- DO NOT COPY - DO NOT REPRODUCE ----------------------------------------------
--------------------------------------------------------------------------------------------------------------------------
"@

$DATE = Get-Date
Write-Output($DATE)
$NEW_LINE = [System.Environment]::NewLine
Write-Output($bitkey)
$content = $HEADER + $NEW_LINE + 'Date:' + $DATE + $NEW_LINE + $bitkey + $FOOTER
Write-Output($content)
Set-Content $file_name $content
get-Content $file_name

#Store the key to USB (Drive D).
try{
$destination  = 'D:\bitkeys\Bitlocker_{0}_{1}.txt' -f $username, $computername 
Write-Output($destination )
}
catch {
Write-Output($Error[0])
[System.Windows.MessageBox]::Show($Error[0],'Info','Ok','None')
}

$ENCODED-TOKEN = { Base64 encoded of your-api-key:X } 

$computer_name = [System.Net.DNS]::GetHostByName('').HostName
#Use get search to get display_id
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic $($ENCODED-TOKEN)")
$headers.Add("Content-Type", "application/json")
try{
$get_url = 'https://change-me.freshservice.com/api/v2/assets?search="name%3A%27' +$computer_name+ '%27"'
$response = Invoke-RestMethod $get_url -Method 'GET' -Headers $headers
$response | ConvertTo-Json
}
catch {
    Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
    Write-Output($Error[0])
    [System.Windows.MessageBox]::Show($Error[0],'Info','Ok','None')
}

# laptop_asset_type_int = 1400034XXXX (or the like).


Function Get-RedirectedUrl {
    Param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )

    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()

    If ($response.StatusCode -eq "Found")
    {
        $response.GetResponseHeader("Location")
    }
}
$serial_number = (Get-WmiObject win32_bios).serialnumber
$sn= Out-String -InputObject $serial_number


$url_lookup = "https://pcsupport.lenovo.com/us/en/warrantylookup?sn=" + $sn + "&upgrade&cid=ww:apps:x6bnuv&utm_source=Vantage&utm_medium=Native&utm_campaign=Warranty_Promo_Tile#/"
Write-Output($url_lookup)
$redirected_url = Get-RedirectedUrl -URL $url_lookup 
$full_redirected_url = 'https://pcsupport.lenovo.com' + $redirected_url
Write-Output($full_redirected_url)
$WebResponse = Invoke-WebRequest $full_redirected_url
$targeted_class_name = "text-capitalize"
$d = $WebResponse.ParsedHtml.body.getElementsByClassName('title-line')
$body = $WebResponse.ParsedHtml.body
#Write-Output($body.innerHTML)
#Write-Output($body.innerHTML.GetType())
#$body.innerHTML -match '(?<="RemainingDays":).*(?=,"EntireWarrantyPeriod":)'
$body.innerHTML -match '(?<="End":")\d{4}-\d{2}-\d{2}(?=","Status":)'
#.Matches.Value
$warranty_expiry_date = $Matches[0]
Write-Output($Matches[0])

$body.innerHTML -match '(?<="Start":")\d{4}-\d{2}-\d{2}(?=","End":")'
$acquisition_date = $Matches[0]
Write-Output($acquisition_date)
#$warranty_expiry_date = $WebResponse.ParsedHtml.body

#Please refer to your FreshService for integer represent for laptops. 
$body2 = @"
{	
	"usage_type": "permanent",
    "description": "Updated on $($DATE)",
	"type_fields": {
	    "product_change_me" : change_me,
        "asset_state_change_me" : "in use",
		"windowskey_change_me": "$($windows_key)",
        "bitlocker_recovery_change_me": "$($bitkey)",
        "warranty_expiry_date_change_me": "$($warranty_expiry_date)"
	}
}
"@
try {
    $PRODUCT_URL = "https://change_me.freshservice.com/api/v2/assets/" + $response.assets.display_id
    $response_update = Invoke-RestMethod $PRODUCT_URL -Method 'PUT' -Headers $headers -Body $body2
    $response_update | ConvertTo-Json 
}
catch {
    Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription
    Write-Output($Error[0])
    [System.Windows.MessageBox]::Show($Error[0],'Info','Ok','None')
}

# Show a prompt when finish!
try {
Copy-Item $file_name -Destination $destination
Remove-Item $file_name
#$wshell.Popup("Job Done Successfully",1,"Done",0x1)
[System.Windows.MessageBox]::Show('Successfully','Info','Ok','None')
}

catch {
Write-Output($Error[0])
[System.Windows.MessageBox]::Show($Error[0],'Info','Ok','None')
}
