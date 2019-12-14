$configJson = Get-Content "config.json" | ConvertFrom-Json
$config = $configJson[0]

$users = Get-ADUser -Filter * -SearchBase $config.SearchPath -Server $config.ADServer | select -ExpandProperty SamAccountName

$count_qr = 0;
$count_noqr = 0;
foreach ($user in $users) {
    $user_obj = Get-ADuser $user -Properties * -Server $config.ADServer
    if ($user_obj.description){
        $count_qr++
        $qr_code_file = $config.QRCodesPath + "\" + $user + ".png"
        New-QRCodeVCard `
            -Name $user_obj.description `
            -Company $(if ($user_obj.company -or $user_obj.department) {"$($user_obj.company), $($user_obj.department)"} Else { " " }) `
            -URL  $config.websiteURL `
            -Tel $config.companyTel `
            -Mobile $(if($user_obj.mobile) { $user_obj.mobile } else {" "}) `
            -City $(if($user_obj.l){$user_obj.l} else { " " }) `
            -Adr $(if($user_obj.streetAddress){$user_obj.streetAddress} else{" "}) `
            -Title $(if($user_obj.title) { $user_obj.title } else {" "}) `
            -Email $(if($user_obj.mail){$user_obj.mail} else {" "}) `
            -Width 10 `
            -OutPath $qr_code_file
            Write-Host "$count_qr > $user -> QR Code Saved to $qr_code_file" -f green
    }
    else {
        $count_noqr++
        Write-Host "$count_noqr > $user -> No description value, QR Code can't be created"  -f red
    }
}

Write-Host "==============================" -f yellow
Write-Host "SUMMARY" -f yellow
Write-Host "==============================" -f yellow
Write-Host "QR Codes Created: $count_qr" -f green
Write-Host "QR Codes Not Created: $count_noqr" -f red
