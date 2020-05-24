
$number = 5
try{
$Uri = "https://www.ffxiah.com/item/"+$number 
$Body = @{
    sid = 28
}
try{
$properRequest = Invoke-WebRequest -Uri $Uri -MaximumRedirection 0 -ErrorAction SilentlyContinue
}
catch{}
$newURL = "https://www.ffxiah.com" + $properRequest.Headers.Location
Invoke-WebRequest -Uri $newURL -Form $Body -Method Post 
}
catch{}
#Write-Host $form.Fields
<#
$Form.Fields["sid"]=$sid
$Form.Fields["server"]=$server
$Form.Fields["name"]=$server
$Form.Fields["ffxi-main-server-select"]=$sid
#>

#Invoke-WebRequest -Uri $location -Form $Body -Method Post