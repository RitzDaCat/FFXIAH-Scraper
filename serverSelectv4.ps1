# Install the module on demand
$startTime = Get-Date
Write-Host $startTime
If (-not (Get-Module -ErrorAction Ignore -ListAvailable PowerHTML)) {
  Write-Verbose "Installing PowerHTML module for the current user..."
  Install-Module PowerHTML -ErrorAction Stop
}
Import-Module -ErrorAction Stop PowerHTML
# Parse the HTML file into an HTML DOM.

#OUTSIDE VARIABLES
$outputPath = "C:\Repos\FFXIAH-Scraper\items.csv"
#change line 61 to change servers
<#
            <option value="28">Asura</option>
            <option value="1">Bahamut</option>
            <option value="25">Bismarck</option>
            <option value="6">Carbuncle</option>
            <option value="23">Cerberus</option>
            <option value="7">Fenrir</option>
            <option value="27">Lakshmi</option>
            <option value="11">Leviathan</option>
            <option value="12">Odin</option>
            <option value="5">Phoenix</option>
            <option value="16">Quetzalcoatl</option>
            <option value="20">Ragnarok</option>
            <option value="2">Shiva</option>
            <option value="17">Siren</option>
            <option value="8">Sylph</option>
            <option value="9">Valefor</option>
            <optgroup label="Inactive">
                <option value="10">Alexander</option>
                <option value="15">Caitsith</option>
                <option value="14">Diabolos</option>
                <option value="30">Fairy</option>
                <option value="22">Garuda</option>
                <option value="19">Gilgamesh</option>
                <option value="32">Hades</option>
                <option value="13">Ifrit</option>
                <option value="24">Kujata</option>
                <option value="29">Midgardsormr</option>
                <option value="21">Pandemonium</option>
                <option value="4">Ramuh</option>
                <option value="31">Remora</option>
                <option value="26">Seraph</option>
                <option value="3">Titan</option>
                <option value="18">Unicorn</option>
#>
#PARALLEL CONFIG
$FromIndex = 1      # Arrays start at 0
$ToIndex = 1499  # Our initial batch goes to index 1499
$Increment = 1500  # We increment each batch by 1500 items
$End = $false       # Bool for whether we're at the end of the FilesToProcess array
$LastItemIndex = 29500 # The index of the last item in the array


$global:emptyCSV = @()
function ProcessData([int]$FromIndex, [int]$ToIndex)
{
    #inside parallel variables
    $results = $FromIndex..$ToIndex | ForEach-Object -Parallel{
    try{
        $number = $_
        $Uri = "https://www.ffxiah.com/item/"+$number 
        $Body = @{
            sid = 28
        }
        $properRequest = Invoke-WebRequest -Method HEAD -Uri $Uri
        #write-host $properRequest.BaseResponse.RequestMessage.RequestUri.AbsolutePath
        
        $newURL = "https://www.ffxiah.com" + $properRequest.BaseResponse.RequestMessage.RequestUri.AbsolutePath
        $response = Invoke-WebRequest -Uri $newURL -Form $Body -Method Post 
    #$htmlDom = ConvertFrom-Html -URI $URI
    $htmlDom = ConvertFrom-Html -Content $response.Content
    $node = $htmlDom.SelectSingleNode("//body")
    # Find a specific table by its column names, using an XPath
    # query to iterate over all tables.""

    #ITEMINFO
    $ID = $number
    $STOCK = 0
    $SOLDDAY = 0
    $PRICE = 0
    $NAME = "Blank"
    if(!($null -eq $node))
    {
        foreach ($item in $node.SelectNodes("//span[@class='" + "stock" + "']"))
        {
            $STOCK = $item.InnerText;
        }
        foreach ($item2 in $node.SelectNodes("//span[@id='" + "sales-per-day" + "']"))
        {
            $split = $item2.ParentNode.ParentNode.InnerText.Split(" ")
            $SOLDDAY = $split[0]
            if($SOLDDAY -eq "")
            {
                $SOLDDAY = 0
            }
        }
        <#
        $selectPath = '//*[@id="tbl-main"]/table/tbody/tr/td[1]/table/tbody/tr[7]/td[2]'
        $path2 = '/html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/table/tbody/tr[7]/td[2]'
        NAMEPATH = /html/body/table/tbody/tr[2]/td[2]/table/tbody/tr/td[1]/table[1]/tbody/tr[1]/td[2]/span[1]/span/span
        #>
        foreach($item3 in $node.SelectNodes("//span"))
        {
            #Write-Host $item3.XPath
            if($item3.XPath.Contains("/tr[7]/td[2]"))
            {
                $PRICE = $item3.InnerText
                if(!($PRICE-match '^[0-9]+$'))
                {
                    $PRICE = 0
                }
            }
        }

        foreach ($item4 in $node.SelectNodes("//span[@class='" + "item-name" + "']"))
        {
            $NAME = $item4.InnerText.Trim().Replace("&nbsp;","").TrimEnd()
        }
        <#
        Write-Host "START"
        Write-Host "ID:" $ID
        Write-Host "PRICE:" $PRICE
        Write-Host "STOCK:" $STOCK
        Write-Host "SOLD PER DAY:" $SOLDDAY
        Write-Host "NAME:" $NAME
        #>
    }
    if(!($PRICE -eq 0) -and !($null -eq $ID))
    {
        $itemData = @(

        [pscustomobject]@{

        itemid = $ID
        name = $NAME
        price01 = $PRICE
        stock01  = $STOCK
        rate01 = $SOLDDAY
        }
        )
        $output = $itemData
        Remove-Variable itemData
        #$emptyCSV += $itemData
    }
    <#
    if($_ % 500 -eq 0)
    {
        #garbage collect attempt?
        Write-Host "GBC @" $_
        [system.gc]::Collect()
    }
    #>
    Remove-Variable htmlDom 
    Remove-Variable properRequest
    Remove-Variable node
    Remove-Variable response
    }
    catch{
        $output = $_
    }
    
    @($output)


    } -ThrottleLimit 6

    return $results
}@($output) #added this



do {
    
    Write-Host "[+] Processing from $FromIndex to $ToIndex" -ForegroundColor Green
    $results = ProcessData -FromIndex $FromIndex -ToIndex $ToIndex
    $emptyCSV+=$results
    Write-Host "[+] Running Garbage Collection" -ForegroundColor Red
    #[system.gc]::Collect() # Manual garbage collection following parallel processing of a batch
    [System.GC]::GetTotalMemory($true) | out-null
    # We increment the FromIndex and ToIndex variables to set them for the next batch
    $FromIndex = $ToIndex + 1
    $ToIndex = $ToIndex + $Increment

    # If the ToIndex value exceeds the index of the last item of the array, we set it to the last item 
    # and flip the $End flag to end batch processing. 
    if ($ToIndex -ge $LastItemIndex) {
        $ToIndex = $LastItemIndex
        $End = $true
    }

} while ($End -ne $true)
$endTime = Get-Date
Write-Host $endTime
$emptyCSV | Select-Object itemid,name,price01,stock01 -Unique | Where-Object itemid | Export-Csv -Path $outputPath -Force
#$emptycsv | Select-Object * | ConvertTo-Csv | % {$_ -replace '"', ""} | Out-File $outputpath