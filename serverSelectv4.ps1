# Install the module on demand
$startTime = Get-Date
Write-Host $startTime
If (-not (Get-Module -ErrorAction Ignore -ListAvailable PowerHTML)) {
  Write-Verbose "Installing PowerHTML module for the current user..."
  Install-Module PowerHTML -ErrorAction Stop
}
Import-Module -ErrorAction Stop PowerHTML
# Parse the HTML file into an HTML DOM.

$global:emptyCSV = @()
function ProcessData([int]$FromIndex, [int]$ToIndex)
{
    $results = $FromIndex..$ToIndex | ForEach-Object -Parallel{
    try{
    $number = $_
    $URI = "https://www.ffxiah.com/item/" + $number
    $htmlDom = ConvertFrom-Html -URI $URI
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
    if(!($PRICE -eq 0))
    {
        $itemData = @(

        [pscustomobject]@{

        ID = $ID
        NAME = $NAME
        PRICE = $PRICE
        STOCK  = $STOCK
        SPD = $SOLDDAY

        }
        )
        $output = $itemData
        $emptyCSV += $itemData
    }
    <#
    if($_ % 500 -eq 0)
    {
        #garbage collect attempt?
        Write-Host "GBC @" $_
        [system.gc]::Collect()
    }
    #>
    }
    catch{
        $output = $_
    }
    @($output)

    } -ThrottleLimit 8

    return $results
}@($output) #added this

#start using parallel function?
$FromIndex = 0      # Arrays start at 0
$ToIndex = 1499    # Our initial batch goes to index 1499
$Increment = 1500   # We increment each batch by 1500 items
$End = $false       # Bool for whether we're at the end of the FilesToProcess array
$LastItemIndex = 30000 # The index of the last item in the array

do {
    
    Write-Host "[+] Processing from $FromIndex to $ToIndex" -ForegroundColor Green
    $results = ProcessData -FromIndex $FromIndex -ToIndex $ToIndex
    $emptyCSV+=$results
    Write-Host "[+] Running Garbage Collection" -ForegroundColor Red
    [system.gc]::Collect() # Manual garbage collection following parallel processing of a batch

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
$emptyCSV | Out-GridView -Title "FFXIAH ITEMS"