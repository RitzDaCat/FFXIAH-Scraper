function get-WebPageTable{
param(
    [Parameter(Mandatory = $true)]

    [Microsoft.PowerShell.Commands.HtmlWebResponseObject] $WebRequest,

   

    [Parameter(Mandatory = $true)]

    [int] $TableNumber

)

## Extract the tables out of the web request

$tables = @($WebRequest.ParsedHtml.getElementsByTagName("TABLE"))

$table = $tables[$TableNumber]

$titles = @()

$rows = @($table.Rows)

## Go through all of the rows in the table

foreach($row in $rows)

{

    $cells = @($row.Cells)

   

    ## If we've found a table header, remember its titles

    if($cells[0].tagName -eq "TH")

    {

        $titles = @($cells | % { ("" + $_.InnerText).Trim() })

        continue

    }

    ## If we haven't found any table headers, make up names "P1", "P2", etc.

    if(-not $titles)

    {

        $titles = @(1..($cells.Count + 2) | % { "P$_" })

    }

    ## Now go through the cells in the the row. For each, try to find the

    ## title that represents that column and create a hashtable mapping those

    ## titles to content

    $resultObject = [Ordered] @{}

    for($counter = 0; $counter -lt $cells.Count; $counter++)

    {

        $title = $titles[$counter]

        if(-not $title) { continue }

       

        $resultObject[$title] = ("" + $cells[$counter].InnerText).Trim()

    }

    ## And finally cast that hashtable to a PSCustomObject

    [PSCustomObject] $resultObject

}

}

$allresults
$firstnumber = 1
while($firstnumber -lt 26)
{
    $secondnumber = 1
    while($secondnumber -lt 26)
    {
        $Uri = "https://www.ffxiah.com/search/item?q=" + [char](65 + $firstnumber) + [char](65 + $secondnumber)
    
        $R = Invoke-WebRequest -Uri $Uri -Body $form -Method Post
        $Form = $R.Forms[1]
        #Write-Host $form.Fields
        $Form.Fields["sid"]=17
        $Form.Fields["server"]="Siren"
        $Form.Fields["name"]="Siren"
        $Form.Fields["ffxi-main-server-select"]=17
        $response = Invoke-WebRequest -Uri $Uri -Body $form.Fields -Method Post
        $allresults += get-WebPageTable -WebRequest $response -TableNumber 3
        $secondnumber++
    }
    $firstnumber ++
}
$allresults | sort "Item Name" -Unique | Export-Csv C:\temp\test.csv -NoTypeInformation

