$ModuleManifestName = 'Vinum.DocumentFormat.psd1'
$ModuleManifestPath = "${PSScriptRoot}\..\${ModuleManifestName}"

Import-Module $ModuleManifestPath

class Order {
    [string] $OrderDate
    [string] $Region
    [string] $Rep
    [string] $Item
    [string] $Units
    [string] $UnitCost
    [string] $Total

    Order (
        [string] $OrderDate,
        [string] $Region,
        [string] $Rep,
        [string] $Item,
        [string] $Units,
        [string] $UnitCost,
        [string] $Total
    )
    {
        $this.OrderDate = $OrderDate
        $this.Region = $Region
        $this.Rep = $Rep
        $this.Item = $Item
        $this.Units = $Units
        $this.UnitCost = $UnitCost
        $this.Total = $Total
    }
}

Describe 'Cmdlet : Convert-ExcelSheetToCSV Tests' {
    It 'Passes Convert Sample Data' {
        # prepare data
        $orders = @(
            [Order]::new("43471", "East", "Jones", "Pencil", "95", "1.99", "189.05")
            [Order]::new("43488", "Central", "Kivell", "Binder", "50", "19.989999999999998", "999.49999999999989")
            [Order]::new("43505", "Central", "Jardine", "Pencil", "36", "4.99", "179.64000000000001")
        )
        $refData = ($orders | ConvertTo-Csv -NoTypeInformation | ConvertFrom-Csv)
        $diffData = Convert-ExcelSheetToCSV -Path "${PSScriptRoot}\data\sampledata.xlsx" -SheetName "SalesOrders" -StartRow 1 | ConvertFrom-Csv

        #$diffData | Should -Be - $refData
    }
}