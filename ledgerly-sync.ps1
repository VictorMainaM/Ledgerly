param(
  [string]$WorkbookPath = (Join-Path ([Environment]::GetFolderPath("Desktop")) "COMPANY TRANSACTIONS AND BOOKKEEPING.xlsx"),
  [int]$Port = 8765
)

$ErrorActionPreference = "Stop"
$script:ExcelApp = $null
$script:Workbook = $null
$script:TransactionHeaders = @("ID", "Number", "Date", "Type", "Party", "Description", "Item", "Quantity", "Unit", "Unit Price (KES)", "Amount (KES)", "Payment", "Reference", "Notes", "Created At", "Item ID")
$script:InventoryHeaders = @("ID", "Item Name", "Unit", "Opening Quantity", "Opening Unit Cost (KES)")
$script:TransportHeaders = @("ID", "Trip Number", "Load Carried", "Carried By", "Departure Date & Time", "Arrival Date & Time", "Total Mileage", "Mileage Unit", "Fuel Cost (KES)", "Additional Costs (KES)", "Additional Cost Details", "Load Value (KES)", "Payment", "Notes", "Created At")
$script:ProductionHeaders = @("ID", "Batch Number", "Date", "Finished Item ID", "Finished Item", "Output Quantity", "Material Item ID", "Raw Material", "Material Quantity", "Additional Costs (KES)", "Payment", "Notes", "Created At")
$script:OpeningBalanceHeaders = @("ID", "Balance Date", "Account Type", "Account Name", "Opening Amount (KES)")
$script:Token = [Guid]::NewGuid().ToString("N")

function Get-Workbook {
  if ($script:Workbook) { return $script:Workbook }
  if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
  }

  $fullPath = [IO.Path]::GetFullPath($WorkbookPath)
  $excel = $null
  try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
  } catch {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
  }
  $excel.Visible = $true

  $book = $null
  foreach ($candidate in $excel.Workbooks) {
    try {
      if ([IO.Path]::GetFullPath([string]$candidate.FullName) -ieq $fullPath) {
        $book = $candidate
        break
      }
    } catch { }
  }
  if (-not $book) {
    $book = $excel.Workbooks.Open($fullPath, 0, $false)
  }
  if ($book.ReadOnly) {
    throw "The workbook is open read-only. Close that copy, reopen it with editing enabled, then sync again."
  }

  $script:ExcelApp = $excel
  $script:Workbook = $book
  return $book
}

function Get-Sheet($Name, [bool]$Create) {
  $book = Get-Workbook
  $sheet = $null
  try { $sheet = $book.Worksheets.Item($Name) } catch { }
  if (-not $sheet -and $Create) {
    $sheet = $book.Worksheets.Add()
    $sheet.Name = $Name
  }
  return $sheet
}

function Get-LastRow($Sheet, [int]$Column = 1) {
  return [int]$Sheet.Cells.Item($Sheet.Rows.Count, $Column).End(-4162).Row
}

function Get-CellValue($Sheet, [int]$Row, [int]$Column) {
  return $Sheet.Cells.Item($Row, $Column).Value2
}

function Set-CellValue {
  param([object]$Sheet, [int]$Row, [int]$Column, [object]$Value)
  $valueType = if ($null -eq $Value) { "null" } else { $Value.GetType().FullName }
  try {
    if ($null -eq $Value) { $Sheet.Cells.Item($Row, $Column).ClearContents() }
    elseif ($Value -is [ValueType] -and $Value -isnot [bool]) {
      $formula = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
      $Sheet.Cells.Item($Row, $Column).Formula = $formula
    } else {
      $Sheet.Cells.Item($Row, $Column).Value2 = $Value
    }
  } catch {
    throw "Could not write $($Sheet.Name) row $Row column $Column ($valueType): $($_.Exception.Message)"
  }
}

function Test-EmptySheet($Sheet) {
  return [string]::IsNullOrWhiteSpace([string](Get-CellValue $Sheet 1 1))
}

function Assert-Headers($Sheet, $Expected, [string]$Name) {
  if (Test-EmptySheet $Sheet) { return }
  for ($column = 1; $column -le $Expected.Count; $column++) {
    $actual = [string](Get-CellValue $Sheet 1 $column)
    if ($actual -cne $Expected[$column - 1]) {
      throw "The $Name tab has unexpected headers. Restore the original column names before syncing."
    }
  }
}

function ConvertTo-OptionalNumber($Value, [string]$Description) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
  $number = 0.0
  if (-not [double]::TryParse([string]$Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
    throw "$Description must be a number."
  }
  if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { throw "$Description must be a finite number." }
  return $number
}

function ConvertTo-IsoDate($Value, [string]$Description) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { throw "$Description needs a date." }
  if ($Value -is [double] -or $Value -is [int] -or $Value -is [long]) {
    try { return [DateTime]::FromOADate([double]$Value).ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture) }
    catch { throw "$Description has an invalid Excel date." }
  }
  $text = ([string]$Value).Trim()
  $formats = @("yyyy-MM-dd", "M/d/yyyy", "MM/dd/yyyy", "d/M/yyyy", "dd/MM/yyyy")
  try {
    return [DateTime]::ParseExact($text, $formats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None).ToString("yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
  } catch {
    throw "$Description has an invalid date. Use yyyy-mm-dd or enter a date in an Excel date cell."
  }
}

function ConvertTo-IsoDateTime($Value, [string]$Description) {
  if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return "" }
  if ($Value -is [double] -or $Value -is [int] -or $Value -is [long]) {
    try { return [DateTime]::FromOADate([double]$Value).ToString("yyyy-MM-ddTHH:mm", [Globalization.CultureInfo]::InvariantCulture) }
    catch { throw "$Description has an invalid Excel date and time." }
  }
  $text = ([string]$Value).Trim()
  $formats = @(
    "yyyy-MM-ddTHH:mm", "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss",
    "M/d/yyyy H:mm", "M/d/yyyy HH:mm", "M/d/yyyy h:mm tt", "M/d/yyyy h:mm:ss tt",
    "MM/dd/yyyy H:mm", "MM/dd/yyyy HH:mm", "MM/dd/yyyy h:mm tt", "MM/dd/yyyy h:mm:ss tt",
    "d/M/yyyy H:mm", "d/M/yyyy HH:mm", "d/M/yyyy h:mm tt", "d/M/yyyy h:mm:ss tt",
    "dd/MM/yyyy H:mm", "dd/MM/yyyy HH:mm", "dd/MM/yyyy h:mm tt", "dd/MM/yyyy h:mm:ss tt"
  )
  foreach ($format in $formats) {
    $parsed = [DateTime]::MinValue
    if ([DateTime]::TryParseExact($text, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
      return $parsed.ToString("yyyy-MM-ddTHH:mm", [Globalization.CultureInfo]::InvariantCulture)
    }
  }
  throw "$Description has an invalid date and time. Use yyyy-mm-dd hh:mm or an Excel date/time cell."
}

function Read-Inventory {
  $script:ItemsById = @{}
  $sheet = Get-Sheet "LEDGERLY INVENTORY" $false
  if (-not $sheet -or (Test-EmptySheet $sheet)) { return ,@() }
  Assert-Headers $sheet $script:InventoryHeaders "LEDGERLY INVENTORY"
  $items = New-Object System.Collections.Generic.List[object]
  $ids = @{}
  $names = @{}
  $lastRow = Get-LastRow $sheet 2
  for ($row = 2; $row -le $lastRow; $row++) {
    $name = [string](Get-CellValue $sheet $row 2)
    $unit = [string](Get-CellValue $sheet $row 3)
    $quantityValue = Get-CellValue $sheet $row 4
    $costValue = Get-CellValue $sheet $row 5
    if ([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($unit) -and $null -eq $quantityValue -and $null -eq $costValue) { continue }
    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($unit)) { throw "Check the name and unit in LEDGERLY INVENTORY row $row." }
    $id = [string](Get-CellValue $sheet $row 1)
    if ([string]::IsNullOrWhiteSpace($id)) { $id = "excel-item-" + [Guid]::NewGuid().ToString("N") }
    if ($ids.ContainsKey($id)) { throw "Duplicate inventory ID in LEDGERLY INVENTORY row $row." }
    $key = $name.Trim().ToLowerInvariant()
    if ($names.ContainsKey($key)) { throw "Duplicate item name '$name' in LEDGERLY INVENTORY. Item names must be unique." }
    $quantity = ConvertTo-OptionalNumber $quantityValue "Opening quantity in inventory row $row"
    if ($null -eq $quantity -or $quantity -lt 0) { throw "Opening quantity in inventory row $row must be zero or greater." }
    $cost = ConvertTo-OptionalNumber $costValue "Opening cost in inventory row $row"
    if ($null -ne $cost -and $cost -lt 0) { throw "Opening cost in inventory row $row must be zero or greater." }
    $item = [pscustomobject]@{
      id = $id
      name = $name.Trim()
      unit = $unit.Trim()
      openingQuantity = $quantity
      openingCost = $cost
    }
    $items.Add($item)
    $ids[$id] = $true
    $names[$key] = $id
  }
  $script:ItemsById = @{}
  foreach ($item in $items) { $script:ItemsById[$item.id] = $item }
  return ,@($items.ToArray())
}

function Read-Transactions($Items) {
  $sheet = Get-Sheet "LEDGERLY TRANSACTIONS" $false
  if (-not $sheet -or (Test-EmptySheet $sheet)) { return ,@() }
  Assert-Headers $sheet $script:TransactionHeaders "LEDGERLY TRANSACTIONS"
  $itemByName = @{}
  foreach ($item in $Items) { $itemByName[$item.name.Trim().ToLowerInvariant()] = $item }
  $transactions = New-Object System.Collections.Generic.List[object]
  $ids = @{}
  $lastRow = Get-LastRow $sheet 3
  for ($row = 2; $row -le $lastRow; $row++) {
    $id = [string](Get-CellValue $sheet $row 1)
    $number = [string](Get-CellValue $sheet $row 2)
    $dateValue = Get-CellValue $sheet $row 3
    $type = ([string](Get-CellValue $sheet $row 4)).Trim().ToLowerInvariant()
    $party = [string](Get-CellValue $sheet $row 5)
    $description = [string](Get-CellValue $sheet $row 6)
    $itemName = [string](Get-CellValue $sheet $row 7)
    $quantityValue = Get-CellValue $sheet $row 8
    $unit = [string](Get-CellValue $sheet $row 9)
    $unitPriceValue = Get-CellValue $sheet $row 10
    $amountValue = Get-CellValue $sheet $row 11
    $payment = [string](Get-CellValue $sheet $row 12)
    $reference = [string](Get-CellValue $sheet $row 13)
    $notes = [string](Get-CellValue $sheet $row 14)
    $createdAt = [string](Get-CellValue $sheet $row 15)
    $itemId = [string](Get-CellValue $sheet $row 16)
    if ([string]::IsNullOrWhiteSpace($number) -and $null -eq $dateValue -and [string]::IsNullOrWhiteSpace($type) -and [string]::IsNullOrWhiteSpace($description) -and [string]::IsNullOrWhiteSpace($itemName) -and $null -eq $amountValue) { continue }
    if ($type -notin @("sale", "purchase", "expense", "income", "customer_receipt", "supplier_payment", "loan_received", "loan_repayment", "capital_introduced", "capital_withdrawal")) { throw "Transaction row $row has an unsupported Type." }
    if ($type -in @("sale", "purchase", "customer_receipt", "supplier_payment", "loan_received", "loan_repayment", "capital_introduced", "capital_withdrawal") -and [string]::IsNullOrWhiteSpace($party)) {
      throw "Transaction row $row needs an account name in the Party column."
    }
    if (($type -eq "sale" -or $type -eq "purchase") -and [string]::IsNullOrWhiteSpace($party)) {
      $requiredParty = if ($type -eq "sale") { "customer" } else { "supplier" }
      throw "Transaction row $row needs the $requiredParty in the Party column."
    }
    if ([string]::IsNullOrWhiteSpace($id)) { $id = "excel-txn-" + [Guid]::NewGuid().ToString("N") }
    if ($ids.ContainsKey($id)) { throw "Duplicate transaction ID in LEDGERLY TRANSACTIONS row $row." }
    $date = ConvertTo-IsoDate $dateValue "Transaction row $row"
    $quantity = ConvertTo-OptionalNumber $quantityValue "Quantity in transaction row $row"
    if ($null -eq $quantity) { $quantity = 0.0 }
    $unitPrice = ConvertTo-OptionalNumber $unitPriceValue "Unit price in transaction row $row"
    if ($null -eq $unitPrice) { $unitPrice = 0.0 }
    $amount = ConvertTo-OptionalNumber $amountValue "Amount in transaction row $row"
    if ($type -eq "sale" -or $type -eq "purchase") {
      if ($itemId -and $script:ItemsById.ContainsKey($itemId)) {
        $item = $script:ItemsById[$itemId]
      } else {
        $itemKey = $itemName.Trim().ToLowerInvariant()
        if (-not $itemByName.ContainsKey($itemKey)) { throw "Transaction row $row names an item not found in LEDGERLY INVENTORY: '$itemName'." }
        $item = $itemByName[$itemKey]
        $itemId = $item.id
      }
      if (-not $itemName) { $itemName = $item.name }
      if ([string]::IsNullOrWhiteSpace($unit)) { $unit = $item.unit }
      if ($null -eq $amount) { $amount = $quantity * $unitPrice }
      if ($quantity -le 0 -or $unitPrice -le 0) { throw "Transaction row $row needs a positive quantity and unit price." }
    } else {
      $itemId = ""
      if ($null -eq $amount) { throw "Transaction row $row needs an Amount (KES)." }
    }
    if ($amount -le 0) { throw "Transaction row $row needs a positive Amount (KES)." }
    if ([string]::IsNullOrWhiteSpace($description)) { $description = if ($itemName) { $itemName } else { $type } }
    if ([string]::IsNullOrWhiteSpace($payment)) { $payment = "Cash" }
    if ($type -in @("customer_receipt", "supplier_payment", "loan_received", "loan_repayment", "capital_introduced", "capital_withdrawal") -and $payment -notin @("Cash", "Bank", "M-Pesa")) {
      throw "Account movement in transaction row $row must use Cash, Bank, or M-Pesa, not Credit."
    }
    if ([string]::IsNullOrWhiteSpace($createdAt)) { $createdAt = [DateTime]::UtcNow.ToString("o") }
    $transaction = [pscustomobject]@{
      id = $id
      number = $number.Trim()
      date = $date
      type = $type
      party = $party.Trim()
      description = $description.Trim()
      itemId = $itemId
      quantity = $quantity
      unit = $unit.Trim()
      unitPrice = $unitPrice
      amount = $amount
      payment = $payment.Trim()
      reference = $reference.Trim()
      notes = $notes.Trim()
      createdAt = $createdAt
    }
    $transactions.Add($transaction)
    $ids[$id] = $true
  }

  $maxByYear = @{}
  foreach ($transaction in $transactions) {
    if ($transaction.number -match "^TX-(\d{4})-(\d+)$") {
      $year = $Matches[1]
      $sequence = [int]$Matches[2]
      if (-not $maxByYear.ContainsKey($year) -or $sequence -gt $maxByYear[$year]) { $maxByYear[$year] = $sequence }
    }
  }
  foreach ($transaction in $transactions) {
    if ([string]::IsNullOrWhiteSpace($transaction.number)) {
      $year = $transaction.date.Substring(0, 4)
      if (-not $maxByYear.ContainsKey($year)) { $maxByYear[$year] = 0 }
      $maxByYear[$year]++
      $transaction.number = "TX-$year-" + $maxByYear[$year].ToString("D4")
    }
  }
  return ,@($transactions.ToArray())
}

function Read-OpeningBalances {
  $empty = [pscustomobject]@{ asOfDate = ""; accounts = @() }
  $sheet = Get-Sheet "LEDGERLY OPENING BALANCES" $false
  if (-not $sheet -or (Test-EmptySheet $sheet)) { return $empty }
  Assert-Headers $sheet $script:OpeningBalanceHeaders "LEDGERLY OPENING BALANCES"
  $accounts = New-Object System.Collections.Generic.List[object]
  $ids = @{}
  $keys = @{}
  $balanceDate = ""
  $lastRow = 1
  for ($column = 1; $column -le $script:OpeningBalanceHeaders.Count; $column++) {
    $lastRow = [Math]::Max($lastRow, (Get-LastRow $sheet $column))
  }
  for ($row = 2; $row -le $lastRow; $row++) {
    $id = [string](Get-CellValue $sheet $row 1)
    $dateValue = Get-CellValue $sheet $row 2
    $type = ([string](Get-CellValue $sheet $row 3)).Trim().ToLowerInvariant()
    $name = ([string](Get-CellValue $sheet $row 4)).Trim()
    $amountValue = Get-CellValue $sheet $row 5
    if ([string]::IsNullOrWhiteSpace($id) -and $null -eq $dateValue -and [string]::IsNullOrWhiteSpace($type) -and [string]::IsNullOrWhiteSpace($name) -and $null -eq $amountValue) { continue }
    $date = ConvertTo-IsoDate $dateValue "Opening balance row $row"
    if ($balanceDate -and $balanceDate -ne $date) { throw "All opening balance rows must use the same Balance Date." }
    $balanceDate = $date
    if ($type -eq "settings" -and $id -eq "__settings__") { continue }
    if ($type -notin @("cash", "bank", "debtor", "creditor", "loan", "capital")) { throw "Opening balance row $row has an unsupported Account Type." }
    if ($type -eq "cash") { $name = "Cash & M-Pesa" }
    elseif ($type -eq "bank") { $name = "Bank" }
    if ([string]::IsNullOrWhiteSpace($name)) { throw "Opening balance row $row needs an Account Name." }
    $amount = ConvertTo-OptionalNumber $amountValue "Opening amount in row $row"
    if ($null -eq $amount -or $amount -lt 0) { throw "Opening amount in row $row must be a non-negative number." }
    if ([string]::IsNullOrWhiteSpace($id)) { $id = "excel-opening-" + [Guid]::NewGuid().ToString("N") }
    if ($ids.ContainsKey($id)) { throw "Duplicate opening balance ID in row $row." }
    $key = "$type`:$($name.ToLowerInvariant())"
    if ($keys.ContainsKey($key)) { throw "Duplicate opening balance for '$name' in LEDGERLY OPENING BALANCES." }
    $accounts.Add([pscustomobject]@{ id = $id; type = $type; name = $name; amount = $amount })
    $ids[$id] = $true
    $keys[$key] = $true
  }
  return [pscustomobject]@{ asOfDate = $balanceDate; accounts = @($accounts.ToArray()) }
}

function Read-WorkbookState {
  $items = Read-Inventory
  $transactions = Read-Transactions $items
  $productions = Read-Productions $items
  $transports = Read-Transports
  $openingBalances = Read-OpeningBalances
  return [pscustomobject]@{ items = @($items); transactions = @($transactions); productions = @($productions); transports = @($transports); openingBalances = $openingBalances }
}

function Read-Transports {
  $sheet = Get-Sheet "LEDGERLY TRANSPORT" $false
  if (-not $sheet -or (Test-EmptySheet $sheet)) { return ,@() }
  Assert-Headers $sheet $script:TransportHeaders "LEDGERLY TRANSPORT"
  $transports = New-Object System.Collections.Generic.List[object]
  $ids = @{}
  $numbers = @{}
  $lastRow = 1
  for ($column = 1; $column -le $script:TransportHeaders.Count; $column++) {
    $lastRow = [Math]::Max($lastRow, (Get-LastRow $sheet $column))
  }
  for ($row = 2; $row -le $lastRow; $row++) {
    $id = [string](Get-CellValue $sheet $row 1)
    $number = [string](Get-CellValue $sheet $row 2)
    $load = [string](Get-CellValue $sheet $row 3)
    $carrier = [string](Get-CellValue $sheet $row 4)
    $departureValue = Get-CellValue $sheet $row 5
    $arrivalValue = Get-CellValue $sheet $row 6
    $mileageValue = Get-CellValue $sheet $row 7
    $mileageUnit = ([string](Get-CellValue $sheet $row 8)).Trim().ToLowerInvariant()
    $fuelValue = Get-CellValue $sheet $row 9
    $additionalValue = Get-CellValue $sheet $row 10
    $additionalDetails = [string](Get-CellValue $sheet $row 11)
    $loadValueInput = Get-CellValue $sheet $row 12
    $payment = [string](Get-CellValue $sheet $row 13)
    $notes = [string](Get-CellValue $sheet $row 14)
    $createdAt = [string](Get-CellValue $sheet $row 15)
    if ([string]::IsNullOrWhiteSpace($id) -and [string]::IsNullOrWhiteSpace($number) -and [string]::IsNullOrWhiteSpace($load) -and [string]::IsNullOrWhiteSpace($carrier) -and $null -eq $departureValue -and $null -eq $arrivalValue -and $null -eq $mileageValue -and [string]::IsNullOrWhiteSpace($mileageUnit) -and $null -eq $fuelValue -and $null -eq $additionalValue -and [string]::IsNullOrWhiteSpace($additionalDetails) -and $null -eq $loadValueInput -and [string]::IsNullOrWhiteSpace($payment) -and [string]::IsNullOrWhiteSpace($notes) -and [string]::IsNullOrWhiteSpace($createdAt)) { continue }
    if ([string]::IsNullOrWhiteSpace($load) -or [string]::IsNullOrWhiteSpace($carrier)) { throw "Check the load and carrier in LEDGERLY TRANSPORT row $row." }
    $departure = ConvertTo-IsoDateTime $departureValue "Departure in transport row $row"
    if (-not $departure) { throw "Transport row $row needs a departure date and time." }
    $arrival = ConvertTo-IsoDateTime $arrivalValue "Arrival in transport row $row"
    if ($arrival -and $arrival -lt $departure) { throw "Arrival in transport row $row must be after departure." }
    $mileage = ConvertTo-OptionalNumber $mileageValue "Mileage in transport row $row"
    $fuelCost = ConvertTo-OptionalNumber $fuelValue "Fuel cost in transport row $row"
    $additionalCosts = ConvertTo-OptionalNumber $additionalValue "Additional costs in transport row $row"
    $loadValue = ConvertTo-OptionalNumber $loadValueInput "Load value in transport row $row"
    if ($null -eq $mileage) { $mileage = 0.0 }
    if ($null -eq $fuelCost) { $fuelCost = 0.0 }
    if ($null -eq $additionalCosts) { $additionalCosts = 0.0 }
    if ($null -eq $loadValue) { $loadValue = 0.0 }
    if ($mileage -lt 0 -or $fuelCost -lt 0 -or $additionalCosts -lt 0 -or $loadValue -lt 0) { throw "Mileage and costs in transport row $row must be zero or greater." }
    if ($mileageUnit -notin @("km", "miles")) { throw "Mileage unit in transport row $row must be km or miles." }
    if ([string]::IsNullOrWhiteSpace($payment)) { $payment = "Cash" }
    if ($payment -notin @("Cash", "Bank", "M-Pesa", "Credit")) { throw "Payment in transport row $row must be Cash, Bank, M-Pesa, or Credit." }
    if ([string]::IsNullOrWhiteSpace($id)) { $id = "excel-trip-" + [Guid]::NewGuid().ToString("N") }
    if ($ids.ContainsKey($id)) { throw "Duplicate trip ID in LEDGERLY TRANSPORT row $row." }
    $trip = [pscustomobject]@{
      id = $id
      number = $number.Trim()
      load = $load.Trim()
      carrier = $carrier.Trim()
      departure = $departure
      arrival = $arrival
      mileage = $mileage
      mileageUnit = $mileageUnit
      fuelCost = $fuelCost
      additionalCosts = $additionalCosts
      additionalCostDetails = $additionalDetails.Trim()
      loadValue = $loadValue
      payment = $payment.Trim()
      notes = $notes.Trim()
      createdAt = $createdAt
    }
    $transports.Add($trip)
    $ids[$id] = $true
  }
  foreach ($trip in $transports) {
    if ([string]::IsNullOrWhiteSpace($trip.number)) {
      $year = $trip.departure.Substring(0, 4)
      $prefix = "TR-$year-"
      $max = 0
      foreach ($existing in $transports) {
        if ($existing.number -match ("^TR-" + [regex]::Escape($year) + "-(\d+)$")) { $max = [Math]::Max($max, [int]$Matches[1]) }
      }
      do { $max++; $trip.number = $prefix + $max.ToString("D4") } while ($numbers.ContainsKey($trip.number))
    }
    if ($numbers.ContainsKey($trip.number)) { throw "Duplicate trip number '$($trip.number)' in LEDGERLY TRANSPORT." }
    $numbers[$trip.number] = $true
    if ([string]::IsNullOrWhiteSpace($trip.createdAt)) { $trip.createdAt = [DateTime]::UtcNow.ToString("o") }
  }
  return ,@($transports.ToArray())
}

function Resolve-ProductionItem($ItemId, $ItemName, $ItemsById, $ItemsByName, [string]$Description) {
  $idItem = $null
  $namedItem = $null
  if (-not [string]::IsNullOrWhiteSpace($ItemId) -and $ItemsById.ContainsKey($ItemId)) {
    $idItem = $ItemsById[$ItemId]
  }
  if (-not [string]::IsNullOrWhiteSpace($ItemName)) {
    $nameKey = $ItemName.Trim().ToLowerInvariant()
    if ($ItemsByName.ContainsKey($nameKey)) { $namedItem = $ItemsByName[$nameKey] }
  }
  if ($namedItem -and $idItem -and $namedItem.id -ne $idItem.id) { return $namedItem }
  if ($idItem) { return $idItem }
  if ($namedItem) { return $namedItem }
  throw "$Description does not match an item in LEDGERLY INVENTORY."
}

function Read-Productions($Items) {
  $sheet = Get-Sheet "LEDGERLY PRODUCTION" $false
  if (-not $sheet -or (Test-EmptySheet $sheet)) { return ,@() }
  Assert-Headers $sheet $script:ProductionHeaders "LEDGERLY PRODUCTION"
  $itemsById = @{}
  $itemsByName = @{}
  foreach ($item in $Items) {
    $itemsById[[string]$item.id] = $item
    $itemsByName[$item.name.Trim().ToLowerInvariant()] = $item
  }
  $groups = New-Object System.Collections.Generic.List[object]
  $groupsByKey = @{}
  $lastRow = 1
  for ($column = 1; $column -le $script:ProductionHeaders.Count; $column++) {
    $lastRow = [Math]::Max($lastRow, (Get-LastRow $sheet $column))
  }
  for ($row = 2; $row -le $lastRow; $row++) {
    $hasValue = $false
    for ($column = 1; $column -le $script:ProductionHeaders.Count; $column++) {
      $cell = Get-CellValue $sheet $row $column
      if ($null -ne $cell -and -not [string]::IsNullOrWhiteSpace([string]$cell)) { $hasValue = $true; break }
    }
    if (-not $hasValue) { continue }

    $id = ([string](Get-CellValue $sheet $row 1)).Trim()
    $number = ([string](Get-CellValue $sheet $row 2)).Trim()
    $dateValue = Get-CellValue $sheet $row 3
    $outputItemId = ([string](Get-CellValue $sheet $row 4)).Trim()
    $outputItemName = ([string](Get-CellValue $sheet $row 5)).Trim()
    $outputQuantityValue = Get-CellValue $sheet $row 6
    $materialItemId = ([string](Get-CellValue $sheet $row 7)).Trim()
    $materialName = ([string](Get-CellValue $sheet $row 8)).Trim()
    $materialQuantityValue = Get-CellValue $sheet $row 9
    $additionalCostsValue = Get-CellValue $sheet $row 10
    $payment = ([string](Get-CellValue $sheet $row 11)).Trim()
    $notes = ([string](Get-CellValue $sheet $row 12)).Trim()
    $createdAt = ([string](Get-CellValue $sheet $row 13)).Trim()
    if ([string]::IsNullOrWhiteSpace($id) -and [string]::IsNullOrWhiteSpace($number)) {
      throw "Production row $row needs a batch ID or batch number so its material lines can be grouped."
    }
    $key = if (-not [string]::IsNullOrWhiteSpace($id)) { "id:$id" } else { "number:$number" }
    if (-not $groupsByKey.ContainsKey($key)) {
      $date = ConvertTo-IsoDate $dateValue "Production row $row"
      $outputItem = Resolve-ProductionItem $outputItemId $outputItemName $itemsById $itemsByName "Finished item in production row $row"
      $outputQuantity = ConvertTo-OptionalNumber $outputQuantityValue "Output quantity in production row $row"
      if ($null -eq $outputQuantity -or $outputQuantity -le 0) { throw "Production row $row needs a positive finished output quantity." }
      $additionalCosts = ConvertTo-OptionalNumber $additionalCostsValue "Additional costs in production row $row"
      if ($null -eq $additionalCosts) { $additionalCosts = 0.0 }
      if ($additionalCosts -lt 0) { throw "Additional production costs in row $row must be zero or greater." }
      if ([string]::IsNullOrWhiteSpace($payment)) { $payment = "Cash" }
      if ($payment -notin @("Cash", "Bank", "M-Pesa", "Credit")) { throw "Payment in production row $row must be Cash, Bank, M-Pesa, or Credit." }
      if ([string]::IsNullOrWhiteSpace($id)) { $id = "excel-production-" + [Guid]::NewGuid().ToString("N") }
      if ([string]::IsNullOrWhiteSpace($createdAt)) { $createdAt = [DateTime]::UtcNow.ToString("o") }
      $production = [pscustomobject]@{
        id = $id
        number = $number
        date = $date
        outputItemId = [string]$outputItem.id
        outputQuantity = $outputQuantity
        materials = New-Object System.Collections.Generic.List[object]
        additionalCosts = $additionalCosts
        payment = $payment
        notes = $notes
        createdAt = $createdAt
      }
      $groups.Add($production)
      $groupsByKey[$key] = $production
    } else {
      $production = $groupsByKey[$key]
      if (-not [string]::IsNullOrWhiteSpace($number) -and $production.number -and $production.number -ne $number) { throw "Production row $row has a different batch number for the same batch ID." }
      if ($null -ne $dateValue -and -not [string]::IsNullOrWhiteSpace([string]$dateValue) -and $production.date -ne (ConvertTo-IsoDate $dateValue "Production row $row")) { throw "Production batch $($production.number) has inconsistent dates across its material rows." }
      if (-not [string]::IsNullOrWhiteSpace($outputItemId) -or -not [string]::IsNullOrWhiteSpace($outputItemName)) {
        $outputItem = Resolve-ProductionItem $outputItemId $outputItemName $itemsById $itemsByName "Finished item in production row $row"
        if ($production.outputItemId -ne $outputItem.id) { throw "Production batch $($production.number) has inconsistent finished items across its material rows." }
      }
      if ($null -ne $outputQuantityValue -and -not [string]::IsNullOrWhiteSpace([string]$outputQuantityValue)) {
        $outputQuantity = ConvertTo-OptionalNumber $outputQuantityValue "Output quantity in production row $row"
        if ($production.outputQuantity -ne $outputQuantity) { throw "Production batch $($production.number) has inconsistent output quantities across its material rows." }
      }
      if ($null -ne $additionalCostsValue -and -not [string]::IsNullOrWhiteSpace([string]$additionalCostsValue)) {
        $additionalCosts = ConvertTo-OptionalNumber $additionalCostsValue "Additional costs in production row $row"
        if ($production.additionalCosts -ne $additionalCosts) { throw "Production batch $($production.number) has inconsistent additional costs across its material rows." }
      }
      if (-not [string]::IsNullOrWhiteSpace($payment) -and $production.payment -ne $payment) { throw "Production batch $($production.number) has inconsistent payment methods across its material rows." }
      if (-not [string]::IsNullOrWhiteSpace($notes) -and $production.notes -and $production.notes -ne $notes) { throw "Production batch $($production.number) has inconsistent notes across its material rows." }
    }

    $production = $groupsByKey[$key]
    if ($null -eq $materialQuantityValue -or [string]::IsNullOrWhiteSpace([string]$materialQuantityValue) -or ([string]::IsNullOrWhiteSpace($materialItemId) -and [string]::IsNullOrWhiteSpace($materialName))) {
      throw "Production row $row needs a raw material and its positive quantity."
    }
    $materialItem = Resolve-ProductionItem $materialItemId $materialName $itemsById $itemsByName "Raw material in production row $row"
    $materialQuantity = ConvertTo-OptionalNumber $materialQuantityValue "Material quantity in production row $row"
    if ($null -eq $materialQuantity -or $materialQuantity -le 0) { throw "Production row $row needs a positive raw-material quantity." }
    $production.materials.Add([pscustomobject]@{ itemId = [string]$materialItem.id; quantity = $materialQuantity })
  }

  $productions = New-Object System.Collections.Generic.List[object]
  $maxByYear = @{}
  foreach ($production in $groups) {
    if ($production.number -match "^PR-(\d{4})-(\d+)$") {
      $year = $Matches[1]
      $sequence = [int]$Matches[2]
      if (-not $maxByYear.ContainsKey($year) -or $sequence -gt $maxByYear[$year]) { $maxByYear[$year] = $sequence }
    }
  }
  foreach ($production in $groups) {
    if ([string]::IsNullOrWhiteSpace($production.number)) {
      $year = $production.date.Substring(0, 4)
      if (-not $maxByYear.ContainsKey($year)) { $maxByYear[$year] = 0 }
      $maxByYear[$year]++
      $production.number = "PR-$year-" + $maxByYear[$year].ToString("D4")
    }
    $production.materials = @($production.materials.ToArray())
    $productions.Add($production)
  }
  return ,@($productions.ToArray())
}

function Ensure-Sheet($Name, $Headers) {
  $sheet = Get-Sheet $Name $true
  Assert-Headers $sheet $Headers $Name
  if (Test-EmptySheet $sheet) {
    for ($column = 1; $column -le $Headers.Count; $column++) {
      Set-CellValue -Sheet $sheet -Row 1 -Column $column -Value $Headers[$column - 1]
    }
    $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item(1,$Headers.Count)).Font.Bold = $true
    $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item(1,$Headers.Count)).Interior.Color = 15792366
    $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item(1,$Headers.Count)).AutoFilter() | Out-Null
    $sheet.Columns.Item(1).Hidden = $true
    if ($Name -eq "LEDGERLY TRANSACTIONS") {
      $sheet.Columns.Item(15).Hidden = $true
      $sheet.Columns.Item(16).Hidden = $true
    } elseif ($Name -eq "LEDGERLY TRANSPORT") {
      $sheet.Columns.Item(1).Hidden = $true
      $sheet.Columns.Item(15).Hidden = $true
    }
    $sheet.Columns.AutoFit() | Out-Null
  }
  return $sheet
}

function Write-WorkbookState($State) {
  $items = @($State.items)
  $transactions = @($State.transactions)
  $productions = @($State.productions)
  $transports = @($State.transports)
  $openingBalances = $State.openingBalances
  if ($null -eq $openingBalances) { $openingBalances = [pscustomobject]@{ asOfDate = ""; accounts = @() } }
  $openingAccounts = @($openingBalances.accounts)
  $itemById = @{}
  foreach ($item in $items) { $itemById[[string]$item.id] = $item }

  $inventorySheet = Ensure-Sheet "LEDGERLY INVENTORY" $script:InventoryHeaders
  $transactionSheet = Ensure-Sheet "LEDGERLY TRANSACTIONS" $script:TransactionHeaders
  $productionSheet = Ensure-Sheet "LEDGERLY PRODUCTION" $script:ProductionHeaders
  $transportSheet = Ensure-Sheet "LEDGERLY TRANSPORT" $script:TransportHeaders
  $openingBalanceSheet = Ensure-Sheet "LEDGERLY OPENING BALANCES" $script:OpeningBalanceHeaders
  foreach ($sheet in @($inventorySheet, $transactionSheet, $transportSheet, $openingBalanceSheet)) {
    $clearThrough = if ($sheet.Name -eq "LEDGERLY INVENTORY") { 5 } elseif ($sheet.Name -eq "LEDGERLY TRANSACTIONS") { 16 } elseif ($sheet.Name -eq "LEDGERLY TRANSPORT") { 15 } else { 5 }
    $lastRow = 1
    for ($column = 1; $column -le $clearThrough; $column++) {
      $lastRow = [Math]::Max($lastRow, (Get-LastRow $sheet $column))
    }
    if ($lastRow -gt 1) {
      $sheet.Range($sheet.Cells.Item(2,1), $sheet.Cells.Item($lastRow,$clearThrough)).ClearContents()
    }
  }
  $lastProductionRow = 1
  for ($column = 1; $column -le $script:ProductionHeaders.Count; $column++) {
    $lastProductionRow = [Math]::Max($lastProductionRow, (Get-LastRow $productionSheet $column))
  }
  if ($lastProductionRow -gt 1) {
    $productionSheet.Range($productionSheet.Cells.Item(2,1), $productionSheet.Cells.Item($lastProductionRow,$script:ProductionHeaders.Count)).ClearContents()
  }

  $row = 2
  foreach ($item in $items) {
    $values = @($item.id, $item.name, $item.unit, $item.openingQuantity, $item.openingCost)
    for ($column = 1; $column -le $values.Count; $column++) {
      $cellValue = $values[$column - 1]
      Set-CellValue -Sheet $inventorySheet -Row $row -Column $column -Value $cellValue
    }
    $row++
  }
  $inventorySheet.Columns.Item(2).ColumnWidth = 28
  $inventorySheet.Columns.Item(3).ColumnWidth = 16
  $inventorySheet.Columns.Item(4).NumberFormat = "#,##0.###"
  $inventorySheet.Columns.Item(5).NumberFormat = "#,##0.00"

  $row = 2
  foreach ($transaction in $transactions) {
    $itemName = ""
    if ($transaction.itemId -and $itemById.ContainsKey([string]$transaction.itemId)) {
      $itemName = [string]$itemById[[string]$transaction.itemId].name
    }
    $values = @(
      $transaction.id, $transaction.number, $transaction.date, $transaction.type, $transaction.party,
      $transaction.description, $itemName, $transaction.quantity, $transaction.unit, $transaction.unitPrice,
      $transaction.amount, $transaction.payment, $transaction.reference, $transaction.notes, $transaction.createdAt, $transaction.itemId
    )
    for ($column = 1; $column -le $values.Count; $column++) {
      $cellValue = $values[$column - 1]
      Set-CellValue -Sheet $transactionSheet -Row $row -Column $column -Value $cellValue
    }
    $row++
  }
  $transactionSheet.Columns.Item(2).ColumnWidth = 17
  $transactionSheet.Columns.Item(3).NumberFormat = "yyyy-mm-dd"
  $transactionSheet.Columns.Item(5).ColumnWidth = 22
  $transactionSheet.Columns.Item(6).ColumnWidth = 26
  $transactionSheet.Columns.Item(7).ColumnWidth = 22
  $transactionSheet.Columns.Item(8).NumberFormat = "#,##0.###"
  $transactionSheet.Columns.Item(10).NumberFormat = "#,##0.00"
  $transactionSheet.Columns.Item(11).NumberFormat = "#,##0.00"

  $row = 2
  foreach ($production in $productions) {
    $finishedName = ""
    if ($production.outputItemId -and $itemById.ContainsKey([string]$production.outputItemId)) {
      $finishedName = [string]$itemById[[string]$production.outputItemId].name
    }
    foreach ($material in @($production.materials)) {
      $materialName = ""
      if ($material.itemId -and $itemById.ContainsKey([string]$material.itemId)) {
        $materialName = [string]$itemById[[string]$material.itemId].name
      }
      $values = @(
        $production.id, $production.number, $production.date, $production.outputItemId, $finishedName,
        $production.outputQuantity, $material.itemId, $materialName, $material.quantity,
        $production.additionalCosts, $production.payment, $production.notes, $production.createdAt
      )
      for ($column = 1; $column -le $values.Count; $column++) {
        Set-CellValue -Sheet $productionSheet -Row $row -Column $column -Value $values[$column - 1]
      }
      $row++
    }
  }
  $productionSheet.Columns.Item(2).ColumnWidth = 17
  $productionSheet.Columns.Item(3).NumberFormat = "yyyy-mm-dd"
  $productionSheet.Columns.Item(5).ColumnWidth = 24
  $productionSheet.Columns.Item(6).NumberFormat = "#,##0.###"
  $productionSheet.Columns.Item(8).ColumnWidth = 24
  $productionSheet.Columns.Item(9).NumberFormat = "#,##0.###"
  $productionSheet.Columns.Item(10).NumberFormat = "#,##0.00"

  $row = 2
  foreach ($trip in $transports) {
    $values = @(
      $trip.id, $trip.number, $trip.load, $trip.carrier, $trip.departure, $trip.arrival, $trip.mileage,
      $trip.mileageUnit, $trip.fuelCost, $trip.additionalCosts, $trip.additionalCostDetails, $trip.loadValue,
      $trip.payment, $trip.notes, $trip.createdAt
    )
    for ($column = 1; $column -le $values.Count; $column++) {
      $cellValue = $values[$column - 1]
      Set-CellValue -Sheet $transportSheet -Row $row -Column $column -Value $cellValue
    }
    $row++
  }
  $transportSheet.Columns.Item(2).ColumnWidth = 17
  $transportSheet.Columns.Item(3).ColumnWidth = 32
  $transportSheet.Columns.Item(4).ColumnWidth = 22
  $transportSheet.Columns.Item(5).NumberFormat = "yyyy-mm-dd hh:mm"
  $transportSheet.Columns.Item(6).NumberFormat = "yyyy-mm-dd hh:mm"
  $transportSheet.Columns.Item(7).NumberFormat = "#,##0.###"
  $transportSheet.Columns.Item(9).NumberFormat = "#,##0.00"
  $transportSheet.Columns.Item(10).NumberFormat = "#,##0.00"
  $transportSheet.Columns.Item(12).NumberFormat = "#,##0.00"

  $row = 2
  foreach ($account in $openingAccounts) {
    $name = [string]$account.name
    if ($account.type -eq "cash") { $name = "Cash & M-Pesa" }
    elseif ($account.type -eq "bank") { $name = "Bank" }
    $values = @($account.id, $openingBalances.asOfDate, $account.type, $name, $account.amount)
    for ($column = 1; $column -le $values.Count; $column++) {
      Set-CellValue -Sheet $openingBalanceSheet -Row $row -Column $column -Value $values[$column - 1]
    }
    $row++
  }
  if ($openingAccounts.Count -eq 0 -and $openingBalances.asOfDate) {
    $values = @("__settings__", $openingBalances.asOfDate, "settings", "", $null)
    for ($column = 1; $column -le $values.Count; $column++) {
      Set-CellValue -Sheet $openingBalanceSheet -Row $row -Column $column -Value $values[$column - 1]
    }
  }
  $openingBalanceSheet.Columns.Item(2).NumberFormat = "yyyy-mm-dd"
  $openingBalanceSheet.Columns.Item(4).ColumnWidth = 28
  $openingBalanceSheet.Columns.Item(5).NumberFormat = "#,##0.00"
  $script:Workbook.Save()
  return [pscustomobject]@{ ok = $true; itemCount = $items.Count; transactionCount = $transactions.Count; productionCount = $productions.Count; transportCount = $transports.Count; openingBalanceCount = $openingAccounts.Count }
}

function Read-Request($Stream) {
  $headerBytes = New-Object System.Collections.Generic.List[byte]
  $one = New-Object byte[] 1
  $marker = ""
  while ($marker -ne "`r`n`r`n") {
    if ($headerBytes.Count -gt 16384) { throw "Request headers are too large." }
    $read = $Stream.Read($one, 0, 1)
    if ($read -eq 0) { throw "The client closed the connection early." }
    $headerBytes.Add($one[0])
    $marker = ([Text.Encoding]::ASCII.GetString($headerBytes.ToArray()))
    if ($marker.Length -gt 3) { $marker = $marker.Substring([Math]::Max(0, $marker.Length - 4)) }
  }
  $headerText = [Text.Encoding]::ASCII.GetString($headerBytes.ToArray())
  $lines = $headerText -split "`r`n"
  $requestLine = $lines[0].Split(" ")
  $headers = @{}
  foreach ($line in $lines | Select-Object -Skip 1) {
    if ($line -match "^([^:]+):\s*(.*)$") { $headers[$Matches[1].ToLowerInvariant()] = $Matches[2] }
  }
  $length = 0
  if ($headers.ContainsKey("content-length")) { $length = [int]$headers["content-length"] }
  if ($length -gt 1048576) { throw "Request body exceeds the 1 MB limit." }
  $bodyBytes = New-Object byte[] $length
  $offset = 0
  while ($offset -lt $length) {
    $count = $Stream.Read($bodyBytes, $offset, $length - $offset)
    if ($count -eq 0) { throw "The client closed the request body early." }
    $offset += $count
  }
  $body = $null
  if ($length -gt 0) { $body = [Text.Encoding]::UTF8.GetString($bodyBytes) | ConvertFrom-Json }
  return [pscustomobject]@{
    method = $requestLine[0]
    path = $requestLine[1]
    origin = [string]$headers["origin"]
    body = $body
  }
}

function Send-Response($Stream, [int]$Status, [string]$Reason, [string]$Body) {
  $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
  $header = "HTTP/1.1 $Status $Reason`r`nContent-Type: application/json; charset=utf-8`r`nContent-Length: $($bytes.Length)`r`nAccess-Control-Allow-Origin: null`r`nAccess-Control-Allow-Methods: POST, OPTIONS`r`nAccess-Control-Allow-Headers: Content-Type`r`nVary: Origin`r`nConnection: close`r`n`r`n"
  $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)
  if ($bytes.Length -gt 0) { $Stream.Write($bytes, 0, $bytes.Length) }
  $Stream.Flush()
}

if ($Port -lt 1024 -or $Port -gt 65535) { throw "Port must be between 1024 and 65535." }
if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
$listener = [Net.Sockets.TcpListener]::new([Net.IPAddress]::Loopback, $Port)
try {
  $listener.Start()
  Write-Host "Ledgerly Excel connector is running on this computer only." -ForegroundColor Green
  Write-Host "Workbook: $([IO.Path]::GetFullPath($WorkbookPath))"
  Write-Host "One-time sync token: $script:Token" -ForegroundColor Yellow
  Write-Host "Paste the token into Ledgerly. Keep this window open; press Ctrl+C to stop."
  while ($true) {
    $client = $listener.AcceptTcpClient()
    try {
      $stream = $client.GetStream()
      try {
        $request = Read-Request $stream
        if ($request.origin -and $request.origin -cne "null") {
          Send-Response $stream 403 "Forbidden" '{"error":"Only the local Ledgerly file is allowed to use this connector."}'
        } elseif ($request.method -eq "OPTIONS" -and $request.path -eq "/sync") {
          Send-Response $stream 204 "No Content" ""
        } elseif ($request.method -ne "POST" -or $request.path -ne "/sync") {
          Send-Response $stream 404 "Not Found" '{"error":"Unknown connector endpoint."}'
        } elseif ($null -eq $request.body -or [string]$request.body.token -cne $script:Token) {
          Send-Response $stream 401 "Unauthorized" '{"error":"The connector token is incorrect or expired. Copy the current token from the PowerShell window."}'
        } else {
          try {
            if ($request.body.action -eq "read") {
              $result = Read-WorkbookState
            } elseif ($request.body.action -eq "write") {
              if ($null -eq $request.body.state) { throw "Sync data is missing." }
              $result = Write-WorkbookState $request.body.state
            } else {
              throw "Unsupported connector action."
            }
            $json = ConvertTo-Json -InputObject $result -Depth 10 -Compress
            Send-Response $stream 200 "OK" $json
          } catch {
            $message = $_.Exception.Message
            Write-Host "Sync request failed: $message" -ForegroundColor Red
            Write-Host $_.ScriptStackTrace -ForegroundColor DarkRed
            $json = ConvertTo-Json -InputObject @{ error = $message } -Compress
            Send-Response $stream 400 "Bad Request" $json
          }
        }
      } finally {
        $stream.Dispose()
      }
    } catch {
      Write-Host "Connection error: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
      $client.Close()
    }
  }
} finally {
  $listener.Stop()
}
