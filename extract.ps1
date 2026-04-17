$text = Get-Content 'C:\Users\fezkongs\.gemini\antigravity\brain\d04d3a78-ccba-49c1-980a-7166e793c845\.system_generated\steps\281\content.md' -Raw
$jsonStr = $text -replace '(?s)^.*?---.*?\n\n', ''
$obj = $jsonStr | ConvertFrom-Json

$expenseItems = $obj.plans | Where-Object { $_.Type -eq 'Expense' }
Write-Output ("Total Expense Items: " + $expenseItems.Length)

if ($expenseItems.Length -gt 0) {
    $expenseItems | Select-Object -First 3 | ConvertTo-Json
}
