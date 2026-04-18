$data = Get-Content transactions.json -Raw | ConvertFrom-Json
Write-Host "Count: $($data.transactions.Count)"
