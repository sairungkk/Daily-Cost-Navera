$ErrorActionPreference = 'Stop'
$base = 'http://localhost:3000'

function Step($label, $script) {
  try {
    & $script
  } catch {
    Write-Output ("FAIL " + $label + " : " + $_.Exception.Message)
  }
}

# 1. login admin
$login = Invoke-RestMethod -Uri "$base/api/login" -Method Post -ContentType 'application/json' -Body '{"password":"admin001405"}'
if (-not $login.success) { Write-Output 'FAIL login admin'; exit 1 }
Write-Output ("PASS login admin role=" + $login.role)
$H = @{ 'x-auth-token' = $login.token; 'Content-Type' = 'application/json' }

# 2. set goal Reveira House 100,000
$g = Invoke-RestMethod -Uri "$base/api/upsell/goal" -Method Post -Headers $H -Body '{"year":2026,"month":9,"outlet":"Reveira House","monthly":100000}'
Write-Output ("PASS set goal: " + $g.message)

# 3. add entry
$e = Invoke-RestMethod -Uri "$base/api/upsell/entry" -Method Post -Headers $H -Body '{"date":"2026-09-17","outlet":"Reveira House","amount":4500,"note":"API smoke test"}'
Write-Output ("PASS add entry: " + $e.message)

# 4. upsert same date = replace amount
$e2 = Invoke-RestMethod -Uri "$base/api/upsell/entry" -Method Post -Headers $H -Body '{"date":"2026-09-17","outlet":"Reveira House","amount":5200,"note":"upsert test"}'
Write-Output ("PASS upsert entry: " + $e2.message)

# 5. GET and verify
$d = Invoke-RestMethod -Uri "$base/api/upsell?year=2026&month=9" -Method Get -Headers $H
$gh = $d.data.goals.PSObject.Properties['Reveira House'].Value
$cnt = @($d.data.entries).Count
$amount = if ($cnt -gt 0) { $d.data.entries[0].amount } else { -1 }
Write-Output ("VERIFY goal monthly=" + $gh.monthly + " daily[0]=" + $gh.daily[0] + " dailyCount=" + $gh.daily.Count + " (expect 100000 / 3333.33 / 30)")
Write-Output ("VERIFY entries=" + $cnt + " amount=" + $amount + " (expect 1 / 5200)")

# 6. invalid outlet rejected
try {
  $bad = Invoke-RestMethod -Uri "$base/api/upsell/goal" -Method Post -Headers $H -Body '{"year":2026,"month":9,"outlet":"Fake Outlet","monthly":1000}'
  Write-Output ("FAIL invalid outlet accepted: " + $bad.message)
} catch {
  Write-Output "PASS invalid outlet rejected (HTTP error)"
}

# 7. staff token: GET ok, POST goal forbidden
$ls = Invoke-RestMethod -Uri "$base/api/login" -Method Post -ContentType 'application/json' -Body '{"password":"1234"}'
$Hs = @{ 'x-auth-token' = $ls.token; 'Content-Type' = 'application/json' }
$gs = Invoke-RestMethod -Uri "$base/api/upsell?year=2026&month=9" -Method Get -Headers $Hs
Write-Output ("PASS staff GET upsell: success=" + $gs.success)
try {
  $ps = Invoke-RestMethod -Uri "$base/api/upsell/goal" -Method Post -Headers $Hs -Body '{"year":2026,"month":9,"outlet":"Reveira House","monthly":5000}'
  Write-Output ("FAIL staff could set goal: " + $ps.message)
} catch {
  Write-Output "PASS staff POST goal forbidden (HTTP error)"
}

# 8. no token -> 401
try {
  Invoke-RestMethod -Uri "$base/api/upsell?year=2026&month=9" -Method Get | Out-Null
  Write-Output "FAIL no-token GET accepted"
} catch {
  Write-Output "PASS no-token GET rejected (HTTP error)"
}

# 9. cleanup: delete entry + clear goal
$eId = if ($cnt -gt 0) { $d.data.entries[0].id } else { 0 }
$del = Invoke-RestMethod -Uri "$base/api/upsell/entry" -Method Delete -Headers $H -Body ('{"id":' + $eId + '}')
$gc = Invoke-RestMethod -Uri "$base/api/upsell/goal" -Method Post -Headers $H -Body '{"year":2026,"month":9,"outlet":"Reveira House","monthly":0}'
$d2 = Invoke-RestMethod -Uri "$base/api/upsell?year=2026&month=9" -Method Get -Headers $H
Write-Output ("PASS cleanup: delete=" + $del.success + " goalClear=" + $gc.success + " entriesNow=" + $d2.data.entries.Count + " goalsNow=" + ($d2.data.goals.PSObject.Properties | Measure-Object).Count + " (expect 0 / 0)")
Write-Output 'DONE'
