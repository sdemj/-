# 단원별 게임 데이터 만들기
# '단원별 게임 단어 정리(1025).xlsx'를 읽어 게임마다 unit-data.js 를 만들고(단어\단원데이터\게임 이름\),
# 단원별 게임 폴더가 있으면 그 폴더에도 unit-data.js 와 단원 주소 폴더(3-1/ 등)를 만든다.
# 엑셀을 고친 뒤 이 파일을 다시 실행하면 게임에 반영된다.
#   실행: PowerShell 에서  powershell -ExecutionPolicy Bypass -File "단원데이터_만들기.ps1"
# (이 파일은 UTF-8 BOM 으로 저장해야 한글이 깨지지 않는다)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$xlsx = Join-Path $here '단원별 게임 단어 정리(1025).xlsx'

# 게임 이름(엑셀) → 게임 폴더
$targets = [ordered]@{
  'Word Race'          = 'word-race-unit'
  'Pick the Word'      = 'picktheword-unit'
  'Word catch! catch!' = 'wordcatchcatch-unit'
  # 다른 게임을 단원별로 만들면 아래처럼 한 줄씩 추가한다 (게임 이름은 엑셀 '단원별 게임' 시트와 똑같이)
  # 'Save the Ice Cream' = 'icecream-unit'
  # 'Anagram Game'       = 'anagram-unit'
  # 'Whack-a-Mole'       = 'whack-a-mole-unit'
  # 'Unjumble Game'      = 'unjumble-unit'
}

function Read-Sheet($wb, $name) {
  $v = $wb.Worksheets.Item($name).UsedRange.Value2
  $rows = @()
  for ($i = 2; $i -le $v.GetLength(0); $i++) {
    $r = @(); for ($j = 1; $j -le $v.GetLength(1); $j++) { $r += [string]$v[$i, $j] }
    if (($r -join '').Trim()) { $rows += ,$r }
  }
  return $rows
}

$xl = New-Object -ComObject Excel.Application; $xl.Visible = $false; $xl.DisplayAlerts = $false
try {
  $wb = $xl.Workbooks.Open($xlsx, 0, $true)
  $units = Read-Sheet $wb '단원별 게임'   # 학년, 단원, 주소, 게임1~3, ...
  $words = Read-Sheet $wb '게임별 단어'   # 게임, 학년, 단원, 주소, 개수, 사용 단어, ...
  $means = Read-Sheet $wb '단어 뜻'       # 게임, 주소, 영어, 뜻, 확인 여부
  $wb.Close($false)
} finally { $xl.Quit() }

$utf8 = New-Object Text.UTF8Encoding $false
function Esc($s) { return ($s -replace '\\', '\\' -replace '"', '\"' -replace "`r?`n", ' ') }

# 엑셀에 나오는 모든 게임을 만든다.
#  - 항상: 단어\단원데이터\<게임 이름>\unit-data.js  (다른 선생님께 그 게임 파일만 보내 줄 때 쓴다)
#  - 위 $targets 에 폴더가 적혀 있고 그 폴더가 있으면: 그 폴더에도 unit-data.js 와 단원 주소 폴더를 만든다
$games = @(); foreach ($u in $units) { foreach ($g in $u[3], $u[4], $u[5]) { if ($g -and $games -notcontains $g) { $games += $g } } }
$shareRoot = Join-Path $here '단원데이터'

foreach ($game in $games) {
  $dir = $null
  if ($targets.Contains($game)) {
    $dir = Join-Path $root $targets[$game]
    if (-not (Test-Path $dir)) { Write-Warning "폴더 없음: $dir"; $dir = $null }
  }

  $lines = @()
  foreach ($u in $units) {
    $key = $u[2]
    if (@($u[3], $u[4], $u[5]) -notcontains $game) { continue }
    $row = $words | Where-Object { $_[0] -eq $game -and $_[3] -eq $key } | Select-Object -First 1
    if (-not $row) { continue }
    # Unjumble 은 문장(줄바꿈으로 구분), 나머지는 단어(쉼표로 구분)
    $sep = if ($game -eq 'Unjumble Game') { "`n" } else { ',' }
    $list = $row[5] -split $sep | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $items = foreach ($w in $list) {
      $m = $means | Where-Object { $_[0] -eq $game -and $_[1] -eq $key -and $_[2] -eq $w } | Select-Object -First 1
      $ko = if ($m) { $m[3] } else { '' }
      '{{"en":"{0}","ko":"{1}"}}' -f (Esc $w), (Esc $ko)
    }
    $lines += '    {{"key":"{0}","grade":{1},"title":"{2}","words":[{3}]}}' -f $key, [int]$u[0], (Esc $u[1]), ($items -join ',')

    # 단원 주소 폴더: 3-1/index.html → ../index.html?unit=3-1 로 이동 (내 컴퓨터에서 파일로 열어도 되게 index.html 까지 적는다)
    if (-not $dir) { continue }
    $ud = Join-Path $dir $key
    New-Item -ItemType Directory -Force $ud | Out-Null
    $html = "<!doctype html><html lang=""ko""><head><meta charset=""utf-8""><title>$game $key</title>" +
            "<meta http-equiv=""refresh"" content=""0; url=../index.html?unit=$key"">" +
            "<script>location.replace('../index.html?unit=$key');</script></head>" +
            "<body><a href=""../index.html?unit=$key"">게임으로 이동</a></body></html>"
    [IO.File]::WriteAllText((Join-Path $ud 'index.html'), $html, $utf8)
  }

  $js = "/* 단원별 단어 데이터 - '단어\단원데이터_만들기.ps1'이 엑셀에서 자동으로 만든 파일입니다. 직접 고치지 말고 엑셀을 고친 뒤 다시 만드세요. */`n" +
        "window.UNIT_DATA = {`n  ""game"": ""$game"",`n  ""units"": [`n" + ($lines -join ",`n") + "`n  ]`n};`n"
  $share = Join-Path $shareRoot $game
  New-Item -ItemType Directory -Force $share | Out-Null
  [IO.File]::WriteAllText((Join-Path $share 'unit-data.js'), $js, $utf8)
  if ($dir) { [IO.File]::WriteAllText((Join-Path $dir 'unit-data.js'), $js, $utf8) }
  "{0}: 단원 {1}개 → 단원데이터\{0}{2}" -f $game, $lines.Count, $(if ($dir) { " + $($targets[$game])" } else { '' })
}
