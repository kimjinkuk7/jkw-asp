# =============================================================================
#  jkw.com  -  Classic ASP 정적 점검 도구
# -----------------------------------------------------------------------------
#  Classic ASP 는 컴파일 단계가 없어 오류가 런타임에야 드러난다.
#  아래 7가지를 코드 수정 직후 미리 잡는다.
#
#    1. include 경로 유효성
#    2. 페이지 병합 후 최상위 Dim 중복            (800a0411)
#    3. Dim 이름 vs 전역 Function/Sub 이름 충돌   (800a0411)
#    4. 매개변수 이름 == 함수 이름                 (800a0411)
#    5. VBScript 내장 함수명과 충돌하는 전역 선언  (800a0411)
#    6. 주석 안의 include 지시자 (순환 include 유발)
#    7. UTF-8 인코딩 / CODEPAGE 지시자
#
#  경로 :  이 리포지토리는 자립형이다. assets / upload / db / docs 가 모두 안에 있다.
#          include 는 전부 상대경로(#include file=)라 "그 지시자가 적힌 파일의
#          디렉터리" 기준으로 푼다. (IIS 실측으로 확인한 ASP 의 해석 규칙)
#
#  -UrlBase :  이 스택이 서비스되는 URL 접두어.  App.Conf.asp 의 APP_ROOT 와 같아야 한다.
#              사이트 루트에 배포했으면 ''  /  하위 asp 폴더면 '/asp'
#
#  사용 :  powershell -ExecutionPolicy Bypass -File tools\CheckAsp.ps1
#          powershell -ExecutionPolicy Bypass -File tools\CheckAsp.ps1 -Http http://localhost:8080
#          (-Http 를 주면 전 페이지를 실제로 요청해 컴파일 오류까지 확인한다)
# =============================================================================
param(
    # ASP 애플리케이션 루트  (이 리포지토리의 최상위)
    [string]$Root = (Split-Path -Parent $PSScriptRoot),
    # 서비스 URL 접두어. 지정하지 않으면 App.Conf.asp 의 APP_ROOT 에서 읽는다.
    [string]$UrlBase = '',
    [string]$Http = ''
)

# --- APP_ROOT 자동 인식 : 소스와 검사 기준이 어긋나는 것을 막는다 --------------
$appConf = Join-Path $Root 'App.Conf.asp'
if ($UrlBase -eq '' -and (Test-Path $appConf)) {
    foreach ($l in [System.IO.File]::ReadAllLines($appConf)) {
        if ($l -match '(?i)^\s*Const\s+APP_ROOT\s*=\s*"([^"]*)"') { $UrlBase = $Matches[1]; break }
    }
}
$UrlBase = $UrlBase.TrimEnd('/')

$ErrorActionPreference = 'Stop'
$fail = 0

function Head($t) { Write-Host ''; Write-Host ('=== ' + $t + ' ===') -ForegroundColor Cyan }
function Ok($t)   { Write-Host ('  [OK] ' + $t) -ForegroundColor Green }
function Bad($t)  { Write-Host ('  [!!] ' + $t) -ForegroundColor Red; $script:fail++ }

$INTRINSIC = @(
 'Abs','Array','Asc','AscB','AscW','Atn','CBool','CByte','CCur','CDate','CDbl','Chr','ChrB','ChrW',
 'CInt','CLng','Cos','CreateObject','CSng','CStr','Date','DateAdd','DateDiff','DatePart','DateSerial',
 'DateValue','Day','Eval','Execute','Exp','Filter','Fix','FormatCurrency','FormatDateTime','FormatNumber',
 'FormatPercent','GetLocale','GetObject','GetRef','Hex','Hour','InputBox','InStr','InStrB','InStrRev','Int',
 'IsArray','IsDate','IsEmpty','IsNull','IsNumeric','IsObject','Join','LBound','LCase','Left','LeftB','Len',
 'LenB','Log','LTrim','Mid','MidB','Minute','Month','MonthName','MsgBox','Now','Oct','Randomize','Replace',
 'RGB','Right','RightB','Rnd','Round','RTrim','ScriptEngine','Second','SetLocale','Sgn','Sin','Space','Split',
 'Sqr','StrComp','String','StrReverse','Tan','Time','Timer','TimeSerial','TimeValue','Trim','TypeName',
 'UBound','UCase','VarType','Weekday','WeekdayName','Year')

$aspFiles = @(Get-ChildItem $Root -Recurse -Include *.asp -File)
$views    = @(Get-ChildItem (Join-Path $Root 'view') -Recurse -Include *.asp -File |
              Where-Object { $_.DirectoryName -notlike '*\view\include' })

# api\*.asp 도 요청되는 진입 페이지다 (React 스택이 쓰는 JSON API).
# 화면이 아니라 HTTP 스모크(-Http) 대상에서는 빼지만,
# "병합 후 Dim/Const 중복" 검사는 view 와 똑같이 받아야 한다.
$apiDir   = Join-Path $Root 'api'
$apis     = @()
if (Test-Path $apiDir) { $apis = @(Get-ChildItem $apiDir -Recurse -Include *.asp -File) }
$entries  = @($views) + @($apis)

# --- 한 줄에서 ASP 구분자를 제거해 순수 VBScript 만 남긴다 --------------------
function Strip($raw) {
    $s = $raw -replace '<%=', ' ' -replace '<%@[^%]*%>', ' ' -replace '<%', ' ' -replace '%>', ' '
    return $s.Trim()
}

$IncCache = @{}; $DeclCache = @{}

# include 경로를 실제 파일 경로로 푼다.
#   file=    그 지시자가 적힌 파일의 디렉터리 기준 (ASP 실측 확인)
#   virtual= 사이트 루트 기준. 이 리포지토리는 쓰지 않지만 남아 있으면 잡아낸다.
function Read-Includes($path) {
    if ($IncCache.ContainsKey($path)) { return $IncCache[$path] }
    $l = New-Object System.Collections.ArrayList
    if (Test-Path $path) {
        $here = Split-Path -Parent $path
        foreach ($line in [System.IO.File]::ReadAllLines($path)) {
            if ($line -match '#include\s+file\s*=\s*"([^"]+)"') {
                $p = Join-Path $here ($Matches[1] -replace '/','\')
                [void]$l.Add([System.IO.Path]::GetFullPath($p))
            }
            elseif ($line -match '#include\s+virtual\s*=\s*"([^"]+)"') {
                # 자립형 배치에서는 사이트 루트를 알 수 없다. 상대경로로 바꿔야 한다.
                [void]$l.Add('<<VIRTUAL>> ' + $Matches[1])
            }
        }
    }
    $IncCache[$path] = $l
    return $l
}

# 파일의 최상위(클래스/프로시저 밖) Dim / Const / Function / Sub 선언
function Read-Decls($path) {
    if ($DeclCache.ContainsKey($path)) { return $DeclCache[$path] }
    $out = New-Object System.Collections.ArrayList
    if (Test-Path $path) {
        $depth = 0
        foreach ($raw in [System.IO.File]::ReadAllLines($path)) {
            $line = Strip $raw
            if ($line.Length -eq 0 -or $line.StartsWith("'")) { continue }

            if ($line -match '(?i)^(Public\s+|Private\s+|Default\s+)*(Class|Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)') {
                # 주의 : $Matches 는 다음 -match 에서 덮어써진다. 먼저 꺼내둘 것.
                $blkKind = $Matches[2]
                $blkName = $Matches[3]
                if ($depth -eq 0 -and (($blkKind -ieq 'Function') -or ($blkKind -ieq 'Sub'))) {
                    [void]$out.Add([pscustomobject]@{ Kind='PROC'; Name=$blkName })
                }
                # 한 줄짜리 프로시저는 같은 줄에서 닫힌다.
                #   Public Function MyNick() : MyNick = Session(...) : End Function
                # 이때도 depth 를 올려버리면 그 뒤의 모든 최상위 선언이 "프로시저 안"으로
                # 잘못 판정되어 검사에서 빠진다. (실제로 이 때문에 Dim bi 와
                # Function BI() 의 충돌을 놓쳐 런타임 800a0411 이 났다)
                if ($line -notmatch '(?i)\bEnd\s+(Class|Function|Sub|Property)\b') { $depth++ }
                continue
            }
            if ($line -match '(?i)^End\s+(Class|Function|Sub|Property)\b') { if ($depth -gt 0) { $depth-- }; continue }
            if ($depth -ne 0) { continue }

            # 문자열 리터럴 제거. 값 안의 콤마를 선언 구분자로 오인하지 않기 위함.
            #   예) Const UPLOAD_DENY_EXT = "asp,php,config,inc"
            $line = $line -replace '"[^"]*"', '""'

            if ($line -match '(?i)^(Public\s+|Private\s+)?Const\s+(.+)$') {
                foreach ($s in ($Matches[2] -split ',')) {
                    if ($s.Trim() -match '^([A-Za-z_]\w*)') { [void]$out.Add([pscustomobject]@{ Kind='CONST'; Name=$Matches[1] }) }
                }
            }
            elseif ($line -match '(?i)^Dim\s+(.+)$') {
                foreach ($s in ((($Matches[1] -split ':')[0]) -split ',')) {
                    if ($s.Trim() -match '^([A-Za-z_]\w*)') { [void]$out.Add([pscustomobject]@{ Kind='DIM'; Name=$Matches[1] }) }
                }
            }
        }
    }
    $DeclCache[$path] = $out
    return $out
}

function Get-PageFiles($entry) {
    $order = New-Object System.Collections.ArrayList
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]'
    function Walk($p) {
        if (-not $seen.Add($p)) { return }
        foreach ($i in (Read-Includes $p)) { if ($i -notlike '<<VIRTUAL>>*') { Walk $i } }
        [void]$order.Add($p)
    }
    Walk $entry
    return $order
}

# =============================================================================
Head '1. include 경로 유효성'
$n = 0; $miss = @()
foreach ($f in $aspFiles) {
    foreach ($i in (Read-Includes $f.FullName)) {
        $n++
        if ($i -like '<<VIRTUAL>>*') {
            $miss += ($f.FullName.Replace($Root,'') + '  virtual include 가 남아 있음 -> ' + $i.Substring(12) + '  (상대경로 file= 로 바꿀 것)')
        }
        elseif (-not (Test-Path $i)) { $miss += ($f.FullName.Replace($Root,'') + ' -> ' + $i.Replace($Root,'')) }
    }
}
if ($miss.Count -eq 0) { Ok ("include " + $n + "건 전부 유효") } else { $miss | ForEach-Object { Bad $_ } }

# =============================================================================
Head '2~3. 최상위 Dim/Const 중복 및 전역 함수명 충돌'
$dupTotal = 0
foreach ($v in $entries) {
    $rows = New-Object System.Collections.ArrayList
    foreach ($f in (Get-PageFiles $v.FullName)) {
        foreach ($d in (Read-Decls $f)) {
            [void]$rows.Add([pscustomobject]@{ Key=$d.Name.ToLower(); Kind=$d.Kind; Name=$d.Name; File=$f.Replace($Root,'') })
        }
    }
    # 이름이 같으면 종류(Dim/Const/Function)가 달라도 VBScript 에서는 충돌이다
    foreach ($g in ($rows | Group-Object Key | Where-Object { $_.Count -gt 1 })) {
        $dupTotal++
        $detail = ($g.Group | ForEach-Object { $_.Kind + '@' + $_.File } | Select-Object -Unique) -join '  |  '
        Bad ($v.FullName.Replace($Root,'') + '  :  ' + $g.Group[0].Name + '   ' + $detail)
    }
}
if ($dupTotal -eq 0) { Ok ($entries.Count.ToString() + ' 페이지 모두 중복 없음 (view ' + $views.Count + ' + api ' + $apis.Count + ')') }

# =============================================================================
Head '4~5. 매개변수==함수명 / 내장 함수명 충돌'
$pc = 0
foreach ($f in $aspFiles) {
    $lines = [System.IO.File]::ReadAllLines($f.FullName)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = Strip $lines[$i]
        if ($line.StartsWith("'")) { continue }
        if ($line -match '(?i)^(Public\s+|Private\s+|Default\s+)*(Function|Sub|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*(\(([^)]*)\))?') {
            $nm = $Matches[3]; $ps = $Matches[5]
            if ($ps) {
                foreach ($p in ($ps -split ',')) {
                    $pn = ($p -replace '(?i)\bByVal\b|\bByRef\b','').Trim()
                    if ($pn -match '^([A-Za-z_]\w*)' -and $Matches[1].ToLower() -eq $nm.ToLower()) {
                        $pc++; Bad ($f.FullName.Replace($Root,'') + ':' + ($i+1) + '  매개변수가 함수명과 같음 -> ' + $nm)
                    }
                }
            }
        }
    }
}
# Function/Sub 를 내장 함수명으로 선언하면 오류. Dim 은 섀도잉이라 허용되지만
# 가독성을 해치므로 경고로만 표시한다.
foreach ($f in $aspFiles) {
    foreach ($d in (Read-Decls $f.FullName)) {
        if ($INTRINSIC -contains $d.Name) {
            if ($d.Kind -eq 'PROC') {
                $pc++; Bad ($f.FullName.Replace($Root,'') + '  내장 함수명과 충돌 -> ' + $d.Kind + ' ' + $d.Name)
            } else {
                Write-Host ('  [경고] ' + $f.FullName.Replace($Root,'') + '  내장 함수명을 가림 -> ' + $d.Kind + ' ' + $d.Name) -ForegroundColor Yellow
            }
        }
    }
}
if ($pc -eq 0) { Ok '없음' }

# =============================================================================
Head '6. 주석 안의 include 지시자'
$cc = 0
foreach ($f in $aspFiles) {
    $ln = 0
    foreach ($raw in [System.IO.File]::ReadAllLines($f.FullName)) {
        $ln++
        if ($raw.Trim().StartsWith("'") -and $raw -match '<!--\s*#include') {
            $cc++; Bad ($f.FullName.Replace($Root,'') + ':' + $ln + '  주석 안 include (순환 위험)')
        }
    }
}
if ($cc -eq 0) { Ok '없음' }

# =============================================================================
Head '7. UTF-8 / CODEPAGE'
$enc = New-Object System.Text.UTF8Encoding($false, $true)
$ec = 0
# .asp 도 .css/.js 도 전부 이 리포지토리 안에 있다 (공용 폴더 없음).
$encTargets = @(Get-ChildItem $Root -Recurse -Include *.asp -File)
$assetsDir  = Join-Path $Root 'assets'
if (Test-Path $assetsDir) { $encTargets += @(Get-ChildItem $assetsDir -Recurse -Include *.css,*.js -File) }
foreach ($f in $encTargets) {
    try { [void]$enc.GetString([System.IO.File]::ReadAllBytes($f.FullName)) }
    catch { $ec++; Bad ($f.FullName.Replace($Root,'') + '  UTF-8 이 아님') }
}
foreach ($v in $entries) {
    $first = ([System.IO.File]::ReadAllLines($v.FullName))[0]
    if ($first -notmatch 'CODEPAGE\s*=\s*"?65001') { $ec++; Bad ($v.FullName.Replace($Root,'') + '  1행에 CODEPAGE=65001 없음') }
}
if ($ec -eq 0) { Ok '전 파일 UTF-8 / 진입 페이지 CODEPAGE 정상' }

# =============================================================================
if ($Http -ne '') {
    Head ('8. HTTP 스모크 테스트  ' + $Http)
    $curl = Join-Path $env:windir 'system32\curl.exe'
    $good = 0
    # 실제 브라우저와 같은 UA 를 보낸다. UA 를 안 보내면 세션 지문 경로가 검사되지 않는다.
    $ua = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36'

    # 필수 파라미터가 있는 페이지. 없으면 앱이 정상적으로 404 를 돌려주므로 오탐이 된다.
    # 시드 데이터 기준값이다. 데이터를 바꿨으면 여기도 맞춰야 한다.
    # 키는 리포지토리 기준 상대 경로다 (UrlBase 는 요청할 때 앞에 붙인다).
    $QsMap = @{
        '/view/board/List.asp'     = '?bid=notice'
        '/view/board/View.asp'     = '?bid=notice&id=1'
        '/view/board/Write.asp'    = '?bid=free'
        '/view/board/PwCheck.asp'  = '?bid=guest&id=7'
        '/view/portfolio/View.asp' = '?id=1'
    }

    # 진입 스텁 :  <UrlBase>/  ->  302  ->  <UrlBase>/view/main/Default.asp
    $wantLoc  = $UrlBase + '/view/main/Default.asp'
    $rootRaw  = & $curl -s -i --max-time 20 -A $ua ($Http + $UrlBase + '/') 2>&1 | Out-String
    if ($rootRaw -match '(?im)^Location:\s*(\S+)') {
        $gotLoc = $Matches[1].Trim()
        if ($gotLoc.EndsWith($wantLoc)) { Ok ($UrlBase + '/  ->  ' + $gotLoc) }
        else { Bad ($UrlBase + '/  리다이렉트 대상이 다름 -> ' + $gotLoc + '  (기대 ' + $wantLoc + ')') }
    } else { Bad ($UrlBase + '/  리다이렉트 없음 (default.asp 확인)') }

    foreach ($v in $views) {
        $rel = $v.FullName.Replace($Root,'').Replace('\','/')
        $qs  = ''
        if ($QsMap.ContainsKey($rel)) { $qs = $QsMap[$rel] }
        # *_p.asp 는 POST + CSRF 전용 처리 페이지다. 단순 GET 은 가드가 막는 것이 정상이므로
        # 4xx 도 통과로 본다 (컴파일 오류라면 500 + 오류 텍스트가 나온다).
        $isProc = $rel -match '_p\.asp$'
        # 상태 코드까지 받는다. 본문만 보면 500 오류 화면(DB 접속 실패 등)을 정상으로 오판한다.
        $raw  = & $curl -s -w '|__STATUS__%{http_code}' --max-time 20 -A $ua ($Http + $UrlBase + $rel + $qs) 2>&1 | Out-String
        $code = ''
        $body = $raw
        $i = $raw.LastIndexOf('|__STATUS__')
        if ($i -ge 0) {
            $body = $raw.Substring(0, $i)
            $code = $raw.Substring($i + 11).Trim()
        }

        $probs = @()
        $okCode = if ($isProc) { '^(200|301|302|304|400|403|404)$' } else { '^(200|301|302|304)$' }
        if ($code -notmatch $okCode) { $probs += ('HTTP ' + $code) }
        if ($body -match 'Microsoft VBScript|800a0|800401|ASP 0') {
            $m = ($body -replace '<[^>]+>',' ' -replace '\s+',' ').Trim()
            $probs += $m.Substring(0, [Math]::Min(160, $m.Length))
        }

        # 상세 링크의 키가 0 이면 목록이 키 컬럼을 못 읽은 것이다.
        # BIGINT 컬럼은 ADO 에서 VT_DECIMAL 로 오고 IsNumeric() 이 이를 False 로 보기 때문에
        # 예전에 전 게시판의 View 링크가 id=0 으로 깨진 적이 있다. hidden input 의 id=0(신규글)은 정상이므로 href 만 본다.
        $zeroKey = [regex]::Matches($body, 'href="[^"]*(?:\?|&amp;|&)(?:id|fid)=0(?![0-9])')
        if ($zeroKey.Count -gt 0) { $probs += ('링크 키가 0 : ' + $zeroKey.Count + '건') }

        if ($probs.Count -gt 0) { Bad ($rel + '  ' + ($probs -join '  |  ')) } else { $good++ }
    }
    if ($good -eq $views.Count) { Ok ($good.ToString() + ' / ' + $views.Count + ' 페이지 정상 응답') }
    else { Bad ($good.ToString() + ' / ' + $views.Count + ' 페이지만 정상 응답') }
}

# =============================================================================
Write-Host ''
if ($fail -eq 0) { Write-Host '전체 통과' -ForegroundColor Green; exit 0 }
else { Write-Host ('문제 ' + $fail + '건') -ForegroundColor Red; exit 1 }
