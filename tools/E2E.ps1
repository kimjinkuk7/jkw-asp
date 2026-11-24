# =============================================================================
#  jkw.com  -  Classic ASP 기능(E2E) 검증
# -----------------------------------------------------------------------------
#  CheckAsp.ps1 이 "페이지가 오류 없이 뜨는가"를 본다면,
#  이 스크립트는 "기능이 실제로 동작하는가"를 본다.
#
#    1. 로그인 (CSRF 토큰 포함 POST)
#    2. 로그인 상태 유지 / 마이페이지 / 관리자 화면 접근
#    3. 게시글 등록 (multipart + 첨부파일)
#    4. 목록 / 상세에 반영 확인
#    5. 첨부 다운로드 (바이트 일치 확인)
#    6. 댓글 등록
#    7. 후원 주문 생성 (AJAX) + 모의 승인
#    8. 게시글 삭제
#    9. 문의 등록
#   10. CSRF 토큰 없는 POST 가 차단되는지
#   11. 로그아웃
#
#  사용 :  powershell -ExecutionPolicy Bypass -File tools\E2E.ps1
#          powershell -ExecutionPolicy Bypass -File tools\E2E.ps1 -Base http://localhost:8080
# =============================================================================
param(
    [string]$Base   = 'http://localhost:8080',
    # ASP 스택의 URL 접두어. App.Conf.asp 의 VIEW_ROOT 와 같아야 한다.
    [string]$View   = '/asp/view',
    [string]$Id     = 'admin',
    [string]$Pw     = 'admin1234!'
)

$ErrorActionPreference = 'Continue'
$fail = 0
$UA   = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36'

function Head($t) { Write-Host ''; Write-Host ('=== ' + $t + ' ===') -ForegroundColor Cyan }
function Ok($t)   { Write-Host ('  [OK] ' + $t) -ForegroundColor Green }
function Bad($t)  { Write-Host ('  [!!] ' + $t) -ForegroundColor Red; $script:fail++ }

# 쿠키(세션)를 유지하는 웹 세션
$sess = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$sess.UserAgent = $UA

function Get-Page($path) {
    try {
        return (Invoke-WebRequest -Uri ($Base + $path) -WebSession $sess -UseBasicParsing -TimeoutSec 30)
    } catch {
        Bad ('GET ' + $path + ' 실패 : ' + $_.Exception.Message)
        return $null
    }
}

# 폼 POST. 한글이 깨지지 않도록 본문을 UTF-8 로 직접 퍼센트 인코딩한다.
# (해시테이블을 그대로 넘기면 PowerShell 이 어떤 인코딩을 쓰는지 보장되지 않는다)
function Post-Form($path, $body) {
    $pairs = @()
    foreach ($k in $body.Keys) {
        $pairs += ([System.Uri]::EscapeDataString([string]$k) + '=' + [System.Uri]::EscapeDataString([string]$body[$k]))
    }
    $raw = [System.Text.Encoding]::UTF8.GetBytes(($pairs -join '&'))
    try {
        return (Invoke-WebRequest -Uri ($Base + $path) -WebSession $sess -UseBasicParsing `
                -Method POST -Body $raw -ContentType 'application/x-www-form-urlencoded; charset=utf-8' -TimeoutSec 30)
    } catch {
        Bad ('POST ' + $path + ' 실패 : ' + $_.Exception.Message)
        return $null
    }
}

# ASP 응답 본문에서 CSRF 토큰을 꺼낸다
function Get-Csrf($html) {
    if ($html -match 'name="_csrf"\s+value="([^"]*)"') { return $Matches[1] }
    return ''
}

# 처리 페이지(*_p.asp)는 성공/실패 모두 HTTP 200 이다.
# 성공은 location.replace, 실패는 alert 로 구분한다.
function Test-NoAlert($html, $label) {
    if ($html -match "alert\('([^']*)'\)") { Bad ($label + '  ->  ' + $Matches[1]); return $false }
    Ok $label
    return $true
}
function Test-Alert($html, $label) {
    if ($html -match "alert\('([^']*)'\)") { Ok ($label + '  ->  ' + $Matches[1]); return $true }
    Bad ($label + '  ->  차단되지 않았다'); return $false
}

# --- multipart/form-data 를 직접 조립한다 -------------------------------------
#     PowerShell 5.1 의 Invoke-WebRequest 에는 -Form 이 없다.
#     한글이 섞이므로 반드시 UTF-8 바이트로 만든다.
function Post-Multipart($path, $fields, $fileField, $fileName, $fileBytes) {
    $boundary = '----jkwE2E' + [guid]::NewGuid().ToString('N')
    $ms  = New-Object System.IO.MemoryStream
    $enc = New-Object System.Text.UTF8Encoding($false)

    function AddText($s) { $b = $enc.GetBytes($s); $ms.Write($b, 0, $b.Length) }

    foreach ($k in $fields.Keys) {
        AddText ("--$boundary`r`n")
        AddText ("Content-Disposition: form-data; name=`"$k`"`r`n`r`n")
        AddText ([string]$fields[$k])
        AddText "`r`n"
    }
    if ($fileField -ne '') {
        AddText ("--$boundary`r`n")
        AddText ("Content-Disposition: form-data; name=`"$fileField`"; filename=`"$fileName`"`r`n")
        AddText ("Content-Type: text/plain`r`n`r`n")
        $ms.Write($fileBytes, 0, $fileBytes.Length)
        AddText "`r`n"
    }
    AddText ("--$boundary--`r`n")

    try {
        return (Invoke-WebRequest -Uri ($Base + $path) -WebSession $sess -UseBasicParsing `
                -Method POST -Body $ms.ToArray() `
                -ContentType ('multipart/form-data; boundary=' + $boundary) -TimeoutSec 60)
    } catch {
        Bad ('POST(multipart) ' + $path + ' 실패 : ' + $_.Exception.Message)
        return $null
    } finally { $ms.Dispose() }
}

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'


# =============================================================================
Head '1. 로그인'
$p = Get-Page ($View + '/member/Login.asp')
if ($p -eq $null) { Write-Host '로그인 화면을 못 받았다. 중단.' -ForegroundColor Red; exit 1 }
$csrf = Get-Csrf $p.Content
if ($csrf -eq '') { Bad 'CSRF 토큰을 찾지 못했다'; exit 1 }
Ok ('CSRF 토큰 수신 : ' + $csrf.Substring(0, [Math]::Min(12, $csrf.Length)) + '...')

$r = Post-Form ($View + '/member/Login_p.asp') @{ _csrf = $csrf; login_id = $Id; user_pw = $Pw; rurl = ($View + '/main/Default.asp') }
if ($r -ne $null) { [void](Test-NoAlert $r.Content ('로그인 : ' + $Id)) }


# =============================================================================
Head '2. 로그인 상태 / 권한 화면'
$p = Get-Page ($View + '/member/Mypage.asp')
if ($p -ne $null) {
    if ($p.Content -match '로그인이 필요' -or $p.Content -match "alert\(") { Bad '마이페이지 : 세션이 유지되지 않았다' }
    else { Ok '마이페이지 진입' }
}
$p = Get-Page ($View + '/admin/Default.asp')
if ($p -ne $null) {
    if ($p.Content -match "alert\(" ) { Bad '관리자 대시보드 : 접근 거부' } else { Ok '관리자 대시보드 진입' }
}


# =============================================================================
Head '3. 게시글 등록 (multipart + 첨부)'
$p = Get-Page ($View + '/board/Write.asp?bid=free')
$csrf = Get-Csrf $p.Content
$title  = 'E2E 테스트 글 ' + $stamp
$content = "한글 본문 확인용입니다.`n라인2"
$fileTxt = 'jkw e2e attachment ' + $stamp + ' 한글포함'
$fileBytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($fileTxt)

$r = Post-Multipart ($View + '/board/Write_p.asp?bid=free') `
        @{ _csrf = $csrf; bid = 'free'; mode = 'write'; id = '0'; page = '1'; sf = ''; sw = '';
           category_id = '0'; title = $title; content = $content; tags = 'e2e'; link_url = '' } `
        'upfile1' ('e2e-' + $stamp + '.txt') $fileBytes

$newId = 0
if ($r -ne $null) {
    if (Test-NoAlert $r.Content '글 등록') {
        if ($r.Content -match "location\.replace\('([^']+)'\)") {
            $goUrl = $Matches[1] -replace '&amp;', '&'
            if ($goUrl -match 'id=(\d+)') { $newId = [int]$Matches[1]; Ok ('새 글 id = ' + $newId) }
            else { Bad ('이동 URL 에 id 가 없다 : ' + $goUrl) }
        } else { Bad '이동 스크립트가 없다' }
    }
}


# =============================================================================
Head '4. 목록 / 상세 반영'
if ($newId -gt 0) {
    $p = Get-Page ($View + '/board/List.asp?bid=free')
    if ($p -ne $null) {
        if ($p.Content -match [regex]::Escape($title)) { Ok '목록에 제목 노출 (한글 정상)' } else { Bad '목록에 새 글이 없다' }
        # cid=0 은 "전체 카테고리" 링크라 정상이다. 상세 키(id/fid)만 본다.
        if ($p.Content -match 'href="[^"]*(?:\?|&amp;|&)(?:id|fid)=0(?![0-9])') { Bad '목록 상세 링크 키가 0 이다' } else { Ok '목록 링크 키 정상' }
    }
    $p = Get-Page ($View + '/board/View.asp?bid=free&id=' + $newId)
    if ($p -ne $null) {
        if ($p.Content -match [regex]::Escape($title)) { Ok '상세 제목 일치' } else { Bad '상세에 제목이 없다' }
        if ($p.Content -match '한글 본문 확인용입니다') { Ok '상세 본문 한글 정상' } else { Bad '상세 본문이 깨졌거나 없다' }
    }
} else { Bad '새 글 id 를 모르므로 4~6 을 건너뛴다' }


# =============================================================================
Head '5. 첨부 다운로드'
$fid = 0
if ($newId -gt 0) {
    $p = Get-Page ($View + '/board/View.asp?bid=free&id=' + $newId)
    if ($p -ne $null -and $p.Content -match 'Download\.asp\?[^"]*fid=(\d+)') { $fid = [int]$Matches[1] }
}
if ($fid -gt 0) {
    try {
        $d = Invoke-WebRequest -Uri ($Base + $View + '/board/Download.asp?bid=free&fid=' + $fid) `
             -WebSession $sess -UseBasicParsing -TimeoutSec 30
        $got = [System.Text.Encoding]::UTF8.GetString($d.Content)
        if ($got -eq $fileTxt) { Ok ('첨부 다운로드 바이트 일치 (' + $d.Content.Length + ' bytes)') }
        else { Bad ('첨부 내용 불일치 : ' + $got) }
    } catch { Bad ('다운로드 실패 : ' + $_.Exception.Message) }
} else { Bad '첨부 fid 를 찾지 못했다' }


# =============================================================================
Head '6. 댓글 등록'
if ($newId -gt 0) {
    $p = Get-Page ($View + '/board/View.asp?bid=free&id=' + $newId)
    $csrf = Get-Csrf $p.Content
    $cmt = 'E2E 댓글 ' + $stamp
    $r = Post-Form ($View + '/board/Comment_p.asp') `
            @{ _csrf = $csrf; bid = 'free'; mode = 'write'; id = $newId; parent_id = '0'; content = $cmt }
    if ($r -ne $null) { [void](Test-NoAlert $r.Content '댓글 등록') }

    $p = Get-Page ($View + '/board/View.asp?bid=free&id=' + $newId)
    if ($p -ne $null -and $p.Content -match [regex]::Escape($cmt)) { Ok '댓글 노출 확인' } else { Bad '댓글이 보이지 않는다' }
}


# =============================================================================
Head '7. 후원 주문 생성 / 모의 승인'
$p = Get-Page ($View + '/donate/Default.asp')
if ($p -ne $null) {
    $csrf = Get-Csrf $p.Content
    $r = Post-Form ($View + '/donate/Order_p.asp') `
            @{ _csrf = $csrf; amount = '5000'; donor_name = 'E2E'; message = '자동검증'; is_public = '1'; pg = 'mock' }
    if ($r -ne $null) {
        $txt = $r.Content.Trim()
        if ($txt -like 'OK|*') {
            $parts = $txt.Split('|')
            Ok ('주문 생성 : ' + $parts[1] + ' / ' + $parts[2] + '원')
            $r2 = Post-Form ($View + '/donate/Approve_p.asp') @{ _csrf = $csrf; order = $parts[1]; result = 'success' }
            if ($r2 -ne $null) {
                if ($r2.Content.Trim() -like 'OK*') { Ok '모의 승인 처리' } else { Bad ('승인 실패 : ' + $r2.Content.Trim()) }
            }
        } else { Bad ('주문 생성 실패 : ' + $txt) }
    }

    # 금액 위조 방어 : 최소금액 미만은 서버가 거부해야 한다
    $r = Post-Form ($View + '/donate/Order_p.asp') `
            @{ _csrf = $csrf; amount = '1'; donor_name = 'E2E'; message = ''; is_public = '0'; pg = 'mock' }
    if ($r -ne $null) {
        if ($r.Content.Trim() -like 'FAIL|*') { Ok ('금액 위조 차단 : ' + $r.Content.Trim()) } else { Bad '1원 주문이 통과했다' }
    }
}


# =============================================================================
Head '8. 문의 등록'
$p = Get-Page ($View + '/contact/Default.asp')
if ($p -ne $null) {
    $csrf = Get-Csrf $p.Content
    $r = Post-Form ($View + '/contact/Send_p.asp') `
            @{ _csrf = $csrf; contact_type = 'etc'; writer_name = 'E2E'; company = 'nable';
               email = 'e2e@example.com'; mobile = '010-0000-0000';
               subject = 'E2E 문의 ' + $stamp; content = '자동 검증 문의입니다.'; is_agree = '1' }
    # Send_p 는 성공도 alert 로 알린다. 문구로 구분한다.
    if ($r -ne $null) {
        if ($r.Content -match '정상적으로 접수') { Ok '문의 등록' }
        elseif ($r.Content -match "alert\('([^']*)'\)") { Bad ('문의 등록  ->  ' + $Matches[1]) }
        else { Bad '문의 등록 : 응답을 해석할 수 없다' }
    }
}


# =============================================================================
Head '9. CSRF 토큰 없는 POST 차단'
$r = Post-Form ($View + '/donate/Order_p.asp') @{ amount = '5000'; donor_name = 'X'; pg = 'mock' }
if ($r -ne $null) {
    if ($r.Content.Trim() -like 'FAIL|*') { Ok ('토큰 없는 주문 차단 : ' + $r.Content.Trim()) } else { Bad '토큰 없이 주문이 생성됐다' }
}
if ($newId -gt 0) {
    $r = Post-Form ($View + '/board/Comment_p.asp') @{ bid = 'free'; mode = 'write'; id = $newId; content = 'no-csrf' }
    if ($r -ne $null) { [void](Test-Alert $r.Content '토큰 없는 댓글') }
}


# =============================================================================
Head '10. 게시글 삭제'
if ($newId -gt 0) {
    $p = Get-Page ($View + '/board/View.asp?bid=free&id=' + $newId)
    $csrf = Get-Csrf $p.Content
    $r = Post-Form ($View + '/board/Delete_p.asp') @{ _csrf = $csrf; bid = 'free'; id = $newId; page = '1' }
    if ($r -ne $null) {
        if ($r.Content -match "alert\('삭제되었습니다") { Ok '글 삭제' } else { Bad ('삭제 실패 : ' + ($r.Content -replace '<[^>]+>',' ').Trim()) }
    }
    $p = Get-Page ($View + '/board/List.asp?bid=free')
    if ($p -ne $null) {
        if ($p.Content -match [regex]::Escape($title)) { Bad '삭제 후에도 목록에 남아 있다' } else { Ok '목록에서 사라짐' }
    }
}


# =============================================================================
Head '11. 로그아웃'
$r = Get-Page ($View + '/member/Logout.asp')
if ($r -ne $null) { Ok '로그아웃 요청' }
$p = Get-Page ($View + '/member/Mypage.asp')
if ($p -ne $null) {
    if ($p.Content -match "alert\(") { Ok '로그아웃 후 마이페이지 차단' } else { Bad '로그아웃 후에도 마이페이지가 열린다' }
}


# =============================================================================
Write-Host ''
if ($fail -eq 0) { Write-Host 'E2E 전체 통과' -ForegroundColor Green; exit 0 }
else { Write-Host ('E2E 문제 ' + $fail + '건') -ForegroundColor Red; exit 1 }
