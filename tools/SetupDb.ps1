# =============================================================================
#  jkw.com  -  데이터베이스 초기 구축 / 검증
# -----------------------------------------------------------------------------
#  SQL Server 가 기동된 뒤 이 스크립트 하나만 실행하면 된다.
#    1) 서비스 상태 확인
#    2) 01_schema.sql / 02_seed.sql / 03_proc.sql 순차 실행 (UTF-8)
#    3) 테이블 / 행 수 / 한글 데이터 정상 여부 검증
#
#  사용 :  powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1
#          powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1 -Server "localhost\SQLEXPRESS"
#          powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1 -Verify   (스크립트 실행 없이 검증만)
# =============================================================================
param(
    # 기본 인스턴스면 'localhost', 명명 인스턴스면 'localhost\SQL2022' 처럼 지정
    [string]$Server  = 'localhost',
    # IIS 응용 프로그램 풀 이름. 지정하면 해당 풀 ID 에 jkwcom 접근 권한을 부여한다.
    [string]$AppPool = '',
    [switch]$Verify
)

# sqlcmd 는 오류를 stderr 로 내보내는데, PowerShell 5.1 에서 ErrorActionPreference=Stop
# 이면 그것만으로 스크립트가 중단되어 아래 친절한 안내가 출력되지 않는다.
# 네이티브 명령의 성공/실패는 $LASTEXITCODE 로만 판정한다.
$ErrorActionPreference = 'Continue'
$root  = Split-Path -Parent $PSScriptRoot
$dbDir = Join-Path $root 'db\mssql'
$fail  = 0

function Head($t) { Write-Host ''; Write-Host ('=== ' + $t + ' ===') -ForegroundColor Cyan }
function Ok($t)   { Write-Host ('  [OK] ' + $t) -ForegroundColor Green }
function Bad($t)  { Write-Host ('  [!!] ' + $t) -ForegroundColor Red; $script:fail++ }

# --- sqlcmd 찾기 --------------------------------------------------------------
$sqlcmd = (Get-Command sqlcmd -ErrorAction SilentlyContinue).Source
if (-not $sqlcmd) {
    foreach ($p in @(
        "$env:ProgramFiles\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\SQLCMD.EXE",
        "$env:ProgramFiles\Microsoft SQL Server\Client SDK\ODBC\180\Tools\Binn\SQLCMD.EXE",
        "${env:ProgramFiles(x86)}\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\SQLCMD.EXE")) {
        if (Test-Path $p) { $sqlcmd = $p; break }
    }
}
if (-not $sqlcmd) { Bad 'sqlcmd 를 찾을 수 없습니다.'; exit 1 }

# --- 1. 서비스 / 접속 확인 -----------------------------------------------------
Head '1. SQL Server 접속 확인'
try {
    $svc = Get-Service MSSQLSERVER -ErrorAction SilentlyContinue
    if ($svc) { Write-Host ('  서비스 상태 : ' + $svc.Status) }
} catch {}

$ver = & $sqlcmd -S $Server -E -h -1 -W -Q "SET NOCOUNT ON; SELECT CONVERT(varchar(200), SERVERPROPERTY('ProductVersion')) + ' / ' + CONVERT(varchar(100), SERVERPROPERTY('Edition'))" 2>&1
if ($LASTEXITCODE -ne 0) {
    Bad ('접속 실패 : ' + ($ver -join ' '))
    Write-Host '  → SQL Server 서비스가 기동되지 않았습니다. 먼저 서비스를 시작하세요.' -ForegroundColor Yellow
    exit 1
}
Ok ('접속 성공 : ' + (($ver | Where-Object { $_.Trim() -ne '' }) -join ' '))

# --- 2. 스크립트 실행 ----------------------------------------------------------
if (-not $Verify) {
    Head '2. 스키마 / 시드 / 프로시저 실행'
    foreach ($f in @('01_schema.sql', '02_seed.sql', '03_proc.sql')) {
        $path = Join-Path $dbDir $f
        if (-not (Test-Path $path)) { Bad ($f + ' 없음'); continue }

        Write-Host ('  ▶ ' + $f + ' 실행 중...') -ForegroundColor Gray
        # -f 65001 : 입력 파일을 UTF-8 로 읽는다 (한글 보존)
        # -b       : 오류 발생 시 즉시 중단하고 종료코드 반환
        $o = & $sqlcmd -S $Server -E -f 65001 -b -i $path 2>&1
        if ($LASTEXITCODE -ne 0) {
            Bad ($f + ' 실행 실패')
            $o | Select-Object -Last 15 | ForEach-Object { Write-Host ('      ' + $_) -ForegroundColor Red }
        } else {
            Ok ($f + ' 완료')
        }
    }
}

# --- 3. 검증 -------------------------------------------------------------------
Head '3. 검증'

$q = @"
SET NOCOUNT ON;
SELECT 'tables=' + CONVERT(varchar(10), COUNT(*)) FROM sys.tables;
SELECT 'config=' + CONVERT(varchar(10), COUNT(*)) FROM tb_config;
SELECT 'member=' + CONVERT(varchar(10), COUNT(*)) FROM tb_member;
SELECT 'menu=' + CONVERT(varchar(10), COUNT(*)) FROM tb_menu;
SELECT 'board=' + CONVERT(varchar(10), COUNT(*)) FROM tb_board_master;
SELECT 'article=' + CONVERT(varchar(10), COUNT(*)) FROM tb_board_article;
SELECT 'skill=' + CONVERT(varchar(10), COUNT(*)) FROM tb_skill;
SELECT 'portfolio=' + CONVERT(varchar(10), COUNT(*)) FROM tb_portfolio;
"@
$r = & $sqlcmd -S $Server -d jkwcom -E -h -1 -W -Q $q 2>&1
if ($LASTEXITCODE -ne 0) { Bad ('검증 쿼리 실패 : ' + ($r -join ' ')) }
else { ($r | Where-Object { $_.Trim() -ne '' }) | ForEach-Object { Write-Host ('  ' + $_) } }

Head '4. 한글 데이터 확인 (인코딩 검증)'
$r2 = & $sqlcmd -S $Server -d jkwcom -E -h -1 -W -Q "SET NOCOUNT ON; SELECT board_id + ' : ' + board_name FROM tb_board_master ORDER BY sort_no" 2>&1
if ($LASTEXITCODE -ne 0) { Bad '조회 실패' }
else {
    $txt = ($r2 | Where-Object { $_.Trim() -ne '' }) -join "`n"
    $txt -split "`n" | ForEach-Object { Write-Host ('  ' + $_) }
    if ($txt -match '\?\?') { Bad '한글이 ?? 로 깨졌습니다. sqlcmd 에 -f 65001 을 주고 다시 실행하세요.' }
    elseif ($txt -match '[가-힣]') { Ok '한글 정상' }
}

Head '5. 비밀번호 해시 규약 확인'
$r3 = & $sqlcmd -S $Server -d jkwcom -E -h -1 -W -Q @"
SET NOCOUNT ON;
SELECT login_id + ' : ' +
       CASE WHEN user_pw = LOWER(CONVERT(CHAR(64), HASHBYTES('SHA2_256', user_salt + CASE login_id WHEN 'admin' THEN 'admin1234!' ELSE 'test1234!' END), 2))
            THEN 'hash OK' ELSE 'hash MISMATCH' END
  FROM tb_member ORDER BY member_id
"@ 2>&1
if ($LASTEXITCODE -eq 0) {
    ($r3 | Where-Object { $_.Trim() -ne '' }) | ForEach-Object { Write-Host ('  ' + $_) }
    if (($r3 -join ' ') -match 'MISMATCH') { Bad '해시 불일치' } else { Ok '초기 계정 해시 정상' }
} else { Bad '해시 검증 실패' }

# --- 6. IIS 응용 프로그램 풀 계정에 권한 부여 -----------------------------------
if ($AppPool -ne '') {
    Head ('6. IIS 앱풀 계정 권한 부여  (IIS APPPOOL\' + $AppPool + ')')
    $acct = 'IIS APPPOOL\' + $AppPool
    $g = @"
SET NOCOUNT ON;
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'$acct')
    CREATE LOGIN [$acct] FROM WINDOWS;
USE jkwcom;
IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'$acct')
    CREATE USER [$acct] FOR LOGIN [$acct];
ALTER ROLE db_datareader ADD MEMBER [$acct];
ALTER ROLE db_datawriter ADD MEMBER [$acct];
GRANT EXECUTE ON SCHEMA::dbo TO [$acct];
SELECT 'granted';
"@
    $r6 = & $sqlcmd -S $Server -E -b -h -1 -W -Q $g
    if ($LASTEXITCODE -ne 0) {
        Bad '권한 부여 실패'
        $r6 | Select-Object -Last 8 | ForEach-Object { Write-Host ('      ' + $_) -ForegroundColor Red }
    } else {
        Ok ($acct + ' : db_datareader + db_datawriter + EXECUTE 부여 완료')
    }
}

Write-Host ''
if ($fail -eq 0) {
    Write-Host '전체 통과 - 사이트에서 확인하세요.' -ForegroundColor Green
    Write-Host '  http://localhost:8080/            (메인)'
    Write-Host '  http://localhost:8080/view/admin/Default.asp   admin / admin1234!'
    exit 0
} else {
    Write-Host ('문제 ' + $fail + '건') -ForegroundColor Red
    exit 1
}
