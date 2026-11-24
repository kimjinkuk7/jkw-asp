# jkw.com — Classic ASP 스택

개인 홈페이지를 **Classic ASP (VBScript) + MS SQL Server** 로 구현한 것입니다.
같은 서비스를 Spring Boot · Django · React 로도 만들었고, 각각 **별도 리포지토리**입니다.

| 스택 | 리포지토리 | DB |
|---|---|---|
| **Classic ASP** | 이 리포지토리 | MS SQL Server |
| Spring Boot | `jkw-springboot` | MS SQL Server |
| Django | `jkw-django` | PostgreSQL |
| React SPA (+ 동봉 ASP JSON API) | `jkw-react` | MS SQL Server |

## 자립형 리포지토리

이 폴더 하나로 사이트가 완성됩니다. **다른 스택의 폴더를 참조하지 않습니다.**

```
asp/                    ← 이 리포지토리 (IIS 사이트 루트로 지정 가능)
├─ default.asp          진입점 (302 → view/main/Default.asp)
├─ web.config           IIS 설정 전부 (ASP 런타임 · 보안 헤더 · 소스 차단 · 캐시)
├─ App.Conf.asp         ★ 유일한 환경 설정 지점
├─ conf/ lib/ model/ code/ view/      애플리케이션
├─ api/                 JSON API 16개 (React 스택이 쓰는 것과 같은 엔드포인트)
├─ assets/              CSS · JS · 이미지     ← 이 스택의 사본
├─ upload/              업로드 파일           ← 이 스택의 폴더
├─ db/mssql/            스키마 · 시드 · 프로시저
├─ docs/                설계 문서
└─ tools/               SetupDb · CheckAsp · E2E
```

### 배치 경로를 바꾸려면 (3줄)

소스의 `#include` 는 **전부 상대경로**(`#include file="../../code/..."`)라 폴더를 어디에 두든
그대로 동작합니다. 바꿀 것은 브라우저가 받는 URL 접두어뿐입니다.

`App.Conf.asp` 의 세 상수를 함께 바꿉니다.

| 배치 | `APP_ROOT` | `VIEW_ROOT` | `UPLOAD_ROOT` |
|---|---|---|---|
| 이 폴더가 IIS 사이트 루트 | `""` | `"/view"` | `"/upload"` |
| 사이트 루트 아래 `asp` 폴더 | `"/asp"` | `"/asp/view"` | `"/asp/upload"` |

> VBScript 는 `Const` 에서 다른 `Const` 를 참조할 수 없습니다(실측: HTTP 500).
> 그래서 한 줄로 묶지 못하고 셋을 각각 리터럴로 적습니다.

---

## 0. 개요

| 항목 | 내용 |
|---|---|
| 런타임 | Classic ASP (VBScript) / IIS 10 |
| DB | MS SQL Server 2019 (15.0) — DB명 `jkwcom` |
| 인코딩 | UTF-8 (CodePage 65001) |
| 구조 | `code / conf / lib / model / view` (ZOCBOv7 관례 준용) |
| 반응형 | Mobile First · 600px / 1024px 2개 브레이크포인트 · 다크모드 |
| 외부 컴포넌트 | **없음.** JSON 파서 · multipart 파서 · SHA-256 모두 순수 VBScript |

---

## 1. 설치

### 1-1. 데이터베이스

```powershell
# 관리자 권한 PowerShell
Start-Service MSSQLSERVER

# -f 65001 : 입력 파일을 UTF-8 로 읽는다 (한글 깨짐 방지)
sqlcmd -S localhost -E -f 65001 -i "db\mssql\01_schema.sql"
sqlcmd -S localhost -E -f 65001 -i "db\mssql\02_seed.sql"
sqlcmd -S localhost -E -f 65001 -i "db\mssql\03_proc.sql"
```

정상 실행되면 테이블 16개와 초기 데이터가 생성됩니다.

> SSMS 로 실행해도 됩니다. `.sql` 파일에는 **UTF-8 BOM** 이 붙어 있어 SSMS 가 인코딩을 자동 인식합니다.

**한 번에 처리하려면** 아래 스크립트를 쓰세요. 3개 스크립트 실행 + 테이블/행 수 · 한글 인코딩 ·
비밀번호 해시 규약까지 검증합니다.

```powershell
powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1

# 이미 구축했고 검증만 하려면
powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1 -Verify

# 명명 인스턴스
powershell -ExecutionPolicy Bypass -File tools\SetupDb.ps1 -Server "localhost\SQLEXPRESS"
```

| 구분 | 아이디 | 비밀번호 |
|---|---|---|
| 최고관리자 | `admin` | `admin1234!` |
| 일반회원 | `test` | `test1234!` |

> ⚠️ 최초 로그인 후 **마이페이지 → 비밀번호 변경**에서 반드시 변경하세요.

### 1-2. IIS

**관리자 권한 PowerShell** 에서 아래 3가지를 먼저 처리해야 합니다. 하나라도 빠지면 사이트가 뜨지 않습니다.

```powershell
# (1) Classic ASP 기능 설치  ← 기본 미설치. 없으면 ASP 가 아예 실행되지 않는다
Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All -NoRestart

# (2) asp 구성 섹션 잠금 해제
#     IIS 는 system.webServer/asp 를 기본 잠금(Deny)해서, web.config 에 <asp> 가 있으면
#     HTTP 500.19 가 난다. 이 프로젝트는 web.config 로 ASP 설정을 관리하므로 해제한다.
& "$env:windir\system32\inetsrv\appcmd.exe" unlock config -section:system.webServer/asp

# (3) 업로드 폴더 쓰기 권한  (<앱풀이름> 을 실제 이름으로 교체)
icacls "<이 리포지토리 경로>\upload" /grant "IIS AppPool\<앱풀이름>:(OI)(CI)M" /T

# (4) 익명 인증이 앱 풀 ID 로 동작하게 한다  ← 빠뜨리면 DB 접속이 IUSR 로 시도돼 실패한다
#     userName 을 비우면 IIS 가 IUSR 대신 앱 풀 ID(IIS APPPOOL\<앱풀이름>)를 쓴다.
& "$env:windir\system32\inetsrv\appcmd.exe" set config "<사이트명>" `
    -section:system.webServer/security/authentication/anonymousAuthentication `
    /userName:"" /commit:apphost

iisreset
```

사이트는 IIS 관리자에서 **실제 경로 = 이 리포지토리 폴더**, 응용 프로그램 풀 = 통합 모드로 추가합니다.
그 경우 `App.Conf.asp` 의 `APP_ROOT` / `VIEW_ROOT` / `UPLOAD_ROOT` 를
`""` / `"/view"` / `"/upload"` 로 바꾸세요 (맨 위 "배치 경로" 표 참고).

> **★ `<asp>` 섹션은 IIS 애플리케이션 루트의 `web.config` 에서만 적용됩니다.**
> 하위 디렉터리에 두면 IIS 가 **오류 없이 조용히 무시**합니다. (2026-11-24 실측)
> 이 리포지토리를 사이트 루트가 아닌 하위 폴더(`<사이트루트>\asp`)에 두면
> `web.config` 의 `enableParentPaths` 가 먹지 않아 **모든 페이지가 500** 이 됩니다.
> 그 배치에서는 사이트 루트 `web.config` 에 같은 `<asp>` 블록을 넣어야 합니다.
> 확인 방법 : `Server.MapPath("..")` 가 `ASP 0175` 를 내면 미적용 상태입니다.

`localhost` 로 접속하려면 호스트 헤더 없는 바인딩이 하나 필요합니다.

```powershell
# 예) 8080 포트 바인딩 추가 (Default Web Site 가 80 을 쓰고 있을 때)
& "$env:windir\system32\inetsrv\appcmd.exe" set site "jkw" /+"bindings.[protocol='http',bindingInformation='*:8080:']"
```

### 1-3. 권한

```powershell
# 응용 프로그램 풀 ID 에 업로드 폴더 쓰기 권한 부여
icacls "<이 리포지토리 경로>\upload" /grant "IIS AppPool\DefaultAppPool:(OI)(CI)M" /T
```

DB 접근은 기본이 **Windows 통합 인증(SSPI)** 입니다. 응용 프로그램 풀 ID(`IIS APPPOOL\<풀이름>`)를
SQL Server 로그인으로 추가하고 `jkwcom` DB 에 `db_datareader` + `db_datawriter` + `EXECUTE` 를 부여하세요.

<details>
<summary>SQL 인증을 쓰려면</summary>

```sql
CREATE LOGIN jkwcom_app WITH PASSWORD = 'ChangeMe!2026';
USE jkwcom;
CREATE USER jkwcom_app FOR LOGIN jkwcom_app;
ALTER ROLE db_datareader ADD MEMBER jkwcom_app;
ALTER ROLE db_datawriter ADD MEMBER jkwcom_app;
GRANT EXECUTE ON SCHEMA::dbo TO jkwcom_app;
```

그리고 `App.Conf.asp` 에서:
```asp
Const DB_TRUSTED = False
Const DB_USER    = "jkwcom_app"
Const DB_PASS    = "ChangeMe!2026"
```
</details>

### 1-4. 접속

이 폴더가 IIS 사이트 루트일 때 (`APP_ROOT = ""`)

```
http://localhost:8080/                           →  /view/main/Default.asp 로 302
http://localhost:8080/view/admin/Default.asp        관리자
```

사이트 루트 아래 `asp` 폴더일 때 (`APP_ROOT = "/asp"`, 현재 개발 PC)

```
http://localhost:8080/asp/                       →  /asp/view/main/Default.asp 로 302
http://localhost:8080/asp/view/admin/Default.asp    관리자
```

### 1-5. 점검 도구

코드를 수정한 뒤에는 아래 스크립트를 돌리세요. Classic ASP 는 컴파일 단계가 없어
오타가 런타임에야 드러나므로, 자주 나는 오류 7종을 미리 잡아줍니다.

```powershell
# 정적 점검만
powershell -ExecutionPolicy Bypass -File tools\CheckAsp.ps1

# 정적 점검 + 전 페이지 실제 요청 (컴파일 오류까지 확인)
powershell -ExecutionPolicy Bypass -File tools\CheckAsp.ps1 -Http http://localhost:8080
```

검사 항목 : include 경로 · 최상위 `Dim`/`Const` 중복 · 전역 함수명 충돌 ·
매개변수와 함수명 동일 · VBScript 내장 함수명 충돌 · 주석 안 include 지시자 · UTF-8/CODEPAGE ·
사이트 루트(`/`) 리다이렉트 · 전 페이지 HTTP 응답

`CheckAsp.ps1` 은 "페이지가 오류 없이 뜨는가"까지만 봅니다.
**기능이 실제로 동작하는가**는 `E2E.ps1` 이 봅니다.

```powershell
powershell -ExecutionPolicy Bypass -File tools\E2E.ps1
```

검사 항목 : 로그인(CSRF) · 세션 유지 · 관리자 접근 · 글 등록(multipart + 첨부) ·
목록/상세 반영 · 첨부 다운로드 바이트 일치 · 댓글 · 후원 주문/승인 · 금액 위조 차단 ·
문의 등록 · CSRF 없는 POST 차단 · 글 삭제 · 로그아웃

> `E2E.ps1` 은 `admin` 계정으로 **실제 데이터를 쓰고 지웁니다.** 운영 서버에 돌리지 마세요.

---

## 2. 폴더 구조

```
asp\                        ← 이 리포지토리
│
├─ default.asp              진입점 (302 → view/main/Default.asp)
├─ web.config               IIS 설정 전부
│                             ASP 런타임(asp 섹션) · 기본 문서 · 보안 헤더
│                             소스 디렉터리 차단(hiddenSegments) · upload/assets 캐시
├─ App.Conf.asp             ★ 유일한 환경 설정 지점 (배치 경로 / DB / 상수 / 전역)
├─ README.md
│
├─ conf\                    설정
│   └─ Common.Conf.asp      요청 공통 초기화 (설정 로드 · 세션 검증 · 방문로그)
│
├─ lib\                     범용 클래스 (도메인 무관)
│   ├─ DBHelper.Class.asp   ★ ADODB 를 다루는 유일한 파일. 파라미터 바인딩 강제
│   ├─ StringUtil.Class.asp 이스케이프 · 검증 · 포맷
│   ├─ DateUtil.Class.asp
│   ├─ RequestUtil.Class.asp
│   ├─ Paging.Class.asp
│   ├─ Crypt.Class.asp      SHA-256 / Base64 / 쿠키 인코딩
│   ├─ Json.Class.asp       ★ 순수 VBScript JSON 파서  (외부 컴포넌트 없음)
│   ├─ JsonOut.Class.asp    ★ JSON 생성기 (api\ 가 쓴다)
│   ├─ Http.Class.asp       ★ ServerXMLHTTP 래퍼 (외부 API 호출)
│   ├─ Cookie.Class.asp
│   ├─ FileUpload.Class.asp multipart 파서 (외부 컴포넌트 없음)
│   └─ FileDownload.Class.asp
│
├─ model\                   도메인 클래스 (DB 접근)
│   ├─ common\   Config · Menu · Auth · Visit
│   ├─ board\    Board · Article · Comment · BoardFile
│   ├─ member\   Member · PhoneAuth · Social (소셜 연동)
│   ├─ donate\   Donate · PgInicis (개발자 지원금)
│   ├─ portfolio\Portfolio
│   ├─ about\    About (경력 + 스킬)
│   └─ contact\  Contact
│
├─ code\                    컨트롤러
│   ├─ Common.Code.asp      ★ 모든 Code 파일이 최초로 include
│   ├─ api\                 API 공통 (응답 봉투 · 본문 파싱 · CSRF · 권한)
│   └─ main\ about\ portfolio\ board\ member\ donate\ contact\ admin\
│
├─ view\                    화면 (URL 진입점)
│   ├─ include\  Header · Gnb · Lnb · Footer · AdminSide
│   └─ main\ about\ portfolio\ board\ member\ donate\ contact\ admin\
│
├─ api\                     JSON API 16개 (화면 대신 JSON 을 돌려주는 진입점)
│
├─ assets\   css\style.css  js\common.js  img\
├─ upload\   board\  portfolio\  profile\
├─ db\mssql\ 01_schema.sql  02_seed.sql  03_proc.sql
├─ docs\     ARCHITECTURE.md  DATABASE.md
└─ tools\    SetupDb.ps1  CheckAsp.ps1  E2E.ps1
```

**include 규칙 — 전부 상대경로**
`#include virtual="/asp/..."` 는 **사이트 루트 기준**이라 폴더를 옮기면 전부 깨진다.
그래서 이 리포지토리는 `#include file="../../code/..."` 처럼 **그 지시자가 적힌 파일의
디렉터리 기준** 상대경로만 쓴다(302건). 폴더를 어디에 두든 그대로 동작한다.

> 전제 : `web.config` 의 `<asp enableParentPaths="true">`.
> 이 값이 꺼져 있으면 `../` 를 쓴 include 가 전부 500 이 된다.
> 그리고 **이 섹션은 IIS 애플리케이션 루트에서만 적용된다** (1-2장 참고).

**URL 규칙 — 접두어 3개로 수렴**
- 화면 링크는 전부 `U("/board/List.asp")` → `VIEW_ROOT` 가 앞에 붙는다
- CSS/JS 는 `<%= APP_ROOT %>/assets/...`
- 업로드는 `UPLOAD_ROOT`, 그리고 **DB 에 저장된 경로는 `UpUrl()` 로 다시 붙인다**

`UpUrl()` (`code\Common.Code.asp`) 이 필요한 이유 : `tb_board_file.save_path` 와
`tb_portfolio.thumb_file` 에는 **그 글을 쓴 스택의 접두어**가 박힌다.
같은 `jkwcom` DB 를 Spring Boot · React 스택과 공유하므로, 읽을 때 현재 스택 기준으로
바꿔줘야 첨부가 깨지지 않는다. 멱등이라 몇 번 거쳐도 안전하다.

```
"/upload/board/202611"        ->  "/asp/upload/board/202611"
"/react/upload/board/202611"  ->  "/asp/upload/board/202611"
"/asp/upload/board/202611"    ->  그대로
```

---

## 3. 페이지 작성 규약

### view (URL 진입점)

`view/board/List.asp` (리포지토리 기준 깊이 2) 의 예 :

```asp
<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/board/List.Code.asp"-->
<!--#include file="../include/Header.asp"-->

    ... HTML ...

<!--#include file="../include/Footer.asp"-->
```

> **`#include file` 은 그 지시자가 적힌 파일의 디렉터리 기준**이다 (IIS 실측 확인).
> 중첩돼도 마찬가지다 — `view/board/List.asp` 가 `../include/Header.asp` 를 읽고,
> 그 `Header.asp` 안의 `Gnb.asp` 는 `view/include/` 에서 풀린다.
> 그래서 깊이만 맞추면 되고, 리포지토리를 어느 URL 아래에 두든 고칠 것이 없다.
>
> `../` 를 쓰므로 `web.config` 의 `<asp enableParentPaths="true">` 가 **필수**다.

> `@LANGUAGE` 지시자와 `Option Explicit` 은 **실제 요청되는 파일(view)** 에 둡니다.
> ZOCBOv7 은 Code 파일에 두지만, include 파일의 `@` 지시자는 IIS 설정에 따라 무시될 수 있어
> 이 프로젝트는 view 에 두는 방식으로 통일했습니다.

### code (컨트롤러)

`code/board/List.Code.asp` (깊이 2) 의 예 :

```asp
<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/board/Board.Class.asp"-->
<%
    g_menuCode  = "comm_free"       ' GNB 하이라이트
    g_pageTitle = "자유게시판"       ' <title>

    ' 1. 권한 확인 → 2. 파라미터 수집/검증 → 3. 모델 호출 → 4. 변수 준비
%>
```

### model (도메인 클래스)

```asp
Class CArticle
    Public DBHelper
    Private Sub Class_Initialize()
        Set DBHelper = New CDBHelper
        DBHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    End Sub

    Public Function Article_l(ByVal pBoardId, ByVal pOffset, ByVal pRows)
        With DBHelper
            .Sql = "SELECT ... WHERE board_id = ? OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
            .SetParamVarchar(pBoardId)
            .SetParamInt(pOffset)
            .SetParamInt(pRows)
        End With
        Article_l = DBHelper.ExecuteGetRows()    ' arr(컬럼, 행) 또는 Empty
    End Function
End Class

Dim Article
Set Article = New CArticle
```

**메서드 접미사** — `_l` 목록 · `_v` 상세 · `_i` 등록 · `_u` 수정 · `_d` 삭제 · `_c` 건수 · `_p` 처리

**결과 접근 헬퍼** (`code/Common.Code.asp`)

| 함수 | 용도 |
|---|---|
| `RowCnt(arr)` | 행 개수 (결과 없으면 0) |
| `AV(arr, c, r)` | 값 (NULL → `""`) |
| `AVH(arr, c, r)` | 값 + HTML 이스케이프 ← **뷰에서 기본** |
| `AVI(arr, c, r)` | 정수 |
| `AVB(arr, c, r)` | BIT → Boolean |
| `IIfStr(조건, a, b)` | 삼항 연산자 대용 |

---

## 3-1. 인코딩 정책 (전 구간 UTF-8)

| 구간 | 설정 | 위치 |
|---|---|---|
| 소스 파일 | **UTF-8 (BOM 없음)** — `.asp` `.css` `.js` `.md` | 전 파일 |
| SQL 스크립트 | **UTF-8 (BOM 있음)** — sqlcmd/SSMS 자동 인식용 | `db/mssql/*.sql` |
| ASP 소스 해석 | `<%@ … CODEPAGE="65001" %>` | 모든 view 진입 파일 1행 |
| 응답 본문 | `Response.CodePage = 65001` / `Response.CharSet = "utf-8"` | `App.Conf.asp` |
| 세션/폼 입력 | `Session.CodePage = 65001` | `App.Conf.asp` |
| HTML | `<meta charset="utf-8">` | `view/include/Header.asp` |
| CSS | `@charset "utf-8";` | `assets/css/style.css` |
| DB | 전 컬럼 `NVARCHAR` (유니코드) | `db/mssql/01_schema.sql` |

**왜 `.asp` 에는 BOM 을 넣지 않는가**
include 파일에 BOM 이 있으면 병합 시 본문 중간에 U+FEFF 가 섞여 들어갑니다.
모든 진입 파일 1행에 `CODEPAGE="65001"` 이 있으므로 IIS 가 include 까지 UTF-8 로 읽습니다.
(만약 편집기에서 저장하다 실수로 CP949 로 바뀌어 한글이 깨지면, 그때는 해당 파일을 UTF-8 로 다시 저장하세요.)

**주의** — 새 `.asp` 파일을 만들 때 반드시 UTF-8 로 저장하고 1행에 `CODEPAGE="65001"` 을 넣으세요.
Visual Studio / VS Code 는 기본이 UTF-8 이지만, 메모장·EditPlus 는 ANSI 로 저장될 수 있습니다.

---

## 4. 반드시 지켜야 할 4가지

1. **SQL 은 항상 `?` 파라미터 바인딩** — 문자열 연결로 값을 넣지 않습니다. (`lib/DBHelper.Class.asp` 외에서 ADODB 직접 사용 금지)
2. **출력은 항상 `HE()` / `AVH()` 통과** — 에디터 본문만 `StringUtil.SafeHtml()`
3. **처리 페이지(`*_p.asp`) 첫 줄에 `Auth.CheckToken`** — 폼에는 `<%= Auth.TokenField() %>`
4. **비밀번호는 ASCII 만** — 4개 스택 간 SHA-256 해시 호환을 위한 규약 (`docs/DATABASE.md` 5장)

## 4-1. 휴대폰 인증 (`PHONE_AUTH_MODE`)

`App.Conf.asp` 의 상수 하나로 방식을 갈아끼운다.

| 값 | 동작 | 상태 |
|---|---|---|
| `"off"` | 인증 단계 없음. 연락처는 그냥 입력값 | — |
| **`"mock"`** | **시연용.** 버튼을 누르면 인증 처리. 문자를 보내지 않으므로 **번호 소유를 확인하지 않는다** | 기본값 |
| `"firebase"` | Firebase Phone Auth (무료 할당량 · 발신번호 등록 불필요) | 자리만 예약 |
| `"sms"` | 국내 SMS API + 인증번호 6자리 | 자리만 예약 |

```
asp/view/member/Join.asp              [인증하기] 버튼 + 상태 표시 (mock 이면 "시연용" 문구 노출)
assets/js/common.js                   initPhoneAuth() - 화면 표시만 담당  ← 스택 공용
asp/view/member/PhoneAuth_p.asp       AJAX 엔드포인트 (순수 텍스트 : OK|번호  또는  FAIL|메시지)
asp/code/member/PhoneAuth_p.Code.asp  ★ 방식 교체 시 손댈 곳은 이 파일의 Select Case 뿐
asp/code/member/Join_p.Code.asp       가입 시 서버측 재검증
```

**흐름 (mock 도 실제 SMS 인증과 같은 구조)**

```
[인증번호 받기] → 서버가 6자리 생성 → tb_phone_auth 저장(만료 3분)
                → 입력칸 등장 + 남은시간 카운트다운
                → 시연용이면 인증번호를 입력칸 아래 회색으로 표시
[확인]          → 서버가 DB 값과 대조 (틀리면 try_count +1, 5회 초과 시 차단)
                → 성공하면 verified_dt 기록 + 세션에 인증 번호 저장
```

`tb_phone_auth` 에 만료시각·시도횟수를 두는 이유는 무한 대입과 무한 재발송을 막기 위한 것이다.
10분당 5회 발송 제한도 같은 이유다. 실제 연동으로 바뀌어도 이 로직은 그대로 쓴다.

**설계 규칙 4가지**

1. **인증 상태는 세션에만 둔다.** `Session("PHONEOK_NUM")`. 폼 hidden 으로 넘기면 조작된다.
   화면 JS 는 표시 전용이고, `Join_p` 가 세션과 입력 번호를 다시 비교한다.
2. **인증 방식을 DB 에 남긴다.** `tb_member.mobile_verify_type` = `mock` / `firebase` / `sms`.
   시연용으로 들어온 값을 실제 인증과 구분할 수 없으면 나중에 데이터를 믿을 수 없다.
3. **번호가 바뀌면 인증을 해제한다.** `Member_u` 가 `CASE WHEN ISNULL(mobile,'') = ISNULL(?,'')`
   로 처리한다. 안 하면 인증 후 마이페이지에서 번호만 바꿔치기할 수 있다.
4. **인증번호는 mock 일 때만 응답에 담는다.** 실제 모드에서 응답에 넣으면
   문자를 받지 않고도 인증을 통과할 수 있다 (`PhoneAuth_p.Code.asp` 의 분기).

> ⚠️ `mock` 은 **번호 소유 확인이 아니다.** 실명·생년월일이 필요한 법적 본인확인은
> 본인확인기관(NICE·KCB 등)과 계약해야 하며 무료 대안이 없다.

## 4-2. 소셜 로그인 (`SOCIAL_MODE`)

| 값 | 동작 |
|---|---|
| **`"mock"`** | **시연용.** 제공자 서버에 가지 않고 고정 데모 계정으로 로그인한다 (기본값) |
| `"live"` | 실제 OAuth 2.0. 아래 키를 채워야 동작한다 |

```
asp/lib/Json.Class.asp                순수 VBScript JSON 파서
asp/lib/Http.Class.asp                ServerXMLHTTP 래퍼 (TLS 1.3)
asp/model/member/Social.Class.asp     제공자별 URL·토큰교환·프로필·연동 관리
asp/view/member/SocialStart.asp       state 생성 → 제공자(또는 모의) 화면으로
asp/view/member/SocialMock.asp        ★ 시연용 계정선택·동의 화면
asp/view/member/SocialCallback.asp    state 대조 → 토큰 → 프로필 → 로그인/가입
```

로그인 버튼은 **별도 창(팝업)** 으로 연다. 실제 OAuth 와 같은 방식이고, 팝업이 콜백까지
처리한 뒤 부모 창을 옮기고 스스로 닫는다. 팝업이 차단되면 같은 창에서 이동한다.

> **모의 화면에 비밀번호 입력칸을 두지 않은 이유**
> 실제 제공자 로그인 화면을 그대로 모사하면서 비밀번호를 받는 UI 는 피싱 페이지와
> 구조가 같아진다. 게다가 OAuth 에서 우리 서버는 원래 비밀번호를 볼 수 없는 것이 정상이다.
> 그래서 계정 선택 + 권한 동의만 두었다. 흐름을 보여주는 목적에는 이게 더 정확하다.

**키 발급** (무료, 각 콘솔에서 직접) — 발급 후 `App.Conf.asp` 에 넣고 `SOCIAL_MODE = "live"`.

| 제공자 | 콘솔 |
|---|---|
| 구글 | `console.cloud.google.com` → OAuth 클라이언트 ID |
| 네이버 | `developers.naver.com/apps` |
| 카카오 | `developers.kakao.com` (카카오 로그인 활성화 + 동의항목 설정) |

★ **리다이렉트 URI 를 콘솔에 그대로 등록해야 한다.** 한 글자만 달라도 실패한다.

```
http://localhost:8080/asp/view/member/SocialCallback.asp?provider=google
```

운영 전환 시 https 주소로 다시 등록한다 (구글은 `localhost` 만 http 를 허용).

**계정 연결 규칙**

1. `(provider, provider_uid)` 가 유일 키다. **이메일은 바뀔 수 있어 키로 쓰지 않는다.**
2. 연동이 없고 이메일이 같은 기존 회원이 있으면 그 계정에 연동을 붙인다.
3. 둘 다 없으면 회원을 새로 만든다. 이때 `user_pw` 는 로그인할 수 없는 임의값을 넣는다
   (스키마를 4스택이 공유하므로 `user_pw` 를 NULL 허용으로 바꾸지 않았다).
4. `state` 를 세션에 넣고 콜백에서 반드시 대조한다 (OAuth 의 CSRF 방어). 1회용이다.

## 4-3. 개발자 지원금 / 결제 (`PAY_PROVIDER`)

| 값 | 동작 |
|---|---|
| **`"mock"`** | **시연용.** 결제창 대신 성공/실패를 고르는 화면. 카드번호 입력 없음 (기본값) |
| `"toss"` | 토스페이먼츠. 테스트 키를 넣으면 실제 결제창까지 뜨고 승인 API 도 호출된다 |
| `"inicis"` | KG이니시스 웹표준. 테스트 상점 정보를 넣으면 iframe 결제창이 뜬다 |

> ### 🔑 PG 키는 비어 있습니다
>
> 이 저장소는 공개용이라 `App.Conf.asp` 의 PG 키를 **빈 문자열로 커밋**합니다.
>
> ```asp
> Const TOSS_CLIENT_KEY = ""     ' test_ck_...
> Const TOSS_SECRET_KEY = ""     ' test_sk_...  ← 서버 전용. 절대 커밋 금지
> Const INICIS_MID      = ""     ' 테스트 상점아이디
> Const INICIS_SIGNKEY  = ""     ' 위 MID 와 짝이 되는 signKey
> ```
>
> 키가 비면 그 모듈은 결제 화면에서 **"키가 설정되지 않아 사용할 수 없습니다"** 로 비활성되고,
> `mock` 이 자동 선택됩니다. 그래서 **clone 직후에도 후원 흐름 전체를 볼 수 있습니다.**
>
> 실제 결제창을 보려면 로컬에서만 채워 쓰세요.
> - 토스 — <https://developers.tosspayments.com> → 내 개발정보 (무료, `test_` 키는 돈이 안 움직임)
> - 이니시스 — <https://manual.inicis.com> 웹표준 연동 가이드에 **공개 테스트 상점 정보**가 있음
>
> Classic ASP 는 조건부 `include` 가 안 되므로 "로컬 설정 파일" 분리가 불가능합니다.
> 키를 채운 뒤에는 **커밋 전에 `git diff App.Conf.asp` 로 반드시 확인**하세요.
>
> `CRYPT_KEY`(쿠키 암호화) 도 플레이스홀더입니다. 배포할 때마다 임의 문자열로 바꾸세요.

```
asp/model/donate/Donate.Class.asp     주문·상태·후원자목록·PG로그
asp/view/donate/Default.asp           안내 + 금액 선택 + 후원자 목록 + ★모의 결제창(모달)
asp/view/donate/Order_p.asp           (AJAX) ★ 금액 재검증 → 주문 생성(ready)
asp/view/donate/Approve_p.asp         (AJAX) 모의 결제창의 승인 처리
asp/view/donate/Pay_p.asp             JS 없는 환경용 폴백 (주문 생성 → Pay.asp)
asp/view/donate/Pay.asp               폴백 화면 / 토스 결제창(JS SDK)
asp/view/donate/PayResult.asp         결과 화면. 토스는 여기서 ★서버 승인★ 을 한다
```

모의 결제창은 실제 PG 창의 순서를 그대로 따라간다.

```
① 주문 생성(서버)  →  ② 결제 정보  →  ③ 승인 요청(서버)  →  ④ 완료
```

모달 상단에 이 4단계가 표시되고 진행에 따라 채워진다. 결제수단 탭(신용카드·계좌이체·간편결제)도
동작한다. **카드번호 칸은 테스트 값이 박힌 읽기 전용이다** — 실제로 입력받으면 그 화면이 곧
카드정보 수집 폼이 되기 때문이고, 실제 결제에서도 카드번호는 PG 창이 직접 받아 우리 서버를
지나가지 않는다. 서버로 전송하지도 않는다.

**결제창을 닫았을 때** — 흔히 빠뜨리는 자리다.

| 상황 | 처리 |
|---|---|
| 토스 창을 닫음 | SDK 가 `USER_CANCEL` 로 reject → `Cancel_p.asp` 로 주문을 `failed` 로 정리 |
| 이니시스 창을 닫음 | `closeUrl` = **`InicisClose.asp`** (레이아웃 없는 전용 페이지) |
| 시연용 모달을 닫음 | 같은 `Cancel_p.asp` 경로 |

두 가지가 이것 때문에 깨졌던 적이 있다.

1. **`closeUrl` 에 일반 페이지를 주면 안 된다.** 이니시스는 결제창 iframe 안에서 `closeUrl` 을
   열기 때문에, 사이트 전체가 결제창 안에 그려져 이상한 화면이 된다.
   `InicisClose.asp` 는 레이아웃을 쓰지 않고 `window.top` 을 옮기는 스크립트만 내보낸다.
2. **AJAX 로 제출하는 폼은 공통 중복제출 잠금을 끈다.** `jkwValidate` 의 잠금은
   "페이지가 이동한다"는 전제다. 결제창만 닫히고 페이지는 그대로면 버튼이 영구히 죽는다.
   그래서 후원 폼에 `data-no-lock` 을 주고 JS 가 직접 잠금을 관리한다.

**결제에서 반드시 지킨 4가지**

1. **금액은 서버가 다시 검증한다.** 화면에서 온 금액을 그대로 쓰면 1원 결제가 가능하다.
2. **승인은 서버가 한다.** 결제창이 "성공"을 알려준 것은 승인이 아니다.
   `PayResult` 가 `paymentKey` 로 승인 API 를 직접 호출한다.
3. **승인 전에 금액을 대조한다.** 결제창이 돌려준 금액이 주문 금액과 다르면 승인하지 않는다.
4. **`ready` 상태만 `paid` 로 바꾼다.** 새로고침이나 중복 콜백으로 두 번 승인되지 않는다.

`tb_payment_log` 에 요청·응답 원문을 남긴다. 결제 분쟁과 디버깅에 원문이 반드시 필요하다.
카드번호는 PG 창에서 직접 입력되므로 우리 서버를 지나가지 않는다.

> ⚠️ `PAY_IS_TEST = True` 인 동안 모든 후원은 `is_test = 1` 로 저장된다.
> 실결제를 켜려면 **사업자등록 + PG 심사**가 필요하다.

### 파일 업로드 시 주의 (Classic ASP 고유 제약)

`Request.BinaryRead` 는 `Request.Form` 을 **한 번이라도 읽은 뒤에는 실패**합니다.
그래서 multipart 폼의 `action` 에는 `?bid=...` 처럼 **필요한 키를 QueryString 으로** 붙이고,
`code/board/Board.Code.asp` 는 `Request.QueryString` 만 읽도록 되어 있습니다.
처리 페이지에서는 `Upload.Load()` 를 먼저 호출한 뒤 `Upload.Field()` 로 값을 읽으세요.

---

## 5. 게시판 추가하기 (코드 수정 0줄)

관리자 → **게시판 관리 → 새 게시판 추가** 로 등록하거나, SQL 로:

```sql
INSERT INTO tb_board_master (board_id, board_name, skin_type, auth_write, sort_no)
VALUES ('qna', N'Q&A', 'list', 3, 70);

INSERT INTO tb_menu (parent_id, menu_code, menu_name, menu_url, menu_depth, sort_no)
VALUES ((SELECT menu_id FROM tb_menu WHERE menu_code = 'community'),
        'comm_qna', N'Q&A', '/board/List.asp?bid=qna', 2, 50);
```

`code/board/Board.Code.asp` 의 `MenuCodeByBid()` 에 매핑을 추가하면 GNB 하이라이트까지 맞습니다.
(추가하지 않아도 `comm_<board_id>` 규칙으로 자동 추정합니다)

---

## 6. 운영 전환 체크리스트

- [ ] `App.Conf.asp` → `DEBUG_MODE = False`
- [ ] `App.Conf.asp` → `PHONE_AUTH_MODE` 가 `"mock"` 이면 **반드시 바꾼다** (4-1장)
- [ ] `App.Conf.asp` → `SOCIAL_MODE` 를 `"live"` 로, 각 제공자 키 등록 (4-2장)
- [ ] `App.Conf.asp` → `PAY_PROVIDER` / `PAY_IS_TEST` 확인. 실결제는 사업자등록·PG 심사 후 (4-3장)
- [ ] 소셜 리다이렉트 URI 를 **https 주소로 다시 등록** (각 제공자 콘솔)
- [ ] `TOSS_SECRET_KEY` 가 소스에 남지 않는지 확인 — `git diff App.Conf.asp` (서버 전용 키)
- [ ] `web.config` → `<asp scriptErrorSentToBrowser="false">` · `httpErrors errorMode="Custom"`
- [ ] `App.Conf.asp` → `CRYPT_KEY` 변경
- [ ] `admin` / `test` 계정 비밀번호 변경 (또는 `test` 삭제)
- [ ] DB 전용 계정(`jkwcom_app`) 사용, `sa` 미사용
- [ ] SQL Server 에이전트에 야간 배치 등록
  ```sql
  EXEC dbo.usp_visit_stat_aggregate;
  EXEC dbo.usp_visit_log_purge 90;
  ```
- [ ] HTTPS 적용 후 `lib/Cookie.Class.asp` 에 `Secure` 속성 추가

---

## 7. 다른 스택

같은 서비스를 다른 언어/렌더링으로 만든 **별도 리포지토리**들입니다.
각각 자기 `assets` · `upload` · `db` · `docs` 를 들고 있어 단독으로 돕니다.

| 리포지토리 | 스택 | 포트 | DB |
|---|---|---|---|
| **이 리포지토리** | Classic ASP | IIS 8080 | **MS SQL Server** `jkwcom` |
| `jkw-springboot` | Java 21 / Spring Boot 3.3 / Thymeleaf | 8090 (내장 톰캣) | MS SQL Server `jkwcom` |
| `jkw-django` | Python 3.12 / Django 5.2 | 8000 | **PostgreSQL** `jkwcom` |
| `jkw-react` | React 19 + Vite (백엔드는 동봉된 ASP JSON API) | IIS 8080 | MS SQL Server `jkwcom` |

**달라지는 축은 넷뿐** — 라우팅 · ORM · 템플릿 · 인증.
CSS/JS 는 같은 파일의 사본이고, 화면도 같습니다. 그래야 비교가 성립합니다.

### 같은 DB 를 공유할 때 주의

ASP · Spring Boot · React 스택은 **같은 `jkwcom` 데이터베이스**를 볼 수 있습니다.
그런데 업로드 폴더는 스택마다 다릅니다(`/asp/upload`, `/react/upload`, Spring 은 자기 `./upload`).
그래서 DB 에 저장된 경로를 **읽을 때 자기 기준으로 다시 붙입니다.**

| 스택 | 정규화 지점 |
|---|---|
| Classic ASP | `code/Common.Code.asp` 의 `UpUrl()` |
| Spring Boot | `common/UploadPath.java` → 템플릿에서 `${up.url(...)}` |
| Django | `common/upload.py` 의 `up_url()` → 템플릿 필터 `|upurl` |

세 구현 모두 **멱등**입니다. 어느 스택이 쓴 값이든 안전하게 읽힙니다.
파일 자체는 공유되지 않으므로, 다른 스택이 올린 첨부를 읽으려면 `upload` 폴더를 복사하세요.

MSSQL → PostgreSQL 타입 매핑표와 이식 규칙은 `docs/DATABASE.md` 8장에 정리돼 있습니다.

---

## 8. 트러블슈팅

| 증상 | 원인 / 해결 |
|---|---|
| **HTTP 500.19** `구성 파일이 잘못된 형식의 XML` | web.config 주석 안에 하이픈 2개(`--`)가 있음. XML 주석은 `--` 를 허용하지 않는다 |
| **HTTP 500.19** `이 구성 섹션은 이 경로에서 사용할 수 없습니다` | `appcmd unlock config -section:system.webServer/asp` (1-2장) |
| **HTTP 500.19** `중복 컬렉션 항목` | `defaultDocument` 의 `value` 는 대소문자 무시 고유값. `default.asp` 와 `Default.asp` 동시 등록 불가 |
| ASP 파일이 그냥 다운로드됨 | Classic ASP 기능 미설치 → `Enable-WindowsOptionalFeature -Online -FeatureName IIS-ASP -All` |
| **800a0411** `다시 정의된 이름` | ① 같은 이름을 두 번 `Dim` ② 매개변수 이름이 함수명과 같음 ③ 전역 함수명과 같은 변수명 → `tools\CheckAsp.ps1` 로 위치 확인 |
| 레이아웃 include 때문에 변수 충돌 | `view/include/*` 의 지역변수는 모든 페이지에 노출된다. 반드시 접두어(`gnb`, `lnb`, `ft`, `hd`)를 붙일 것 |
| `DB 접속` 오류 화면 | SQL Server 서비스 중지 → `Start-Service MSSQLSERVER` |
| `DB 접속` 오류에 `Login failed for user 'NT AUTHORITY\IUSR'` | 익명 인증이 앱 풀 ID 가 아니라 IUSR 로 동작 중. 1-2장 (4) 실행 |
| **SQL Server 서비스가 시작 직후 종료**<br>ERRORLOG 에 `misaligned log IOs` | 디스크가 **물리 섹터 4KB 초과**로 보고될 때 SQL Server 가 기동하지 못한다. NVMe 에서 흔하다.<br>`Get-Disk` 로 `PhysicalSectorSize` 확인 → 4096 초과면 아래를 넣고 **재부팅**<br>`HKLM\SYSTEM\CurrentControlSet\Services\stornvme\Parameters\Device`<br>`ForcedPhysicalSectorSizeInBytes` = `REG_MULTI_SZ` = `* 4095`<br>되돌리려면 이 값을 지우고 재부팅 |
| ERRORLOG 의 msdb/model 경로가 `d:\dbs\sh\s19s\...\mkmastr.proj\` | 설치 후처리가 실패해 `master` 가 **템플릿 상태**로 남은 것. (위 섹터 문제로 설치 중 엔진이 죽으면 이렇게 된다)<br>① 그 경로를 실제로 만들고 `DATA` 의 `model.mdf` `modellog.ldf` `MSDBData.mdf` `MSDBLog.ldf` 를 복사 → 엔진 기동<br>② `sqlservr.exe -c -m -sMSSQLSERVER` (단일 사용자) 로 접속해 `ALTER DATABASE … MODIFY FILE` 로 tempdb·model·msdb 경로를 `DATA` 로 되돌림<br>③ 서비스 재시작 후 임시 폴더 삭제<br>템플릿 master 에는 로그인이 없으므로 ②에서 `CREATE LOGIN … FROM WINDOWS` + `ALTER SERVER ROLE sysadmin ADD MEMBER` 도 함께 해줘야 한다 |
| SSMS/sqlcmd 는 되는데 관리자 권한 없이 실행하면 권한 부족 | `BUILTIN\Administrators` 로만 sysadmin 이 걸려 있으면 UAC 필터 토큰 때문에 비승격 실행 시 권한이 없다. 계정 자체를 sysadmin 에 직접 추가할 것 |
| `개체가 필요합니다: 'Config'` | include 순서 문제. Code 파일 첫 줄이 상대경로 include (`../Common.Code.asp` 류) 인지 확인 |
| `Option Explicit 은 첫 구문` | `<% Option Explicit %>` 가 view 파일 2번째 줄에 있는지 확인 |
| 화면에서 한글이 깨짐 | ① 파일이 **UTF-8** 로 저장됐는지 ② view 1행에 `CODEPAGE="65001"` 이 있는지 (3-1장 참고) |
| DB 데이터 한글이 `???` | `sqlcmd` 에 `-f 65001` 을 빼고 실행했을 때. 스크립트를 다시 실행하면 복구됩니다 |
| 폼 입력 한글이 깨짐 | `Session.CodePage = 65001` 확인. 파일 업로드 폼이면 `Upload.Field()` 를 쓰고 있는지 확인 |
| **multipart 폼의 한글만 깨져 저장됨** | ADODB.Stream 은 `iso-8859-1` 을 **읽을 때 CP1252, 쓸 때 ISO-8859-1** 로 처리해 바이트 왕복이 무손실이 아니다. UTF-8 한글의 뒤따르는 바이트(0x80~0xBF)가 정면으로 걸린다.<br>값은 문자열로 바꾼 본문이 아니라 **원본 바이트 오프셋에서 UTF-8 로 직접** 읽어야 한다 (`lib/FileUpload.Class.asp` 의 `ReadUtf8At`) |
| **800a0006** `오버플로` (해시·누적 계산) | VBScript `Mod` 는 피연산자를 Long(32비트)으로 변환한다. 중간값이 2,147,483,647 을 넘을 수 있으면 `h - Int(h / m) * m` 으로 계산할 것 (`model/common/Auth.Class.asp` 의 `Hash32`) |
| 첫 게시글/첫 항목의 id 가 0 이라 상세 화면이 404 | `DBCC CHECKIDENT(…, RESEED, 0)` 은 **INSERT 이력이 없는 테이블**에서는 첫 id 를 0 으로 만든다 (이력이 있으면 1). `id <= 0` 을 잘못된 요청으로 보는 코드와 충돌한다. `db/mssql/02_seed.sql` 의 IDENTITY 초기화 블록은 `last_value` 로 두 경우를 구분한다 |
| 파일 업로드가 안 됨 | ① `web.config` 의 `maxRequestEntityAllowed` (★ `<asp>` 섹션이 애플리케이션 루트에서만 적용된다 — 하위 폴더면 IIS 기본 200KB 가 그대로 살아있다) ② 폼 `action` 에 `?bid=` 가 있는지 ③ `upload` 폴더 쓰기 권한 |
| 업로드는 되는데 필드값이 빈값 | multipart 인데 `Request.Form` 을 먼저 읽었을 가능성. `Upload.Field()` 사용 |
| **첨부 다운로드가 오류 없이 연결만 끊김**<br>(브라우저 "연결할 수 없음", curl exit 56) | `Response.Buffer = False` 상태에서 `Response.Flush` 를 부르면 Classic ASP 가 런타임 오류를 낸다. 그 시점엔 헤더가 이미 전송돼 오류 화면조차 못 보내고 IIS 가 연결을 끊는다.<br>→ 청크 전송은 **버퍼링을 켠 채로** Flush 한다 (`lib/FileDownload.Class.asp`). 전송량 누적도 `mChunkSize` 가산이 아니라 `st.Position` 으로 해야 마지막 청크가 어긋나지 않는다 |
| 관리자가 방명록(비회원) 글을 삭제할 수 없음 | 비회원 글은 비밀번호 확인(`PwCheck`)을 거치는데, 관리자까지 그 경로로 보내면 남의 비밀번호를 모르니 스팸을 지울 수 없다.<br>→ `view/board/View.asp` 는 `artWriterId = 0 And Not bdCanAdmin` 일 때만 `PwCheck` 로 보낸다 (`Delete_p` 서버측은 원래부터 관리자 허용) |
| 스타일이 안 바뀜 | 관리자 → 사이트 설정 → `asset_ver` 값 변경 (캐시 무효화). `tb_config` 에 이 행이 없으면 `App.Conf.asp` 기본값으로 고정되어 영원히 캐시된다 |
| **목록의 상세 링크가 전부 `id=0`** | `IsNumeric()` 이 **BIGINT/DECIMAL 컬럼을 숫자로 보지 않는다.** ADO 는 이 컬럼을 `VT_DECIMAL`(VarType 14)로 돌려주고, VBScript 의 `IsNumeric` 은 여기에 `False` 를 반환한다. `TypeName()` 은 아예 오류 458 을 낸다.<br>→ 숫자 판정은 `CStr(v)` 로 바꿔 한 뒤에 한다. `StringUtil.ToInt` / `StringUtil.IsNum` / `DBHelper` 의 `SetParamInt`·`SetParamBigInt` 가 이 방식이다.<br>※ `article_id` `comment_id` `file_id` `group_no` 가 BIGINT 이므로 이 함정에 걸린다 |
| 서버 검증에 걸린 뒤 **전송 버튼이 계속 비활성** | `JsAlert` 은 `alert()` 후 `history.back()` 으로 되돌린다. 그런데 `common.js` 가 중복 제출 방지로 걸어둔 `disabled` 가 복원된 DOM 에 그대로 남아 재전송이 막힌다.<br>→ `assets/js/common.js` (이 리포지토리 안의 사본) 가 `pageshow` 에서 `disabled` 를 해제한다 (최초 로드·뒤로가기 복원 모두에서 발생) |
| 500 오류 상세를 보고 싶음 | `web.config` 의 `httpErrors errorMode="Detailed"` + `<asp scriptErrorSentToBrowser="true">`.<br>★ 상세 오류가 안 보이면 `<asp>` 섹션이 적용되지 않은 것이다 — 이 폴더가 IIS **애플리케이션 루트**인지 확인할 것 (1-2장) |
