<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/donate/Donate.Class.asp"-->
<%
'==============================================================================
' code/donate/Default.Code.asp  -  개발자 지원금 안내 + 후원자 목록
'==============================================================================
g_menuCode  = "donate"
g_bodyClass = "page-donate"
g_pageTitle = "개발자 지원금"
g_pageDesc  = "이 사이트가 마음에 드셨다면 커피 한 잔 값으로 응원해 주세요."

'--- 후원자 목록 / 집계 ---
Dim dnList, dnCnt, dnStat, dnTotalCnt, dnTotalAmt
dnList = Donate.Public_l(30)
dnCnt  = RowCnt(dnList)

dnStat     = Donate.Stat_v()
dnTotalCnt = 0
dnTotalAmt = 0
If RowCnt(dnStat) > 0 Then
    dnTotalCnt = AVI(dnStat, 0, 0)
    dnTotalAmt = AVI(dnStat, 1, 0)
End If

'--- 금액 선택 버튼 ---
Dim dnPresets
dnPresets = Array(1000, 3000, 5000, 10000)

'--- 선택 가능한 결제 모듈 ---
Dim dnPgList, dnPgCount, dnDefault, dnHasMock

dnPgList  = Split(LCase(Replace(PAY_PROVIDERS, " ", "")), ",")
dnDefault = LCase(PAY_PROVIDER)
dnHasMock = (InStr("," & LCase(PAY_PROVIDERS) & ",", ",mock,") > 0)

'<Desc>결제 모듈 표시명</Desc>
Function PgLabel(ByVal p)
    Select Case LCase(CStr(p))
        Case "mock"   : PgLabel = "시연용 (자체 모달)"
        Case "toss"   : PgLabel = "토스페이먼츠"
        Case "inicis" : PgLabel = "KG이니시스"
        Case Else     : PgLabel = CStr(p)
    End Select
End Function

'<Desc>결제 모듈 설명</Desc>
Function PgDesc(ByVal p)
    Select Case LCase(CStr(p))
        Case "mock"   : PgDesc = "PG 를 거치지 않고 흐름만 확인합니다."
        Case "toss"   : PgDesc = "토스가 제공하는 실제 결제창이 열립니다."
        Case "inicis" : PgDesc = "이니시스가 제공하는 실제 결제창이 열립니다."
        Case Else     : PgDesc = ""
    End Select
End Function

'<Desc>그 모듈이 실제로 준비됐는지 (키 확인)</Desc>
Function PgReady(ByVal p)
    Select Case LCase(CStr(p))
        Case "mock"   : PgReady = True
        Case "toss"   : PgReady = (Len(TOSS_CLIENT_KEY) > 0 And Len(TOSS_SECRET_KEY) > 0)
        Case "inicis" : PgReady = (Len(INICIS_MID) > 0 And Len(INICIS_SIGNKEY) > 0)
        Case Else     : PgReady = False
    End Select
End Function

'--- 기본 후원자명 (로그인했으면 닉네임) ---
Dim dnDefName
dnDefName = ""
If GLB_IS_LOGIN Then dnDefName = GLB_MY_NAME
If Len(dnDefName) = 0 And GLB_IS_LOGIN Then dnDefName = GLB_MY_ID
%>
