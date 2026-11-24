<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<!--#include file="../../model/member/Social.Class.asp"-->
<%
'==============================================================================
' code/member/SocialCallback.Code.asp  -  소셜 로그인 콜백
'------------------------------------------------------------------------------
' 처리 순서
'   1) state 대조 (세션과 다르면 즉시 중단)
'   2) code → 토큰 → 사용자 정보
'   3) (provider, uid) 로 연동 조회
'        있으면            → 그 회원으로 로그인
'        없고 이메일 일치  → 기존 회원에 연동을 붙이고 로그인
'        둘 다 없으면      → 회원을 새로 만들고 연동 후 로그인
'==============================================================================

Dim scProvider, scCode, scState, scErr, scToken
Dim scUid, scEmail, scName, scMemberId, scArr, scRurl, scNew, scPopup, scAcct

scProvider = LCase(RequestUtil.ValueInit("provider", ""))
scCode     = RequestUtil.ValueInit("code", "")
scState    = RequestUtil.ValueInit("state", "")
scAcct     = RequestUtil.ValueInit("acct", "1")
scPopup    = (RequestUtil.ValueInit("popup", "") = "1")
scNew      = False

'<Desc>팝업이면 부모 창을 옮기고 팝업을 닫는다. 아니면 평소처럼 이동한다.</Desc>
Sub GoDone(ByVal pUrl, ByVal pMsg)
    If Not scPopup Then
        If Len(pMsg) > 0 Then JsAlertGo pMsg, pUrl Else JsGo pUrl
        Exit Sub
    End If
    Response.Clear
    Response.ContentType = "text/html"
    Response.Write "<!DOCTYPE html><html lang=""ko""><head><meta charset=""utf-8"">" & _
                   "<title>로그인 처리 중</title></head><body>" & _
                   "<p style=""font:14px sans-serif;padding:24px"">로그인 처리 중입니다…</p>" & _
                   "<script>(function(){" & _
                   "  var msg=" & IIfStr(Len(pMsg) > 0, "'" & JE(pMsg) & "'", "''") & ";" & _
                   "  if(msg) alert(msg);" & _
                   "  try{ if(window.opener && !window.opener.closed){" & _
                   "    window.opener.location.replace('" & JE(pUrl) & "'); window.close(); return; } }catch(e){}" & _
                   "  location.replace('" & JE(pUrl) & "');" & _
                   "})();</script></body></html>"
    Response.End
End Sub

If Not Social.IsSupported(scProvider) Then JsAlert "지원하지 않는 로그인 방식입니다."

'--- 제공자가 거절한 경우 (사용자가 동의 취소 등) ---
If Len(RequestUtil.ValueInit("error", "")) > 0 Then
    JsAlertGo Social.Label(scProvider) & " 로그인이 취소되었습니다.", U("/member/Login.asp")
End If

'--- 1) state 대조 ---
If Len(scState) = 0 Or scState <> (Session("SOCIAL_STATE") & "") Then
    JsAlertGo "로그인 요청이 유효하지 않습니다. 다시 시도해 주세요.", U("/member/Login.asp")
End If
If scProvider <> (Session("SOCIAL_PROVIDER") & "") Then
    JsAlertGo "로그인 요청이 유효하지 않습니다. 다시 시도해 주세요.", U("/member/Login.asp")
End If
Session("SOCIAL_STATE") = ""            ' 재사용 방지

scRurl = Session("SOCIAL_RURL") & ""
If Len(scRurl) = 0 Then scRurl = U("/main/Default.asp")

'--- 2) 사용자 정보 확보 ---
If LCase(SOCIAL_MODE) = "mock" Then
    ' 모의 화면에서 고른 계정. uid 가 다르면 서로 다른 회원이 된다.
    If scAcct <> "2" Then scAcct = "1"
    scUid   = "mock-" & scProvider & "-" & scAcct
    scEmail = scProvider & IIfStr(scAcct = "2", ".demo2", ".demo") & "@example.com"
    scName  = Social.Label(scProvider) & " 데모" & IIfStr(scAcct = "2", "2", "")
Else
    If Len(scCode) = 0 Then JsAlertGo "인증 코드가 없습니다.", U("/member/Login.asp")

    scToken = Social.Exchange(scProvider, scCode, scState, scErr)
    If Len(scToken) = 0 Then JsAlertGo "로그인 실패 : " & scErr, U("/member/Login.asp")

    If Not Social.Profile(scProvider, scToken, scUid, scEmail, scName, scErr) Then
        JsAlertGo "로그인 실패 : " & scErr, U("/member/Login.asp")
    End If
End If

'--- 3) 연동 조회 → 로그인 ---
scMemberId = Social.FindMemberId(scProvider, scUid)

If scMemberId <= 0 Then
    ' 이메일이 같은 기존 회원이 있으면 그 계정에 연동을 붙인다
    scMemberId = Social.FindMemberIdByEmail(scEmail)

    If scMemberId <= 0 Then
        scMemberId = Social.CreateMember(scProvider, scUid, scEmail, scName, scErr)
        If scMemberId <= 0 Then JsAlertGo "가입 처리 실패 : " & scErr, U("/member/Login.asp")
        scNew = True
    End If

    Social.Link_i scMemberId, scProvider, scUid, scEmail, scName
End If

Social.TouchLogin_u scProvider, scUid

'--- 세션 발급 ---
scArr = Member.Member_v(scMemberId)
If Not IsArray(scArr) Then JsAlertGo "회원 정보를 찾을 수 없습니다.", U("/member/Login.asp")

If Not AVB(scArr, MB_ISUSE, 0) Then
    JsAlertGo "정지되었거나 사용할 수 없는 계정입니다.", U("/member/Login.asp")
End If

Auth.SetSession scMemberId, AV(scArr, MB_LOGINID, 0), AVI(scArr, MB_LEVEL, 0), _
                AV(scArr, MB_NICK, 0), AV(scArr, MB_NAME, 0)

Session("SOCIAL_RURL") = ""

If scNew Then
    GoDone scRurl, Social.Label(scProvider) & " 계정으로 가입되었습니다. 환영합니다!"
Else
    GoDone scRurl, ""
End If
%>
