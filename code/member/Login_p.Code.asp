<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<%
'==============================================================================
' code/member/Login_p.Code.asp  -  로그인 처리
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."
Auth.CheckToken Trim(Request.Form("_csrf") & "")

Dim inId, inPw, inSave, rUrl, rc
inId   = Left(LCase(Trim(Request.Form("login_id") & "")), 30)
inPw   = Left(Trim(Request.Form("user_pw") & ""), 64)
inSave = (Trim(Request.Form("save_id") & "") = "1")
rUrl   = Trim(Request.Form("rurl") & "")

'--- 오픈 리다이렉트 방지 ---
If Len(rUrl) = 0 Or Left(rUrl, 1) <> "/" Or Left(rUrl, 2) = "//" Then
    rUrl = U("/main/Default.asp")
End If

If StringUtil.IsBlank(inId) Then JsAlert "아이디를 입력해 주세요."
If StringUtil.IsBlank(inPw) Then JsAlert "비밀번호를 입력해 주세요."

'--- 아이디 저장 쿠키 ---
If inSave Then
    Cookie.Encrypt "JKW_SAVE_ID", inId
Else
    Cookie.Remove "JKW_SAVE_ID"
End If

'--- 인증 ---
rc = Auth.Login(inId, inPw)

Select Case rc
    Case 1
        JsGo rUrl
    Case -1
        JsAlert "정지되었거나 탈퇴한 계정입니다. 관리자에게 문의해 주세요."
    Case -2
        JsAlert "로그인 실패가 " & LOGIN_FAIL_LIMIT & "회 누적되어 " & LOGIN_LOCK_MIN & "분간 로그인이 제한됩니다."
    Case Else
        JsAlert "아이디 또는 비밀번호가 일치하지 않습니다."
End Select
%>
