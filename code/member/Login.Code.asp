<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<!--#include file="../../model/member/Social.Class.asp"-->
<%
'==============================================================================
' code/member/Login.Code.asp  -  로그인 폼
'==============================================================================
g_pageTitle = "로그인"
g_pageDesc  = "회원 로그인"
g_bodyClass = "page-member"

'--- 이미 로그인 상태면 메인으로 ---
If GLB_IS_LOGIN Then JsGo U("/main/Default.asp")

'--- 로그인 후 돌아갈 주소 (오픈 리다이렉트 방지 : 내부 경로만 허용) ---
Dim rUrl
rUrl = RequestUtil.ValueInit("rurl", "")
If Len(rUrl) > 0 Then
    If Left(rUrl, 1) <> "/" Or Left(rUrl, 2) = "//" Then rUrl = ""
End If
If Len(rUrl) = 0 Then rUrl = U("/main/Default.asp")

'--- 아이디 저장 쿠키 ---
Dim savedId
savedId = Cookie.Decrypt("JKW_SAVE_ID")
%>
