<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<%
'==============================================================================
' code/member/Join.Code.asp  -  회원가입 폼
'==============================================================================
g_pageTitle = "회원가입"
g_pageDesc  = "회원가입 후 게시판 작성과 댓글 기능을 이용할 수 있습니다."
g_bodyClass = "page-member"

If GLB_IS_LOGIN Then JsAlertGo "이미 로그인되어 있습니다.", U("/main/Default.asp")

'--- 입력값 되살리기 (검증 실패 후 돌아온 경우) ---
Dim jLoginId, jName, jNick, jEmail, jMobile, jIntro
jLoginId = Session("JOIN_TMP_ID")    & ""
jName    = Session("JOIN_TMP_NAME")  & ""
jNick    = Session("JOIN_TMP_NICK")  & ""
jEmail   = Session("JOIN_TMP_EMAIL") & ""
jMobile  = Session("JOIN_TMP_MOBILE")& ""
jIntro   = Session("JOIN_TMP_INTRO") & ""

'--- 1회 사용 후 폐기 ---
Session("JOIN_TMP_ID")     = ""
Session("JOIN_TMP_NAME")   = ""
Session("JOIN_TMP_NICK")   = ""
Session("JOIN_TMP_EMAIL")  = ""
Session("JOIN_TMP_MOBILE") = ""
Session("JOIN_TMP_INTRO")  = ""
%>
