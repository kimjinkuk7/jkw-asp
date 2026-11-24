<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/contact/Contact.Class.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<%
'==============================================================================
' code/contact/Default.Code.asp  -  문의하기 폼
'==============================================================================
g_menuCode  = "contact"
g_pageTitle = "문의하기"
g_pageDesc  = "프로젝트 의뢰, 채용 제안, 기타 문의를 남겨주세요."
g_bodyClass = "page-contact"

'--- 로그인 상태면 폼 기본값 채우기 ---
Dim ctName, ctEmail
ctName  = ""
ctEmail = ""
If GLB_IS_LOGIN Then
    ctName  = Auth.MyName()
    ctEmail = ""
End If

Dim ctTypes
ctTypes = Array("프로젝트 문의", "채용 제안", "협업 제안", "기타")

Dim ctResponseTime
ctResponseTime = Config.ValueInit("contact_reply_time", "보통 1~2 영업일 이내에 회신드립니다.")
%>
