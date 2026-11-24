<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/board/Board.Class.asp"-->
<!--#include file="../../model/board/Article.Class.asp"-->
<!--#include file="../../model/board/Comment.Class.asp"-->
<!--#include file="../../model/board/BoardFile.Class.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<!--#include file="../../model/contact/Contact.Class.asp"-->
<%
'==============================================================================
' code/admin/Admin.Code.asp  -  관리자 모듈 공통
'------------------------------------------------------------------------------
' 모든 admin Code 파일이 첫 줄에 include 한다. (Login 제외)
'==============================================================================

Auth.CheckAdmin

g_bodyClass   = "admin"
g_isAdminPage = True
g_menuCode    = ""

'--- 사이드바 활성 항목 (각 페이지에서 지정) ---
Dim g_adminNav
g_adminNav = ""

'--- 미확인 문의 수 (사이드바 배지) ---
Dim admUnread
admUnread = Contact.Unread_c()
%>
