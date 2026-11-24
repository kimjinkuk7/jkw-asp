<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="App.Conf.asp"-->
<%
'==============================================================================
' asp/default.asp  -  ASP 스택 진입점  ( /asp/ )
'------------------------------------------------------------------------------
' 실제 화면은 view/ 아래에 있다. 여기서는 메인으로 보내기만 한다.
' 보내는 주소는 App.Conf.asp 의 VIEW_ROOT 가 정한다 (U 함수).
'
' 이 폴더가 IIS 사이트 루트면  /         -> /view/main/Default.asp
' 사이트 루트 아래 asp 폴더면  /asp/     -> /asp/view/main/Default.asp
'==============================================================================
Response.Status = "302 Found"
Response.AddHeader "Location", U("/main/Default.asp")
Response.End
%>
