<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Portfolio.Code.asp  -  포트폴리오 목록 관리
'==============================================================================
g_adminNav  = "pf"
g_pageTitle = "포트폴리오 관리"

Dim apArr, apCnt
apArr = Portfolio.AdminAll_l()
apCnt = RowCnt(apArr)

'--- AdminAll_l 의 추가 컬럼 인덱스 ---
Const PFA_CONTENT = 15
Const PFA_TEAM    = 16
Const PFA_ISUSE   = 17
Const PFA_SORTNO  = 18
%>
