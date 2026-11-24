<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<%
'==============================================================================
' code/about/Career.Code.asp  -  경력 · 학력 · 수상 · 자격
'==============================================================================
g_menuCode  = "about_career"
g_lnbParent = "about"
g_pageTitle = "경력 · 학력"
g_pageDesc  = GLB_OWNER_NAME & "의 경력, 학력, 수상, 자격 사항"
g_bodyClass = "page-about"

Dim crWork, crEdu, crAward, crCert, crAct
Dim crWorkCnt, crEduCnt, crAwardCnt, crCertCnt, crActCnt

crWork  = About.Career_l("work")     : crWorkCnt  = RowCnt(crWork)
crEdu   = About.Career_l("edu")      : crEduCnt   = RowCnt(crEdu)
crAward = About.Career_l("award")    : crAwardCnt = RowCnt(crAward)
crCert  = About.Career_l("cert")     : crCertCnt  = RowCnt(crCert)
crAct   = About.Career_l("activity") : crActCnt   = RowCnt(crAct)

Dim crYears
crYears = About.TotalYears()
%>
