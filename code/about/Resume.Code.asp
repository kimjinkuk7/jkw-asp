<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<%
'==============================================================================
' code/about/Resume.Code.asp  -  이력서 (인쇄 최적화)
'------------------------------------------------------------------------------
' style.css 의 @media print 규칙으로 헤더/푸터/버튼이 빠진 상태로 인쇄된다.
'==============================================================================
g_menuCode  = "about_resume"
g_lnbParent = "about"
g_pageTitle = "이력서"
g_pageDesc  = GLB_OWNER_NAME & " 이력서"
g_bodyClass = "page-about page-resume"

Dim rsIntro, rsYears
rsIntro = Config.ValueInit("resume_summary", Config.ValueInit("about_intro", GLB_SITE_DESC))
rsYears = About.TotalYears()

Dim rsWork, rsEdu, rsCert, rsAward, rsSkillGrp, rsSkill, rsPf
Dim rsWorkCnt, rsEduCnt, rsCertCnt, rsAwardCnt, rsSkillGrpCnt, rsSkillCnt, rsPfCnt

rsWork     = About.Career_l("work")  : rsWorkCnt     = RowCnt(rsWork)
rsEdu      = About.Career_l("edu")   : rsEduCnt      = RowCnt(rsEdu)
rsCert     = About.Career_l("cert")  : rsCertCnt     = RowCnt(rsCert)
rsAward    = About.Career_l("award") : rsAwardCnt    = RowCnt(rsAward)
rsSkillGrp = About.SkillGroup_l()    : rsSkillGrpCnt = RowCnt(rsSkillGrp)
rsSkill    = About.Skill_l("")       : rsSkillCnt    = RowCnt(rsSkill)
rsPf       = Portfolio.Main_l(6)     : rsPfCnt       = RowCnt(rsPf)
%>
