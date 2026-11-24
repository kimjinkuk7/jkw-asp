<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<%
'==============================================================================
' code/about/Default.Code.asp  -  소개
'==============================================================================
g_menuCode    = "about_intro"
g_lnbParent   = "about"
g_pageTitle   = "소개"
g_pageDesc    = GLB_OWNER_NAME & " 소개 - " & GLB_SITE_DESC
g_bodyClass   = "page-about"

Dim abIntro, abYears, abProjectCnt
abIntro      = Config.ValueInit("about_intro", _
    "안녕하세요. 백엔드 개발자 " & GLB_OWNER_NAME & "입니다." & vbCrLf & _
    "레거시와 모던 스택을 오가며, 같은 문제를 언어마다 어떻게 다르게 푸는지 정리하는 일을 좋아합니다." & vbCrLf & _
    "이 사이트는 동일한 서비스를 Classic ASP · Spring Boot · Python · PHP 네 갈래로 구현하며 만든 결과물입니다.")

abYears      = About.TotalYears()
abProjectCnt = Portfolio.Portfolio_c("")

'--- 대표 스킬 (그룹별 상위) ---
Dim abSkillArr, abSkillCnt
abSkillArr = About.Skill_l("")
abSkillCnt = RowCnt(abSkillArr)

'--- 현재 재직 ---
Dim abWorkArr, abWorkCnt
abWorkArr = About.Career_l("work")
abWorkCnt = RowCnt(abWorkArr)
%>
