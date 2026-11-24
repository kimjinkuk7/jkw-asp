<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<%
'==============================================================================
' code/about/Skill.Code.asp  -  기술 스택
'==============================================================================
g_menuCode  = "about_skill"
g_lnbParent = "about"
g_pageTitle = "기술 스택"
g_pageDesc  = GLB_OWNER_NAME & "가 사용하는 언어 · 프레임워크 · 데이터베이스 · 인프라"
g_bodyClass = "page-about"

Dim skGrpArr, skGrpCnt, skAllArr, skAllCnt
skGrpArr = About.SkillGroup_l()
skGrpCnt = RowCnt(skGrpArr)
skAllArr = About.Skill_l("")
skAllCnt = RowCnt(skAllArr)
%>
