<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<%
'==============================================================================
' code/portfolio/View.Code.asp  -  포트폴리오 상세 (케이스 스터디)
'==============================================================================
g_menuCode  = "portfolio"
g_bodyClass = "page-portfolio"

Dim pvId, pvArr
pvId = RequestUtil.ValueIntInit("id", 0)
If pvId <= 0 Then Go404 "잘못된 접근입니다."

pvArr = Portfolio.Portfolio_v(pvId)
If Not IsArray(pvArr) Then Go404 "존재하지 않는 프로젝트입니다."

Dim pvTitle, pvSubtitle, pvSummary, pvContent, pvCate, pvClient, pvRole, pvTeam
Dim pvSiteUrl, pvRepoUrl, pvThumb, pvPeriod, pvHit

pvTitle    = AV(pvArr,  PF_TITLE, 0)
pvSubtitle = AV(pvArr,  PF_SUBTITLE, 0)
pvSummary  = AV(pvArr,  PF_SUMMARY, 0)
pvContent  = AV(pvArr,  PFV_CONTENT, 0)
pvCate     = AV(pvArr,  PF_CATEGORY, 0)
pvClient   = AV(pvArr,  PF_CLIENT, 0)
pvRole     = AV(pvArr,  PF_ROLE, 0)
pvTeam     = AV(pvArr,  PFV_TEAM, 0)
pvSiteUrl  = AV(pvArr,  PF_SITEURL, 0)
pvRepoUrl  = AV(pvArr,  PF_REPOURL, 0)
pvThumb    = UpUrl(AV(pvArr,  PF_THUMB, 0))
pvHit      = AVI(pvArr, PF_HIT, 0)
pvPeriod   = DateUtil.Period(AV(pvArr, PF_START, 0), AV(pvArr, PF_END, 0), AVB(pvArr, PF_ONGOING, 0))

'--- 조회수 (세션 중복 제외) ---
Dim pvKey, pvLog
pvKey = "|pf:" & pvId & "|"
pvLog = Session("HITLOG") & ""
If InStr(pvLog, pvKey) = 0 Then
    Portfolio.HitUp pvId
    pvHit = pvHit + 1
    If Len(pvLog) > 3000 Then pvLog = ""
    Session("HITLOG") = pvLog & pvKey
End If

'--- 사용 기술 ---
Dim pvSkillArr, pvSkillCnt
pvSkillArr = Portfolio.Skill_l(pvId)
pvSkillCnt = RowCnt(pvSkillArr)

'--- 이전 / 다음 ---
Dim pvPrev, pvNext
pvPrev = Portfolio.Sibling_v(pvId, "prev")
pvNext = Portfolio.Sibling_v(pvId, "next")

g_pageTitle = pvTitle
g_pageDesc  = pvSummary
%>
