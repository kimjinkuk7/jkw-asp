<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/PortfolioForm.Code.asp  -  포트폴리오 등록 / 수정 폼
'==============================================================================
g_adminNav  = "pf"

Dim pfId, pfArr, pfIsNew
pfId    = RequestUtil.ValueIntInit("id", 0)
pfIsNew = (pfId <= 0)

'--- 폼 값 ---
Dim vTitle, vSubtitle, vSummary, vContent, vCategory, vClient, vRole, vTeam
Dim vStart, vEnd, vOngoing, vSiteUrl, vRepoUrl, vThumb, vIsMain, vSortNo, vIsUse
Dim vSkillIds

vTitle = "" : vSubtitle = "" : vSummary = "" : vContent = ""
vCategory = "Web" : vClient = "" : vRole = "" : vTeam = ""
vStart = "" : vEnd = "" : vOngoing = False
vSiteUrl = "" : vRepoUrl = "" : vThumb = ""
vIsMain = False : vSortNo = 10 : vIsUse = True
vSkillIds = ""

If Not pfIsNew Then
    pfArr = Portfolio.Admin_v(pfId)
    If Not IsArray(pfArr) Then JsAlertGo "존재하지 않는 프로젝트입니다.", U("/admin/Portfolio.asp")

    vTitle    = AV(pfArr,  PF_TITLE, 0)
    vSubtitle = AV(pfArr,  PF_SUBTITLE, 0)
    vSummary  = AV(pfArr,  PF_SUMMARY, 0)
    vContent  = AV(pfArr,  15, 0)
    vCategory = AV(pfArr,  PF_CATEGORY, 0)
    vClient   = AV(pfArr,  PF_CLIENT, 0)
    vRole     = AV(pfArr,  PF_ROLE, 0)
    vTeam     = AV(pfArr,  16, 0)
    vStart    = DateUtil.Fmt(AV(pfArr, PF_START, 0), "ymd")
    vEnd      = DateUtil.Fmt(AV(pfArr, PF_END, 0), "ymd")
    vOngoing  = AVB(pfArr, PF_ONGOING, 0)
    vSiteUrl  = AV(pfArr,  PF_SITEURL, 0)
    vRepoUrl  = AV(pfArr,  PF_REPOURL, 0)
    vThumb    = UpUrl(AV(pfArr,  PF_THUMB, 0))
    vIsMain   = AVB(pfArr, PF_ISMAIN, 0)
    vSortNo   = AVI(pfArr, 18, 0)
    vIsUse    = AVB(pfArr, 17, 0)
    vSkillIds = Portfolio.SkillIds(pfId)
End If

g_pageTitle = IIfStr(pfIsNew, "새 프로젝트", "프로젝트 수정")

'--- 기술 스택 선택지 ---
Dim pfSkillArr, pfSkillCnt
pfSkillArr = About.SkillAll_l()
pfSkillCnt = RowCnt(pfSkillArr)

Function SkillChecked(ByVal pSkillId)
    SkillChecked = (InStr("," & vSkillIds & ",", "," & pSkillId & ",") > 0)
End Function

'--- 카테고리 선택지 ---
Dim pfCates
pfCates = Array("Web", "Mobile", "Solution", "Data", "DevOps", "ETC")
%>
