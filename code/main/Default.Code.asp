<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/board/Board.Class.asp"-->
<!--#include file="../../model/board/Article.Class.asp"-->
<!--#include file="../../model/portfolio/Portfolio.Class.asp"-->
<!--#include file="../../model/about/About.Class.asp"-->
<%
'==============================================================================
' code/main/Default.Code.asp  -  메인 화면 컨트롤러
'==============================================================================

g_menuCode  = "home"
g_pageTitle = GLB_SITE_TITLE
g_pageDesc  = GLB_SITE_DESC
g_bodyClass = "page-main"

'--- 히어로 문구 ---
Dim heroHeadline, heroSub
heroHeadline = Config.ValueInit("hero_headline", GLB_SITE_TITLE)
heroSub      = Config.ValueInit("hero_sub", GLB_SITE_DESC)

'--- 대표 포트폴리오 ---
Dim mainPfArr, mainPfCnt
mainPfArr = Portfolio.Main_l(3)
mainPfCnt = RowCnt(mainPfArr)

'--- 최근 글 ---
Dim blogArr, blogCnt, noticeArr, noticeCnt
blogArr   = Article.Recent_l("blog", 5)
blogCnt   = RowCnt(blogArr)
noticeArr = Article.Recent_l("notice", 5)
noticeCnt = RowCnt(noticeArr)

'--- 대표 기술 스택 (히어로 하단 칩) ---
Dim skillArr, skillCnt, mainSkillMax
skillArr     = About.Skill_l("Language")
skillCnt     = RowCnt(skillArr)
mainSkillMax = 7

'--- 지표 ---
Dim statYears, statProject, statPost, statVisit, vcArr
statYears   = About.TotalYears()
statProject = Portfolio.Portfolio_c("")
statPost    = Article.Article_c("blog", 0, "", "")

vcArr = Visit.Counter_v()
If IsArray(vcArr) Then statVisit = AVI(vcArr, 1, 0) Else statVisit = 0
%>
