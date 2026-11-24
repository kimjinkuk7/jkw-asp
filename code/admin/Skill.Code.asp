<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Skill.Code.asp  -  기술 스택 관리
'==============================================================================
g_adminNav  = "skill"
g_pageTitle = "기술 스택 관리"

Dim asArr, asCnt
asArr = About.SkillAll_l()
asCnt = RowCnt(asArr)

'--- SkillAll_l 추가 컬럼 ---
Const SKA_SORTNO = 7
Const SKA_ISUSE  = 8

'--- 편집 대상 (목록에서 찾는다) ---
Dim asId, asIdx, asIsEdit, k
asId     = RequestUtil.ValueIntInit("id", 0)
asIdx    = -1
asIsEdit = False

If asId > 0 Then
    For k = 0 To asCnt - 1
        If AVI(asArr, SK_ID, k) = asId Then asIdx = k : asIsEdit = True
    Next
End If

Dim sGroup, sName, sLevel, sExp, sDesc, sColor, sSort, sUse
If asIsEdit Then
    sGroup = AV(asArr,  SK_GROUP, asIdx)
    sName  = AV(asArr,  SK_NAME, asIdx)
    sLevel = AVI(asArr, SK_LEVEL, asIdx)
    sExp   = AV(asArr,  SK_EXP, asIdx)
    sDesc  = AV(asArr,  SK_DESC, asIdx)
    sColor = AV(asArr,  SK_COLOR, asIdx)
    sSort  = AVI(asArr, SKA_SORTNO, asIdx)
    sUse   = AVB(asArr, SKA_ISUSE, asIdx)
Else
    sGroup = "Language" : sName = "" : sLevel = 70 : sExp = 1
    sDesc = "" : sColor = "" : sSort = 10 : sUse = True
End If

'--- 그룹 선택지 (기존 그룹 + 기본값) ---
Dim asGroups
asGroups = Array("Language", "Framework", "Database", "DevOps", "Tool", "ETC")
%>
