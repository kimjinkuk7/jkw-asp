<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Skill_p.Code.asp  -  기술 스택 저장 / 삭제
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim spMode, spId
spMode = LCase(RequestUtil.ValueInit("mode", ""))
spId   = RequestUtil.ValueIntInit("id", 0)

If spMode = "delete" Then
    If spId <= 0 Then JsAlert "잘못된 접근입니다."
    About.Skill_d spId
    JsAlertGo "삭제되었습니다.", U("/admin/Skill.asp")
End If

Dim kGroup, kName, kLevel, kExp, kDesc, kColor, kSort, kUse
kGroup = Left(Trim(Request.Form("skill_group") & ""), 50)
kName  = Left(Trim(Request.Form("skill_name")  & ""), 50)
kLevel = StringUtil.ToInt(Request.Form("level_pct"), 50)
kExp   = Trim(Request.Form("experience_yr") & "")
kDesc  = Left(Trim(Request.Form("skill_desc") & ""), 300)
kColor = Left(Trim(Request.Form("color_hex")  & ""), 7)
kSort  = StringUtil.ToInt(Request.Form("sort_no"), 0)
kUse   = (Trim(Request.Form("is_use") & "") = "1")

If StringUtil.IsBlank(kGroup) Then JsAlert "그룹을 입력해 주세요."
If StringUtil.IsBlank(kName)  Then JsAlert "기술명을 입력해 주세요."
If kLevel < 0   Then kLevel = 0
If kLevel > 100 Then kLevel = 100
If Not IsNumeric(kExp) Then kExp = Null
If Len(kColor) > 0 Then
    If Not StringUtil.RegTest("^#[0-9A-Fa-f]{6}$", kColor) Then kColor = ""
End If

If spMode = "update" And spId > 0 Then
    About.Skill_u spId, kGroup, kName, kLevel, kExp, kDesc, kColor, kSort, kUse
    JsAlertGo "수정되었습니다.", U("/admin/Skill.asp") & "?id=" & spId
ElseIf spMode = "insert" Then
    About.Skill_i kGroup, kName, kLevel, kExp, kDesc, kColor, kSort
    JsAlertGo "추가되었습니다.", U("/admin/Skill.asp")
Else
    JsAlert "잘못된 접근입니다."
End If
%>
