<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Career_p.Code.asp  -  경력 항목 저장 / 삭제
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim cpMode, cpId
cpMode = LCase(RequestUtil.ValueInit("mode", ""))
cpId   = RequestUtil.ValueIntInit("id", 0)

If cpMode = "delete" Then
    If cpId <= 0 Then JsAlert "잘못된 접근입니다."
    About.Career_d cpId
    JsAlertGo "삭제되었습니다.", U("/admin/Career.asp")
End If

' cCur 라는 이름은 VBScript 내장 함수 CCur() 를 가리므로 쓰지 않는다.
Dim cType, cTitle, cPos, cUrl, cDesc, cStart, cEnd, cIsCurrent, cSort, cUse
cType      = LCase(Trim(Request.Form("career_type") & ""))
cTitle     = Left(Trim(Request.Form("title")        & ""), 200)
cPos       = Left(Trim(Request.Form("position")     & ""), 100)
cUrl       = Left(Trim(Request.Form("org_url")      & ""), 300)
cDesc      = Left(Trim(Request.Form("description")  & ""), 2000)
cStart     = Trim(Request.Form("date_start") & "")
cEnd       = Trim(Request.Form("date_end")   & "")
cIsCurrent = (Trim(Request.Form("is_current") & "") = "1")
cSort  = StringUtil.ToInt(Request.Form("sort_no"), 0)
cUse   = (Trim(Request.Form("is_use") & "") = "1")

If InStr(",work,edu,cert,award,activity,", "," & cType & ",") = 0 Then JsAlert "구분 값이 올바르지 않습니다."
If StringUtil.IsBlank(cTitle) Then JsAlert "제목을 입력해 주세요."
If Not IsDate(cStart) Then cStart = Null
If Not IsDate(cEnd)   Then cEnd   = Null

If cpMode = "update" And cpId > 0 Then
    About.Career_u cpId, cType, cTitle, cPos, cUrl, cDesc, cStart, cEnd, cIsCurrent, cSort, cUse
    JsAlertGo "수정되었습니다.", U("/admin/Career.asp") & "?id=" & cpId
ElseIf cpMode = "insert" Then
    About.Career_i cType, cTitle, cPos, cUrl, cDesc, cStart, cEnd, cIsCurrent, cSort
    JsAlertGo "추가되었습니다.", U("/admin/Career.asp")
Else
    JsAlert "잘못된 접근입니다."
End If
%>
