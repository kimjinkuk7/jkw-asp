<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/PwCheck.Code.asp  -  비회원 글 비밀번호 확인
'------------------------------------------------------------------------------
' act : read(비밀글 열람) | modify(수정) | delete(삭제)
' 확인에 성공하면 Session("PWOK_<bid>_<id>") = "1" 을 세팅하고 해당 화면으로 보낸다.
'==============================================================================

Dim pcId, pcAct, pcArr
pcId  = RequestUtil.ValueIntInit("id", 0)
pcAct = LCase(RequestUtil.ValueInit("act", "read"))
If pcAct <> "modify" And pcAct <> "delete" Then pcAct = "read"

If pcId <= 0 Then JsAlert "잘못된 접근입니다."

pcArr = Article.Article_v(pcId)
If Not IsArray(pcArr) Then JsAlert "삭제되었거나 존재하지 않는 게시물입니다."
If AV(pcArr, AVW_BID, 0) <> bdId Then JsAlert "잘못된 접근입니다."

Dim pcTitle, pcWid, pcHasPw
pcTitle  = AV(pcArr,  AVW_TITLE, 0)
pcWid    = AVI(pcArr, AVW_WID, 0)
pcHasPw  = AVB(pcArr, AVW_HASPW, 0)

If Not pcHasPw Then JsAlert "비밀번호가 설정되지 않은 게시물입니다."

'==============================================================================
' POST : 비밀번호 검증
'==============================================================================
Dim pcErr
pcErr = ""

If RequestUtil.IsPost() Then

    Auth.CheckToken Trim(Request.Form("_csrf") & "")

    Dim inPw
    inPw = Left(Trim(Request.Form("pw") & ""), 20)

    If StringUtil.IsBlank(inPw) Then
        pcErr = "비밀번호를 입력해 주세요."
    ElseIf Not Article.ChkGuestPw(pcId, inPw) Then
        pcErr = "비밀번호가 일치하지 않습니다."
    Else
        Session("PWOK_" & bdId & "_" & pcId) = "1"

        Select Case pcAct
            Case "modify"
                JsGo U("/board/Write.asp") & "?bid=" & UE(bdId) & "&id=" & pcId & "&mode=modify"
            Case "delete"
                JsGo U("/board/Delete_p.asp") & "?bid=" & UE(bdId) & "&id=" & pcId & "&_csrf=" & UE(Auth.Token())
            Case Else
                JsGo U("/board/View.asp") & "?bid=" & UE(bdId) & "&id=" & pcId
        End Select
    End If
End If

'--- 화면 문구 ---
Dim pcLabel
Select Case pcAct
    Case "modify" : pcLabel = "수정"
    Case "delete" : pcLabel = "삭제"
    Case Else     : pcLabel = "열람"
End Select

g_pageTitle = "비밀번호 확인"
g_pageDesc  = "비회원 게시물 " & pcLabel & "을 위한 비밀번호 확인"
%>
