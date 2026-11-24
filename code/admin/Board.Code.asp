<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Board.Code.asp  -  게시판 관리 (목록 + 설정 편집)
'------------------------------------------------------------------------------
' bid 가 있으면 편집 화면, 없으면 목록 화면.
'==============================================================================
g_adminNav  = "board"
g_pageTitle = "게시판 관리"

Dim abBid, abArr, abIsEdit
abBid    = LCase(Trim(RequestUtil.ValueInit("bid", "")))
abIsEdit = False

Dim abListArr, abListCnt
abListArr = Board.Board_l(True)
abListCnt = RowCnt(abListArr)

Dim abCateArr, abCateCnt
abCateCnt = 0

If Len(abBid) > 0 Then
    ' 미사용 게시판도 편집할 수 있어야 하므로 목록에서 찾는다
    Dim k
    For k = 0 To abListCnt - 1
        If AV(abListArr, BM_ID, k) = abBid Then
            abIsEdit = True
            Exit For
        End If
    Next
    If Not abIsEdit Then JsAlertGo "존재하지 않는 게시판입니다.", U("/admin/Board.asp")

    abArr     = Board.Board_v(abBid)
    abCateArr = Board.Category_l(abBid)
    abCateCnt = RowCnt(abCateArr)
End If

'--- 편집 대상 인덱스 (Board_l 결과에서) ---
Dim abIdx
abIdx = -1
If abIsEdit Then
    For k = 0 To abListCnt - 1
        If AV(abListArr, BM_ID, k) = abBid Then abIdx = k
    Next
End If

'--- 스킨 선택지 ---
Dim abSkins
abSkins = Array("list", "gallery", "blog", "guest")

'--- 권한 선택지 ---
Dim abLevels
abLevels = Array(0, 3, 5, 9)

Function LevelName(ByVal lv)
    Select Case CLng(lv)
        Case 0 : LevelName = "0 · 누구나"
        Case 3 : LevelName = "3 · 일반회원"
        Case 5 : LevelName = "5 · 운영자"
        Case 9 : LevelName = "9 · 관리자"
        Case Else : LevelName = CStr(lv)
    End Select
End Function
%>
