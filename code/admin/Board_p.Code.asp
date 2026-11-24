<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Board_p.Code.asp  -  게시판 / 분류 저장 처리
'   mode : insert | update | cate_insert | cate_delete
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim pMode, pBid
pMode = LCase(RequestUtil.ValueInit("mode", ""))
pBid  = LCase(Trim(RequestUtil.ValueInit("board_id", "")))


'==============================================================================
' 게시판 신규
'==============================================================================
If pMode = "insert" Then

    Dim iName, iDesc, iSkin, iSort
    iName = Left(Trim(Request.Form("board_name") & ""), 50)
    iDesc = Left(Trim(Request.Form("board_desc") & ""), 300)
    iSkin = LCase(Trim(Request.Form("skin_type") & ""))
    iSort = StringUtil.ToInt(Request.Form("sort_no"), 0)

    If Not StringUtil.RegTest("^[a-z][a-z0-9_]{1,19}$", pBid) Then
        JsAlert "게시판 ID 형식이 올바르지 않습니다." & vbCrLf & "영소문자로 시작하는 영문·숫자·_ 조합 2~20자."
    End If
    If StringUtil.IsBlank(iName) Then JsAlert "게시판 이름을 입력해 주세요."
    If InStr(",list,gallery,blog,guest,", "," & iSkin & ",") = 0 Then iSkin = "list"

    If IsArray(Board.Board_v(pBid)) Then JsAlert "이미 존재하는 게시판 ID 입니다."

    Board.Board_i pBid, iName, iDesc, iSkin, iSort
    JsAlertGo "게시판이 추가되었습니다. 이어서 설정을 확인하세요.", U("/admin/Board.asp") & "?bid=" & UE(pBid)


'==============================================================================
' 게시판 설정 수정
'==============================================================================
ElseIf pMode = "update" Then

    If Not IsArray(Board.Board_l(True)) Then JsAlert "게시판 정보를 찾을 수 없습니다."

    Dim uName, uDesc, uSkin
    uName = Left(Trim(Request.Form("board_name") & ""), 50)
    uDesc = Left(Trim(Request.Form("board_desc") & ""), 300)
    uSkin = LCase(Trim(Request.Form("skin_type") & ""))

    If StringUtil.IsBlank(uName) Then JsAlert "게시판 이름을 입력해 주세요."
    If InStr(",list,gallery,blog,guest,", "," & uSkin & ",") = 0 Then uSkin = "list"

    Board.Board_u pBid, uName, uDesc, uSkin, _
        StringUtil.ToInt(Request.Form("list_count"),  DEF_LIST_COUNT), _
        StringUtil.ToInt(Request.Form("page_block"),  DEF_PAGE_BLOCK), _
        StringUtil.ToInt(Request.Form("gallery_col"), 3), _
        (Trim(Request.Form("use_comment")  & "") = "1"), _
        (Trim(Request.Form("use_file")     & "") = "1"), _
        (Trim(Request.Form("use_category") & "") = "1"), _
        (Trim(Request.Form("use_secret")   & "") = "1"), _
        (Trim(Request.Form("use_reply")    & "") = "1"), _
        (Trim(Request.Form("use_editor")   & "") = "1"), _
        (Trim(Request.Form("use_notice")   & "") = "1"), _
        StringUtil.ToInt(Request.Form("file_max_count"), 3), _
        StringUtil.ToInt(Request.Form("file_max_size"), 10240), _
        Left(Trim(Request.Form("file_ext_allow") & ""), 300), _
        StringUtil.ToInt(Request.Form("auth_list"),    0), _
        StringUtil.ToInt(Request.Form("auth_read"),    0), _
        StringUtil.ToInt(Request.Form("auth_write"),   3), _
        StringUtil.ToInt(Request.Form("auth_comment"), 3), _
        StringUtil.ToInt(Request.Form("auth_admin"),   9), _
        StringUtil.ToInt(Request.Form("sort_no"), 0), _
        (Trim(Request.Form("is_use") & "") = "1")

    JsAlertGo "저장되었습니다.", U("/admin/Board.asp") & "?bid=" & UE(pBid)


'==============================================================================
' 분류 추가
'==============================================================================
ElseIf pMode = "cate_insert" Then

    Dim cName, cSort
    cName = Left(Trim(Request.Form("category_name") & ""), 50)
    cSort = StringUtil.ToInt(Request.Form("sort_no"), 0)

    If StringUtil.IsBlank(cName) Then JsAlert "분류명을 입력해 주세요."
    If Not IsArray(Board.Board_v(pBid)) Then JsAlert "게시판을 찾을 수 없습니다."

    Board.Category_i pBid, cName, cSort
    JsGo U("/admin/Board.asp") & "?bid=" & UE(pBid)


'==============================================================================
' 분류 삭제
'==============================================================================
ElseIf pMode = "cate_delete" Then

    Dim dCid
    dCid = RequestUtil.ValueIntInit("category_id", 0)
    If dCid <= 0 Then JsAlert "잘못된 접근입니다."
    If Not Board.IsValidCategory(pBid, dCid) Then JsAlert "이 게시판의 분류가 아닙니다."

    Board.Category_d dCid
    JsAlertGo "분류가 삭제되었습니다.", U("/admin/Board.asp") & "?bid=" & UE(pBid)

Else
    JsAlert "잘못된 접근입니다."
End If
%>
