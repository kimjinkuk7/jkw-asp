<!--#include file="Board.Code.asp"-->
<%
'==============================================================================
' code/board/Comment_p.Code.asp  -  댓글 등록 / 삭제 처리
'==============================================================================

Auth.CheckToken RequestUtil.ValueInit("_csrf", "")

Dim cMode, artId, artArr, backTo
cMode = LCase(RequestUtil.ValueInit("mode", "write"))
artId = RequestUtil.ValueIntInit("id", 0)

If artId <= 0 Then JsAlert "잘못된 접근입니다."
If Not bdUseComment Then JsAlert "이 게시판은 댓글을 사용하지 않습니다."

artArr = Article.Article_v(artId)
If Not IsArray(artArr) Then JsAlert "삭제되었거나 존재하지 않는 게시물입니다."
If AV(artArr, AVW_BID, 0) <> bdId Then JsAlert "잘못된 접근입니다."

backTo = U("/board/View.asp") & "?bid=" & UE(bdId) & "&id=" & artId


'==============================================================================
' 등록
'==============================================================================
If cMode = "write" Then

    If Not bdCanComment Then JsAlert "댓글 작성 권한이 없습니다."

    Dim cContent, cParentId, cName, cPw, myId2, cNameFinal, cPwHash
    cContent  = Left(Trim(Request.Form("content") & ""), 1000)
    cParentId = StringUtil.ToInt(Request.Form("parent_id"), 0)
    cName     = Left(Trim(Request.Form("writer_name") & ""), 20)
    cPw       = Left(Trim(Request.Form("writer_pw") & ""), 20)

    If StringUtil.IsBlank(cContent) Then JsAlert "댓글 내용을 입력해 주세요."

    '--- 대댓글 대상이 이 글의 댓글인지 확인 ---
    If cParentId > 0 Then
        Dim pArr
        pArr = Comment.Comment_v(cParentId)
        If Not IsArray(pArr) Then JsAlert "원 댓글을 찾을 수 없습니다."
        If AVI(pArr, 0, 0) <> artId Then JsAlert "잘못된 접근입니다."
    End If

    myId2 = Auth.MyId()
    If myId2 > 0 Then
        cNameFinal = Auth.DisplayName()
        cPwHash    = Null
    Else
        If StringUtil.IsBlank(cName) Then JsAlert "이름을 입력해 주세요."
        If Len(cPw) < 4 Then JsAlert "비밀번호를 4자 이상 입력해 주세요."
        cNameFinal = cName
        cPwHash    = Crypt.HashGuestPw(cPw)
    End If

    Comment.Comment_i artId, bdId, cParentId, cContent, myId2, cNameFinal, cPwHash, False

    JsGo backTo & "#comment"


'==============================================================================
' 삭제
'==============================================================================
ElseIf cMode = "delete" Then

    Dim cmtId, cmtArr, cWid, cHasPw, okDel, inPw
    cmtId = StringUtil.ToInt(RequestUtil.ValueInit("cid", "0"), 0)
    If cmtId <= 0 Then JsAlert "잘못된 접근입니다."

    cmtArr = Comment.Comment_v(cmtId)
    If Not IsArray(cmtArr) Then JsAlert "이미 삭제된 댓글입니다."
    If AVI(cmtArr, 0, 0) <> artId Then JsAlert "잘못된 접근입니다."

    cWid   = AVI(cmtArr, 1, 0)
    cHasPw = AVB(cmtArr, 2, 0)

    okDel = False
    If bdCanAdmin Then okDel = True
    If Auth.MyId() > 0 And Auth.MyId() = cWid Then okDel = True

    If Not okDel And cWid = 0 And cHasPw Then
        inPw = RequestUtil.ValueInit("pw", "")
        If StringUtil.IsBlank(inPw) Then JsAlert "비밀번호를 입력해 주세요."
        If Comment.ChkGuestPw(cmtId, inPw) Then okDel = True Else JsAlert "비밀번호가 일치하지 않습니다."
    End If

    If Not okDel Then JsAlert "삭제 권한이 없습니다."

    Comment.Comment_d cmtId, artId
    JsGo backTo & "#comment"

Else
    JsAlert "잘못된 접근입니다."
End If
%>
