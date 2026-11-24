<!--#include file="Board.Code.asp"-->
<!--#include file="../../lib/FileUpload.Class.asp"-->
<%
'==============================================================================
' code/board/Write_p.Code.asp  -  글 등록 / 수정 / 답변 처리
'------------------------------------------------------------------------------
' ★ multipart 요청이면 Upload.Load() 로 본문을 먼저 파싱해야 한다.
'   그 이전에 Request.Form 을 건드리면 BinaryRead 가 실패한다.
'   (Board.Code.asp 는 QueryString 만 읽도록 되어 있다)
'==============================================================================

'--- POST 강제 ---
If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."

'--- 본문 파싱 ---
Dim isMp
isMp = RequestUtil.IsMultipart()
If isMp Then
    If Not Upload.Load() Then JsAlert Upload.ErrMsg
End If

'--- 파라미터 --------------------------------------------------------------
Dim wMode, artId, pTitle, pContent, pCateId, pTags, pLink
Dim pWriterName, pWriterPw, pIsNotice, pIsSecret, pDelFile
Dim rPage, rSf, rSw

Auth.CheckToken PV("_csrf")

wMode = LCase(PV("mode"))
If wMode <> "modify" And wMode <> "reply" Then wMode = "write"

artId       = StringUtil.ToInt(PV("id"), 0)
pTitle      = Left(PV("title"), 200)
pContent    = PV("content")
pCateId     = StringUtil.ToInt(PV("category_id"), 0)
pTags       = Left(PV("tags"), 300)
pLink       = Left(PV("link_url"), 300)
pWriterName = Left(PV("writer_name"), 20)
pWriterPw   = Left(PV("writer_pw"), 20)
pIsNotice   = (PV("is_notice") = "1")
pIsSecret   = (PV("is_secret") = "1")
pDelFile    = PV("del_file")

rPage = StringUtil.ToInt(PV("page"), 1)
rSf   = Left(PV("sf"), 20)
rSw   = Left(PV("sw"), 50)

'--- 권한 ------------------------------------------------------------------
If Not bdCanWrite Then JsAlert "글쓰기 권한이 없습니다."
If pIsNotice And Not bdCanAdmin Then pIsNotice = False
If Not bdUseSecret Then pIsSecret = False
If wMode = "reply" And Not bdUseReply Then JsAlert "이 게시판은 답변 기능을 사용하지 않습니다."

'--- 방명록은 제목 고정 ---
If bdSkin = "guest" Then pTitle = "방명록"

'--- 유효성 ----------------------------------------------------------------
If StringUtil.IsBlank(pTitle)   Then JsAlert "제목을 입력해 주세요."
If StringUtil.IsBlank(pContent) Then JsAlert "내용을 입력해 주세요."

If pCateId > 0 Then
    If Not Board.IsValidCategory(bdId, pCateId) Then pCateId = 0
End If

'--- 비회원 처리 ------------------------------------------------------------
Dim myId, writerNm, writerPwHash, writerEmail
myId = Auth.MyId()

If myId > 0 Then
    writerNm     = Auth.DisplayName()
    writerPwHash = Null
    writerEmail  = Null
Else
    If StringUtil.IsBlank(pWriterName) Then JsAlert "작성자 이름을 입력해 주세요."
    writerNm    = pWriterName
    writerEmail = Null

    If wMode = "modify" Then
        ' 비워두면 기존 비밀번호 유지
        If StringUtil.IsBlank(pWriterPw) Then
            writerPwHash = Null
        Else
            writerPwHash = Crypt.HashGuestPw(pWriterPw)
        End If
    Else
        If Len(pWriterPw) < 4 Then JsAlert "비밀번호를 4자 이상 입력해 주세요."
        writerPwHash = Crypt.HashGuestPw(pWriterPw)
    End If
End If

'--- 본문 정화 (스크립트성 태그 제거) ---
pContent = StringUtil.SafeHtml(pContent)

'==============================================================================
' 저장
'==============================================================================
Dim newId, oldArr

Select Case wMode

    Case "modify"
        If artId <= 0 Then JsAlert "잘못된 접근입니다."
        oldArr = Article.Article_v(artId)
        If Not IsArray(oldArr) Then JsAlert "삭제되었거나 존재하지 않는 게시물입니다."
        If AV(oldArr, AVW_BID, 0) <> bdId Then JsAlert "잘못된 접근입니다."

        '--- 수정 권한 재확인 (폼 우회 방지) ---
        Dim oWid, oHasPw, okMod
        oWid   = AVI(oldArr, AVW_WID, 0)
        oHasPw = AVB(oldArr, AVW_HASPW, 0)
        okMod  = False
        If bdCanAdmin Then okMod = True
        If myId > 0 And myId = oWid Then okMod = True
        If oWid = 0 And oHasPw And (Session("PWOK_" & bdId & "_" & artId) & "" = "1") Then okMod = True
        If Not okMod Then JsAlert "수정 권한이 없습니다."

        Article.Article_u artId, pCateId, pTitle, pContent, pTags, _
                          AV(oldArr, AVW_THUMB, 0), pLink, pIsNotice, pIsSecret

        ' 비회원 비밀번호를 새로 입력했으면 갱신
        If myId = 0 And Not IsNull(writerPwHash) Then UpdateGuestPw artId, writerPwHash

        newId = artId

    Case "reply"
        If artId <= 0 Then JsAlert "잘못된 접근입니다."
        newId = Article.Reply_i(bdId, artId, pTitle, pContent, myId, writerNm, writerPwHash, writerEmail)
        If newId <= 0 Then JsAlert "답변 등록에 실패했습니다."

    Case Else
        newId = Article.Article_i(bdId, pCateId, pTitle, pContent, myId, writerNm, writerPwHash, _
                                  writerEmail, pTags, "", pLink, pIsNotice, pIsSecret)
        If newId <= 0 Then JsAlert "등록에 실패했습니다."

End Select

'==============================================================================
' 첨부파일
'==============================================================================
If bdUseFile And bdFileCnt > 0 Then

    '--- 삭제 체크된 기존 첨부 제거 ---
    If wMode = "modify" And Len(pDelFile) > 0 Then
        Dim delArr, di, delId
        delArr = Split(pDelFile, ",")
        For di = 0 To UBound(delArr)
            delId = StringUtil.ToInt(Trim(delArr(di)), 0)
            If delId > 0 Then BoardFile.File_d delId
        Next
    End If

    '--- 신규 업로드 ---
    If isMp Then
        '--- physPath 라는 이름은 쓸 수 없다. 전역 함수 PhysPath() 와 충돌한다 ---
        Dim urlPath, physFolder, curCnt, ui2, fld, ext, saveNm
        urlPath    = BoardUploadPath()
        physFolder = Server.MapPath(urlPath)
        curCnt     = BoardFile.File_c(newId)

        For ui2 = 1 To bdFileCnt
            fld = "upfile" & ui2
            If Upload.HasFile(fld) Then

                If curCnt >= bdFileCnt Then
                    Exit For
                End If

                If Not Upload.ChkExt(fld, bdFileExt) Then JsAlert Upload.ErrMsg
                If Not Upload.ChkSize(fld, bdFileSize) Then JsAlert Upload.ErrMsg

                ext    = Upload.FileExt(fld)
                saveNm = StringUtil.NewKey() & "." & ext

                If Upload.SaveAs(fld, physFolder, saveNm) Then
                    BoardFile.File_i newId, bdId, Upload.FileName(fld), saveNm, urlPath, _
                                     Upload.FileSize(fld), ext, curCnt
                    curCnt = curCnt + 1
                Else
                    JsAlert Upload.ErrMsg
                End If
            End If
        Next
    End If

    '--- 갤러리 / 블로그 : 첫 이미지를 대표 썸네일로 ---
    If bdSkin = "gallery" Or bdSkin = "blog" Then
        Article.Thumb_u newId, BoardFile.FirstImage_v(newId)
    End If
End If

Upload.Dispose

'==============================================================================
' 이동
'==============================================================================
Dim goUrl
If bdSkin = "guest" Then
    goUrl = U("/board/List.asp") & "?bid=" & UE(bdId)
Else
    goUrl = U("/board/View.asp") & "?bid=" & UE(bdId) & "&id=" & newId & _
            "&page=" & rPage & "&sf=" & UE(rSf) & "&sw=" & UE(rSw)
End If

JsGo goUrl


'##############################################################################
' 내부 함수
'##############################################################################

'<Desc>multipart 여부에 상관없이 POST 값을 읽는다</Desc>
Function PV(ByVal pName)
    If isMp Then
        PV = Upload.Field(pName)
    Else
        PV = Trim(Request.Form(pName) & "")
    End If
End Function

'<Desc>비회원 글 비밀번호 갱신</Desc>
Sub UpdateGuestPw(ByVal pArticleId, ByVal pHash)
    Dim h
    Set h = New CDBHelper
    h.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    With h
        .Sql = "UPDATE tb_board_article SET writer_pw = ? WHERE article_id = ?"
        .SetParamVarchar(pHash)
        .SetParamBigInt(pArticleId)
    End With
    h.Execute()
    Set h = Nothing
End Sub
%>
