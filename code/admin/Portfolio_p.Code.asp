<!--#include file="Admin.Code.asp"-->
<!--#include file="../../lib/FileUpload.Class.asp"-->
<%
'==============================================================================
' code/admin/Portfolio_p.Code.asp  -  포트폴리오 저장 / 삭제
'   mode : save | delete     (mode / id 는 QueryString 으로 받는다)
'==============================================================================

Dim ppMode, ppId, ppIsMp
ppMode = LCase(Trim(Request.QueryString("mode") & ""))
ppId   = StringUtil.ToInt(Request.QueryString("id"), 0)

'--- 삭제 ---------------------------------------------------------------------
If ppMode = "delete" Then
    Auth.CheckToken Trim(Request.QueryString("_csrf") & "")
    If ppId <= 0 Then JsAlert "잘못된 접근입니다."
    Portfolio.Portfolio_d ppId
    JsAlertGo "삭제되었습니다.", U("/admin/Portfolio.asp")
End If

If ppMode <> "save" Then JsAlert "잘못된 접근입니다."
If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."

'--- 본문 파싱 (파일 업로드 대응) ---
ppIsMp = RequestUtil.IsMultipart()
If ppIsMp Then
    If Not Upload.Load() Then JsAlert Upload.ErrMsg
End If

Auth.CheckToken PPV("_csrf")

'--- 값 수집 ---
Dim qTitle, qSubtitle, qSummary, qContent, qCategory, qClient, qRole, qTeam
Dim qStart, qEnd, qOngoing, qSiteUrl, qRepoUrl, qThumb, qIsMain, qSortNo, qIsUse
Dim qSkillIds, qDelThumb

qTitle    = Left(PPV("title"), 200)
qSubtitle = Left(PPV("subtitle"), 200)
qSummary  = Left(PPV("summary"), 500)
qContent  = PPV("content")
qCategory = Left(PPV("category"), 50)
qClient   = Left(PPV("client_name"), 100)
qRole     = Left(PPV("job_role"), 200)
qTeam     = Left(PPV("team_size"), 50)
qStart    = PPV("period_start")
qEnd      = PPV("period_end")
qOngoing  = (PPV("is_ongoing") = "1")
qSiteUrl  = Left(PPV("site_url"), 300)
qRepoUrl  = Left(PPV("repo_url"), 300)
qIsMain   = (PPV("is_main") = "1")
qIsUse    = (PPV("is_use") = "1")
qSortNo   = StringUtil.ToInt(PPV("sort_no"), 0)
qSkillIds = PPV("skill_ids")
qThumb    = Left(PPV("thumb_cur"), 300)
qDelThumb = (PPV("del_thumb") = "1")

If StringUtil.IsBlank(qTitle) Then JsAlert "제목을 입력해 주세요."
If Not IsDate(qStart) Then qStart = Null
If Not IsDate(qEnd)   Then qEnd   = Null
qContent = StringUtil.SafeHtml(qContent)

'--- 대표 이미지 ---
If qDelThumb Then qThumb = ""

If ppIsMp Then
    If Upload.HasFile("thumb_file_up") Then
        If Not Upload.ChkExt("thumb_file_up", "jpg,jpeg,gif,png,webp") Then JsAlert Upload.ErrMsg
        If Not Upload.ChkSize("thumb_file_up", 5120) Then JsAlert Upload.ErrMsg

        Dim tUrlPath, tPhysPath, tExt, tSaveNm
        tUrlPath  = UPLOAD_ROOT & "/portfolio"
        tPhysPath = Server.MapPath(tUrlPath)
        tExt      = Upload.FileExt("thumb_file_up")
        tSaveNm   = StringUtil.NewKey() & "." & tExt

        If Upload.SaveAs("thumb_file_up", tPhysPath, tSaveNm) Then
            qThumb = tUrlPath & "/" & tSaveNm
        Else
            JsAlert Upload.ErrMsg
        End If
    End If
End If

'--- 저장 ---
Dim savedId
If ppId > 0 Then
    Portfolio.Portfolio_u ppId, qTitle, qSubtitle, qSummary, qContent, qCategory, qClient, qRole, qTeam, _
                          qStart, qEnd, qOngoing, qSiteUrl, qRepoUrl, qThumb, qIsMain, qSortNo, qIsUse
    savedId = ppId
Else
    savedId = Portfolio.Portfolio_i(qTitle, qSubtitle, qSummary, qContent, qCategory, qClient, qRole, qTeam, _
                                    qStart, qEnd, qOngoing, qSiteUrl, qRepoUrl, qThumb, qIsMain, qSortNo, qIsUse)
    If savedId <= 0 Then JsAlert "등록에 실패했습니다."
End If

Portfolio.Skill_p savedId, qSkillIds

Upload.Dispose

JsAlertGo "저장되었습니다.", U("/admin/PortfolioForm.asp") & "?id=" & savedId


'##############################################################################
Function PPV(ByVal pName)
    If ppIsMp Then
        PPV = Upload.Field(pName)
    Else
        PPV = Trim(Request.Form(pName) & "")
    End If
End Function
%>
