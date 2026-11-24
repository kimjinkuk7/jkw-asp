<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/contact/Contact.Class.asp"-->
<%
'==============================================================================
' code/contact/Send_p.Code.asp  -  문의 접수 처리
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."

Auth.CheckToken Trim(Request.Form("_csrf") & "")

Dim sType, sName, sCompany, sEmail, sMobile, sSubject, sContent, sAgree
sType    = Left(Trim(Request.Form("contact_type") & ""), 30)
sName    = Left(Trim(Request.Form("writer_name")  & ""), 50)
sCompany = Left(Trim(Request.Form("company")      & ""), 100)
sEmail   = Left(Trim(Request.Form("email")        & ""), 100)
sMobile  = Left(Trim(Request.Form("mobile")       & ""), 20)
sSubject = Left(Trim(Request.Form("subject")      & ""), 200)
sContent = Left(Trim(Request.Form("content")      & ""), 4000)
sAgree   = (Trim(Request.Form("is_agree") & "") = "1")

'--- 유효성 ---
If Not sAgree                        Then JsAlert "개인정보 수집 및 이용에 동의해 주세요."
If StringUtil.IsBlank(sName)         Then JsAlert "이름을 입력해 주세요."
If StringUtil.IsBlank(sEmail)        Then JsAlert "이메일을 입력해 주세요."
If Not StringUtil.ChkEmail(sEmail)   Then JsAlert "이메일 형식이 올바르지 않습니다."
If StringUtil.IsBlank(sSubject)      Then JsAlert "제목을 입력해 주세요."
If Len(sContent) < 10                Then JsAlert "문의 내용을 10자 이상 입력해 주세요."

'--- 도배 방지 : 같은 IP 로 10분 내 3건 초과 차단 ---
If Contact.RecentByIp_c(10) >= 3 Then
    JsAlert "짧은 시간에 너무 많은 문의를 보내셨습니다. 잠시 후 다시 시도해 주세요."
End If

'--- 스팸 필터 : 본문에 링크가 과도하게 많으면 차단 ---
Dim linkCnt
linkCnt = UBound(Split(LCase(sContent) & " ", "http"))
If linkCnt > 5 Then JsAlert "본문에 링크가 너무 많습니다. 내용을 확인해 주세요."

'--- 저장 ---
Dim newId
newId = Contact.Contact_i(sType, sName, sCompany, sEmail, sMobile, sSubject, sContent, sAgree)
If newId <= 0 Then JsAlert "접수에 실패했습니다. 잠시 후 다시 시도해 주세요."

'--- 메일 발송 (tb_config.mail_use = 1 일 때만) -------------------------------
'    CDO.Message 는 서버에 SMTP 설정이 있어야 동작하므로 실패해도 무시한다.
If Config.Value("mail_use") = "1" Then
    On Error Resume Next
    Dim cdo, cfg
    Set cdo = Server.CreateObject("CDO.Message")
    Set cfg = Server.CreateObject("CDO.Configuration")

    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")      = 2
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")     = Config.Value("mail_smtp")
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Config.ValueInt("mail_port", 587)
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")   = Config.Value("mail_user")
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")   = Config.Value("mail_pass")
    cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")     = True
    cfg.Fields.Update

    Set cdo.Configuration = cfg
    cdo.From     = Config.Value("mail_user")
    cdo.To       = GLB_OWNER_EMAIL
    cdo.ReplyTo  = sEmail
    cdo.Subject  = "[" & GLB_SITE_TITLE & " 문의] " & sSubject
    cdo.BodyPart.Charset = "utf-8"
    cdo.HTMLBody = "<p><b>유형</b> : " & HE(sType)    & "</p>" & _
                   "<p><b>이름</b> : " & HE(sName)    & " / " & HE(sCompany) & "</p>" & _
                   "<p><b>이메일</b> : " & HE(sEmail) & "</p>" & _
                   "<p><b>연락처</b> : " & HE(sMobile) & "</p>" & _
                   "<hr><p>" & Replace(HE(sContent), vbCrLf, "<br>") & "</p>"
    cdo.Send

    Set cdo = Nothing
    Set cfg = Nothing
    Err.Clear
    On Error GoTo 0
End If

JsAlertGo "문의가 정상적으로 접수되었습니다. 빠른 시일 내에 회신드리겠습니다.", U("/main/Default.asp")
%>
