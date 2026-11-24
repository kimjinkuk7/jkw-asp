<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<%
'==============================================================================
' code/member/Join_p.Code.asp  -  회원가입 처리
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."
Auth.CheckToken Trim(Request.Form("_csrf") & "")
If GLB_IS_LOGIN Then JsAlertGo "이미 로그인되어 있습니다.", U("/main/Default.asp")

Dim jId, jPw, jPw2, jName, jNick, jEmail, jMobile, jIntro, jAgree, jMailRecv
jId       = Left(LCase(Trim(Request.Form("login_id") & "")), 30)
jPw       = Left(Trim(Request.Form("user_pw")  & ""), 64)
jPw2      = Left(Trim(Request.Form("user_pw2") & ""), 64)
jName     = Left(Trim(Request.Form("user_name")& ""), 50)
jNick     = Left(Trim(Request.Form("nickname") & ""), 50)
jEmail    = Left(Trim(Request.Form("email")    & ""), 100)
jMobile   = Left(Trim(Request.Form("mobile")   & ""), 20)
jIntro    = Left(Trim(Request.Form("intro")    & ""), 500)
jAgree    = (Trim(Request.Form("agree")     & "") = "1")
jMailRecv = (Trim(Request.Form("mail_recv") & "") = "1")

'--- 입력값 보존 (오류 시 폼 복원용) ---
Session("JOIN_TMP_ID")     = jId
Session("JOIN_TMP_NAME")   = jName
Session("JOIN_TMP_NICK")   = jNick
Session("JOIN_TMP_EMAIL")  = jEmail
Session("JOIN_TMP_MOBILE") = jMobile
Session("JOIN_TMP_INTRO")  = jIntro

'--- 유효성 ---
If Not jAgree                          Then JsAlert "약관에 동의해 주세요."
If Not StringUtil.ChkLoginId(jId)      Then JsAlert "아이디 형식이 올바르지 않습니다." & vbCrLf & "영소문자로 시작하는 영문·숫자·_ 조합 4~30자."
If Not StringUtil.ChkPassword(jPw)     Then JsAlert "비밀번호는 ASCII 문자 8자 이상이며 영문과 숫자를 포함해야 합니다." & vbCrLf & "한글은 사용할 수 없습니다."
If jPw <> jPw2                         Then JsAlert "비밀번호가 일치하지 않습니다."
If StringUtil.IsBlank(jName)           Then JsAlert "이름을 입력해 주세요."
If StringUtil.IsBlank(jNick)           Then JsAlert "닉네임을 입력해 주세요."
If Not StringUtil.ChkEmail(jEmail)     Then JsAlert "이메일 형식이 올바르지 않습니다."
If Len(jMobile) > 0 Then
    If Not StringUtil.ChkMobile(jMobile) Then JsAlert "연락처 형식이 올바르지 않습니다."
End If

'--- 휴대폰 인증 확인 (화면 JS 는 우회될 수 있으므로 서버가 다시 본다) ---------
' 인증 상태는 세션에만 있다. 폼 hidden 으로 넘기면 조작할 수 있다.
Dim jPhoneOk
jPhoneOk = False
If Len(jMobile) > 0 And LCase(PHONE_AUTH_MODE) <> "off" Then
    If StringUtil.DigitsOnly(jMobile) = (Session("PHONEOK_NUM") & "") Then jPhoneOk = True
    If Not jPhoneOk Then JsAlert "휴대폰 인증을 완료해 주세요."
End If

'--- 중복 검사 ---
If Member.IsDupLoginId(jId)         Then JsAlert "이미 사용 중인 아이디입니다."
If Member.IsDupNickname(jNick, 0)   Then JsAlert "이미 사용 중인 닉네임입니다."
If Member.IsDupEmail(jEmail, 0)     Then JsAlert "이미 가입된 이메일입니다."

'--- 예약어 아이디 차단 ---
If InStr(",admin,administrator,root,master,system,manager,test,guest,null,", "," & jId & ",") > 0 Then
    If jId <> "test" Then JsAlert "사용할 수 없는 아이디입니다."
End If

'--- 등록 ---
Dim newId
newId = Member.Member_i(jId, jPw, jName, jNick, jEmail, jMobile, jIntro, jMailRecv)
If newId <= 0 Then JsAlert "가입 처리에 실패했습니다. 잠시 후 다시 시도해 주세요."

'--- 휴대폰 인증 기록 (어떤 방식으로 인증했는지 함께 남긴다) ---
If jPhoneOk Then
    Member.MobileVerify_u newId, jMobile, LCase(Session("PHONEOK_TYPE") & "")
End If
Session("PHONEOK_NUM")  = ""
Session("PHONEOK_DT")   = ""
Session("PHONEOK_TYPE") = ""

'--- 임시 세션 정리 후 자동 로그인 ---
Session("JOIN_TMP_ID")     = ""
Session("JOIN_TMP_NAME")   = ""
Session("JOIN_TMP_NICK")   = ""
Session("JOIN_TMP_EMAIL")  = ""
Session("JOIN_TMP_MOBILE") = ""
Session("JOIN_TMP_INTRO")  = ""

Auth.SetSession newId, jId, LV_MEMBER, jNick, jName

JsAlertGo "가입이 완료되었습니다. 환영합니다!", U("/main/Default.asp")
%>
