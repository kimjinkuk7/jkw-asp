<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/PhoneAuth.Class.asp"-->
<%
'==============================================================================
' code/member/PhoneAuth_p.Code.asp  -  휴대폰 인증 (AJAX, 순수 텍스트 응답)
'------------------------------------------------------------------------------
' act=send    인증번호 발급
'   OK|SENT|<유효초>|<인증번호 또는 빈값>
'                                    ← 인증번호는 mock 일 때만 돌려준다
'   FAIL|<메시지>
'
' act=verify  인증번호 확인
'   OK|VERIFIED|<정규화번호>
'   FAIL|<메시지>|<남은시도>
'
' ★ 실제 SMS/Firebase 로 바꿀 때 손댈 곳은 이 파일의 SendSms 한 군데다.
'   발급·저장·대조·만료·시도제한 로직은 그대로 쓴다.
'
' 인증 결과는 세션에 남긴다 (폼 hidden 은 조작 가능하므로).
'   Session("PHONEOK_NUM") / ("PHONEOK_DT") / ("PHONEOK_TYPE")
'==============================================================================

Const PA_VALID_MIN   = 3        ' 인증번호 유효시간(분)
Const PA_MAX_TRY     = 5        ' 인증번호당 최대 시도
Const PA_MAX_SEND    = 5        ' 10분당 최대 발송
Const PA_SEND_WINDOW = 10

Response.Clear
Response.ContentType = "text/plain"
Response.CharSet     = "utf-8"
Response.Expires     = -1

Sub Out(ByVal s)
    Response.Write s
    Response.End
End Sub

'<Desc>실제 발송 지점. mock 은 보내지 않고 True 만 돌려준다.</Desc>
Function SendSms(ByVal pMobile, ByVal pCode, ByRef outErr)
    outErr = ""
    Select Case LCase(PHONE_AUTH_MODE)
        Case "mock"
            ' 시연용 : 문자를 보내지 않는다. 화면이 인증번호를 직접 보여준다.
            SendSms = True
        Case "sms"
            ' TODO : 솔라피 / 알리고 / NCP SENS REST 호출 (lib/Http, lib/Json 사용)
            outErr = "SMS 발송이 아직 연결되지 않았습니다."
            SendSms = False
        Case "firebase"
            ' Firebase 는 브라우저가 직접 SMS 를 받고 idToken 을 준다.
            ' 이 엔드포인트를 쓰지 않고 별도 검증 경로가 필요하다.
            outErr = "Firebase 인증은 아직 준비 중입니다."
            SendSms = False
        Case Else
            outErr = "휴대폰 인증이 비활성화되어 있습니다."
            SendSms = False
    End Select
End Function

'--- 공통 검증 --------------------------------------------------------------
If Not RequestUtil.IsPost() Then Out "FAIL|잘못된 접근입니다."
If Len(Auth.Token()) = 0 Then Out "FAIL|세션이 만료되었습니다. 새로고침 후 다시 시도해 주세요."
If CStr(Trim(Request.Form("_csrf") & "")) <> Auth.Token() Then Out "FAIL|잘못된 접근입니다."

Dim paAct, paMobile, paNum
paAct    = LCase(Trim(Request.Form("act") & ""))
paMobile = Left(Trim(Request.Form("mobile") & ""), 20)

If StringUtil.IsBlank(paMobile)       Then Out "FAIL|휴대폰 번호를 입력해 주세요."
If Not StringUtil.ChkMobile(paMobile) Then Out "FAIL|휴대폰 번호 형식이 올바르지 않습니다."
paNum = StringUtil.DigitsOnly(paMobile)

'==============================================================================
' 1) 인증번호 발급
'==============================================================================
If paAct = "send" Then

    Dim paCode, paErr, paId

    If PhoneAuth.RecentCount_c(paNum, PA_SEND_WINDOW) >= PA_MAX_SEND Then
        Out "FAIL|인증 요청이 너무 많습니다. " & PA_SEND_WINDOW & "분 후 다시 시도해 주세요."
    End If

    paCode = PhoneAuth.NewCode()

    If Not SendSms(paNum, paCode, paErr) Then Out "FAIL|" & paErr

    paId = PhoneAuth.Send_i(paNum, paCode, PHONE_AUTH_MODE, PA_VALID_MIN)
    If paId <= 0 Then Out "FAIL|인증번호 발급에 실패했습니다."

    ' 번호를 바꿔 다시 받는 경우를 대비해 기존 인증 상태를 지운다
    Session("PHONEOK_NUM")  = ""
    Session("PHONEOK_TYPE") = ""

    PhoneAuth.Purge_p

    ' mock 일 때만 인증번호를 화면으로 돌려준다 (회색 안내로 표시된다)
    If LCase(PHONE_AUTH_MODE) = "mock" Then
        Out "OK|SENT|" & (PA_VALID_MIN * 60) & "|" & paCode
    Else
        Out "OK|SENT|" & (PA_VALID_MIN * 60) & "|"
    End If
End If

'==============================================================================
' 2) 인증번호 확인
'==============================================================================
If paAct = "verify" Then

    Dim paIn, paArr, paId2, paDbCode, paExpire, paTry, paLeft

    paIn = Left(Trim(Request.Form("code") & ""), 10)
    If StringUtil.IsBlank(paIn) Then Out "FAIL|인증번호를 입력해 주세요.|"

    paArr = PhoneAuth.Latest_v(paNum)
    If Not IsArray(paArr) Then Out "FAIL|먼저 인증번호를 받아 주세요.|"

    paId2    = AVI(paArr, PA_ID, 0)
    paDbCode = AV(paArr,  PA_CODE, 0)
    paExpire = AV(paArr,  PA_EXPIRE, 0)
    paTry    = AVI(paArr, PA_TRY, 0)

    If paTry >= PA_MAX_TRY Then
        Out "FAIL|입력 횟수를 초과했습니다. 인증번호를 다시 받아 주세요.|0"
    End If

    If IsDate(paExpire) Then
        If CDate(paExpire) < Now() Then
            Out "FAIL|인증번호가 만료되었습니다. 다시 받아 주세요.|"
        End If
    End If

    If CStr(paIn) <> CStr(paDbCode) Then
        PhoneAuth.TryUp_u paId2
        paLeft = PA_MAX_TRY - (paTry + 1)
        If paLeft < 0 Then paLeft = 0
        Out "FAIL|인증번호가 일치하지 않습니다.|" & paLeft
    End If

    If Not PhoneAuth.Verified_u(paId2) Then
        Out "FAIL|이미 사용된 인증번호입니다.|"
    End If

    Session("PHONEOK_NUM")  = paNum
    Session("PHONEOK_DT")   = Now()
    Session("PHONEOK_TYPE") = LCase(PHONE_AUTH_MODE)

    Out "OK|VERIFIED|" & paNum
End If

Out "FAIL|알 수 없는 요청입니다."
%>
