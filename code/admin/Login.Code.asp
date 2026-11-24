<!--#include file="../Common.Code.asp"-->
<%
'==============================================================================
' code/admin/Login.Code.asp  -  관리자 로그인
'------------------------------------------------------------------------------
' 일반 로그인과 동일한 인증을 쓰되, 최고관리자 등급만 통과시킨다.
'==============================================================================
g_pageTitle = "관리자 로그인"
g_bodyClass = "admin page-member"

If GLB_MY_LEVEL >= LV_ADMIN Then JsGo U("/admin/Default.asp")

Dim adRurl, adErr
adRurl = RequestUtil.ValueInit("rurl", "")
If Len(adRurl) = 0 Or Left(adRurl, 1) <> "/" Or Left(adRurl, 2) = "//" Then
    adRurl = U("/admin/Default.asp")
End If
adErr = ""

If RequestUtil.IsPost() Then
    Auth.CheckToken Trim(Request.Form("_csrf") & "")

    Dim aId, aPw, rc
    aId = Left(LCase(Trim(Request.Form("login_id") & "")), 30)
    aPw = Left(Trim(Request.Form("user_pw") & ""), 64)

    If StringUtil.IsBlank(aId) Or StringUtil.IsBlank(aPw) Then
        adErr = "아이디와 비밀번호를 입력해 주세요."
    Else
        rc = Auth.Login(aId, aPw)
        If rc = 1 Then
            If Auth.MyLevel() >= LV_ADMIN Then
                JsGo adRurl
            Else
                Auth.Logout
                adErr = "관리자 권한이 없는 계정입니다."
            End If
        ElseIf rc = -1 Then
            adErr = "정지되었거나 탈퇴한 계정입니다."
        ElseIf rc = -2 Then
            adErr = "로그인 실패가 누적되어 " & LOGIN_LOCK_MIN & "분간 제한됩니다."
        Else
            adErr = "아이디 또는 비밀번호가 일치하지 않습니다."
        End If
    End If
End If
%>
