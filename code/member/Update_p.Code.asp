<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<%
'==============================================================================
' code/member/Update_p.Code.asp  -  내 정보 수정 / 비밀번호 변경 / 탈퇴
'------------------------------------------------------------------------------
' mode : info | password | withdraw
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."
Auth.CheckLevel LV_MEMBER
Auth.CheckToken Trim(Request.Form("_csrf") & "")

Dim uMode
uMode = LCase(Trim(Request.Form("mode") & ""))


'==============================================================================
' 내 정보 수정
'==============================================================================
If uMode = "info" Then

    Dim uName, uNick, uEmail, uMobile, uHome, uIntro, uMailRecv
    uName     = Left(Trim(Request.Form("user_name") & ""), 50)
    uNick     = Left(Trim(Request.Form("nickname")  & ""), 50)
    uEmail    = Left(Trim(Request.Form("email")     & ""), 100)
    uMobile   = Left(Trim(Request.Form("mobile")    & ""), 20)
    uHome     = Left(Trim(Request.Form("homepage")  & ""), 200)
    uIntro    = Left(Trim(Request.Form("intro")     & ""), 500)
    uMailRecv = (Trim(Request.Form("mail_recv") & "") = "1")

    If StringUtil.IsBlank(uName)       Then JsAlert "이름을 입력해 주세요."
    If StringUtil.IsBlank(uNick)       Then JsAlert "닉네임을 입력해 주세요."
    If Not StringUtil.ChkEmail(uEmail) Then JsAlert "이메일 형식이 올바르지 않습니다."
    If Len(uMobile) > 0 Then
        If Not StringUtil.ChkMobile(uMobile) Then JsAlert "연락처 형식이 올바르지 않습니다."
    End If

    If Member.IsDupNickname(uNick, GLB_MY_ID) Then JsAlert "이미 사용 중인 닉네임입니다."
    If Member.IsDupEmail(uEmail, GLB_MY_ID)   Then JsAlert "이미 사용 중인 이메일입니다."

    Member.Member_u GLB_MY_ID, uName, uNick, uEmail, uMobile, uHome, uIntro, uMailRecv

    '--- 세션의 표시 이름 갱신 ---
    Session(SKEY_NICK) = uNick
    Session(SKEY_NAME) = uName

    JsAlertGo "정보가 수정되었습니다.", U("/member/Mypage.asp") & "?tab=info"


'==============================================================================
' 비밀번호 변경
'==============================================================================
ElseIf uMode = "password" Then

    Dim curPw, newPw, newPw2
    curPw  = Left(Trim(Request.Form("cur_pw")  & ""), 64)
    newPw  = Left(Trim(Request.Form("new_pw")  & ""), 64)
    newPw2 = Left(Trim(Request.Form("new_pw2") & ""), 64)

    ' 소셜 전용 계정은 비밀번호가 없다. 현재 비밀번호를 물을 수 없으므로 변경도 막는다.
    If Not Member.HasPwLogin(GLB_MY_ID) Then
        JsAlert "소셜 로그인 전용 계정은 비밀번호를 사용하지 않습니다."
    End If

    If StringUtil.IsBlank(curPw)                  Then JsAlert "현재 비밀번호를 입력해 주세요."
    If Not Member.ChkPassword(GLB_MY_ID, curPw)   Then JsAlert "현재 비밀번호가 일치하지 않습니다."
    If Not StringUtil.ChkPassword(newPw)          Then JsAlert "새 비밀번호는 ASCII 문자 8자 이상이며 영문과 숫자를 포함해야 합니다."
    If newPw <> newPw2                            Then JsAlert "새 비밀번호가 일치하지 않습니다."
    If newPw = curPw                              Then JsAlert "현재 비밀번호와 다른 비밀번호를 입력해 주세요."

    Member.Password_u GLB_MY_ID, newPw

    JsAlertGo "비밀번호가 변경되었습니다. 다시 로그인해 주세요.", U("/member/Logout.asp")


'==============================================================================
' 회원 탈퇴
'==============================================================================
ElseIf uMode = "withdraw" Then

    Dim wPw, wAgree
    wPw    = Left(Trim(Request.Form("cur_pw") & ""), 64)
    wAgree = (Trim(Request.Form("agree") & "") = "1")

    If Not wAgree               Then JsAlert "탈퇴 동의에 체크해 주세요."
    If GLB_MY_LEVEL >= LV_ADMIN Then JsAlert "관리자 계정은 탈퇴할 수 없습니다."

    ' 소셜 전용 계정은 비밀번호가 없다. 비밀번호를 요구하면 탈퇴를 영구히 못 하게 되므로
    ' 대신 닉네임을 입력해 본인 확인을 대체한다.
    If Member.HasPwLogin(GLB_MY_ID) Then
        If StringUtil.IsBlank(wPw)                Then JsAlert "비밀번호를 입력해 주세요."
        If Not Member.ChkPassword(GLB_MY_ID, wPw) Then JsAlert "비밀번호가 일치하지 않습니다."
    Else
        ' 세션의 표시명이 아니라 DB 의 닉네임과 대조한다 (표시명은 대체값일 수 있다)
        Dim wNick, wArr
        wNick = Left(Trim(Request.Form("confirm_nick") & ""), 50)
        If StringUtil.IsBlank(wNick) Then JsAlert "확인을 위해 닉네임을 입력해 주세요."

        wArr = Member.Member_v(GLB_MY_ID)
        If Not IsArray(wArr) Then JsAlert "회원 정보를 찾을 수 없습니다."
        If wNick <> AV(wArr, MB_NICK, 0) Then JsAlert "닉네임이 일치하지 않습니다."
    End If

    Member.Member_d GLB_MY_ID
    Auth.Logout

    JsAlertGo "탈퇴 처리되었습니다. 그동안 이용해 주셔서 감사합니다.", U("/main/Default.asp")

Else
    JsAlert "잘못된 접근입니다."
End If
%>
