<%
'==============================================================================
' model/member/Social.Class.asp  -  소셜 로그인 (구글 / 네이버 / 카카오)
'------------------------------------------------------------------------------
' 흐름
'   1) AuthorizeUrl()  제공자 로그인 화면으로 보낼 URL 을 만든다 (state 로 CSRF 방어)
'   2) 콜백에서 code 를 받아 Exchange() → 액세스 토큰
'   3) Profile()       토큰으로 사용자 정보 조회 → uid / email / name
'   4) FindMember()    (provider, uid) 로 기존 연동 조회
'      없으면 이메일로 기존 회원을 찾아 연결하거나 새 회원을 만든다
'
' 설계 메모
'   · 제공자 식별자(uid)를 유일 키로 쓴다. 이메일은 바뀔 수 있어 키로 부적합하다.
'   · 소셜 회원도 tb_member 에 행이 생긴다. 비밀번호는 로그인에 쓸 수 없는
'     임의값으로 채운다 (user_pw 가 NOT NULL 이고, 스키마를 4스택이 공유한다).
'   · SOCIAL_MODE = "mock" 이면 제공자 서버에 가지 않는다. 화면 흐름은 동일하다.
'==============================================================================
Class CSocial

    Public DBHelper

    Private Sub Class_Initialize()
        Set DBHelper = New CDBHelper
        DBHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    End Sub

    '==========================================================================
    ' 제공자 메타
    '==========================================================================
    Public Function IsSupported(ByVal pProvider)
        IsSupported = (InStr(",google,naver,kakao,", "," & LCase(CStr(pProvider)) & ",") > 0)
    End Function

    Public Function Label(ByVal pProvider)
        Select Case LCase(CStr(pProvider))
            Case "google" : Label = "구글"
            Case "naver"  : Label = "네이버"
            Case "kakao"  : Label = "카카오"
            Case Else     : Label = CStr(pProvider)
        End Select
    End Function

    '<Desc>해당 제공자가 실제 연동 가능한 상태인가 (키가 채워져 있는가)</Desc>
    Public Function IsLive(ByVal pProvider)
        IsLive = False
        If LCase(SOCIAL_MODE) <> "live" Then Exit Function
        Select Case LCase(CStr(pProvider))
            Case "google" : IsLive = (Len(GOOGLE_CLIENT_ID) > 0 And Len(GOOGLE_CLIENT_SECRET) > 0)
            Case "naver"  : IsLive = (Len(NAVER_CLIENT_ID)  > 0 And Len(NAVER_CLIENT_SECRET)  > 0)
            Case "kakao"  : IsLive = (Len(KAKAO_CLIENT_ID)  > 0)
        End Select
    End Function

    Private Function ClientId(ByVal p)
        Select Case LCase(CStr(p))
            Case "google" : ClientId = GOOGLE_CLIENT_ID
            Case "naver"  : ClientId = NAVER_CLIENT_ID
            Case "kakao"  : ClientId = KAKAO_CLIENT_ID
            Case Else     : ClientId = ""
        End Select
    End Function

    Private Function ClientSecret(ByVal p)
        Select Case LCase(CStr(p))
            Case "google" : ClientSecret = GOOGLE_CLIENT_SECRET
            Case "naver"  : ClientSecret = NAVER_CLIENT_SECRET
            Case "kakao"  : ClientSecret = KAKAO_CLIENT_SECRET
            Case Else     : ClientSecret = ""
        End Select
    End Function

    '<Desc>콜백 주소. 제공자 콘솔에 등록한 값과 한 글자도 달라선 안 된다.</Desc>
    Public Function RedirectUri(ByVal pProvider)
        RedirectUri = GLB_SITE_URL & VIEW_ROOT & "/member/SocialCallback.asp?provider=" & LCase(CStr(pProvider))
    End Function

    '==========================================================================
    ' 1) 로그인 화면 URL
    '==========================================================================
    Public Function AuthorizeUrl(ByVal pProvider, ByVal pState)
        Dim p, cid, ru
        p   = LCase(CStr(pProvider))
        cid = ClientId(p)
        ru  = Server.URLEncode(RedirectUri(p))

        Select Case p
            Case "google"
                AuthorizeUrl = "https://accounts.google.com/o/oauth2/v2/auth" & _
                               "?client_id=" & Server.URLEncode(cid) & _
                               "&redirect_uri=" & ru & _
                               "&response_type=code" & _
                               "&scope=" & Server.URLEncode("openid email profile") & _
                               "&state=" & Server.URLEncode(pState)
            Case "naver"
                AuthorizeUrl = "https://nid.naver.com/oauth2.0/authorize" & _
                               "?client_id=" & Server.URLEncode(cid) & _
                               "&redirect_uri=" & ru & _
                               "&response_type=code" & _
                               "&state=" & Server.URLEncode(pState)
            Case "kakao"
                AuthorizeUrl = "https://kauth.kakao.com/oauth/authorize" & _
                               "?client_id=" & Server.URLEncode(cid) & _
                               "&redirect_uri=" & ru & _
                               "&response_type=code" & _
                               "&state=" & Server.URLEncode(pState)
            Case Else
                AuthorizeUrl = ""
        End Select
    End Function

    '==========================================================================
    ' 2) code → 액세스 토큰
    '==========================================================================
    Public Function Exchange(ByVal pProvider, ByVal pCode, ByVal pState, ByRef outErr)
        Dim p, url, body, tok
        Exchange = ""
        outErr   = ""
        p = LCase(CStr(pProvider))

        Select Case p
            Case "google" : url = "https://oauth2.googleapis.com/token"
            Case "naver"  : url = "https://nid.naver.com/oauth2.0/token"
            Case "kakao"  : url = "https://kauth.kakao.com/oauth/token"
            Case Else     : outErr = "지원하지 않는 제공자입니다." : Exit Function
        End Select

        body = "grant_type=authorization_code" & _
               "&client_id=" & Server.URLEncode(ClientId(p)) & _
               "&client_secret=" & Server.URLEncode(ClientSecret(p)) & _
               "&redirect_uri=" & Server.URLEncode(RedirectUri(p)) & _
               "&code=" & Server.URLEncode(pCode)
        If p = "naver" Then body = body & "&state=" & Server.URLEncode(pState)

        Http.Reset
        Http.TimeoutMs = 10000
        If Not Http.Post(url, body, "application/x-www-form-urlencoded") Then
            outErr = "토큰 요청 실패 : " & Http.ErrMsg
            Exit Function
        End If

        If Not Json.Parse(Http.Body) Then
            outErr = "토큰 응답을 해석할 수 없습니다."
            Exit Function
        End If

        tok = Json.Str("access_token", "")
        If Len(tok) = 0 Then
            outErr = "토큰 발급 실패 : " & Json.Str("error_description", Json.Str("error", "알 수 없는 오류"))
            Exit Function
        End If
        Exchange = tok
    End Function

    '==========================================================================
    ' 3) 토큰 → 사용자 정보
    '==========================================================================
    Public Function Profile(ByVal pProvider, ByVal pToken, ByRef outUid, ByRef outEmail, ByRef outName, ByRef outErr)
        Dim p, url
        Profile  = False
        outUid   = "" : outEmail = "" : outName = "" : outErr = ""
        p = LCase(CStr(pProvider))

        Select Case p
            Case "google" : url = "https://openidconnect.googleapis.com/v1/userinfo"
            Case "naver"  : url = "https://openapi.naver.com/v1/nid/me"
            Case "kakao"  : url = "https://kapi.kakao.com/v2/user/me"
            Case Else     : outErr = "지원하지 않는 제공자입니다." : Exit Function
        End Select

        Http.Reset
        Http.TimeoutMs = 10000
        Http.SetHeader "Authorization", "Bearer " & pToken
        If Not Http.HttpGet(url) Then
            outErr = "사용자 정보 요청 실패 : " & Http.ErrMsg
            Exit Function
        End If

        If Not Json.Parse(Http.Body) Then
            outErr = "사용자 정보를 해석할 수 없습니다."
            Exit Function
        End If

        Select Case p
            Case "google"
                outUid   = Json.Str("sub", "")
                outEmail = Json.Str("email", "")
                outName  = Json.Str("name", "")
            Case "naver"
                outUid   = Json.Str("response.id", "")
                outEmail = Json.Str("response.email", "")
                outName  = Json.Str("response.nickname", Json.Str("response.name", ""))
            Case "kakao"
                outUid   = Json.Str("id", "")
                outEmail = Json.Str("kakao_account.email", "")
                outName  = Json.Str("kakao_account.profile.nickname", Json.Str("properties.nickname", ""))
        End Select

        If Len(outUid) = 0 Then
            outErr = "제공자에서 식별자를 받지 못했습니다."
            Exit Function
        End If
        Profile = True
    End Function

    '==========================================================================
    ' 4) 연동 조회 / 등록
    '==========================================================================
    '<Desc>(provider, uid) 로 연결된 member_id. 없으면 0</Desc>
    Public Function FindMemberId(ByVal pProvider, ByVal pUid)
        Dim v
        With DBHelper
            .Sql = "SELECT s.member_id FROM tb_member_social s " & _
                   "  INNER JOIN tb_member m ON m.member_id = s.member_id " & _
                   " WHERE s.provider = ? AND s.provider_uid = ? " & _
                   "   AND m.is_del = 0 AND m.is_use = 1"
            .SetParamVarchar(LCase(CStr(pProvider)))
            .SetParamVarchar(CStr(pUid))
        End With
        v = DBHelper.ExecuteScalar()
        FindMemberId = StringUtil.ToInt(v, 0)
    End Function

    '<Desc>이메일로 기존 회원 찾기 (소셜 계정을 기존 계정에 붙일 때)</Desc>
    Public Function FindMemberIdByEmail(ByVal pEmail)
        Dim v
        FindMemberIdByEmail = 0
        If StringUtil.IsBlank(pEmail) Then Exit Function
        With DBHelper
            .Sql = "SELECT member_id FROM tb_member " & _
                   " WHERE email = ? AND is_del = 0 AND is_use = 1"
            .SetParamVarchar(CStr(pEmail))
        End With
        v = DBHelper.ExecuteScalar()
        FindMemberIdByEmail = StringUtil.ToInt(v, 0)
    End Function

    Public Sub Link_i(ByVal pMemberId, ByVal pProvider, ByVal pUid, ByVal pEmail, ByVal pName)
        With DBHelper
            .Sql = "INSERT INTO tb_member_social " & _
                   " (member_id, provider, provider_uid, provider_email, provider_name, last_login_dt) " & _
                   "VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)"
            .SetParamInt(pMemberId)
            .SetParamVarchar(LCase(CStr(pProvider)))
            .SetParamVarchar(CStr(pUid))
            .SetParamVarchar(pEmail)
            .SetParamVarchar(pName)
        End With
        DBHelper.Execute()
    End Sub

    Public Sub TouchLogin_u(ByVal pProvider, ByVal pUid)
        With DBHelper
            .Sql = "UPDATE tb_member_social SET last_login_dt = CURRENT_TIMESTAMP " & _
                   " WHERE provider = ? AND provider_uid = ?"
            .SetParamVarchar(LCase(CStr(pProvider)))
            .SetParamVarchar(CStr(pUid))
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>내 연동 목록 (마이페이지)  arr(0)=provider arr(1)=email arr(2)=reg_dt</Desc>
    Public Function MyLinks_l(ByVal pMemberId)
        With DBHelper
            .Sql = "SELECT provider, COALESCE(provider_email,''), reg_dt " & _
                   "  FROM tb_member_social WHERE member_id = ? ORDER BY provider"
            .SetParamInt(pMemberId)
        End With
        MyLinks_l = DBHelper.ExecuteGetRows()
    End Function

    '<Desc>소셜 신규 회원 생성. 아이디/닉네임 충돌을 피해 만든다.</Desc>
    Public Function CreateMember(ByVal pProvider, ByVal pUid, ByVal pEmail, ByVal pName, ByRef outErr)
        Dim baseId, newLoginId, newNick, i, pw, email
        CreateMember = 0
        outErr = ""

        ' 로그인 아이디 : 제공자 + uid 앞부분. 규칙(영소문자 시작, 4~30자)을 지킨다.
        baseId = LCase(Left(CStr(pProvider), 1)) & StringUtil.DigitsOnly(pUid)
        If Len(baseId) < 4 Then baseId = LCase(CStr(pProvider)) & baseId
        baseId = Left(baseId, 24)

        newLoginId = baseId
        i = 0
        Do While Member.IsDupLoginId(newLoginId)
            i = i + 1
            If i > 50 Then outErr = "아이디를 만들 수 없습니다." : Exit Function
            newLoginId = Left(baseId, 24) & CStr(i)
        Loop

        ' 닉네임 : 제공자가 준 이름을 쓰고, 겹치면 뒤에 숫자를 붙인다
        newNick = Trim(CStr(pName & ""))
        If Len(newNick) = 0 Then newNick = Label(pProvider) & "사용자"
        newNick = Left(newNick, 40)
        i = 0
        Do While Member.IsDupNickname(newNick, 0)
            i = i + 1
            If i > 50 Then outErr = "닉네임을 만들 수 없습니다." : Exit Function
            newNick = Left(CStr(pName & Label(pProvider)), 40) & CStr(i)
        Loop

        ' 이메일이 없으면 로그인 불가능한 자리표시 주소를 만든다 (email 이 NOT NULL)
        email = Trim(CStr(pEmail & ""))
        If Not StringUtil.ChkEmail(email) Then email = newLoginId & "@social.invalid"
        If Member.IsDupEmail(email, 0) Then email = newLoginId & "@social.invalid"

        ' 비밀번호로 로그인할 수 없게 임의값을 넣는다 (user_pw 는 NOT NULL)
        pw = "SOCIAL-" & StringUtil.NewKey()

        CreateMember = Member.Member_i(newLoginId, pw, Left(newNick, 50), newNick, email, "", "", False)
        If CreateMember <= 0 Then
            outErr = "회원 생성에 실패했습니다."
            Exit Function
        End If

        ' ★ 소셜 전용 계정으로 표시한다.
        '   이 표시가 없으면 비밀번호 확인을 요구하는 화면(비밀번호 변경·탈퇴)에서
        '   본인이 모르는 임의 비밀번호를 요구받아 아무것도 할 수 없게 된다.
        Member.PwLoginOff_u CreateMember
    End Function

End Class

Dim Social
Set Social = New CSocial
%>
