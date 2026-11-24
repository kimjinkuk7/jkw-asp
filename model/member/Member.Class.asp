<%
'==============================================================================
' model/member/Member.Class.asp  -  회원 (tb_member)
'------------------------------------------------------------------------------
' 비밀번호 : LOWER(HEX(SHA2_256(user_salt || 평문))) - Crypt.HashPassword 사용
'==============================================================================

'--- Member_v() 결과 인덱스 --------------------------------------------------
Const MB_ID       = 0
Const MB_LOGINID  = 1
Const MB_NAME     = 2
Const MB_NICK     = 3
Const MB_EMAIL    = 4
Const MB_MOBILE   = 5
Const MB_HOMEPAGE = 6
Const MB_PROFILE  = 7
Const MB_INTRO    = 8
Const MB_LEVEL    = 9
Const MB_MAILRECV = 10
Const MB_LOGINCNT = 11
Const MB_LASTLOGIN= 12
Const MB_ISUSE    = 13
Const MB_REGDT    = 14
Const MB_PWLOGIN  = 15      ' 0 = 소셜 전용 계정 (비밀번호 로그인 불가)

Class CMember

    Public DBHelper

    Private Sub Class_Initialize()
        Set DBHelper = New CDBHelper
        DBHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    End Sub

    Private Function Cols()
        Cols = " member_id, login_id, user_name, nickname, email, COALESCE(mobile,''), " & _
               " COALESCE(homepage,''), COALESCE(profile_img,''), COALESCE(intro,''), " & _
               " member_level, is_email_recv, login_count, last_login_dt, is_use, reg_dt, " & _
               " pw_login_use "
    End Function

    '<Desc>비밀번호 로그인 사용 여부를 끈다 (소셜 전용 계정)</Desc>
    Public Sub PwLoginOff_u(ByVal pMemberId)
        With DBHelper
            .Sql = "UPDATE tb_member SET pw_login_use = 0 WHERE member_id = ?"
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>이 회원이 비밀번호로 로그인할 수 있는가</Desc>
    Public Function HasPwLogin(ByVal pMemberId)
        Dim v
        With DBHelper
            .Sql = "SELECT pw_login_use FROM tb_member WHERE member_id = ?"
            .SetParamInt(pMemberId)
        End With
        v = DBHelper.ExecuteScalar()
        HasPwLogin = StringUtil.ToBool(v)
    End Function

    '==========================================================================
    ' 조회
    '==========================================================================
    Public Function Member_v(ByVal pMemberId)
        With DBHelper
            .Sql = "SELECT " & Cols() & " FROM tb_member WHERE member_id = ? AND is_del = 0"
            .SetParamInt(pMemberId)
        End With
        Member_v = DBHelper.ExecuteGetRows()
    End Function

    Public Function MemberByLoginId_v(ByVal pLoginId)
        With DBHelper
            .Sql = "SELECT " & Cols() & " FROM tb_member WHERE login_id = ? AND is_del = 0"
            .SetParamVarchar(LCase(pLoginId))
        End With
        MemberByLoginId_v = DBHelper.ExecuteGetRows()
    End Function

    '==========================================================================
    ' 중복 검사
    '==========================================================================
    Public Function IsDupLoginId(ByVal pLoginId)
        IsDupLoginId = Exists("login_id", LCase(pLoginId), 0)
    End Function

    Public Function IsDupEmail(ByVal pEmail, ByVal pExceptId)
        IsDupEmail = Exists("email", pEmail, pExceptId)
    End Function

    Public Function IsDupNickname(ByVal pNick, ByVal pExceptId)
        IsDupNickname = Exists("nickname", pNick, pExceptId)
    End Function

    Private Function Exists(ByVal pCol, ByVal pVal, ByVal pExceptId)
        Dim v, sql
        ' pCol 은 내부 호출로만 전달되는 고정 문자열이다 (외부 입력 아님)
        sql = "SELECT COUNT(*) FROM tb_member WHERE " & pCol & " = ? AND is_del = 0"
        If StringUtil.ToInt(pExceptId, 0) > 0 Then sql = sql & " AND member_id <> ?"

        With DBHelper
            .Sql = sql
            .SetParamVarchar(pVal)
            If StringUtil.ToInt(pExceptId, 0) > 0 Then .SetParamInt(pExceptId)
        End With
        v = DBHelper.ExecuteScalar()
        Exists = (StringUtil.ToInt(v, 0) > 0)
    End Function

    '==========================================================================
    ' 가입 / 수정
    '==========================================================================

    '<Desc>회원 가입. 새 member_id 반환 (실패 0)</Desc>
    Public Function Member_i(ByVal pLoginId, ByVal pPlainPw, ByVal pName, ByVal pNick, _
                             ByVal pEmail, ByVal pMobile, ByVal pIntro, ByVal pMailRecv)
        Dim salt, hash
        salt = Crypt.MakeSalt()
        hash = Crypt.HashPassword(salt, pPlainPw)

        With DBHelper
            .Sql = "INSERT INTO tb_member " & _
                   " (login_id, user_pw, user_salt, user_name, nickname, email, mobile, " & _
                   "  intro, member_level, is_email_recv, reg_ip) " & _
                   "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            .SetParamVarchar(LCase(pLoginId))
            .SetParamVarchar(hash)
            .SetParamVarchar(salt)
            .SetParamVarchar(pName)
            .SetParamVarchar(pNick)
            .SetParamVarchar(pEmail)
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pIntro)
            .SetParamInt(LV_MEMBER)
            .SetParamBit(pMailRecv)
            .SetParamVarchar(RequestUtil.ClientIp())
        End With
        Member_i = DBHelper.ExecuteIdentity()
    End Function

    '<Desc>휴대폰 인증 기록. pType 은 mock / firebase / sms (App.Conf 의 PHONE_AUTH_MODE)</Desc>
    Public Sub MobileVerify_u(ByVal pMemberId, ByVal pMobile, ByVal pType)
        With DBHelper
            .Sql = "UPDATE tb_member " & _
                   "   SET mobile             = ?, " & _
                   "       mobile_verified_dt = CURRENT_TIMESTAMP, " & _
                   "       mobile_verify_type = ?, " & _
                   "       upd_dt             = CURRENT_TIMESTAMP " & _
                   " WHERE member_id = ?"
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pType)
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>내 정보 수정</Desc>
    Public Sub Member_u(ByVal pMemberId, ByVal pName, ByVal pNick, ByVal pEmail, _
                        ByVal pMobile, ByVal pHomepage, ByVal pIntro, ByVal pMailRecv)
        ' 번호가 바뀌면 휴대폰 인증 기록을 지운다. 그러지 않으면 다른 번호로
        ' 바꿔치기했는데 예전 인증이 그대로 남아 "인증된 번호"로 통한다.
        ' SET 절의 오른쪽 식은 모두 UPDATE 이전 값으로 평가되므로 이렇게 쓸 수 있다.
        With DBHelper
            .Sql = "UPDATE tb_member SET user_name = ?, nickname = ?, email = ?, " & _
                   "       mobile_verified_dt = CASE WHEN ISNULL(mobile,'') = ISNULL(?,'') " & _
                   "                                 THEN mobile_verified_dt ELSE NULL END, " & _
                   "       mobile_verify_type = CASE WHEN ISNULL(mobile,'') = ISNULL(?,'') " & _
                   "                                 THEN mobile_verify_type ELSE NULL END, " & _
                   "       mobile = ?, " & _
                   "       homepage = ?, intro = ?, is_email_recv = ?, upd_dt = CURRENT_TIMESTAMP " & _
                   " WHERE member_id = ?"
            .SetParamVarchar(pName)
            .SetParamVarchar(pNick)
            .SetParamVarchar(pEmail)
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pHomepage)
            .SetParamVarchar(pIntro)
            .SetParamBit(pMailRecv)
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>현재 비밀번호 확인</Desc>
    Public Function ChkPassword(ByVal pMemberId, ByVal pPlainPw)
        Dim arr, salt, dbPw
        ChkPassword = False

        With DBHelper
            .Sql = "SELECT user_salt, user_pw FROM tb_member WHERE member_id = ? AND is_del = 0"
            .SetParamInt(pMemberId)
        End With
        arr = DBHelper.ExecuteGetRows()
        If Not IsArray(arr) Then Exit Function

        salt = Trim(arr(0, 0) & "")
        dbPw = LCase(Trim(arr(1, 0) & ""))
        ChkPassword = (Crypt.HashPassword(salt, pPlainPw) = dbPw)
    End Function

    '<Desc>비밀번호 변경 (salt 도 함께 재발급)</Desc>
    Public Sub Password_u(ByVal pMemberId, ByVal pNewPlainPw)
        Dim salt, hash
        salt = Crypt.MakeSalt()
        hash = Crypt.HashPassword(salt, pNewPlainPw)

        With DBHelper
            .Sql = "UPDATE tb_member SET user_pw = ?, user_salt = ?, pw_fail_count = 0, " & _
                   "       lock_until_dt = NULL, upd_dt = CURRENT_TIMESTAMP WHERE member_id = ?"
            .SetParamVarchar(hash)
            .SetParamVarchar(salt)
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>탈퇴 (논리 삭제. 작성글은 유지)</Desc>
    Public Sub Member_d(ByVal pMemberId)
        With DBHelper
            .Sql = "UPDATE tb_member SET is_del = 1, is_use = 0, upd_dt = CURRENT_TIMESTAMP WHERE member_id = ?"
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '==========================================================================
    ' 마이페이지
    '==========================================================================

    '<Desc>내가 쓴 글  arr : article_id, board_id, board_name, title, cmt, reg_dt</Desc>
    Public Function MyArticle_l(ByVal pMemberId, ByVal pOffset, ByVal pRows)
        With DBHelper
            .Sql = "SELECT a.article_id, a.board_id, b.board_name, a.title, a.comment_count, a.reg_dt " & _
                   "  FROM tb_board_article a " & _
                   " INNER JOIN tb_board_master b ON b.board_id = a.board_id " & _
                   " WHERE a.writer_id = ? AND a.is_del = 0 " & _
                   " ORDER BY a.article_id DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
            .SetParamInt(pMemberId)
            .SetParamInt(pOffset)
            .SetParamInt(pRows)
        End With
        MyArticle_l = DBHelper.ExecuteGetRows()
    End Function

    Public Function MyArticle_c(ByVal pMemberId)
        Dim v
        With DBHelper
            .Sql = "SELECT COUNT(*) FROM tb_board_article WHERE writer_id = ? AND is_del = 0"
            .SetParamInt(pMemberId)
        End With
        v = DBHelper.ExecuteScalar()
        MyArticle_c = StringUtil.ToInt(v, 0)
    End Function

    '<Desc>내가 쓴 댓글</Desc>
    Public Function MyComment_l(ByVal pMemberId, ByVal pOffset, ByVal pRows)
        With DBHelper
            .Sql = "SELECT c.comment_id, c.board_id, c.article_id, c.content, a.title, c.reg_dt " & _
                   "  FROM tb_board_comment c " & _
                   " INNER JOIN tb_board_article a ON a.article_id = c.article_id AND a.is_del = 0 " & _
                   " WHERE c.writer_id = ? AND c.is_del = 0 " & _
                   " ORDER BY c.comment_id DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
            .SetParamInt(pMemberId)
            .SetParamInt(pOffset)
            .SetParamInt(pRows)
        End With
        MyComment_l = DBHelper.ExecuteGetRows()
    End Function

    '==========================================================================
    ' 관리자
    '==========================================================================
    Public Function Member_l(ByVal pWord, ByVal pLevel, ByVal pOffset, ByVal pRows)
        Dim ww, wl
        ww = "" : wl = ""
        If Not StringUtil.IsBlank(pWord) Then ww = " AND (login_id LIKE ? OR user_name LIKE ? OR nickname LIKE ? OR email LIKE ?) "
        If StringUtil.ToInt(pLevel, -1) >= 0 Then wl = " AND member_level = ? "

        With DBHelper
            .Sql = "SELECT " & Cols() & " FROM tb_member WHERE is_del = 0 " & ww & wl & _
                   " ORDER BY member_id DESC OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
            If Not StringUtil.IsBlank(pWord) Then
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
            End If
            If StringUtil.ToInt(pLevel, -1) >= 0 Then .SetParamInt(pLevel)
            .SetParamInt(pOffset)
            .SetParamInt(pRows)
        End With
        Member_l = DBHelper.ExecuteGetRows()
    End Function

    Public Function Member_c(ByVal pWord, ByVal pLevel)
        Dim ww, wl, v
        ww = "" : wl = ""
        If Not StringUtil.IsBlank(pWord) Then ww = " AND (login_id LIKE ? OR user_name LIKE ? OR nickname LIKE ? OR email LIKE ?) "
        If StringUtil.ToInt(pLevel, -1) >= 0 Then wl = " AND member_level = ? "

        With DBHelper
            .Sql = "SELECT COUNT(*) FROM tb_member WHERE is_del = 0 " & ww & wl
            If Not StringUtil.IsBlank(pWord) Then
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
                .SetParamVarchar("%" & pWord & "%")
            End If
            If StringUtil.ToInt(pLevel, -1) >= 0 Then .SetParamInt(pLevel)
        End With
        v = DBHelper.ExecuteScalar()
        Member_c = StringUtil.ToInt(v, 0)
    End Function

    Public Sub Level_u(ByVal pMemberId, ByVal pLevel, ByVal pIsUse)
        With DBHelper
            .Sql = "UPDATE tb_member SET member_level = ?, is_use = ?, upd_dt = CURRENT_TIMESTAMP WHERE member_id = ?"
            .SetParamInt(pLevel)
            .SetParamBit(pIsUse)
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>로그인 잠금 해제</Desc>
    Public Sub Unlock_u(ByVal pMemberId)
        With DBHelper
            .Sql = "UPDATE tb_member SET pw_fail_count = 0, lock_until_dt = NULL WHERE member_id = ?"
            .SetParamInt(pMemberId)
        End With
        DBHelper.Execute()
    End Sub

End Class

Dim Member
Set Member = New CMember
%>
