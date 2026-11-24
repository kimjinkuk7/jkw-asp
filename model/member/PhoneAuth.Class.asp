<%
'==============================================================================
' model/member/PhoneAuth.Class.asp  -  휴대폰 인증번호 발급 / 검증
'------------------------------------------------------------------------------
' 실제 SMS 인증과 같은 구조다. mock 이든 실제든 이 클래스는 바뀌지 않는다.
' 달라지는 것은 "문자를 실제로 보내는가" 한 가지뿐이다.
'
'   Send_i()     인증번호를 만들어 DB 에 넣는다
'   Latest_v()   그 번호의 가장 최근 미인증 건
'   Verify_p()   입력값 대조 (만료 / 시도횟수 확인 포함)
'==============================================================================

'--- Latest_v() 결과 인덱스 ---
Const PA_ID       = 0
Const PA_MOBILE   = 1
Const PA_CODE     = 2
Const PA_MODE     = 3
Const PA_EXPIRE   = 4
Const PA_TRY      = 5
Const PA_REGDT    = 6

Class CPhoneAuth

    Public DBHelper

    Private Sub Class_Initialize()
        Set DBHelper = New CDBHelper
        DBHelper.ConnectionString = ConnectionString(DBCONN_JKWCOM)
    End Sub

    '<Desc>6자리 인증번호. 0 으로 시작하지 않게 만든다.</Desc>
    Public Function NewCode()
        Dim s, i
        Randomize Timer * (1 + Len(CStr(Request.ServerVariables("REMOTE_PORT") & "")))
        s = CStr(Int(Rnd() * 9) + 1)
        For i = 2 To 6
            s = s & CStr(Int(Rnd() * 10))
        Next
        NewCode = s
    End Function

    '<Desc>인증번호 발급. 반환값 = auth_id (실패 0)</Desc>
    Public Function Send_i(ByVal pMobile, ByVal pCode, ByVal pMode, ByVal pValidMin)
        With DBHelper
            .Sql = "INSERT INTO tb_phone_auth " & _
                   " (mobile, auth_code, auth_mode, expire_dt, reg_ip) " & _
                   "VALUES (?, ?, ?, DATEADD(MINUTE, ?, CURRENT_TIMESTAMP), ?)"
            .SetParamVarchar(pMobile)
            .SetParamVarchar(pCode)
            .SetParamVarchar(LCase(CStr(pMode)))
            .SetParamInt(pValidMin)
            .SetParamVarchar(RequestUtil.ClientIp())
        End With
        Send_i = DBHelper.ExecuteIdentity()
    End Function

    '<Desc>해당 번호의 최근 미인증 건 (만료 여부는 호출측에서 판단)</Desc>
    Public Function Latest_v(ByVal pMobile)
        With DBHelper
            .Sql = "SELECT TOP 1 auth_id, mobile, auth_code, auth_mode, expire_dt, try_count, reg_dt " & _
                   "  FROM tb_phone_auth " & _
                   " WHERE mobile = ? AND verified_dt IS NULL " & _
                   " ORDER BY auth_id DESC"
            .SetParamVarchar(pMobile)
        End With
        Latest_v = DBHelper.ExecuteGetRows()
    End Function

    '<Desc>최근 pMin 분 안에 발급한 건수 (재발송 남용 방지)</Desc>
    Public Function RecentCount_c(ByVal pMobile, ByVal pMin)
        Dim v
        With DBHelper
            .Sql = "SELECT COUNT(*) FROM tb_phone_auth " & _
                   " WHERE mobile = ? AND reg_dt > DATEADD(MINUTE, -?, CURRENT_TIMESTAMP)"
            .SetParamVarchar(pMobile)
            .SetParamInt(pMin)
        End With
        v = DBHelper.ExecuteScalar()
        RecentCount_c = StringUtil.ToInt(v, 0)
    End Function

    Public Sub TryUp_u(ByVal pAuthId)
        With DBHelper
            .Sql = "UPDATE tb_phone_auth SET try_count = try_count + 1 WHERE auth_id = ?"
            .SetParamInt(pAuthId)
        End With
        DBHelper.Execute()
    End Sub

    '<Desc>인증 완료 표시. 미인증 건만 바꾼다.</Desc>
    Public Function Verified_u(ByVal pAuthId)
        With DBHelper
            .Sql = "UPDATE tb_phone_auth SET verified_dt = CURRENT_TIMESTAMP " & _
                   " WHERE auth_id = ? AND verified_dt IS NULL"
            .SetParamInt(pAuthId)
        End With
        DBHelper.Execute()
        Verified_u = (DBHelper.Affected > 0)
    End Function

    '<Desc>만료된 옛 건 정리 (하루 이상 지난 것)</Desc>
    Public Sub Purge_p()
        On Error Resume Next
        With DBHelper
            .Sql = "DELETE FROM tb_phone_auth WHERE reg_dt < DATEADD(DAY, -1, CURRENT_TIMESTAMP)"
        End With
        DBHelper.Execute()
        On Error GoTo 0
    End Sub

End Class

Dim PhoneAuth
Set PhoneAuth = New CPhoneAuth
%>
