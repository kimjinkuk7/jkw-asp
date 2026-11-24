<%
'==============================================================================
' model/donate/PgInicis.Class.asp  -  KG이니시스 웹표준(INIStdPay) 어댑터
'------------------------------------------------------------------------------
' 웹표준 결제 흐름
'   1) 서버가 결제 파라미터 + 서명을 만든다
'   2) 화면의 hidden form 에 채우고 INIStdPay.pay(formId) 호출 → 이니시스 iframe 결제창
'   3) 결제 후 returnUrl 로 POST 로 돌아온다 (authToken, authUrl 등)
'   4) ★ 서버가 authUrl 로 승인 요청을 보낸다. 이 단계가 실제 승인이다.
'   5) resultCode = 0000 이면 성공
'
' 서명 (모두 SHA-256 hex 소문자)
'   signature = SHA256("oid=<oid>&price=<price>&timestamp=<ts>")
'   verification = SHA256("oid=<oid>&price=<price>&signKey=<signKey>&timestamp=<ts>")
'   mKey = SHA256(signKey)
'   승인단계 signature = SHA256("authToken=<token>&timestamp=<ts>")
'
' 전용 DLL 을 쓰지 않는다. SHA-256 은 Crypt(HASHBYTES) 로, 통신은 Http 로 한다.
'==============================================================================
Class CPgInicis

    '<Desc>이니시스가 요구하는 밀리초 타임스탬프</Desc>
    Public Function Timestamp()
        ' 1970-01-01 부터의 초 * 1000. 서명과 폼에 같은 값을 써야 한다.
        Timestamp = CStr(CDbl(DateDiff("s", "1970-01-01 00:00:00", Now())) * 1000)
    End Function

    Public Function ScriptUrl()
        If PAY_IS_TEST Then ScriptUrl = INICIS_TEST_JS Else ScriptUrl = INICIS_LIVE_JS
    End Function

    '<Desc>결제창 호출용 서명</Desc>
    Public Function Signature(ByVal pOid, ByVal pPrice, ByVal pTs)
        Signature = Crypt.Sha256Hex("oid=" & pOid & "&price=" & pPrice & "&timestamp=" & pTs)
    End Function

    '<Desc>결제창 호출용 검증값 (signKey 포함)</Desc>
    Public Function Verification(ByVal pOid, ByVal pPrice, ByVal pTs)
        Verification = Crypt.Sha256Hex("oid=" & pOid & "&price=" & pPrice & _
                                       "&signKey=" & INICIS_SIGNKEY & "&timestamp=" & pTs)
    End Function

    Public Function MKey()
        MKey = Crypt.Sha256Hex(INICIS_SIGNKEY)
    End Function

    '==========================================================================
    ' 승인 요청
    '==========================================================================
    '<Desc>authUrl 로 승인 요청. 성공하면 True. outTid / outMethod / outErr 를 채운다.</Desc>
    Public Function Approve(ByVal pAuthUrl, ByVal pAuthToken, ByVal pOrderId, ByVal pAmount, _
                            ByRef outTid, ByRef outMethod, ByRef outErr)
        Dim ts, sign, verify, body, code, paid

        Approve   = False
        outTid    = ""
        outMethod = ""
        outErr    = ""

        ' authUrl 은 이니시스가 준 값을 그대로 쓴다. 다만 우리 도메인으로 위조된 값이
        ' 들어오지 못하게 이니시스 도메인인지 확인한다.
        If InStr(LCase(pAuthUrl), "inicis.com") = 0 Then
            outErr = "승인 주소가 올바르지 않습니다."
            Exit Function
        End If

        ts     = Timestamp()
        sign   = Crypt.Sha256Hex("authToken=" & pAuthToken & "&timestamp=" & ts)
        verify = Crypt.Sha256Hex("authToken=" & pAuthToken & "&signKey=" & INICIS_SIGNKEY & "&timestamp=" & ts)

        body = "mid=" & Server.URLEncode(INICIS_MID) & _
               "&authToken=" & Server.URLEncode(pAuthToken) & _
               "&signature=" & Server.URLEncode(sign) & _
               "&verification=" & Server.URLEncode(verify) & _
               "&timestamp=" & Server.URLEncode(ts) & _
               "&charset=UTF-8&format=JSON"

        Http.Reset
        Http.TimeoutMs = 20000
        If Not Http.Post(pAuthUrl, body, "application/x-www-form-urlencoded") Then
            outErr = "승인 요청 실패 : " & Http.ErrMsg
            Donate.Log_i pOrderId, "inicis", "error", 0, body, Http.ErrMsg
            Exit Function
        End If

        Donate.Log_i pOrderId, "inicis", "approve", Http.Status, body, Http.Body

        If Not Json.Parse(Http.Body) Then
            outErr = "승인 응답을 해석할 수 없습니다."
            Exit Function
        End If

        code = Json.Str("resultCode", "")
        If code <> "0000" Then
            outErr = Json.Str("resultMsg", "승인이 거절되었습니다. (" & code & ")")
            Exit Function
        End If

        ' ★ 금액 대조. 승인 응답 금액이 주문 금액과 다르면 성공으로 보지 않는다.
        paid = CLng(Json.Num("TotPrice", 0))
        If paid <> CLng(pAmount) Then
            outErr = "승인 금액이 주문 금액과 다릅니다. (" & paid & " / " & pAmount & ")"
            Exit Function
        End If

        outTid    = Json.Str("tid", "")
        outMethod = Json.Str("payMethod", "카드")
        Approve   = True
    End Function

End Class

Dim PgInicis
Set PgInicis = New CPgInicis
%>
