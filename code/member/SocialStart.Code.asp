<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Member.Class.asp"-->
<!--#include file="../../model/member/Social.Class.asp"-->
<%
'==============================================================================
' code/member/SocialStart.Code.asp  -  소셜 로그인 시작
'------------------------------------------------------------------------------
' state 를 만들어 세션에 넣고 제공자 로그인 화면으로 보낸다.
' state 는 콜백에서 반드시 대조한다 (OAuth 의 CSRF 방어 수단).
'==============================================================================

Dim ssProvider, ssState, ssRurl, ssUrl

ssProvider = LCase(RequestUtil.ValueInit("provider", ""))
If Not Social.IsSupported(ssProvider) Then JsAlert "지원하지 않는 로그인 방식입니다."

If GLB_IS_LOGIN Then JsAlertGo "이미 로그인되어 있습니다.", U("/main/Default.asp")

'--- 로그인 후 돌아갈 주소 (열린 리다이렉트 방지 : 내부 경로만 허용) ---
ssRurl = RequestUtil.ValueInit("rurl", "")
If Len(ssRurl) = 0 Or Left(ssRurl, 1) <> "/" Or Left(ssRurl, 2) = "//" Then
    ssRurl = U("/main/Default.asp")
End If

' ★ 소셜 로그인은 iframe 모달로 띄울 수 없다.
'   구글·네이버·카카오 모두 X-Frame-Options / CSP frame-ancestors 로 프레임 삽입을
'   차단한다 (클릭재킹 방어). 제공자가 지원하는 방식은 팝업 또는 전체 이동뿐이다.
Dim ssPopup
ssPopup = (RequestUtil.ValueInit("popup", "") = "1")

ssState = StringUtil.NewKey()
Session("SOCIAL_STATE")    = ssState
Session("SOCIAL_PROVIDER") = ssProvider
Session("SOCIAL_RURL")     = ssRurl

If LCase(SOCIAL_MODE) = "mock" Then
    ' 시연용 : 제공자 대신 우리 도메인의 모의 계정선택/동의 화면으로 보낸다.
    Response.Redirect U("/member/SocialMock.asp") & "?provider=" & UE(ssProvider) & _
                      "&state=" & UE(ssState) & IIfStr(ssPopup, "&popup=1", "")
End If

If Not Social.IsLive(ssProvider) Then
    JsAlert Social.Label(ssProvider) & " 로그인 키가 설정되지 않았습니다." & vbCrLf & _
            "App.Conf.asp 에서 CLIENT_ID / SECRET 을 등록해 주세요."
End If

ssUrl = Social.AuthorizeUrl(ssProvider, ssState)
If Len(ssUrl) = 0 Then JsAlert "로그인 주소를 만들 수 없습니다."

Response.Redirect ssUrl
%>
