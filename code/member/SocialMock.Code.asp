<!--#include file="../Common.Code.asp"-->
<!--#include file="../../model/member/Social.Class.asp"-->
<%
'==============================================================================
' code/member/SocialMock.Code.asp  -  시연용 모의 로그인 / 동의 화면
'------------------------------------------------------------------------------
' 실제 OAuth 에서 제공자 도메인이 보여주는 "계정 선택 → 권한 동의" 단계를
' 우리 도메인에서 흉내낸 화면이다. 흐름을 눈으로 보기 위한 것이다.
'
' ★ 비밀번호 입력칸을 두지 않는다.
'   실제 제공자 로그인 화면을 그대로 모사하면서 비밀번호를 받는 UI 는
'   피싱 페이지와 구조가 같아진다. OAuth 에서 우리 서버는 원래 비밀번호를
'   볼 수 없는 것이 정상이므로, 계정 선택과 동의만 두는 것이 오히려 정확하다.
'==============================================================================
g_bodyClass = "page-social-mock"
g_pageTitle = "모의 로그인"

Dim smProvider, smState, smLabel, smAccounts, smPopup

smProvider = LCase(RequestUtil.ValueInit("provider", ""))
smState    = RequestUtil.ValueInit("state", "")
smPopup    = (RequestUtil.ValueInit("popup", "") = "1")

If Not Social.IsSupported(smProvider) Then JsAlert "지원하지 않는 로그인 방식입니다."
If LCase(SOCIAL_MODE) <> "mock" Then JsAlert "모의 로그인은 시연 모드에서만 사용합니다."

'--- state 대조. 여기서도 확인해야 임의 호출로 들어오는 것을 막는다 ---
If Len(smState) = 0 Or smState <> (Session("SOCIAL_STATE") & "") Then
    JsAlert "로그인 요청이 유효하지 않습니다. 다시 시도해 주세요."
End If

smLabel = Social.Label(smProvider)

'--- 고를 수 있는 데모 계정. uid 가 다르면 서로 다른 회원이 된다 ---
'    arr(n) = Array(uid접미사, 표시이름, 이메일)
smAccounts = Array( _
    Array("1", smLabel & " 데모",     smProvider & ".demo@example.com"), _
    Array("2", smLabel & " 데모2",    smProvider & ".demo2@example.com") )
%>
