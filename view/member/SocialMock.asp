<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/member/SocialMock.Code.asp"-->
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>모의 로그인 · <%= HE(smLabel) %></title>
<link rel="stylesheet" href="<%= APP_ROOT %>/assets/css/style.css?v=<%= HE(GLB_ASSET_VER) %>">
</head>
<body class="mock-body">

<div class="mock-card">

    <div class="mock-ribbon">시연용 모의 로그인 · 실제 <%= HE(smLabel) %> 서버가 아닙니다</div>

    <div class="mock-head mock-<%= HE(smProvider) %>">
        <span class="mock-brand"><%= HE(smLabel) %></span>
        <span class="mock-sub">계정으로 로그인</span>
    </div>

    <div class="mock-body-inner">
        <p class="mock-desc">
            <strong>jkw.com</strong> 에서 로그인을 요청했습니다.<br>
            사용할 계정을 선택해 주세요.
        </p>

        <ul class="mock-accounts">
<%      Dim mi, acc
        For mi = 0 To UBound(smAccounts)
            acc = smAccounts(mi) %>
            <li>
                <a href="<%= U("/member/SocialCallback.asp") %>?provider=<%= UE(smProvider) %>&amp;code=MOCK&amp;state=<%= UE(smState) %>&amp;acct=<%= UE(acc(0)) %><%= IIfStr(smPopup, "&amp;popup=1", "") %>">
                    <span class="ava"><%= HE(Left(acc(1), 1)) %></span>
                    <span class="info">
                        <span class="nm"><%= HE(acc(1)) %></span>
                        <span class="em"><%= HE(acc(2)) %></span>
                    </span>
                    <span class="arw">›</span>
                </a>
            </li>
<%      Next %>
        </ul>

        <div class="mock-scope">
            <div class="ttl">jkw.com 에 제공되는 정보</div>
            <ul>
                <li>프로필 이름</li>
                <li>이메일 주소</li>
            </ul>
            <p class="note">
                실제 OAuth 에서도 우리 서버는 비밀번호를 받지 않습니다.
                제공자가 발급한 토큰으로 위 정보만 조회합니다.
                그래서 이 모의 화면에도 비밀번호 입력칸을 두지 않았습니다.
            </p>
        </div>

        <div class="mock-foot">
            <a href="<%= U("/member/Login.asp") %>"<%= IIfStr(smPopup, " onclick=""window.close();return false;""", "") %>>취소</a>
        </div>
    </div>
</div>

</body>
</html>
