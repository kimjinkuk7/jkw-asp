<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/member/Login.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">
        <h2>로그인</h2>
        <p class="desc"><%= HE(GLB_SITE_TITLE) %>에 오신 것을 환영합니다.</p>

        <div class="auth-box">
            <form method="post" action="<%= U("/member/Login_p.asp") %>" data-validate>
                <%= Auth.TokenField() %>
                <input type="hidden" name="rurl" value="<%= HE(rUrl) %>">

                <div class="form-group">
                    <label class="form-label" for="login_id">아이디</label>
                    <input class="input" type="text" id="login_id" name="login_id"
                           value="<%= HE(savedId) %>" maxlength="30" autocomplete="username"
                           data-required data-label="아이디">
                </div>

                <div class="form-group">
                    <label class="form-label" for="user_pw">비밀번호</label>
                    <input class="input" type="password" id="user_pw" name="user_pw"
                           maxlength="64" autocomplete="current-password"
                           data-required data-label="비밀번호">
                </div>

                <label class="check">
                    <input type="checkbox" name="save_id" value="1" <%= IIfStr(Len(savedId) > 0, "checked", "") %>>
                    <span>아이디 저장</span>
                </label>

                <button type="submit" class="btn btn-primary btn-block btn-lg mt-3">로그인</button>
            </form>

            <div class="auth-links">
                <a href="<%= U("/member/Join.asp") %>">회원가입</a>
                <span>·</span>
                <a href="<%= U("/board/List.asp") %>?bid=guest">방명록</a>
            </div>

<%          '--- 소셜 로그인 --------------------------------------------------
            Dim soList, soI, soP
            soList = Array("google", "naver", "kakao")
%>
            <div class="social-login">
                <div class="social-sep"><span>또는</span></div>
<%              For soI = 0 To UBound(soList)
                    soP = soList(soI) %>
                <a class="btn btn-block social-btn social-<%= soP %>" data-social="<%= soP %>"
                   href="<%= U("/member/SocialStart.asp") %>?provider=<%= UE(soP) %>&amp;rurl=<%= UE(rUrl) %>">
                    <%= HE(Social.Label(soP)) %> 계정으로 로그인
                </a>
<%              Next %>
<% If LCase(SOCIAL_MODE) = "mock" Then %>
                <p class="form-help tc">
                    <strong>시연용 로그인입니다.</strong>
                    별도 창이 열리지만 실제 제공자 서버가 아니라
                    이 사이트의 모의 동의 화면입니다.
                </p>
<% End If %>
            </div>
        </div>

        <div class="alert alert-info mt-4 fs-xs">
            비밀번호 <%= LOGIN_FAIL_LIMIT %>회 연속 오류 시 <%= LOGIN_LOCK_MIN %>분간 로그인이 제한됩니다.
        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
