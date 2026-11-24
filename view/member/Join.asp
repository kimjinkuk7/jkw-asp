<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/member/Join.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>회원가입</li>
        </ul>
        <h2>회원가입</h2>
        <p class="desc">가입하시면 게시판 작성과 댓글 기능을 이용할 수 있습니다.</p>
    </div>

    <form method="post" action="<%= U("/member/Join_p.asp") %>" data-validate>
        <%= Auth.TokenField() %>

        <div class="form-table">

            <div class="row">
                <label class="form-label" for="j-id">아이디<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="j-id" name="login_id" value="<%= HE(jLoginId) %>"
                           maxlength="30" autocomplete="username"
                           data-required data-label="아이디" data-type="loginid">
                    <p class="form-help">영소문자로 시작하는 영문·숫자·언더바(_) 조합 4~30자</p>
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-pw">비밀번호<span class="req">*</span></label>
                <div>
                    <input class="input" type="password" id="j-pw" name="user_pw"
                           maxlength="64" autocomplete="new-password"
                           data-required data-label="비밀번호" data-type="pw">
                    <p class="form-help">
                        8자 이상, 영문+숫자 포함.<br>
                        <strong>ASCII 문자(영문·숫자·특수문자)만 사용할 수 있습니다.</strong>
                        4개 스택(ASP · Spring Boot · Python · PHP)에서 해시값이 동일해야 하므로 한글은 허용하지 않습니다.
                    </p>
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-pw2">비밀번호 확인<span class="req">*</span></label>
                <div>
                    <input class="input" type="password" id="j-pw2" name="user_pw2"
                           maxlength="64" autocomplete="new-password"
                           data-required data-label="비밀번호 확인" data-type="pwconfirm" data-target="user_pw">
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-name">이름<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="j-name" name="user_name" value="<%= HE(jName) %>"
                           maxlength="50" data-required data-label="이름">
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-nick">닉네임<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="j-nick" name="nickname" value="<%= HE(jNick) %>"
                           maxlength="50" data-required data-label="닉네임">
                    <p class="form-help">게시판에 표시되는 이름입니다.</p>
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-email">이메일<span class="req">*</span></label>
                <div>
                    <input class="input" type="email" id="j-email" name="email" value="<%= HE(jEmail) %>"
                           maxlength="100" data-required data-label="이메일" data-type="email">
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-mobile">연락처</label>
                <div>
<% If LCase(PHONE_AUTH_MODE) = "off" Then %>
                    <input class="input" type="tel" id="j-mobile" name="mobile" value="<%= HE(jMobile) %>"
                           maxlength="20" placeholder="010-0000-0000">
<% Else %>
                    <div class="form-inline">
                        <input class="input" type="tel" id="j-mobile" name="mobile" value="<%= HE(jMobile) %>"
                               maxlength="20" placeholder="010-0000-0000" autocomplete="tel">
                        <button type="button" class="btn" id="j-phone-btn">인증번호 받기</button>
                    </div>

                    <%'--- 인증번호 입력 (인증번호를 받으면 나타난다) ---%>
                    <div class="otp-box" id="j-otp-box" hidden>
                        <div class="form-inline">
                            <input class="input otp-input" type="text" id="j-otp" maxlength="6"
                                   inputmode="numeric" autocomplete="one-time-code" placeholder="6자리">
                            <button type="button" class="btn" id="j-otp-btn">확인</button>
                            <span class="otp-timer" id="j-otp-timer"></span>
                        </div>
                        <div class="otp-hint" id="j-otp-hint"></div>
                    </div>

                    <div class="form-help" id="j-phone-state" role="status" aria-live="polite"></div>

                    <div class="form-help">
                        번호를 입력하면 인증이 필요합니다. 비워두면 건너뜁니다.
<% If LCase(PHONE_AUTH_MODE) = "mock" Then %>
                        <br><strong>시연용 인증입니다.</strong>
                        문자를 실제로 보내지 않는 대신, 서버가 만든 인증번호를 입력칸 아래에 보여줍니다.
                        인증번호는 <code>tb_phone_auth</code> 에 저장되고 만료·시도횟수까지 실제와 같게 검사합니다.
<% End If %>
                    </div>
<% End If %>
                </div>
            </div>

            <div class="row">
                <label class="form-label" for="j-intro">자기소개</label>
                <div>
                    <textarea class="textarea" id="j-intro" name="intro" rows="4" maxlength="500"
                              style="min-height:100px"><%= HE(jIntro) %></textarea>
                </div>
            </div>

            <div class="row">
                <span class="form-label">약관 동의<span class="req">*</span></span>
                <div>
                    <div class="privacy-box">
                        <strong>개인정보 수집 및 이용 동의</strong><br>
                        · 수집 항목 : 아이디, 비밀번호(단방향 암호화), 이름, 닉네임, 이메일, 연락처, 접속 IP<br>
                        · 수집 목적 : 회원 식별, 게시판 서비스 제공, 문의 응대<br>
                        · 보유 기간 : 회원 탈퇴 시까지 (관계 법령에 따라 일정 기간 보관될 수 있음)<br>
                        · 동의를 거부할 수 있으나, 이 경우 회원가입이 불가합니다.
                    </div>
                    <div class="check-row mt-2">
                        <label class="check">
                            <input type="checkbox" name="agree" value="1" id="j-agree">
                            <span>위 내용에 동의합니다.</span>
                        </label>
                        <label class="check">
                            <input type="checkbox" name="mail_recv" value="1">
                            <span>이메일 수신 동의 (선택)</span>
                        </label>
                    </div>
                </div>
            </div>

        </div>

        <div class="btn-row between mt-5">
            <a class="btn" href="<%= U("/main/Default.asp") %>">취소</a>
            <button type="submit" class="btn btn-primary btn-lg">가입하기</button>
        </div>
    </form>

    <div class="auth-links mt-4">
        이미 계정이 있으신가요? <a href="<%= U("/member/Login.asp") %>" class="fw7">로그인</a>
    </div>

</div>

<%
g_extraScript = "<script>" & _
"document.querySelector('form[data-validate]').addEventListener('submit',function(e){" & _
"  if(!document.getElementById('j-agree').checked){" & _
"    alert('약관에 동의해 주세요.'); e.preventDefault(); e.stopImmediatePropagation();" & _
"  }" & _
"},true);" & _
"</script>"
%>
<!--#include file="../include/Footer.asp"-->
