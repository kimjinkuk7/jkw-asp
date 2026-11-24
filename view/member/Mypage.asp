<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/member/Mypage.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>마이페이지</li>
        </ul>
        <h2>마이페이지</h2>
        <p class="desc"><strong><%= AVH(meArr, MB_NICK, 0) %></strong> 님, 반갑습니다.</p>
    </div>

    <nav class="mypage-tabs" aria-label="마이페이지 메뉴">
        <a href="<%= U("/member/Mypage.asp") %>?tab=info"     class="<%= IIfStr(myTab = "info",     "is-on", "") %>">내 정보</a>
        <a href="<%= U("/member/Mypage.asp") %>?tab=article"  class="<%= IIfStr(myTab = "article",  "is-on", "") %>">내가 쓴 글</a>
        <a href="<%= U("/member/Mypage.asp") %>?tab=comment"  class="<%= IIfStr(myTab = "comment",  "is-on", "") %>">내 댓글</a>
        <a href="<%= U("/member/Mypage.asp") %>?tab=password" class="<%= IIfStr(myTab = "password", "is-on", "") %>">비밀번호 변경</a>
        <a href="<%= U("/member/Mypage.asp") %>?tab=withdraw" class="<%= IIfStr(myTab = "withdraw", "is-on", "") %>">회원 탈퇴</a>
    </nav>


<%'==================== 내 정보 ====================%>
<% If myTab = "info" Then %>

    <form method="post" action="<%= U("/member/Update_p.asp") %>" data-validate>
        <%= Auth.TokenField() %>
        <input type="hidden" name="mode" value="info">

        <div class="form-table">
            <div class="row">
                <span class="form-label">아이디</span>
                <div><input class="input" type="text" value="<%= AVH(meArr, MB_LOGINID, 0) %>" readonly></div>
            </div>
            <div class="row">
                <span class="form-label">등급</span>
                <div>
                    <span class="badge badge-primary">
<%                      Select Case AVI(meArr, MB_LEVEL, 0)
                            Case LV_ADMIN  : Response.Write "최고관리자"
                            Case LV_STAFF  : Response.Write "운영자"
                            Case LV_MEMBER : Response.Write "일반회원"
                            Case Else      : Response.Write "제한"
                        End Select %>
                    </span>
                    <span class="fs-xs muted">가입일 <%= HE(DateUtil.Fmt(AV(meArr, MB_REGDT, 0), "ymd")) %> · 로그인 <%= AVI(meArr, MB_LOGINCNT, 0) %>회</span>
                </div>
            </div>
            <div class="row">
                <label class="form-label" for="m-name">이름<span class="req">*</span></label>
                <div><input class="input" type="text" id="m-name" name="user_name" value="<%= AVH(meArr, MB_NAME, 0) %>" maxlength="50" data-required data-label="이름"></div>
            </div>
            <div class="row">
                <label class="form-label" for="m-nick">닉네임<span class="req">*</span></label>
                <div><input class="input" type="text" id="m-nick" name="nickname" value="<%= AVH(meArr, MB_NICK, 0) %>" maxlength="50" data-required data-label="닉네임"></div>
            </div>
            <div class="row">
                <label class="form-label" for="m-email">이메일<span class="req">*</span></label>
                <div><input class="input" type="email" id="m-email" name="email" value="<%= AVH(meArr, MB_EMAIL, 0) %>" maxlength="100" data-required data-label="이메일" data-type="email"></div>
            </div>
            <div class="row">
                <label class="form-label" for="m-mobile">연락처</label>
                <div><input class="input" type="tel" id="m-mobile" name="mobile" value="<%= AVH(meArr, MB_MOBILE, 0) %>" maxlength="20"></div>
            </div>
            <div class="row">
                <label class="form-label" for="m-home">홈페이지</label>
                <div><input class="input" type="url" id="m-home" name="homepage" value="<%= AVH(meArr, MB_HOMEPAGE, 0) %>" maxlength="200" placeholder="https://"></div>
            </div>
            <div class="row">
                <label class="form-label" for="m-intro">자기소개</label>
                <div><textarea class="textarea" id="m-intro" name="intro" rows="4" maxlength="500" style="min-height:100px"><%= AVH(meArr, MB_INTRO, 0) %></textarea></div>
            </div>
            <div class="row">
                <span class="form-label">메일 수신</span>
                <div>
                    <label class="check">
                        <input type="checkbox" name="mail_recv" value="1" <%= IIfStr(AVB(meArr, MB_MAILRECV, 0), "checked", "") %>>
                        <span>이메일 수신에 동의합니다.</span>
                    </label>
                </div>
            </div>
        </div>

        <div class="btn-row right mt-5">
            <button type="submit" class="btn btn-primary btn-lg">저장</button>
        </div>
    </form>


<%'==================== 내가 쓴 글 ====================%>
<% ElseIf myTab = "article" Then %>

    <div class="board-head"><span class="count"><%= Paging.Summary() %></span></div>

<% If myCnt = 0 Then %>
    <div class="empty">작성한 글이 없습니다.</div>
<% Else %>
    <table class="board-list">
        <caption class="sr-only">내가 쓴 글</caption>
        <thead>
            <tr>
                <th scope="col" class="col-cate">게시판</th>
                <th scope="col" class="col-title">제목</th>
                <th scope="col" class="col-date">작성일</th>
            </tr>
        </thead>
        <tbody>
<%      Dim ai
        For ai = 0 To myCnt - 1 %>
            <tr>
                <td class="col-cate"><span class="badge"><%= AVH(myArr, 2, ai) %></span></td>
                <td class="col-title">
                    <a href="<%= BoardViewUrl(AV(myArr, 1, ai), AVI(myArr, 0, ai)) %>"><%= AVH(myArr, 3, ai) %></a>
<%                  If AVI(myArr, 4, ai) > 0 Then %><span class="cmt">[<%= AVI(myArr, 4, ai) %>]</span><% End If %>
                    <div class="meta">
                        <span><%= AVH(myArr, 2, ai) %></span>
                        <span><%= HE(DateUtil.Fmt(AV(myArr, 5, ai), "ymd")) %></span>
                    </div>
                </td>
                <td class="col-date"><%= HE(DateUtil.Fmt(AV(myArr, 5, ai), "ymd")) %></td>
            </tr>
<%      Next %>
        </tbody>
    </table>
    <%= Paging.Render() %>
<% End If %>


<%'==================== 내 댓글 ====================%>
<% ElseIf myTab = "comment" Then %>

<% If myCnt = 0 Then %>
    <div class="empty">작성한 댓글이 없습니다.</div>
<% Else %>
    <div class="comment-list">
<%  Dim mi
    For mi = 0 To myCnt - 1 %>
        <div class="comment-item">
            <div class="head">
                <a class="name" href="<%= BoardViewUrl(AV(myArr, 1, mi), AVI(myArr, 2, mi)) %>"><%= AVH(myArr, 4, mi) %></a>
                <span class="date"><%= HE(DateUtil.Fmt(AV(myArr, 5, mi), "ymdhi")) %></span>
            </div>
            <div class="body"><%= AVH(myArr, 3, mi) %></div>
        </div>
<%  Next %>
    </div>
<% End If %>


<%'==================== 비밀번호 변경 ====================%>
<% ElseIf myTab = "password" Then %>

<% If Not AVB(meArr, MB_PWLOGIN, 0) Then %>
    <div class="alert alert-info">
        <strong>소셜 로그인 전용 계정입니다.</strong><br>
        비밀번호를 사용하지 않으므로 변경할 것이 없습니다.
        연결된 소셜 계정으로 계속 로그인해 주세요.
    </div>
<% Else %>
    <form method="post" action="<%= U("/member/Update_p.asp") %>" data-validate>
        <%= Auth.TokenField() %>
        <input type="hidden" name="mode" value="password">

        <div class="form-table">
            <div class="row">
                <label class="form-label" for="p-cur">현재 비밀번호<span class="req">*</span></label>
                <div><input class="input" type="password" id="p-cur" name="cur_pw" maxlength="64" autocomplete="current-password" data-required data-label="현재 비밀번호"></div>
            </div>
            <div class="row">
                <label class="form-label" for="p-new">새 비밀번호<span class="req">*</span></label>
                <div>
                    <input class="input" type="password" id="p-new" name="new_pw" maxlength="64" autocomplete="new-password" data-required data-label="새 비밀번호" data-type="pw">
                    <p class="form-help">ASCII 문자 8자 이상, 영문+숫자 포함</p>
                </div>
            </div>
            <div class="row">
                <label class="form-label" for="p-new2">새 비밀번호 확인<span class="req">*</span></label>
                <div><input class="input" type="password" id="p-new2" name="new_pw2" maxlength="64" autocomplete="new-password" data-required data-label="비밀번호 확인" data-type="pwconfirm" data-target="new_pw"></div>
            </div>
        </div>

        <div class="btn-row right mt-5">
            <button type="submit" class="btn btn-primary btn-lg">비밀번호 변경</button>
        </div>
    </form>
<% End If %>


<%'==================== 회원 탈퇴 ====================%>
<% Else %>

    <div class="alert alert-warn">
        <strong>탈퇴 시 유의사항</strong><br>
        · 계정은 즉시 비활성화되며 다시 로그인할 수 없습니다.<br>
        · 작성하신 게시글과 댓글은 삭제되지 않고 그대로 유지됩니다.<br>
        · 동일한 아이디로 재가입할 수 없습니다.
    </div>

    <form method="post" action="<%= U("/member/Update_p.asp") %>" data-validate>
        <%= Auth.TokenField() %>
        <input type="hidden" name="mode" value="withdraw">

        <div class="form-table">
<%          '--- 소셜 전용 계정은 비밀번호가 없다. 닉네임 입력으로 본인 확인을 대체한다. ---
            If AVB(meArr, MB_PWLOGIN, 0) Then %>
            <div class="row">
                <label class="form-label" for="w-pw">비밀번호 확인<span class="req">*</span></label>
                <div><input class="input" type="password" id="w-pw" name="cur_pw" maxlength="64" autocomplete="current-password" data-required data-label="비밀번호"></div>
            </div>
<%          Else %>
            <div class="row">
                <label class="form-label" for="w-nick">닉네임 확인<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="w-nick" name="confirm_nick" maxlength="50"
                           autocomplete="off" data-required data-label="닉네임">
                    <p class="form-help">
                        소셜 로그인 전용 계정이라 비밀번호가 없습니다.
                        확인을 위해 닉네임 <strong><%= AVH(meArr, MB_NICK, 0) %></strong> 을 그대로 입력해 주세요.
                    </p>
                </div>
            </div>
<%          End If %>
            <div class="row">
                <span class="form-label">동의</span>
                <div>
                    <label class="check">
                        <input type="checkbox" name="agree" value="1" required>
                        <span>위 내용을 확인했으며 탈퇴에 동의합니다.</span>
                    </label>
                </div>
            </div>
        </div>

        <div class="btn-row right mt-5">
            <button type="submit" class="btn btn-danger btn-lg" data-confirm="정말 탈퇴하시겠습니까? 되돌릴 수 없습니다.">회원 탈퇴</button>
        </div>
    </form>

<% End If %>

</div>

<!--#include file="../include/Footer.asp"-->
