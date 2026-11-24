<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Contact.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>문의 관리</h2>
                <p class="desc">사이트 문의하기로 접수된 내역입니다.</p>
            </div>

<%'==================== 상세 ====================%>
<% If ctIsView Then %>
            <div class="admin-card">
                <h3>
                    <%= AVH(ctArr, CT_SUBJECT, 0) %>
<%                  If AVB(ctArr, CT_ISREPLY, 0) Then %><span class="badge badge-primary">답변완료</span><% End If %>
                </h3>

                <div class="form-table">
                    <div class="row"><span class="form-label">유형</span><div><%= AVH(ctArr, CT_TYPE, 0) %></div></div>
                    <div class="row"><span class="form-label">이름</span><div><%= AVH(ctArr, CT_NAME, 0) %> <span class="muted"><%= AVH(ctArr, CT_COMPANY, 0) %></span></div></div>
                    <div class="row"><span class="form-label">이메일</span><div><a href="mailto:<%= AVH(ctArr, CT_EMAIL, 0) %>"><%= AVH(ctArr, CT_EMAIL, 0) %></a></div></div>
                    <div class="row"><span class="form-label">연락처</span><div><%= AVH(ctArr, CT_MOBILE, 0) %></div></div>
                    <div class="row"><span class="form-label">접수일</span><div><%= HE(DateUtil.Fmt(AV(ctArr, CT_REGDT, 0), "ymdhis")) %> <span class="muted fs-xs">IP <%= AVH(ctArr, CT_IP, 0) %></span></div></div>
                    <div class="row">
                        <span class="form-label">내용</span>
                        <div style="white-space:pre-wrap;background:var(--c-bg-sub);padding:var(--sp-4);border-radius:var(--radius-sm)"><%= AVH(ctArr, CT_CONTENT, 0) %></div>
                    </div>
                </div>
            </div>

            <div class="admin-card">
                <h3>답변 <%= IIfStr(AVB(ctArr, CT_ISREPLY, 0), "(" & HE(DateUtil.Fmt(AV(ctArr, CT_REPLYDT, 0), "ymdhi")) & " 작성)", "") %></h3>
                <form method="post" action="<%= U("/admin/Contact_p.asp") %>" data-validate>
                    <%= Auth.TokenField() %>
                    <input type="hidden" name="mode" value="reply">
                    <input type="hidden" name="id"   value="<%= ctId %>">
                    <textarea class="textarea" name="reply_content" rows="8" data-required data-label="답변"><%= AVH(ctArr, CT_REPLY, 0) %></textarea>
                    <p class="form-help">
                        여기에 저장한 답변은 관리자 화면 기록용입니다.
                        실제 회신은 <a href="mailto:<%= AVH(ctArr, CT_EMAIL, 0) %>?subject=RE: <%= UE(AV(ctArr, CT_SUBJECT, 0)) %>">이메일로 직접 발송</a>하세요.
                    </p>
                    <div class="btn-row between mt-3">
                        <a class="btn" href="<%= U("/admin/Contact.asp") %>">목록</a>
                        <span class="btn-row">
                            <a class="btn btn-danger" href="<%= U("/admin/Contact_p.asp") %>?mode=delete&amp;id=<%= ctId %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                               data-confirm="이 문의를 삭제하시겠습니까?">삭제</a>
                            <button type="submit" class="btn btn-primary">답변 저장</button>
                        </span>
                    </div>
                </form>
            </div>
<% End If %>

<%'==================== 목록 ====================%>
            <div class="admin-card">
                <form class="form-inline mb-3" method="get" action="<%= U("/admin/Contact.asp") %>">
                    <input class="input" type="search" name="sw" value="<%= HE(ctWord) %>" maxlength="50" placeholder="제목 / 이름 / 이메일">
                    <label class="check">
                        <input type="checkbox" name="unread" value="1" <%= IIfStr(ctUnread, "checked", "") %>>
                        <span>미확인만</span>
                    </label>
                    <button type="submit" class="btn btn-primary">검색</button>
                    <a class="btn" href="<%= U("/admin/Contact.asp") %>">초기화</a>
                </form>

                <h3>문의 목록 <span class="muted fs-sm">(<%= StringUtil.Comma(ctTotal) %>건)</span></h3>

<% If ctListCnt = 0 Then %>
                <div class="empty">접수된 문의가 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr><th class="num">ID</th><th class="ctr">상태</th><th>유형</th><th class="wide">제목</th><th>이름</th><th>이메일</th><th>접수일</th></tr>
                        </thead>
                        <tbody>
<%      Dim li
        For li = 0 To ctListCnt - 1 %>
                            <tr<%= IIfStr(AVI(ctListArr, CT_ID, li) = ctId, " style=""background:var(--c-primary-soft)""", "") %>>
                                <td class="num"><%= AVI(ctListArr, CT_ID, li) %></td>
                                <td class="ctr">
<%                                  If AVB(ctListArr, CT_ISREPLY, li) Then %><span class="badge badge-primary">답변</span>
                                    <% ElseIf AVB(ctListArr, CT_ISREAD, li) Then %><span class="badge">확인</span>
                                    <% Else %><span class="badge badge-new">신규</span><% End If %>
                                </td>
                                <td><%= AVH(ctListArr, CT_TYPE, li) %></td>
                                <td class="wide"><a href="<%= U("/admin/Contact.asp") %>?id=<%= AVI(ctListArr, CT_ID, li) %>" class="fw7"><%= AVH(ctListArr, CT_SUBJECT, li) %></a></td>
                                <td><%= AVH(ctListArr, CT_NAME, li) %></td>
                                <td><%= AVH(ctListArr, CT_EMAIL, li) %></td>
                                <td><%= HE(DateUtil.Fmt(AV(ctListArr, CT_REGDT, li), "ymdhi")) %></td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
                <%= Paging.Render() %>
<% End If %>
            </div>

        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
