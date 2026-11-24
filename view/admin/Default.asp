<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>대시보드</h2>
                <p class="desc"><%= HE(DateUtil.Fmt(Now(), "ymdhi")) %> 기준</p>
            </div>

            <div class="dash-grid">
                <div class="dash-card">
                    <div class="lbl">전체 회원</div>
                    <div class="num"><%= StringUtil.Comma(dsMemberT) %></div>
                    <div class="sub">오늘 +<%= dsMemberD %></div>
                </div>
                <div class="dash-card">
                    <div class="lbl">전체 게시글</div>
                    <div class="num"><%= StringUtil.Comma(dsArticleT) %></div>
                    <div class="sub">오늘 +<%= dsArticleD %></div>
                </div>
                <div class="dash-card">
                    <div class="lbl">오늘 방문</div>
                    <div class="num"><%= StringUtil.Comma(dsVisitD) %></div>
                    <div class="sub">누적 <%= StringUtil.Comma(dsVisitT) %></div>
                </div>
                <div class="dash-card">
                    <div class="lbl">미확인 문의</div>
                    <div class="num" style="<%= IIfStr(dsUnread > 0, "color:var(--c-danger)", "") %>"><%= StringUtil.Comma(dsUnread) %></div>
                    <div class="sub"><a href="<%= U("/admin/Contact.asp") %>">확인하기 →</a></div>
                </div>
            </div>

            <%'--- 최근 7일 방문 ---%>
            <div class="admin-card">
                <h3>최근 7일 방문</h3>
<% If vsCnt = 0 Then %>
                <div class="empty">방문 기록이 없습니다.</div>
<% Else %>
                <div style="display:flex;align-items:flex-end;gap:8px;height:140px;padding-top:8px">
<%      Dim gi
        For gi = vsCnt - 1 To 0 Step -1 %>
                    <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;height:100%">
                        <div style="flex:1;display:flex;align-items:flex-end;width:100%">
                            <div title="<%= AVI(vsArr, 1, gi) %> PV / <%= AVI(vsArr, 2, gi) %> UV"
                                 style="width:100%;background:var(--c-primary);border-radius:4px 4px 0 0;height:<%= Int(AVI(vsArr, 1, gi) * 100 / vsMax) %>%;min-height:3px"></div>
                        </div>
                        <span class="fs-xs muted"><%= HE(Right(AV(vsArr, 0, gi), 5)) %></span>
                        <span class="fs-xs fw7"><%= AVI(vsArr, 1, gi) %></span>
                    </div>
<%      Next %>
                </div>
<% End If %>
            </div>

            <%'--- 게시판 현황 ---%>
            <div class="admin-card">
                <h3>게시판 현황</h3>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>ID</th><th>게시판</th><th>스킨</th><th class="ctr">사용</th><th class="ctr">쓰기권한</th><th></th>
                            </tr>
                        </thead>
                        <tbody>
<%      Dim bi
        For bi = 0 To bdListCnt - 1 %>
                            <tr>
                                <td><code><%= AVH(bdListArr, BM_ID, bi) %></code></td>
                                <td class="fw7"><%= AVH(bdListArr, BM_NAME, bi) %></td>
                                <td><span class="badge"><%= AVH(bdListArr, BM_SKIN, bi) %></span></td>
                                <td class="ctr"><%= IIfStr(AVB(bdListArr, BM_IS_USE, bi), "O", "-") %></td>
                                <td class="ctr">Lv.<%= AVI(bdListArr, BM_AUTH_WRITE, bi) %></td>
                                <td>
                                    <a class="btn btn-sm" href="<%= U("/board/List.asp") %>?bid=<%= UE(AV(bdListArr, BM_ID, bi)) %>">보기</a>
                                    <a class="btn btn-sm" href="<%= U("/admin/Board.asp") %>?bid=<%= UE(AV(bdListArr, BM_ID, bi)) %>">설정</a>
                                </td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
            </div>

            <%'--- 최근 게시글 ---%>
            <div class="admin-card">
                <h3>최근 게시글</h3>
<% If recArtCnt = 0 Then %>
                <div class="empty">게시글이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead><tr><th>게시판</th><th class="wide">제목</th><th>작성자</th><th class="num">조회</th><th>작성일</th></tr></thead>
                        <tbody>
<%      Dim ri
        For ri = 0 To recArtCnt - 1 %>
                            <tr>
                                <td><span class="badge"><%= AVH(recArtArr, 2, ri) %></span></td>
                                <td class="wide">
                                    <a href="<%= BoardViewUrl(AV(recArtArr, 1, ri), AVI(recArtArr, 0, ri)) %>"><%= AVH(recArtArr, 3, ri) %></a>
<%                                  If AVB(recArtArr, 8, ri) Then %> <span class="badge badge-notice">삭제됨</span><% End If %>
                                </td>
                                <td><%= AVH(recArtArr, 4, ri) %></td>
                                <td class="num"><%= AVI(recArtArr, 5, ri) %></td>
                                <td><%= HE(DateUtil.Fmt(AV(recArtArr, 9, ri), "ymdhi")) %></td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
<% End If %>
            </div>

            <%'--- 최근 댓글 / 문의 ---%>
            <div class="admin-card">
                <h3>최근 댓글</h3>
<% If recCmtCnt = 0 Then %>
                <div class="empty">댓글이 없습니다.</div>
<% Else %>
                <div class="comment-list">
<%      Dim mi
        For mi = 0 To recCmtCnt - 1 %>
                    <div class="comment-item">
                        <div class="head">
                            <span class="name"><%= AVH(recCmtArr, 4, mi) %></span>
                            <span class="date"><%= HE(DateUtil.TimeAgo(AV(recCmtArr, 5, mi))) %></span>
                            <a class="fs-xs" href="<%= BoardViewUrl(AV(recCmtArr, 1, mi), AVI(recCmtArr, 2, mi)) %>">[<%= AVH(recCmtArr, 6, mi) %>]</a>
                        </div>
                        <div class="body"><%= HE(StringUtil.Cut(AV(recCmtArr, 3, mi), 100)) %></div>
                    </div>
<%      Next %>
                </div>
<% End If %>
            </div>

            <div class="admin-card">
                <h3>최근 문의</h3>
<% If recCtCnt = 0 Then %>
                <div class="empty">문의가 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead><tr><th class="ctr">상태</th><th>유형</th><th class="wide">제목</th><th>이름</th><th>접수일</th></tr></thead>
                        <tbody>
<%      Dim ti
        For ti = 0 To recCtCnt - 1 %>
                            <tr>
                                <td class="ctr">
<%                                  If AVB(recCtArr, CT_ISREPLY, ti) Then %><span class="badge badge-primary">답변</span>
                                    <% ElseIf AVB(recCtArr, CT_ISREAD, ti) Then %><span class="badge">확인</span>
                                    <% Else %><span class="badge badge-new">신규</span><% End If %>
                                </td>
                                <td><%= AVH(recCtArr, CT_TYPE, ti) %></td>
                                <td class="wide"><a href="<%= U("/admin/Contact.asp") %>?id=<%= AVI(recCtArr, CT_ID, ti) %>"><%= AVH(recCtArr, CT_SUBJECT, ti) %></a></td>
                                <td><%= AVH(recCtArr, CT_NAME, ti) %></td>
                                <td><%= HE(DateUtil.Fmt(AV(recCtArr, CT_REGDT, ti), "ymdhi")) %></td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
<% End If %>
            </div>

        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
