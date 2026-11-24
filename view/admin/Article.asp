<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Article.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>게시글 관리</h2>
                <p class="desc">삭제된 글도 함께 표시됩니다. 논리 삭제된 글은 복구할 수 있습니다.</p>
            </div>

            <div class="admin-card">
                <form class="form-inline" method="get" action="<%= U("/admin/Article.asp") %>">
                    <select class="select" name="bid" style="flex:0 0 160px">
                        <option value="">전체 게시판</option>
<%      Dim bi
        For bi = 0 To aaBoardCnt - 1 %>
                        <option value="<%= AVH(aaBoards, BM_ID, bi) %>" <%= IIfStr(aaBid = AV(aaBoards, BM_ID, bi), "selected", "") %>><%= AVH(aaBoards, BM_NAME, bi) %></option>
<%      Next %>
                    </select>
                    <input class="input" type="search" name="sw" value="<%= HE(aaWord) %>" maxlength="50" placeholder="제목 검색">
                    <button type="submit" class="btn btn-primary">검색</button>
                    <a class="btn" href="<%= U("/admin/Article.asp") %>">초기화</a>
                </form>
            </div>

            <div class="admin-card">
                <h3>목록 <span class="muted fs-sm">(<%= StringUtil.Comma(aaTotal) %>건)</span></h3>

<% If aaCnt = 0 Then %>
                <div class="empty">게시글이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th class="num">ID</th><th>게시판</th><th class="wide">제목</th>
                                <th>작성자</th><th class="num">조회</th><th class="num">댓글</th>
                                <th class="ctr">상태</th><th>작성일</th><th>관리</th>
                            </tr>
                        </thead>
                        <tbody>
<%      Dim ai, isDel
        For ai = 0 To aaCnt - 1
            isDel = AVB(aaArr, 8, ai) %>
                            <tr>
                                <td class="num"><%= AVI(aaArr, 0, ai) %></td>
                                <td><span class="badge"><%= AVH(aaArr, 2, ai) %></span></td>
                                <td class="wide">
                                    <a href="<%= BoardViewUrl(AV(aaArr, 1, ai), AVI(aaArr, 0, ai)) %>" target="_blank"><%= AVH(aaArr, 3, ai) %></a>
<%                                  If AVB(aaArr, 7, ai) Then %> <span class="badge badge-notice">공지</span><% End If %>
                                </td>
                                <td><%= AVH(aaArr, 4, ai) %></td>
                                <td class="num"><%= AVI(aaArr, 5, ai) %></td>
                                <td class="num"><%= AVI(aaArr, 6, ai) %></td>
                                <td class="ctr">
<%                                  If isDel Then %><span class="badge badge-notice">삭제</span><% Else %><span class="badge badge-primary">정상</span><% End If %>
                                </td>
                                <td><%= HE(DateUtil.Fmt(AV(aaArr, 9, ai), "ymdhi")) %></td>
                                <td>
<% If isDel Then %>
                                    <a class="btn btn-sm" href="<%= U("/admin/Article_p.asp") %>?mode=restore&amp;id=<%= AVI(aaArr, 0, ai) %>&amp;_csrf=<%= UE(Auth.Token()) %>">복구</a>
                                    <a class="btn btn-sm btn-danger" href="<%= U("/admin/Article_p.asp") %>?mode=purge&amp;id=<%= AVI(aaArr, 0, ai) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="영구 삭제합니다. 댓글과 첨부파일도 함께 삭제되며 복구할 수 없습니다. 계속할까요?">영구삭제</a>
<% Else %>
                                    <a class="btn btn-sm" href="<%= U("/admin/Article_p.asp") %>?mode=notice&amp;id=<%= AVI(aaArr, 0, ai) %>&amp;_csrf=<%= UE(Auth.Token()) %>"><%= IIfStr(AVB(aaArr, 7, ai), "공지해제", "공지") %></a>
                                    <a class="btn btn-sm btn-danger" href="<%= U("/admin/Article_p.asp") %>?mode=delete&amp;id=<%= AVI(aaArr, 0, ai) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="이 게시물을 삭제하시겠습니까? (복구 가능)">삭제</a>
<% End If %>
                                </td>
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
