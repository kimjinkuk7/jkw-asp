<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Member.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>회원 관리</h2>
                <p class="desc">등급 변경과 계정 잠금 해제를 처리합니다.</p>
            </div>

            <div class="admin-card">
                <form class="form-inline" method="get" action="<%= U("/admin/Member.asp") %>">
                    <select class="select" name="lv" style="flex:0 0 150px">
                        <option value="-1" <%= IIfStr(amLevel = -1, "selected", "") %>>전체 등급</option>
                        <option value="9"  <%= IIfStr(amLevel = 9,  "selected", "") %>>최고관리자</option>
                        <option value="5"  <%= IIfStr(amLevel = 5,  "selected", "") %>>운영자</option>
                        <option value="3"  <%= IIfStr(amLevel = 3,  "selected", "") %>>일반회원</option>
                        <option value="0"  <%= IIfStr(amLevel = 0,  "selected", "") %>>정지</option>
                    </select>
                    <input class="input" type="search" name="sw" value="<%= HE(amWord) %>" maxlength="50" placeholder="아이디 / 이름 / 닉네임 / 이메일">
                    <button type="submit" class="btn btn-primary">검색</button>
                    <a class="btn" href="<%= U("/admin/Member.asp") %>">초기화</a>
                </form>
            </div>

            <div class="admin-card">
                <h3>회원 목록 <span class="muted fs-sm">(<%= StringUtil.Comma(amTotal) %>명)</span></h3>

<% If amCnt = 0 Then %>
                <div class="empty">회원이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th class="num">ID</th><th>아이디</th><th>이름</th><th>닉네임</th><th>이메일</th>
                                <th class="ctr">등급</th><th class="num">로그인</th><th>최종접속</th><th>가입일</th><th>관리</th>
                            </tr>
                        </thead>
                        <tbody>
<%      Dim mi
        For mi = 0 To amCnt - 1 %>
                            <tr>
                                <td class="num"><%= AVI(amArr, MB_ID, mi) %></td>
                                <td class="fw7"><%= AVH(amArr, MB_LOGINID, mi) %></td>
                                <td><%= AVH(amArr, MB_NAME, mi) %></td>
                                <td><%= AVH(amArr, MB_NICK, mi) %></td>
                                <td><%= AVH(amArr, MB_EMAIL, mi) %></td>
                                <td class="ctr">
                                    <span class="badge <%= IIfStr(AVI(amArr, MB_LEVEL, mi) >= LV_ADMIN, "badge-admin", "badge-primary") %>"><%= HE(LvName(AVI(amArr, MB_LEVEL, mi))) %></span>
<%                                  If Not AVB(amArr, MB_ISUSE, mi) Then %><span class="badge badge-notice">중지</span><% End If %>
                                </td>
                                <td class="num"><%= AVI(amArr, MB_LOGINCNT, mi) %></td>
                                <td><%= HE(DateUtil.Fmt(AV(amArr, MB_LASTLOGIN, mi), "ymdhi")) %></td>
                                <td><%= HE(DateUtil.Fmt(AV(amArr, MB_REGDT, mi), "ymd")) %></td>
                                <td>
<% If AVI(amArr, MB_ID, mi) = GLB_MY_ID Then %>
                                    <span class="fs-xs muted">본인</span>
<% Else %>
                                    <form method="post" action="<%= U("/admin/Member_p.asp") %>" style="display:flex;gap:4px;align-items:center">
                                        <%= Auth.TokenField() %>
                                        <input type="hidden" name="mode" value="level">
                                        <input type="hidden" name="id" value="<%= AVI(amArr, MB_ID, mi) %>">
                                        <select class="select" name="member_level" style="min-height:34px;font-size:12px;padding:4px 24px 4px 8px">
                                            <option value="9" <%= IIfStr(AVI(amArr, MB_LEVEL, mi) = 9, "selected", "") %>>관리자</option>
                                            <option value="5" <%= IIfStr(AVI(amArr, MB_LEVEL, mi) = 5, "selected", "") %>>운영자</option>
                                            <option value="3" <%= IIfStr(AVI(amArr, MB_LEVEL, mi) = 3, "selected", "") %>>회원</option>
                                            <option value="0" <%= IIfStr(AVI(amArr, MB_LEVEL, mi) = 0, "selected", "") %>>정지</option>
                                        </select>
                                        <label class="check" style="min-height:auto">
                                            <input type="checkbox" name="is_use" value="1" <%= IIfStr(AVB(amArr, MB_ISUSE, mi), "checked", "") %>>
                                            <span class="fs-xs">사용</span>
                                        </label>
                                        <button type="submit" class="btn btn-sm btn-primary">적용</button>
                                    </form>
                                    <a class="btn btn-sm" href="<%= U("/admin/Member_p.asp") %>?mode=unlock&amp;id=<%= AVI(amArr, MB_ID, mi) %>&amp;_csrf=<%= UE(Auth.Token()) %>">잠금해제</a>
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
