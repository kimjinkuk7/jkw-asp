<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Portfolio.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>포트폴리오 관리</h2>
                <p class="desc">메인 노출 여부와 정렬 순서를 여기서 조정합니다.</p>
            </div>

            <div class="btn-row right mb-4">
                <a class="btn btn-primary" href="<%= U("/admin/PortfolioForm.asp") %>?id=0">＋ 새 프로젝트</a>
            </div>

            <div class="admin-card">
                <h3>목록 <span class="muted fs-sm">(<%= apCnt %>건)</span></h3>

<% If apCnt = 0 Then %>
                <div class="empty">등록된 프로젝트가 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th class="num">ID</th><th class="num">정렬</th><th class="wide">제목</th>
                                <th>카테고리</th><th>기간</th><th class="ctr">메인</th><th class="ctr">사용</th>
                                <th class="num">조회</th><th>관리</th>
                            </tr>
                        </thead>
                        <tbody>
<%      Dim pi
        For pi = 0 To apCnt - 1 %>
                            <tr>
                                <td class="num"><%= AVI(apArr, PF_ID, pi) %></td>
                                <td class="num"><%= AVI(apArr, PFA_SORTNO, pi) %></td>
                                <td class="wide">
                                    <a href="<%= U("/admin/PortfolioForm.asp") %>?id=<%= AVI(apArr, PF_ID, pi) %>" class="fw7"><%= AVH(apArr, PF_TITLE, pi) %></a>
                                    <div class="fs-xs muted"><%= AVH(apArr, PF_SUBTITLE, pi) %></div>
                                </td>
                                <td><span class="badge"><%= AVH(apArr, PF_CATEGORY, pi) %></span></td>
                                <td class="fs-xs"><%= HE(DateUtil.Period(AV(apArr, PF_START, pi), AV(apArr, PF_END, pi), AVB(apArr, PF_ONGOING, pi))) %></td>
                                <td class="ctr"><%= IIfStr(AVB(apArr, PF_ISMAIN, pi), "O", "-") %></td>
                                <td class="ctr"><%= IIfStr(AVB(apArr, PFA_ISUSE, pi), "O", "X") %></td>
                                <td class="num"><%= AVI(apArr, PF_HIT, pi) %></td>
                                <td>
                                    <a class="btn btn-sm btn-primary" href="<%= U("/admin/PortfolioForm.asp") %>?id=<%= AVI(apArr, PF_ID, pi) %>">수정</a>
                                    <a class="btn btn-sm" href="<%= U("/portfolio/View.asp") %>?id=<%= AVI(apArr, PF_ID, pi) %>" target="_blank">보기</a>
                                    <a class="btn btn-sm btn-danger"
                                       href="<%= U("/admin/Portfolio_p.asp") %>?mode=delete&amp;id=<%= AVI(apArr, PF_ID, pi) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="삭제하시겠습니까? 복구할 수 없습니다.">삭제</a>
                                </td>
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
