<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/board/List.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li><%= HE(bdName) %></li>
        </ul>
        <h2><%= HE(bdName) %></h2>
<% If Len(bdDesc) > 0 Then %>
        <p class="desc"><%= HE(bdDesc) %></p>
<% End If %>
    </div>

<%'============ 분류 탭 ============%>
<% If cateCnt > 0 Then %>
    <nav class="cate-tabs" aria-label="분류">
        <a href="<%= U("/board/List.asp") %>?bid=<%= UE(bdId) %>" class="<%= IIfStr(cateId = 0, "is-on", "") %>">전체</a>
<%      Dim ci
        For ci = 0 To cateCnt - 1 %>
        <a href="<%= U("/board/List.asp") %>?bid=<%= UE(bdId) %>&amp;cid=<%= AVI(cateArr, CAT_ID, ci) %>"
           class="<%= IIfStr(cateId = AVI(cateArr, CAT_ID, ci), "is-on", "") %>"><%= AVH(cateArr, CAT_NAME, ci) %></a>
<%      Next %>
    </nav>
<% End If %>

    <div class="board-head">
        <span class="count"><%= Paging.Summary() %></span>
<% If bdCanWrite And bdSkin <> "guest" Then %>
        <a class="btn btn-primary btn-sm" href="<%= HE(writeUrl) %>">글쓰기</a>
<% End If %>
    </div>

<% If isSearching Then %>
    <div class="alert alert-info">
        <strong>&ldquo;<%= HE(srchWord) %>&rdquo;</strong> 검색 결과 <%= StringUtil.Comma(totalCount) %>건
        <a href="<%= U("/board/List.asp") %>?bid=<%= UE(bdId) %>" style="margin-left:8px;text-decoration:underline">검색 초기화</a>
    </div>
<% End If %>


<%'==================================================================
  ' SKIN : list  (공지 · 자유게시판 · 자료실)
  ' 데스크톱 = 표 / 모바일 = 카드. HTML 은 한 벌, CSS 로만 전환한다.
  '==================================================================%>
<% If bdSkin = "list" Then %>

    <table class="board-list">
        <caption class="sr-only"><%= HE(bdName) %> 목록</caption>
        <thead>
            <tr>
                <th scope="col" class="col-no">번호</th>
<% If cateCnt > 0 Then %>   <th scope="col" class="col-cate">분류</th><% End If %>
                <th scope="col" class="col-title">제목</th>
                <th scope="col" class="col-writer">작성자</th>
                <th scope="col" class="col-date">작성일</th>
                <th scope="col" class="col-hit">조회</th>
            </tr>
        </thead>
        <tbody>
<%
    '--- rowCnt 라는 이름은 쓸 수 없다. 전역 함수 RowCnt() 와 충돌한다(800a0411) ---
    Dim pass, rowArr, rCnt, ri, isNoticeRow, dp
    For pass = 0 To 1
        If pass = 0 Then
            rowArr = noticeArr : rCnt = noticeCnt : isNoticeRow = True
        Else
            rowArr = listArr   : rCnt = listCnt   : isNoticeRow = False
        End If

        For ri = 0 To rCnt - 1
%>
            <tr class="<%= IIfStr(isNoticeRow, "is-notice", "") %>">
                <td class="col-no">
<%              If isNoticeRow Then %><span class="badge badge-notice">공지</span>
<%              ElseIf AVI(rowArr, ART_DEPTH, ri) > 0 Then %>-<%
                ' 답변글은 번호를 표시하지 않는다. 번호는 "몇 번째 원글인가"를 뜻하고
                ' 답변글은 원글에 딸린 것이므로, 번호를 주면 원글 번호가 어긋나 보인다.
                ' 계층은 제목 앞 ↳ 들여쓰기로 이미 드러난다.
                %>
<%              Else %><%= AVI(rowArr, ART_NO, ri) %><% End If %>
                </td>
<% If cateCnt > 0 Then %>
                <td class="col-cate"><%= AVH(rowArr, ART_CNAME, ri) %></td>
<% End If %>
                <td class="col-title">
<%              dp = AVI(rowArr, ART_DEPTH, ri)
                If dp > 0 Then %><span class="reply-arrow" aria-hidden="true"><%= String(dp, "-") %>&nbsp;↳</span><% End If %>
                    <a href="<%= BoardViewUrl(bdId, AVI(rowArr, ART_ID, ri)) %><%= ListStateQs() %>">
<%                  If AVB(rowArr, ART_SECRET, ri) Then %><span class="ico" title="비밀글">🔒</span><% End If %>
                        <%= Highlighted(AVH(rowArr, ART_TITLE, ri)) %>
                    </a>
<%              If AVI(rowArr, ART_CMT, ri) > 0 Then %><span class="cmt">[<%= AVI(rowArr, ART_CMT, ri) %>]</span><% End If %>
<%              If AVI(rowArr, ART_FILE, ri) > 0 Then %><span class="ico" title="첨부파일">📎</span><% End If %>
<%              If DateUtil.IsNew(AV(rowArr, ART_REGDT, ri), 24) Then %><span class="badge badge-new">N</span><% End If %>

                    <div class="meta">
                        <span class="writer"><%= AVH(rowArr, ART_WNAME, ri) %></span>
                        <span><%= HE(DateUtil.ListDate(AV(rowArr, ART_REGDT, ri))) %></span>
                        <span>조회 <%= AVI(rowArr, ART_HIT, ri) %></span>
                    </div>
                </td>
                <td class="col-writer"><%= AVH(rowArr, ART_WNAME, ri) %></td>
                <td class="col-date"><%= HE(DateUtil.ListDate(AV(rowArr, ART_REGDT, ri))) %></td>
                <td class="col-hit"><%= AVI(rowArr, ART_HIT, ri) %></td>
            </tr>
<%
        Next
    Next

    If (noticeCnt + listCnt) = 0 Then
%>
            <tr>
                <td colspan="6" style="padding:0;border:0">
                    <div class="empty"><%= IIfStr(isSearching, "검색 결과가 없습니다.", "등록된 게시물이 없습니다.") %></div>
                </td>
            </tr>
<%  End If %>
        </tbody>
    </table>


<%'==================================================================
  ' SKIN : gallery
  '==================================================================%>
<% ElseIf bdSkin = "gallery" Then %>

<% If listCnt = 0 Then %>
    <div class="empty"><%= IIfStr(isSearching, "검색 결과가 없습니다.", "등록된 게시물이 없습니다.") %></div>
<% Else %>
    <div class="gallery-grid">
<%  Dim gi2, gth
    For gi2 = 0 To listCnt - 1
        gth = UpUrl(AV(listArr, ART_THUMB, gi2)) %>
        <a class="gallery-item" href="<%= BoardViewUrl(bdId, AVI(listArr, ART_ID, gi2)) %><%= ListStateQs() %>">
            <span class="thumb"<% If Len(gth) > 0 Then %> style="background-image:url('<%= HE(gth) %>')"<% End If %>>
<%          If Len(gth) = 0 Then %><span class="no-img">이미지 없음</span><% End If %>
            </span>
            <span class="tit"><%= AVH(listArr, ART_TITLE, gi2) %></span>
            <span class="meta">
                <%= AVH(listArr, ART_WNAME, gi2) %> · <%= HE(DateUtil.ListDate(AV(listArr, ART_REGDT, gi2))) %>
<%              If AVI(listArr, ART_CMT, gi2) > 0 Then %> · 댓글 <%= AVI(listArr, ART_CMT, gi2) %><% End If %>
            </span>
        </a>
<%  Next %>
    </div>
<% End If %>


<%'==================================================================
  ' SKIN : blog
  '==================================================================%>
<% ElseIf bdSkin = "blog" Then %>

<% If tagCnt > 0 Then %>
    <div class="mb-4">
<%      Dim ti
        For ti = 0 To tagCnt - 1 %>
        <a class="chip" href="<%= U("/board/List.asp") %>?bid=<%= UE(bdId) %>&amp;sf=tags&amp;sw=<%= UE(AV(tagArr, 0, ti)) %>">#<%= AVH(tagArr, 0, ti) %> <%= AVI(tagArr, 1, ti) %></a>
<%      Next %>
    </div>
<% End If %>

<% If listCnt = 0 Then %>
    <div class="empty"><%= IIfStr(isSearching, "검색 결과가 없습니다.", "등록된 글이 없습니다.") %></div>
<% Else %>
    <div class="blog-list">
<%  Dim bi2, bth
    For bi2 = 0 To listCnt - 1
        bth = UpUrl(AV(listArr, ART_THUMB, bi2)) %>
        <article class="blog-item">
<%      If Len(bth) > 0 Then %>
            <a class="thumb" href="<%= BoardViewUrl(bdId, AVI(listArr, ART_ID, bi2)) %>" style="background-image:url('<%= HE(bth) %>')" aria-hidden="true"></a>
<%      End If %>
            <div>
<%          If Len(AV(listArr, ART_CNAME, bi2)) > 0 Then %>
                <span class="cate"><%= AVH(listArr, ART_CNAME, bi2) %></span>
<%          End If %>
                <h3><a href="<%= BoardViewUrl(bdId, AVI(listArr, ART_ID, bi2)) %>"><%= Highlighted(AVH(listArr, ART_TITLE, bi2)) %></a></h3>
                <p class="summary"><%= HE(StringUtil.Cut(StringUtil.StripTags(AV(listArr, ART_SUMMARY, bi2)), 140)) %></p>
                <div class="meta">
                    <span><%= HE(DateUtil.Fmt(AV(listArr, ART_REGDT, bi2), "ymd")) %></span>
                    <span>조회 <%= AVI(listArr, ART_HIT, bi2) %></span>
<%                  If AVI(listArr, ART_CMT, bi2) > 0 Then %><span>댓글 <%= AVI(listArr, ART_CMT, bi2) %></span><% End If %>
                </div>
<%          If Len(AV(listArr, ART_TAGS, bi2)) > 0 Then %>
                <div class="mt-2"><%= StringUtil.TagChips(AV(listArr, ART_TAGS, bi2), U("/board/List.asp") & "?bid=" & UE(bdId) & "&sf=tags&sw={T}") %></div>
<%          End If %>
            </div>
        </article>
<%  Next %>
    </div>
<% End If %>


<%'==================================================================
  ' SKIN : guest  (방명록 - 목록 위에 바로 작성 폼)
  '==================================================================%>
<% Else %>

<% If bdCanWrite Then %>
    <form class="guest-write" method="post" action="<%= U("/board/Write_p.asp") %>?bid=<%= UE(bdId) %>" data-validate>
        <%= Auth.TokenField() %>
        <input type="hidden" name="bid"  value="<%= HE(bdId) %>">
        <input type="hidden" name="mode" value="write">
        <input type="hidden" name="title" value="방명록">

<% If Not GLB_IS_LOGIN Then %>
        <div class="form-inline mb-3">
            <input type="text"     class="input" name="writer_name" placeholder="이름" maxlength="20" data-required data-label="이름">
            <input type="password" class="input" name="writer_pw"   placeholder="비밀번호 (수정·삭제용)" maxlength="20" data-required data-label="비밀번호">
        </div>
<% End If %>

        <div class="form-group mb-3">
            <label class="sr-only" for="gb-content">방명록 내용</label>
            <textarea class="textarea" id="gb-content" name="content" rows="4" maxlength="1000"
                      placeholder="다녀가신 흔적을 남겨주세요." data-required data-label="내용" data-counter="gb-cnt"></textarea>
            <div class="form-help tr"><span id="gb-cnt">0</span> / 1000</div>
        </div>

        <div class="btn-row right">
            <button type="submit" class="btn btn-primary">남기기</button>
        </div>
    </form>
<% Else %>
    <div class="alert alert-info">방명록 작성 권한이 없습니다.</div>
<% End If %>

<% If listCnt = 0 Then %>
    <div class="empty">첫 번째 방명록을 남겨보세요.</div>
<% Else %>
    <div class="guest-list">
<%  Dim ui
    For ui = 0 To listCnt - 1 %>
        <article class="guest-item">
            <div class="head">
                <span class="name"><%= AVH(listArr, ART_WNAME, ui) %></span>
                <span class="date"><%= HE(DateUtil.Fmt(AV(listArr, ART_REGDT, ui), "ymdhi")) %></span>
            </div>
            <div class="body"><%= HE(StringUtil.StripTags(AV(listArr, ART_SUMMARY, ui))) %></div>
<%          If CanEditArticle(AVI(listArr, ART_WID, ui), True) Then %>
            <div class="btn-row right mt-2">
                <a class="btn btn-sm" href="<%= U("/board/View.asp") %>?bid=<%= UE(bdId) %>&amp;id=<%= AVI(listArr, ART_ID, ui) %>">관리</a>
            </div>
<%          End If %>
        </article>
<%  Next %>
    </div>
<% End If %>

<% End If %>
<%'================= /SKIN =================%>


<%'============ 페이징 ============%>
<% If bdSkin = "blog" Or bdSkin = "guest" Then %>
    <%= Paging.Render() %>
<% Else %>
    <%= Paging.Render() %>
<% End If %>

<%'============ 검색 ============%>
    <form class="search-box" method="get" action="<%= U("/board/List.asp") %>">
        <input type="hidden" name="bid" value="<%= HE(bdId) %>">
<% If cateId > 0 Then %><input type="hidden" name="cid" value="<%= cateId %>"><% End If %>
        <label class="sr-only" for="sf">검색 조건</label>
        <select class="select" id="sf" name="sf">
            <option value="all"     <%= IIfStr(srchField = "all",     "selected", "") %>>제목+내용</option>
            <option value="title"   <%= IIfStr(srchField = "title",   "selected", "") %>>제목</option>
            <option value="content" <%= IIfStr(srchField = "content", "selected", "") %>>내용</option>
            <option value="writer"  <%= IIfStr(srchField = "writer",  "selected", "") %>>작성자</option>
<% If bdSkin = "blog" Then %>
            <option value="tags"    <%= IIfStr(srchField = "tags",    "selected", "") %>>태그</option>
<% End If %>
        </select>
        <label class="sr-only" for="sw">검색어</label>
        <input class="input" type="search" id="sw" name="sw" value="<%= HE(srchWord) %>" maxlength="50" placeholder="검색어를 입력하세요">
        <button type="submit" class="btn btn-primary">검색</button>
    </form>

<% If bdCanWrite And bdSkin <> "guest" Then %>
    <div class="btn-row right mt-4">
        <a class="btn btn-primary" href="<%= HE(writeUrl) %>">글쓰기</a>
    </div>
<% End If %>

</div>

<%
'--- 검색어 하이라이트 헬퍼 (이미 HtmlEnc 된 문자열에 적용) ---
Function Highlighted(ByVal pEscaped)
    If isSearching And (srchField = "title" Or srchField = "all") Then
        Highlighted = StringUtil.Highlight(pEscaped, srchWord)
    Else
        Highlighted = pEscaped
    End If
End Function
%>
<!--#include file="../include/Footer.asp"-->
