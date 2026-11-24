<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/board/View.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li><a href="<%= HE(backUrl) %>"><%= HE(bdName) %></a></li>
            <li>본문</li>
        </ul>
    </div>

    <article>
        <header class="view-head">
<% If Len(artCateName) > 0 Then %>
            <span class="cate"><%= HE(artCateName) %></span>
<% End If %>
            <h2>
<%              If artIsNotice Then %><span class="badge badge-notice">공지</span> <% End If %>
<%              If artIsSecret Then %><span class="ico" title="비밀글">🔒</span> <% End If %>
                <%= HE(artTitle) %>
            </h2>
            <div class="view-meta">
                <span class="writer"><%= HE(artWriterName) %></span>
                <span><%= HE(DateUtil.Fmt(artRegDt, "ymdhi")) %></span>
<% If IsDate(artUpdDt) Then %>
                <span>수정 <%= HE(DateUtil.Fmt(artUpdDt, "ymdhi")) %></span>
<% End If %>
                <span>조회 <%= StringUtil.Comma(artHit) %></span>
<% If bdUseComment Then %>
                <span>댓글 <%= artCmtCnt %></span>
<% End If %>
            </div>
        </header>

<%'============ 첨부파일 ============%>
<% If fileCnt > 0 Then %>
        <div class="attach">
            <div class="tit">첨부파일 <%= fileCnt %>개</div>
            <ul>
<%          Dim fi
            For fi = 0 To fileCnt - 1 %>
                <li>
                    <a href="<%= U("/board/Download.asp") %>?fid=<%= AVI(fileArr, BF_ID, fi) %>">
                        <span aria-hidden="true">📎</span>
                        <span><%= AVH(fileArr, BF_ORGNAME, fi) %></span>
                        <span class="size">(<%= HE(StringUtil.SizeStr(AVI(fileArr, BF_SIZE, fi))) %> · <%= AVI(fileArr, BF_DOWNCNT, fi) %>회)</span>
                    </a>
                </li>
<%          Next %>
            </ul>
        </div>
<% End If %>

<%'============ 본문 ============%>
        <div class="view-body">
            <%= StringUtil.SafeHtml(artContent) %>
        </div>

<% If Len(artLink) > 0 Then %>
        <div class="view-tags">
            <a class="btn btn-sm" href="<%= HE(artLink) %>" target="_blank" rel="noopener noreferrer">관련 링크 열기 ↗</a>
        </div>
<% End If %>

<% If Len(artTags) > 0 Then %>
        <div class="view-tags">
            <%= StringUtil.TagChips(artTags, U("/board/List.asp") & "?bid=" & UE(bdId) & "&sf=tags&sw={T}") %>
        </div>
<% End If %>
    </article>

<%'============ 버튼 ============%>
    <div class="btn-row between mt-4">
        <a class="btn" href="<%= HE(backUrl) %>">목록</a>
        <span class="btn-row">
<% If canReply Then %>
            <a class="btn" href="<%= HE(replyUrl) %>">답변</a>
<% End If %>
<% If canEdit Then %>
            <a class="btn" href="<%= HE(modifyUrl) %>">수정</a>
    <% If artWriterId = 0 And Not bdCanAdmin Then %>
            <%' 비회원 글은 비밀번호 확인을 거친다. 단 게시판 관리자는 바로 삭제한다.
              ' (관리자까지 PwCheck 로 보내면 남의 비밀번호를 모르니 방명록 스팸을 지울 수 없다) %>
            <a class="btn btn-danger" href="<%= U("/board/PwCheck.asp") %>?bid=<%= UE(bdId) %>&amp;id=<%= artId %>&amp;act=delete"
               title="비밀번호 확인 후 삭제됩니다">삭제</a>
    <% Else %>
            <a class="btn btn-danger" href="<%= HE(deleteUrl) %>" data-confirm="이 게시물을 삭제하시겠습니까?">삭제</a>
    <% End If %>
<% End If %>
        </span>
    </div>

<%'============ 이전 / 다음 ============%>
    <ul class="view-nav">
<% If IsArray(prevArr) Then %>
        <li>
            <span class="dir">이전글</span>
            <a class="tit" href="<%= BoardViewUrl(bdId, AVI(prevArr, 0, 0)) %>"><%= AVH(prevArr, 1, 0) %></a>
        </li>
<% End If %>
<% If IsArray(nextArr) Then %>
        <li>
            <span class="dir">다음글</span>
            <a class="tit" href="<%= BoardViewUrl(bdId, AVI(nextArr, 0, 0)) %>"><%= AVH(nextArr, 1, 0) %></a>
        </li>
<% End If %>
<% If Not IsArray(prevArr) And Not IsArray(nextArr) Then %>
        <li><span class="dir muted">-</span><span class="tit muted">이전/다음 글이 없습니다.</span></li>
<% End If %>
    </ul>


<%'==================================================================
  ' 댓글
  '==================================================================%>
<% If bdUseComment Then %>
    <section class="comment-area">
        <h3>댓글 <span class="num"><%= cmtCnt %></span></h3>

<% If cmtCnt > 0 Then %>
        <div class="comment-list">
<%      Dim mi2, mDepth, mWid, mCanDel
        For mi2 = 0 To cmtCnt - 1
            mDepth  = AVI(cmtArr, CMT_DEPTH, mi2)
            mWid    = AVI(cmtArr, CMT_WID, mi2)
            mCanDel = CanEditArticle(mWid, AVB(cmtArr, CMT_HASPW, mi2))
%>
            <div class="comment-item <%= IIfStr(mDepth > 0, "is-reply", "") %>">
                <div class="head">
                    <span class="name">
                        <%= AVH(cmtArr, CMT_WNAME, mi2) %>
<%                      If AVI(cmtArr, CMT_WLEVEL, mi2) >= LV_ADMIN Then %><span class="badge badge-admin">관리자</span><% End If %>
                    </span>
                    <span class="date"><%= HE(DateUtil.TimeAgo(AV(cmtArr, CMT_REGDT, mi2))) %></span>
                </div>
                <div class="body"><%= HE(AV(cmtArr, CMT_CONTENT, mi2)) %></div>

                <div class="act">
<%              If mDepth = 0 And bdCanComment Then %>
                    <button type="button" onclick="jkwReplyToggle(<%= AVI(cmtArr, CMT_ID, mi2) %>)">답글</button>
<%              End If %>
<%              If mCanDel Then %>
                    <button type="button" onclick="jkwCommentDelete(<%= AVI(cmtArr, CMT_ID, mi2) %>, <%= IIfStr(mWid = 0, "true", "false") %>)">삭제</button>
<%              End If %>
                </div>

<%              If mDepth = 0 And bdCanComment Then %>
                <div class="reply-form mt-3" id="reply-form-<%= AVI(cmtArr, CMT_ID, mi2) %>" style="display:none">
                    <form method="post" action="<%= U("/board/Comment_p.asp") %>" data-validate>
                        <%= Auth.TokenField() %>
                        <input type="hidden" name="mode"      value="write">
                        <input type="hidden" name="bid"       value="<%= HE(bdId) %>">
                        <input type="hidden" name="id"        value="<%= artId %>">
                        <input type="hidden" name="parent_id" value="<%= AVI(cmtArr, CMT_ID, mi2) %>">
<%                      If Not GLB_IS_LOGIN Then %>
                        <div class="form-inline mb-3">
                            <input type="text"     class="input" name="writer_name" placeholder="이름" maxlength="20" data-required data-label="이름">
                            <input type="password" class="input" name="writer_pw"   placeholder="비밀번호" maxlength="20" data-required data-label="비밀번호">
                        </div>
<%                      End If %>
                        <textarea class="textarea" name="content" rows="2" maxlength="1000"
                                  placeholder="답글을 입력하세요." data-required data-label="답글"></textarea>
                        <div class="btn-row right mt-2">
                            <button type="submit" class="btn btn-sm btn-primary">답글 등록</button>
                        </div>
                    </form>
                </div>
<%              End If %>
            </div>
<%      Next %>
        </div>
<% End If %>

<% If bdCanComment Then %>
        <form class="comment-form" method="post" action="<%= U("/board/Comment_p.asp") %>" data-validate>
            <%= Auth.TokenField() %>
            <input type="hidden" name="mode" value="write">
            <input type="hidden" name="bid"  value="<%= HE(bdId) %>">
            <input type="hidden" name="id"   value="<%= artId %>">
            <input type="hidden" name="parent_id" value="0">

<% If Not GLB_IS_LOGIN Then %>
            <div class="form-inline mb-3">
                <input type="text"     class="input" name="writer_name" placeholder="이름" maxlength="20" data-required data-label="이름">
                <input type="password" class="input" name="writer_pw"   placeholder="비밀번호 (삭제용)" maxlength="20" data-required data-label="비밀번호">
            </div>
<% End If %>

            <label class="sr-only" for="cmt-content">댓글 내용</label>
            <textarea class="textarea" id="cmt-content" name="content" rows="3" maxlength="1000"
                      placeholder="댓글을 입력하세요." data-required data-label="댓글" data-counter="cmt-cnt"></textarea>
            <div class="form-help tr"><span id="cmt-cnt">0</span> / 1000</div>

            <div class="btn-row right mt-2">
                <button type="submit" class="btn btn-primary">댓글 등록</button>
            </div>
        </form>
<% Else %>
        <div class="alert alert-info">
            댓글을 남기려면 <a href="<%= U("/member/Login.asp") %>?rurl=<%= UE(GLB_CUR_URL) %>" style="text-decoration:underline">로그인</a>이 필요합니다.
        </div>
<% End If %>
    </section>
<% End If %>

</div>

<%
g_extraScript = "<script>" & _
"function jkwCommentDelete(cid, isGuest){" & _
"  if(isGuest){" & _
"    var pw = prompt('댓글 작성 시 입력한 비밀번호를 입력하세요.');" & _
"    if(pw===null||pw==='') return;" & _
"    location.href='" & JE(U("/board/Comment_p.asp")) & "?mode=delete&bid=" & JE(bdId) & "&id=" & artId & "&cid='+cid+'&pw='+encodeURIComponent(pw)+'&_csrf=" & JE(Auth.Token()) & "';" & _
"  }else{" & _
"    if(!confirm('댓글을 삭제하시겠습니까?')) return;" & _
"    location.href='" & JE(U("/board/Comment_p.asp")) & "?mode=delete&bid=" & JE(bdId) & "&id=" & artId & "&cid='+cid+'&_csrf=" & JE(Auth.Token()) & "';" & _
"  }" & _
"}" & _
"</script>"
%>
<!--#include file="../include/Footer.asp"-->
