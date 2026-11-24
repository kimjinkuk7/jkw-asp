<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/board/Write.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li><a href="<%= HE(backUrl) %>"><%= HE(bdName) %></a></li>
            <li><%= HE(modeLabel) %></li>
        </ul>
        <h2><%= HE(bdName) %> <%= HE(modeLabel) %></h2>
    </div>

    <%' bid 를 action 의 QueryString 에 둔다 - 파일 업로드(BinaryRead) 전에 Form 을 읽지 않기 위함 %>
    <form method="post" action="<%= U("/board/Write_p.asp") %>?bid=<%= UE(bdId) %>"
          enctype="multipart/form-data" data-validate>

        <%= Auth.TokenField() %>
        <input type="hidden" name="bid"  value="<%= HE(bdId) %>">
        <input type="hidden" name="mode" value="<%= HE(wMode) %>">
        <input type="hidden" name="id"   value="<%= artId %>">
        <input type="hidden" name="page" value="<%= HE(RequestUtil.ValueInit("page", "1")) %>">
        <input type="hidden" name="sf"   value="<%= HE(RequestUtil.ValueInit("sf", "")) %>">
        <input type="hidden" name="sw"   value="<%= HE(RequestUtil.ValueInit("sw", "")) %>">

        <div class="form-table">

<%'============ 분류 ============%>
<% If cateCnt > 0 Then %>
            <div class="row">
                <label class="form-label" for="f-cid">분류</label>
                <div>
                    <select class="select" id="f-cid" name="category_id">
                        <option value="0">선택하세요</option>
<%                      Dim ci
                        For ci = 0 To cateCnt - 1 %>
                        <option value="<%= AVI(cateArr, CAT_ID, ci) %>" <%= IIfStr(fCateId = AVI(cateArr, CAT_ID, ci), "selected", "") %>><%= AVH(cateArr, CAT_NAME, ci) %></option>
<%                      Next %>
                    </select>
                </div>
            </div>
<% End If %>

<%'============ 작성자 (비회원) ============%>
<% If Not GLB_IS_LOGIN Then %>
            <div class="row">
                <label class="form-label" for="f-wname">작성자<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="f-wname" name="writer_name"
                           value="<%= HE(fWriterName) %>" maxlength="20"
                           data-required data-label="작성자">
                </div>
            </div>
            <div class="row">
                <label class="form-label" for="f-wpw">비밀번호<span class="req">*</span></label>
                <div>
                    <input class="input" type="password" id="f-wpw" name="writer_pw"
                           maxlength="20" autocomplete="new-password"
                           <%= IIfStr(wMode <> "modify", "data-required data-label=""비밀번호""", "") %>>
                    <p class="form-help">글 수정·삭제 시 필요합니다.<%= IIfStr(wMode = "modify", " 변경하지 않으려면 비워두세요.", "") %></p>
                </div>
            </div>
<% End If %>

<%'============ 제목 ============%>
            <div class="row">
                <label class="form-label" for="f-title">제목<span class="req">*</span></label>
                <div>
                    <input class="input" type="text" id="f-title" name="title"
                           value="<%= HE(fTitle) %>" maxlength="200"
                           data-required data-label="제목">
                </div>
            </div>

<%'============ 내용 ============%>
            <div class="row">
                <label class="form-label" for="f-content">내용<span class="req">*</span></label>
                <div>
                    <textarea class="textarea" id="f-content" name="content" rows="16"
                              data-required data-label="내용"><%= HE(fContent) %></textarea>
                    <p class="form-help">
                        기본 HTML 태그(&lt;p&gt;, &lt;b&gt;, &lt;ul&gt;, &lt;img&gt; 등)를 사용할 수 있습니다.
                        스크립트성 태그는 저장 시 자동으로 제거됩니다.
                    </p>
                </div>
            </div>

<%'============ 태그 ============%>
<% If bdSkin = "blog" Or bdSkin = "gallery" Then %>
            <div class="row">
                <label class="form-label" for="f-tags">태그</label>
                <div>
                    <input class="input" type="text" id="f-tags" name="tags"
                           value="<%= HE(fTags) %>" maxlength="300"
                           placeholder="콤마로 구분 : ASP, 보안, 아키텍처">
                </div>
            </div>
            <div class="row">
                <label class="form-label" for="f-link">관련 링크</label>
                <div>
                    <input class="input" type="url" id="f-link" name="link_url"
                           value="<%= HE(fLinkUrl) %>" maxlength="300" placeholder="https://">
                </div>
            </div>
<% End If %>

<%'============ 첨부파일 ============%>
<% If bdUseFile And bdFileCnt > 0 Then %>
            <div class="row">
                <label class="form-label">첨부파일</label>
                <div>
<% If fileCnt > 0 Then %>
                    <ul class="attach mb-3" style="padding-top:0">
<%                  Dim fi
                    For fi = 0 To fileCnt - 1 %>
                        <li>
                            <label class="check">
                                <input type="checkbox" name="del_file" value="<%= AVI(fileArr, BF_ID, fi) %>">
                                <span><%= AVH(fileArr, BF_ORGNAME, fi) %> <span class="size">(<%= HE(StringUtil.SizeStr(AVI(fileArr, BF_SIZE, fi))) %>)</span> 삭제</span>
                            </label>
                        </li>
<%                  Next %>
                    </ul>
<% End If %>
                    <div id="file-wrap">
                        <div class="file-row">
                            <input type="file" name="upfile1" class="input">
                        </div>
                    </div>
                    <div class="btn-row mt-2">
                        <button type="button" class="btn btn-sm" onclick="jkwAddFile('file-wrap', <%= bdFileCnt %>)">파일 추가</button>
                    </div>
                    <p class="form-help">
                        최대 <%= bdFileCnt %>개 / 개당 <%= HE(StringUtil.SizeStr(CDbl(bdFileSize) * 1024)) %> 까지<br>
                        허용 확장자 : <%= HE(bdFileExt) %>
                    </p>
                </div>
            </div>
<% End If %>

<%'============ 옵션 ============%>
<% If canSetNotice Or bdUseSecret Then %>
            <div class="row">
                <span class="form-label">옵션</span>
                <div class="check-row">
<% If canSetNotice Then %>
                    <label class="check">
                        <input type="checkbox" name="is_notice" value="1" <%= IIfStr(fIsNotice, "checked", "") %>>
                        <span>공지로 상단 고정</span>
                    </label>
<% End If %>
<% If bdUseSecret Then %>
                    <label class="check">
                        <input type="checkbox" name="is_secret" value="1" <%= IIfStr(fIsSecret, "checked", "") %>>
                        <span>비밀글</span>
                    </label>
<% End If %>
                </div>
            </div>
<% End If %>

        </div>

        <div class="btn-row between mt-5">
            <a class="btn" href="<%= HE(backUrl) %>">취소</a>
            <button type="submit" class="btn btn-primary btn-lg"><%= HE(modeLabel) %></button>
        </div>
    </form>

</div>

<!--#include file="../include/Footer.asp"-->
