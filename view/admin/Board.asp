<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Board.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

<%'==================== 목록 ====================%>
<% If Not abIsEdit Then %>

            <div class="page-head">
                <h2>게시판 관리</h2>
                <p class="desc">
                    행 1개가 게시판 1개입니다. 새 게시판을 추가하면 메뉴에 등록만 하면 되고,
                    <strong>ASP 코드는 수정할 필요가 없습니다.</strong>
                </p>
            </div>

            <div class="admin-card">
                <h3>게시판 목록 (<%= abListCnt %>)</h3>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr>
                                <th>ID</th><th>이름</th><th>스킨</th>
                                <th class="ctr">댓글</th><th class="ctr">첨부</th><th class="ctr">분류</th>
                                <th class="ctr">비밀</th><th class="ctr">답변</th>
                                <th class="ctr">목록</th><th class="ctr">읽기</th><th class="ctr">쓰기</th><th class="ctr">댓글권한</th>
                                <th class="ctr">사용</th><th></th>
                            </tr>
                        </thead>
                        <tbody>
<%      Dim li
        For li = 0 To abListCnt - 1 %>
                            <tr>
                                <td><code><%= AVH(abListArr, BM_ID, li) %></code></td>
                                <td class="fw7"><%= AVH(abListArr, BM_NAME, li) %></td>
                                <td><span class="badge"><%= AVH(abListArr, BM_SKIN, li) %></span></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_USE_COMMENT,  li), "O", "-") %></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_USE_FILE,     li), "O", "-") %></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_USE_CATEGORY, li), "O", "-") %></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_USE_SECRET,   li), "O", "-") %></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_USE_REPLY,    li), "O", "-") %></td>
                                <td class="ctr"><%= AVI(abListArr, BM_AUTH_LIST,    li) %></td>
                                <td class="ctr"><%= AVI(abListArr, BM_AUTH_READ,    li) %></td>
                                <td class="ctr"><%= AVI(abListArr, BM_AUTH_WRITE,   li) %></td>
                                <td class="ctr"><%= AVI(abListArr, BM_AUTH_COMMENT, li) %></td>
                                <td class="ctr"><%= IIfStr(AVB(abListArr, BM_IS_USE, li), "O", "X") %></td>
                                <td>
                                    <a class="btn btn-sm btn-primary" href="<%= U("/admin/Board.asp") %>?bid=<%= UE(AV(abListArr, BM_ID, li)) %>">설정</a>
                                    <a class="btn btn-sm" href="<%= U("/board/List.asp") %>?bid=<%= UE(AV(abListArr, BM_ID, li)) %>">보기</a>
                                </td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
            </div>

            <div class="admin-card">
                <h3>새 게시판 추가</h3>
                <form method="post" action="<%= U("/admin/Board_p.asp") %>" data-validate>
                    <%= Auth.TokenField() %>
                    <input type="hidden" name="mode" value="insert">
                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="n-bid">게시판 ID<span class="req">*</span></label>
                            <div>
                                <input class="input" type="text" id="n-bid" name="board_id" maxlength="20"
                                       placeholder="qna" data-required data-label="게시판 ID">
                                <p class="form-help">영소문자로 시작하는 영문·숫자·_ 조합 2~20자. URL 에 노출됩니다.</p>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="n-name">이름<span class="req">*</span></label>
                            <div><input class="input" type="text" id="n-name" name="board_name" maxlength="50" placeholder="Q&amp;A" data-required data-label="이름"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="n-desc">설명</label>
                            <div><input class="input" type="text" id="n-desc" name="board_desc" maxlength="300"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="n-skin">스킨</label>
                            <div>
                                <select class="select" id="n-skin" name="skin_type">
<%      Dim si
        For si = 0 To UBound(abSkins) %>
                                    <option value="<%= abSkins(si) %>"><%= abSkins(si) %></option>
<%      Next %>
                                </select>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="n-sort">정렬 순서</label>
                            <div><input class="input" type="number" id="n-sort" name="sort_no" value="70"></div>
                        </div>
                    </div>
                    <div class="btn-row right mt-4">
                        <button type="submit" class="btn btn-primary">게시판 추가</button>
                    </div>
                </form>
                <p class="form-help mt-3">
                    추가 후 <strong>사이트 설정 &gt; 메뉴</strong> 또는 DB 의 <code>tb_menu</code> 에
                    <code>/board/List.asp?bid=&lt;ID&gt;</code> 링크를 등록하면 GNB 에 노출됩니다.
                </p>
            </div>


<%'==================== 편집 ====================%>
<% Else %>

            <div class="page-head">
                <ul class="breadcrumb">
                    <li><a href="<%= U("/admin/Board.asp") %>">게시판 관리</a></li>
                    <li><%= HE(abBid) %></li>
                </ul>
                <h2><%= AVH(abListArr, BM_NAME, abIdx) %> <span class="muted fs-sm">(<%= HE(abBid) %>)</span></h2>
            </div>

            <form method="post" action="<%= U("/admin/Board_p.asp") %>" data-validate>
                <%= Auth.TokenField() %>
                <input type="hidden" name="mode" value="update">
                <input type="hidden" name="board_id" value="<%= HE(abBid) %>">

                <div class="admin-card">
                    <h3>기본 정보</h3>
                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="e-name">이름<span class="req">*</span></label>
                            <div><input class="input" type="text" id="e-name" name="board_name" value="<%= AVH(abListArr, BM_NAME, abIdx) %>" maxlength="50" data-required data-label="이름"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="e-desc">설명</label>
                            <div><input class="input" type="text" id="e-desc" name="board_desc" value="<%= AVH(abListArr, BM_DESC, abIdx) %>" maxlength="300"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="e-skin">스킨</label>
                            <div>
                                <select class="select" id="e-skin" name="skin_type">
<%      Dim si2
        For si2 = 0 To UBound(abSkins) %>
                                    <option value="<%= abSkins(si2) %>" <%= IIfStr(AV(abListArr, BM_SKIN, abIdx) = abSkins(si2), "selected", "") %>><%= abSkins(si2) %></option>
<%      Next %>
                                </select>
                                <p class="form-help">list = 표형 / gallery = 썸네일 / blog = 요약카드 / guest = 방명록</p>
                            </div>
                        </div>
                        <div class="row">
                            <span class="form-label">목록 설정</span>
                            <div class="form-inline">
                                <input class="input" type="number" name="list_count"  value="<%= AVI(abListArr, BM_LIST_COUNT, abIdx) %>"  min="1" max="100" title="페이지당 글 수">
                                <input class="input" type="number" name="page_block"  value="<%= AVI(abListArr, BM_PAGE_BLOCK, abIdx) %>"  min="1" max="20"  title="페이지 번호 개수">
                                <input class="input" type="number" name="gallery_col" value="<%= AVI(abListArr, BM_GAL_COL, abIdx) %>"     min="1" max="6"   title="갤러리 컬럼 수">
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="e-sort">정렬 / 사용</label>
                            <div class="form-inline">
                                <input class="input" type="number" id="e-sort" name="sort_no" value="<%= AVI(abListArr, BM_SORT_NO, abIdx) %>">
                                <label class="check">
                                    <input type="checkbox" name="is_use" value="1" <%= IIfStr(AVB(abListArr, BM_IS_USE, abIdx), "checked", "") %>>
                                    <span>사용</span>
                                </label>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="admin-card">
                    <h3>기능</h3>
                    <div class="check-row">
                        <label class="check"><input type="checkbox" name="use_comment"  value="1" <%= IIfStr(AVB(abListArr, BM_USE_COMMENT,  abIdx), "checked", "") %>><span>댓글</span></label>
                        <label class="check"><input type="checkbox" name="use_file"     value="1" <%= IIfStr(AVB(abListArr, BM_USE_FILE,     abIdx), "checked", "") %>><span>첨부파일</span></label>
                        <label class="check"><input type="checkbox" name="use_category" value="1" <%= IIfStr(AVB(abListArr, BM_USE_CATEGORY, abIdx), "checked", "") %>><span>분류</span></label>
                        <label class="check"><input type="checkbox" name="use_secret"   value="1" <%= IIfStr(AVB(abListArr, BM_USE_SECRET,   abIdx), "checked", "") %>><span>비밀글</span></label>
                        <label class="check"><input type="checkbox" name="use_reply"    value="1" <%= IIfStr(AVB(abListArr, BM_USE_REPLY,    abIdx), "checked", "") %>><span>답변형</span></label>
                        <label class="check"><input type="checkbox" name="use_editor"   value="1" <%= IIfStr(AVB(abListArr, BM_USE_EDITOR,   abIdx), "checked", "") %>><span>HTML 에디터</span></label>
                        <label class="check"><input type="checkbox" name="use_notice"   value="1" <%= IIfStr(AVB(abListArr, BM_USE_NOTICE,   abIdx), "checked", "") %>><span>공지 고정</span></label>
                    </div>
                </div>

                <div class="admin-card">
                    <h3>첨부파일</h3>
                    <div class="form-table">
                        <div class="row">
                            <span class="form-label">개수 / 용량(KB)</span>
                            <div class="form-inline">
                                <input class="input" type="number" name="file_max_count" value="<%= AVI(abListArr, BM_FILE_CNT, abIdx) %>" min="0" max="20">
                                <input class="input" type="number" name="file_max_size"  value="<%= AVI(abListArr, BM_FILE_SIZE, abIdx) %>" min="0">
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="e-ext">허용 확장자</label>
                            <div>
                                <input class="input" type="text" id="e-ext" name="file_ext_allow" value="<%= AVH(abListArr, BM_FILE_EXT, abIdx) %>" maxlength="300">
                                <p class="form-help">
                                    콤마 구분. 실행 가능한 확장자는 여기에 넣어도
                                    <code>App.Conf.asp</code> 의 <code>UPLOAD_DENY_EXT</code> 가 항상 차단합니다.
                                </p>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="admin-card">
                    <h3>권한</h3>
                    <div class="form-table">
<%      Dim authNames, authCols, ai2
        authNames = Array("목록 열람", "본문 열람", "글쓰기", "댓글", "게시판 관리")
        authCols  = Array("auth_list", "auth_read", "auth_write", "auth_comment", "auth_admin")
        Dim authVals
        authVals = Array(AVI(abListArr, BM_AUTH_LIST, abIdx), AVI(abListArr, BM_AUTH_READ, abIdx), _
                         AVI(abListArr, BM_AUTH_WRITE, abIdx), AVI(abListArr, BM_AUTH_COMMENT, abIdx), _
                         AVI(abListArr, BM_AUTH_ADMIN, abIdx))
        For ai2 = 0 To UBound(authNames) %>
                        <div class="row">
                            <label class="form-label" for="e-<%= authCols(ai2) %>"><%= authNames(ai2) %></label>
                            <div>
                                <select class="select" id="e-<%= authCols(ai2) %>" name="<%= authCols(ai2) %>">
<%          Dim lv
            For lv = 0 To UBound(abLevels) %>
                                    <option value="<%= abLevels(lv) %>" <%= IIfStr(authVals(ai2) = abLevels(lv), "selected", "") %>><%= HE(LevelName(abLevels(lv))) %></option>
<%          Next %>
                                </select>
                            </div>
                        </div>
<%      Next %>
                    </div>
                </div>

                <div class="btn-row between mt-4">
                    <a class="btn" href="<%= U("/admin/Board.asp") %>">목록</a>
                    <button type="submit" class="btn btn-primary btn-lg">설정 저장</button>
                </div>
            </form>

            <%'--- 분류 관리 ---%>
            <div class="admin-card mt-5">
                <h3>분류 관리</h3>
<% If abCateCnt = 0 Then %>
                <div class="empty">등록된 분류가 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead><tr><th class="ctr">ID</th><th>분류명</th><th></th></tr></thead>
                        <tbody>
<%      Dim ci
        For ci = 0 To abCateCnt - 1 %>
                            <tr>
                                <td class="ctr"><%= AVI(abCateArr, CAT_ID, ci) %></td>
                                <td class="fw7"><%= AVH(abCateArr, CAT_NAME, ci) %></td>
                                <td>
                                    <a class="btn btn-sm btn-danger"
                                       href="<%= U("/admin/Board_p.asp") %>?mode=cate_delete&amp;board_id=<%= UE(abBid) %>&amp;category_id=<%= AVI(abCateArr, CAT_ID, ci) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="분류를 삭제하시겠습니까? 이 분류의 글은 '분류 없음' 이 됩니다.">삭제</a>
                                </td>
                            </tr>
<%      Next %>
                        </tbody>
                    </table>
                </div>
<% End If %>

                <form method="post" action="<%= U("/admin/Board_p.asp") %>" class="form-inline mt-3" data-validate>
                    <%= Auth.TokenField() %>
                    <input type="hidden" name="mode" value="cate_insert">
                    <input type="hidden" name="board_id" value="<%= HE(abBid) %>">
                    <input class="input" type="text" name="category_name" maxlength="50" placeholder="새 분류명" data-required data-label="분류명">
                    <input class="input" type="number" name="sort_no" value="10" style="flex:0 0 90px" title="정렬">
                    <button type="submit" class="btn btn-primary">추가</button>
                </form>
            </div>

<% End If %>

        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
