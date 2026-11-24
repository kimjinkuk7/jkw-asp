<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/PortfolioForm.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <ul class="breadcrumb">
                    <li><a href="<%= U("/admin/Portfolio.asp") %>">포트폴리오 관리</a></li>
                    <li><%= IIfStr(pfIsNew, "등록", "수정") %></li>
                </ul>
                <h2><%= IIfStr(pfIsNew, "새 프로젝트 등록", "프로젝트 수정") %></h2>
            </div>

            <%' 파일 업로드가 있으므로 id 는 QueryString 으로 전달 (BinaryRead 제약) %>
            <form method="post" action="<%= U("/admin/Portfolio_p.asp") %>?mode=save&amp;id=<%= pfId %>"
                  enctype="multipart/form-data" data-validate>
                <%= Auth.TokenField() %>

                <div class="admin-card">
                    <h3>기본 정보</h3>
                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="p-title">제목<span class="req">*</span></label>
                            <div><input class="input" type="text" id="p-title" name="title" value="<%= HE(vTitle) %>" maxlength="200" data-required data-label="제목"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-sub">부제</label>
                            <div><input class="input" type="text" id="p-sub" name="subtitle" value="<%= HE(vSubtitle) %>" maxlength="200"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-summary">요약</label>
                            <div>
                                <textarea class="textarea" id="p-summary" name="summary" rows="3" maxlength="500" style="min-height:90px"><%= HE(vSummary) %></textarea>
                                <p class="form-help">목록 카드에 노출됩니다.</p>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-cate">카테고리</label>
                            <div>
                                <select class="select" id="p-cate" name="category">
<%      Dim ci
        For ci = 0 To UBound(pfCates) %>
                                    <option value="<%= pfCates(ci) %>" <%= IIfStr(vCategory = pfCates(ci), "selected", "") %>><%= pfCates(ci) %></option>
<%      Next %>
                                </select>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-client">발주 / 소속</label>
                            <div><input class="input" type="text" id="p-client" name="client_name" value="<%= HE(vClient) %>" maxlength="100"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-role">담당 역할</label>
                            <div><input class="input" type="text" id="p-role" name="job_role" value="<%= HE(vRole) %>" maxlength="200"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-team">규모</label>
                            <div><input class="input" type="text" id="p-team" name="team_size" value="<%= HE(vTeam) %>" maxlength="50" placeholder="4명"></div>
                        </div>
                        <div class="row">
                            <span class="form-label">기간</span>
                            <div class="form-inline">
                                <input class="input" type="date" name="period_start" value="<%= HE(vStart) %>">
                                <input class="input" type="date" name="period_end"   value="<%= HE(vEnd) %>">
                                <label class="check">
                                    <input type="checkbox" name="is_ongoing" value="1" <%= IIfStr(vOngoing, "checked", "") %>>
                                    <span>진행중</span>
                                </label>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="admin-card">
                    <h3>링크 / 이미지</h3>
                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="p-site">사이트 URL</label>
                            <div><input class="input" type="url" id="p-site" name="site_url" value="<%= HE(vSiteUrl) %>" maxlength="300" placeholder="https://"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-repo">저장소 URL</label>
                            <div><input class="input" type="url" id="p-repo" name="repo_url" value="<%= HE(vRepoUrl) %>" maxlength="300" placeholder="https://github.com/"></div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="p-thumbfile">대표 이미지</label>
                            <div>
<% If Len(vThumb) > 0 Then %>
                                <img src="<%= HE(vThumb) %>" alt="현재 대표 이미지" style="max-width:280px;border:1px solid var(--c-border);border-radius:var(--radius-sm);margin-bottom:8px">
                                <label class="check">
                                    <input type="checkbox" name="del_thumb" value="1">
                                    <span>현재 이미지 삭제</span>
                                </label>
<% End If %>
                                <input type="file" id="p-thumbfile" name="thumb_file_up" class="input" accept="image/*">
                                <input type="hidden" name="thumb_cur" value="<%= HE(vThumb) %>">
                                <p class="form-help">jpg / png / gif / webp. 업로드하면 기존 이미지를 대체합니다.</p>
                            </div>
                        </div>
                    </div>
                </div>

                <div class="admin-card">
                    <h3>사용 기술</h3>
<% If pfSkillCnt = 0 Then %>
                    <div class="empty">등록된 기술 스택이 없습니다. 먼저 <a href="<%= U("/admin/Skill.asp") %>">기술 스택</a>을 등록하세요.</div>
<% Else %>
                    <div class="check-row">
<%      Dim si
        For si = 0 To pfSkillCnt - 1 %>
                        <label class="check">
                            <input type="checkbox" name="skill_ids" value="<%= AVI(pfSkillArr, SK_ID, si) %>" <%= IIfStr(SkillChecked(AVI(pfSkillArr, SK_ID, si)), "checked", "") %>>
                            <span><%= AVH(pfSkillArr, SK_NAME, si) %></span>
                        </label>
<%      Next %>
                    </div>
<% End If %>
                </div>

                <div class="admin-card">
                    <h3>상세 내용 (케이스 스터디)</h3>
                    <textarea class="textarea" name="content" rows="18"><%= HE(vContent) %></textarea>
                    <p class="form-help">
                        HTML 사용 가능. 예) &lt;h3&gt;배경&lt;/h3&gt;&lt;p&gt;...&lt;/p&gt;&lt;ul&gt;&lt;li&gt;...&lt;/li&gt;&lt;/ul&gt;
                    </p>
                </div>

                <div class="admin-card">
                    <h3>노출 설정</h3>
                    <div class="form-inline">
                        <input class="input" type="number" name="sort_no" value="<%= vSortNo %>" style="flex:0 0 110px" title="정렬 순서">
                        <label class="check">
                            <input type="checkbox" name="is_main" value="1" <%= IIfStr(vIsMain, "checked", "") %>>
                            <span>메인 화면 노출</span>
                        </label>
                        <label class="check">
                            <input type="checkbox" name="is_use" value="1" <%= IIfStr(vIsUse, "checked", "") %>>
                            <span>사용</span>
                        </label>
                    </div>
                </div>

                <div class="btn-row between mt-4">
                    <a class="btn" href="<%= U("/admin/Portfolio.asp") %>">목록</a>
                    <button type="submit" class="btn btn-primary btn-lg">저장</button>
                </div>
            </form>

        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
