<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Skill.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>기술 스택 관리</h2>
                <p class="desc">About &gt; 기술 스택, 이력서, 포트폴리오 매핑에 사용됩니다.</p>
            </div>

            <div class="admin-card">
                <h3><%= IIfStr(asIsEdit, "기술 수정 (ID " & asId & ")", "새 기술 추가") %></h3>
                <form method="post" action="<%= U("/admin/Skill_p.asp") %>" data-validate>
                    <%= Auth.TokenField() %>
                    <input type="hidden" name="mode" value="<%= IIfStr(asIsEdit, "update", "insert") %>">
                    <input type="hidden" name="id"   value="<%= asId %>">

                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="s-group">그룹<span class="req">*</span></label>
                            <div>
                                <input class="input" type="text" id="s-group" name="skill_group" value="<%= HE(sGroup) %>"
                                       maxlength="50" list="group-list" data-required data-label="그룹">
                                <datalist id="group-list">
<%      Dim gi
        For gi = 0 To UBound(asGroups) %>
                                    <option value="<%= asGroups(gi) %>"></option>
<%      Next %>
                                </datalist>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="s-name">기술명<span class="req">*</span></label>
                            <div><input class="input" type="text" id="s-name" name="skill_name" value="<%= HE(sName) %>" maxlength="50" data-required data-label="기술명"></div>
                        </div>
                        <div class="row">
                            <span class="form-label">숙련도 / 경력(년)</span>
                            <div class="form-inline">
                                <input class="input" type="number" name="level_pct" value="<%= sLevel %>" min="0" max="100" title="숙련도 %">
                                <input class="input" type="number" name="experience_yr" value="<%= HE(sExp) %>" min="0" max="50" step="0.5" title="경력 년수">
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="s-color">색상</label>
                            <div class="form-inline">
                                <input class="input" type="text" id="s-color" name="color_hex" value="<%= HE(sColor) %>" maxlength="7" placeholder="#4F46E5" style="flex:0 0 130px">
                                <input type="color" class="input" value="<%= HE(IIfStr(Len(sColor) = 7, sColor, "#2b5cff")) %>" style="flex:0 0 60px;padding:4px"
                                       onchange="document.getElementById('s-color').value=this.value">
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="s-desc">설명</label>
                            <div><input class="input" type="text" id="s-desc" name="skill_desc" value="<%= HE(sDesc) %>" maxlength="300"></div>
                        </div>
                        <div class="row">
                            <span class="form-label">정렬 / 사용</span>
                            <div class="form-inline">
                                <input class="input" type="number" name="sort_no" value="<%= sSort %>" style="flex:0 0 110px">
                                <label class="check">
                                    <input type="checkbox" name="is_use" value="1" <%= IIfStr(sUse, "checked", "") %>>
                                    <span>노출</span>
                                </label>
                            </div>
                        </div>
                    </div>

                    <div class="btn-row right mt-4">
<% If asIsEdit Then %>
                        <a class="btn" href="<%= U("/admin/Skill.asp") %>">새로 작성</a>
<% End If %>
                        <button type="submit" class="btn btn-primary"><%= IIfStr(asIsEdit, "수정", "추가") %></button>
                    </div>
                </form>
            </div>

            <div class="admin-card">
                <h3>전체 기술 <span class="muted fs-sm">(<%= asCnt %>건)</span></h3>
<% If asCnt = 0 Then %>
                <div class="empty">등록된 기술이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr><th class="num">ID</th><th>그룹</th><th>기술명</th><th>숙련도</th><th class="num">경력</th><th class="num">정렬</th><th class="ctr">노출</th><th>관리</th></tr>
                        </thead>
                        <tbody>
<%      Dim si
        For si = 0 To asCnt - 1 %>
                            <tr>
                                <td class="num"><%= AVI(asArr, SK_ID, si) %></td>
                                <td><span class="badge"><%= AVH(asArr, SK_GROUP, si) %></span></td>
                                <td class="fw7"><%= AVH(asArr, SK_NAME, si) %></td>
                                <td style="min-width:140px">
                                    <div class="skill-bar" style="width:110px;display:inline-block;vertical-align:middle">
                                        <i style="width:<%= AVI(asArr, SK_LEVEL, si) %>%<% If Len(AV(asArr, SK_COLOR, si)) > 0 Then %>;background:<%= HE(AV(asArr, SK_COLOR, si)) %><% End If %>"></i>
                                    </div>
                                    <span class="fs-xs muted"><%= AVI(asArr, SK_LEVEL, si) %>%</span>
                                </td>
                                <td class="num"><%= HE(AV(asArr, SK_EXP, si)) %></td>
                                <td class="num"><%= AVI(asArr, SKA_SORTNO, si) %></td>
                                <td class="ctr"><%= IIfStr(AVB(asArr, SKA_ISUSE, si), "O", "X") %></td>
                                <td>
                                    <a class="btn btn-sm btn-primary" href="<%= U("/admin/Skill.asp") %>?id=<%= AVI(asArr, SK_ID, si) %>">수정</a>
                                    <a class="btn btn-sm btn-danger"
                                       href="<%= U("/admin/Skill_p.asp") %>?mode=delete&amp;id=<%= AVI(asArr, SK_ID, si) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="삭제하시겠습니까? 포트폴리오 매핑도 함께 삭제됩니다.">삭제</a>
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
