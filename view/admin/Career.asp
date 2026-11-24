<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Career.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>경력 · 학력 관리</h2>
                <p class="desc">About &gt; 경력·학력 및 이력서 페이지에 노출되는 항목입니다.</p>
            </div>

            <div class="admin-card">
                <h3><%= IIfStr(acIsEdit, "항목 수정 (ID " & acId & ")", "새 항목 추가") %></h3>
                <form method="post" action="<%= U("/admin/Career_p.asp") %>" data-validate>
                    <%= Auth.TokenField() %>
                    <input type="hidden" name="mode" value="<%= IIfStr(acIsEdit, "update", "insert") %>">
                    <input type="hidden" name="id"   value="<%= acId %>">

                    <div class="form-table">
                        <div class="row">
                            <label class="form-label" for="c-type">구분<span class="req">*</span></label>
                            <div>
                                <select class="select" id="c-type" name="career_type">
<%      Dim ti
        For ti = 0 To UBound(acTypes) %>
                                    <option value="<%= acTypes(ti) %>" <%= IIfStr(eType = acTypes(ti), "selected", "") %>><%= HE(acTypeNames(ti)) %></option>
<%      Next %>
                                </select>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="c-title">제목<span class="req">*</span></label>
                            <div>
                                <input class="input" type="text" id="c-title" name="title" value="<%= HE(eTitle) %>" maxlength="200" data-required data-label="제목">
                                <p class="form-help">회사명 / 학교명 / 자격명 / 수상명</p>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="c-pos">부제</label>
                            <div>
                                <input class="input" type="text" id="c-pos" name="position" value="<%= HE(ePosition) %>" maxlength="100">
                                <p class="form-help">직위 / 전공 / 발급기관</p>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="c-url">관련 링크</label>
                            <div><input class="input" type="url" id="c-url" name="org_url" value="<%= HE(eOrgUrl) %>" maxlength="300" placeholder="https://"></div>
                        </div>
                        <div class="row">
                            <span class="form-label">기간</span>
                            <div class="form-inline">
                                <input class="input" type="date" name="date_start" value="<%= HE(eStart) %>">
                                <input class="input" type="date" name="date_end"   value="<%= HE(eEnd) %>">
                                <label class="check">
                                    <input type="checkbox" name="is_current" value="1" <%= IIfStr(eCurrent, "checked", "") %>>
                                    <span>재직/재학 중</span>
                                </label>
                            </div>
                        </div>
                        <div class="row">
                            <label class="form-label" for="c-desc">설명</label>
                            <div>
                                <textarea class="textarea" id="c-desc" name="description" rows="5" maxlength="2000" style="min-height:120px"><%= HE(eDesc) %></textarea>
                                <p class="form-help">줄바꿈이 그대로 반영됩니다. 주요 업무를 한 줄씩 적어주세요.</p>
                            </div>
                        </div>
                        <div class="row">
                            <span class="form-label">정렬 / 사용</span>
                            <div class="form-inline">
                                <input class="input" type="number" name="sort_no" value="<%= eSortNo %>" style="flex:0 0 110px">
                                <label class="check">
                                    <input type="checkbox" name="is_use" value="1" <%= IIfStr(eIsUse, "checked", "") %>>
                                    <span>노출</span>
                                </label>
                            </div>
                        </div>
                    </div>

                    <div class="btn-row right mt-4">
<% If acIsEdit Then %>
                        <a class="btn" href="<%= U("/admin/Career.asp") %>">새로 작성</a>
<% End If %>
                        <button type="submit" class="btn btn-primary"><%= IIfStr(acIsEdit, "수정", "추가") %></button>
                    </div>
                </form>
            </div>

            <div class="admin-card">
                <h3>전체 항목 <span class="muted fs-sm">(<%= acCnt %>건)</span></h3>
<% If acCnt = 0 Then %>
                <div class="empty">등록된 항목이 없습니다.</div>
<% Else %>
                <div class="data-table-wrap">
                    <table class="data-table">
                        <thead>
                            <tr><th class="num">ID</th><th>구분</th><th class="wide">제목</th><th>부제</th><th>기간</th><th class="num">정렬</th><th class="ctr">노출</th><th>관리</th></tr>
                        </thead>
                        <tbody>
<%      Dim ci
        For ci = 0 To acCnt - 1 %>
                            <tr>
                                <td class="num"><%= AVI(acArr, CR_ID, ci) %></td>
                                <td><span class="badge"><%= HE(CareerTypeName(AV(acArr, CR_TYPE, ci))) %></span></td>
                                <td class="wide fw7"><%= AVH(acArr, CR_TITLE, ci) %></td>
                                <td><%= AVH(acArr, CR_POSITION, ci) %></td>
                                <td class="fs-xs"><%= HE(DateUtil.Period(AV(acArr, CR_START, ci), AV(acArr, CR_END, ci), AVB(acArr, CR_CURRENT, ci))) %></td>
                                <td class="num"><%= AVI(acArr, CRA_SORTNO, ci) %></td>
                                <td class="ctr"><%= IIfStr(AVB(acArr, CRA_ISUSE, ci), "O", "X") %></td>
                                <td>
                                    <a class="btn btn-sm btn-primary" href="<%= U("/admin/Career.asp") %>?id=<%= AVI(acArr, CR_ID, ci) %>">수정</a>
                                    <a class="btn btn-sm btn-danger"
                                       href="<%= U("/admin/Career_p.asp") %>?mode=delete&amp;id=<%= AVI(acArr, CR_ID, ci) %>&amp;_csrf=<%= UE(Auth.Token()) %>"
                                       data-confirm="삭제하시겠습니까?">삭제</a>
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
