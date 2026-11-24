<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/contact/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>CONTACT</li>
        </ul>
        <h2>문의하기</h2>
        <p class="desc"><%= HE(ctResponseTime) %></p>
    </div>

    <div class="contact-grid">

        <aside>
            <div class="card">
                <h3 style="font-size:var(--fs-md);margin-bottom:var(--sp-3)">연락처</h3>
                <div class="contact-info">
<% If Len(GLB_OWNER_EMAIL) > 0 Then %>
                    <div class="item">
                        <span class="lbl">이메일</span>
                        <a href="mailto:<%= HE(GLB_OWNER_EMAIL) %>"><%= HE(GLB_OWNER_EMAIL) %></a>
                    </div>
<% End If %>
<% If Len(GLB_OWNER_GITHUB) > 0 Then %>
                    <div class="item">
                        <span class="lbl">GitHub</span>
                        <a href="<%= HE(GLB_OWNER_GITHUB) %>" target="_blank" rel="noopener noreferrer">바로가기 ↗</a>
                    </div>
<% End If %>
                    <div class="item">
                        <span class="lbl">방명록</span>
                        <a href="<%= U("/board/List.asp") %>?bid=guest">가볍게 인사 남기기</a>
                    </div>
                </div>
            </div>
        </aside>

        <div>
            <form method="post" action="<%= U("/contact/Send_p.asp") %>" data-validate>
                <%= Auth.TokenField() %>

                <div class="form-table">

                    <div class="row">
                        <label class="form-label" for="ct-type">문의 유형</label>
                        <div>
                            <select class="select" id="ct-type" name="contact_type">
<%                          Dim ti
                            For ti = 0 To UBound(ctTypes) %>
                                <option value="<%= HE(ctTypes(ti)) %>"><%= HE(ctTypes(ti)) %></option>
<%                          Next %>
                            </select>
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-name">이름<span class="req">*</span></label>
                        <div>
                            <input class="input" type="text" id="ct-name" name="writer_name"
                                   value="<%= HE(ctName) %>" maxlength="50"
                                   data-required data-label="이름">
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-company">회사 / 소속</label>
                        <div>
                            <input class="input" type="text" id="ct-company" name="company" maxlength="100">
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-email">이메일<span class="req">*</span></label>
                        <div>
                            <input class="input" type="email" id="ct-email" name="email"
                                   value="<%= HE(ctEmail) %>" maxlength="100"
                                   data-required data-label="이메일" data-type="email">
                            <p class="form-help">회신받으실 주소를 정확히 입력해 주세요.</p>
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-mobile">연락처</label>
                        <div>
                            <input class="input" type="tel" id="ct-mobile" name="mobile" maxlength="20" placeholder="010-0000-0000">
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-subject">제목<span class="req">*</span></label>
                        <div>
                            <input class="input" type="text" id="ct-subject" name="subject" maxlength="200"
                                   data-required data-label="제목">
                        </div>
                    </div>

                    <div class="row">
                        <label class="form-label" for="ct-content">내용<span class="req">*</span></label>
                        <div>
                            <textarea class="textarea" id="ct-content" name="content" rows="10" maxlength="4000"
                                      placeholder="프로젝트 개요, 일정, 예산 범위 등을 함께 적어주시면 회신이 빨라집니다."
                                      data-required data-label="내용" data-minlen="10" data-counter="ct-cnt"></textarea>
                            <div class="form-help tr"><span id="ct-cnt">0</span> / 4000</div>
                        </div>
                    </div>

                    <div class="row">
                        <span class="form-label">개인정보 수집</span>
                        <div>
                            <div class="privacy-box">
                                <strong>개인정보 수집 및 이용 동의</strong><br>
                                · 수집 항목 : 이름, 이메일, 연락처, 회사명, 문의 내용, 접속 IP<br>
                                · 수집 목적 : 문의 접수 및 회신<br>
                                · 보유 기간 : 문의 처리 완료 후 1년, 이후 지체 없이 파기<br>
                                · 동의를 거부할 수 있으나, 이 경우 문의 접수가 불가합니다.
                            </div>
                            <label class="check mt-2">
                                <input type="checkbox" name="is_agree" value="1" id="ct-agree">
                                <span>위 내용에 동의합니다.<span class="req">*</span></span>
                            </label>
                        </div>
                    </div>

                </div>

                <div class="btn-row right mt-5">
                    <button type="submit" class="btn btn-primary btn-lg">문의 보내기</button>
                </div>
            </form>
        </div>

    </div>
</div>

<%
g_extraScript = "<script>" & _
"document.querySelector('form[data-validate]').addEventListener('submit',function(e){" & _
"  if(!document.getElementById('ct-agree').checked){" & _
"    alert('개인정보 수집 및 이용에 동의해 주세요.');" & _
"    e.preventDefault(); e.stopImmediatePropagation();" & _
"  }" & _
"},true);" & _
"</script>"
%>
<!--#include file="../include/Footer.asp"-->
