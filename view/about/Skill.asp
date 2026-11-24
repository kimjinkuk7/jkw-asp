<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/about/Skill.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap wrap-narrow">

    <div class="page-head">
        <ul class="breadcrumb">
            <li><a href="<%= U("/main/Default.asp") %>">HOME</a></li>
            <li>ABOUT</li>
            <li>기술 스택</li>
        </ul>
        <h2>기술 스택</h2>
        <p class="desc">막대는 숙련도, 우측 숫자는 실무 사용 기간입니다.</p>
    </div>

    <!--#include file="../include/Lnb.asp"-->

<% If skGrpCnt = 0 Then %>
    <div class="empty">등록된 기술 스택이 없습니다.</div>
<% Else %>
<%  Dim gi, sj, gName %>
<%  For gi = 0 To skGrpCnt - 1
        gName = AV(skGrpArr, 0, gi) %>
    <section class="skill-group">
        <h4><%= HE(gName) %></h4>
        <div class="skill-list">
<%      For sj = 0 To skAllCnt - 1
            If AV(skAllArr, SK_GROUP, sj) = gName Then %>
            <div class="skill-item">
                <div class="top">
                    <span class="nm"><%= AVH(skAllArr, SK_NAME, sj) %></span>
                    <span class="yr">
<%                      If CDbl("0" & AV(skAllArr, SK_EXP, sj)) > 0 Then %><%= HE(AV(skAllArr, SK_EXP, sj)) %>년<% End If %>
                    </span>
                </div>
                <div class="skill-bar">
                    <i style="width:<%= AVI(skAllArr, SK_LEVEL, sj) %>%<% If Len(AV(skAllArr, SK_COLOR, sj)) > 0 Then %>;background:<%= HE(AV(skAllArr, SK_COLOR, sj)) %><% End If %>"></i>
                </div>
<%              If Len(AV(skAllArr, SK_DESC, sj)) > 0 Then %>
                <p class="form-help"><%= AVH(skAllArr, SK_DESC, sj) %></p>
<%              End If %>
            </div>
<%          End If
        Next %>
        </div>
    </section>
<%  Next %>
<% End If %>

    <div class="btn-row center mt-5">
        <a class="btn btn-primary" href="<%= U("/portfolio/Default.asp") %>">이 기술로 만든 프로젝트 보기</a>
    </div>

</div>

<!--#include file="../include/Footer.asp"-->
