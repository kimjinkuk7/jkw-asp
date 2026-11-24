<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/donate/InicisPay.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">
        <h2>이니시스 결제</h2>

        <div class="auth-box">
            <dl class="pay-summary">
                <dt>주문번호</dt><dd class="mono"><%= HE(ipOrder) %></dd>
                <dt>후원자</dt>  <dd><%= HE(ipName) %></dd>
                <dt>금액</dt>    <dd class="amt"><%= StringUtil.Comma(ipAmount) %>원</dd>
            </dl>

<% If PAY_IS_TEST Then %>
            <div class="alert alert-warn mt-3">
                <strong>이니시스 테스트 상점(<%= HE(INICIS_MID) %>)으로 동작합니다.</strong>
                실제 결제가 발생하지 않습니다.
            </div>
<% End If %>

            <div class="tc mt-3" id="ini-wait">
                <p class="form-help">이니시스 결제창을 불러오는 중입니다…</p>
            </div>

            <button type="button" class="btn btn-primary btn-block btn-lg mt-3" id="ini-open">
                결제창 다시 열기
            </button>
            <div class="tc mt-3">
                <a href="<%= U("/donate/Default.asp") %>" class="fs-sm">돌아가기</a>
            </div>
        </div>
    </div>
</div>

<%'============ 이니시스 웹표준 요청 폼 (화면에 보이지 않는다) ============%>
<form id="SendPayForm_id" method="post">
    <input type="hidden" name="version"      value="1.0">
    <input type="hidden" name="mid"          value="<%= HE(INICIS_MID) %>">
    <input type="hidden" name="oid"          value="<%= HE(ipOrder) %>">
    <input type="hidden" name="price"        value="<%= ipAmount %>">
    <input type="hidden" name="timestamp"    value="<%= HE(ipTs) %>">
    <input type="hidden" name="signature"    value="<%= HE(ipSign) %>">
    <input type="hidden" name="verification" value="<%= HE(ipVerify) %>">
    <input type="hidden" name="mKey"         value="<%= HE(ipMKey) %>">
    <input type="hidden" name="currency"     value="WON">
    <input type="hidden" name="goodname"     value="개발자 지원금">
    <input type="hidden" name="buyername"    value="<%= HE(ipName) %>">
    <input type="hidden" name="buyertel"     value="00000000000">
    <input type="hidden" name="buyeremail"   value="noreply@localhost">
    <input type="hidden" name="gopaymethod"  value="Card">
    <input type="hidden" name="acceptmethod" value="HPP(1):below1000:centerCd(Y)">
    <input type="hidden" name="returnUrl"    value="<%= HE(ipReturnUrl) %>">
    <input type="hidden" name="closeUrl"     value="<%= HE(ipCloseUrl) %>">
</form>

<script src="<%= HE(PgInicis.ScriptUrl()) %>"></script>
<script>
(function () {
    function openPay() {
        if (!window.INIStdPay) {
            var w = document.getElementById('ini-wait');
            if (w) w.innerHTML = '<p class="form-error">이니시스 스크립트를 불러오지 못했습니다. 네트워크를 확인해 주세요.</p>';
            return;
        }
        INIStdPay.pay('SendPayForm_id');
    }
    var btn = document.getElementById('ini-open');
    if (btn) btn.addEventListener('click', openPay);
    /* 페이지가 열리면 바로 결제창을 띄운다 */
    window.addEventListener('load', function () { setTimeout(openPay, 200); });
})();
</script>

<!--#include file="../include/Footer.asp"-->
