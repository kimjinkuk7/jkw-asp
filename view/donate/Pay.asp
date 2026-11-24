<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/donate/Pay.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="auth-wrap">

        <h2>결제 확인</h2>

        <div class="auth-box">
            <dl class="pay-summary">
                <dt>주문번호</dt><dd class="mono"><%= HE(pyOrderId) %></dd>
                <dt>후원자</dt>  <dd><%= HE(pyName) %></dd>
                <dt>금액</dt>    <dd class="amt"><%= StringUtil.Comma(pyAmount) %>원</dd>
            </dl>

<% If LCase(pyProvider) = "mock" Then %>
            <%'============ 시연용 결제창 ============%>
            <div class="alert alert-warn mt-3">
                <strong>시연용 결제 화면입니다.</strong><br>
                실제 결제창이 아니며 카드번호를 입력하지 않습니다. 아래에서 결과를 직접 선택하면
                실제 PG 가 돌려주는 응답과 같은 경로로 처리됩니다.
            </div>

            <form method="post" action="<%= U("/donate/PayResult.asp") %>?order=<%= UE(pyOrderId) %>">
                <%= Auth.TokenField() %>
                <input type="hidden" name="order" value="<%= HE(pyOrderId) %>">

                <button type="submit" name="result" value="success"
                        class="btn btn-primary btn-block btn-lg mt-3">결제 성공으로 처리</button>
                <button type="submit" name="result" value="fail"
                        class="btn btn-block mt-2">결제 실패로 처리</button>
            </form>

            <div class="tc mt-3">
                <a href="<%= U("/donate/Default.asp") %>" class="fs-sm">돌아가기</a>
            </div>

<% ElseIf LCase(pyProvider) = "toss" Then %>
            <%'============ 토스페이먼츠 결제창 ============%>
            <div class="alert alert-info mt-3">
                토스페이먼츠 결제창이 열립니다.
<% If PAY_IS_TEST Then %>
                <strong>테스트 키로 동작하므로 실제 결제가 발생하지 않습니다.</strong>
<% End If %>
            </div>

            <button type="button" id="btn-toss" class="btn btn-primary btn-block btn-lg mt-3">
                <%= StringUtil.Comma(pyAmount) %>원 결제하기
            </button>
            <div class="tc mt-3">
                <a href="<%= U("/donate/Default.asp") %>" class="fs-sm">돌아가기</a>
            </div>

            <script src="https://js.tosspayments.com/v1/payment"></script>
            <script>
            (function () {
                var btn = document.getElementById('btn-toss');
                if (!btn || !window.TossPayments) return;
                var tp = TossPayments('<%= HE(TOSS_CLIENT_KEY) %>');
                btn.addEventListener('click', function () {
                    btn.disabled = true;
                    tp.requestPayment('카드', {
                        amount:      <%= pyAmount %>,
                        orderId:     '<%= HE(pyOrderId) %>',
                        orderName:   '개발자 지원금',
                        customerName:'<%= JE(pyName) %>',
                        successUrl:  '<%= GLB_SITE_URL & U("/donate/PayResult.asp") %>?order=<%= UE(pyOrderId) %>',
                        failUrl:     '<%= GLB_SITE_URL & U("/donate/PayResult.asp") %>?order=<%= UE(pyOrderId) %>'
                    }).catch(function (e) {
                        btn.disabled = false;
                        alert('결제를 시작할 수 없습니다 : ' + (e && e.message ? e.message : ''));
                    });
                });
            })();
            </script>
<% Else %>
            <div class="alert alert-error mt-3">
                지원하지 않는 결제 수단입니다 : <%= HE(pyProvider) %>
            </div>
<% End If %>
        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
