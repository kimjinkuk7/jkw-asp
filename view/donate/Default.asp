<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/donate/Default.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">

    <div class="page-head">
        <h2>개발자 지원금 ☕</h2>
        <p class="desc">
            이 사이트는 같은 서비스를 Classic ASP · Spring Boot · Python · PHP 네 갈래로
            만들어 보는 삽질의 결과물입니다.<br>
            재미있게 보셨다면 커피 한 잔 값으로 응원해 주세요. 밤샘 연료로 쓰겠습니다.
        </p>
    </div>

<% If PAY_IS_TEST Then %>
    <div class="alert alert-warn">
        <strong>테스트 결제입니다.</strong>
        모든 결제 모듈이 테스트 키로 동작하므로 <strong>실제로 돈이 빠져나가지 않습니다.</strong>
        토스·이니시스는 진짜 결제창이 열리지만 승인만 흉내내고 청구는 되지 않습니다.
    </div>
<% End If %>

    <div class="donate-grid">

        <%'============ 후원 폼 ============%>
        <div class="donate-form-box">
            <h3>후원하기</h3>

            <%' data-toss-key : 토스 클라이언트 키는 브라우저에 노출되는 것이 정상이다.
              '                 시크릿 키(TOSS_SECRET_KEY)는 절대 화면에 내보내지 않는다. %>
            <%' data-no-lock : 제출을 AJAX 로 직접 처리하므로 공통 중복제출 잠금을 끈다.
              '               (결제창을 닫고 돌아왔을 때 버튼이 죽는 것을 막는다) %>
            <form method="post" action="<%= U("/donate/Pay_p.asp") %>" data-validate data-no-lock
                  id="donate-form" data-toss-key="<%= HE(TOSS_CLIENT_KEY) %>">
                <%= Auth.TokenField() %>

                <div class="form-group">
                    <span class="form-label">금액</span>
                    <div class="amount-row">
<%                  Dim ai
                    For ai = 0 To UBound(dnPresets) %>
                        <button type="button" class="btn amount-btn" data-amount="<%= dnPresets(ai) %>">
                            <%= StringUtil.Comma(dnPresets(ai)) %>원
                        </button>
<%                  Next %>
                    </div>
                    <input class="input mt-2" type="number" id="dn-amount" name="amount"
                           value="3000" min="<%= PAY_MIN_AMOUNT %>" max="<%= PAY_MAX_AMOUNT %>" step="100"
                           data-required data-label="금액">
                    <div class="form-help">
                        <%= StringUtil.Comma(PAY_MIN_AMOUNT) %>원 ~ <%= StringUtil.Comma(PAY_MAX_AMOUNT) %>원
                    </div>
                </div>

                <div class="form-group">
                    <label class="form-label" for="dn-name">후원자 이름</label>
                    <input class="input" type="text" id="dn-name" name="donor_name"
                           value="<%= HE(dnDefName) %>" maxlength="30"
                           placeholder="익명" data-required data-label="후원자 이름">
                </div>

                <div class="form-group">
                    <label class="form-label" for="dn-msg">응원 한마디</label>
                    <textarea class="textarea" id="dn-msg" name="message" rows="3" maxlength="200"
                              placeholder="남기지 않아도 됩니다." data-counter="dn-msg-cnt"></textarea>
                    <div class="form-help tr"><span id="dn-msg-cnt">0</span> / 200</div>
                </div>

                <div class="form-group">
                    <label class="check">
                        <input type="checkbox" name="is_public" value="1" checked>
                        후원자 목록에 이름과 한마디를 공개합니다
                    </label>
                </div>

                <div class="form-group">
                    <span class="form-label">결제 모듈</span>
                    <div class="pg-pick">
<%                  Dim pi, pv, pOn
                    For pi = 0 To UBound(dnPgList)
                        pv = Trim(dnPgList(pi))
                        If Len(pv) > 0 Then
                            pOn = (pv = dnDefault) %>
                        <label class="pg-pick-item<%= IIfStr(PgReady(pv), "", " is-off") %>">
                            <input type="radio" name="pg" value="<%= HE(pv) %>"
                                   <%= IIfStr(pOn, "checked", "") %>
                                   <%= IIfStr(PgReady(pv), "", "disabled") %>>
                            <span class="ttl"><%= HE(PgLabel(pv)) %></span>
                            <span class="dsc">
                                <%= HE(PgDesc(pv)) %>
<%                              If Not PgReady(pv) Then %>
                                <br><strong>키가 설정되지 않아 사용할 수 없습니다.</strong>
<%                              End If %>
                            </span>
                        </label>
<%                      End If
                    Next %>
                    </div>
                </div>

                <button type="submit" class="btn btn-primary btn-block btn-lg">후원하기</button>
            </form>
        </div>

        <%'============ 후원자 목록 ============%>
        <div class="donate-list-box">
            <h3>응원해 주신 분들</h3>

            <div class="donate-stat">
                <div class="dash-card">
                    <div class="lbl">후원 건수</div>
                    <div class="num"><%= StringUtil.Comma(dnTotalCnt) %></div>
                </div>
                <div class="dash-card">
                    <div class="lbl">누적 금액</div>
                    <div class="num"><%= StringUtil.Comma(dnTotalAmt) %>원</div>
                </div>
            </div>

<% If dnCnt = 0 Then %>
            <div class="empty">아직 후원 내역이 없습니다. 첫 커피의 주인이 되어보세요.</div>
<% Else %>
            <ul class="donate-list">
<%          Dim di
            For di = 0 To dnCnt - 1 %>
                <li>
                    <div class="head">
                        <span class="name"><%= AVH(dnList, DN_NAME, di) %></span>
                        <span class="amt"><%= StringUtil.Comma(AVI(dnList, DN_AMOUNT, di)) %>원</span>
<%                  If AVB(dnList, DN_ISTEST, di) Then %>
                        <span class="badge">테스트</span>
<%                  End If %>
                        <span class="date"><%= HE(DateUtil.Fmt(AV(dnList, DN_PAIDDT, di), "ymdhi")) %></span>
                    </div>
<%                  If Len(AV(dnList, DN_MESSAGE, di)) > 0 Then %>
                    <div class="msg"><%= AVH(dnList, DN_MESSAGE, di) %></div>
<%                  End If %>
                </li>
<%          Next %>
            </ul>
<% End If %>
        </div>

    </div>
</div>

<% If Len(TOSS_CLIENT_KEY) > 0 Then %>
<script src="https://js.tosspayments.com/v1/payment"></script>
<% End If %>

<% If dnHasMock Then %>
<%'======================================================================
  ' 모의 결제창 (모달)
  ' 실제 PG 결제창의 구조를 그대로 따라간다 :
  '   주문 생성 → 결제창 → 승인 요청 → 완료
  ' ★ 카드번호 칸은 테스트 값이 박힌 읽기 전용이다.
  '   실제로 입력받으면 그 화면이 곧 카드정보 수집 폼이 되므로 그렇게 하지 않는다.
  '   서버로도 카드번호를 보내지 않는다 (실제 결제에서도 PG 창이 직접 받는다).
  '======================================================================%>
<%' <dialog> + showModal() 을 쓴다. backdrop / 포커스 트랩 / ESC / 바깥 요소 inert 를
  ' 브라우저가 처리해 주므로 직접 만든 오버레이보다 정확하다. %>
<dialog class="pg-dialog" id="pg-modal" aria-labelledby="pg-title">
    <div class="pg-window">

        <div class="pg-top">
            <div class="pg-mall">
                <span id="pg-title">jkw.com</span>
                <em>개발자 지원금</em>
            </div>
            <button type="button" class="pg-x" id="pg-close" aria-label="닫기">×</button>
        </div>

        <div class="pg-demo-bar">시연용 결제창 · 실제 결제가 발생하지 않습니다</div>

        <ol class="pg-steps" id="pg-steps">
            <li data-step="1">주문 생성</li>
            <li data-step="2">결제 정보</li>
            <li data-step="3">승인 요청</li>
            <li data-step="4">완료</li>
        </ol>

        <div class="pg-amount">
            <span class="lbl">결제 금액</span>
            <strong id="pg-amount">0원</strong>
        </div>

        <div class="pg-order">
            주문번호 <span class="mono" id="pg-order">-</span>
        </div>

        <div class="pg-methods" id="pg-methods">
            <button type="button" class="pg-m is-on" data-method="신용카드">신용카드</button>
            <button type="button" class="pg-m" data-method="계좌이체">계좌이체</button>
            <button type="button" class="pg-m" data-method="간편결제">간편결제</button>
        </div>

        <div class="pg-fields" id="pg-fields">
            <label>카드번호
                <input type="text" value="4330-1234-****-5678" readonly tabindex="-1">
            </label>
            <div class="pg-row">
                <label>유효기간
                    <input type="text" value="12/29" readonly tabindex="-1">
                </label>
                <label>비밀번호
                    <input type="text" value="**" readonly tabindex="-1">
                </label>
            </div>
            <p class="pg-note">
                시연용 테스트 카드가 미리 채워져 있습니다. 실제 카드번호를 입력하는 칸이 아니며
                이 값은 서버로 전송되지 않습니다.
            </p>
        </div>

        <div class="pg-progress" id="pg-progress" hidden>
            <span class="spin" aria-hidden="true"></span>
            <span id="pg-progress-txt">승인 요청 중…</span>
        </div>

        <div class="pg-done" id="pg-done" hidden>
            <div class="mark">✓</div>
            <div class="txt">결제가 승인되었습니다</div>
            <div class="sub mono" id="pg-tid"></div>
        </div>

        <div class="pg-err" id="pg-err" hidden></div>

        <div class="pg-actions" id="pg-actions">
            <button type="button" class="btn btn-primary btn-block btn-lg" id="pg-pay">결제하기</button>
            <button type="button" class="btn btn-block mt-2" id="pg-fail">결제 실패 시나리오로 진행</button>
        </div>
    </div>
</dialog>
<% End If %>

<%
g_extraScript = "<script>" & _
"(function(){" & _
"  var inp=document.getElementById('dn-amount');" & _
"  if(!inp) return;" & _
"  var btns=document.querySelectorAll('.amount-btn');" & _
"  function mark(){" & _
"    for(var i=0;i<btns.length;i++){" & _
"      btns[i].classList.toggle('is-on', btns[i].getAttribute('data-amount')===String(inp.value));" & _
"    }" & _
"  }" & _
"  for(var i=0;i<btns.length;i++){" & _
"    btns[i].addEventListener('click',function(){ inp.value=this.getAttribute('data-amount'); mark(); });" & _
"  }" & _
"  inp.addEventListener('input', mark);" & _
"  mark();" & _
"})();" & _
"</script>"
%>
<!--#include file="../include/Footer.asp"-->
