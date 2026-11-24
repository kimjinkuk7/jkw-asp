<%@ LANGUAGE="VBScript" CODEPAGE="65001" %>
<% Option Explicit %>
<!--#include file="../../code/admin/Config.Code.asp"-->
<!--#include file="../include/Header.asp"-->

<div class="wrap">
    <div class="admin-layout">
        <!--#include file="../include/AdminSide.asp"-->

        <div class="admin-main">

            <div class="page-head">
                <h2>사이트 설정</h2>
                <p class="desc">
                    <code>tb_config</code> 의 key/value 를 편집합니다.
                    새 설정 항목이 필요하면 DB 에 행을 추가하면 이 화면에 자동으로 나타납니다.
                </p>
            </div>

            <form method="post" action="<%= U("/admin/Config_p.asp") %>">
                <%= Auth.TokenField() %>

<%      Dim gi, ci
        For gi = 0 To cfGroupCnt - 1 %>
                <div class="admin-card">
                    <h3><%= HE(GroupName(cfGroups(gi))) %></h3>
                    <div class="form-table">
<%          For ci = 0 To cfCnt - 1
                If AV(cfArr, CF_GROUP, ci) = cfGroups(gi) Then %>
                        <div class="row">
                            <label class="form-label" for="cf-<%= AVH(cfArr, CF_KEY, ci) %>">
                                <%= AVH(cfArr, CF_DESC, ci) %>
                                <span class="d-block fs-xs muted" style="font-weight:400"><code><%= AVH(cfArr, CF_KEY, ci) %></code></span>
                            </label>
                            <div>
<%              If IsSecretKey(AV(cfArr, CF_KEY, ci)) Then %>
                                <input class="input" type="password" id="cf-<%= AVH(cfArr, CF_KEY, ci) %>"
                                       name="cfg_<%= AVH(cfArr, CF_KEY, ci) %>" value="<%= AVH(cfArr, CF_VALUE, ci) %>"
                                       autocomplete="new-password">
<%              ElseIf IsLongKey(AV(cfArr, CF_KEY, ci)) Then %>
                                <textarea class="textarea" id="cf-<%= AVH(cfArr, CF_KEY, ci) %>"
                                          name="cfg_<%= AVH(cfArr, CF_KEY, ci) %>" rows="3"
                                          style="min-height:80px"><%= AVH(cfArr, CF_VALUE, ci) %></textarea>
<%              Else %>
                                <input class="input" type="text" id="cf-<%= AVH(cfArr, CF_KEY, ci) %>"
                                       name="cfg_<%= AVH(cfArr, CF_KEY, ci) %>" value="<%= AVH(cfArr, CF_VALUE, ci) %>">
<%              End If %>
                            </div>
                        </div>
<%              End If
            Next %>
                    </div>
                </div>
<%      Next %>

                <div class="btn-row right mt-4">
                    <button type="submit" class="btn btn-primary btn-lg">설정 저장</button>
                </div>
            </form>

            <div class="admin-card mt-5">
                <h3>참고</h3>
                <ul class="fs-sm sub" style="list-style:disc;padding-left:1.3em">
                    <li><code>asset_ver</code> 값을 바꾸면 CSS/JS 캐시가 강제로 갱신됩니다. (배포 후 스타일이 안 바뀔 때)</li>
                    <li><code>mail_use</code> 를 <code>1</code> 로 두면 문의 접수 시 CDO 로 메일을 보냅니다. SMTP 계정 설정이 필요합니다.</li>
                    <li>DB 접속 정보와 디버그 모드는 보안상 DB 가 아니라 <code>App.Conf.asp</code> 에 있습니다.</li>
                </ul>
            </div>

        </div>
    </div>
</div>

<!--#include file="../include/Footer.asp"-->
