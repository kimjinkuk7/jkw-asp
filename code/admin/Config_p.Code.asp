<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Config_p.Code.asp  -  사이트 설정 저장
'------------------------------------------------------------------------------
' cfg_<key> 형태로 넘어온 폼 필드만 반영한다.
' 존재하지 않는 key 는 무시되므로 임의 삽입이 불가능하다.
'==============================================================================

If Not RequestUtil.IsPost() Then JsAlert "잘못된 접근입니다."
Auth.CheckToken Trim(Request.Form("_csrf") & "")

Dim cfgArr, cfgCnt, i, k, formName, newVal, changed
cfgArr  = Config.Config_l("")
cfgCnt  = RowCnt(cfgArr)
changed = 0

For i = 0 To cfgCnt - 1
    k        = AV(cfgArr, 0, i)
    formName = "cfg_" & k

    '--- 폼에 해당 필드가 있을 때만 갱신 ---
    If Len(Request.Form(formName) & "") > 0 Or Request.Form(formName).Count > 0 Then
        newVal = Trim(Request.Form(formName) & "")
        If newVal <> (AV(cfgArr, 1, i) & "") Then
            Config.Config_u k, newVal
            changed = changed + 1
        End If
    End If
Next

JsAlertGo changed & "개 항목이 저장되었습니다.", U("/admin/Config.asp")
%>
