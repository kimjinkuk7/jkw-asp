<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Stat.Code.asp  -  접속 통계
'==============================================================================
g_adminNav  = "stat"
g_pageTitle = "접속 통계"

Dim stDays
stDays = RequestUtil.ValueIntInit("days", 30)
If stDays <> 7 And stDays <> 30 And stDays <> 90 Then stDays = 30

Dim stArr, stCnt, stDev, stDevCnt
stArr    = Visit.Stat_l(stDays)       : stCnt    = RowCnt(stArr)
stDev    = Visit.DeviceStat_l(stDays) : stDevCnt = RowCnt(stDev)

'--- 합계 / 최대값 ---
Dim stTotalPv, stTotalUv, stMax, si
stTotalPv = 0 : stTotalUv = 0 : stMax = 1
For si = 0 To stCnt - 1
    stTotalPv = stTotalPv + AVI(stArr, 1, si)
    stTotalUv = stTotalUv + AVI(stArr, 2, si)
    If AVI(stArr, 1, si) > stMax Then stMax = AVI(stArr, 1, si)
Next

Dim stAvg
If stCnt > 0 Then stAvg = Int(stTotalPv / stCnt) Else stAvg = 0

'--- 기기 합계 ---
Dim stDevTotal, di
stDevTotal = 0
For di = 0 To stDevCnt - 1
    stDevTotal = stDevTotal + AVI(stDev, 1, di)
Next
If stDevTotal < 1 Then stDevTotal = 1
%>
