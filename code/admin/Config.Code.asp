<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Config.Code.asp  -  사이트 설정 (tb_config)
'==============================================================================
g_adminNav  = "config"
g_pageTitle = "사이트 설정"

Dim cfArr, cfCnt
cfArr = Config.Config_l("")
cfCnt = RowCnt(cfArr)

'--- Config_l 인덱스 ---
Const CF_KEY   = 0
Const CF_VALUE = 1
Const CF_GROUP = 2
Const CF_DESC  = 3

'--- 그룹 목록 (중복 제거) ---
Dim cfGroups, cfGroupCnt, gi2, gk, found, gj2
cfGroups   = Array("", "", "", "", "", "", "", "")
cfGroupCnt = 0
For gi2 = 0 To cfCnt - 1
    gk    = AV(cfArr, CF_GROUP, gi2)
    found = False
    For gj2 = 0 To cfGroupCnt - 1
        If cfGroups(gj2) = gk Then found = True
    Next
    If Not found And cfGroupCnt < 8 Then
        cfGroups(cfGroupCnt) = gk
        cfGroupCnt = cfGroupCnt + 1
    End If
Next

Function GroupName(ByVal g)
    Select Case g
        Case "site"   : GroupName = "사이트 기본"
        Case "seo"    : GroupName = "SEO / 메타"
        Case "mail"   : GroupName = "메일 발송"
        Case "upload" : GroupName = "업로드"
        Case Else     : GroupName = g
    End Select
End Function

'<Desc>여러 줄 입력이 필요한 키</Desc>
Function IsLongKey(ByVal k)
    IsLongKey = (InStr(",site_desc,hero_sub,about_intro,resume_summary,copyright,", "," & k & ",") > 0)
End Function

'<Desc>비밀번호처럼 가려야 하는 키</Desc>
Function IsSecretKey(ByVal k)
    IsSecretKey = (k = "mail_pass")
End Function
%>
