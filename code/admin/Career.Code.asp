<!--#include file="Admin.Code.asp"-->
<%
'==============================================================================
' code/admin/Career.Code.asp  -  경력 · 학력 · 수상 · 자격 관리
'==============================================================================
g_adminNav  = "career"
g_pageTitle = "경력 · 학력 관리"

Dim acArr, acCnt
acArr = About.CareerAll_l()
acCnt = RowCnt(acArr)

'--- CareerAll_l 추가 컬럼 ---
Const CRA_SORTNO = 9
Const CRA_ISUSE  = 10

'--- 편집 대상 ---
Dim acId, acEdit, acIsEdit
acId     = RequestUtil.ValueIntInit("id", 0)
acIsEdit = False

Dim eType, eTitle, ePosition, eOrgUrl, eDesc, eStart, eEnd, eCurrent, eSortNo, eIsUse
eType = "work" : eTitle = "" : ePosition = "" : eOrgUrl = "" : eDesc = ""
eStart = "" : eEnd = "" : eCurrent = False : eSortNo = 10 : eIsUse = True

If acId > 0 Then
    acEdit = About.Career_v(acId)
    If IsArray(acEdit) Then
        acIsEdit  = True
        eType     = AV(acEdit,  CR_TYPE, 0)
        eTitle    = AV(acEdit,  CR_TITLE, 0)
        ePosition = AV(acEdit,  CR_POSITION, 0)
        eOrgUrl   = AV(acEdit,  CR_ORGURL, 0)
        eDesc     = AV(acEdit,  CR_DESC, 0)
        eStart    = DateUtil.Fmt(AV(acEdit, CR_START, 0), "ymd")
        eEnd      = DateUtil.Fmt(AV(acEdit, CR_END, 0), "ymd")
        eCurrent  = AVB(acEdit, CR_CURRENT, 0)
        eSortNo   = AVI(acEdit, CRA_SORTNO, 0)
        eIsUse    = AVB(acEdit, CRA_ISUSE, 0)
    End If
End If

Dim acTypes, acTypeNames
acTypes     = Array("work", "edu", "cert", "award", "activity")
acTypeNames = Array("경력", "학력", "자격증", "수상", "활동")

'<Desc>career_type → 한글 라벨</Desc>
'<Desc>※ 함수명을 TypeName 으로 하면 VBScript 내장 함수와 충돌한다 (Name redefined)</Desc>
Function CareerTypeName(ByVal t)
    Select Case t
        Case "work"     : CareerTypeName = "경력"
        Case "edu"      : CareerTypeName = "학력"
        Case "cert"     : CareerTypeName = "자격증"
        Case "award"    : CareerTypeName = "수상"
        Case "activity" : CareerTypeName = "활동"
        Case Else       : CareerTypeName = t
    End Select
End Function
%>
