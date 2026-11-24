<%
'==============================================================================
' view/include/Header.asp  -  공통 상단 (DOCTYPE ~ <main> 시작)
'------------------------------------------------------------------------------
' Code 파일에서 미리 채워두는 값
'   g_pageTitle   <title> 앞부분
'   g_pageDesc    meta description
'   g_menuCode    현재 1depth 메뉴 코드 (하이라이트)
'   g_subMenuCode 현재 2depth 메뉴 코드
'   g_bodyClass   <body class="">
'   g_extraHead   추가 head 태그
'==============================================================================
Dim hdTitle
If g_pageTitle = GLB_SITE_TITLE Or Len(g_pageTitle) = 0 Then
    hdTitle = GLB_SITE_TITLE
    If Len(GLB_SITE_SUBTITLE) > 0 Then hdTitle = hdTitle & " - " & GLB_SITE_SUBTITLE
Else
    hdTitle = g_pageTitle & " | " & GLB_SITE_TITLE
End If
%><!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="format-detection" content="telephone=no">
<title><%= HE(hdTitle) %></title>
<meta name="description" content="<%= HE(StringUtil.Cut(StringUtil.StripTags(g_pageDesc), 150)) %>">
<% If Len(GLB_SITE_KEYWORDS) > 0 Then %>
<meta name="keywords" content="<%= HE(GLB_SITE_KEYWORDS) %>">
<% End If %>
<meta name="author" content="<%= HE(GLB_OWNER_NAME) %>">
<meta property="og:type" content="website">
<meta property="og:site_name" content="<%= HE(GLB_SITE_TITLE) %>">
<meta property="og:title" content="<%= HE(hdTitle) %>">
<meta property="og:description" content="<%= HE(StringUtil.Cut(StringUtil.StripTags(g_pageDesc), 150)) %>">
<meta property="og:url" content="<%= HE(GLB_SITE_URL & GLB_CUR_URL) %>">
<link rel="canonical" href="<%= HE(GLB_SITE_URL & GLB_PATH_INFO) %>">
<link rel="stylesheet" href="<%= APP_ROOT %>/assets/css/style.css?v=<%= HE(GLB_ASSET_VER) %>">
<script>
/* 다크모드 깜빡임(FOUC) 방지 : CSS 적용 직후 즉시 테마 확정 */
(function(){try{var t=localStorage.getItem('jkw-theme');if(t==='dark'||t==='light'){document.documentElement.setAttribute('data-theme',t);}}catch(e){}})();

/* assets/js/common.js 는 4개 스택이 공유한다. 서버 주소는 스택마다 다르므로
   이 스택의 엔드포인트를 여기서 알려준다. (common.js 의 ep() 가 읽는다) */
window.JKW = { urls: {
    phoneAuth:     '<%= JE(U("/member/PhoneAuth_p.asp")) %>',
    donateOrder:   '<%= JE(U("/donate/Order_p.asp")) %>',
    donateApprove: '<%= JE(U("/donate/Approve_p.asp")) %>',
    donateCancel:  '<%= JE(U("/donate/Cancel_p.asp")) %>',
    inicisPay:     '<%= JE(U("/donate/InicisPay.asp")) %>?order=',
    payResult:     '<%= JE(U("/donate/PayResult.asp")) %>?order='
} };
</script>
<%= g_extraHead %>
</head>
<body class="<%= HE(g_bodyClass) %>">
<a class="skip-link" href="#content">본문 바로가기</a>

<header id="header">
    <div class="wrap header-inner">

        <a href="<%= U("/main/Default.asp") %>" class="logo">
            <span><%= HE(GLB_SITE_TITLE) %><b class="dot">.</b></span>
            <% If Len(GLB_SITE_SUBTITLE) > 0 Then %><span class="sub"><%= HE(GLB_SITE_SUBTITLE) %></span><% End If %>
        </a>

        <!--#include file="Gnb.asp"-->

        <div class="header-util">
            <button type="button" class="btn btn-sm btn-ghost theme-toggle" aria-label="다크 모드로 전환" title="테마 전환">◐</button>
<% If GLB_IS_LOGIN Then %>
            <span class="util-name"><%= HE(GLB_MY_NAME) %></span>
            <a class="util-link" href="<%= U("/member/Mypage.asp") %>">마이페이지</a>
    <% If GLB_MY_LEVEL >= LV_ADMIN Then %>
            <a class="util-link" href="<%= U("/admin/Default.asp") %>">관리자</a>
    <% End If %>
            <a class="util-link" href="<%= U("/member/Logout.asp") %>">로그아웃</a>
<% Else %>
            <a class="util-link" href="<%= U("/member/Login.asp") %>">로그인</a>
            <a class="util-link" href="<%= U("/member/Join.asp") %>">회원가입</a>
<% End If %>
            <button type="button" class="nav-toggle" aria-label="메뉴 열기" aria-expanded="false" aria-controls="gnb"><span></span></button>
        </div>

    </div>
</header>

<main id="content">
