<% If scEnableBundling = "1" Then %>

    <link type="text/css" rel="stylesheet" href="<%=pcf_getCSSPath("css","combined.min.css")%>" />

<% Else %>

    <!-- Removed PC ref to /shop/pc/css/bootstrap.min.css to avoid double file load with own below -->
    
    <% If pcStrPageName = "configureprd.asp" Or _
          pcStrPageName = "reconfigure.asp" Or _
          pcStrPageName = "prdaddcharges.asp" Or _
          pcStrPageName = "reprdaddcharges.asp" Or _
          InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "configureprd.asp") Or _
          InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "reconfigure.asp") Or _
          InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "prdaddcharges.asp") Or _
          InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "reprdaddcharges.asp") Then %>   
        <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","animate.css")%>" />
        <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","configuratorPricingBox.css")%>" />
        <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","pcBTO.css")%>" />
        <% If statusCM="1" Then %>
        <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","pcCM.css")%>" />
        <% End If %>
        
    <% End If %>

    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","datepicker3.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","facebook.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","screen.css")%>" />
    <!-- Removed pcstorefront.css reference -->
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","quickview.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","pcSearchFields.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","search.min.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","bulkaddtocart.css")%>" />
    <% If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "onepagecheckout.asp") Then %>
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","onepagecheckout.css")%>" />
    <% End If%>
    <% If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "ordercomplete.asp") Then %>
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","pcStoreFront.css")%>" />
    <% End If%>
    <% If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "custviewpastd.asp") Then %>
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("css","pcStoreFront.css")%>" />
    <% End If%>
        
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("../includes/mojozoom","mojozoom.css")%>" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("../htmleditor/scripts/style","awesome.css")%>" media="screen" />
    <link type="text/css" rel="stylesheet" href="/shop/pc/<%=pcf_getCSSPath("../includes/javascripts/flickity","flickity.min.css")%>" media="screen" />

<!-- DA - EDIT -->
    <link href="/css/bootstrap.min.css" rel="stylesheet" type="text/css">
    <link href="/css/font-awesome.min.css" rel="stylesheet" type="text/css" />
	<link href="/css/ekko-lightbox.min.css" rel="stylesheet" />
	<link href="/css/animate.css" rel="stylesheet" />
    <link href="/css/style.css" rel="stylesheet">
    <link href="/css/responsive.css" rel="stylesheet">

	<link href="/css/blue.css" rel="stylesheet">


<% End If %>
