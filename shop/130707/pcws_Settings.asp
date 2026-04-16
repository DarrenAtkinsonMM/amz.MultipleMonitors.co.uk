<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<%
pageTitle=dictLanguage.Item(Session("language")&"_pcAppBtnSettings")
pageIcon="pcv4_icon_settings.png"
%>
<!--#include file="AdminHeader.asp"-->
<!--#include file="../htmleditor/editor.asp" -->
<%
pcPageName="pcws_Settings.asp"
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<link href="https://gitcdn.github.io/bootstrap-toggle/2.2.0/css/bootstrap-toggle.min.css" rel="stylesheet">

<script src="https://gitcdn.github.io/bootstrap-toggle/2.2.0/js/bootstrap-toggle.min.js"></script>

<%
query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    pcv_strUid = rs("pcPCWS_Uid")
    pcv_AuthToken = rs("pcPCWS_AuthToken")  
    pcv_strUsername = rs("pcPCWS_Username")  
    pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
End If
Set rs=nothing

pcv_strThisFeatureCode = getUserInput(request("fc"), 25)
%>

<div class="apps" data-ng-controller="appsCtrl">

    <!--#include file="pcws_Navigation.asp"--> 

    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <htmldiv content="myhtml">
                    <%
                    execute (pcf_dynamicInclude(pcf_getMappedFileAsString("../includes/apps/" & pcv_strThisFeatureCode & "/settings.asp")))
                    %>
                </htmldiv>
            </div>
        </div>
    </div>

</div>

<!--#include file="AdminFooter.asp"-->