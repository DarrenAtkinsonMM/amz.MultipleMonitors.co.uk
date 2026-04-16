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
<%
pageTitle=dictLanguage.Item(Session("language")&"_pcAppBtnMarket")
pageIcon="pcv4_icon_settings.png"
%>
<!--#include file="AdminHeader.asp"-->
<%
pcPageName="pcws_Market.asp"
%>

<style type="text/css">
.apps .navigation {
    margin-top: 1em;
    text-align: center;
    padding: 1em; 
}
.panel-title a {
 	display: block; 
}
</style>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

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

call pcf_UpdateToken()
%>

<div id="appsCtrl" class="apps" data-ng-controller="appsCtrl">

    <!--#include file="pcws_Navigation.asp"--> 

    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <htmldiv content="myhtml">
                    <%= pcf_displayMarket("", pcv_AuthToken) %>
                </htmldiv>
            </div>
        </div>
    </div>

</div>

<script>
// Local Variable Overrides
<%
query="SELECT [pcPCWS_FeatureCode], pcPCWS_IsActive FROM pcWebServiceFeatures;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    Do While Not rs.eof
        pcv_strFeatureCode = rs("pcPCWS_FeatureCode")
        pcv_intIsActive = rs("pcPCWS_IsActive")
        If pcv_intIsActive = 1 Then
            pcv_boolIsActive = true
        Else
            pcv_boolIsActive = false
        End If
        %>
        <%=pcv_strFeatureCode %> = <%=lcase(pcv_boolIsActive) %>;
        <%
        rs.movenext
    Loop
End If
Set rs=nothing
%>
</script>

<!--#include file="AdminFooter.asp"-->