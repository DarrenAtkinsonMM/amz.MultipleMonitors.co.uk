<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle = "Banner Management" 
pageIcon = ""
pcStrPageName = "bannermanagement.asp"
%>
<%PmAdmin=0%>


<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="AdminHeader.asp"-->
<style>
#pcCPmainLeft .btn {
margin-top:20px;
}
.glyphicon.glyphicon-remove{
	color:red;
}
.div.row :nth-child(even){
  background-color: #dcdcdc;
}
.div.row :nth-child(odd){
  background-color: #aaaaaa;
}
.row.header {
	font-weight:bold;
}
.row {
line-height:22px;
}
</style>
<%

FUNCTION stripHTML(strHTML)
  Dim objRegExp, strOutput, tempStr
  Set objRegExp = New Regexp
  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|n)+?>"
  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  'Replace all < and > with &lt; and &gt;
  strOutput = Replace(strOutput, "<", "&lt;")
  strOutput = Replace(strOutput, ">", "&gt;")
  stripHTML = strOutput    'Return the value of strOutput
  Set objRegExp = Nothing
END FUNCTION

query = "SELECT * FROM mod_bannermanagement"

set rsBanners=server.CreateObject("ADODB.Recordset")
set rsBanners=conntemp.execute(query)
If err.number<>0 Then
  call LogErrorToDatabase()
  set rsProducts=nothing
  call closedb()
  response.redirect "techErr.asp?err="&pcStrCustRefID
End If
If NOT rsBanners.eof Then
  bannersList = rsBanners.getRows()
  bannersCount = UBound(bannersList,2)
Else
  set rsBanners = nothing
  call closeDb()
  nobanner = 1
End If
set rsBanners = nothing

if request("action")="upd" then%>
  
  <div class="pcCPmessageSuccess">
    The banner has been updated.
  </div>


<%end if%>

<%if request("action")="del" then%>
  
<div class="pcCPmessageSuccess">
    The banner has been deleted.
  </div>


<%end if%>
<div class="container-fluid">
  <div class="row">
    <div class="col-xs-12">
	<p>Click the Banner Text to edit, <span class="glyphicon glyphicon-remove"></span> if you wish to delete the banner.</p>
	</div>
  </div>
  <div class="row">
    <div class="col-xs-12 line">
        <div class="row header">
          <div class="col-xs-4">Banner Text</div>
          <div class="col-xs-2">Start Date</div>
          <div class="col-xs-2">End Date</div>
          <div class="col-xs-1"></div>
        </div>      
        <% 
        if nobanner <> 1 then
          for i = 0 to bannersCount 
          %>
            <div class="row">
              <div class="col-xs-4" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"><a href="bannereditor.asp?bid=<%=bannersList(0,i)%>"><%=stripHTML(bannersList(5,i))%></a></div>
              <div class="col-xs-2"><%=bannersList(1,i)%></div>
              <div class="col-xs-2"><%=bannersList(2,i)%></div>
              <div class="col-xs-1"><a href="deletebanner.asp?bid=<%=bannersList(0,i)%>" onclick="if (!confirm('Delete this banner?')) { return false }"><span class="glyphicon glyphicon-remove"></span></a></div>
            </div>
          <% next %>
        <% else %>
		  
          <p>Currently you have no banners created.  Click the New Banner button below to create one.</p>
		
        <% end if %>
    </div>
  </div>
</div>

<div style="margin:10px;">
<form action="bannereditor.asp" method="post">
<button class="btn btn-primary">New Banner</button>
</form>
</div>

<%
Call pcs_hookCPanelFooterJS()
%>
<!--#include file="AdminFooter.asp"-->