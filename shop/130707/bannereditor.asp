<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% 
pageTitle = "Banner Management - Edit" 
pageIcon = ""
pcStrPageName = "bannereditor.asp"
%>
<%PmAdmin=0%>


<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"-->
<!--#include file="AdminHeader.asp"-->
<!-- #include file="../htmleditor/editor.asp" -->

<style>
  .inputRow{
    margin-top:20px;
	line-height:24px;
  }
  .col-xs-3.txt {
  	text-align:right;
	padding-right:0;
	line-height:34px;
  }
  .in {
  	width:80%;
  }
  .inputRow .btn {
  float:left;
  margin-right:20px;
  }
</style>

<%

bid = Request.QueryString("bid")

if request("action")="new" then
  startdate = request.form("startdate")
  enddate = request.form("enddate")
  active = request.form("active")
  background = request.form("pcBannerColor")
  html = request("bannertext")
  query = "INSERT INTO mod_bannermanagement (startdate, enddate, active, background, html) VALUES ('" & startdate & "', '" & enddate & "', '" & active & "', '" & background & "', '" & html & "')"
  set rs=server.CreateObject("ADODB.Recordset")
  set rs=conntemp.execute(query)
  if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing
    call closedb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
  else
    response.redirect "bannermanagement.asp?action=new"
  end if

end if

if request("action")="upd" then
  bannerid = request.form("bid")
  startdate = request.form("startdate")
  enddate = request.form("enddate")
  active = request.form("active")
  background = request.form("pcBannerColor")
  html = request("bannertext")
  query = "UPDATE mod_bannermanagement SET startdate = '" & startdate & "', enddate = '" & enddate & "', active = '" & active & "', background = '" & background & "', html = '" & html & "' WHERE bannerid = " & bannerid
  set rs=server.CreateObject("ADODB.Recordset")
  set rs=conntemp.execute(query)
  if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing
    call closedb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
  else
    response.redirect "bannermanagement.asp?action=upd"
  end if
end if

if bid <> "" then
  query = "SELECT * FROM mod_bannermanagement WHERE bannerid = " & bid
  set rs=server.CreateObject("ADODB.Recordset")
  set rs=conntemp.execute(query)
  if err.number<>0 then
    call LogErrorToDatabase()
    set rs=nothing
    call closedb()
    response.redirect "techErr.asp?err="&pcStrCustRefID
  end if

  if NOT rs.eof then  
    bannerid = rs("bannerid")
    bannerstartdate = rs("startdate")
    bannerenddate = rs("enddate")
    banneractive = rs("active")
    bannerbackground = rs("background")
  bannerhtml = rs("html")
  else
    response.write "there is no banner :("
  end if
  set rs = nothing
end if

%>
<% if bid <> "" then %>
  <form action="bannereditor.asp?action=upd&bid=<%=bid%>" method="post">
<% else %>
  <form action="bannereditor.asp?action=new" method="post">
<% end if %>

<input type="hidden" value="<%=bannerid%>" name="bid">
<input type="hidden" value="1" name="active">


<div class="container-fluid">
  <div class="row">
	  <div class="col-xs-12">
 		 <p>Enter Start Date and End Date for your new banner.  You may create a text banner or image if you provide an image link.</p>
	  </div>
  </div>

  <div class="row  inputRow">
    <div class="col-xs-3 txt">Start Date (MM/DD/YYYY):</div>
    <div class="col-xs-3"><input type="text" value="<%=bannerstartdate%>" name="startdate" class="datepicker in"></div>
    <div class="col-xs-3 txt">End Date (MM/DD/YYYY):</div>
    <div class="col-xs-3"><input type="text" value="<%=bannerenddate%>" name="enddate" class="datepicker in"></div>
  </div>
  <div class="row  inputRow">
    <div class="col-xs-3 txt">Banner Background Color:</div>
    <div class="col-xs-3">
      <script src='../pc/js/spectrumjs/spectrum.js'></script>
      <link rel='stylesheet' href='../pc/js/spectrumjs/spectrum.css' />
      <input type='text' id="pcBannerColor" name="pcBannerColor" value="<%=bannerbackground%>" class="in"/>
      <script>
      $("#pcBannerColor").spectrum({
          preferredFormat: "hex",
          showInput: true,
          showPalette: true,
          palette: [["#ff0000", "#00ff00", "#0000ff", "<%=bannerbackground%>"]]
      });
      </script>
    </div>
  </div>
  <div class="row  inputRow">
    <div class="col-xs-12">
      <div class="editorcontainer" id="editorcontainer" name="editorcontainer">
        <textarea class="htmleditor" name="bannertext" id="bannertext" rows="6" cols="56" tabindex="103" ><%=bannerhtml%></textarea>
        <script language="javascript" type="text/javascript">
            window["oEdit1"].REPLACE("bannertext", "editorcontainer");
        </script>
      </div>
    </div>
  </div>
</div>
<div class="inputRow">
 <button class="btn btn-primary">Save</button>
</div>
</form>
<div class="inputRow">
<form action="bannermanagement.asp" method="post">
<button class="btn btn-Secondary">Close Without Saving</button>
</form>
</div>


<%
Call pcs_hookCPanelFooterJS()
%>
<!--#include file="AdminFooter.asp"-->