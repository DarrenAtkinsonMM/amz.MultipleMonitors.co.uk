<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=10%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/sendmail.asp"--> 
<% response.buffer=true %>
<% pageTitle="Reply to Customer's Quote" %>
<% Section="genRpts" %>
<% if request("action")="post" then
	pemail=request("toemail")
	msubject=request("subject")
	mbody=request("messageText")
	mbody=replace(mbody,"<XMP>","")
	mbody=replace(mbody,"</XMP>",vbcrlf)
	mbody=replace(mbody,"<PRE>","")
	mbody=replace(mbody,"</PRE>",vbcrlf)
	mbody=replace(mbody,"<xmp>","")
	mbody=replace(mbody,"</xmp>",vbcrlf)
	mbody=replace(mbody,"<pre>","")
	mbody=replace(mbody,"</pre>",vbcrlf)
	mbody=replace(mbody,"<BR>",vbcrlf)
	mbody=replace(mbody,"<br>",vbcrlf)
	session("News_MsgType")="0"
	call sendmail (scCompanyName, scEmail, pemail, msubject, mbody)
	%>
	<!--#include file="AdminHeader.asp"-->
    <div class="pcCPmessageSuccess">Your message was sent successfully. <a href="srcQuotesa.asp?datefrom=<%=request("datefrom")%>&dateto=<%=request("dateto")%>&idcustomer=<%=request("idcustomer")%>">Return to quotes</a>.</div>
	<!--#include file="AdminFooter.asp"-->
<% end if %>
