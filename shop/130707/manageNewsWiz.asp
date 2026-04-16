<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% pageTitle="Newsletter Wizard" %>
<% section="mngAcc" %>
<!--#include file="AdminHeader.asp"-->
<%'Start SDBA
pcv_pageType=request("pagetype")
'End SDBA
%>
<table class="pcCPcontent">
	<tr>
		<td>
		<p>The Newsletter Wizard allows you obtain a list of <%if pcv_pageType="0" then%>suppliers<%else%><%if pcv_pageType="1" then%>drop-shippers<%else%>customers<%end if%><%end if%> using a number of filters, and then export the list or send a message within ProductCart. You can also use a previously sent message to send a new message to the same list.
		<ul>
		<li><a href="<%if pcv_pageType<>"" then%>sds_newsWizStep1.asp?pagetype=<%=pcv_pageType%><%else%>newsWizStep1.asp<%end if%>">Start the Wizard</a> to create a new message</li>
		<li>View <a href="manageNews.asp?pagetype=<%=pcv_pageType%>">previously sent messages</a></li>
		</ul>
		</td>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
    <th>NO SPAM</th>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
    <p>DO NOT USE this feature to send SPAM e-mail. Regardless of whether or not SPAM is considered illegal in your State or Country, sending unsolicited messages is not what this feature is meant for. It is also not a good marketing practice and it will harm your business in the long run.</p>
    <p><strong>US STORES: You must comply with  <a href="http://www.ftc.gov/spam/" target="_blank">CAN-SPAM law</a></strong></p>
    <p>Make sure that you comply with the CAN-SPAM.  <a href="http://www.ftc.gov/spam/" target="_blank">Click here for more details</a>. Failure to comply could result in fines and possible imprisonment. In a nutshell, all commercial e-mail messages:</p>
    <ul>
        <li>Must not present misleading information in the From field or header information.</li>
        <li>Must include a link for and honor unsubscribe requests.</li>
        <li>Must conspicuously state that all commercial, promotional mail is an advertisement, unless all recipients have opted in </li>
        <li>Must Include a valid, physical mailing address (postal address) in all email campaigns.</li>
    </ul>
    </td>
  </tr>

	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
    <th>Not a professional e-mail marketing system</th>
	</tr>
	<tr>
    <td class="pcCPspacer"></td>
	</tr>
	<tr>
		<td>
            <p>Please keep in mind that the ProductCart Newsletter Wizard is not a professional e-mail marketing system and is not intended to handle large e-mail lists. There are a number of reasons why it might make sense for you to upgrade to a more robust and feature-rich e-mail marketing system</p>
    <ul>
      <li>Double-opt in mechanism (subscription pending until confirmed by e-mail)</li>
      <li>Message tracking (reads, opens, clicks, etc.)</li>
      <li>Separate subscription management for different multiple lists (e.g. 'Product Updates' vs. 'Specials and Promotions')</li>
      <li>List-specific, one-click unsubscribe and bounced messages management</li>
      <li>Robust infrastructure to send a message to a high number of recipients</li>
     </ul>

     <p>&nbsp;</p>
    </td>
	</tr>
</table>
<!--#include file="AdminFooter.asp"-->
