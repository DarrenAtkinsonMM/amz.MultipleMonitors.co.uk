<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<%
' Check if the store is on. If store is turned off display store message
 
%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->
<div id="pcMain">
	<div class="pcMainContent">
		<form name="Form1" action="ggg_srcGRb.asp?action=search" method="POST" class="pcForms">
			<div class="pcShowContent">
				<h1><%= dictLanguage.Item(Session("language")&"_SrcGR_1")%></h1>
				<div class="pcFormItem">
					<div class="pcFormItemFull"><%= dictLanguage.Item(Session("language")&"_SrcGR_8")%></div>
				</div>

				<div class="pcSpacer"></div>

				<% '// First Name %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_SrcGR_2")%></div>
					<div class="pcFormField"><input type=text name="cname" value="" size="30"></div>
				</div>

				<% '// Last Name %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_SrcGR_3")%></div>
					<div class="pcFormField"><input type=text name="clastname" value="" size="30"></div>
				</div>

				<% '// Registry Name %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_SrcGR_9")%></div>
					<div class="pcFormField"><input type=text name="cregname" value="" size="30"></div>
				</div>
				
				<div class="pcSpacer"></div>
				
				<% '// Event Date %>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%= dictLanguage.Item(Session("language")&"_SrcGR_4")%></div>
					<div class="pcFormField">
						<%= dictLanguage.Item(Session("language")&"_SrcGR_5")%>
						<select name="emonth">
							<option value="" selected><%= dictLanguage.Item(Session("language")&"_SrcGR_5")%></option>
							<option value="1"><%= dictLanguage.Item(Session("language")&"_SrcGR_11")%></option>
							<option value="2"><%= dictLanguage.Item(Session("language")&"_SrcGR_12")%></option>
							<option value="3"><%= dictLanguage.Item(Session("language")&"_SrcGR_13")%></option>
							<option value="4"><%= dictLanguage.Item(Session("language")&"_SrcGR_14")%></option>
							<option value="5"><%= dictLanguage.Item(Session("language")&"_SrcGR_15")%></option>
							<option value="6"><%= dictLanguage.Item(Session("language")&"_SrcGR_16")%></option>
							<option value="7"><%= dictLanguage.Item(Session("language")&"_SrcGR_17")%></option>
							<option value="8"><%= dictLanguage.Item(Session("language")&"_SrcGR_18")%></option>
							<option value="9"><%= dictLanguage.Item(Session("language")&"_SrcGR_19")%></option>
							<option value="10"><%= dictLanguage.Item(Session("language")&"_SrcGR_20")%></option>
							<option value="11"><%= dictLanguage.Item(Session("language")&"_SrcGR_21")%></option>
							<option value="12"><%= dictLanguage.Item(Session("language")&"_SrcGR_22")%></option>
						</select>
						&nbsp;
						<%= dictLanguage.Item(Session("language")&"_SrcGR_6")%>
						<select name="eyear">
							<option value="" selected><%= dictLanguage.Item(Session("language")&"_SrcGR_6")%></option>
							<option value="<%=year(date())%>"><%=year(date())%></option>
							<option value="<%=year(date())+1%>"><%=year(date())+1%></option>
							<option value="<%=year(date())+2%>"><%=year(date())+2%></option>
							<option value="<%=year(date())+3%>"><%=year(date())+3%></option>
						</select>
					</div>
				</div>

				<div class="pcSpacer"></div>
				
				<div class="pcFormButtons">
					
					<button class="pcButton pcButtonContinue" id="submit" name="Submit" value="<%= dictLanguage.Item(Session("language")&"_css_submit")%>">
						<img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_submit") %>">
						<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_submit") %></span>
					</button>

				</div>
			</div>
		</form>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
