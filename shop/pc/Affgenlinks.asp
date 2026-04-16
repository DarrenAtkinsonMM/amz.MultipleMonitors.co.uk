<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="AffLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/SocialNetworkWidgetConstants.asp"-->
<!--#include file="../includes/pcSeoFunctions.asp"-->
<%
' Load affiliate ID
affVar=session("pc_idaffiliate")
if not validNum(affVar) then
	response.redirect "AffiliateLogin.asp"
end if
%>
<!--#include file="header_wrapper.asp"-->
<%sMode=request("action")
	if sMode <> "" then
		sMode="1"
		idproduct=request.Form("product")
		idaffiliate=session("pc_IDAffiliate")
	end If %><%
Dim strSQL

%>
<script type=text/javascript>
var copytoclip=1

function HighlightAll(theField) {
	var tempval=eval("document."+theField)
	tempval.focus()
	tempval.select()
	if (document.all&&copytoclip==1){
		therange=tempval.createTextRange()
		therange.execCommand("Copy")
		window.status="Contents highlighted and copied to clipboard!"
		setTimeout("window.status=''",1800)
	}
}
</script>
<div id="pcMain">
	<div class="pcMainContent">
		<h1><%=dictLanguage.Item(Session("language")&"_AffgenLinks_1")%></h1>

		<form method="post" name="links" action="Affgenlinks.asp?action=1" class="pcForms">
			<div class="pcShowContent">
				<div class="pcSpacer"></div>
				<div class="pcFormItem">
					<strong><%=dictLanguage.Item(Session("language")&"_AffgenLinks_8")%></strong>
				</div>
				<div class="pcSpacer"></div>
				<div class="pcFormItem">
					<div class="pcFormLabel"><%=dictLanguage.Item(Session("language")&"_AffgenLinks_2")%></div>
					<div class="pcFormField">
						<select name="product">  
							<%
							query="SELECT idproduct,description FROM products WHERE active=-1 AND configOnly=0 AND removed=0 ORDER BY description ASC"
							set rsPrd=Server.CreateObject("adodb.recordset")
							set rsPrd=conntemp.execute(query)
							if err.number <> 0 then
    							call LogErrorToDatabase()
    							set rsPrd = Nothing
    							call closeDb()
    							response.redirect "techErr.asp?err="&pcStrCustRefID
							end If
								
								do until rsPrd.eof
								intTempIdProduct=rsPrd("idproduct")
								strTempDescription=rsPrd("description")
								if sMode="1" And Cint(idproduct)= Cint(intTempIdProduct) then 
									pDescription=strTempDescription %>
									<option value="<%= intTempIdProduct%>" selected> 
								<% else %>
									<option value="<%= intTempIdProduct%>"> 
								<% end if %>
								<%=strTempDescription%>
								</option>
								<%
								rsPrd.movenext
								loop
								set rsPrd=nothing
							%>
							</select>
					</div>
				</div>
						
				<% If sMode="1" then						
							
					query="SELECT idproduct, description FROM products WHERE idproduct="&idproduct
					set rsPrd=Server.CreateObject("adodb.recordset")
					set rsPrd=conntemp.execute(query)
					pProductDesc=rsPrd("description")
					set rsPrd=nothing
						
					'// SEO Links
					'// Build Navigation Product Link
					'// Get the first category that the product has been assigned to, filtering out hidden categories
					query="SELECT categories_products.idCategory FROM categories_products INNER JOIN categories ON categories_products.idCategory = categories.idCategory WHERE categories_products.idProduct="& idproduct &" AND categories.iBTOhide<>1 AND categories.pccats_RetailHide<>1"
					set rs=server.CreateObject("ADODB.RecordSet")
					set rs=conntemp.execute(query)
					if not rs.EOF then
						pIdCategory=rs("idCategory")
					else
						pIdCategory=1
					end if
					set rs=nothing

					if scSeoURLs=1 then
						pcStrPrdLink=pProductDesc & "-" & pIdCategory & "p" & idproduct & ".htm"
						pcStrPrdLink=removeChars(pcStrPrdLink)
						pcStrPrdLink=pcStrPrdLink & "?"
					else
						pcStrPrdLink="viewPrd.asp?idproduct=" & idproduct &"&"
					end if
					'//
							
					tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/"&pcStrPrdLink&"idaffiliate="&idaffiliate),"//","/")
					tempURL=replace(tempURL,"http:/","http://")
					tempURL=replace(tempURL,"https:/","https://")
						
				%>
					<div class="pcSpacer"></div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<%=dictLanguage.Item(Session("language")&"_AffgenLinks_4")%>&nbsp;<strong><%=pProductDesc%></strong><%=dictLanguage.Item(Session("language")&"_AffgenLinks_5")%><%=pAffiliateName%>:
						</div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<input type="text" name="link1" size="80" value="<%=tempURL%>">
							<a class="highlighttext" href="javascript:HighlightAll('links.link1')"><img src="<%=pcf_getImagePath("images","edit2.gif")%>" alt="Highlight All" style="width: 25px; height: 23px;"></a>
						</div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<%=dictLanguage.Item(Session("language")&"_AffgenLinks_6")%><%=pAffiliateName%>:
						</div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<%
							tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/home.asp?idaffiliate="&idaffiliate),"//","/")
							tempURL=replace(tempURL,"http:/","http://")
							tempURL=replace(tempURL,"https:/","https://")
							%>
							<input type="text" name="link2" size="80" value="<%=tempURL%>">
							<a class="highlighttext" href="javascript:HighlightAll('links.link2')"><img src="<%=pcf_getImagePath("images","edit2.gif")%>" alt="Highlight All" style="width: 25px; height: 23px;"></a>
						</div>
					</div>

				<% end if %> 
                        
				<% If SNW_AFFILIATE="1" then %>
					<div class="pcSpacer"></div>
					<div class="pcFormItem">
						<strong><%=dictLanguage.Item(Session("language")&"_AffgenLinks_9")%></strong>
					</div>
					<div class="pcSpacer"></div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<%=dictLanguage.Item(Session("language")&"_AffgenLinks_7")%>:
						</div>
					</div>

					<div class="pcFormItem">
						<div class="pcFormFull">
							<%
								tempURL=replace((scStoreURL&"/"&scPcFolder),"//","/")
								tempURL=replace(tempURL,"http:/","http://")
								tempURL=replace(tempURL,"https:/","https://")
								tempCode="<script type=text/javascript>idaffiliate="""& session("pc_IDAffiliate") &""";</script><script type=text/javascript src="""&tempURL&"/pc/pcSyndication.js""></script>"									
							%>
							<textarea name="link3" cols="50" rows="10"><%=tempCode%></textarea>		
							<a class="highlighttext" href="javascript:HighlightAll('links.link3')"><img src="<%=pcf_getImagePath("images","edit2.gif")%>" alt="Highlight All" style="width: 25px; height: 23px; vertical-align: inherit"></a>
						</div>
					</div>
				<% end if %> 
			</div>
				
			<div class="pcSpacer"></div>

			<div class="pcFormButtons">
				<button class="pcButton pcButtonGenerateLinks" name="submit1" id="submit" value="<%=dictLanguage.Item(Session("language")&"_AffgenLinks_3")%>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_AffgenLinks_3") %></span>
				</button>

				<a class="pcButton pcButtonBack" href="javascript:history.back(-1);">
					<img src="<%=pcf_getImagePath("",rslayout("back"))%>" alt="<%= dictLanguage.Item(Session("language")&"_css_back") %>">
					<span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_back") %></span>
				</a>
			</div>
		</form>
	</div>
</div>
<!--#include file="footer_wrapper.asp"-->
