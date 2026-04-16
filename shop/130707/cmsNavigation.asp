<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=11%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSeoFunctions.asp"-->

<%
pageTitle="Generate Content Pages Navigation"
dim pcvParentPageName, pcArray, pcInt_Parent, pcInt_Published, pcInt_Inactive, pcInt_Active, pcv_IntNumrowsCount, pcv_IntNumrows, m, n
%>

<!--#include file="AdminHeader.asp"-->
<%


' Load Pages
query="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 AND pcCont_InActive=0 AND pcCont_MenuExclude=0 ORDER BY pcCont_Order, pcCont_Parent, pcCont_PageName ASC;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error loading parent content pages") 
end If
%>



<%
'Check URL

SPath1=Request.ServerVariables("PATH_INFO")
strHttps = ucase(Request.ServerVariables("HTTPS"))
mycount1=0
do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
		mycount1=mycount1+1
	end if
	if mycount1<2 then
		SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
loop
if strHttps="ON" then
	strURLPrefix="https://"
else
	strURLPrefix="http://"
end if
SPathInfo=strURLPrefix & Request.ServerVariables("HTTP_HOST") & SPath1

if Right(SPathInfo,1)="/" then
	pcv_strViewContents=SPathInfo & "pc/viewContent.asp"
else
	pcv_strViewContents=SPathInfo & "/pc/viewContent.asp"
end if
%>

<form name="form1" action="cmsNavigation.asp" method="post" class="pcForms">
	<table class="pcCPcontent">
    
		<%
		'-------------------------
		' NO Content Pages Found
		'-------------------------
		IF rstemp.eof THEN
			set rstemp=nothing
		%>
			<tr> 
				<td align="center">
					<div class="pcCPmessage">No Content Pages Found. <a href="cmsAddEdit.asp">Add New</a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=436"></a></div>
				</td>
			</tr>                  
		<% 
		ELSE
		
		'-------------------------
		' NO FORM Submitted
		'-------------------------	
			
			IF request("submit")="" THEN
		%>
			<tr> 
				<td colspan="2">
                	<h2>Which pages are included</h2>
                	The system will generate an unordered list with all content pages that are:
                	<ul>
                    	<li>Active</li>
                        <li>Not excluded from the navigation</li>
                    </ul>
				</td>
			</tr>  
			<tr> 
				<td colspan="2">
                	<h2>JQuery Navigation</h2>
				</td>
			</tr>  
			<tr> 
				<td valign="top">Prepare for JQuery menu bar:<br /><a href="http://www.dynamicdrive.com/dynamicindex1/ddsmoothmenu.htm" target="_blank">Examples</a>, <a href="http://www.dynamicdrive.com/dynamicindex1/ddsmoothmenu.htm" target="_blank">Documentation</a></td>
                <td>
                <input type="radio" name="JQNav" value="1"> JQuery Horizontal Menu Bar <br />
                <input type="radio" name="JQNav" value="2"> JQuery Vertical Menu Bar <br />
                <input type="radio" name="JQNav" value="0" checked> None
                </td>
			</tr>
			<tr> 
				<td>ID of JQuery menu bar:</td>
                <td>
                <input type="text" name="JQNavID" value="menubar1" size="30"> <span class="pcSmallText">See JQuery documentation for details.</span>
                </td>
			</tr>
			<tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
			<tr> 
				<td colspan="2">
                	<h2>JQuery Navigation Display Settings</h2>
                    You can assign custom background corlors to JQuery menu bar,  one for the default state, the other, when the mouse rolls over the menu items.
				</td>
			</tr>
			<tr> 
				<td nowrap>'Default State' background color:</td>
                <td><input type="text" name="menubg" value="" size="30">
			</tr>
			<tr> 
				<td nowrap>'Active State' background color:</td>
                <td><input type="text" name="menubga" value="" size="30">
			</tr>
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
			<tr> 
				<td colspan="2">
                	<h2>Advanced Settings</h2>
                    You can assign a CSS class to the most relevant elements in the unordered list. These settings are <u>ignored</u> if using the JQuery menu option above.
				</td>
			</tr>  
			<tr> 
				<td nowrap>&lt;UL&gt; Tag ID</td>
                <td><input type="text" name="ulid" value="" size="30"> <span class="pcSmallText">This is often referenced in JavaScript used to activate the menu.</span>
			</tr>
			<tr> 
				<td nowrap>&lt;UL&gt; CSS Class</td>
                <td><input type="text" name="ulclass" value="" size="30">
			</tr> 
			<tr> 
				<td nowrap>Top-level &lt;LI&gt; CSS Class</td>
                <td><input type="text" name="liclass" value="" size="30">
			</tr> 
			<tr> 
				<td nowrap>Top-level &lt;LI&gt; with Sub-Items CSS Class</td>
                <td><input type="text" name="lisbclass" value="" size="30">
			</tr>  
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
            <tr>
            	<td colspan="2"><input type="submit" name="submit" value="Generate Content Pages Navigation" onClick="return(confirm('You are about to overwrite the existing Content Pages navigation with a new list of pages. Back up the file pc/cmsNavigationLinks.inc if you need to keep a copy of the existing navigation. Are you sure you want to continue?'));" class="btn btn-primary"></td>
            </tr>
		
		<%
			ELSE
			'-------------------------
			' BUILD Content Pages Navigation
			'-------------------------
			
			pcIntJQNav = request("JQNav") ' 1 = horizontal, 2 = vertical
			pcvULID = request("JQNavID")
			if not validNum(pcIntJQNav) or pcIntJQNav=0  then
				pcvULclass = request("ulclass")
				pcvULID = request("ulid")
				pcvLIclass = request("liclass")
				pcvLIsbclass = request("lisbclass")
			else
				if pcIntJQNav = 1 then
					pcvULclass = "ddsmoothmenu"
				else
					pcvULclass = "ddsmoothmenu-v"
				end if					
				pcvLIclass = ""
				pcvLIsbclass = ""
				pcv_menubg=request("menubg")
				pcv_menubga=request("menubga")
			end if
			
			strNavigationStart = "<div id='" & pcvULID & "' class='" & pcvULclass & "'>" & vbcrlf & "<ul>"
			strNavigationSubMenu = "<ul>"
			strNavigationEnd = "</ul>"
			
			pcArray = rstemp.getRows()
			set rstemp=nothing
			
			pcv_IntNumrows = UBound(pcArray, 2)

			pcv_IntNumrowsCount=0
			
			FOR m = 0 to pcv_IntNumrows

				pcv_IntNumrowsCount=pcv_IntNumrowsCount+1
				pcv_lngIDPage= pcArray(0,m)
				pcv_strPageName = pcArray(1,m)
				
				'// Check to see if there are subpages
				query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage & " AND pcCont_InActive=0 AND pcCont_MenuExclude=0;"
				set rstemp = Server.CreateObject("ADODB.Recordset")
				set rstemp = conntemp.execute(query)
				if err.number <> 0 then
					set rstemp=nothing
					
					call closeDb()
					response.redirect "techErr.asp?error="& Server.Urlencode("Error loading child pages") 
				end If
				if not rstemp.eof then
					pcIntHasSubPages = 1
					else
					pcIntHasSubPages = 0
				end if
				
				'// SEO Links
				'// Build Navigation Product Link
				if scSeoURLs=1 then
					if pcIntHasSubPages = 1 then
						pcStrCntPageLink=pcv_strPageName & "-e" & pcv_lngIDPage & ".htm"
					else
						pcStrCntPageLink=pcv_strPageName & "-d" & pcv_lngIDPage & ".htm"
					end if
					pcStrCntPageLink=removeChars(pcStrCntPageLink)
					pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink
				else
					pcStrCntPageLink=pcv_strViewContents&"?idpage="&pcv_lngIDPage
				end if
				'//			
				
				
				if pcv_IntNumrowsCount=1 then strNavigation=strNavigationStart&strNavigation & Vbcrlf
				
				if pcIntHasSubPages = 1 then ' Don't close the list item if there is a submenu
					if pcIntJQNav > 0 then
						strNavigation = strNavigation & "<li><a href=""" & pcStrCntPageLink & """>" & pcv_strPageName & "</a>" & Vbcrlf
					else
						strNavigation = strNavigation & "<li class='" & pcvLIclass & "'><a href=" & pcStrCntPageLink & ">" & pcv_strPageName & "</a>" & Vbcrlf
					end if
				else
					if pcIntJQNav > 0 then
						strNavigation = strNavigation & "<li><a href=""" & pcStrCntPageLink & """>" & pcv_strPageName & "</a></li>" & Vbcrlf
					else
						strNavigation = strNavigation & "<li class='" & pcvLIclass & "'><a href=" & pcStrCntPageLink & ">" & pcv_strPageName & "</a></li>" & Vbcrlf
					end if
				end if	
					
					'// If there are subpages, build menu subsection
					if pcIntHasSubPages = 1 then
					
						query="SELECT pcCont_IDPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage & " ORDER BY pcCont_Order, pcCont_PageName ASC;"
						set rstemp = Server.CreateObject("ADODB.Recordset")
						set rstemp = conntemp.execute(query)
						if err.number <> 0 then
							set rstemp=nothing
							
							call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error loading child pages") 
						end If
						
						pcArraySubPages = rstemp.getRows()
						set rstemp=nothing
						
						pcv_IntNumrowsSP = UBound(pcArraySubPages, 2)
						pcv_IntNumrowsCountSP=0
						
						FOR n = 0 to pcv_IntNumrowsSP
			
							pcv_IntNumrowsCountSP=pcv_IntNumrowsCountSP+1
							pcv_lngIDPageSP= pcArraySubPages(0,n)
							pcv_strPageNameSP = pcArraySubPages(1,n)
							
							'// SEO Links
							'// Build Navigation Product Link
							if scSeoURLs=1 then
								pcStrCntPageLink=pcv_strPageNameSP & "-d" & pcv_lngIDPageSP & ".htm"
								pcStrCntPageLink=removeChars(pcStrCntPageLink)
								pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink
							else
								pcStrCntPageLink=pcv_strViewContents&"?idpage="&pcv_lngIDPageSP
							end if
							'//			
							
							if pcv_IntNumrowsCountSP=1 then strNavigation=strNavigation & strNavigationSubMenu & Vbcrlf
							if pcIntJQNav > 0 then
								strNavigation = strNavigation & "<li><a href=""" & pcStrCntPageLink & """>" & pcv_strPageNameSP & "</a></li>" & Vbcrlf
								else
								strNavigation = strNavigation & "<li class='" & pcvLIsbclass & "'><a href=""" & pcStrCntPageLink & """>" & pcv_strPageNameSP & "</a></li>" & Vbcrlf								
							end if
							'// Close submenu and add closing list item for parent menu item
							if (pcv_IntNumrowsCountSP-1)=pcv_IntNumrowsSP then strNavigation=strNavigation & strNavigationEnd & "</li>"
							
						NEXT
							
					end if ' End if there are subpages
				
				if (pcv_IntNumrowsCount-1)=pcv_IntNumrows then
					 strNavigation=strNavigation & strNavigationEnd
				end if
				
				NEXT
				
				strNavigation=strNavigation & "</div>"
				
				set rstemp = nothing
				
				
				if not validNum(pcIntJQNav) or pcIntJQNav=0 then
				else
					if pcIntJQNav=1 then
						strNavigation=strNavigation & vbcrlf & "<script type=text/javascript>" & vbcrlf
						strNavigation=strNavigation & "ddsmoothmenu.init({"
						strNavigation=strNavigation & "mainmenuid: """ & pcvULID & """," & vbcrlf
						strNavigation=strNavigation & "orientation: 'h'," & vbcrlf
						strNavigation=strNavigation & "classname: 'ddsmoothmenu'," & vbcrlf
						if pcv_menubg<>"" AND pcv_menubga<>"" then
						strNavigation=strNavigation & "customtheme: [""" & pcv_menubg & """, """ & pcv_menubga & """]," & vbcrlf
						end if
						strNavigation=strNavigation & "contentsource: ""markup""" & vbcrlf
						strNavigation=strNavigation & "})" & vbcrlf
						strNavigation=strNavigation & "</script>" & vbcrlf
					else
						strNavigation=strNavigation & vbcrlf & "<script type=text/javascript>" & vbcrlf
						strNavigation=strNavigation & "ddsmoothmenu.init({"
						strNavigation=strNavigation & "mainmenuid: """ & pcvULID & """," & vbcrlf
						strNavigation=strNavigation & "orientation: 'v'," & vbcrlf
						strNavigation=strNavigation & "classname: 'ddsmoothmenu-v'," & vbcrlf
						if pcv_menubg<>"" AND pcv_menubga<>"" then
						strNavigation=strNavigation & "customtheme: [""" & pcv_menubg & """, """ & pcv_menubga & """]," & vbcrlf
						end if
						strNavigation=strNavigation & "arrowswap: true," & vbcrlf
						strNavigation=strNavigation & "contentsource: ""markup""" & vbcrlf
						strNavigation=strNavigation & "})" & vbcrlf
						strNavigation=strNavigation & "</script>" & vbcrlf
					end if
				end if

				
				'-------------------------
				' WRITE to FILE
				'-------------------------
				
				if PPD="1" then
					pcStrFolder="/"&scPcFolder&"/pc"
				else
					pcStrFolder="../pc"
				end if
				
				call pcs_SaveUTF8(pcStrFolder & "\cmsNavigationLinks.inc",pcStrFolder & "\cmsNavigationLinks.inc",strNavigation)
			
				%>
            <tr>
                <td>
                <div class="pcCPmessageSuccess">Content Pages navigation successfully saved to &quot;cmsNavigationLinks.inc&quot;</div>
                The unordered list containing the Content Pages navigation has been saved to the file <strong>cmsNavigationLinks.inc</strong> located in the &quot;<strong>pc</strong>&quot; folder. There are many ways to use an ordered list to create a navigation menu.
                </td>
            </tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
            <tr>
                <td>
                You can find the raw HTML code for the unordered list (UL) that contains the selected Content Pages below:
                <br /><br />
                <textarea cols="80" rows="10"><%=strNavigation %></textarea>
                <br /><br />
                <%
				if pcIntJQNav = 1 then
				%>
                <a href="cmsJQPreview.asp" target="_blank">See JQuery Horizontal Menu Bar Preview</a>
                &nbsp;|&nbsp;
                <%
				elseif pcIntJQNav = 2 then
				%>
                <a href="cmsJQPreview.asp" target="_blank">See JQuery Vertical Menu Bar Preview</a>
                &nbsp;|&nbsp;
                <%
				else
				%>
                <a href="cmsPreview.asp" target="_blank">View HTML</a>
                &nbsp;|&nbsp;
                <%
				end if
				%>
                <a href="cmsNavigation.asp">Back</a>
                </td>
            </tr>
		<%
			END IF
		END IF
		%>					
		<tr>
			<td class="pcCPspacer"></td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
