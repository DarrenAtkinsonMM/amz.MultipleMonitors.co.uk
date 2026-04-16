<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin="11*12*"%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSeoFunctions.asp"-->
<%
dim pcvParentPageName, pcArray, pcInt_Parent, pcInt_Published, pcInt_Inactive, pcInt_Active, pcv_IntNumrowsCount, pcv_IntNumrows, m, pcInt_AlertStoreManager, pcInt_DraftPresent

pcInt_Parent=request("DACatFilter")
if not validNum(pcInt_Parent) then pcInt_Parent = 99
if not pcInt_Parent=99 then
		query1=" WHERE pcCont_BlogCat="&pcInt_Parent
end if

'pcInt_Parent=request("parent")
'if not validNum(pcInt_Parent) then pcInt_Parent = 0
'if pcInt_Parent>0 then
	' Load Parent Page Name
	
	'query="SELECT pcCont_PageName FROM pcContents WHERE pcCont_IDPage="&pcInt_Parent
	'set rs=server.CreateObject("ADODB.RecordSet")
	'set rs=conntemp.execute(query)
	'pcvParentPageName=rs("pcCont_PageName")
	'pageTitle="Manage Content Pages under " & pcvParentPageName
	'query1=" WHERE pcCont_Parent="&pcInt_Parent
	'set rs=nothing
	
'else
	pcvParentPageName=""
	pageTitle="Manage Content Pages"
'end if

%>

<!--#include file="AdminHeader.asp"-->

<%


'// START - Determine the type of user
pcInt_LimitedUser=0
if session("PmAdmin") <> "19" and (not isNull(findUser(pcUserArr,12,pcUserArrCount))) then
	pcInt_LimitedUser=1
end if
'// END - Determine the type of user

IF request("submit1")<>"" OR request("submit2")<>"" THEN

	pcv_IntNumrowsCount=request("IntNumrowsCount")
	pcInt_Parent=request.form("parent")
	if not validNum(pcInt_Parent) then pcInt_Parent = 0
	
	IF validNum(pcv_IntNumrowsCount) then
	
	For k=1 to clng(pcv_IntNumrowsCount)
	
		if request("CT"&k)="1" then
			pcv_id=request("CT"&k&"_id")
			pcInt_Active=request("active"&k)
			pcInt_priority=request.form("priority"&k)
			pcInt_Published=request.form("published"&k)
			
			if not validNum(pcInt_Active) then pcInt_Active=0
			if pcInt_Active="0" then
				pcInt_Inactive="1"
			else
				pcInt_Inactive="0"
			end if
			
			if not validNum(pcInt_Published) then pcInt_Published=0
			
			'// UPDATE Selected Pages
			if request("submit1")<>"" then
				query="UPDATE pcContents SET pcCont_InActive=" & pcInt_Inactive & ", pcCont_Order=" & pcInt_priority & ", pcCont_Published=" & pcInt_Published & " WHERE pcCont_IDPage=" & pcv_id  
				set rstemp=Server.CreateObject("ADODB.Recordset")
				set rstemp=conntemp.execute(query)
				if err.number <> 0 then
					strErrDescription = Err.Description
					set rstemp = nothing
					
					call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "& strErrDescription) 
				else
					msg="Content Pages updated successfully"
				end if
			end if
			
			'// DELETE Selected Pages
			if request("submit2")<>"" then
				query="DELETE FROM pcContents WHERE pcCont_IDPage=" & pcv_id  
				set rstemp=conntemp.execute(query)
				if err.number <> 0 then
					strErrDescription = Err.Description
					set rstemp = nothing
					
					call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error: "& strErrDescription) 
				else
					msg="Content Pages deleted successfully"
				end if
			end if
		end if	
	
	Next
	
	END IF
	
	set rstemp=nothing
	
	call closeDb()
response.redirect "cmsManage.asp?s=1&msg=" & server.URLEncode(msg)
	
END IF

' Load Pages
query="SELECT pcCont_IDPage, pcCont_PageName, pcCont_InActive, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_DraftStatus FROM pcContents" & query1 & " ORDER BY pcCont_Order, pcCont_Parent, pcCont_PageName ASC;"
set rstemp=server.CreateObject("ADODB.RecordSet")
set rstemp=conntemp.execute(query)

if err.number <> 0 then
	set rstemp=nothing
	
	call closeDb()
response.redirect "techErr.asp?error="& Server.Urlencode("Error loading content pages") 
end If

'// START - TO DO items
	if pcInt_LimitedUser=0 then
		' //Check to see if any pages need to be reviewed
		pcInt_AlertStoreManager=0
		query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_Published=0 or pcCont_Published=2;"
		set rstempCheck=server.CreateObject("ADODB.RecordSet")
		set rstempCheck=conntemp.execute(query)
		if not rstempCheck.eof then pcInt_AlertStoreManager=1
		if pcInt_AlertStoreManager=1 then
		%>
			<div class="pcCPmessage">One or more Content Pages need to be <strong>reviewed</strong>.<br><em>It's the pages for which the checkbox in the 'Pub' column below is not checked.</em></div>
		<%
		end if
	
		' Check to see if any pages need to be reviewed
		pcInt_AlertStoreManager=0
		query="SELECT pcCont_IDPage FROM pcContents WHERE pcCont_DraftStatus=1;"
		set rstempCheck=conntemp.execute(query)
		if not rstempCheck.eof then pcInt_AlertStoreManager=1
		set rstempCheck=nothing
		if pcInt_AlertStoreManager=1 then
		%>
			<div class="pcCPmessage">One or more Content Pages have a <strong>draft</strong> saved to the database, which might need to be completed, reviewed, and published.<br><em>It's the pages for which the column 'Draft' says 'Yes' below.</em></div>
		<%
		end if
	end if
'// END - TO DO items
%>


<%
'Check URL

	SPath1=Request.ServerVariables("PATH_INFO")
	mycount1=0
	do while mycount1<2
	if mid(SPath1,len(SPath1),1)="/" then
	mycount1=mycount1+1
	end if
	if mycount1<2 then
	SPath1=mid(SPath1,1,len(SPath1)-1)
	end if
	loop
	SPathInfo="http://" & Request.ServerVariables("HTTP_HOST") & SPath1

%>

<form name="form1" action="cmsManage.asp" method="post" class="pcForms">
                <% ' START show message, if any %>
                    <!--#include file="pcv4_showMessage.asp"-->
                <% 	' END show message %>

  <div class="pcCPsortableTableHeader">
    <div class="pcCPsortableTableIndex">#</div>
		<div class="pcCPcontentCheck">&nbsp;</div>
		<div class="pcCPcontentActive">Active</div>
 		<div class="pcCPcontentPub">Pub</div>
    <div class="pcCPcontentDraft">Draft</div>
		<div class="pcCPcontentName">Name</div>
    <div class="pcCPcontentParent">
      Parent
    </div>
    <div class="pcCPcontentActions">
			<%
            ' Load parent pages - Start
			'if pcInt_Parent>0 then
			%>
            <span class="cpLinksList"><a href="cmsManage.asp">Show All</a></span>
            <%
			'else
        Dim pcPageParentExist, intPageCount
        query="SELECT pcCont_idPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 ORDER BY pcCont_PageName ASC"
        set rs=Server.CreateObject("ADODB.Recordset")
        set rs=connTemp.execute(query)
        if rs.EOF then
            pcPageParentExist=0
        else
            pcPageParentExist=1
            pcPageArr=rs.getRows()
        end if
        set rs=nothing
        'if pcPageParentExist=1 then
        %>
          <select name="DACatFilter" tabindex="104" onChange="this.form.submit()">
                                <option value="99">Show All Pages</option>
                                <option value="99">--- Blog ---</option>
                                <option value="0">Hardware</option>
                                <option value="1">Product Guides</option>
                                <option value="2">Setup Guides</option>
                                <option value="3">Software</option>
                                <option value="4">Stands</option>
                                <option value="5">Trading</option>
                                <option value="99">--- Pages ---</option>
                                <option value="6">Hardware</option>
                                <option value="7">Software</option>
                                <option value="99">--- Main Site ---</option>
                                <option value="9">General Pages</option>
                                <option value="8">Landing Pages</option>
                        
                            </select>
      <%
     ' end if
    'end if
    ' Load parent pages - End
    %>
    </div>
  </div>

		<%
		'-------------------------
		' NO Content Pages Found
		'-------------------------
		If rstemp.eof Then
			set rstemp=nothing
			
		%>
	    <div class="pcCPmessage">No Content Pages Found. <a href="cmsAddEdit.asp">Add New</a>&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=436"></a></div>          
		<% 
		'-------------------------
		' LIST Content Pages
		'-------------------------
		Else 
		  %><ul class="pcCPsortable pcCPsortableTable"><%
			pcArray = rstemp.getRows()
			set rstemp=nothing
			
			
			pcv_IntNumrows = UBound(pcArray, 2)

			pcv_IntNumrowsCount=0
			FOR m = 0 to pcv_IntNumrows

				pcv_IntNumrowsCount=pcv_IntNumrowsCount+1
				pcv_lngIDPage= pcArray(0,m)
				pcv_strPageName = pcArray(1,m)
				pcInt_Inactive = pcArray(2,m)
				pcInt_priority = pcArray(3,m)
				pcInt_Parent = pcArray(4,m)
				pcInt_Published = pcArray(5,m)
				pcInt_DraftPresent = pcArray(6,m)
				
				'// SEO Links
				'// Build Navigation Page Link
				if scSeoURLs=1 then
					pcStrCntPageLink=pcv_strPageName & "-d" & pcv_lngIDPage & ".htm"
					pcStrCntPageLink=removeChars(pcStrCntPageLink)
					pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink &"?adminPreview=1"
				else
					pcStrCntPageLink="../pc/viewcontent.asp?idpage="&pcv_lngIDPage&"&adminPreview=1"
				end if
				
				'// Change links if this is a parent page
				Dim intAlreadyParent
				
				query="SELECT pcCont_idPage FROM pcContents WHERE pcCont_Parent=" & pcv_lngIDPage
				set rsParent=Server.CreateObject("ADODB.Recordset")
				set rsParent=connTemp.execute(query)
				if rsParent.EOF then
					intAlreadyParent=0
				else
					intAlreadyParent=1
				end if
				if intAlreadyParent=1 then 
					if scSeoURLs=1 then
						pcStrCntPageLink=pcv_strPageName & "-e" & pcv_lngIDPage & ".htm"
						pcStrCntPageLink=removeChars(pcStrCntPageLink)
						pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink &"?adminPreview=1"
					else
						pcStrCntPageLink="../pc/viewcontent.asp?idpage="&pcv_lngIDPage&"&adminPreview=1"
					end if
				end if
				set rsParent=nothing
				
				'//
				
				if not validNum(pcInt_Inactive) then
					pcInt_Inactive=0
				end if
				
				if not validNum(pcInt_Parent) then
					pcInt_Parent=0
				end if
				
				if not validNum(pcInt_Published) then
					pcInt_Published=0
				end if
				
				if not validNum(pcInt_DraftPresent) then
					pcInt_DraftPresent=0
				end if
				
			%>           
				<li class="cpItemlist"> 
          <div class="pcCPsortableTableIndex">
            <span class="pcCPsortableIndex"><%= pcv_IntNumrowsCount %></span>
						<input type="hidden" class="pcCPsortableOrder" name="priority<%=pcv_IntNumrowsCount%>" value="<%=pcInt_priority%>">
          </div>
					<div class="pcCPcontentCheck">
						<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
                            <input name="ct<%=pcv_IntNumrowsCount%>" type="checkbox" value="1" class="clearBorder">
                        <% end if %>
						<input type="hidden" name="ct<%=pcv_IntNumrowsCount%>_id" value="<%=pcv_lngIDPage%>">
					</div>
					<div class="pcCPcontentActive">
						<input type="checkbox" name="active<%=pcv_IntNumrowsCount%>" value="1" <%if pcInt_Inactive="0" then%>checked<%end if%><%if pcInt_LimitedUser=1 then%> disabled <%end if%> class="clearBorder">
					</div>
					<div class="pcCPcontentPub">
						<input type="checkbox" name="published<%=pcv_IntNumrowsCount%>" value="1" <%if pcInt_Published="1" then%>checked<%end if%><%if pcInt_LimitedUser=1 then%> disabled <%end if%> class="clearBorder">
					</div>
					<div class="pcCPcontentDraft">
						<% if pcInt_DraftPresent<>0 then %>Yes<% else %>No<% end if %>
					</div>
					<div class="pcCPcontentName">
						<a href="cmsAddEdit.asp?idpage=<%=pcv_lngIDPage%>"><%=pcv_strPageName%></a>
					</div>
          <div class="pcCPcontentParent">
            <% 
					  if pcInt_Parent > 0 then
						  
						  query="SELECT pcCont_PageName FROM pcContents WHERE pcCont_IDPage=" & pcInt_Parent
						  set rstemp = Server.CreateObject("ADODB.Recordset")
						  set rstemp = conntemp.execute(query)
						  if not rstemp.eof then
						  pcv_ParentPageName = rstemp("pcCont_PageName")
						  else
						  pcv_ParentPageName = "N/A"
						  end if
						  set rstemp = nothing
						  
					  %>
              <a href="cmsAddEdit.asp?idpage=<%=pcInt_Parent%>"><%=pcv_ParentPageName%></a>
            <%
					end if
					%>
                    </div>
					<div class="pcCPcontentActions cpLinksList">
            <a href="cmsAddEdit.asp?idpage=<%=pcv_lngIDPage%>"><img src="images/pcIconGo.jpg" border="0" alt="Edit" title="Edit this Content Page"></a>&nbsp;<a href="cmsAddEdit.asp?idpage=<%=pcv_lngIDPage%>">Edit</a> |
            <a href="<%=pcStrCntPageLink%>" target="_blank"><img src="images/pcIconPreview.jpg" border="0" alt="Preview" title="Preview this Content Page"></a>&nbsp;<a href="<%=pcStrCntPageLink%>" target="_blank">Preview</a>
					</div>
				</li>
				<%
				NEXT
				%>
        </ul>

				<table class="pcCPcontent">
                <% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<tr>
					<td colspan="8">
                    	<% if pcv_IntNumrowsCount<>"" then %>
						<script type=text/javascript>
              function checkAll() {
              for (var j = 1; j <= <%=pcv_IntNumrowsCount%>; j++) {
              box = eval("document.form1.ct" + j); 
              if (box.checked == false) box.checked = true;
                    }
              }
                            
              function uncheckAll() {
              for (var j = 1; j <= <%=pcv_IntNumrowsCount%>; j++) {
              box = eval("document.form1.ct" + j); 
              if (box.checked == true) box.checked = false;
                    }
              }
          </script>
						<%end if%>
						<span class="cpLinksList"><a href="javascript:checkAll();">Check All</a> | <a href="javascript:uncheckAll();">Uncheck All</a></span>
					</td>
				</tr>	
		<%
				end if
			END IF
		'-------------------------
		' END listing content pages
		'-------------------------
		%>					

		<tr>
			<td colspan="8" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td align="center" colspan="8">
            	<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<input name="submit1" type="submit" value="Update Selected" class="btn btn-primary">&nbsp;
				<input name="submit2" class="btn btn-default" type="submit" value="Delete Selected" onclick="return(confirm('You are about to remove selected content pages from your database. Are you sure you want to complete this action?'));">&nbsp;
               	<input type="hidden" name="IntNumrowsCount" value="<%=pcv_IntNumrowsCount%>">
                <% end if %>
				<input type="button" class="btn btn-default"  value="Add New" onclick="location='cmsAddEdit.asp';">&nbsp;
				<% if session("PmAdmin") = "19" or (isNull(findUser(pcUserArr,12,pcUserArrCount))) then %>
				<input type="button" class="btn btn-default"  value="Generate Navigation" onclick="location='cmsNavigation.asp';">&nbsp;
                <% end if %>
                <input type="button" class="btn btn-default"  value="Browse Pages" onclick="window.open('../pc/viewcontent.asp');">&nbsp;
                <input type="button" class="btn btn-default"  value="Help" onclick="window.open('http://wiki.productcart.com/productcart/settings-content-pages');">&nbsp;
				<input type="button" class="btn btn-default"  value="Back" onClick="javascript:history.back()">
			</td>
		</tr>
	</table>
</form>
<!--#include file="AdminFooter.asp"-->
