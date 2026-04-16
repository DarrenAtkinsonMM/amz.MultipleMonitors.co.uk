<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin="11*12*"%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcSeoFunctions.asp"-->
<%
dim pcInt_idpage, pcIntIncHeader, pcIntNewPage, pcIntHideBackButton, queryDraft, queryDraft2, pcv_PageDraft, pcInt_PageDraftPublish, pcInt_LimitedUser

'// START - Determine the type of user
pcUserArr = split(session("PmAdmin"),"*")
pcUserArrCount = ubound(pcUserArr)-1
pcInt_LimitedUser=0
if session("PmAdmin") <> "19" and (not isNull(findUser(pcUserArr,12,pcUserArrCount))) then
	pcInt_LimitedUser=1
end if
'// END - Determine the type of user

IF request("action")="add" THEN

	pcv_PageName=left(request("PageName"),250)
	pcv_PageName=pcf_ReplaceCharacters(pcv_PageName)
	pcv_PageTitle=left(request("PageTitle"),250)
	pcv_PageTitle=pcf_ReplaceCharacters(pcv_PageTitle)
	
	'// Main page content and draft
	pcv_PageDesc=request("details")
	pcv_PageDesc=pcf_SaveHTMLField(pcv_PageDesc)     
	pcv_PageDraft=request("pageDraft")
	pcv_PageDraft=pcf_ReplaceCharacters(pcv_PageDraft)
	if pcInt_LimitedUser=1 then
		if trim(pcv_PageDraft)="" then pcv_PageDraft=pcv_PageDesc
	end if
	pcInt_PageDraftPublish=request("draftPublish")
	
	pcv_MetaDesc=left(request("metadesc"),500)
	pcv_MetaDesc=pcf_ReplaceCharacters(pcv_MetaDesc)
	pcv_MetaKeywords=left(request("metakeywords"),500)
	pcv_MetaKeywords=pcf_ReplaceCharacters(pcv_MetaKeywords)
	pcv_MetaTitle=left(request("metatitle"),250)
	pcv_MetaTitle=pcf_ReplaceCharacters(pcv_MetaTitle)
	
	pcv_Active=request("Active")
	pcInt_Published=request("Published")
	pcv_PageComments=request("PageComments")
	pcv_PageComments=pcf_ReplaceCharacters(pcv_PageComments)
	pcInt_Parent=request("Parent")
	pcInt_Order=request("Order")
	pcv_PageThumbnail=request("PageThumbnail")
	pcv_PageThumbnail=pcf_ReplaceCharacters(pcv_PageThumbnail)
	pcIntIncHeader=request("IncHeader")
	pcIntMenuExclude=request("MenuExclude")
	pcv_CustomerType=request("CustomerType")
	pcIntHideBackButton=request("HideBackButton")
	pcvUrl=request("pcv_PUrl")
	pcv_Blog=request("blogPage")
	pcv_BlogVis=request("blogVisibility")
	pcv_BlogCat=request("blogCategory")
	pcv_BlogPubDate=request("blogPubDate")
	PCV_BlogIntro=request("blogIntro")

	if (pcv_Active="") or (pcv_Active="0") then
		pcv_InActive="1"
		else
		pcv_InActive="0"
	end if

	if not validNum(pcInt_Parent) then pcInt_Parent=0
	if pcIntIncHeader="" then pcIntIncHeader=0
	if not validNum(pcIntIncHeader) then pcIntIncHeader=1
	if not validNum(pcInt_Order) then pcInt_Order=0
	if not validNum(pcInt_Published) then pcInt_Published=0
	if not validNum(pcIntMenuExclude) then pcIntMenuExclude=0
	if not validNum(pcIntHideBackButton) then pcIntHideBackButton=0
	if not validNum(pcInt_PageDraftPublish) then pcInt_PageDraftPublish=0
	
	pcInt_idpage=request("idpage")
	
	if pcInt_idpage<>"" then

		'// EDIT existing content page
		if not validNum(pcInt_idpage) then
			call closeDb()
response.redirect "cmsManage.asp?msg=" & server.URLEncode("Not a valid content page ID")
		end if
		
		
		'// A user with limited permissions (Add/Edit) modified the page -> keep current, save draft
		if pcInt_LimitedUser=1 then
			queryDraft = "" ' The current page content is not edited
			queryDraft2 = ",pcCont_Draft=N'" & pcv_PageDraft & "',pcCont_DraftStatus=1"
			pcStrMessage="A draft of your edits has been saved to the database. The live content has not changed. It will be changed when your edits are approved and published."
		'// A user with publishing permissions is editing the page -> if there is a draft, will it be published?
		else
			if pcInt_PageDraftPublish=1 then ' Publish the draft
				queryDraft = "',pcCont_Description=N'" & pcv_PageDraft ' The draft becomes the live content
				queryDraft2 = ",pcCont_Draft='',pcCont_DraftStatus=0" ' Current draft is removed since it's now the live content
				pcStrMessage="Content Page edited successfully! The draft copy of the page has become the live content."
				
			elseif pcInt_PageDraftPublish=2 then ' User with publishing rights edited the draft, but did not publish it
				queryDraft = "',pcCont_Description=N'" & pcv_PageDesc ' Changes might have been made to the live content
				queryDraft2 = ",pcCont_Draft=N'" & pcv_PageDraft & "',pcCont_DraftStatus=1" ' Save changes to the draft
				pcStrMessage="Content Page edited successfully! <br /><br />A draft of your edits has been saved to the database. The live content has not changed. Remember to publish the draft when you are ready to do so."
				
			elseif pcInt_PageDraftPublish=4 then ' User with publishing rights wants to save a draft
				queryDraft = "" ' The current page content is not edited
				queryDraft2 = ",pcCont_Draft=N'" & pcv_PageDesc & "',pcCont_DraftStatus=1" ' It is instead saved as a draft
				pcStrMessage="Content Page edited successfully! <br /><br />A draft of your edits has been saved to the database. The live content has not changed. Remember to publish the draft when you are ready to do so."
				
			else ' There is no draft
				queryDraft = "',pcCont_Description=N'" & pcv_PageDesc
				queryDraft2 = ",pcCont_Draft='',pcCont_DraftStatus=0"
				pcStrMessage="Content Page edited successfully!"
			end if
		end if

		query="UPDATE pcContents SET pcCont_PageName=N'" & pcv_PageName & queryDraft & "',pcCont_IncHeader=" & pcIntIncHeader & ",pcCont_InActive=" & pcv_InActive & ",pcCont_MetaTitle=N'" & pcv_MetaTitle & "',pcCont_MetaDesc=N'" & pcv_MetaDesc & "',pcCont_MetaKeywords=N'" & pcv_MetaKeywords & "', pcCont_Order=" & pcInt_Order & ", pcCont_Parent=" & pcInt_Parent & ", pcCont_Published=" & pcInt_Published & ",pcCont_Thumbnail='" & pcv_PageThumbnail & "', pcCont_Comments=N'" & pcv_PageComments & "', pcCont_PageTitle=N'" & pcv_PageTitle & "', pcCont_MenuExclude=" & pcIntMenuExclude & ", pcCont_CustomerType='" & pcv_CustomerType & "', pcCont_HideBackButton =" & pcIntHideBackButton & queryDraft2 & ", pcUrl ='" & pcvUrl & "', pcCont_Blog=" & pcv_Blog & ", pcCont_BlogVis=" & pcv_BlogVis & ", pcCont_BlogCat=" & pcv_BlogCat & ", pcCont_PubDate='" & pcv_BlogPubDate & "', pcCont_BlogIntro='" & pcv_BlogIntro & "' WHERE pcCont_IDPage=" & pcInt_idpage & ";"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		
		
		call closeDb()
		response.redirect "cmsAddEdit.asp?s=1&idpage="& pcInt_idpage & "&idparent=" & pcInt_Parent & "&msg=" & server.URLEncode(pcStrMessage)
		
	else

		'// ADD new content page
		
		query="INSERT INTO pcContents (pcCont_PageName, pcCont_Description, pcCont_IncHeader, pcCont_InActive, pcCont_MetaTitle, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_Order, pcCont_Parent, pcCont_Published, pcCont_Thumbnail, pcCont_Comments, pcCont_PageTitle, pcCont_MenuExclude, pcCont_CustomerType, pcCont_HideBackButton,pcCont_Draft,pcCont_DraftStatus) values (N'" & pcv_PageName & "',N'" & pcv_PageDesc & "'," & pcIntIncHeader & "," & pcv_InActive & ",N'" & pcv_MetaTitle & "',N'" & pcv_MetaDesc & "',N'" & pcv_MetaKeywords & "'," & pcInt_Order & "," & pcInt_Parent & "," & pcInt_Published &",'" & pcv_PageThumbnail & "',N'" & pcv_PageComments & "',N'" & pcv_PageTitle & "'," & pcIntMenuExclude & ",'" & pcv_CustomerType & "'," & pcIntHideBackButton & ",'',0);"
		set rstemp=Server.CreateObject("ADODB.Recordset")
		set rstemp=connTemp.execute(query)
		set rstemp=nothing
		
		call closeDb()
		response.redirect "cmsManage.asp?s=1&msg=" & server.URLEncode("Content Page added successfully!")
	end if

END IF 
		
IF request("idpage")<>"" THEN

	pcIntNewPage=0
	pcInt_idpage=request("idpage")
	if not validNum(pcInt_idpage) then
		call closeDb()
response.redirect "cmsManage.asp?msg=" & server.URLEncode("Not a valid content page ID")
	end if
	
	
	query="SELECT pcCont_PageName, pcCont_PageTitle, pcCont_Comments, pcCont_Thumbnail, pcCont_IncHeader, pcCont_InActive, pcCont_Published, pcCont_Order, pcCont_Parent, pcCont_MetaTitle, pcCont_MetaDesc, pcCont_MetaKeywords, pcCont_MenuExclude, pcCont_CustomerType, pcCont_HideBackButton, pcCont_Description, pcCont_Draft, pcUrl, pcCont_Blog,pcCont_BlogVis,pcCont_BlogCat,pcCont_PubDate,pcCont_BlogIntro FROM pcContents WHERE pcCont_IDPage=" & pcInt_idpage
	set rstemp=server.CreateObject("ADODB.Recordset")
	set rstemp=connTemp.execute(query)
	
	pcv_PageName=rstemp("pcCont_PageName")
	pcv_PageName=replace(pcv_PageName,"""","&quot;")
	
	pcv_PUrl=rstemp("pcUrl")
	
	pcv_PageTitle=rstemp("pcCont_PageTitle")
	if pcv_PageTitle<>"" then
		pcv_PageTitle=replace(pcv_PageTitle,"""","&quot;")
	end if
	
	pcv_PageComments=rstemp("pcCont_Comments")
	if pcv_PageComments<>"" and not isNull(pcv_PageComments) then
		pcv_PageComments=replace(pcv_PageComments,"<","&lt;")
		pcv_PageComments=replace(pcv_PageComments,">","&gt;")
		pcv_PageComments=replace(pcv_PageComments,"""","&quot;")
	end if
	
	pcv_PageThumbnail=rstemp("pcCont_Thumbnail")

	pcIntIncHeader=rstemp("pcCont_IncHeader")
	if not validNum(pcIntIncHeader) then pcIntIncHeader=1
	
	pcv_InActive=rstemp("pcCont_InActive")
	if not validNum(pcv_InActive) then pcv_InActive=0

	if pcv_InActive="0" then
		pcv_Active="1"
		else
		pcv_Active="0"
	end if
	
	pcInt_Published=rstemp("pcCont_Published")
	if not validNum(pcInt_Published) then pcInt_Published=0
	
	pcInt_Order=rstemp("pcCont_Order")
	if not validNum(pcInt_Order) then pcInt_Order=0
	
	pcInt_Parent=rstemp("pcCont_Parent")
	if not validNum(pcInt_Parent) then pcInt_Parent=0

	pcv_MetaTitle=rstemp("pcCont_MetaTitle")
	if pcv_MetaTitle<>"" and not isNull(pcv_MetaTitle) then
		pcv_MetaTitle=replace(pcv_MetaTitle,"""","&quot;")
	end if
	
	pcv_MetaDesc=rstemp("pcCont_MetaDesc")
	if pcv_MetaDesc<>"" and not isNull(pcv_MetaDesc) then
		pcv_MetaDesc=replace(pcv_MetaDesc,"""","&quot;")
	end if
	
	pcv_MetaKeywords=rstemp("pcCont_MetaKeywords")
	if pcv_MetaKeywords<>"" and not isNull(pcv_MetaKeywords) then		
		pcv_MetaKeywords=replace(pcv_MetaKeywords,"""","&quot;")
	end if
	
	pcIntMenuExclude=rstemp("pcCont_MenuExclude")
	if not validNum(pcIntMenuExclude) then pcIntMenuExclude=0
	
	pcv_CustomerType=rstemp("pcCont_CustomerType")
	
	pcIntHideBackButton=rstemp("pcCont_HideBackButton")
	if not validNum(pcIntHideBackButton) then pcIntHideBackButton=0
	
	pcv_PageDesc=rstemp("pcCont_Description")
	if pcv_PageDesc<>"" and not isNull(pcv_PageDesc) then
		pcv_PageDesc=replace(pcv_PageDesc,"<","&lt;")
		pcv_PageDesc=replace(pcv_PageDesc,">","&gt;")
		pcv_PageDesc=replace(pcv_PageDesc,"""","&quot;")
	end if
	
	pcv_PageDraft=rstemp("pcCont_Draft")
	if pcv_PageDraft<>"" and not isNull(pcv_PageDesc) then
		pcv_PageDesc=replace(pcv_PageDesc,"<","&lt;")
		pcv_PageDesc=replace(pcv_PageDesc,">","&gt;")
		pcv_PageDesc=replace(pcv_PageDesc,"""","&quot;")
	end if
	
	pcCont_Blog=rstemp("pcCont_Blog")
	
	if pcCont_Blog = 0 then
		pcvPageSelect = " selected=""selected"""
		pcvBlogSelect = ""
	Else
		pcvBlogSelect = " selected=""selected"""
		pcvPageSelect = ""
	end if
	
	if rstemp("pcCont_BlogVis") = 0 then
		pcvBlogVisPublic = " selected=""selected"""
		pcvBlogVisPrivateAll = ""
		pcvBlogVisPrivatePC = ""
	elseif rstemp("pcCont_BlogVis") = 1 then
		pcvBlogVisPublic = ""
		pcvBlogVisPrivateAll = " selected=""selected"""
		pcvBlogVisPrivatePC = ""
	else
		pcvBlogVisPublic = ""
		pcvBlogVisPrivateAll = ""
		pcvBlogVisPrivatePC = " selected=""selected"""
	end if
	
	pcvBlogCatHW = ""
	pcvBlogCatPG = ""
	pcvBlogCatSG = ""
	pcvBlogCatSW = ""
	pcvBlogCatST = ""
	pcvBlogCatTR = ""
	pcvBlogCatPHW = ""
	pcvBlogCatPSW = ""
	pcvBlogCatSLP = ""
	pcvBlogCatSGP = ""
	
	select case rstemp("pcCont_BlogCat")
		case 0
			pcvBlogCatHW = " selected=""selected"""
		case 1
			pcvBlogCatPG = " selected=""selected"""
		case 2
			pcvBlogCatSG = " selected=""selected"""
		case 3
			pcvBlogCatSW = " selected=""selected"""
		case 4
			pcvBlogCatST = " selected=""selected"""
		case 5
			pcvBlogCatTR = " selected=""selected"""
		case 6
			pcvBlogCatPHW = " selected=""selected"""
		case 7
			pcvBlogCatPSW = " selected=""selected"""
		case 8
			pcvBlogCatSLP = " selected=""selected"""
		case 9
			pcvBlogCatSGP = " selected=""selected"""
	end select
	
	pcvBlogPubDate = rstemp("pcCont_PubDate")
	
	pcv_BlogIntro = rstemp("pcCont_BlogIntro")
	
	set rstemp = nothing
	
	
ELSE

	pcIntNewPage=1

END IF

'// Create Page Title
if pcIntNewPage=0 then
	pageTitle="Edit Content Page: <strong>" & pcv_PageName & "</strong>"
	else
	pageTitle="Add New Content Page"
end if
%>
<!--#include file="AdminHeader.asp"-->

<!-- #include file="../htmleditor/editor.asp" -->

<script type=text/javascript>
function Form1_Validator(theForm) {
	// InnovaStudio HTML Editor Workaround for this keyword
	theForm = document.form1;

	if (theForm.PageName.value == "") {
		alert("Please enter a page name.");
		theForm.PageName.focus();
		return (false);
	}

	// Workaround for inline HTML editor
	var editor1 = $pc(theForm).find("#idContenteditor1").contents().find("body");

	if (editor1.html() == "") {
		alert("Please enter the page content.");
		editor1.focus();
		return (false);
	}

	return (true);
}

function newWindow(file,window)
{
	msgWindow=open(file,window,'resizable=no,width=400,height=500');
	if (msgWindow.opener == null) msgWindow.opener = self;
}
</script>

<%
'// Show alert message to user with limited permissions
if request.QueryString("msg")="" and pcInt_LimitedUser=1 then
%>
	<div class="pcCPmessage">When you <strong>add a new Content Page</strong>, the page will be offline and <em>Under Review</em>.<br />When you <strong>edit an existing page</strong>, the changes will be saved to a draft (<em>live page is not edited</em>). <br /><br /> When you are done with your edits, <a href="mailto:<%=scFrmEmail%>">e-mail the store manager</a> (<em>an e-mail is not sent automatically to avoid &quot;e-mail overload&quot;</em>) so that they can be reviewed and published.</div>
<%
end if
%>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>

<form name="form1" action="cmsAddEdit.asp?action=add" method="post" onSubmit="return Form1_Validator(this)" class="pcForms">
    
		<%
		'// TABBED PANELS - MAIN DIV START
		%>
        <div id="TabbedPanels1" class="tabbable-left">
		
		<%
		'// TABBED PANELS - START NAVIGATION
		%>
		<div class="col-xs-3">
            <ul class="nav nav-tabs tabs-left">
				<li class="active"><a href="#tabs-1" data-toggle="tab">Name, Title, &amp; Content</a></li>
				<li><a href="#tabs-2" data-toggle="tab">Settings</a></li>
				<li><a href="#tabs-3" data-toggle="tab">Meta Tags</a></li>
				<li>
                	<div style="margin-top:10px; margin-bottom:10px; text-align: center">
					<input type="submit" name="submit" value="<%if request("idpage")<>"" then%>Update Page<%else%>Add Content page<%end if%>" class="btn btn-primary">
                    
                    <% if validNum(pcInt_idpage) then %>
                        <div style="margin-top: 5px">
                        <%

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
							
							if Right(SPathInfo,1)="/" then
							pcv_strViewContents=SPathInfo & "pc/viewcontent.asp"
							else
							pcv_strViewContents=SPathInfo & "/pc/viewcontent.asp"
							end if
							
                            '// SEO Links
                            '// Build Navigation Product Link
                            if scSeoURLs=1 then
                                pcStrCntPageLink=pcv_PageName & "-d" & pcInt_idpage & ".htm"
                                pcStrCntPageLink=removeChars(pcStrCntPageLink)
                                pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink &"?adminPreview=1"
                            else
                                pcStrCntPageLink=pcv_strViewContents&"?idpage="&pcInt_idpage&"&adminPreview=1"
                            end if
                            
							'// Change links if this is a parent page
							Dim intAlreadyParent
							
							query="SELECT pcCont_idPage FROM pcContents WHERE pcCont_Parent=" & pcInt_idpage
							set rsParent=Server.CreateObject("ADODB.Recordset")
							set rsParent=connTemp.execute(query)
							if rsParent.EOF then
								intAlreadyParent=0
							else
								intAlreadyParent=1
							end if
							if intAlreadyParent=1 then
								if scSeoURLs=1 then
									pcStrCntPageLink=pcv_PageName & "-e" & pcInt_idpage & ".htm"
									pcStrCntPageLink=removeChars(pcStrCntPageLink)
									pcStrCntPageLink=SPathInfo & "pc/" & pcStrCntPageLink &"?adminPreview=1"
								else
									pcStrCntPageLink="../pc/viewcontent.asp?idpage="&pcInt_idpage&"&adminPreview=1"
								end if
							end if
							set rsParent=nothing
                        %>
                        <input type="button" class="btn btn-default"  name="Button" value="Preview" onClick="window.open('<%=pcStrCntPageLink%>');">
                        </div>
                    <% end if %>

                    <div style="margin-top: 5px">
					<input type="hidden" name="idpage" value="<%=request("idpage")%>">
					<input type="button" class="btn btn-default"  name="Button" value="Manage Pages" onClick="location='cmsManage.asp';">
                    </div>
                    </div>
				</li>
			</ul>
		</div>	
		<%
		'// TABBED PANELS - END NAVIGATION
		
		'// TABBED PANELS - START PANELS
		%>
        <div class="col-xs-9">
            <div class="tab-content">
		
			<%
			'// =========================================
			'// FIRST PANEL - START - Name, Descriptions, Images
			'// =========================================
			%>
				<div id="tabs-1" class="tab-pane active">
				
					<table class="pcCPcontent">

                        <tr>
                        	<td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<th colspan="2">Name, Title, and Content</th>
                        </tr>
                        <tr>
                        	<td colspan="2" class="pcCPspacer"></td>
                        </tr>
                    	<tr>
                        	<td colspan="2">Use this feature to create &quot;content&quot; pages that you want to manage from your Control Panel. <a href="http://wiki.productcart.com/productcart/settings-content-pages" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this feature" border="0"></a></td>
                        </tr>
                        <tr>							
                            <td valign="top" align="right" nowrap>Page Name (Link):</td>
                            <td>
                                <input type="text" name="PageName" size="50" maxlength="255" value="<%=pcv_PageName%>">
                                <div class="pcSmallText" style="padding-top: 6px; padding-bottom: 10px;"><i>Max 255 characters. Shown in the CP and in the storefront when linking to the page.</i></div>
                            </td>
                        </tr>
                        <tr>							
                            <td valign="top" align="right" nowrap>Page Url:</td>
                            <td>
                                <input type="text" name="pcv_PUrl" size="50" maxlength="255" value="<%=pcv_PUrl%>">
                                <div class="pcSmallText" style="padding-top: 6px; padding-bottom: 10px;"><i>Max 255 characters. Optional. Displayed using an H1 HTML tag on the page itself.</i></div>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="right">
                            Page Content:
                            </td>
                            <td valign="top">  
                                <textarea id="details" name="details" rows="14" cols="60"><%=pcv_PageDesc%></textarea>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                         <tr>
                            <td valign="top" align="right">
                            Blog Intro:
                            </td>
                            <td valign="top">  
                                <textarea name="blogIntro" rows="5" cols="60"><%=pcv_BlogIntro%></textarea>
                                 <div class="pcSmallText" style="padding-top: 6px; padding-bottom: 10px;"><i>Include the 'p' html tag. Also used as page meta desc with p tag stripped out.</i></div>
                           </td>
                        </tr>
                        <% 
						'// START - Draft content
						if pcv_PageDraft<>"" then
						%>
                        <tr>
                            <td valign="top" align="right">
                            Draft Content:
                            </td>
                            <td valign="top">  
                                <textarea class="htmleditor" id="pageDraft" name="pageDraft" rows="14" cols="60"><%=pcv_PageDraft%></textarea>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <% 
							if pcInt_LimitedUser=0 then
						%>
                            <tr>							
                                <td valign="top" align="right" nowrap>Draft Content options:</td>
                                <td>
                                    <input type="radio" value="1" name="draftPublish" class="clearBorder"> Publish draft (<em>draft becomes live page</em>) <br />
                                    <input type="radio" value="2" name="draftPublish" class="clearBorder" checked> Save draft (<em>both live page changes and changes to the draft are saved</em>)<br />
                                    <input type="radio" value="3" name="draftPublish" class="clearBorder"> Remove draft (<em>deletes the draft</em>)<br />
                                </td>
                            </tr>
                        <%
							end if 
						else 
							if pcInt_LimitedUser=0 then
						%>
                            <tr>							
                                <td valign="top" align="right">
                                    <input type="checkbox" value="4" name="draftPublish" class="clearBorder">
                                </td>
                                <td>
                                	Save <em>Page Content</em> as a <strong>draft</strong><br /><em>Page content that is currently live is not modified. Modifications are saved to a draft</em>.
                                </td>
                            </tr>
                        <% 
							end if
						end if 
						'// END - Draft content
						%>
					</table>
					
                    <table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
                        <tr>
                            <th colspan="2">Settings</th>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td>Active:</td>
                            <td>
                                <input type="checkbox" name="Active" value="1" <%if pcv_Active="1" then%>checked<%end if%> class="clearBorder"> Yes
                            </td>
                        </tr>
                         <tr>
                        	<td nowrap>Template / URL:</td>
                            <td>
                                <select name="blogPage">
                                    <option value="0"<%=pcvPageSelect%>>Page</option>
                                    <option value="1"<%=pcvBlogSelect%>>Blog</option>
                                </select>
                            </td>
                        </tr>
                        <tr>  
                        	<td>Remove Title:</td>
                            <td><input type="checkbox" name="HideBackButton" value="1" <%if pcIntHideBackButton="1" then%>checked<%end if%> class="clearBorder"> Setting only affects pages template, used on some landing pages.</td>
                        </tr>
                         <tr>
                        	<td nowrap>Blog Visibility:</td>
                            <td>
                                <select name="blogVisibility">
                                    <option value="0"<%=pcvBlogVisPublic%>>Public</option>
                                    <option value="1"<%=pcvBlogVisPrivateAll%>>Private All</option>
                                    <option value="2"<%=pcvBlogVisPrivatePC%>>Private PC / Bundle Only</option>
                                </select>
                            </td>
                        </tr>
                         <tr>
                        	<td nowrap>Category:</td>
                            <td>
                                <select name="blogCategory">
                                    <option value="0"<%=pcvBlogCatHW%>>Blog | Hardware</option>
                                    <option value="1"<%=pcvBlogCatPG%>>Blog | Product Guides</option>
                                    <option value="2"<%=pcvBlogCatSG%>>Blog | Setup Guides</option>
                                    <option value="3"<%=pcvBlogCatSW%>>Blog | Software</option>
                                    <option value="4"<%=pcvBlogCatST%>>Blog | Stands</option>
                                    <option value="5"<%=pcvBlogCatTR%>>Blog | Trading</option>
                                    <option value="6"<%=pcvBlogCatPHW%>>Page | Hardware</option>
                                    <option value="7"<%=pcvBlogCatPSW%>>Page | Software</option>
                                    <option value="9"<%=pcvBlogCatSGP%>>Site | General Page</option>
                                    <option value="8"<%=pcvBlogCatSLP%>>Site | Landing Page</option>
                                </select>
                            </td>
                        </tr>
                         <tr>
                        	<td nowrap>Published Date:</td>
                            <td>
                                <input type="text" name="blogPubDate" size="12" value="<%=pcvBlogPubDate%>"> For sorting only (yyyy-mm-dd)
                            </td>
                        </tr>
                        <tr>
                                <td nowrap>Review Status:</td>
                                <td>
                                    <input type="radio" name="Published" value="1" <%if pcInt_Published="1" then%>checked<%end if%> class="clearBorder"> Published
                                    &nbsp;
                                    <input type="radio" name="Published" value="0" <%if pcInt_Published="0" then%>checked<%end if%> class="clearBorder"> Under Review
                                    &nbsp;
                                    <input type="radio" name="Published" value="2" <%if pcInt_Published="2" then%>checked<%end if%> class="clearBorder"> Reviewed: Changes Needed
                                </td>
                            </tr>
					</table>
				</div>
			<%
			'// =========================================
			'// FIRST PANEL - END
			'// =========================================

			'// =========================================
			'// SECOND PANEL - START - Settings
			'// =========================================
			%>
				<div id="tabs-2" class="tab-pane">

					<table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
                        <tr>
                            <th colspan="2">Settings</th>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td colspan="2"><h2>Page Status &amp; Visibility</h2></td>
                        </tr>
                        <tr>
                        	<td>Active?</td>
                            <td>
                                <input type="checkbox" name="Active" value="1" <%if pcv_Active="1" then%>checked<%end if%> class="clearBorder"> Yes
                            </td>
                        </tr>
                        <tr>
                        	<td nowrap>Only accessible by...</td>
                            <td>
                                <select name="customerType">
                                    <option value="ALL" selected>All Customers</option>
                                    <option value="W"<% if pcv_CustomerType="W" then response.write " selected"%>>Wholesale Customer</option>
                                    <% 'if there are pricing categories - List them here
                                    
                                    query="SELECT idcustomerCategory, pcCC_Name, pcCC_CategoryType FROM pcCustomerCategories;"
                                    SET rs=Server.CreateObject("ADODB.RecordSet")
                                    SET rs=conntemp.execute(query)
                                    if NOT rs.eof then 
                                        do until rs.eof 
                                            intIdcustomerCategory=rs("idcustomerCategory")
                                            strpcCC_Name=rs("pcCC_Name")
                                            %>
                                            <option value="CC_<%=intIdcustomerCategory%>"<% if pcv_CustomerType="CC_" & intIdcustomerCategory then response.write " selected"%>><%=strpcCC_Name%></option>
                                            <% rs.moveNext
                                        loop
                                    end if
                                    SET rs=nothing
                                    
                                    %>
                                </select>
                            </td>
                        </tr>
                        <%
                        '// Don't allow to change "Published" status based on permissions
                        if session("PmAdmin") = "19" or (not isNull(findUser(pcUserArr,11,pcUserArrCount))) then %>	
                            <tr>
                                <td colspan="2" class="pcCPspacer"></td>
                            </tr>
                        	<tr>
                            	<td colspan="2"><h2>Review Status and Notes</h2></td>
                            </tr>
                        	<tr>
                            	<td colspan="2"><strong>Review Status</strong> and <strong>Review Notes</strong> are only shown to Control Panel users with <em>publishing</em> permissions.</td>
                            </tr>
                            <tr>
                                <td nowrap>Review Status:</td>
                                <td>
                                    <input type="radio" name="Published" value="1" <%if pcInt_Published="1" then%>checked<%end if%> class="clearBorder"> Published
                                    &nbsp;
                                    <input type="radio" name="Published" value="0" <%if pcInt_Published="0" then%>checked<%end if%> class="clearBorder"> Under Review
                                    &nbsp;
                                    <input type="radio" name="Published" value="2" <%if pcInt_Published="2" then%>checked<%end if%> class="clearBorder"> Reviewed: Changes Needed
                                </td>
                            </tr>
                            <tr>
                            	<td valign="top">Review Notes: <div style="margin-top: 3px;" class="pcSmallText">Not shown in the storefront</div></td>
                                <td><textarea name="PageComments" rows="7" cols="60"><%=pcv_PageComments%></textarea></td>
                            </tr>
                        <%
						else
						%>
                        <tr>
                            <td nowrap>Review Status:</td>
                            <td>
                            	<% 
								select case pcInt_Published
									case 1 
									response.write "Published"
									case 0
									response.write "Under Review"
									case 2
									response.write "Reviewed: Changes Needed"
								end select
								%>
                            	<input type="hidden" name="Published" value="<%=pcInt_Published%>">
                            </td>
                        </tr>
                        <%
                        end if
                        %>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td colspan="2"><h2>Parent (Optional)</h2></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                            To simplify <a href="../pc/viewcontent.asp" target="_blank">browsing Content Pages</a>, you can organize them in a two-level tree.
							<div style="margin-top: 6px;">
							<%
							if intAlreadyParent=0 then
								Dim pcPageParentExist, intPageCount
								
								if validNum(pcInt_idpage) then
									query="SELECT pcCont_idPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 AND pcCont_idPage<>" & pcInt_idpage & " ORDER BY pcCont_PageName ASC"
								else
									query="SELECT pcCont_idPage, pcCont_PageName FROM pcContents WHERE pcCont_Parent=0 ORDER BY pcCont_PageName ASC"	
								end if
								set rs=Server.CreateObject("ADODB.Recordset")
								set rs=connTemp.execute(query)
								if rs.EOF then
									pcPageParentExist=0
								else
									pcPageParentExist=1
									pcPageArr=rs.getRows()
								end if
								set rs=nothing
								
								if pcPageParentExist=0 then
								%>
									No Content Pages available.
									<br />
									First add a Content Page, then you can use it as a &quot;Parent&quot; of another page.
								<%
								else
								%>
									<select name="Parent" tabindex="104">
										<option value="0">None</option>
								<%
									intPageCount=ubound(pcPageArr,2)
									For m=0 to intPageCount %>
										<option value="<%=pcPageArr(0,m)%>"<% if pcPageArr(0,m)=pcInt_Parent then %>selected<% end if %>><%=pcPageArr(1,m)%></option>
								<%
									Next
								%>
									</select>
								<%
								end if
							else
								response.write "<em>This is already a Parent page. You cannot select a Parent for it.</em>"
							end if
							%>
							</div>
                            </td>
						</tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td colspan="2"><h2>Thumbnail (Optional)</h2></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->
                                <%If HaveImgUplResizeObjs=1 then%>
                                <%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_8")%>&nbsp;<a href="#" onClick="window.open('uploadresize/catResizea.asp','popup','toolbar=no,status=no,location=no,menubar=no,height=350,width=400,scrollbars=no'); return false;">click here</a>.
                                <% Else %>
                                    <%=dictLanguageCP.Item(Session("language")&"_cpInstPrd_9")%>&nbsp;<a href="#" onClick="window.open('imageuploada_popup.asp','_blank', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=400,height=360')">click here</a>.
                                <% End If %>
                                &nbsp;The thumbnail is used when browsing Content Pages.
                                <div style="margin-top: 6px;">
                                    <input type="text" name="PageThumbnail" id="PageThumbnail" value="<%=pcv_PageThumbnail%>" size="40"> <a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=PageThumbnail&fid=form1','window2')"><img src="images/pcIconSearch.jpg" alt="Locate previously uploaded images" border=0 hspace="2"></a>
                                    <%
									if trim(pcv_PageThumbnail<>"") then
									%>
                                    <a href="javascript:;" onClick="chgWin('../pc/catalog/<%=pcv_PageThumbnail%>','window3')"><img src="images/pcIconPreview.jpg" border="0"></a>
                                    <%
									end if
									%>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                        	<td colspan="2"><h2>Graphical Interface</h2></td>
                        </tr>
                        <tr>
                        	<td align="right" valign="top"><input type="text" name="Order" value="<%=pcInt_Order%>" size="3"></td>
                            <td>Display Order <br><span class="pcSmallText">Used when displaying a list of pages (e.g. navigation).</span></td>		
                        <tr>                      
                            <td align="right"><input type="checkbox" name="IncHeader" value="1" <%if pcIntIncHeader="1" or pcIntNewPage=1 then%>checked<%end if%> class="clearBorder"></td>
                            <td>Include store header &amp; footer</td>
                        </tr> 
                        <tr>                      
                            <td align="right"><input type="checkbox" name="MenuExclude" value="1" <%if pcIntMenuExclude="1" then%>checked<%end if%> class="clearBorder"></td>
                            <td>Exclude from Content Pages <a href="cmsNavigation.asp">navigation</a> and &quot;<a href="../pc/viewcontent.asp" target="_blank">Browse Pages</a>&quot;</td>
                        </tr>
                        <tr>                      
                            <td align="right" valign="top"><input type="checkbox" name="HideBackButton" value="1" <%if pcIntHideBackButton="1" then%>checked<%end if%> class="clearBorder"></td>
                            <td valign="top">Hide &quot;Back&quot; button.<br><span class="pcSmallText">Helpful if you are not using the <a href="../pc/viewcontent.asp">Browse Content Pages</a> page.</span></td>
                        </tr> 
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
					</table>
					
				</div>
			<%
			'// =========================================
			'// SECOND PANEL - END
			'// =========================================

			'// =========================================
			'// THIRD PANEL - START - Meta Tags
			'// =========================================
			%>
				<div id="tabs-3" class="tab-pane">

					<table class="pcCPcontent">	
						<tr>
							<td colspan="2" class="pcCPspacer"></td>
						</tr>
                        <tr>
                            <th colspan="3">Meta Tags</th>
                        </tr>
                        <tr>
                            <td colspan="3" class="pcCPspacer"></td>
                        </tr>
                        <tr>
                            <td valign="top" align="right">Title:</td>
                            <td valign="top" colspan="2">  
                                <input type="text" name="metatitle" size="50" value="<%=pcv_MetaTitle%>"> <span class="pcSmallText">(max 250 char.)</span>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="right">Description:</td>
                            <td valign="top" colspan="2">  
                                <textarea name="metadesc" rows="5" cols="50"><%=pcv_MetaDesc%></textarea> <span class="pcSmallText">(max 500 char.)</span>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="right">Keywords:</td>
                            <td valign="top" colspan="2">  
                                <textarea name="metakeywords" rows="5" cols="50"><%=pcv_MetaKeywords%></textarea> <span class="pcSmallText">(max 500 char.)</span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="pcCPspacer"></td>
                        </tr>
            
                    </table>
					
				</div>
			<%
			'// =========================================
			'// THIRD PANEL - END
			'// =========================================
			%>
			
			</div>
			
		<%
		'// TABBED PANELS - MAIN DIV END
		%>
    
        </div>
    </div>
	<div style="clear: both;">&nbsp;</div>
</form>
<!--#include file="AdminFooter.asp"-->
