<%'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Manage Slideshow Settings" %>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 

<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<!--#include file="../includes/pcServerSideValidation.asp" -->
<!--#include file="../includes/javascripts/pcClientSideValidation.asp" -->
<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->    
<%
pcPageName = Request.ServerVariables("SCRIPT_NAME")
pcUploadName = "AdminSlideShowUpload.asp"
			
if PPD="1" then
	catalogpath=Server.Mappath ("\"&scPcFolder&"\pc\catalog\")
else
	catalogpath=Server.Mappath ("..\pc\catalog\")
end if
catalogpath = catalogpath & "\"

if Request.QueryString("action") = "delete" then
	slideID = Request.QueryString("slide")
	
	query = "SELECT slideImage FROM pcSlideShow WHERE idSlide = " & slideID
	set rs = conntemp.execute(query)
	if not rs.eof then
		slideName = rs("slideImage")
	    on error resume next
		Set FS = Server.CreateObject("Scripting.FileSystemObject")
		If FS.FileExists(catalogpath & slideName) Then FS.DeleteFile(catalogpath & slideName)
		Set FS = nothing
	end if
			
	query = "DELETE FROM pcSlideShow WHERE idSlide = " & slideID
	conntemp.execute(query)
	
  Session("pcCPCheckCode") = 1
  Session("pcCPCheckText") = "Slide deleted successfully!"

	call closeDb()
	response.redirect pcPageName
end if

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
response.write "<script type=""text/javascript"">"&vbcrlf

response.write "function Form1_Validator(theForm)"&vbcrlf
response.write "{"&vbcrlf

StrGenericJSError=dictLanguageCP.Item(Session("language")&"_cpCommon_403")

pcs_JavaNumberBox "pauseTime", 500, 120000, true, StrGenericJSError
pcs_JavaNumberBox "animSpeed", 0, 5000, true, StrGenericJSError

response.write "return (true);"&vbcrlf
response.write "}"&vbcrlf

response.write "</script>"&vbcrlf
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' End Config Client-Side Validation
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>


<script type=text/javascript>
	$pc(document).ready(function() {
		
		$pc(".pcCPslideTools .pcCPslideReloadLink").click(function(e) {
			$pc(this).parents('li').find(".pcCPslideUpload").slideToggle('fast');
			$pc(this).parents('li').find(".pcCPslideUrl").slideToggle('fast');
			$pc(this).parents('li').find(".pcCPslideAlt").slideToggle('fast');
			
			e.preventDefault();
		});
		
		$pc(".pcCPslideLinks .pcCPslideReloadLink").click(function(e) {
			$pc(this).parents('li').find(".pcCPslideUpload").show();
			$pc(this).parents('li').find(".pcCPslideUrl").hide();
			$pc(this).parents('li').find(".pcCPslideAlt").hide();
			$pc(this).parents('li').find(".pcCPslideMore").slideToggle('fast');

			e.preventDefault();
		});
		
		$pc(".pcCPslideLinks .pcCPslideMoreLink").click(function(e) {
			$pc(this).toggleClass("expanded");
			$pc(this).parents('li').find(".pcCPslideMore").slideToggle('fast');

			e.preventDefault();
		});
		
		$pc("input").keypress(function(e) {
			if (e.which == 13) {				
				$pc("#slideSettingsSave").click();
				
				e.preventDefault();
			}
		});

		$pc("#useDefault").click(function() {
			var enabled = $pc(this).is(":checked");

			useDefaultSet(enabled);
		});
	});

	function useDefaultSet(enabled)
	{
		$pc("#mobile").find("input, select").each(function() {
			if ($pc(this).attr("id") != "useDefault" && !$pc(this).hasClass("slideSettingsSave")) {
				$pc(this).attr('disabled', enabled);
			}
		});

		$pc("#mobile").find(".pcCPslides").each(function() {
			if (enabled) {
				$pc(this).sortable("disable");
			} else {
				$pc(this).sortable("enable");
			}
		});

		$pc("#mobile").find(".pcCPslideForm").each(function() {
			if (enabled) {
				$pc(this).addClass("disabled");
			} else {
				$pc(this).removeClass("disabled");
			}
		});
	}
	
	function deleteSlide(id) {
		if (confirm("Are you sure you want to delete this slide?")) {
			window.location = "<%= pcPageName %>?slide=" + id + "&action=delete";
		}
		
		return false;
	}
	
	function addFormParams(caller, params) {
		var form = $pc(caller).closest("form");
		var newAction = form.attr('action');
		var tmpStr="";
		if (newAction.indexOf("action=") >= 0)
		{
			var tmpArr=newAction.split("&");
			for (var i = 0; i < tmpArr.length; i++)
			{
   				if (tmpArr[i].indexOf("action=")<0)
				{
					if (i==0)
					{
					 	tmpStr=tmpArr[i];
					}
					else
					{
						tmpStr=tmpStr+ "&" + tmpArr[i];
					}
				}
				else
				{
					if (tmpArr[i].indexOf("?action=")>=0) 
					{
						tmpArr1=tmpArr[i].split("?");
						tmpStr=tmpArr1[0];
					}
				}
 			}
			newAction=tmpStr;
		}

		for (var param in params) {
			if (newAction.indexOf("?") > 0) {
				newAction += "&";
			} else {
				newAction += "?";
			}
			newAction += param + "=" + params[param];
		}
		
		form.attr('action', newAction);

		return true;
	}
</script>

<%if CUmsg<>"" then
	msg="Some upload/resize components have errors:<br><ul>" & CUmsg & "</ul>"
	msgType=0
end if%>
<div>
<!--#include file="pcv4_showMessage.asp"-->
</div>

<div id="TabbedPanels2" class="tabbable">
	<ul class="nav nav-tabs">
		<li><a href="#desktop" data-toggle="tab">Desktop</a></li>
		<li><a href="#mobile" data-toggle="tab">Mobile</a></li>
	</ul>

	<div class="tab-content">
		<%Dim FCount
		FCount=0
			AddSlideShowPage "desktop", 1
			AddSlideShowPage "mobile", 2	 
		%>
	</div>

	<%
		Sub AddSlideShowPage(pageName, settingId)
	%>
		<div id="<%= pageName %>" class="tab-pane">
			<%
			'load settings
			query="SELECT * FROM pcSlideShowSettings WHERE idSetting = " & settingId & ";"
			set rstemp=Server.CreateObject("ADODB.Recordset")     
			set rstemp=conntemp.execute(query)
			pcIntSlideSettingId=rstemp("idSetting")
			pcIntSlideWidth=rstemp("slideWidth")
			pcIntSlideHeight=rstemp("slideHeight")
			pcStrEffect=rstemp("effect")
			pcIntAnimSpeed=rstemp("animSpeed")
			pcIntPauseTime=rstemp("pauseTime")
			pcIntSlideUseDefault=rstemp("useDefault")
			set rstemp = nothing

			if err.number <> 0 Then
				Session("pcCPCheckText") = "Error loading slideshow settings: " & err.description
				Session("pcCPCheckCode") = 2
  
				call closeDb()
				response.redirect pcPageName
			end if
			%>

			<form method="post" id="hForm<%=settingId%>" name="hForm<%=settingId%>" enctype="multipart/form-data" action="<%= pcUploadName %>?setting=<%= pcIntSlideSettingId %>" class="pcForms pcCPslideForm">
				<div class="pcCPslideSettings">
				<table class="pcCPcontent">
    
					<% If pageName = "mobile" Then %>
					<script type="text/javascript">
						$pc(document).ready(function() {
							useDefaultSet(<% If pcIntSlideUseDefault = 1 Then Response.Write "true" Else Response.Write "false" End If %>);
						});
					</script>
					<tr>
						<td colspan="2">
							<div class="bs-callout bs-callout-warning">
								<p>Check the option below to use the desktop settings and slide images on the mobile site. This is useful if both your desktop and mobile site will be sharing the same configuration.</p>
								<p>
									<input type="checkbox" id="useDefault" name="useDefault" value="1" <% If pcIntSlideUseDefault = 1 Then Response.Write "checked" %>/>
									<label for="useDefault">Use Desktop Settings</label>
								</p>
							</div>
						</td>
					</tr>
					<tr>
						<td colspan="2" class="pcSpacer"></td>
					</tr>
					<% End If %>

					<% If pcvMessageType <> 2 Then %>
						<tr>
  						<th colspan="2">Upload New Slides</th>
						</tr>
						<tr>
							<td colspan="2">
								<% If HaveImgUplResizeObjs = 0 Then %>
									<div class="pcCPmessage">We have detected that your server does not have any compatible image upload/resize/crop components. You will still be able to upload images, but resizing and cropping will be disabled. <br/><br />NOTE: There may also be limitations with the size of uploaded images in IIS depending on your server configuration. Please view the <a href="http://support.microsoft.com//kb/942074" target="_blank">KB article</a> for more information.</div>
								<% ElseIf pcv_ResizeObj = 0 Then %>
									<div class="pcCPmessage">We have detected that your server does not have any compatible resize/crop components. You will still be able to upload images, but resizing and cropping with be disabled.
								<% End If %>
							</td>
						</tr>
						<tr valign="top">
							<td>Choose Image: </td>
							<td>
								<input type="file" id="upload_image<%=settingId%>" name="imgupload<%=settingId%>">
								&nbsp;
								<input name="submit" type="submit" value="Upload" class="btn btn-primary" onClick="return addFormParams(this, {action: 'upload'});">
							</td>
						</tr>
						<%if pcv_ResizeObj > 0 Then%>
						<tr valign="top">
							<td>Crop New Slide: </td>
							<td>
								<input type="checkbox" name="cropslide" id="cropslide" value="1" class="clearBorder">&nbsp;&nbsp;&nbsp;<input type="number" id="slideWidth" min="16" max="20000" name="slideWidth" value="<%= pcIntSlideWidth %>" title="Slide Width"/>
								x <input type="number" id="slideHeight" min="16" max="20000" name="slideHeight" value="<%= pcIntSlideHeight %>" title="Slide Height"/>
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=850"></a><br>
								<i>Note: You will be able to select the area to crop after the upload.</i>
							</td>
						</tr>
						<%end if%>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						<tr>
  						<th colspan="2">Add Existing Slide to Slideshow</th>
						</tr>
						<tr valign="top">
							<td>Choose Image: </td>
							<td>
								<input type="text" id="addimage" name="addimage" size="30"><a href="javascript:;" onClick="chgWin('../pc/imageDir.asp?ffid=addimage&fid=hForm<%=settingId%>','window2')"><img src="images/search.gif" alt="locate images previously uploaded" width="16" height="16" border=0 hspace="3"></a>
								<br><br>
								<input name="submit" type="submit" value="Add Selected" class="btn btn-primary" onClick="return addFormParams(this, {action: 'addexist'});">
							</td>
						</tr>
						<tr>
							<td colspan="2" class="pcSpacer"></td>
						</tr>
						
						<tr>
							<td>Pause Time: </td>
							<td>              
					<input type="number" name="pauseTime" value="<%= pcIntPauseTime %>" min="500" max="1200000"/>&nbsp;ms              
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=852"></a>
							</td>
						</tr>
						<tr>
							<td>Animation Speed: </td>
							<td>
					<input type="number" name="animSpeed" value="<%= pcIntAnimSpeed %>" />&nbsp;ms
								&nbsp;<a class="pcCPhelp" href="helpOnline.asp?ref=853"></a>
							</td>
						</tr>
						<tr>
							<td colspan="2"><input name="submit" type="submit" class="slideSettingsSave btn btn-primary" value="Save Settings" onClick="return addFormParams(this, {action: 'save'});"></td>
						</tr>
						<tr> 
							<td colspan="2"><hr /></td>
						</tr>
						<tr> 
							<td class="pcCPspacer"></td>
						</tr>
						<tr>
  						<th colspan="2">Slides <span class="pcSmallText">(Click and drag the images below to sort)</span></th>
						</tr>
						<tr>
							<td colspan="2">
    						<ul class="pcCPslides pcCPsortable">
								<%
									query="SELECT idSlide,slideOrder,slideCaption,slideImage,slideUrl,slideAlt,slideNewWindow, slideStart, slideEnd FROM pcSlideShow WHERE idSetting = " & pcIntSlideSettingId
									If pageName = "desktop" Then
										query=query&" OR idSetting IS NULL"
									End If
									query=query&" ORDER BY slideOrder"
									set rstemp=Server.CreateObject("ADODB.Recordset")
									set rstemp=conntemp.execute(query)
  
									if err.number <> 0 Then
										Session("pcCPCheckCode") = 2
										Session("pcCPCheckText") = "Error loading slides: " & err.description
              
										call closeDb()
										response.redirect pcPageName
									end if
        
									if rstemp.eof then
									%>
										<li class="pcCPslideMessage pcCPnotSortable">
											<p>No slideshow images were found. Use the upload area above to start adding slides!</p>   
										</li> 
									<%
									else
										slideCount = 1
										do while not rstemp.eof
											slideId = rstemp("idSlide")
											slideImage = rstemp("slideImage")
											slideCaption = rstemp("slideCaption") & ""
											slideUrl = rstemp("slideUrl") & ""
											slideAlt = rstemp("slideAlt") & ""
											slideOrder = rstemp("slideUrl")
											slideStart = rstemp("slideStart")
											slideEnd = rstemp("slideEnd")
											If Not IsNumeric(slideOrder) Then
												slideOrder = slideCount
											End If
											if slideStart = "1900-01-01" Then
												slideStart = ""
											End if
											if slideEnd = "1900-01-01" Then
												slideEnd = ""
											end if
											if slideEnd = "2099-12-31" then
												slideEnd = ""
											end if
											slideNewWindow = rstemp("slideNewWindow")

											if InStr(slideStart,"-")  then
											d = split(slideStart,"-")
											slideStart = d(1)&"/"&d(2)&"/"&d(0)
											end if

											if InStr(slideEnd,"-")  then
											d = split(slideEnd,"-")
											slideEnd = d(1)&"/"&d(2)&"/"&d(0)
											end if


											'// Unescape from the database
											slideCaption = Replace(slideCaption, "''", "'")
											slideUrl = Replace(slideUrl, "''", "'")
											slideAlt = Replace(slideAlt, "''", "'")
						
											imageWidth = 0
											imageHeight = 0
					
											Set FS = Server.CreateObject("Scripting.FileSystemObject")
											If FS.FileExists(catalogpath & slideImage) Then
						   					GetImageDimensions catalogpath & slideImage, imageWidth, imageHeight
											End If
											Set FS = Nothing
						
											slideCount = slideCount + 1

											imageNeedReload = imageWidth > 0 And imageHeight > 0 And (imageWidth <> pcIntSlideWidth Or imageHeight <> pcIntSlideHeight)
											%>        
												<li>
													<div class="pcCPsortableHandle pcCPslideImage">
														<input type="hidden" class="pcCPsortableOrder" name="slideOrder_<%= slideId %>" value="<%= slideOrder %>" />
														<img src="../pc/catalog/<%= slideImage %>" alt="<%= slideCaption %>" />
													</div>
                
													<div class="pcCPslideInfo">
														<div class="pcCPslideCaption">
															
															<label>Caption (optional): </label>
															<input type="text" name="slideCaption_<%= slideId %>" value="<%= slideCaption %>" />
														</div>
                  
														<div class="pcCPslideLinks">
															<a href="#" onClick="deleteSlide(<%= slideId %>); return false;">Delete</a>
															<span>|</span>
															<a href="#" class="pcCPslideMoreLink">More Settings</a>
														</div>
													</div>
                
													<div class="pcCPslideMore">
                						<div class="pcCPslideTools">
															<% If imageWidth > 0 And imageHeight > 0 Then %>
																<div class="pcCPslideSize">
																	Size: <strong><%= imageWidth %> x <%= imageHeight %></strong>
																</div>

															<% End if %>
													
                				  						<ul>
				                  							<li><a href="../pc/catalog/<%= slideImage %>" target="_blank">View Slide Image</a></li>
														</ul>
														</div>
                  
														<div class="pcCPslideUrl">
															<label>Link (optional): </label>
															<input type="text" name="slideUrl_<%= slideId %>" value="<%= slideUrl %>" />
														</div>
									
														<div class="pcCPslideAlt">
															<label>Alt Tag Text (optional): </label>
															<input type="text" name="slideAlt_<%= slideId %>" value="<%= slideAlt %>" />
														</div>




                              
                              <div class="pcCPslideTools" style="margin-top:10px;">
                              	Enter the first day you want to display this slide and the last day you want this slide to display (mm/dd/yyy).
                              </div>

                              <div>
                              	<input type="checkbox" name="slideNewWindow_<%= slideId %>" value="1" <% if slideNewWindow = 1 then %>checked<% end if %>/>Open link in a new window
                              </div>




                              <div class="pcCPslideAlt" style="margin-top:10px;">
                              Display from: <input type="text" class="datepicker" name="slideStart_<%= slideId %>" value="<%= slideStart %>" readonly>
                              </div>
                              <div class="pcCPslideAlt" style="margin-top:10px;">
                              Display through: <input type="text" class="datepicker" name="slideEnd_<%= slideId %>" value="<%= slideEnd %>" readonly>
                              </div>

													</div>
                
													<div style="clear: both"></div>
												</li>
											<%
											rstemp.MoveNext
										loop
       						end if
            
									set rstemp=nothing
     						%>
								</ul>
      
							</td>
 						</tr>      
						<tr>
  						<td colspan="2"><hr></td>
						</tr>
						<tr> 
							<td colspan="2" align="center"> 
								<input name="submit" type="submit" class="btn btn-primary" value="Save" onClick="return addFormParams(this, {action: 'save'});">
								&nbsp;
								<input type="button" class="btn btn-default"  name="Button" value="Back" onClick="javascript:history.back()">
							</td>
						</tr>
					<% End If %>
				</table>
				</div>
			</form>
		</div>
	<% End Sub %>

	<script type=text/javascript>
		var tab = window.location.hash;
		if (tab != "#desktop" && tab != "#mobile") {
			tab = "#desktop";
		}
		$pc('.nav-tabs').on('click', 'a', function() {
			window.location.hash = $pc(this).attr('href');
		});

		$pc('#TabbedPanels2 a[href="' + tab + '"]').tab('show');
	</script>

</div>

<!--#include file="AdminFooter.asp"-->