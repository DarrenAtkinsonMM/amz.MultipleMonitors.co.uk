<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% PmAdmin=19 %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcMobileSettings.asp"-->
<% 
pageTitle="Theme Editor"
pageIcon="pcv4_icon_settings.png"
section="layout"

Dim pcv_strPageName
pcv_strPageName = "ThemeEditor.asp"
%>

<!--#include file="AdminHeader.asp"-->

<%
Dim idTheme, editFile
idTheme = getUserInput(Request("idtheme"),0)
editFile = getUserInput(Request("file"),0)
if editFile = "" then
    editFile = "theme.css"
end if

if idTheme = "" then
    query = "SELECT pcThemes_Id, pcThemes_Name FROM pcThemes WHERE pcThemes_Active = 1;"
else
    query = "SELECT pcThemes_Id, pcThemes_Name FROM pcThemes WHERE pcThemes_Id = " & idTheme
end if
set rs = connTemp.execute(query)
If Not rs.eof Then
    idTheme = rs("pcThemes_Id")
    pcStrTheme = rs("pcThemes_Name")
End If

If PPD = "1" Then
    ThemePath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/" & pcStrTheme
    StylePath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/" & pcStrTheme & "/css"
	JSPath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/" & pcStrTheme & "/js"
Else
    ThemePath = Server.MapPath("../pc/theme/") & "/" & pcStrTheme
    StylePath = Server.MapPath("../pc/theme/") & "/" & pcStrTheme & "/css"
	JSPath = Server.MapPath("../pc/theme/") & "/" & pcStrTheme & "/js"
End If

Dim objFSO, objFile, objFolderContents
Set objFSO = CreateObject("Scripting.FileSystemObject")

If lcase(right(editFile, 4)) = ".asp" then
    Set objFile = objFSO.OpenTextFile(ThemePath & "/" & editFile, 1, false)
    objFolderContents = objFile.ReadAll
    objFile.close
    editorMode = "vbscript"
elseif lcase(right(editFile, 4)) = ".css" then
    Set objFile = objFSO.OpenTextFile(StylePath & "/" & editFile, 1, false)
    objFolderContents = objFile.ReadAll
    objFile.close
    editorMode = "css"
elseIf lcase(right(editFile, 3)) = ".js" then
    Set objFile = objFSO.OpenTextFile(JSPath & "/" & editFile, 1, false)
    objFolderContents = objFile.ReadAll
    objFile.close
    editorMode = "javascript"
end if

If Request.Form("content") <> "" Then

    'objFolderContents = getUserInput(Request.Form("content"),0)
    'objFolderContents = replace(objFolderContents,"&lt;%","<%")
    'objFolderContents = replace(objFolderContents,"%&gt;","%"& ">")
    'objFolderContents = replace(objFolderContents,"&lt;","<")
    'objFolderContents = replace(objFolderContents,"&gt;",">")        
    objFolderContents = Request.Form("content")
    
    If editorMode = "css" Then
        findfile = StylePath & "/" & editFile
	ElseIf editorMode = "javascript" Then
        findfile = JSPath & "/" & editFile        
	Else
		findfile = ThemePath & "/" & editFile
    End If

    Set fso=server.CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(findfile)
    Err.number = 0
    if Err.number > 0 then
        response.redirect "../" & scAdminFolderName & "/techErr.asp?error=" & Server.URLEncode("Permissions Not Set to Modify Template files.")
    end if
    Set f = nothing
    
    Set f = fso.OpenTextFile(findfile, 2, True)
    f.Write trim(objFolderContents)
    f.Close
    
    msg = "<strong>" & editFile & "</strong> has been successfully updated."
    
    Set fso = nothing
    Set f = nothing
End If
%>

<% If msg <> "" Then %>
    <div class="pcCPmessageSuccess"><% =msg %></div>
<% End If %>

<% If scEnableBundling = "1" Then %>
    <div class="pcCPmessageInfo">
        Enable "Combine &amp; Minify CSS / JavaScript" is currently enabled under <a href="AdminSettings.asp">Store Settings</a>.  Any changes you make with the Theme Editor will not take effect until you click "Combine &amp; Minify All" under <a href="pcTSUtility.asp#rebundle">Developer Tools</a>.
    </div>
<% End If %>

<form action="ThemeEditor.asp" method="post" name="form1">
    <pre id="editor" style="height: 550px;"><% =Server.HTMLEncode(objFolderContents) %></pre>
    <input type="hidden" name="file" value="<% =editFile %>" />
    <input type="hidden" name="idtheme" value="<% =idTheme %>" />
    <input type="hidden" name="content" id="content" value="" />
    <input type="submit" value="Save" class="btn btn-primary" style="float: right;" />
</form>
<!-- load ace -->
<script src="../includes/javascripts/ace/ace.js"></script>
<!-- load ace language tools -->
<script src="../includes/javascripts/ace/ext-language_tools.js"></script>
<script>
    // trigger extension
    ace.require("ace/ext/language_tools");
    var editor = ace.edit("editor");
    editor.session.setMode("ace/mode/<%=editorMode%>");
    editor.setTheme("ace/theme/tomorrow");
    // enable autocompletion and snippets
    editor.setOptions({
        enableBasicAutocompletion: true,
        enableSnippets: true,
        enableLiveAutocompletion: false
    });
    editor.getSession().setUseWrapMode(true);
	
	var content = $("#content");
	editor.getSession().on('change', function () {
       content.val(editor.getSession().getValue());
   });
</script>

</div>
<div id="pcCPmainRight">
    <br /><br />
    Select theme :
    <%
        query = "SELECT pcThemes_Id, pcThemes_Name, pcThemes_Active FROM pcThemes;"
        set rs = connTemp.execute(query)
        If Not rs.eof Then
          arrThemes = rs.GetRows()
        End If
    %>
    <select name="ThemeFolder" onChange="window.location.href = 'ThemeEditor.asp?idtheme='+this.options[this.selectedIndex].value;">
    <%
        For i = 0 to UBound(arrThemes, 2)
            ThemeId = arrThemes(0, i)
            ThemeName = arrThemes(1, i)
            ThemeStatus = arrThemes(2, i)
    %>
            <option value="<%= ThemeId %>" <% If ThemeId = idTheme Then Response.Write "selected" %>><%= ThemeName %></option>
    <%
        Next
    %>
    </select>
    
    <%    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    %>
	<!--
    <h3>Templates</h3>
    <hr />
    <ul>
        <%
        Set objFolder = objFSO.GetFolder(ThemePath)
        Set objFolderContents = objFolder.Files
        
        For Each objFileItem In objFolderContents
            If lcase(right(objFileItem.Name, 4)) = ".asp" then
                if objFileItem.Name = editFile then
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """ style=""font-weight: 900;"">" & objFileItem.Name & "</a></li>"
                else
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """>" & objFileItem.Name & "</a></li>"
                end if
            End If
        Next
        
        Set objFolder = Nothing
        Set objFolderContents = Nothing
        %>
    </ul>
	-->
    
    <% If objFSO.FolderExists(StylePath) then %>
    <h3>Styles</h3>
    <hr />
    <ul>
        <%
        Set objFolder = objFSO.GetFolder(StylePath)
        Set objFolderContents = objFolder.Files
        
        For Each objFileItem In objFolderContents
            If lcase(right(objFileItem.Name, 4)) = ".css" then
                if objFileItem.Name = editFile then
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """ style=""font-weight: 800;"">" & objFileItem.Name & "</a></li>"
                else
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """>" & objFileItem.Name & "</a></li>"
                end if
            End If
        Next
        
        Set objFolder = Nothing
        Set objFolderContents = Nothing
        %>
    </ul>
    <% end if %>
    
    <% If objFSO.FolderExists(JSPath) then %>
    <h3>Javascripts</h3>
    <hr />
    <ul>
        <%
        Set objFolder = objFSO.GetFolder(JSPath)
        Set objFolderContents = objFolder.Files
        
        For Each objFileItem In objFolderContents
            If lcase(right(objFileItem.Name, 3)) = ".js" then
                if objFileItem.Name = editFile then
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """ style=""font-weight: 800;"">" & objFileItem.Name & "</a></li>"
                else
                    Response.Write "<li><a href=""" & pcv_strPageName & "?idtheme=" & idTheme & "&file=" & objFileItem.Name & """>" & objFileItem.Name & "</a></li>"
                end if
            End If
        Next
        
        Set objFolder = Nothing
        Set objFolderContents = Nothing
        %>
    </ul>
    <% end if %>
    
<%
Set objFSO = Nothing
pcv_strDisplayType = "1"
%>

<!--#include file="AdminFooter.asp"-->