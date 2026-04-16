<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/pcMobileSettings.asp"-->
<% 
pageTitle="Theme Settings"
pageIcon="pcv4_icon_settings.png"
section="layout"
%>
<%
Dim pcv_strPageName
pcv_strPageName="ThemeSettings.asp"

msg = getUserInput(Request("msg"), 0)

'// Index Themes
call pcs_IndexThemeFolder()

If Request("idtheme") > 0 Then

  query = "UPDATE pcThemes SET pcThemes_Active = 0;"
  Set rs = server.CreateObject("ADODB.RecordSet")
  Set rs = conntemp.execute(query)
  Set rs = Nothing
  
  query = "UPDATE pcThemes SET pcThemes_Active = 1 WHERE pcThemes_Id = " & Request("idtheme")
  Set rs = server.CreateObject("ADODB.RecordSet")
  Set rs = conntemp.execute(query)
  Set rs = Nothing
  
  query = "SELECT pcThemes_Name FROM pcThemes WHERE pcThemes_Id = " & Request("idtheme")
  Set rs = server.CreateObject("ADODB.RecordSet")
  Set rs = conntemp.execute(query)
  If Not rs.Eof Then
    pcStrThemeFolder = rs("pcThemes_Name")
  End If
  Set rs = Nothing
  
  call pcs_SaveThemeToSettings("theme/" & pcStrThemeFolder)
  
  If err.number <> 0 Then
    call LogErrorToDatabase()
    Set rs = Nothing
    call closedb()
    response.redirect "techErr.asp?err= " & pcStrCustRefID
  Else
    msg="success"
  End If
  
  Set rs = nothing
End If
%>
<!--#include file="AdminHeader.asp"-->

<%
If Request("action") = "search" Then
    query = "SELECT pcThemes_Id, pcThemes_Name, pcThemes_Active FROM pcThemes WHERE pcThemes_Name LIKE '%" & Request("keyword") & "%';"
Else
    query = "SELECT pcThemes_Id, pcThemes_Name, pcThemes_Active FROM pcThemes;"
End If

Set rs = server.CreateObject("ADODB.RecordSet")
Set rs = conntemp.execute(query)

Dim arrThemes
If Not rs.Eof Then
    arrThemes = rs.GetRows()
    pcv_intThemeCount = Ubound(arrThemes, 2) + 1
    pcv_boolHasResults = True
End If
Set rs = nothing

if IsNull(pcStrThemeFolder) OR pcStrThemeFolder="" then
    pcStrThemeFolder="theme/basic_blue"
end if
%>
<div class="pcThemeDetail">

    <form method="post" name="form_search" action="<%=pcv_strPageName%>?action=search" class="pcForms">
        <table class="pcCPcontent">
            <tr>
                <th colspan="2">Search Theme</th>
            </tr>
            <tr>
                <td colspan="2" class="pcSpacer"></td>
            </tr>
            <tr>
                <td width="270"><input type="text" name="keyword" value="" size="35" /></td>
                <td><input name="Submit" type="submit" value="Go" class="btn btn-info"></td>
            </tr>
        </table>
    </form>

    <form method="post" name="form1" action="<%=pcv_strPageName%>?action=add" class="pcForms">
        <table class="pcCPcontent">
            <% if msg<>"" then %>
                <tr>
                    <td colspan="2">
                        <% If msg <> "success" Then %>
                            <div class="pcCPmessage"><%=msg%></div>
                        <% Else %>
                            <div class="pcCPmessageSuccess">Theme successfully changed!</div>
                        <% End If %>
                    </td>
                </tr>
            <% end if %>
            <tr>
                <td class="pcCPspacer" colspan="2"></td>
            </tr>
            <tr>
                <th colspan="2"><%=dictLanguageCP.Item(Session("language")&"_cpSettings_81")%></th>
            </tr>
            <tr>
                <td colspan="2" class="pcSpacer"></td>
            </tr>
            <tr>
                <td>
                <% If pcv_boolHasResults = True Then %>
                    <div style="width: 100%;">
                        <%
                        If PPD="1" Then
                            ThemePath = Server.MapPath("/" & scPcFolder & "/pc/theme/") & "/"
                        Else
                            ThemePath = Server.MapPath("../pc/theme/") & "/"
                        End If
                        
                        For i = 0 to UBound(arrThemes, 2)
                            ThemeId = arrThemes(0, i)
                            ThemeName = arrThemes(1, i)
                            ThemeStatus = arrThemes(2, i)
                            
                            ThumbnailImage = ThemePath & "/" & ThemeName & "/" & ThemeName & ".jpg"
    
                            If ThemeStatus = True Then
                                CurrentTheme = ThemeName
                                InactiveClass = "activated"
                                HideActivateButton = "style=""display: none;"""
                            Else
                                InactiveClass = "inactive"
                                HideActivateButton = ""
                            End If
                            %>
                            <div class="col-md-4">
                                <div class="thumbnail">
                                    <div class="<%=InactiveClass%>">
                                        <h4>&nbsp;</h4>
                                        <a href="../pc/sandbox.asp?theme=<%=ThemeName%>" class="label label-danger" title="Live Preview" rel="tooltip" target="_blank">Preview</a>
                                        <img src="images/spacer.gif" width="10" height="1">
                                        <a href="ThemeSettings.asp?idtheme=<%=ThemeId%>" class="label label-default" rel="tooltip" title="Activate Now" <%=HideActivateButton%>>Activate</a>
                                    </div>
                                    <% 
                                    SPathInfo = pcf_getTruePath()	
                                    if Right(SPathInfo, 1)="/" then
                                        SPathInfo=SPathInfo & "pc/home.asp?theme="
                                    else
                                        SPathInfo=SPathInfo & "/pc/home.asp?theme="
                                    end if
                                    pcv_SafePath = Server.URLEncode(SPathInfo) 

                                    Dim ThemeFS
                                    Set ThemeFS = Server.CreateObject("Scripting.FileSystemObject")
                                    If ThemeFS.FileExists(ThumbnailImage) Then
                                        pcv_strThemeIcon = "../pc/theme/" & ThemeName & "/" & ThemeName & ".jpg"
                                        %>
                                        <img class="lazy" data-original="<%=pcv_strThemeIcon%>" width="225" height="169" />
                                        <%
                                    Else
                                        pcv_strThemeIcon = "../pc/theme/_common/images/icon.png"
                                        %>
                                        <div style="width: 225px; height: 140px;">
                                            <script type="text/javascript"
                                            src="//api.grabz.it/services/javascript.ashx?key=OTg1ZmNhN2NkN2YwNDg0ZGEwNTBhOTg3MzE1YjNmYzc=&url=<%=pcv_SafePath%><%=ThemeName%>&onfinish=loaded<%=i%>&width=225&height=169&format=jpg">
                                            </script>
                                            <script type="text/javascript">
                                            function loaded<%=i%>(id) { $.get( "ThemeDownload.asp?id=" + id + "&theme=<%=ThemeName%>", function( data ) {}); }
                                            </script>
                                        </div>
                                        <%
                                    End If
                                    Set ThemeFS = Nothing
                                    %>
  
                                </div>
                                <h4><%=pcf_displayThemeName(ThemeName) %></h4>
                            </div>
                        <%
                        Next
                        %>
                    </div>
                <% Else %>
                    <input type="text" name="ThemeFolder" value="<%=pcStrThemeFolder%>">
                <% End If %>
    
                </td>
            </tr>
            <tr>
                <td colspan="2" class="pcCPspacer"></td>
            </tr>
        </table>
    </form>
</div>
<script src="../includes/javascripts/jquery.lazyload.js" type="text/javascript"></script>
<script>
$( document ).ready(function() {
    $("[rel='tooltip']").tooltip();    
 
    $('.thumbnail').hover(
        function(){
            $(this).find('.inactive').fadeIn(250);
        },
        function(){
            $(this).find('.inactive').fadeOut(250);
        }
    ); 
});
$("img.lazy").lazyload({
    effect : "fadeIn"
});
</script>

<!--#include file="AdminFooter.asp"-->
