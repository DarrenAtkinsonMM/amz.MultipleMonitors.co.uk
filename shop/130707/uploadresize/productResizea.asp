<% pageTitle="Image Upload & Auto Resize" %>
<% Section="products" %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../../includes/common.asp"-->
<!--#include file="../../includes/languagesCP.asp"-->
<!--#include file="../../includes/common_checkout.asp"-->
<% 
on error resume next

thumbnailSize = 100
generalSize = 200
detailSize = 350

If Len(Session("prdThumbSize")) > 0 Then thumbnailSize = Session("prdThumbSize")
If Len(Session("prdGeneralSize")) > 0 Then generalSize = Session("prdGeneralSize")
If Len(Session("prdDetailSize")) > 0 Then detailSize = Session("prdDetailSize")
%>

<% Dim PID, barref%>

<!DOCTYPE html>
<html>
<head>
	<title>Upload Images</title>
  <link href="../css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" style="background-image: none;">
<%checkSubFolder="1"%>
<!--#include file="checkImgUplResizeObjs.asp"-->
<%If HaveImgUplResizeObjs=0 then%>
<table class="pcCPcontent">
<tr>
	<td>
		<div class="pcCPmessage">
			We have detected that your server does not have any compatible image upload/resize/crop components. You will still be able to upload images, but resizing and cropping will be disabled. <br/><br />NOTE: There may also be limitations with the size of uploaded images in IIS depending on your server configuration. Please view the <a href="http://support.microsoft.com//kb/942074" target="_blank">KB article</a> for more information.
		</div>
	</td>
</tr>
</table>
<%End If%>
<script type=text/javascript>
 var submitted = false;
 function check_submit(theform) {
   if (submitted) return false;
   theform.Submit.disabled=true;
   theform.Submit.value="Uploading...";
   return (submitted = true);
 }
</script>
<form action="productResizeb.asp" name="MyForm" method="post" enctype="multipart/form-data" onSubmit="return check_submit(this)">
	<table width="400" border="0" cellspacing="0" align="center" bgcolor="#FFFFFF">
	<tr> 
      <td> 
        <table width="90%" border="0" cellspacing="0" cellpadding="3" align="center">
          <tr> 
            <td colspan="3" bgcolor="#e5e5e5"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#000000">&nbsp;Image Upload & Auto Resize</font></b></font></td>
          </tr>
          <tr> 
            <td height="10" colspan="2"></td>
          </tr>
          <tr> 
            <td height="18" colspan="2"><font face="Arial, Helvetica, sans-serif" size="2"> 
              Select an image using the &quot;Browse&quot; button. Then click 
              on &quot;Upload&quot;. All images are automatically uploaded to 
              the &quot;<b><%= scPcFolder %>/pc/catalog</b>&quot; folder on your Web 
              server and sizes are set.</font></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"></td>
          </tr>
		  <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Image: </font></div>
            </td>
            <td width="80%"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input class=ibtng type="file" name="FILE1" size="25">
              </font></b></td>
          </tr>
		  <tr> 
            <td colspan="2" height="10"></td>
          </tr>
		<%if pcv_ResizeObj > 0 Then%>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Thumbnail Size: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="thumbnailsize" size="4" maxlength="4" class="ibtng" value="<%= thumbnailSize %>"> pixels
              </font></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">General Size: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="generalsize" size="4" maxlength="4" class="ibtng" value="<%= generalSize %>"> pixels
              </font></td>
          </tr>
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Detail Size: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="text" name="detailsize" size="4" maxlength="4" class="ibtng" value="<%= detailSize %>"> pixels
              </font></td>
          </tr>		  		     
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Resize Based On: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="radio" name="resizexy" value="Width" checked> Width&nbsp;&nbsp;&nbsp;<input type="radio" name="resizexy" value="Height" > Height
              </font></td>
          </tr>		  
          <tr> 
            <td width="20%" nowrap> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" size="2">Sharpen Image: </font></div>
            </td>
            <td width="80%"><font face="Arial, Helvetica, sans-serif" size="2"> 
              <input type="radio" name="sharpen" value="1"> Yes&nbsp;&nbsp;&nbsp;<input type="radio" name="sharpen" value="0" checked> No
              </font></td>
          </tr>		  
          <tr> 
            <td colspan="2" height="15"></td>
          </tr>
		<%end if%>
          <tr>
            <td width="20%">&nbsp;</td>		  
            <td width="80%"> 
              <div align="left"> 
                <font face="Arial, Helvetica, sans-serif" size="2"> 
                  <input type="submit" name="Submit" value="Upload">
                  <input type="button" class="btn btn-default"  value="Close Window" onClick="javascript:window.close();">
                 </font>
              </div>
            </td>
          </tr>
        </table>
      </td>
	</tr>
	</table>
</form>
</body>
</html>
