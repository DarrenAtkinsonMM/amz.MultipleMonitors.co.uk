<!DOCTYPE html>
<html>

<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Shipping Wizard - Print Label" %>
<% Section="mngAcc" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/FedExWSconstants.asp"-->
<!--#include file="../includes/pcFedExWSClass.asp"-->
<!--#include file="uploadresize/checkImgUplResizeObjs.asp"-->

<head>
	<title></title>
	<!--#include file="inc_jquery.asp"-->

<%
Const iPageSize=5

Dim iPageCurrent, varFlagIncomplete, strORD, pcv_intOrderID


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: ON LOAD
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'// SET PAGE NAMES
pcPageName = "FedExWS_ManageShipmentsPrinting.asp"
ErrPageName = "FedExWS_ManageShipmentsPrinting.asp"

'// PAGE SETTINGS
fedExLabelsDir = "FedExLabels"
defaultScreenDPI = 96
defaultLabelResolution = 200

'// SET THE FEDEX OBJECT
set objFedExClass = New pcFedExWSClass

'// GET LABEL AND RESOLUTION
pLabelFile = getUserInput(Request.QueryString("label"), 0)
pLabelResolution = getUserInput(Request.QueryString("res"), 0)
pCurrentPath =  fedExLabelsDir & "/" & pLabelFile

If pLabelResolution & "" = "" Or pLabelResolution < 75 Then
	pLabelResolution = defaultLabelResolution
End If

If InStr(pCurrentPath, "PDF") > 0 Then
	Response.Redirect pCurrentPath
Else
	
	imageTransform = Round(defaultScreenDPI / pLabelResolution, 2)

	'// LABEL SIZE
	If pcv_ResizeObj <> 0 Then
		origImageWidth = 0
		origImageHeight = 0

		call GetImageDimensions(Server.MapPath(fedExLabelsDir) & "/" & pLabelFile, origImageWidth, origImageHeight)

		iHorizontal=(imageTransform*origImageWidth)
		iVertical=(imageTransform*origImageHeight)

		imageHeight = "height: " & iVertical & "px;"
		imageWidth = "width: " & iHorizontal & "px;"
	Else 
		'// Revert to using JavaScript if no image resizer available
		%>
		<script type=text/javascript>
			$pc(window).on('load', function() {
				$pc("img").each(function () {
					var width = $pc(this).width();

					new_width = Math.round(width * <%= imageTransform %>);
					$pc(this).width(new_width);

					$pc(this).css("visibility", "visible");
				});
			});
		</script>
		<% 
	End If %>
</head>
<body style="margin: 0px; padding: 0px;">
	<div class="pcFedExPrintContainer">
		<img src="<%= pCurrentPath %>" style="<%= imageHeight %> <%= imageWidth %>; <% if pcv_ResizeObj = 0 then response.write "visibility: hidden;" %>" border="0" />
	</div>
	<%
	End If

	'// DESTROY THE FEDEX OBJECT
	set objFedExClass = nothing
	%>
</body>
</html>
