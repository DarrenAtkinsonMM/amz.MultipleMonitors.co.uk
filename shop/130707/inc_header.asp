<script src="../includes/javascripts/pcControlPanelFunctions.js" type="text/javascript"></script>

<% IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN %>
<!--#include file="inc_jquery.asp" -->
<%if scKeepSession="1" then%>
<script type=text/javascript>
	function keepSessionAlive() {
		httpReq = new XMLHttpRequest();
		httpReq.open("GET", "keep-session-alive.asp");
		httpReq.send("");
	};
	setInterval(keepSessionAlive,120000);
</script>
<%end if%>
<!--<link rel="stylesheet" type="text/css" href="JQueryCP.css" />-->
<link rel="stylesheet" type="text/css" href="../includes/jquery/smoothmenu/ddsmoothmenuCP.css" />
<script type="text/javascript" src="../includes/jquery/smoothmenu/ddsmoothmenu.js">

/***********************************************
* Smooth Navigational Menu- (c) Dynamic Drive DHTML code library (www.dynamicdrive.com)
* This notice MUST stay intact for legal use
* Visit Dynamic Drive at http://www.dynamicdrive.com/ for full source code
***********************************************/

</script>

<script type=text/javascript>

	ddsmoothmenu.init({
		mainmenuid: "smoothmenu1", //menu DIV id
		orientation: 'h', //Horizontal or vertical menu: Set to "h" or "v"
		classname: 'ddsmoothmenu', //class added to menu's outer DIV
		contentsource: "markup", //"markup" or ["container_id", "path_to_menu_file"]
		arrowswap: true
	});

</script>
<% END IF %>

<script src="XHConn.js"></script>

<script type="text/javascript">
	// Pulled from the settings
	var scDateFrmt = "<%= scDateFrmt %>";

  // Shared Variables (productcart.js)
  var facebookActive;
  
    <%
    pcv_strRootPath = ""
    If scSSL="1" And scIntSSLPage="1" Then
        If (Request.ServerVariables("HTTPS") = "on") Then
            pcv_strRootPath = scSSLUrl
        End If
    End If
    If len(pcv_strRootPath)=0 Then
        pcv_strRootPath = scStoreURL
    End If
    If (Right(pcv_strRootPath, 1) = "/") Then
        pcv_strRootPath = left(pcv_strRootPath,len(pcv_strRootPath)-1) 
    End If
    %>   
    var pcRootUrl = '<%=pcv_strRootPath & "/" & scPcFolder %>';
</script>

<script src="../includes/javascripts/productcart.js?token=20140910900300" type="text/javascript"></script>
<script src="../includes/javascripts/productcartCP.js" type="text/javascript"></script>
<script src="../includes/javascripts/bootstrap.min.js" type="text/javascript"></script>
<script src="../includes/javascripts/bootstrap-datepicker.js" type="text/javascript"></script>
<script src="../includes/jquery/opentip/opentip-jquery.min.js" type="text/javascript"></script>
<script src="../includes/jquery/jquery.validate.min.js" type="text/javascript"></script>

<link href="screen.css" media="screen" rel="stylesheet" type="text/css">
<link href="../includes/jquery/opentip/opentip.css" rel="stylesheet" type="text/css">
<link href="../pc/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="../pc/css/datepicker3.css" rel="stylesheet" type="text/css">
<link href="bootstrap-theme.min.css" rel="stylesheet" type="text/css">
<link href="css/pcv4_ControlPanel.css" rel="stylesheet" type="text/css">
<link href="css/pcApps.css" rel="stylesheet" type="text/css">
<link href="../pc/css/pcPrint.css" media="print" rel="stylesheet" type="text/css">

<style type="text/css">
.CollapsiblePanel {
	width: 750;
}
.pcPanelTitle1 {
	font-size:14px;
	background-color:#EEE;
	font-weight:bold;
}
.pcPanelDesc {
	font-size: 12px;
	background-color: #EEE;
}
.CollapsiblePanelTab1 {
	background-color:#fff;
	border:dotted;
	border-width:thin;
	font-family:Verdana, Geneva, sans-serif;
}
.CollapsiblePanelContent {
	margin: 0px;
	padding: 0px;
	background-color:#CCC;
}
.CollapsiblePanelContentEnabled {
	margin: 0px;
	padding: 4px;
	background-color:#6F6;
}
.CollapsiblePanelContentDisabled {
	margin: 0px;
	padding: 4px;
	background-color:#FFF;
}
.pcPanelItalic {
	font-style:italic;
	color:#F60;
	font-weight:bold;
}
.pcSubmenuHeader {
	font-family:Verdana, Geneva, sans-serif;
	font-size:12px;
	font-weight:bold;
}
.pcSubmenuContent {
	font-family:Verdana, Geneva, sans-serif;
	font-size:11px;
	font-weight:normal;
	text-align:center;
}
</style>
<%call pcs_genReCaHeader()%>