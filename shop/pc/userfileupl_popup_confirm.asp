<!--#include file="CustLIv.asp"-->
<!--#include file="../includes/common.asp"-->
<%session("uploaded")="1"%>
<!DOCTYPE html>
<html>
<head>
<title>Upload Data File(s)</title>
<script type=text/javascript>
function winclose()
{
opener.document.hForm.uploaded.value="1";
opener.document.hForm.submit();
self.close();
}
</script>
<!--#include file="inc_headerv5.asp" -->
</head>

<body id="pcPopup" onUnload="javascript:winclose();">
	<div id="pcMain">
		<div class="pcMainContent">
			<h2>Upload Data File(s)</h2>
					
			<div class="pcSuccessMessage">
				File(s) uploaded successfully!
			</div>
			
			<div class="pcFormButtons">
				<button class="pcButtonCloseWindow" onClick="javascript:winclose(); return false;">Close Window</button>
			</div>
		</div>
	</div>
</body>
</html>
<% call closeDb() %>