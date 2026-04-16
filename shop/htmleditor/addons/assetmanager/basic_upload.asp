<!--#include file="server/upload.asp"-->

<%
	activeFolder = Request.QueryString("folder")
%>

<html>
	<head>
		<link href='../../../pc/css/bootstrap.min.css' rel='stylesheet' type='text/css'>
		<link href='../../../<%= scAdminFolderName %>/css/pcv4_ControlPanel.css' rel='stylesheet' type='text/css'>
		<style type="text/css">
			html, body {
				height: 100px;
			}
			body {
				margin: 0px;
				padding: 0px;
				overflow: hidden;
			}

			input {
				float: left;
				margin-right: 5px;
			}

			.pcCPmessageWarning {
				margin-top: 0px;
			}
		</style>
	</head>

	<body>
		<form name='form1' method='post' action='<%= Request.ServerVariables("SCRIPT_NAME") %>?folder=<%= activeFolder %>&message=1' enctype='multipart/form-data' onsubmit='return (document.getElementById("File1").value!="") '>
			
			<% If UploadErrorStr&"" <> "" Then %>
				<div class="pcCPmessageWarning">
					<%= UploadErrorStr %>
				</div>
			<% End If %>

			<input id='folder' name='folder' type='hidden' value='<%= activeFolder %>' />
			<input id='File1' name='Filedata' type='file' />
			<input type='submit' name='act' value='Upload' />
			<div style="clear: both"></div>
		</form>
	</body>
</html>
