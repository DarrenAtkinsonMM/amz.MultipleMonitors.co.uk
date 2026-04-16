<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
standid = request.querystring("s")

		Select Case standid
			Case "mm-s2v"
				video = "244802427"
			Case "mm-s2h"
				video = "225223896"
			Case "mm-s3h"
				video = "231508464"
			Case "mm-s3p"
				video = "244802461"
			Case "mm-s4s"
				video = "244802709"
			Case "mm-s4h"
				video = "244802515"
			Case "mm-s4sp"
				video = "244802657"
			Case "mm-s4p"
				video = "244802581"
			Case "mm-s5p"
				video = "244802780"
			Case "mm-s6r"
				video = "244802915"
			Case "mm-s6rp"
				video = "244802848"
			Case "mm-s8r"
				video = "244802995"
			Case else
				video = "225223896"
		End Select
%>
<iframe src="https://player.vimeo.com/video/<%=video%>" width="736" height="414" style="margin-left:12px;" frameborder="0" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>
