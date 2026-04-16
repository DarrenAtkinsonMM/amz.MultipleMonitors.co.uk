<% pageTitle = "Category Import Wizard - Instructions" %>
<% section = "products" %>
<%PmAdmin=2%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
on error resume next
%>
<table class="pcCPcontent">
<tr>
	<td>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
        
		<div class="pcCPmessage" style="width: 500px;">IMPORTANT <span style="font-weight: normal;">- Carefully read the <a href="http://wiki.productcart.com/productcart/products_category_import" target="_blank">Category Import Wizard documentation</a> before attempting to import or update category information.</span></div>
		<p align="center">
		<input type="button" class="btn btn-default"  value="Proceed to Category Import Wizard" class="btn btn-primary" onClick="location.href='catindex_import.asp'">&nbsp;
		<%
		CSVFile = "importlogs/categorylogs.txt"
		findit = Server.MapPath(CSVfile)
		Set fso = server.CreateObject("Scripting.FileSystemObject")
		Err.number=0
		MyTest=1
		Set f = fso.OpenTextFile(findit, 1)
		if Err.number>0 then
			MyTest=0
			Err.number=0
			Err.Description=0
		end if
		if MyTest=1 then
			Topline = f.Readline
			InsTop=""
			if TopLine="IMPORT" then
				InsTop="Import"
			end if
			if TopLine="UPDATE" then
				InsTop="Update"
			end if
			if InsTop<>"" then
			%>
				<input type="button" class="btn btn-default"  value="Undo Last <%=InsTop%>" class="btn btn-primary" onClick="javascript:if (confirm('You are about to undo your last Category <%=InsTop%>. All the information added to the database during the import/update will be removed. ProductCart saved a log of the information imported/updated in the file pcadmin/importlogs/categorylogs.txt. You should NOT use this feature if you have further updated the Category information after having imported/updated Category data. Are you sure you want to complete this action?')) location='undocatimport.asp'">&nbsp;
			<%
			end if
			f.close
			set f=nothing
		end if
		set fso=nothing%>
	<input type="button" class="btn btn-default"  value="Help" onClick="window.open('http://wiki.productcart.com/productcart/products_category_import')">
	</p>
	</td>
</tr>
<tr>
	<td>
		<p><input type="button" class="btn btn-default"  value="Export for Re-Import" onClick="location.href='ReverseCatImport_step1.asp'"></p>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->