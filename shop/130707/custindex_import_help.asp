<% pageTitle = "Customer Import Wizard - Instructions" %>
<% section = "mngAcc" %>
<%PmAdmin=7%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
    <td class="pcCPspacer">
        <% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
    </td>
</tr>
<tr>
	<td align="center">
		<div class="pcCPmessage">IMPORTANT <span style="font-weight: normal;">- Carefully read the <a href="http://wiki.productcart.com/productcart/customer-import" target="_blank">Customer Import Wizard documentation</a> before attempting to import or update customer information. <b>NOTE: as part of a new security standard introduced in v5.2, even if you include Customer passwords in the Import, they will still be required to reset them the first time they login to their Account in the storefront!</b></span></div>
        <div style="margin-top: 25px; text-align: center;">
		<input type="button" class="btn btn-default"  value="Proceed to Customer Import Wizard" class="btn btn-primary" onClick="location.href='custindex_import.asp'">&nbsp;
		<%
		on error resume next
		CSVFile = "importlogs/custlogs.txt"
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
				<input type="button" class="btn btn-default"  value="Undo Last <%=InsTop%>" onClick="javascript:if (confirm('You are about to undo your last customer <%=InsTop%>. All the information added to the database during the import/update will be removed. ProductCart saved a log of the information imported/updated in the file pcadmin/importlogs/custlogs.txt. You should NOT use this feature if you have further updated the customer information after having imported/updated customer data. Are you sure you want to complete this action?')) location='undocustimport.asp'">&nbsp;
			<%
			end if
			f.close
			set f=nothing
		end if
		set fso=nothing%>
		<input type="button" class="btn btn-default"  value="Help" onClick="window.open('http://wiki.productcart.com/productcart/customer-import')">
        </div>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->