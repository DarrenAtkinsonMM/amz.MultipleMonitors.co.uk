<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle = "Import 'Order Shipped' Information - Map fields" %>
<% section = "orders" %>
<%PmAdmin=9%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="ship-checkfields.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
if ucase(right(session("importfile"),4))=".XLS" then
	call closeDb()
    response.redirect "ship-step2-xls.asp?append=" & request("append") & "&movecat=" & request("movecat")
end if
%>
<% 
append=request("append")
if append<>"" then
	session("append")=append
else
	append=session("append")
end if
movecat=request("movecat")
if movecat<>"" then
else
	movecat="1"
end if
session("movecat")=movecat
if append="1" then
	requiredfields = 1
else
	requiredfields = 4
end if

sub displayerror(msg) %>
	<!--#include file="pcv4_showMessage.asp"-->
<% end sub %>

<table class="pcCPcontent">
    <tr>  
        <td colspan="2"><h2>Steps:</h2></td>
    </tr>
    <tr>
        <td  width="5%" align="right"><img border="0" src="images/step1.gif"></td>
        <td width="95%"><font color="#A8A8A8">Upload data file</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step2a.gif"></td>
        <td><font color="#000000"><strong>Map fields</strong></font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step3.gif"></td>
        <td><font color="#A8A8A8">Confirm mapping</font></td>
    </tr>
    <tr>
        <td align="right"><img border="0" src="images/step4.gif"></td>
        <td><font color="#A8A8A8"><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</font></td>
    </tr>
</table>

		<%
		FileCSV = "../pc/catalog/" & session("importfile")
		if PPD="1" then
			FileCSV="/"&scPcFolder&"/pc/catalog/"& session("importfile")
		end if
		findit = Server.MapPath(FileCSV)
		Dim objStream, strData
		Err.number=0
		Set objStream = CreateObject("ADODB.Stream")
		
		objStream.CharSet = "utf-8"
		objStream.Open
		objStream.LoadFromFile(findit)
		
		strData = objStream.ReadText()
		
		objStream.Close
		Set objStream = Nothing
		if Err.number>0 then
			session("importfilename")=""%>
			<script type=text/javascript>
			location="msg.asp?message=31";
			</script><%
		end if
		
		ReadLine=split(strData,vbcrlf)
		
		Topline = Readline(0)

		A=split(Topline,",")
		if ubound(a)-lbound(a)+1<requiredfields then
			session("importfilename")=""%>
			<script type=text/javascript>
			location="msg.asp?message=28";
             </script><%
		end if
		validfields=0
		for i=lbound(a) to ubound(a)
			if trim(a(i))<>"" then
				validfields=validfields+1
			end if
		next
		if validfields<requiredfields then
			session("importfilename")=""%>
			<script type=text/javascript>
			location="msg.asp?message=28";
			</script><%
		end if
		session("totalfields")=ubound(a)-lbound(a)+1
		if a(ubound(a))="" then
			session("totalfields")=session("totalfields")-1
		end if
		msg=request.querystring("msg")
		if msg<>"" then 
			displayerror(msg)%>
		<% end if %>

        <div style="margin: 10px;">Use the drop-down menus below to map existing fields in your data file, located on the left side of the page under 'From' to ProductCart database fields, which are located on the right side of the page under 'To'.</div>
        <form method="post" action="ship-step3.asp" class="pcForms"> 
        <table class="pcCPcontent">
            <tr>
                <th width="50%">From:</th>
                <th width="50%">To:</th>
            </tr>
            <tr>
            	<td colspan="2" class="pcCPspacer"></td>
            </tr>
				<% validfields=0
                for i=lbound(a) to ubound(a)
                    if trim(a(i))<>"" then
                        if left(a(i),1)=chr(34) then
                            a(i)=mid(a(i),2,len(a(i)))
                        end if
                        if right(a(i),1)=chr(34) then
                            a(i)=mid(a(i),1,len(a(i))-1)
                        end if    	
                        validfields=validfields+1%>
                        <tr>
                            <td width="50%" style="border-bottom: 1px solid #ccc"><%=a(i)%>
                            <input type=hidden name="F<%=validfields%>" value="<%=a(i)%>" > 
                            <input type=hidden name="P<%=validfields%>" value="<%=i%>" >
                            </td>
                            <td width="50%" style="border-bottom: 1px solid #ccc">
                            	<select size="1" name="T<%=validfields%>">
                                    <option value="   ">   </option>
                                    <option value="Order ID">Order ID</option>
                                    <option value="Ship">Ship</option>
                                    <option value="Send Mail">Send Mail</option>
                                    <option value="Ship Date">Ship Date</option>
                                    <option value="Method">Method</option>
                                    <option value="Tracking Number">Tracking Number</option>
									<%if request("T" & validfields)<>"" then%>
                                        <option value="<%=request("T" & validfields)%>" selected><%=request("T" & validfields)%></option>
									<% else
                                        FiName=""
                                        FiName=CheckField(a(i))
                                        if FiName<>"" then%>
                                            <option value="<%=FiName%>" selected><%=FiName%></option>
                                        <%end if
                                    end if%>
                                </select>
                            </td>
                        </tr>
					<% end if
				next %>   
                <tr>
                    <td colspan="2" class="pcCPspacer"></td>
                </tr>
                <tr>
                    <td colspan="2">
                    
                        <input type="hidden" name="validfields" value="<%=validfields%>">         
                        <input type="submit" name="submit" value="Map Fields" class="btn btn-primary">&nbsp; 
                        <input type="reset" name="reset" value="Reset">  
            
                    </td>
                </tr>
            </table>
        </form>

<!--#include file="AdminFooter.asp"-->