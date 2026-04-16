<%@LANGUAGE="VBSCRIPT"%>
<% Server.ScriptTimeout = 5400 %>
<% 'On Error Resume Next %>
<%PmAdmin=19%>
<% pageTitle = "ProductCart - Apparel Add-On - Fix Sub-Product Names Issues" %>
<% Section = "" %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<%

msg=""
IF request("action")="upd" THEN

query="SELECT idProduct,Description FROM Products WHERE pcProd_Apparel=1 AND removed=0;"
set rs=connTemp.execute(query)

IF not rs.eof then
	tmpArr=rs.getRows()
	set rs=nothing
	intC=ubound(tmpArr,2)
	FOR k=0 to intC
	
		PR_id=tmpArr(0,k)
		PR_name=tmpArr(1,k)
	
		ReDim Opts(5)
		
		Opts(0)=""
		Opts(1)=""
		Opts(2)=""
		Opts(3)=""
		Opts(4)=""

		
		query="SELECT idOptionGroup FROM pcProductsOptions WHERE idproduct=" & PR_id & " ORDER BY pcProdOpt_order ASC;"
		set rs1=connTemp.execute(query)
		
		if not rs1.eof then
			pcArr=rs1.getRows()
			intCount=ubound(pcArr,2)
			set rs1=nothing
			For i=0 to intCount
				query = "SELECT options_optionsGroups.idoptoptgrp, options.optiondescrip, options.pcOpt_Code, options_optionsGroups.price, options_optionsGroups.wprice "
				query = query & "FROM options_optionsGroups "
				query = query & "INNER JOIN options "
				query = query & "ON options_optionsGroups.idOption = options.idOption "
				query = query & "WHERE options_optionsGroups.idOptionGroup=" & pcArr(0,i) &" "
				query = query & "AND options_optionsGroups.idProduct=" & PR_id &" AND options_optionsGroups.InActive=0 "
				query = query & "ORDER BY options_optionsGroups.sortOrder;"
				set rs1=server.CreateObject("ADODB.RecordSet")
				set rs1=conntemp.execute(query)
		
				do while not rs1.eof
					Opts(0)=Opts(0) & rs1("idoptoptgrp") & "||"
					Opts(1)=Opts(1) & rs1("optiondescrip") & "||"
					Opts(2)=Opts(2) & rs1("pcOpt_Code") & "||"
					Opts(3)=Opts(3) & rs1("price") & "||"
					opt_wprice=rs1("wprice")
					if opt_wprice="0" then
						opt_wprice=rs1("price")
					end if
					Opts(4)=Opts(4) & opt_wprice & "||"
					rs1.MoveNext
				loop
				set rs1=nothing
			Next
			
			tmp1=split(Opts(2),"||")
			tmp2=split(Opts(0),"||")
			tmp3=split(Opts(1),"||")
			tmp5=split(Opts(3),"||")
			tmp6=split(Opts(4),"||")
			
			query="SELECT idProduct,pcprod_Relationship FROM Products WHERE pcprod_ParentPrd=" & PR_id & " AND removed=0;"
			set rs1=connTemp.execute(query)
			
			if not rs1.eof then
				tmpSubPrd=rs1.getRows()
				intSC=ubound(tmpSubPrd,2)
				set rs1=nothing
				For i=0 to intSC
					tmpSubID=tmpSubPrd(0,i)
					tmpRela=tmpSubPrd(1,i)
					tmp7=split(tmpRela,"_")
					tmpSubName=""
					For j=1 to ubound(tmp7)
						For m=lbound(tmp2) to ubound(tmp2)
							if Clng(tmp7(j))=Clng(tmp2(m)) then
								if tmpSubName<>"" then
									tmpSubName=tmpSubName & " - "
								end if
								tmpSubName=tmpSubName & tmp3(m)
								exit for
							end if
						Next
					Next
					tmpSubName=PR_name & " (" & tmpSubName & ")"
					tmpSubName=replace(tmpSubName,"'","''")
					tmpSubName=replace(tmpSubName,"""","&quot;")
					query="UPDATE Products SET Description='" & tmpSubName & "' WHERE idProduct=" & tmpSubID & ";"
					set rs1=connTemp.execute(query)
					set rs1=nothing
				Next
			end if
			set rs1=nothing
		end if
		set rs1=nothing
		
		
		
	NEXT
END IF
set rs=nothing
call closedb()
msg="Apparel Products and Sub-Products were updated successfully!"
msgtype=1
END IF
%>
<!--#include file="Adminheader.asp"-->
<table class="pcCPcontent">
<%if msg<>"" then%>
	<tr>
		<td>
			<div class="pcCPmessageSuccess">
				<%=msg%>
			</div>
		</td>
	</tr>
	<tr>
		<td>
			<input type="button" name="backbtn" value="Back to Main menu" onclick="location='menu.asp';" class="ibtnGrey">
		</td>
	</tr>
<%else%>
<form action="FixAppSubName.asp?action=upd" method="post" name="form1" class="pcForms">
	<tr>
		<td>
			<br>
			You can use this script to fix database issues related to Apparel sub-products name.
		</td>
	</tr>
	<tr>
		<td>
			<input type="submit" name="submit1" value=" Fix now  " class="submit2">
		</td>
	</tr>
</form>
<%end if%>
</table>
<!--#include file="AdminFooter.asp"-->
