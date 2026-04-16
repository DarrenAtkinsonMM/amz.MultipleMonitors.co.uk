<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%> 
<% pageTitle = "Copy Product Configuration" %>
<% section = "services" %>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<% 
dim f, pidProduct

' form parameter 
pidProduct=request("idProduct")

if pidProduct="" or pidProduct="0" then
	call closeDb()
	response.redirect "menu.asp"
end if

if request("action")="update" then
	CBTOProduct=request("prdlist")
	CBTOProduct=replace(CBTOProduct,",","")
	CBTOProduct=replace(CBTOProduct," ","")
	if CBTOProduct<>"" then
		query="SELECT configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU,pcConfPro_UseRadio FROM configSpec_products WHERE specproduct=" & CBTOProduct
		set rstemp=conntemp.execute(query)
		IF not rstemp.eof THEN 'Have configuration
			tmpArr=rstemp.GetRows()
			intCount=ubound(tmpArr,2)
			set rstemp=nothing
			For mk=0 to intCount
				query2="" & pidproduct & ","
				query2=query2 & tmpArr(0,mk) & ","
				query2=query2 &tmpArr(1,mk)& ","
				query2=query2 &tmpArr(2,mk)& ","
				if tmpArr(3,mk)<>0 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if tmpArr(4,mk)=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if tmpArr(5,mk)=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if tmpArr(6,mk)=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				query2=query2 &tmpArr(7,mk)& ","
				query2=query2 &tmpArr(8,mk)& ","
				query2=query2 &tmpArr(9,mk) & ","
				if tmpArr(10,mk)=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				tmp_Notes=tmpArr(11,mk)
				if tmp_Notes<>"" then
					tmp_Notes=replace(tmp_Notes,"'","''")
				end if
				query2=query2 & "N'" & tmp_Notes & "',"
				
				if tmpArr(12,mk)="1" then
					query2=query2 & "1,"
				else
					query2=query2 & "0,"
				end if
				if tmpArr(13,mk)="1" then
					query2=query2 & "1,"
				else
					query2=query2 & "0,"
				end if
				if IsNull(tmpArr(14,mk)) or tmpArr(14,mk)="" then
					query2=query2 & "35,"
				else
					query2=query2 & tmpArr(14,mk) & ","
				end if
				if tmpArr(15,mk)="1" then
					query2=query2 & "1,"
				else
					query2=query2 & "0,"
				end if
				if tmpArr(16,mk)="1" then
					query2=query2 & "1"
				else
					query2=query2 & "0"
				end if
				
				query1="insert into configSpec_products (specProduct,configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU,pcConfPro_UseRadio) values (" & query2 & ")" 
				set rstemp1=conntemp.execute(query1)
			Next

			'Cloning Additional Charges
			query2=""
			query="SELECT configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfCha_ShowDesc,pcConfCha_ShowImg,pcConfCha_ImgWidth,pcConfCha_ShowSKU FROM configSpec_Charges WHERE specproduct=" & CBTOProduct
			set rstemp=conntemp.execute(query)
			do while not rstemp.eof
				query2="" & pidproduct & ","
				query2=query2 & rstemp("configProduct") & ","
				query2=query2 &rstemp("price")& ","
				query2=query2 &rstemp("Wprice")& ","
				if rstemp("cdefault")<>0 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if rstemp("showInfo")=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if rstemp("requiredCategory")=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				if rstemp("multiSelect")=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				query2=query2 &rstemp("prdSort")& ","
				query2=query2 &rstemp("catSort")& ","
				query2=query2 &rstemp("configProductCategory") & ","
				if rstemp("displayQF")=-1 then
					query2=query2 & "-1,"
				else
					query2=query2 & "0,"
				end if
				query2=query2 & "N'" & rstemp("Notes") & "',"
				
				if rstemp("pcConfCha_ShowDesc")="1" then
					query2=query2 & "1,"
				else
					query2=query2 & "0,"
				end if
				
				if rstemp("pcConfCha_ShowImg")="1" then
					query2=query2 & "1,"
				else
					query2=query2 & "0,"
				end if
				
				if IsNull(rstemp("pcConfCha_ImgWidth")) or rstemp("pcConfCha_ImgWidth")="" then
					query2=query2 & "35,"
				else
					query2=query2 & rstemp("pcConfCha_ImgWidth") & ","
				end if
				
				if rstemp("pcConfCha_ShowSKU")="1" then
					query2=query2 & "1"
				else
					query2=query2 & "0"
				end if

				query1="insert into configSpec_Charges (specProduct,configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfCha_ShowDesc,pcConfCha_ShowImg,pcConfCha_ImgWidth,pcConfCha_ShowSKU) values (" & query2 & ")" 
				set rstemp1=conntemp.execute(query1)
				rstemp.movenext
			loop
			
			call updPrdEditedDate(pidproduct)
			'End Cloning Additional Charges


			'// Cloning Configurator+ Conflict Management Rules
			If statusCM="0" OR scCM=1 Then

				query="SELECT pcBR_ID,pcBR_IDBTOPrd,pcBR_IDSourcePrd,pcBR_isCAT,pcBR_Must_Exists,pcBR_CanNot_Exists,pcBR_CatMust_Exists,pcBR_CatCanNot_Exists FROM pcBTORules WHERE pcBR_IDBTOPrd=" & CBTOProduct
				set rstemp=conntemp.execute(query)
				if not rstemp.eof then
					pcv_tmpArr=rstemp.getRows()
					intCount=ubound(pcv_tmpArr,2)
					set rstemp=nothing
	
					For i=0	to intCount
						query="INSERT INTO pcBTORules (pcBR_IDBTOPrd,pcBR_IDSourcePrd,pcBR_isCAT,pcBR_Must_Exists,pcBR_CanNot_Exists,pcBR_CatMust_Exists,pcBR_CatCanNot_Exists) VALUES (" & pidproduct & "," & pcv_tmpArr(2,i) & "," & pcv_tmpArr(3,i) & "," & pcv_tmpArr(4,i) & "," & pcv_tmpArr(5,i) & "," & pcv_tmpArr(6,i) & "," & pcv_tmpArr(7,i) & ");"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
						query="SELECT pcBR_ID FROM pcBTORules WHERE pcBR_IDBTOPrd=" & pidproduct & " AND pcBR_IDSourcePrd=" & pcv_tmpArr(2,i) & " AND pcBR_isCAT=" & pcv_tmpArr(3,i)
						set rstemp=connTemp.execute(query)
						pcv_NewBRID=rstemp("pcBR_ID")
						set rstemp=nothing
						if pcv_tmpArr(4,i)>"0" then
							query="SELECT pcBRMust_Item FROM pcBRMust WHERE pcBR_ID=" & pcv_tmpArr(0,i)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For j=0 to intCount1
									query="INSERT INTO pcBRMust (pcBR_ID,pcBRMust_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,j) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(5,i)>"0" then
							query="SELECT pcBRCanNot_Item FROM pcBRCanNot WHERE pcBR_ID=" & pcv_tmpArr(0,i)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For j=0 to intCount1
									query="INSERT INTO pcBRCanNot (pcBR_ID,pcBRCanNot_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,j) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(6,i)>"0" then
							query="SELECT pcBRCatMust_Item FROM pcBRCatMust WHERE pcBR_ID=" & pcv_tmpArr(0,i)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For j=0 to intCount1
									query="INSERT INTO pcBRCatMust (pcBR_ID,pcBRCatMust_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,j) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(7,i)>"0" then
							query="SELECT pcBRCatCanNot_Item FROM pcBRCatCanNot WHERE pcBR_ID=" & pcv_tmpArr(0,i)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For j=0 to intCount1
									query="INSERT INTO pcBRCatCanNot (pcBR_ID,pcBRCatCanNot_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,j) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
					Next
				end if
				set rstemp=nothing

			End If

			call closeDb()
			response.redirect "modBTOconfiga.asp?idproduct=" & pidproduct
		ELSE 'Does not have configuration
			msg="The selected configurable product has not been configured yet."
		END IF
	else 'Did not select Source Configurable Product
		msg="Please select an existing configurable product."
	end if
end if 'End of Update Action
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<%if msg<>"" then%>
<tr>
	<td>
		<div class="pcCPmessage"><%=msg%></div>
	</td>
</tr>
<%end if%>
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Copy Product Configuration"
				src_FormTips1="Use the following filters to look for the configurable product whose settings you would like to copy to other configurable products."
				src_FormTips2="Select the configurable product whose configuration you would like to copy to other configurable products.."
				src_IncNormal=0
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=2
				src_ShowLinks=0
				src_FromPage="ApplyBTOConfiga.asp?idproduct=" & pidproduct
				src_ToPage="ApplyBTOConfiga.asp?action=update&idproduct=" & pidproduct
				src_Button1=" Search "
				src_Button2=" Apply product configuration "
				src_Button3=" Back "
				src_PageSize=15
				UseSpecial=1
				session("srcprd_from")=""
				session("srcprd_where")=" AND products.idproduct<>" & pidproduct & " "
			%>
				<!--#include file="inc_srcPrds.asp"-->
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
set rstemp=nothing
set rstemp1=nothing

%>
<!--#include file="Adminfooter.asp"-->