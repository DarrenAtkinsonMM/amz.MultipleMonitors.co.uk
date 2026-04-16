<% 'CONFIGURATOR ONLY FILE 
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<% pageTitle = "Copy Product Configuration to another Configurable Products" %>
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
	CBTOProduct=replace(CBTOProduct," ","")
	pcArr=split(CBTOProduct,",")
	if pidProduct<>"" then
		query="SELECT configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU,pcConfPro_UseRadio FROM configSpec_products WHERE specproduct=" & pidProduct
		set rstemp=conntemp.execute(query)
		IF not rstemp.eof THEN 'Have configuration
			tmpArr=rstemp.GetRows()
			intCount=ubound(tmpArr,2)
			set rstemp=nothing
			For mk=0 to intCount
				query2=""
				query4=""
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
					query4=query4 & "showInfo=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "showInfo=0,"
				end if
				if tmpArr(5,mk)=-1 then
					query2=query2 & "-1,"
					query4=query4 & "requiredCategory=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "requiredCategory=0,"
				end if
				if tmpArr(6,mk)=-1 then
					query2=query2 & "-1,"
					query4=query4 & "multiSelect=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "multiSelect=0,"
				end if
				query2=query2 &tmpArr(7,mk)& ","
				query2=query2 &tmpArr(8,mk)& ","
				query4=query4 & "catSort=" & tmpArr(8,mk)& ","
				query2=query2 &tmpArr(9,mk) & ","
				if tmpArr(10,mk)=-1 then
					query2=query2 & "-1,"
					query4=query4 & "displayQF=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "displayQF=0,"
				end if
				tmp_Notes=tmpArr(11,mk)
				if tmp_Notes<>"" then
					tmp_Notes=replace(tmp_Notes,"'","''")
				end if
				query2=query2 & "N'" & tmp_Notes & "',"
				query4=query4 & "Notes=N'" & tmp_Notes & "',"
				
				if tmpArr(12,mk)="1" then
					query2=query2 & "1,"
					query4=query4 & "pcConfPro_ShowDesc=1,"
				else
					query2=query2 & "0,"
					query4=query4 & "pcConfPro_ShowDesc=0,"
				end if
				if tmpArr(13,mk)="1" then
					query2=query2 & "1,"
					query4=query4 & "pcConfPro_ShowImg=1,"
				else
					query2=query2 & "0,"
					query4=query4 & "pcConfPro_ShowImg=0,"
				end if
				if IsNull(tmpArr(14,mk)) or tmpArr(14,mk)="" then
					query2=query2 & "35,"
					query4=query4 & "pcConfPro_ImgWidth=35,"
				else
					query2=query2 & tmpArr(14,mk) & ","
					query4=query4 & "pcConfPro_ImgWidth=" & tmpArr(14,mk) & ","
				end if
				if tmpArr(15,mk)="1" then
					query2=query2 & "1,"
					query4=query4 & "pcConfPro_ShowSKU=1,"
				else
					query2=query2 & "0,"
					query4=query4 & "pcConfPro_ShowSKU=0,"
				end if
				if tmpArr(16,mk)="1" then
					query2=query2 & "1"
					query4=query4 & "pcConfPro_UseRadio=1"
				else
					query2=query2 & "0"
					query4=query4 & "pcConfPro_UseRadio=0"
				end if
				
				For i=lbound(pcArr) to ubound(pcArr)
					if pcArr(i)<>"" then
						query3=""
						query3=query3 & pcArr(i) & "," & query2
						
						query1="DELETE FROM configSpec_products WHERE configProduct=" & tmpArr(0,mk) & " AND configProductCategory=" & tmpArr(9,mk) & " AND specProduct=" & pcArr(i)
						set rstemp1=conntemp.execute(query1)
						set rstemp1=nothing
						
						query1="INSERT INTO configSpec_products (specProduct,configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfPro_ShowDesc,pcConfPro_ShowImg,pcConfPro_ImgWidth,pcConfPro_ShowSKU,pcConfPro_UseRadio) values (" & query3 & ")"
						set rstemp1=conntemp.execute(query1)
						set rstemp1=nothing
						
						if query4<>"" then
							query1="UPDATE configSpec_products SET " & query4 & " WHERE configProductCategory=" & tmpArr(9,mk) & " AND specProduct=" & pcArr(i)
							set rstemp1=conntemp.execute(query1)
							set rstemp1=nothing
						end if
						
						call updPrdEditedDate(pcArr(i))
						
					end if
				Next
			Next

			'Cloning Additional Charges
			query="SELECT configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfCha_ShowDesc,pcConfCha_ShowImg,pcConfCha_ImgWidth,pcConfCha_ShowSKU FROM configSpec_Charges WHERE specproduct=" & pidProduct
			set rstemp=conntemp.execute(query)
			do while not rstemp.eof
				query2=""
				query4=""
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
					query4=query4 & "showInfo=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "showInfo=0,"
				end if
				if rstemp("requiredCategory")=-1 then
					query2=query2 & "-1,"
					query4=query4 & "requiredCategory=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "requiredCategory=0,"
				end if
				if rstemp("multiSelect")=-1 then
					query2=query2 & "-1,"
					query4=query4 & "multiSelect=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "multiSelect=0,"
				end if
				query2=query2 &rstemp("prdSort")& ","
				query2=query2 &rstemp("catSort")& ","
				query4=query4 & "catSort=" & rstemp("catSort")& ","
				query2=query2 &rstemp("configProductCategory") & ","
				if rstemp("displayQF")=-1 then
					query2=query2 & "-1,"
					query4=query4 & "displayQF=-1,"
				else
					query2=query2 & "0,"
					query4=query4 & "displayQF=0,"
				end if
				tmp_Notes=rstemp("Notes")
				if tmp_Notes<>"" then
					tmp_Notes=replace(tmp_Notes,"'","''")
				end if
				query2=query2 & "N'" & tmp_Notes & "',"
				query4=query4 & "Notes=N'" & tmp_Notes & "',"
				
				if rstemp("pcConfCha_ShowDesc")="1" then
					query2=query2 & "1,"
					query4=query4 & "pcConfCha_ShowDesc=1,"
				else
					query2=query2 & "0,"
					query4=query4 & "pcConfCha_ShowDesc=0,"
				end if
				
				if rstemp("pcConfCha_ShowImg")="1" then
					query2=query2 & "1,"
					query4=query4 & "pcConfCha_ShowImg=1,"
				else
					query2=query2 & "0,"
					query4=query4 & "pcConfCha_ShowImg=0,"
				end if
				
				if IsNull(rstemp("pcConfCha_ImgWidth")) or rstemp("pcConfCha_ImgWidth")="" then
					query2=query2 & "35,"
					query4=query4 & "pcConfCha_ImgWidth=35,"
				else
					query2=query2 & rstemp("pcConfCha_ImgWidth") & ","
					query4=query4 & "pcConfCha_ImgWidth=" & rstemp("pcConfCha_ImgWidth") & ","
				end if
				
				if rstemp("pcConfCha_ShowSKU")="1" then
					query2=query2 & "1"
					query4=query4 & "pcConfCha_ShowSKU=1"
				else
					query2=query2 & "0"
					query4=query4 & "pcConfCha_ShowSKU=0"
				end if

				For i=lbound(pcArr) to ubound(pcArr)
					if pcArr(i)<>"" then
						query3=""
						query3=query3 & pcArr(i) & "," & query2
						query1="DELETE FROM configSpec_Charges WHERE configProduct=" & rstemp("configProduct") & " AND configProductCategory=" & rstemp("configProductCategory") & " AND specProduct=" & pcArr(i)
						set rstemp1=conntemp.execute(query1)
						set rstemp1=nothing
						query1="insert into configSpec_Charges (specProduct,configProduct,price,Wprice,cdefault,showInfo,requiredCategory,multiSelect,prdSort,catSort,configProductCategory,displayQF,Notes,pcConfCha_ShowDesc,pcConfCha_ShowImg,pcConfCha_ImgWidth,pcConfCha_ShowSKU) values (" & query3 & ")"
						set rstemp1=conntemp.execute(query1)
						set rstemp1=nothing
						
						if query4<>"" then
							query1="UPDATE configSpec_Charges SET " & query4 & " WHERE configProductCategory=" & rstemp("configProductCategory") & " AND specProduct=" & pcArr(i)
							set rstemp1=conntemp.execute(query1)
							set rstemp1=nothing
						end if
					end if
				Next
				rstemp.movenext
			loop
			'End Cloning Additional Charges
			
			'// Cloning Configurator+ Conflict Management Rules
			If statusCM="0" OR scCM=1 Then

				query="SELECT pcBR_ID,pcBR_IDBTOPrd,pcBR_IDSourcePrd,pcBR_isCAT,pcBR_Must_Exists,pcBR_CanNot_Exists,pcBR_CatMust_Exists,pcBR_CatCanNot_Exists FROM pcBTORules WHERE pcBR_IDBTOPrd=" & pidProduct
				set rstemp=conntemp.execute(query)
				if not rstemp.eof then
					pcv_tmpArr=rstemp.getRows()
					intCount=ubound(pcv_tmpArr,2)
					set rstemp=nothing
	
					For k=0	to intCount
					For i=lbound(pcArr) to ubound(pcArr)
					IF pcArr(i)<>"" THEN
						query="SELECT pcBR_ID FROM pcBTORules WHERE pcBR_IDBTOPrd=" & pcArr(i) & " AND pcBR_IDSourcePrd="& pcv_tmpArr(2,k)
						set rstemp=conntemp.execute(query)
						if not rstemp.eof then
							pcv_oldBRID=rstemp("pcBR_ID")
							query="DELETE FROM pcBRMust WHERE pcBR_ID=" & pcv_oldBRID
							set rstemp=conntemp.execute(query)
							set rstemp=nothing
							query="DELETE FROM pcBRCanNot WHERE pcBR_ID=" & pcv_oldBRID
							set rstemp=conntemp.execute(query)
							set rstemp=nothing
							query="DELETE FROM pcBRCatMust WHERE pcBR_ID=" & pcv_oldBRID
							set rstemp=conntemp.execute(query)
							set rstemp=nothing
							query="DELETE FROM pcBRCatCanNot WHERE pcBR_ID=" & pcv_oldBRID
							set rstemp=conntemp.execute(query)
							set rstemp=nothing
							query="DELETE FROM pcBTORules WHERE pcBR_IDBTOPrd=" & pcArr(i) & " AND pcBR_IDSourcePrd="& pcv_tmpArr(2,k)
							set rstemp=conntemp.execute(query)
							set rstemp=nothing
						end if
						set rstemp=nothing
						
						query="INSERT INTO pcBTORules (pcBR_IDBTOPrd,pcBR_IDSourcePrd,pcBR_isCAT,pcBR_Must_Exists,pcBR_CanNot_Exists,pcBR_CatMust_Exists,pcBR_CatCanNot_Exists) VALUES (" & pcArr(i) & "," & pcv_tmpArr(2,k) & "," & pcv_tmpArr(3,k) & "," & pcv_tmpArr(4,k) & "," & pcv_tmpArr(5,k) & "," & pcv_tmpArr(6,k) & "," & pcv_tmpArr(7,k) & ");"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
						query="SELECT pcBR_ID FROM pcBTORules WHERE pcBR_IDBTOPrd=" & pcArr(i) & " AND pcBR_IDSourcePrd=" & pcv_tmpArr(2,k) & " AND pcBR_isCAT=" & pcv_tmpArr(3,k)
						set rstemp=connTemp.execute(query)
						pcv_NewBRID=rstemp("pcBR_ID")
						set rstemp=nothing
						if pcv_tmpArr(4,k)>"0" then
							query="SELECT pcBRMust_Item FROM pcBRMust WHERE pcBR_ID=" & pcv_tmpArr(0,k)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For m=0 to intCount1
									query="INSERT INTO pcBRMust (pcBR_ID,pcBRMust_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,m) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(5,k)>"0" then
							query="SELECT pcBRCanNot_Item FROM pcBRCanNot WHERE pcBR_ID=" & pcv_tmpArr(0,k)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For m=0 to intCount1
									query="INSERT INTO pcBRCanNot (pcBR_ID,pcBRCanNot_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,m) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(6,k)>"0" then
							query="SELECT pcBRCatMust_Item FROM pcBRCatMust WHERE pcBR_ID=" & pcv_tmpArr(0,k)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For m=0 to intCount1
									query="INSERT INTO pcBRCatMust (pcBR_ID,pcBRCatMust_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,m) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
						if pcv_tmpArr(7,k)>"0" then
							query="SELECT pcBRCatCanNot_Item FROM pcBRCatCanNot WHERE pcBR_ID=" & pcv_tmpArr(0,k)
							set rstemp=conntemp.execute(query)
							if not rstemp.eof then
								pcv_tmpArr1=rstemp.getRows()
								intCount1=ubound(pcv_tmpArr1,2)
								set rstemp=nothing
								For m=0 to intCount1
									query="INSERT INTO pcBRCatCanNot (pcBR_ID,pcBRCatCanNot_Item) VALUES (" & pcv_NewBRID & "," & pcv_tmpArr1(0,m) & ");"
									set rstemp=connTemp.execute(query)
									set rstemp=nothing
								Next
							end if
						end if
					END IF
					Next
					Next
				end if
				set rstemp=nothing
			
			End If
			
			msg="The settings were successfully applied to the selected configurable products."
			msgType=1
		ELSE 'Does not have configuration
			msg="This configurable product has not been setup yet."
		END IF
	else 'Did not select Source Configurable Product
		msg="Please select an existing configurable product."
	end if
end if 'End of Update Action
%>
<!--#include file="AdminHeader.asp"-->
<%if msg<>"" then%>
<table class="pcCPcontent">
<tr>
	<td>
		<% ' START show message, if any %>
            <!--#include file="pcv4_showMessage.asp"-->
        <% 	' END show message %>
	</td>
</tr>
	<tr>
		<td class="pcSpacer">&nbsp;</td>
	</tr>
	<tr>
		<td>
			<ul>
				<li><a href="LocateProducts.asp?cptype=1">Locate a configurable product</a></li>
				<li><a href="menu.asp">Back to the start page</a></li>
			</ul>
		</td>
	</tr>
<table class="pcCPcontent">
<%end if%>
<table class="pcCPcontent" <%if msgType=1 then%>style="display:none"<%end if%>>
<tr>
	<td>
		<table id="FindProducts" class="pcCPcontent">
		<tr>
			<td>
			<%
				src_FormTitle1="Find Products"
				src_FormTitle2="Copy Product Configuration to other Configurable Products"
				src_FormTips1="Use the following filters to look for configurable products in your store that you would like to apply settings to."
				src_FormTips2="Select one or more configurable products that you would like to apply settings to."
				src_IncNormal=0
				src_IncBTO=1
				src_IncItem=0
				src_DisplayType=1
				src_ShowLinks=0
				src_FromPage="ApplyBTOConfigMulti.asp?idproduct=" & pidproduct
				src_ToPage="ApplyBTOConfigMulti.asp?action=update&idproduct=" & pidproduct
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