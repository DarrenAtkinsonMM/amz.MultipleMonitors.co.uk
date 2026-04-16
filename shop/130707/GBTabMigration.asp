<%@LANGUAGE="VBSCRIPT"%>
<% 'On Error Resume Next %>
<% pageTitle = "Upgrade GB Tabbed Product to ProductCart Customized Layout Feature" %>
<% Section = "" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../pc/gbdg_ProductTabsConfig.asp"-->
<%
Dim i, j, k

call opendb()

'Global Settings
Dim pcv_TabNames(6)
pcv_TabNames(0)="Tab1"
pcv_TabNames(1)="Tab2"
pcv_TabNames(2)="Tab3"
pcv_TabNames(3)="Tab4"
pcv_TabNames(4)="Tab5"
pcv_TabNames(5)="Tab6"

'scViewPrdStyle = "C"

C_Top="PrdName,CatTree,"
C_TopLeft="PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdInput,PrdATC,PrdWL,"
C_TopRight="PrdAT,PrdImg,PrdQDisc,"
C_Bottom="PrdBtns,PrdCS,PrdLDesc,PrdRev,"

L_Top="PrdName,CatTree,"
L_TopLeft="PrdAT,PrdImg,PrdQDisc,"
L_TopRight="PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdInput,PrdATC,PrdWL,"
L_Bottom="PrdBtns,PrdCS,PrdLDesc,PrdRev,"

O_Top="PrdName,CatTree,PrdImg,PrdAT,PrdQDisc,PrdSKU,PrdRate,PrdW,PrdBrand,PrdStock,PrdDesc,PrdConfig,PrdSearch,PrdRP,PrdPrice,PrdSB,PrdPromo,PrdNoShip,PrdOSM,PrdBOM,PrdOpt,PrdInput,PrdATC,PrdWL,"
O_TopLeft=""
O_TopRight=""
O_Bottom="PrdBtns,PrdCS,PrdLDesc,PrdRev,"

query="SELECT idProduct,pcProd_TabbedContent1,pcProd_TabbedContent2,pcProd_TabbedContent3,pcProd_TabbedContent4,pcProd_TabbedContent5,pcProd_TabbedContent6,pcProd_TabContentTitle1,pcProd_TabContentTitle2,pcProd_TabContentTitle3,pcProd_TabContentTitle4,pcProd_TabContentTitle5,pcProd_TabContentTitle6,pcprod_DisplayLayout FROM Products WHERE removed=0;"
set rs=connTemp.execute(query)
HaveRecords=0
if not rs.eof then
	pcArr=rs.getRows()
	set rs=nothing
	intCount=ubound(pcArr,2)
	HaveRecords=1
end if
set rs=nothing

Dim tmpTabs(6),cat_TabNames(6),LastTab


IF HaveRecords=1 THEN
	'Check Cat Level
	On error resume next
	HaveCatLevel=0
	err.number=0
	err.description=""
	query="SELECT TOP 1 gbdgCatTabText1 FROM Categories;"
	set rs=connTemp.execute(query)
	if err.number<>0 then
		HaveCatLevel=0
		err.number=0
		err.description=""
	else
		HaveCatLevel=1
	end if
	set rs=nothing


	For i=0 to intCount
		P_Top=""
		P_TopLeft=""
		P_TopRight=""
		P_Tabs=""
		P_Bottom=""
		
		tmpTabs(0)=""
		tmpTabs(1)=""
		tmpTabs(2)=""
		tmpTabs(3)=""
		tmpTabs(4)=""
		tmpTabs(5)=""
		
		cat_TabNames(0)=""
		cat_TabNames(1)=""
		cat_TabNames(2)=""
		cat_TabNames(3)=""
		cat_TabNames(4)=""
		cat_TabNames(5)=""
		
		LastTab=-1
		tabcount=-1
		
		pidProduct=pcArr(0,i)
		pLayout=pcArr(13,i)
		if IsNull(pLayout) or pLayout="" then
			pLayout=scViewPrdStyle
		end if
		
		Select Case uCase(pLayout)
			Case "C":
				P_Top=C_Top
				P_TopLeft=C_TopLeft
				P_TopRight=C_TopRight
				P_Bottom=C_Bottom
			Case "L":
				P_Top=L_Top
				P_TopLeft=L_TopLeft
				P_TopRight=L_TopRight
				P_Bottom=L_Bottom
			Case "O":
				P_Top=O_Top
				P_TopLeft=O_TopLeft
				P_TopRight=O_TopRight
				P_Bottom=O_Bottom
			Case Else:
				P_Top=C_Top
				P_TopLeft=C_TopLeft
				P_TopRight=C_TopRight
				P_Bottom=C_Bottom
		End Select
		
		'Get Tab Names - Cat Level
		if HaveCatLevel=1 then
			query="SELECT gbdgCatTabText1,gbdgCatTabText2,gbdgCatTabText3,gbdgCatTabText4,gbdgCatTabText5,gbdgCatTabText6 FROM Categories INNER JOIN Categories_Products ON Categories.idCategory=Categories_Products.idCategory WHERE Categories_Products.idProduct=" & pIdProduct & ";"
			set rs=connTemp.execute(query)
			if not rs.eof then
				cat_TabNames(0)=replace(rs("gbdgCatTabText1"),"'","''")
				cat_TabNames(1)=replace(rs("gbdgCatTabText2"),"'","''")
				cat_TabNames(2)=replace(rs("gbdgCatTabText3"),"'","''")
				cat_TabNames(3)=replace(rs("gbdgCatTabText4"),"'","''")
				cat_TabNames(4)=replace(rs("gbdgCatTabText5"),"'","''")
				cat_TabNames(5)=replace(rs("gbdgCatTabText6"),"'","''")
			end if
			set rs=nothing
		end if
		
		tmpGBTabs=trim(pcArr(1,i) & pcArr(2,i) & pcArr(3,i) & pcArr(4,i) & pcArr(5,i) & pcArr(6,i))
		
		IF IsNull(tmpGBTabs) OR (tmpGBTabs="") THEN
		
			'Override Tab 1 with Product SDesc or LDesc
			if ((gbdgSetSdescAsTab1="1") OR (gbdgSetDetailAsTab1="1")) then
				tabcount=0
				LastTab=1
				if gbdgSetSdescAsTab1="1" then
					tmpTabs(0)="Description``PrdDesc``||"
					P_Top=replace(P_Top,"PrdDesc,","")
					P_TopLeft=replace(P_TopLeft,"PrdDesc,","")
					P_TopRight=replace(P_TopRight,"PrdDesc,","")
					P_Bottom=replace(P_Bottom,"PrdDesc,","")
				else
					tmpTabs(0)="Description``PrdLDesc``||"
					P_Top=replace(P_Top,"PrdLDesc,","")
					P_TopLeft=replace(P_TopLeft,"PrdLDesc,","")
					P_TopRight=replace(P_TopRight,"PrdLDesc,","")
					P_Bottom=replace(P_Bottom,"PrdLDesc,","")
				end if
			end if
			
			'Override Tab J with Cross Selling
			if (gbdgCrossSellInTabs="1") then
				tabcount=tabcount+1
				tmpTabs(tabcount)="Cross Selling``PrdCS``||"
				LastTab=tabcount+1
				P_Top=replace(P_Top,"PrdCS,","")
				P_TopLeft=replace(P_TopLeft,"PrdCS,","")
				P_TopRight=replace(P_TopRight,"PrdCS,","")
				P_Bottom=replace(P_Bottom,"PrdCS,","")
			end if
			
			'Override Last Tab with Product Reviews
			if gbdgReviewsInTabs = "1" then
				if LastTab<6 then
					LastTab=LastTab+1
					tabcount=tabcount+1
				end if
				tmpTabs(LastTab-1)="Reviews``PrdRev``||"
				P_Top=replace(P_Top,"PrdRev,","")
				P_TopLeft=replace(P_TopLeft,"PrdRev,","")
				P_TopRight=replace(P_TopRight,"PrdRev,","")
				P_Bottom=replace(P_Bottom,"PrdRev,","")
			end if
			
		ELSE
		
		For j=1 to 6
			'Override Tab J with Cross Selling
			if (gbdgCrossSellInTabs="1") AND (Clng(gbdgCrossSellColumn)=Clng(j)) then
				tabcount=tabcount+1
				tmpTabs(j-1)="Cross Selling``PrdCS``||"
				LastTab=j
				P_Top=replace(P_Top,"PrdCS,","")
				P_TopLeft=replace(P_TopLeft,"PrdCS,","")
				P_TopRight=replace(P_TopRight,"PrdCS,","")
				P_Bottom=replace(P_Bottom,"PrdCS,","")
			else
				if pcArr(j,i)<>"" then
					tabcount=tabcount+1
					tmpTabName=replace(pcArr(j+6,i),"'","''")
					'Override Tab Name with Cat Level
					if IsNull(tmpTabName) OR tmpTabName="" then
						if HaveCatLevel=1 then
							tmpTabName=cat_TabNames(j-1)
						end if
					end if
					
					'Override Tab Name with Store Level
					if IsNull(tmpTabName) OR tmpTabName="" then
						Select Case j
							Case 1: tmpTabName=gbdgTabTitle1
							Case 2: tmpTabName=gbdgTabTitle2
							Case 3: tmpTabName=gbdgTabTitle3
							Case 4: tmpTabName=gbdgTabTitle4
							Case 5: tmpTabName=gbdgTabTitle5
							Case 6: tmpTabName=gbdgTabTitle6
						End Select
					end if
					
					'Override Tab Name with new ProductCart Default Tab Names
					if IsNull(tmpTabName) OR tmpTabName="" then
						tmpTabName=pcv_TabNames(tabcount)
					end if
					
					tmpTabs(j-1)=tmpTabName & "``CUSTOMHTML,``" & replace(pcArr(j,i),"'","''") & "||"

					LastTab=j
				end if
			end if
			'Override Tab 1 with Product SDesc or LDesc
			if (j=1) AND ((gbdgSetSdescAsTab1="1") OR (gbdgSetDetailAsTab1="1")) then
				tabcount=0
				LastTab=j
				if gbdgSetSdescAsTab1="1" then
					tmpTabs(j-1)="Description``PrdDesc,``||"
					P_Top=replace(P_Top,"PrdDesc,","")
					P_TopLeft=replace(P_TopLeft,"PrdDesc,","")
					P_TopRight=replace(P_TopRight,"PrdDesc,","")
					P_Bottom=replace(P_Bottom,"PrdDesc,","")
				else
					tmpTabs(j-1)="Description``PrdLDesc,``||"
					P_Top=replace(P_Top,"PrdLDesc,","")
					P_TopLeft=replace(P_TopLeft,"PrdLDesc,","")
					P_TopRight=replace(P_TopRight,"PrdLDesc,","")
					P_Bottom=replace(P_Bottom,"PrdLDesc,","")
				end if
			end if
		Next
		
		'Override Last Tab with Product Reviews
		if gbdgReviewsInTabs = "1" then
			if LastTab<6 then
				LastTab=LastTab+1
				tabcount=tabcount+1
			end if
			tmpTabs(LastTab-1)="Reviews``PrdRev``||"
			P_Top=replace(P_Top,"PrdRev,","")
			P_TopLeft=replace(P_TopLeft,"PrdRev,","")
			P_TopRight=replace(P_TopRight,"PrdRev,","")
			P_Bottom=replace(P_Bottom,"PrdRev,","")
		end if
		
		END IF
		
		For j=1 to 6
			if tmpTabs(j-1)<>"" then
				P_Tabs=P_Tabs & tmpTabs(j-1)
			end if
		Next
		
		if P_Tabs<>"" then
			query="UPDATE Products SET pcprod_DisplayLayout='t',pcProd_Top='" & P_Top & "',pcProd_TopLeft='" & P_TopLeft & "',pcProd_TopRight='" & P_TopRight & "',pcProd_Tabs=N'" & P_Tabs & "',pcProd_Bottom='" & P_Bottom & "' WHERE idProduct=" & pIdProduct & ";"
			set rs=connTemp.execute(query)
			set rs=nothing
		end if
	Next

END IF

call closedb()
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td>
		<%IF HaveRecords=1 THEN%>
		<div class="pcCPmessageSuccess">
			Upgraded GB Tabbed Product to ProductCart Customized Layout Feature successfully!
		</div>
		<%ELSE%>
		<div class="pcCPmessageInfo">
			GB Tabbed Products not found!
		</div>
		<%END IF%>
	</td>
</tr>
</table>
<!--#include file="AdminFooter.asp"-->