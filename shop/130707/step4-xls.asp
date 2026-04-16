<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<%
on error resume next
Server.ScriptTimeout = 5400

dim f, rstemp1,TopRecord(100), IDcustom(2), Customcontent(2)
Dim IDPriceCat(4)
Dim PrdWithoutOpts, CheckCount
Dim ErrorsReport

PrdWithoutOpts=0
CheckCount=0

ErrorsReport=""
TempProducts=""
TempProducts=session("TempProducts") 
ErrorsReport=session("ErrorsReport")

if session("PrdWithoutOpts")="" then
	PrdWithoutOpts=0
else
	PrdWithoutOpts=session("PrdWithoutOpts")
end if

Function ImportPrdOptions(IDPrd,tmp_Opt1,tmp_Attr1,tmp_Opt1Req,tmp_Opt1Order)
	If tmp_Opt1<>"" then
		pcv_IDGrp1=checkOptGrp(tmp_Opt1)
		Call checkPrdGrp(IDPrd,pcv_IDGrp1,tmp_Opt1Req,tmp_Opt1Order)
		IF tmp_Attr1<>"" then
			pcv_Arr1=split(tmp_Attr1,"**")
			testErr=0
			For i=lbound(pcv_Arr1) to ubound(pcv_Arr1)
				IF pcv_Arr1(i)<>"" THEN
					pcv_Arr2=split(pcv_Arr1(i),"*")
					if ubound(pcv_Arr2)>4 then
						testErr=1
					else
						if pcv_Arr2(0)="" then
							testErr=1
						else
							'Replace Quotes
							rd_Option1=replace(trim(pcv_Arr2(0)),"""","&quot;")
						end if
						if ubound(pcv_Arr2)>=1 then
							if pcv_Arr2(1)="" then
								rd_price=0
							else
								rd_price=trim(pcv_Arr2(1))
							end if
							if IsNumeric(rd_price)=false then
								testErr=1
							end if
						else
							rd_price=0
						end if
						
						if ubound(pcv_Arr2)>=2 then
							if pcv_Arr2(2)="" then
								rd_wprice=0
							else
								rd_wprice=trim(pcv_Arr2(2))
							end if
							if IsNumeric(rd_wprice)=false then
								testErr=1
							end if
						else
							rd_wprice=0
						end if
						
						if ubound(pcv_Arr2)>=3 then
							if pcv_Arr2(3)="" then
								rd_order=0
							else
								rd_order=trim(pcv_Arr2(3))
							end if
							if IsNumeric(rd_order)=false then
								testErr=1
							end if
						else
							rd_order=0
						end if
						
						if ubound(pcv_Arr2)=4 then
							if pcv_Arr2(4)="" then
								rd_active=1
							else
								rd_active=trim(pcv_Arr2(4))
							end if
							if IsNumeric(rd_active)=false then
								testErr=1
							end if
						else
							rd_active=1
						end if
						if rd_active<>"0" then
							rd_inactive="0"
						else
							rd_inactive="1"
						end if
					end if
					
					IF testErr=1 THEN
						if CheckCount=0 then
							ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - attributes list format is incorrect." & "</td></tr>" & vbcrlf
							PrdWithoutOpts=PrdWithoutOpts+1
							CheckCount=1
						end if
						exit for
					END IF
				END IF 'pcv_Arr1(i)<>""

			Next
			
			IF testErr=0 THEN
			
				testErr=0
				For i=lbound(pcv_Arr1) to ubound(pcv_Arr1)
					IF pcv_Arr1(i)<>"" THEN
						pcv_Arr2=split(pcv_Arr1(i),"*")
						if ubound(pcv_Arr2)>4 then
							testErr=1
						else
							if pcv_Arr2(0)="" then
								testErr=1
							else
								'Replace Quotes
								rd_Option1=replace(trim(pcv_Arr2(0)),"""","&quot;")
								if len(rd_Option1)>250 then
									ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": The option attribute name '" & rd_Option1 & "' was longer than 250 characters. It has been truncated." & "</td></tr>" & vbcrlf
									rd_Option1=mid(rd_Option1,1,250)
								end if
							end if
							if ubound(pcv_Arr2)>=1 then
								if pcv_Arr2(1)="" then
									rd_price=0
								else
									rd_price=trim(pcv_Arr2(1))
								end if
								if IsNumeric(rd_price)=false then
									testErr=1
								end if
							else
								rd_price=0
							end if
							
							if ubound(pcv_Arr2)>=2 then
								if pcv_Arr2(2)="" then
									rd_wprice=0
								else
									rd_wprice=trim(pcv_Arr2(2))
								end if
								if IsNumeric(rd_wprice)=false then
									testErr=1
								end if
							else
								rd_wprice=0
							end if
							
							if ubound(pcv_Arr2)>=3 then
								if pcv_Arr2(3)="" then
									rd_order=0
								else
									rd_order=trim(pcv_Arr2(3))
								end if
								if IsNumeric(rd_order)=false then
									testErr=1
								end if
							else
								rd_order=0
							end if
							
							if ubound(pcv_Arr2)=4 then
								if pcv_Arr2(4)="" then
									rd_active=1
								else
									rd_active=trim(pcv_Arr2(4))
								end if
								if IsNumeric(rd_active)=false then
									testErr=1
								end if
							else
								rd_active=1
							end if
							if rd_active<>"0" then
								rd_inactive="0"
							else
								rd_inactive="1"
							end if
						end if
						
						IF testErr=1 THEN
							if CheckCount=0 then
								ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - attributes list format is incorrect." & "</td></tr>" & vbcrlf
								PrdWithoutOpts=PrdWithoutOpts+1
								CheckCount=1
							end if
							exit for
						ELSE
							pcv_IDOpt1=checkAttr(pcv_IDGrp1,rd_Option1)
							Call ImUpOptGrp(IDPrd,pcv_IDGrp1,pcv_IDOpt1,rd_price,rd_wprice,rd_order,rd_inactive)
						END IF
					END IF
				Next
			
			END IF 'testErr=0
		END IF
		if pcv_IDGrp1>"0" then
			query="SELECT idoptoptgrp FROM options_optionsGroups WHERE IDProduct=" & IDPrd & " AND idOptionGroup=" & pcv_IDGrp1 & ";"
			set rstemp=connTemp.execute(query)
			if rstemp.eof then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because it does not have any imported attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
				query="DELETE FROM pcProductsOptions WHERE idProduct=" & IDPrd & " AND idOptionGroup=" & pcv_IDGrp1 & ";"
				set rstemp=conntemp.execute(query)
				set rstemp=nothing
			end if
			set rstemp=nothing
		end if
	End if
End Function

function checkOptGrp(GrpName)
	query="SELECT idOptionGroup FROM optionsGroups WHERE OptionGroupDesc='" & GrpName & "'"
	set rstemp=conntemp.execute(query)	
	if rstemp.eof then
		query="insert into optionsGroups (OptionGroupDesc) values (N'" & GrpName & "')"
		set rstemp=conntemp.execute(query)
		query="SELECT idOptionGroup FROM optionsGroups WHERE OptionGroupDesc='" & GrpName & "'"
		set rstemp=conntemp.execute(query)
		checkOptGrp=rstemp("idOptionGroup")
	else
		checkOptGrp=rstemp("idOptionGroup")
	end if
	set rstemp=nothing
end function

Function checkAttr(IDGrp,AttrName)
	Dim IDOption
	query="SELECT options.idOption FROM options INNER JOIN optGrps ON options.idOption=optGrps.idoption WHERE optGrps.idOptionGroup=" & IDGrp & " AND options.optionDescrip='" & AttrName & "';"
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		query="insert into options (optionDescrip) values (N'" & AttrName & "')"
		set rstemp=conntemp.execute(query)
		query="SELECT idOption FROM options WHERE optionDescrip='" & AttrName & "' ORDER BY idOption DESC;"
		set rstemp=conntemp.execute(query)
		IDOption=rstemp("idOption")
	else
		IDOption=rstemp("idOption")
	end if

	query="SELECT idoption FROM optGrps WHERE idoption=" & IDOption & " AND idOptionGroup=" & IDGrp
	set rstemp=connTemp.execute(query)
	if rstemp.eof then
		query="insert into optGrps (idOptionGroup,idoption) values (" & IDGrp & "," & IDOption & ")"
		set rstemp=conntemp.execute(query)
	end if
	checkAttr=IDOption
	set rstemp=nothing
end function

Sub checkPrdGrp(IDPrd,IDGrp,GrpReq,GrpOrder)
	query="SELECT idOptionGroup FROM pcProductsOptions WHERE idProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & ";"
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		query="INSERT INTO pcProductsOptions (idProduct,idOptionGroup,pcProdOpt_Required,pcProdOpt_order) VALUES (" & IDPrd & "," & IDGrp & "," & GrpReq & "," & GrpOrder & ");"
		set rstemp=conntemp.execute(query)	
	else
		query="UPDATE pcProductsOptions SET idProduct=" & IDPrd & ",idOptionGroup=" & IDGrp & ",pcProdOpt_Required=" & GrpReq & ",pcProdOpt_order=" & GrpOrder & " WHERE idProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & ";"
		set rstemp=conntemp.execute(query)	
	end if
	set rstemp=nothing
End Sub

Sub ImUpOptGrp(IDPrd,IDGrp,IDOpt,OptPrice,OptWPrice,DOrder,InActive)
	query="SELECT idoptoptgrp FROM options_optionsGroups WHERE IDProduct=" & IDPrd & " AND idOptionGroup=" & IDGrp & " AND idOption=" & IDOpt & ";"
	set rstemp=connTemp.execute(query)	
	if rstemp.eof then
		query="INSERT INTO options_optionsGroups (IDProduct,idOptionGroup,idOption,price,Wprice,sortOrder,InActive) VALUES (" & IDPrd & "," & IDGrp & "," & IDOpt & "," & OptPrice & "," & OptWPrice & "," & DOrder & "," & InActive & ")"
		set rstemp=connTemp.execute(query)
	else
		query="UPDATE options_optionsGroups SET price=" & OptPrice & ",Wprice=" & OptWprice & ",sortOrder=" & DOrder & ",InActive=" & InActive & " WHERE idoptoptgrp=" & rstemp("idoptoptgrp")
		set rstemp=connTemp.execute(query)
	end if	
	set rstemp=nothing
End Sub

function checkparent(pname)
	Dim mypname,mypname1
	Dim pIDCategory

	mypname=pname
	mypname=replace(mypname,"&amp;","&")
	mypname=replace(mypname,"&","&amp;")
	
	mypname1=replace(pname,"&amp;","&")

	query="Select idCategory from categories where (categorydesc='" & mypname & "' or categorydesc='" & mypname1 & "')"
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		imagename="no_image.gif"
		query="insert into categories (categorydesc,idParentCategory,image,largeimage) values (N'" & mypname & "',1,'" & imagename & "','" & imagename & "')"
		set rstemp1=conntemp.execute(query)
		query="Select idCategory from categories where categorydesc='" & mypname & "'"
		set rstemp1=conntemp.execute(query)
		pIDCategory=rstemp1("idCategory")
		
		call updCatCreatedDate(pIDCategory,"")
		
		checkparent=pIDCategory
	else
		checkparent=rstemp("idCategory")
	end if
end function

function new_checkparent(subname)
	Dim mysubname,mysubname1
	Dim pIDCategory,query,rstemp
	mysubname=subname
	mysubname=replace(mysubname,"&amp;","&")
	mysubname=replace(mysubname,"&","&amp;")	
	mysubname1=replace(subname,"&amp;","&")
	query="SELECT idParentCategory FROM categories WHERE (categorydesc LIKE '" & mysubname & "' OR categorydesc LIKE '" & mysubname1 & "');"
	set rstemp=conntemp.execute(query)	
	if not rstemp.eof then
		pIDCategory=rstemp("idParentCategory")
		new_checkparent=pIDCategory
	else
		new_checkparent=1
	end if
	set rstemp=nothing
end function

function checkbrand(pname,pimg)
	query="SELECT idBrand FROM Brands WHERE BrandName LIKE '" & pname & "'"
	set rstemp=conntemp.execute(query)	
	if rstemp.eof then
		if pimg="" then
			bimage="no_image.gif"
		else
			bimage=pimg
		end if
		query="INSERT INTO Brands (BrandName,BrandLogo) VALUES (N'" & pname & "','" & bimage & "')"
		set rstemp1=conntemp.execute(query)
		set rstemp1=nothing
		query="SELECT idBrand FROM Brands WHERE BrandName LIKE '" & pname & "'"
		set rstemp1=conntemp.execute(query)
		if not rstemp1.eof then
			checkbrand=rstemp1("IDBrand")
		else
			checkbrand=0
		end if
		set rstemp1=nothing
	else
		checkbrand=rstemp("IDBrand")
		set rstemp=nothing
		if pimg<>"" then
			bimage=pimg
			query="UPDATE Brands SET BrandLogo='" & bimage & "' WHERE BrandName LIKE '" & pname & "'"
			set rstemp1=conntemp.execute(query)
			set rstemp1=nothing
		end if
	end if
end function

function checkPCat(pcatname)
	Dim query,rstemp
	query="SELECT idCustomerCategory FROM pcCustomerCategories WHERE pcCC_Name like '" & pcatname & "'"
	set rstemp=conntemp.execute(query)
	if not rstemp.eof then
		checkPCat=rstemp("idCustomerCategory")
	else
		checkPCat=-1
	end if
	set rstemp=nothing
end function

function checkcustom(cfname)
	query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & cfname & "'"
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		query="INSERT INTO pcSearchFields (pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder,pcSearchFieldCPShow,pcSearchFieldSearch,pcSearchFieldCPSearch) VALUES (N'" & cfname & "',1,0,1,1,1)"
		set rstemp1=conntemp.execute(query)
		query="SELECT idSearchField FROM pcSearchFields WHERE pcSearchFieldName like '" & cfname & "'"
		set rstemp1=conntemp.execute(query)
		checkcustom=rstemp1("idSearchField")
	else
		checkcustom=rstemp("idSearchField")
	end if
	set rstemp=nothing
	set rstemp1=nothing
end function

function checkcustomvalue(idcustom,searchvalue)
	query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & searchvalue & "'"
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & idcustom & ",N'" & searchvalue & "',0)"
		set rstemp1=conntemp.execute(query)
		query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & idcustom & " AND pcSearchDataName like '" & searchvalue & "'"
		set rstemp1=conntemp.execute(query)
		checkcustomvalue=rstemp1("idSearchData")
	else
		checkcustomvalue=rstemp("idSearchData")
	end if
	set rstemp=nothing
	set rstemp1=nothing
end function

function checkcategory(cname,pcid,simage,limage,SDesc1,LDesc1)

	Dim mycname,mycname1
	
	mycname=cname
	mycname=replace(mycname,"&amp;","&")
	mycname=replace(mycname,"&","&amp;")	
	mycname1=replace(cname,"&amp;","&")
	
	query="Select idCategory from categories where (categorydesc='" & mycname & "' or categorydesc='" & mycname1 & "') and idParentCategory=" & pcid 
	set rstemp=conntemp.execute(query)
	if rstemp.eof then
		query1="categoryDesc,idParentCategory,image,largeimage,SDesc,LDesc"
		if simage<>"" then
			smallimg=simage
		else
			smallimg="no_image.gif"
		end if
		if limage<>"" then
			largeimg=limage
		else
			largeimg="no_image.gif"
		end if
		query2="N'" & mycname & "'," & pcid & ",'" & smallimg & "','" & largeimg & "',N'" & SDesc1 & "',N'" & LDesc1 & "'"
		query="insert into categories (" & query1 & ") values (" & query2 & ")"
		set rstemp1=conntemp.execute(query)
		query="Select idCategory from categories where categorydesc='" & mycname & "' and idParentCategory=" & pcid
		set rstemp1=conntemp.execute(query)
		tcheckcategory=rstemp1("idCategory")
		
		call updCatCreatedDate(tcheckcategory,"")
		
	else
		tcheckcategory=rstemp("idCategory")
		query1=""
		if mycname<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "categoryDesc=N'" & mycname & "'"
		end if
		if pcid<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "idParentCategory=" & pcid
		end if
		if simage<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "image='" & simage & "'"
		end if
		if limage<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "largeimage='" & limage & "'"
		end if
		if SDesc1<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "SDesc=N'" & SDesc1 & "'"
		end if
		if LDesc1<>"" then
			if query1<>"" then
				query1=query1 & ","
			end if
			query1=query1 & "LDesc=N'" & LDesc1 & "'"
		end if
		if query1<>"" then
			query="UPDATE categories SET " & query1 & " WHERE idcategory=" & tcheckcategory & ";"
			set rstemp1=connTemp.execute(query)
			set rstemp1=nothing
			
			call updCatEditedDate(tcheckcategory,"")
		end if
	end if
	set rstemp=nothing
	set rstemp1=nothing

	checkcategory=tcheckcategory

end function
	
function checktempcategory()

	Dim pIDCategory

	TempCategory="ImportedProducts"
	
	query="Select idCategory from categories where categorydesc='" & TempCategory & "' and idParentCategory=1"
	set rstemp=conntemp.execute(query)
	
	if rstemp.eof then
		imagename="no_image.gif"
		query="insert into categories (categorydesc,idParentCategory,image,largeimage) values (N'" & TempCategory & "',1,'" & imagename & "','" & imagename & "')"
		set rstemp1=conntemp.execute(query)
		query="Select idCategory from categories where categorydesc='" & TempCategory & "' and idParentCategory=1"
		set rstemp1=conntemp.execute(query)
		pIDCategory=rstemp1("idCategory")
		checktempcategory=pIDCategory
		
		call updCatCreatedDate(pIDCategory,"")
		
	else
		checktempcategory=rstemp("idCategory")
	end if
end function
%>

<!--#include file="common.asp"-->

<%
iPageSize=3000
iPageCurrent=session("iPageCurrent") 
if iPagecurrent="" then 
	iPageCurrent=1 
end if 

Append=session("append")
if PPD="1" then
	FileXLS = "/"&scPcFolder&"/pc/catalog/" & session("importfile")
else
	FileXLS = "../pc/catalog/" & session("importfile")
end if

Set cnnExcel = Server.CreateObject("ADODB.Connection")
cnnExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(FileXLS) & ";Extended Properties=Excel 8.0;"
Set rsExcel = Server.CreateObject("ADODB.Recordset")
	
rsExcel.CacheSize=iPageSize 
rsExcel.PageSize=iPageSize  

'/*rsExcel.open "SELECT * FROM IMPORT;", cnnExcel
'/*Altered by Sheri
rsExcel.open "SELECT * FROM IMPORT;", cnnExcel , adOpenStatic, adLockReadOnly, adCmdText

dim iPageCount 
iPageCount=rsExcel.PageCount 

If Cint(iPageCurrent) > Cint(iPageCount) Then iPageCurrent=Cint(iPageCount) 
If Cint(iPageCurrent) < 1 Then iPageCurrent=Cint(1) 
'rsExcel.MoveFirst 
rsExcel.AbsolutePage=iPageCurrent 

' counting variable for our recordset 
dim count 

	
if Err.number<>0 then
	session("importfilename")=""
	call closeDb()
response.redirect "msg.asp?message=30"
end if

TotalXLSlines=session("TotalXLSlines")
if TotalXLSlines="" then
	TotalXLSlines=0
end if
ImportedRecords=session("ImportedRecords")
if ImportedRecords="" then
	ImportedRecords=0
end if
fields=session("totalfields")
iCols = rsExcel.Fields.Count
if (customfieldsid(0)>-1) or (customfieldsid(1)>-1) or (customfieldsid(2)>-1) then
	if customfieldsid(0) > -1 then
		customfieldsname(0)=rsExcel.Fields.Item(int(customfieldsid(0))).Name
		if customfieldsname(0)<>"" then
			customfieldsname(0)=replace(customfieldsname(0),"'","''")
		end if
	end if
	if customfieldsid(1) > -1 then
		customfieldsname(1)=rsExcel.Fields.Item(int(customfieldsid(1))).Name
		if customfieldsname(1)<>"" then
			customfieldsname(1)=replace(customfieldsname(1),"'","''")
		end if
	end if
	if customfieldsid(2) > -1 then
		customfieldsname(2)=rsExcel.Fields.Item(int(customfieldsid(2))).Name
		if customfieldsname(2)<>"" then
			customfieldsname(2)=replace(customfieldsname(2),"'","''")
		end if
	end if	
end if

if (pricecatsid(0)>-1) OR (pricecatsid(1)>-1) OR (pricecatsid(2)>-1) OR (pricecatsid(3)>-1) OR (pricecatsid(4)>-1) then
	if pricecatsid(0) > -1 then
		pricecatsname(0)=rsExcel.Fields.Item(int(pricecatsid(0))).Name
		if pricecatsname(0)<>"" then
			pricecatsname(0)=replace(pricecatsname(0),"'","''")
		end if
	end if
	
	if pricecatsid(1) > -1 then
		pricecatsname(1)=rsExcel.Fields.Item(int(pricecatsid(1))).Name
		if pricecatsname(1)<>"" then
			pricecatsname(1)=replace(pricecatsname(1),"'","''")
		end if
	end if
	
	if pricecatsid(2) > -1 then
		pricecatsname(2)=rsExcel.Fields.Item(int(pricecatsid(2))).Name
		if pricecatsname(2)<>"" then
			pricecatsname(2)=replace(pricecatsname(2),"'","''")
		end if
	end if
	
	if pricecatsid(3) > -1 then
		pricecatsname(3)=rsExcel.Fields.Item(int(pricecatsid(3))).Name
		if pricecatsname(3)<>"" then
			pricecatsname(3)=replace(pricecatsname(3),"'","''")
		end if
	end if
	
	if pricecatsid(4) > -1 then
		pricecatsname(4)=rsExcel.Fields.Item(int(pricecatsid(4))).Name
		if pricecatsname(4)<>"" then
			pricecatsname(4)=replace(pricecatsname(4),"'","''")
		end if
	end if
end if
	
if rsExcel.EOF then
	session("importfilename")=""
	call closeDb()
response.redirect "msg.asp?message=32"
end if

	

'Get previous information before import/update products
query="Select * from products order by IDproduct desc"
set rstemp4=connTemp.execute(query)

if not rstemp4.eof then
	PreIDProduct="" & rstemp4("IDproduct")
else
	PreIDProduct="0"
end if
	
query="Select * from categories order by IDCategory desc"
set rstemp4=connTemp.execute(query)

if not rstemp4.eof then
	PreIDCategory="" & rstemp4("IDcategory")
else
	PreIDCategory="0"
end if
	
query="Select * from brands order by IDBrand desc"
set rstemp4=connTemp.execute(query)

if not rstemp4.eof then
	PreIDBrand="" & rstemp4("IDBrand")
else
	PreIDBrand="0"
end if
	
if session("append")="1" then
	UpdateType="UPDATE"
else
	UpdateType="IMPORT"
end if
PreRecords=""
CATRecords=""
	
SKUError=0
' set count equal to zero 
count=0 

do while not rsExcel.eof and count < rsExcel.pageSize '/*Altered by Sheri
'/*Do While not rsExcel.EOF
	
	RecordError=false
	TotalXLSlines=TotalXLSlines+1
	
	if RecordError=False then%>
		<!--#include file="common2.asp"-->
	<%end if%>
		
<%
if RecordError=false then
	'Start SDBA
	if prd_SupplierID>-1 then
		if prd_Supplier>0 then
			query="SELECT pcSupplier_IsDropShipper FROM pcSuppliers WHERE pcSupplier_ID=" & prd_Supplier
			set rsQ=conntemp.execute(query)
			if not rsQ.eof then
				prd_IsDropShipper=rsQ("pcSupplier_IsDropShipper")
				if prd_IsDropShipper="1" then
					prd_DropShipper=prd_Supplier
				end if
			end if
			set rsQ=nothing
		end if
	end if
	'End SDBA
	

	For m=0 to 2
		if customfieldsname(m)<>"" then
			if NOT (len(Session("IDcustom"&m))>0) then
				IDcustom(m)=checkcustom(customfieldsname(m))
				Session("IDcustom"&m)=IDcustom(m)	
			else
				IDcustom(m)=Session("IDcustom"&m)
			end if			
			if customfields(m)<>"" then
				Customcontent(m)=checkcustomvalue(IDcustom(m),customfields(m))					
			else
				Customcontent(m)=0
			end if
		else
			IDcustom(m)=-1
			Customcontent(m)=""
		end if
	Next
	
	For m=0 to 4
		if pricecatsname(m)<>"" then
			if NOT (len(Session("IDPriceCat"&m))>0) then
				IDPriceCat(m)=checkPCat(pricecatsname(m))
				if IDPriceCat(m)<>-1 then
					Session("IDPriceCat"&m)=IDPriceCat(m)
				else
					Session("IDPriceCat"&m)=""
				end if
			else
				IDPriceCat(m)=Session("IDPriceCat"&m)
			end if			
		else
			IDPriceCat(m)=-1
		end if
	Next

	
	temp1=""
	temp2=""
	temp3=""
	
	IsDownloadable=0
	
	if pptype<>"" then
		if ucase(pptype)="BTO" then
			temp1=temp1 & ",serviceSpec"
			temp2=temp2 & ",-1"
			temp3=temp3 & ",serviceSpec=-1"
		else
			if ucase(pptype)="ITEM" then
				temp1=temp1 & ",configOnly"
				temp2=temp2 & ",-1"
				temp3=temp3 & ",configOnly=-1"
			else
				if ucase(pptype)="DP" then
					temp1=temp1 & ",downloadable"
					temp2=temp2 & ",1"
					temp3=temp3 & ",downloadable=1"
					IsDownloadable=1
				else
					if ucase(pptype)="STANDARD" then
						temp1=temp1 & ",serviceSpec,configOnly"
						temp2=temp2 & ",0,0"
						temp3=temp3 & ",serviceSpec=0,configOnly=0"
					end if
				end if
			end if
		end if
	end if
	
	if IsDownloadable=0 then
		if downprdID>-1 then
			temp1=temp1 & ",downloadable"
			temp2=temp2 & "," & prd_downprd
			temp3=temp3 & ",downloadable=" & prd_downprd
			if prd_downprd="1" then
				IsDownloadable=1
			else
				IsDownloadable=0
			end if
		end if
	end if
		
	if BrandName<>"" then
		pIDBrand=checkbrand(BrandName,BrandLogo)
	else
		pIDBrand="0"
	end if
	
	
	query="Select idproduct,removed from products where sku='" & psku & "'"
	set rstemp4=connTemp.execute(query)
	testSKU=0
	if not rstemp4.eof then
		testSKU=1
		IDSKU=rstemp4("idproduct")
		PrdRmv=rstemp4("removed")
	end if	
		
	pAppend=0
	
	'BackInStock-S
	ChangeStock=0
	'BackInStock-E
	
	IF session("append")="1" THEN		
	
	'********************************************************************************
	' START:  APPEND IMORTED PRODUCTS
	'********************************************************************************
	
		query="Select removed from products where sku='" & psku & "'"
		set rstemp4=connTemp.execute(query)
		if not rstemp4.eof then


			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Append Product
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			set StringBuilderObj = new StringBuilder
			PrdRmv=0
			PrdRmv=rstemp4("removed")
			if brandnameid>-1 then
				StringBuilderObj.append ",IDBrand=" & pidBrand
			end if
			if lpriceid>-1 then
				StringBuilderObj.append ",listPrice=" & plprice
			end if
			if wpriceid>-1 then
				StringBuilderObj.append ", bToBPrice=" & pwprice
			end if
			if weightid>-1 then
				StringBuilderObj.append ", weight=" & pweight
			end if
			if unitslbID>-1 then
				StringBuilderObj.append ", pcprod_QtyToPound=" & unitslb
			end if
			if stockid>-1 then
				StringBuilderObj.append ", stock=" & pstock
				'BackInStock-S
				ChangeStock=1
				'BackInStock-E
			end if
			if timageid>-1 then
				StringBuilderObj.append ",smallImageUrl='" & ptimage & "'"
			end if
			if gimageid>-1 then
				StringBuilderObj.append ", imageUrl='" & pgimage & "'"
			end if
			if dimageid>-1 then
				StringBuilderObj.append ",largeImageUrl='" & pdimage & "'"
			end if
			if alttagtextid>-1 then
				StringBuilderObj.append ", pcProd_AltTagText = '" & palttagtext & "'"
			end if
			
			if activeid>-1 then
				StringBuilderObj.append ", active = " & pactive
			else
				if PrdRmv<>0 then
					StringBuilderObj.append ", active = -1"
				end if
			end if
			
			if featuredid>-1 then
				StringBuilderObj.append ", showInHome = " & pfeatured 
			end if
			
			if layTopID>-1 then
				StringBuilderObj.append ", pcProd_Top = '" & layTop & "'"
			end if
			
			if layTopLID>-1 then
				StringBuilderObj.append ", pcProd_TopLeft = '" & layTopL & "'"
			end if
			
			if layTopRID>-1 then
				StringBuilderObj.append ", pcProd_TopRight = '" & layTopR & "'"
			end if

			if layMiddleID>-1 then
				StringBuilderObj.append ", pcProd_Middle = '" & layMiddle & "'"
			end if
			
			if layTabsID>-1 then
				StringBuilderObj.append ", pcProd_Tabs = N'" & layTabs & "'"
			end if
			
			if layBottomID>-1 then
				StringBuilderObj.append ", pcProd_Bottom = '" & layBottom & "'"
			end if
			
			if mt_titleID>-1 then
				StringBuilderObj.append ", pcProd_MetaTitle = N'" & mt_title & "'"
			end if
			
			if mt_descID>-1 then
				StringBuilderObj.append ", pcProd_MetaDesc = N'" & mt_desc & "'"
			end if
			
			if mt_keyID>-1 then
				StringBuilderObj.append ", pcProd_MetaKeywords = N'" & mt_key & "'"
			end if
			
			'CONFIGURATOR-S
			if scBTO=1 then
			
				if hidebtopriceid>-1 then
					StringBuilderObj.append ", pcprod_hidebtoprice = " & prd_hidebtoprice 
				end if
				if hideconfid>-1 then
					StringBuilderObj.append ", pcprod_HideDefConfig = " & prd_hideconf 
				end if
				if dispurchaseid>-1 then
					StringBuilderObj.append ", NoPrices = " & prd_dispurchase 
				end if
				if skipdetailsid>-1 then
					StringBuilderObj.append ", pcProd_SkipDetailsPage = " & prd_skipdetails 
				end if
			
			end if
			'CONFIGURATOR-E
			
			if giftcertID>-1 then
				StringBuilderObj.append ", pcprod_GC = " & prd_giftcert 
			end if
			
			if surcharge1ID>-1 then
				StringBuilderObj.append ", pcProd_Surcharge1 = " & surcharge1
			end if
			
			if surcharge2ID>-1 then
				StringBuilderObj.append ", pcProd_Surcharge2 = " & surcharge2
			end if
			
			if prdnoteid>-1 then
				StringBuilderObj.append ", pcProd_PrdNotes = N'" & prdnote & "'"
			end if
			
			if playoutid>-1 then
				StringBuilderObj.append ", pcProd_DisplayLayout = '" & playout & "'"
			end if
			
			if eimagid>-1 then
				StringBuilderObj.append ", pcPrd_MojoZoom = " & eimag
			end if
			
			if hideaddimagesid>-1 then
				StringBuilderObj.append ", pcProd_AdditionalImages = " & hideaddimages
			end if
			
			if hideskuid>-1 then
				StringBuilderObj.append ", pcProd_HideSKU = " & hidesku
			end if
			
			'//Google Shopping
			if goCatid>-1 then
				StringBuilderObj.append ", pcProd_GoogleCat = '" & goCat & "'"
			end if
			
			if goGenid>-1 then
				StringBuilderObj.append ", pcProd_GoogleGender = '" & goGen & "'"
			end if
			
			if goAgeid>-1 then
				StringBuilderObj.append ", pcProd_GoogleAge = '" & goAge & "'"
			end if
			
			if goColorid>-1 then
				StringBuilderObj.append ", pcProd_GoogleColor = '" & goColor & "'"
			end if
			
			if goSizeid>-1 then
				StringBuilderObj.append ", pcProd_GoogleSize = '" & goSize & "'"
			end if
			
			if goPatid>-1 then
				StringBuilderObj.append ", pcProd_GooglePattern = '" & goPat & "'"
			end if
			
			if goMatid>-1 then
				StringBuilderObj.append ", pcProd_GoogleMaterial = '" & goMat & "'"
			end if
			
			'Start SDBA
			if prd_CostID>-1 then
				StringBuilderObj.append ", cost = " & prd_Cost 
			end if
			
			if prd_BackOrderID>-1 then
				StringBuilderObj.append ", pcProd_BackOrder = " & prd_BackOrder 
			end if

			if prd_ShipNDaysID>-1 then
				StringBuilderObj.append ", pcProd_ShipNDays = " & prd_ShipNDays 
			end if
			
			if prd_NotifyStockID>-1 then
				StringBuilderObj.append ", pcProd_NotifyStock = " & prd_NotifyStock 
			end if
			
			if prd_ReorderLevelID>-1 then
				StringBuilderObj.append ", pcProd_ReorderLevel = " & prd_ReorderLevel 
			end if
			
			if prd_IsDropShippedID>-1 then
				StringBuilderObj.append ", pcProd_IsDropShipped = " & prd_IsDropShipped 
			end if
			
			if prd_SupplierID>-1 then
				StringBuilderObj.append ", pcSupplier_ID = " & prd_Supplier 
			end if
			
			if (prd_DropShipperID>-1) or (prd_IsDropShipper="1") then
				StringBuilderObj.append ", pcDropShipper_ID = " & prd_DropShipper 
			end if
			'End SDBA
			
			if savingid>-1 then
				StringBuilderObj.append ", listhidden=" & psaving
			end if
			if specialid>-1 then
				StringBuilderObj.append ", hotDeal=" & pspecial 
			end if
			if rwpid>-1 then
				StringBuilderObj.append ", iRewardPoints=" & prwp 
			end if
			if ntaxid>-1 then
				StringBuilderObj.append ",notax=" & pntax
			end if
			if nshipid>-1 then
				StringBuilderObj.append ", noshipping=" & pnship
			end if
			if nforsaleid>-1 then
				StringBuilderObj.append ",formquantity=" & pnforsale
			end if
			if nforsalecopyid>-1 then
			StringBuilderObj.append ",emailtext=N'" & pnforsalecopy & "'"
			end if
			if nameid>-1 then
				StringBuilderObj.append ",description=N'" & pname & "'"
			end if
	
			if descid>-1 then
				StringBuilderObj.append ", details=N'" & pdesc & "'"
			end if
			
			if sdescid>-1 then
				StringBuilderObj.append ", sDesc=N'" & sdesc & "'"
			end if			
		
			if opriceid>-1 then
				StringBuilderObj.append ", price=" & poprice
			end if
			
			if distockid>-1 then
				StringBuilderObj.append ", noStock=" & distock
			end if
			
			if dishiptextid>-1 then
				StringBuilderObj.append ", noshippingtext=" & dishiptext
			end if
			
			if MQtyID>-1 then
				StringBuilderObj.append ",pcprod_minimumqty=" & MQty
			end if
			
			if VQtyID>-1 then
				StringBuilderObj.append ",pcprod_qtyvalidate=" & VQty
			end if
			
			if MQtyID>-1 and VQtyID>-1 then
				StringBuilderObj.append ",pcProd_multiQty=" & MQty
			end if
			'APP-S
			'**** Apparel Product Fields
				
			if ApparelID<>-1 then
				StringBuilderObj.append ",pcProd_Apparel=" & app_Apparel
			end if
		
			if ShowoutID<>-1 then
				StringBuilderObj.append ",pcProd_ShowStockMsg=" & app_Showout
			end if
		
			if app_OutMsg<>"" then
				StringBuilderObj.append ",pcProd_StockMsg=N'" & app_OutMsg & "'"
			end if
		
			if DTypeID<>-1 then
				StringBuilderObj.append ",pcprod_ApparelRadio=" & app_DType
			end if
		
			if ChartTextID<>-1 then
				StringBuilderObj.append ",pcProd_SizeLink='" & app_ChartText & "'"
			end if
		
			if ChartDescID<>-1 then
				StringBuilderObj.append ",pcProd_SizeInfo=N'" & app_ChartDesc & "'"
			end if
		
			if ChartImgID<>-1 then
				StringBuilderObj.append ",pcProd_SizeImg='" & app_ChartImg & "'"
			end if
		
			if ChartIURLID<>-1 then
				StringBuilderObj.append ",pcProd_SizeURL='" & app_ChartIURL & "'"
			end if
			'APP-E
			
			'Get product information before update
			query="select * from products where sku='" & psku & "'"
			set rstemp=conntemp.execute(query)			
			IF not rstemp.eof THEN
				PreRecord1=""
				PreRecord1=PreRecord1 & rstemp("idProduct") & "****"
				PreRecord1=PreRecord1 & "Pro" & "****"
				
				iCols = rstemp.Fields.Count
				
				for dd=1 to iCols-1
				
					FType="" & Rstemp.Fields.Item(dd).Type
					if (Ftype="202") or (Ftype="203") or (Ftype="135") then
						PTemp=Rstemp.Fields.Item(dd).Value
						if PTemp<>"" then
							PTemp=replace(PTemp,"'","''")
							PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
						end if
						if dd=1 then
							PreRecord1=PreRecord1 & "'" & PTemp & "'"
						else
							PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
						end if
					else
						PTemp="" & Rstemp.Fields.Item(dd).Value
						if PTemp<>"" then
						else
							PTemp="0"
						end if
						if dd=1 then
							PreRecord1=PreRecord1 & PTemp
						else
							PreRecord1=PreRecord1 & "@@@@@" & PTemp
						end if
					end if
				
				next
				PreRecords=PreRecords & PreRecord1 & vbcrlf
			END IF
			
			err.clear
			query="update products set sku=N'" & psku & "',removed=0" & StringBuilderObj.toString & temp3 & " where sku='" & psku & "'"
			query=replace(query,chr(34),"&quot;")
			query=replace(query,"**DD**",chr(34))
			set rstemp=conntemp.execute(query)
			
			If err.number<>0 Then
				if (err.number=-2147217904) then
					ErrorsReport=ErrorsReport & "<tr><td>" & "<strong>Record " & TotalXLSlines & " was NOT updated</strong>. There was an error in the import file.  Make sure you are not using text in a number field." & "</td></tr>" & vbcrlf '// display error
				else
					ErrorsReport=ErrorsReport & "<tr><td>" & "<strong>Record " & TotalXLSlines & " was NOT updated</strong>. Check the import file. Error reported:  " & err.description & "</td></tr>" & vbcrlf '// display error
				end if
				RecordError=true '// do not count import
			End If
			set rstemp = nothing
			pAppend=1
			
			call pcs_hookProductModified("", psku)
			call pcs_hookStockChanged("", psku)

			query="SELECT idProduct, weight FROM products WHERE sku='" &psku& "' AND pcprod_ParentPrd=0 ORDER by idProduct DESC"
			set rstemp=conntemp.execute(query)
			pIdProduct = rstemp("idProduct")
			pweight=rstemp("weight")
			if pweight<>"" then
			else
				pweight=0
			end if
			set rstemp = nothing
			
			call updPrdEditedDate(pIdProduct)
			
			IF OverSizeID>-1 then			
				if instr(OverSize,"||")>0 then
					OSArray=split(OverSize,"||")
					OverSize=""
					if ubound(OSArray)>=3 then
						For ds=0 to 3
							OverSize=OverSize & OSArray(ds) & "||"
						Next
					else
						For ds=0 to ubound(OSArray)
							OverSize=OverSize & OSArray(ds) & "||"
						Next
						For ds=ubound(OSArray)+1 to 3
							OverSize=OverSize & "0||"
						Next
					end if
					OverSize=OverSize & pweight
				end if				
				query="UPDATE products set removed=0,OverSizeSpec='" & OverSize & "' where idproduct=" & pidproduct
				set rstemp=connTemp.execute(query)
				call pcs_hookProductModified(pidproduct, "")
			END IF

			set StringBuilderObj = nothing
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Downloadable
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			If IsDownloadable=1 Then
						
				query="SELECT idProduct FROM DProducts WHERE idproduct=" & pIdProduct & " ORDER by idProduct DESC"
				set rstemp=conntemp.execute(query)
				
				IF not rstemp.eof THEN
				
					set StringBuilderObj = new StringBuilder
					if fileurlid>-1 then
						StringBuilderObj.append ",ProductURL='" & fileurl & "'"
					end if
					if urlexpireid>-1 then
						StringBuilderObj.append ",URLExpire=" & urlexpire 
					end if
					if expiredaysid>-1 then
						StringBuilderObj.append ",ExpireDays=" & expiredays 
					end if			
					if licenseid>-1 then
						StringBuilderObj.append ",License=" & license 
					end if
					if LocalLGid>-1 then
						StringBuilderObj.append ",LocalLG='" & localLG & "'"
					end if
					if RemoteLGid>-1 then
						StringBuilderObj.append ",RemoteLG='" & RemoteLG & "'" 
					end if
					if LFN1id>-1 then
						StringBuilderObj.append ",LicenseLabel1='" & LFN1 & "'"
					end if
					if LFN2id>-1 then
						StringBuilderObj.append ",LicenseLabel2='" & LFN2 & "'"
					end if
					if LFN3id>-1 then
						StringBuilderObj.append ",LicenseLabel3='" & LFN3 & "'"
					end if
					if LFN4id>-1 then
						StringBuilderObj.append ",LicenseLabel4='" & LFN4 & "'"
					end if
					if LFN5id>-1 then
						StringBuilderObj.append ",LicenseLabel5='" & LFN5 & "'"
					end if						
					
					if AddCopyid>-1 then
						StringBuilderObj.append ",AddToMail='" & AddCopy & "'"
					end if
					
					'Get downloadable product information before update
					query="select * from Dproducts where idproduct=" & pIDProduct
					set rstemp=conntemp.execute(query)
				
					If not rstemp.eof Then
		
						PreRecord1=""
						PreRecord1=PreRecord1 & pIdProduct & "****"
						PreRecord1=PreRecord1 & "DownPro" & "****"
						
						iCols = rstemp.Fields.Count
						for dd=1 to iCols-1
							FType="" & Rstemp.Fields.Item(dd).Type
							if (Ftype="202") or (Ftype="203") or (Ftype="135") then
								PTemp=Rstemp.Fields.Item(dd).Value
								if PTemp<>"" then
									PTemp=replace(PTemp,"'","''")
									PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
								end if
								if dd=1 then
									PreRecord1=PreRecord1 & "'" & PTemp & "'"
								else
									PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
								end if
							else
								PTemp="" & Rstemp.Fields.Item(dd).Value
								if PTemp<>"" then
								else
									PTemp="0"
								end if
								if dd=1 then
									PreRecord1=PreRecord1 & PTemp
								else
									PreRecord1=PreRecord1 & "@@@@@" & PTemp
								end if
							end if
						next
						PreRecords=PreRecords & PreRecord1 & vbcrlf
					End If
				
					query="update DProducts set idproduct=" & pIDProduct & StringBuilderObj.toString & " where idproduct=" & pIDProduct
					query=replace(query,chr(34),"&quot;")
					query=replace(query,"**DD**",chr(34))
					set rstemp=conntemp.execute(query)
					
					set StringBuilderObj = nothing
				
				ELSE
				
					PreRecord1=""
					PreRecord1=PreRecord1 & pIdProduct & "****"
					PreRecord1=PreRecord1 & "DelDownPro" & "****"
					PreRecords=PreRecords & PreRecord1 & vbcrlf
					
					query="INSERT INTO DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & pIdProduct & ",'" & fileurl & "'," & urlexpire & "," & expiredays & "," & license & ",'" & localLG & "','" & remoteLG & "','" & LFN1 & "','" & LFN2 & "','" & LFN3 & "','" & LFN4 & "','" & LFN5 & "','" & AddCopy & "')"
					query=replace(query,chr(34),"&quot;")
					query=replace(query,"**DD**",chr(34))
					set rstemp=conntemp.execute(query)
				
				END IF 'Update DProducts Table

			End If '// If IsDownloadable=1 Then
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// End: Downloadable
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END: Append Product
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


		else '// if not rstemp4.eof then
		
		
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Error in Append
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			SKUError=1
			ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": This SKU is not in the database." & "</td></tr>" & vbcrlf
			RecordError=true
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END: Error in Append
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
		end if
		
	'********************************************************************************
	' END:  APPEND IMORTED PRODUCTS
	'********************************************************************************

	ELSE '// IF session("append")="1" THEN	
	
	'********************************************************************************
	' START:  ADD IMORTED PRODUCTS
	'********************************************************************************
	

		If testSKU=1 Then	
		'///////////////////////////////////////////////////////////////////////////
		'// START:  SKU EXISTS
		'///////////////////////////////////////////////////////////////////////////
		
		
			if PrdRmv<>0 then	
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START:  Sku Deleted
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~		
			
				set StringBuilderObj = new StringBuilder
				
				if brandnameid>-1 then
					StringBuilderObj.append ",IDBrand=" & pidBrand
				end if
				if lpriceid>-1 then
					StringBuilderObj.append ",listPrice=" & plprice
				end if
				if wpriceid>-1 then
					StringBuilderObj.append ", bToBPrice=" & pwprice
				end if
				if weightid>-1 then
					StringBuilderObj.append ", weight=" & pweight
				end if
				if unitslbID>-1 then
					StringBuilderObj.append ", pcprod_QtyToPound=" & unitslb
				end if
				if stockid>-1 then
					StringBuilderObj.append ", stock=" & pstock
					'BackInStock-S
					ChangeStock=1
					'BackInStock-E
				end if
				if timageid>-1 then
					StringBuilderObj.append ",smallImageUrl='" & ptimage & "'"
				end if
				if gimageid>-1 then
					StringBuilderObj.append ", imageUrl='" & pgimage & "'"
				end if
				if dimageid>-1 then
					StringBuilderObj.append ",largeImageUrl='" & pdimage & "'"
				end if
				if alttagtextid>-1 then
					StringBuilderObj.append ", pcProd_AltTagText = '" & palttagtext & "'"
				end if
				if activeid>-1 then
					StringBuilderObj.append ", active = " & pactive
				else
					StringBuilderObj.append ", active = -1" 
				end if
				if featuredid>-1 then
					StringBuilderObj.append ", showInHome = " & pfeatured 
				end if
				
				if layTopID>-1 then
					StringBuilderObj.append ", pcProd_Top = '" & layTop & "'"
				end if
				
				if layTopLID>-1 then
					StringBuilderObj.append ", pcProd_TopLeft = '" & layTopL & "'"
				end if
				
				if layTopRID>-1 then
					StringBuilderObj.append ", pcProd_TopRight = '" & layTopR & "'"
				end if
				
				if layMiddleID>-1 then
					StringBuilderObj.append ", pcProd_Middle = '" & layMiddle & "'"
				end if

				if layTabsID>-1 then
					StringBuilderObj.append ", pcProd_Tabs = N'" & layTabs & "'"
				end if
				
				if layBottomID>-1 then
					StringBuilderObj.append ", pcProd_Bottom = '" & layBottom & "'"
				end if
				
				if mt_titleID>-1 then
					StringBuilderObj.append ", pcProd_MetaTitle = N'" & mt_title & "'"
				end if
				
				if mt_descID>-1 then
					StringBuilderObj.append ", pcProd_MetaDesc = N'" & mt_desc & "'"
				end if
				
				if mt_keyID>-1 then
					StringBuilderObj.append ", pcProd_MetaKeywords = N'" & mt_key & "'"
				end if
				
				'CONFIGURATOR-S
				if scBTO=1 then
			
					if hidebtopriceid>-1 then
						StringBuilderObj.append ", pcprod_hidebtoprice = " & prd_hidebtoprice 
					end if
					if hideconfid>-1 then
						StringBuilderObj.append ", pcprod_HideDefConfig = " & prd_hideconf 
					end if
					if dispurchaseid>-1 then
						StringBuilderObj.append ", NoPrices = " & prd_dispurchase 
					end if
					if skipdetailsid>-1 then
						StringBuilderObj.append ", pcProd_SkipDetailsPage = " & prd_skipdetails 
					end if
				
				end if
				'CONFIGURATOR-E
				
				if giftcertID>-1 then
					StringBuilderObj.append ", pcprod_GC = " & prd_giftcert 
				end if
				
				if surcharge1ID>-1 then
					StringBuilderObj.append ", pcProd_Surcharge1 = " & surcharge1
				end if
				
				if surcharge2ID>-1 then
					StringBuilderObj.append ", pcProd_Surcharge2 = " & surcharge2
				end if
				
				if prdnoteid>-1 then
					StringBuilderObj.append ", pcProd_PrdNotes = N'" & prdnote & "'"
				end if
				
				if playoutid>-1 then
					StringBuilderObj.append ", pcProd_DisplayLayout = '" & playout & "'"
				end if
				
				if eimagid>-1 then
					StringBuilderObj.append ", pcPrd_MojoZoom = " & eimag
				end if
				
				if hideaddimagesid>-1 then
					StringBuilderObj.append ", pcProd_AdditionalImages = " & hideaddimages
				end if
				
				if hideskuid>-1 then
					StringBuilderObj.append ", pcProd_HideSKU = " & hidesku
				end if
				
				'//Google Shopping
				if goCatid>-1 then
					StringBuilderObj.append ", pcProd_GoogleCat = '" & goCat & "'"
				end if
				
				if goGenid>-1 then
					StringBuilderObj.append ", pcProd_GoogleGender = '" & goGen & "'"
				end if
				
				if goAgeid>-1 then
					StringBuilderObj.append ", pcProd_GoogleAge = '" & goAge & "'"
				end if
				
				if goColorid>-1 then
					StringBuilderObj.append ", pcProd_GoogleColor = '" & goColor & "'"
				end if
				
				if goSizeid>-1 then
					StringBuilderObj.append ", pcProd_GoogleSize = '" & goSize & "'"
				end if
				
				if goPatid>-1 then
					StringBuilderObj.append ", pcProd_GooglePattern = '" & goPat & "'"
				end if
				
				if goMatid>-1 then
					StringBuilderObj.append ", pcProd_GoogleMaterial = '" & goMat & "'"
				end if
				
				'Start SDBA
				if prd_CostID>-1 then
					StringBuilderObj.append ", cost = " & prd_Cost 
				end if
				
				if prd_BackOrderID>-1 then
					StringBuilderObj.append ", pcProd_BackOrder = " & prd_BackOrder 
				end if
	
				if prd_ShipNDaysID>-1 then
					StringBuilderObj.append ", pcProd_ShipNDays = " & prd_ShipNDays 
				end if
				
				if prd_NotifyStockID>-1 then
					StringBuilderObj.append ", pcProd_NotifyStock = " & prd_NotifyStock 
				end if
				
				if prd_ReorderLevelID>-1 then
					StringBuilderObj.append ", pcProd_ReorderLevel = " & prd_ReorderLevel 
				end if
				
				if prd_IsDropShippedID>-1 then
					StringBuilderObj.append ", pcProd_IsDropShipped = " & prd_IsDropShipped 
				end if
				
				if prd_SupplierID>-1 then
					StringBuilderObj.append ", pcSupplier_ID = " & prd_Supplier 
				end if
				
				if (prd_DropShipperID>-1) or (prd_IsDropShipper="1") then
					StringBuilderObj.append ", pcDropShipper_ID = " & prd_DropShipper 
				end if
				'End SDBA
				
				if savingid>-1 then
					StringBuilderObj.append ", listhidden=" & psaving
				end if
				if specialid>-1 then
					StringBuilderObj.append ", hotDeal=" & pspecial 
				end if
				if rwpid>-1 then
					StringBuilderObj.append ", iRewardPoints=" & prwp 
				end if
				
				if ntaxid>-1 then
					StringBuilderObj.append ",notax=" & pntax
				end if
				if nshipid>-1 then
					StringBuilderObj.append ", noshipping=" & pnship
				end if
				if nforsaleid>-1 then
					StringBuilderObj.append ",formquantity=" & pnforsale
				end if
				if nforsalecopyid>-1 then
					StringBuilderObj.append ",emailtext=N'" & pnforsalecopy & "'"
				end if
				if nameid>-1 then
					StringBuilderObj.append ",description=N'" & pname & "'"
				end if
		
				if descid>-1 then
					StringBuilderObj.append ", details=N'" & pdesc & "'"
				end if
			
				if sdescid>-1 then
					StringBuilderObj.append ", sDesc=N'" & sdesc & "'"
				end if
			
				if opriceid>-1 then
					StringBuilderObj.append ", price=" & poprice
				end if
				
				if distockid>-1 then
					StringBuilderObj.append ", noStock=" & distock
				end if
				
				if dishiptextid>-1 then
					StringBuilderObj.append ", noshippingtext=" & dishiptext
				end if
				
				if MQtyID>-1 then
					StringBuilderObj.append ",pcprod_minimumqty=" & MQty
				end if
				
				if VQtyID>-1 then
					StringBuilderObj.append ",pcprod_qtyvalidate=" & VQty
				end if
				
				if MQtyID>-1 and VQtyID>-1 then
					StringBuilderObj.append ",pcProd_multiQty=" & MQty
				end if

				'APP-S
				'**** Apparel Product Fields
					
				if ApparelID<>-1 then
					StringBuilderObj.append ",pcProd_Apparel=" & app_Apparel
				end if
			
				if ShowoutID<>-1 then
					StringBuilderObj.append ",pcProd_ShowStockMsg=" & app_Showout
				end if
			
				if app_OutMsg<>"" then
					StringBuilderObj.append ",pcProd_StockMsg=N'" & app_OutMsg & "'"
				end if
			
				if DTypeID<>-1 then
					StringBuilderObj.append ",pcprod_ApparelRadio=" & app_DType
				end if
			
				if ChartTextID<>-1 then
					StringBuilderObj.append ",pcProd_SizeLink='" & app_ChartText & "'"
				end if
			
				if ChartDescID<>-1 then
					StringBuilderObj.append ",pcProd_SizeInfo=N'" & app_ChartDesc & "'"
				end if
			
				if ChartImgID<>-1 then
					StringBuilderObj.append ",pcProd_SizeImg='" & app_ChartImg & "'"
				end if
			
				if ChartIURLID<>-1 then
					StringBuilderObj.append ",pcProd_SizeURL='" & app_ChartIURL & "'"
				end if
				'APP-E
				
				'Get product information before update
				query="select * from products where sku='" & psku & "'"
				set rstemp=conntemp.execute(query)
				
				IF not rstemp.eof THEN

					PreRecord1=""
					PreRecord1=PreRecord1 & rstemp("idProduct") & "****"
					PreRecord1=PreRecord1 & "Pro" & "****"
					
					iCols = rstemp.Fields.Count
					for dd=1 to iCols-1
						FType="" & Rstemp.Fields.Item(dd).Type
						if (Ftype="202") or (Ftype="203") or (Ftype="135") then
							PTemp=Rstemp.Fields.Item(dd).Value
							if PTemp<>"" then
								PTemp=replace(PTemp,"'","''")
								PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
							end if
							if dd=1 then
								PreRecord1=PreRecord1 & "'" & PTemp & "'"
							else
								PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
							end if
						else
							PTemp="" & Rstemp.Fields.Item(dd).Value
							if PTemp<>"" then
							else
								PTemp="0"
							end if
							if dd=1 then
								PreRecord1=PreRecord1 & PTemp
							else
								PreRecord1=PreRecord1 & "@@@@@" & PTemp
							end if
						end if
					next
					PreRecords=PreRecords & PreRecord1 & vbcrlf
				END IF
	
				query="update products set sku='" & psku & "',removed=0" & StringBuilderObj.toString & temp3 & " where sku='" & psku & "'"
				query=replace(query,chr(34),"&quot;")
				query=replace(query,"**DD**",chr(34))
				set rstemp=conntemp.execute(query)
				pAppend=1
				
				call pcs_hookProductModified("", psku)
				call pcs_hookStockChanged("", psku)
				
				set StringBuilderObj = nothing

				query="SELECT idProduct,weight FROM products WHERE sku='" &psku& "' AND pcprod_ParentPrd=0 ORDER by idProduct DESC"
				set rstemp=conntemp.execute(query)
				pIdProduct = rstemp("idProduct")
				pweight=rstemp("weight")
				if pweight<>"" then
				else
					pweight=0
				end if
				
				call updPrdEditedDate(pIdProduct)
				
				IF OverSizeID>-1 then
				
					if instr(OverSize,"||")>0 then
						OSArray=split(OverSize,"||")
						OverSize=""
						if ubound(OSArray)>=3 then
							For ds=0 to 3
								OverSize=OverSize & OSArray(ds) & "||"
							Next
						else
							For ds=0 to ubound(OSArray)
								OverSize=OverSize & OSArray(ds) & "||"
							Next
							For ds=ubound(OSArray)+1 to 3
								OverSize=OverSize & "0||"
							Next
						end if
					
						OverSize=OverSize & pweight
					end if
					
					query="UPDATE products set removed=0,OverSizeSpec='" & OverSize & "' where idproduct=" & pidproduct
					set rstemp=connTemp.execute(query)
					call pcs_hookProductModified(pidproduct, "")
				
				END IF
				
				IF IsDownloadable=1 then
							
					query="SELECT idProduct FROM DProducts WHERE idproduct=" & pIdProduct & " ORDER by idProduct DESC"
					set rstemp=conntemp.execute(query)
					
					IF not rstemp.eof THEN
					
						set StringBuilderObj = new StringBuilder
						
						if fileurlid>-1 then
							StringBuilderObj.append ",ProductURL='" & fileurl & "'"
						end if
						if urlexpireid>-1 then
							StringBuilderObj.append ",URLExpire=" & urlexpire 
						end if
						if expiredaysid>-1 then
							StringBuilderObj.append ",ExpireDays=" & expiredays 
						end if			
						if licenseid>-1 then
							StringBuilderObj.append ",License=" & license 
						end if
						if LocalLGid>-1 then
							StringBuilderObj.append ",LocalLG='" & localLG & "'"
						end if
						if RemoteLGid>-1 then
							StringBuilderObj.append ",RemoteLG='" & RemoteLG & "'" 
						end if
						if LFN1id>-1 then
							StringBuilderObj.append ",LicenseLabel1='" & LFN1 & "'"
						end if
						if LFN2id>-1 then
							StringBuilderObj.append ",LicenseLabel2='" & LFN2 & "'"
						end if
						if LFN3id>-1 then
							StringBuilderObj.append ",LicenseLabel3='" & LFN3 & "'"
						end if
						if LFN4id>-1 then
							StringBuilderObj.append ",LicenseLabel4='" & LFN4 & "'"
						end if
						if LFN5id>-1 then
							StringBuilderObj.append ",LicenseLabel5='" & LFN5 & "'"
						end if											
					
						if AddCopyid>-1 then
							StringBuilderObj.append ",AddToMail='" & AddCopy & "'"
						end if
						
						'Get downloadable product information before update
						query="select * from Dproducts where idproduct=" & pIDProduct
						set rstemp=conntemp.execute(query)
						
						IF not rstemp.eof THEN
			
							PreRecord1=""
							PreRecord1=PreRecord1 & pIdProduct & "****"
							PreRecord1=PreRecord1 & "DownPro" & "****"
							
							iCols = rstemp.Fields.Count
							for dd=1 to iCols-1
								FType="" & Rstemp.Fields.Item(dd).Type
								if (Ftype="202") or (Ftype="203") or (Ftype="135") then
									PTemp=Rstemp.Fields.Item(dd).Value
									if PTemp<>"" then
										PTemp=replace(PTemp,"'","''")
										PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
									end if
									if dd=1 then
										PreRecord1=PreRecord1 & "'" & PTemp & "'"
									else
										PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
									end if
								else
									PTemp="" & Rstemp.Fields.Item(dd).Value
									if PTemp<>"" then
									else
										PTemp="0"
									end if
									if dd=1 then
										PreRecord1=PreRecord1 & PTemp
									else
										PreRecord1=PreRecord1 & "@@@@@" & PTemp
									end if
								end if
							next
							PreRecords=PreRecords & PreRecord1 & vbcrlf
						END IF			
						
						query="update DProducts set idproduct=" & pIDProduct & StringBuilderObj.toString & " where idproduct=" & pIDProduct
						query=replace(query,chr(34),"&quot;")
						query=replace(query,"**DD**",chr(34))
						set rstemp=conntemp.execute(query)
						
						set StringBuilderObj = nothing
					
					ELSE
					
						PreRecord1=""
						PreRecord1=PreRecord1 & pIdProduct & "****"
						PreRecord1=PreRecord1 & "DelDownPro" & "****"
						PreRecords=PreRecords & PreRecord1 & vbcrlf
						
						query="INSERT INTO DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & pIdProduct & ",'" & fileurl & "'," & urlexpire & "," & expiredays & "," & license & ",'" & localLG & "','" & remoteLG & "','" & LFN1 & "','" & LFN2 & "','" & LFN3 & "','" & LFN4 & "','" & LFN5 & "','" & AddCopy & "')"
						query=replace(query,chr(34),"&quot;")
						query=replace(query,"**DD**",chr(34))
						set rstemp=conntemp.execute(query)
					
					END IF 'Update DProducts Table
	
				END IF
	
				testSKU=0			
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END: Sku Deleted
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			else
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Sku Exists
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
				ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": SKU " & psku & " could not be imported because it already exists." & "</td></tr>" & vbcrlf
				RecordError=true
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Sku Exists
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			end if		

		'///////////////////////////////////////////////////////////////////////////
		'// END:  SKU EXISTS
		'///////////////////////////////////////////////////////////////////////////
		
		Else
		
		'///////////////////////////////////////////////////////////////////////////
		'// START:  SKU DOES NOT EXIST
		'///////////////////////////////////////////////////////////////////////////
		
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Create Insert Query
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~			
			
			dim dtTodaysDate			
			dtTodaysDate=Date()
			if SQL_Format="1" then
				dtTodaysDate=(day(dtTodaysDate)&"/"&month(dtTodaysDate)&"/"&year(dtTodaysDate))
			else
				dtTodaysDate=(month(dtTodaysDate)&"/"&day(dtTodaysDate)&"/"&year(dtTodaysDate))
			end if
			
			'CONFIGURATOR-S
			tmp_str1=""
			tmp_str2=""
			if scBTO=1 then
			
				if hidebtopriceid>-1 then
					tmp_str1=tmp_str1 & ",pcprod_hidebtoprice"
					tmp_str2=tmp_str2 & "," & prd_hidebtoprice 
				end if
				if hideconfid>-1 then
					tmp_str1=tmp_str1 & ",pcprod_HideDefConfig"
					tmp_str2=tmp_str2 & "," & prd_hideconf 
				end if
				if dispurchaseid>-1 then
					tmp_str1=tmp_str1 & ",NoPrices"
					tmp_str2=tmp_str2 & "," & prd_dispurchase 
				end if
				if skipdetailsid>-1 then
					tmp_str1=tmp_str1 & ",pcProd_SkipDetailsPage"
					tmp_str2=tmp_str2 & "," & prd_skipdetails 
				end if
			
			end if
			'CONFIGURATOR-E
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END: Create Insert Query
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// START: Run Insert Query
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			query="INSERT INTO products (pcProd_Top,pcProd_TopLeft,pcProd_TopRight,pcProd_Middle,pcProd_Tabs,pcProd_Bottom,pcProd_PrdNotes, pcProd_DisplayLayout, pcPrd_MojoZoom, pcProd_HideSKU, pcProd_Surcharge1, pcProd_Surcharge2, IDBrand,sku, description, details, price, listPrice, bToBPrice, imageUrl, listhidden, hotDeal,iRewardPoints,weight, stock, active,showInHome, idSupplier, smallImageUrl,largeImageUrl, notax, noshipping,formquantity,emailtext" & temp1 & ",sDesc,nostock,noshippingtext, pcprod_EnteredOn,pcprod_qtyvalidate,pcprod_minimumqty,cost,pcProd_BackOrder,pcProd_ShipNDays,pcProd_NotifyStock,pcProd_ReorderLevel,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,pcprod_GC,pcProd_MetaTitle,pcProd_MetaDesc,pcProd_MetaKeywords" & tmp_str1 & ",pcprod_QtyToPound,pcProd_multiQty,pcProd_GoogleCat,pcProd_GoogleGender,pcProd_GoogleAge,pcProd_GoogleSize,pcProd_GoogleColor,pcProd_GooglePattern,pcProd_GoogleMaterial,pcProd_Apparel,pcProd_ShowStockMsg,pcProd_StockMsg,pcprod_ApparelRadio,pcProd_SizeLink,pcProd_SizeInfo,pcProd_SizeImg,pcProd_SizeURL,pcProd_AdditionalImages,pcProd_AltTagText) VALUES ('" & layTop & "','" & layTopL & "','" & layTopR & "','" & layMiddle & "',N'" & layTabs & "','" & layBottom & "',N'" & prdnote & "','" & playout & "'," & eimag & "," & hidesku & "," & surcharge1 & "," & surcharge2 & "," & pIDBrand & ",N'" &psku& "',N'" &pname& "',N'" & pdesc& "'," &poprice& "," &plprice& "," &pwprice& ",'" &pgimage& "'," & psaving & "," & pspecial & "," & prwp & "," &pweight& "," &pstock& "," &pactive& "," & pfeatured & ",10,'" &ptimage& "','"&pdimage&"',"&pntax&","&pnship&"," & pnforsale & ",N'" & pnforsalecopy & "'" & temp2 & ",N'" & sDesc & "'," & distock & "," & dishiptext & ",'"&dtTodaysDate&"'," & VQty & "," & MQty & "," & prd_Cost & "," & prd_BackOrder & "," & prd_ShipNDays & "," & prd_NotifyStock & "," & prd_ReorderLevel & "," & prd_Supplier & "," & prd_IsDropShipped & "," & prd_DropShipper & "," & prd_giftcert & ",N'" & mt_title & "',N'" & mt_desc & "',N'" & mt_key & "'" & tmp_str2 & "," & unitslb & "," & MQty & ",'" & goCat & "','" & goGen & "','" & goAge & "','" & goSize & "','" & goColor & "','" & goPat & "','" & goMat & "'," & app_Apparel & "," & app_Showout & ",N'" & app_OutMsg & "'," & app_DType & ",'" & app_ChartText & "',N'" & app_ChartDesc & "','" & app_ChartImg & "','" & app_ChartIURL & "'," & hideaddimages & ",'" & palttagtext & "')"
			query=replace(query,chr(34),"&quot;")
			query=replace(query,"**DD**",chr(34))
			err.clear
			set rstemp=server.CreateObject("ADODB.RecordSet")
			set rstemp=conntemp.execute(query)
			If err.number<>0 Then
				ErrorsReport=ErrorsReport & "<tr><td>" & "<strong>Record " & TotalXLSlines & " NOT imported.</strong> Error Details:  " & err.description & "</td></tr>" & vbcrlf '// display error
				RecordError=true '// do not count import
			End If
			set rstemp=nothing
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'// END: Run Insert Query
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


			query="SELECT idProduct,weight FROM products WHERE sku='" &psku& "' AND price=" &poprice& " AND pcprod_ParentPrd=0 ORDER by idProduct DESC"
			set rstemp=conntemp.execute(query)
			pIdProduct = rstemp("idProduct")
			pweight=rstemp("weight")
			if pweight<>"" then
			else
			pweight=0
			end if
			
			if instr(OverSize,"||")>0 then
				OSArray=split(OverSize,"||")
				OverSize=""
				if ubound(OSArray)>=3 then
					For ds=0 to 3
						OverSize=OverSize & OSArray(ds) & "||"
					Next
				else
					For ds=0 to ubound(OSArray)
						OverSize=OverSize & OSArray(ds) & "||"
					Next
					For ds=ubound(OSArray)+1 to 3
						OverSize=OverSize & "0||"
					Next
				end if
				OverSize=OverSize & pweight
			end if
			
			query="UPDATE products set removed=0,OverSizeSpec='" & OverSize & "' where idproduct=" & pidproduct
			set rstemp=connTemp.execute(query)
			call pcs_hookProductModified(pidproduct, "")
			
			IF IsDownloadable=1 then
						
				query="INSERT INTO DProducts (IdProduct,ProductURL,URLExpire,ExpireDays,License,LocalLG,RemoteLG,LicenseLabel1,LicenseLabel2,LicenseLabel3,LicenseLabel4,LicenseLabel5,AddToMail) values (" & pIdProduct & ",'" & fileurl & "'," & urlexpire & "," & expiredays & "," & license & ",'" & localLG & "','" & remoteLG & "','" & LFN1 & "','" & LFN2 & "','" & LFN3 & "','" & LFN4 & "','" & LFN5 & "','" & AddCopy & "')"
				query=replace(query,chr(34),"&quot;")
				query=replace(query,"**DD**",chr(34))
				set rstemp=conntemp.execute(query)
	
			END IF


		'///////////////////////////////////////////////////////////////////////////
		'// END:  SKU DOES NOT EXIST
		'///////////////////////////////////////////////////////////////////////////			
		End If '// if testSKU=1 then



	'********************************************************************************
	' END:  ADD IMORTED PRODUCTS
	'********************************************************************************
	END IF '// IF session("append")="1" THEN


	
	'********************************************************************************
	' START:  NO ERRORS
	'********************************************************************************
	IF RecordError=false THEN
		if pAppend=1 then
			query="SELECT idProduct FROM products WHERE sku='" &psku& "' AND pcprod_ParentPrd=0 ORDER by idProduct DESC"
		else
			query="SELECT idProduct FROM products WHERE sku='" &psku& "' AND price=" &poprice& " AND pcprod_ParentPrd=0 ORDER by idProduct DESC"
		end if

		set rstemp=conntemp.execute(query)
	 
		pIdProduct = rstemp("idProduct")
		
		'BackInStock-S
		if ChangeStock=1 then
			Call pcs_hookInStockEvent(pIdProduct, "")
		end if
		'BackInStock-E
	
		IF (session("append")="1") and (session("movecat")="3") then
		
		ELSE
			
			if (session("append")="1") and (session("movecat")="2") and (pcategory&pcategory1&pcategory2<>"") and (pAppend=1) then
			
				'Get category-product information before update
				query="select * from categories_products where idProduct=" &pIdProduct
				set rstemp=conntemp.execute(query)
				
				do while not rstemp.eof
		
					PreRecord1=""
					PreRecord1=PreRecord1 & "Add" & "****"
					
					iCols = rstemp.Fields.Count
					for dd=0 to 1
						FType="" & Rstemp.Fields.Item(dd).Type
						if (Ftype="202") or (Ftype="203") or (Ftype="135") then
							PTemp=Rstemp.Fields.Item(dd).Value
							if PTemp<>"" then
								PTemp=replace(PTemp,"'","''")
								PTemp=replace(PTemp,vbcrlf,"DuLTVDu")
							end if
							if dd=0 then
								PreRecord1=PreRecord1 & "'" & PTemp & "'"
							else
								PreRecord1=PreRecord1 & "@@@@@'" & PTemp & "'"
							end if
						else
							PTemp="" & Rstemp.Fields.Item(dd).Value
							if PTemp<>"" then
							else
							PTemp="0"
							end if
							if dd=0 then
							PreRecord1=PreRecord1 & PTemp
							else
							PreRecord1=PreRecord1 & "@@@@@" & PTemp
							end if
						end if
					next
					CATRecords=CATRecords & PreRecord1 & vbcrlf
					rstemp.MoveNext
				loop
			
				query="DELETE from categories_products where idProduct=" &pIdProduct
				set rstemp=conntemp.execute(query)
			end if
			
			if ppcategory<>"" then
				pidParentCategory=checkparent(ppcategory)
			else
				if pcategory<>"" then
					pidParentCategory=new_checkparent(pcategory)
				else
					pidParentCategory=1
				end if
			end if
			
			pIdCategory=""
			if pcategory<>"" then
				pIdCategory=checkcategory(pcategory,pidParentCategory,pcsimage,pclimage,SCATDesc,LCATDesc)
				if pIdCategory=-1 then
					pIdCategory=checktempcategory
					TempProducts=TempProducts & "Record " & TotalXLSlines & ": Product SKU: " & psku & " - Product Name: " & pname & vbcrlf
				end if
			else
				if (session("append")<>"1") or (pAppend<>1) then
					pIdCategory=checktempcategory
					TempProducts=TempProducts & "Record " & TotalXLSlines & ": Product SKU: " & psku & " - Product Name: " & pname & vbcrlf
				end if
			end if
			
			if pIdCategory<>"" then
				testCAT=0
				if testSKU=1 then
					query="select * from categories_products where idProduct=" & IDSKU & " and idCategory=" &pIdCategory
					set rstemp99=conntemp.execute(query)
					if not rstemp99.eof then
						testCAT=1
					end if
				end if
				
				if testCAT=0 then
				
					'Get category-product information before add
					PreRecord1=""
					PreRecord1=PreRecord1 & "Del" & "****" & pIdproduct & "@@@@@" & pIdCategory
					CATRecords=PreRecord1 & vbcrlf & CATRecords
				
					query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory& ")"
					set rstemp=conntemp.execute(query)
				end if
			end if
			
			if pcategory1<>"" then
				if ppcategory1<>"" then
					pidParentCategory1=checkparent(ppcategory1)
				else
					if pcategory1<>"" then
						pidParentCategory1=new_checkparent(pcategory1)
					else
						pidParentCategory1=1
					end if
				end if
				pIdCategory1=checkcategory(pcategory1,pidParentCategory1,pcsimage1,pclimage1,SCATDesc1,LCATDesc1)
				if pIdCategory1<>-1 then
					testCAT=0
					if testSKU=1 then
						query="select * from categories_products where idProduct=" & IDSKU & " and idCategory=" &pIdCategory1
						set rstemp99=conntemp.execute(query)
						if not rstemp99.eof then
							testCAT=1
						end if
					end if
					if testCAT=0 then
					
						'Get category-product information before add
						PreRecord1=""
						PreRecord1=PreRecord1 & "Del" & "****" & pIdproduct & "@@@@@" & pIdCategory1
						CATRecords=PreRecord1 & vbcrlf & CATRecords
			
						query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory1& ")"
						set rstemp=conntemp.execute(query)
					end if		
				end if
			end if		
		
			if pcategory2<>"" then
				if ppcategory2<>"" then
					pidParentCategory2=checkparent(ppcategory2)
				else
					if pcategory2<>"" then
						pidParentCategory2=new_checkparent(pcategory2)
					else
						pidParentCategory2=1
					end if
				end if
				pIdCategory2=checkcategory(pcategory2,pidParentCategory2,pcsimage2,pclimage2,SCATDesc2,LCATDesc2)
				if pIdCategory2<>-1 then
					testCAT=0
					if testSKU=1 then
						query="select * from categories_products where idProduct=" & IDSKU & " and idCategory=" &pIdCategory2
						set rstemp99=conntemp.execute(query)
						if not rstemp99.eof then
							testCAT=1
						end if
					end if
					if testCAT=0 then			
						'Get category-product information before add
						PreRecord1=""
						PreRecord1=PreRecord1 & "Del" & "****" & pIdproduct & "@@@@@" & pIdCategory2
						CATRecords=PreRecord1 & vbcrlf & CATRecords			
						query="INSERT INTO categories_products (idProduct, idCategory) VALUES (" &pIdProduct& "," &pIdCategory2& ")"
						set rstemp=conntemp.execute(query)
					end if		
				end if
			end if
		
		END IF
		
		
	END IF '// IF RecordError=false THEN
	'********************************************************************************
	' END:  NO ERRORS
	'********************************************************************************
	
	
	'Start SDBA
	IF RecordError=false THEN
		if (clng(prd_Supplier)>0) OR (clng(prd_DropShipper)>0) then
			myquery="DELETE FROM pcDropShippersSuppliers WHERE idproduct=" & pIdProduct
			set rstemp=connTemp.execute(myquery)
			set rstemp=nothing
			myquery="INSERT INTO pcDropShippersSuppliers (idproduct,pcDS_IsDropShipper) VALUES (" & pIdProduct & "," & prd_IsDropShipper & ")"
			set rstemp=connTemp.execute(myquery)
			set rstemp=nothing
		end if
	END IF
	'End SDBA

	'S-UPDATE PRODUCT SEARCH FIELDS
	IF RecordError=false THEN
		For m=0 to 2
			if (IDcustom(m)>"0") AND (Customcontent(m)>"0") then
				query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & pIdProduct & " AND idSearchData IN (SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & IDcustom(m) & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
	
				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & pIdProduct & "," & Customcontent(m) & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
			end if
		Next
	END IF
	'E-UPDATE PRODUCT SEARCH FIELDS

	'S-UPDATE PRODUCT PRICING CATEGORIES
	IF RecordError=false THEN
		For m=0 to 4
			if (IDPriceCat(m)>"0") AND (pricecats(m)>"0") then
				query="DELETE FROM pcCC_Pricing WHERE idproduct=" & pIdProduct & " AND idcustomerCategory=" & IDPriceCat(m) & ";"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
	
				query="INSERT INTO pcCC_Pricing (idproduct,idcustomerCategory,pcCC_Price) VALUES (" & pIdProduct & "," & IDPriceCat(m) & "," & pricecats(m) & ");"
				Set rstemp=conntemp.execute(query)
				Set rstemp=nothing
			end if
		Next
	END IF
	'E-UPDATE PRODUCT PRICING CATEGORIES
	
	'**** Import/Update Product Options
	CheckCount=0
	IF RecordError=false THEN
		
		'*** Import/Update Option Group 1
		If prd_Opt1<>"" AND InvalidGrp1=0 then
			call ImportPrdOptions(pIdProduct,prd_Opt1,prd_Attr1,prd_Opt1Req,prd_Opt1Order)
		else
			if Opt1ID<>-1 AND prd_Opt1<>"" AND InvalidGrp1=1 then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because the fields are not properly formatted or it does not have any attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
			end if
		end if
		'*** End of Import/Update Option Group 1
		'*** Import/Update Option Group 2
		If prd_Opt2<>"" AND InvalidGrp2=0 then
			call ImportPrdOptions(pIdProduct,prd_Opt2,prd_Attr2,prd_Opt2Req,prd_Opt2Order)
		else
			if Opt2ID<>-1 AND prd_Opt2<>"" AND InvalidGrp2=1 then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because the fields are not properly formatted or it does not have any attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
			end if
		end if
		'*** End of Import/Update Option Group 2
		'*** Import/Update Option Group 3
		If prd_Opt3<>"" AND InvalidGrp3=0 then
			call ImportPrdOptions(pIdProduct,prd_Opt3,prd_Attr3,prd_Opt3Req,prd_Opt3Order)
		else
			if Opt3ID<>-1 AND prd_Opt3<>"" AND InvalidGrp3=1 then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because the fields are not properly formatted or it does not have any attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
			end if
		end if
		'*** End of Import/Update Option Group 3
		'*** Import/Update Option Group 4
		If prd_Opt4<>"" AND InvalidGrp4=0 then
			call ImportPrdOptions(pIdProduct,prd_Opt4,prd_Attr4,prd_Opt4Req,prd_Opt4Order)
		else
			if Opt4ID<>-1 AND prd_Opt4<>"" AND InvalidGrp4=1 then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because the fields are not properly formatted or it does not have any attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
			end if
		end if
		'*** End of Import/Update Option Group 4
		'*** Import/Update Option Group 5
		If prd_Opt5<>"" AND InvalidGrp5=0 then
			call ImportPrdOptions(pIdProduct,prd_Opt5,prd_Attr5,prd_Opt5Req,prd_Opt5Order)
		else
			if Opt5ID<>-1 AND prd_Opt5<>"" AND InvalidGrp5=1 then
				if CheckCount=0 then
					ErrorsReport=ErrorsReport & "<tr><td>" & "Record " & TotalXLSlines & ": Product SKU " & psku & " - can not add/update product option group because the fields are not properly formatted or it does not have any attributes." & "</td></tr>" & vbcrlf
					PrdWithoutOpts=PrdWithoutOpts+1
					CheckCount=1
				end if
			end if
		end if
		'*** End of Import/Update Option Group 5
	
	END IF
	'**** End of Import/Update Product Options

	'**** Import/Update Gift Certificates Settings
	IF RecordError=false THEN
		if (giftexpID<>-1) or (giftelectID<>-1) or (giftgenID<>-1) or (giftexpdateID<>-1) or (giftexpdaysID<>-1) or (giftcustgenfileID<>-1) then
		
			query="SELECT pcGC_IDProduct FROM pcGC WHERE pcGC_IDProduct=" & pIdProduct & ";"
			set rstemp=connTemp.execute(query)
			
			if not rstemp.eof then
				tmp4="pcGC_IDProduct=" & pIdProduct
				
				if (giftexpID<>-1) then
					tmp4=tmp4 & ", pcGC_Exp=" & prd_giftexp
				end if
				if (giftelectID<>-1) then
					tmp4=tmp4 & ", pcGC_EOnly=" & prd_giftelect
				end if
				if (giftgenID<>-1) then
					tmp4=tmp4 & ", pcGC_CodeGen=" & prd_giftgen
				end if
				if (giftexpdateID<>-1) then
					if trim(prd_giftexpdate)<>"" then
						tmp4=tmp4 & ", pcGC_ExpDate='" & prd_giftexpdate & "'"
					end if
				end if
				if (giftexpdaysID<>-1) then
					tmp4=tmp4 & ", pcGC_ExpDays=" & prd_giftexpdays
				end if
				if (giftcustgenfileID<>-1) then
					tmp4=tmp4 & ", pcGC_GenFile='" & prd_giftcustgenfile & "'"
				end if
				
				query="UPDATE pcGC SET " & tmp4 & " WHERE pcGC_IDProduct=" & pIdProduct & ";"
				set rstemp=connTemp.execute(query)
				set rstemp=nothing
			else
				if prd_giftcert<>"0" then
					tmp4=""
					tmp5=""
			
					if (giftexpID<>-1) then
						tmp4=tmp4 & ", pcGC_Exp"
						tmp5=tmp5 & "," & prd_giftexp
					end if
					if (giftelectID<>-1) then
						tmp4=tmp4 & ", pcGC_EOnly"
						tmp5=tmp5 & "," & prd_giftelect
					end if
					if (giftgenID<>-1) then
						tmp4=tmp4 & ", pcGC_CodeGen"
						tmp5=tmp5 & "," & prd_giftgen
					end if
					if (giftexpdateID<>-1) then
						if trim(prd_giftexpdate)<>"" then
							tmp4=tmp4 & ", pcGC_ExpDate"
							tmp5=tmp5 & ",'" & prd_giftexpdate & "'"
						end if
					end if
					if (giftexpdaysID<>-1) then
						tmp4=tmp4 & ", pcGC_ExpDays"
						tmp5=tmp5 & "," & prd_giftexpdays
					end if
					if (giftcustgenfileID<>-1) then
						tmp4=tmp4 & ", pcGC_GenFile"
						tmp5=tmp5 & ",'" & prd_giftcustgenfile & "'"
					end if
					
					if (tmp4<>"") and (tmp5<>"") then
						tmp4="pcGC_IDProduct" & tmp4
						tmp5=pIdProduct & tmp5
						query="INSERT INTO pcGC (" & tmp4 & ") VALUES (" & tmp5 & ");"
						set rstemp=connTemp.execute(query)
						set rstemp=nothing
					end if
				end if
			end if
			set rstemp=nothing
		
		end if
	END IF
	'**** End of Import/Update Gift Certificates Settings

	If statusAPP="1" OR scAPP=1 Then
	
		'// Update Apparel Sub-products Prices & Description, pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID
		query="SELECT price,bToBPrice,pcSupplier_ID,pcProd_IsDropShipped,pcDropShipper_ID,description FROM Products WHERE idproduct=" & pIdProduct & " AND removed=0 AND pcProd_Apparel=1;"
		set rsT=connTemp.execute(query)
		
		if not rsT.eof then
			app_Price=rsT("price")
			app_WPrice=rsT("bToBPrice")
			app_SupplierID=rsT("pcSupplier_ID")
			app_IsDropShipped=rsT("pcProd_IsDropShipped")
			app_DropShipperID=rsT("pcDropShipper_ID")
			app_Description=rsT("description")
			if app_WPrice<>"" then
			else
				app_WPrice="0"
			end if
			if cdbl(app_WPrice)>0 then
			else
				app_WPrice=app_Price
			end if
			set rsT=nothing
			
			query="UPDATE Products SET pcSupplier_ID=" & app_SupplierID & ",pcProd_IsDropShipped=" & app_IsDropShipped & ",pcDropShipper_ID=" & app_DropShipperID & ",price=pcProd_AddPrice+" & app_Price & ",bToBPrice=pcProd_AddWPrice+" & app_WPrice & " WHERE pcProd_ParentPrd=" & pIdProduct & " AND removed=0 AND pcProd_Apparel=0;"
			set rsT=connTemp.execute(query)
			set rsT=nothing
			call pcs_hookProductModified(pIdProduct, "")
			
			'Only update sub-product name(s) when parent product name was changed
			IF nameid<>-1 THEN
			
				query="Select idProduct, pcprod_Relationship from Products where pcprod_ParentPrd=" & pIdProduct & " AND removed=0 AND active=0"
				set rsQ=connTemp.execute(query)
				
				if not rsQ.eof then
					tmpArr=rsQ.getRows()
					intCount=ubound(tmpArr,2)
					set rsQ=nothing
					
					For ik=0 to intCount
						pcv_SIdproduct=tmpArr(0,ik)
						
						pcv_TempArr=split(tmpArr(1,ik),"_")	
						
						pcv_newName="("	
						For i=1 to ubound(pcv_TempArr)
							pcv_Opt1=pcv_TempArr(i)
					
							query="select idOption from options_optionsGroups where idoptoptgrp=" & pcv_Opt1
							set rs=connTemp.execute(query)
					
							if not rs.eof then
								pcv_ROpt1=rs("idOption")
							end if
					
							query="select optionDescrip from Options where idOption=" & pcv_ROpt1
							set rs=connTemp.execute(query)
					
							if not rs.eof then
								pcv_Code1=rs("optionDescrip")
							end if
					
							pcv_newName=pcv_newName & pcv_Code1 
							if (i<ubound(pcv_TempArr)) then
								pcv_newName=pcv_newName & " - "
							end if
						Next
						
						pcv_newName=pcv_newName & ")"
						pcv_newName=replace(pcv_newName,"'","''")
						pcv_newName=app_Description & " " & pcv_newName
						
						query="UPDATE products SET description=N'" & pcv_newName & "' WHERE idproduct=" & pcv_SIdproduct
						set rsT2=connTemp.execute(query)
						set rsT2=Nothing
						
						call pcs_hookProductModified(pcv_SIdproduct, "")
					
					Next
				
				end if
				set rsQ=nothing
			
			END IF
			
		end if
		set rsT=nothing
		'**** End of Update Apparel Sub-products Prices
	
	End If
			
	if RecordError = false then
		ImportedRecords = ImportedRecords+1
	end if

end if

	count=count + 1 
	rsExcel.MoveNext

Loop
rsexcel.Close	
set rsexcel=nothing
cnnExcel.close
set cnnExcel=nothing

	
	
session("TempProducts")=TempProducts 
session("ErrorsReport")=ErrorsReport 
session("iPageCurrent")=iPageCurrent+1 
session("TotalXLSlines")=TotalXLSlines 
session("ImportedRecords")=ImportedRecords 

If Cint(iPageCurrent) < Cint(iPageCount) Then
	session("PrdWithoutOpts")=PrdWithoutOpts
	call closeDb()
response.redirect "step4-xls.asp?" & pcv_Query
else
		session("importfile")=""
		session("totalfields")=0
end if
	
	if ImportedRecords>0 then
	
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	Set afi=fs.CreateTextFile(server.MapPath(".") & "\importlogs\prologs.txt",True)
		
	afi.Writeline(UpdateType)
	afi.Writeline(PreIDProduct)
	afi.Writeline(PreIDCategory)
	afi.Writeline(PreIDBrand)
	afi.Writeline(PreRecords)
	afi.Close
	
	Set afi = fs.GetFile(server.MapPath(".") & "\importlogs\catlogs.txt")
	afi.Delete
	
	if err.number<>0 then
	err.Description=""
	err.number=0
	end if
	
	if CATRecords<>"" then
	Set afi=fs.CreateTextFile(server.MapPath(".") & "\importlogs\catlogs.txt",True)
	afi.Writeline(CATRecords)
	afi.Close
	end if
	end if
	
	if SKUError=1 then
	ErrorsReport="<tr><td>One of the records you are importing does not currently exist in the database. The Append feature is strictly for modifying existing product information. Please correct the error and try again." & "</td></tr>" & vbcrlf&vbcrlf &ErrorsReport
	session("ErrorsReport")=ErrorsReport
	end if

if session("append")="1" then 
	pageTitle = "UPDATE"
else
	pageTitle = "IMPORT" 
end if 
pageTitle = pageTitle & " PRODUCT DATA WIZARD - Review Import Results"
section = "products" %>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<tr>
	<td valign="top">
        <table class="pcCPcontent">
	<tr>
            <td colspan="2"><h2>Steps:</h2></td>
	</tr>
	<tr>
            <td width="5%" align="right"><img border="0" src="images/step1.gif"></td>
            <td width="95%"><font color="#A8A8A8">Select product data file</font></td>
	</tr>
	<tr>
            <td align="right"><img border="0" src="images/step2.gif"></td>
            <td><font color="#A8A8A8">Map fields</font></td>
	</tr>
	<tr>
            <td align="right"><img border="0" src="images/step3.gif"></td>
            <td><font color="#A8A8A8">Confirm mapping</font></td>
	</tr>
	<tr>
            <td align="right"><img border="0" src="images/step4a.gif"></td>
            <td><strong><%if session("append")="1" then%>Update<%else%>Import<%end if%> results</strong></td>
	</tr>
	</table>
	<div class="pcCPmessage">
		<%if ImportedRecords-PrdWithoutOpts>0 then%>
      <div>A total of <b><%=ImportedRecords-PrdWithoutOpts%></b> records were <%if session("append")="1" then%>updated<%else%>imported<%end if%> successfully!</div>
    <%end if%>
		<%if PrdWithoutOpts>0 then%>
      <div>A total of <b><font color="#FF0000"><%=PrdWithoutOpts%></font></b> records were <%if session("append")="1" then%>updated<%else%>imported<%end if%>, but <u>without product options</u></div>
    <%end if%>
		<%if TotalXLSlines-ImportedRecords>0 then%>
      <div>A total of <b><font color="#FF0000"><%=TotalXLSlines-ImportedRecords%></font></b> records <u>could NOT</u> be <%if session("append")="1" then%>updated<%else%>imported<%end if%>. See the Error(s) Report section below for details</div>
    <%end if%>
  </div>

	<% if TempProducts<>"" then%> 
	<br /><br />
		<table class="pcCPcontent">
			<tr> 
				<td> 
					<table border="0" cellspacing="0" width="100%" cellpadding="2">
						<tr>
							<th>Invalid Category Report</th>
						</tr>
						<tr>
            <td><p align="justify">The following products could not be assigned to a valid category and	were therefore assigned to a temporary category called	'ImportedProducts'. You can assign these products to any other category in your store by using the 'Manage Categories' feature.</p>						
							</td>
						</tr>
						<tr>
                        	<td>
                                <p align="center">
                                    <textarea name="S1" rows="6" style="font-family: Arial; font-size: 9px; width: 98%;"><%=TempProducts%></textarea>
                                </p>
							</td>
						</tr>
                    </table>
				</td>
			</tr>
		</table>   
		<%end if%>

	<% if ErrorsReport<>"" then%> 
	<br /><br />
	<table class="pcCPcontent">
	<tr> 
		<td> 
			<table border="0" cellspacing="0" width="100%" cellpadding="2">
				<tr>
					<th>
						Error(s) Report
					</th>
				</tr>
                <tr>
                    <td align="center">
                        <div style="width: 98%; height: 150px; overflow: scroll; border: 1px dotted #E1E1E1;">
                            <table id="noheaderodd" style="background: #F5F5F5; font-family: Arial; font-size: 9px; width: 100%;">
                                <%=ErrorsReport%>
                            </table>
                        </div>
                    </td>
                </tr>
			</table> 
		</td>
	</tr>
	</table>
	<%end if%>
  <br /><br />
	<p align="center">
	<input type="button" class="btn btn-default" name=mainmenu value="Back to Main menu" onClick="location='menu.asp';">
	</p>
	</td>
</tr>
</table>
<% If session("importfile")="" Then
	session("TempProducts")=""
	session("ErrorsReport")=""
	session("iPageCurrent")=""
	session("TotalXLSlines")=""
	session("ImportedRecords")=""
	session("PrdWithoutOpts")=""
	session("append")=0
	Session("IDcustom0")=""
	Session("IDcustom1")=""
	Session("IDcustom2")=""
	Session("IDPriceCat0")=""
	Session("IDPriceCat1")=""
	Session("IDPriceCat2")=""
	Session("IDPriceCat3")=""
	Session("IDPriceCat4")=""
end if %>
<!--#include file="AdminFooter.asp"-->