<%@ LANGUAGE="VBSCRIPT" %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=3%><!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<% pageTitle = "Reward Points - Assign Points to Multiple Products" %>
<% Section = "specials" %>
<%
If request.form("Submit")<>"" then
	ieqtype=request.form("eqtype")
	if ieqtype="" then
		ieqtype="1"
	end if
	
	if ieqtype="2" then
		imultiplier=request.form("multiplier")
		
		if scDecSign="," then
			imultiplier=replace(imultiplier,".","")
			imultiplier=replace(imultiplier,",",".")
		else
			imultiplier=replace(imultiplier,",","")
		end if

		if imultiplier="" then
			call closeDb()
			response.redirect "RpProducts.asp?msg="&server.URLEncode("You must insert a multiplier into the form to use the multiplier option.")
		end if
	end if

	'// Get Category List
	pcv_IdCategory=request("idcategory")	
	if pcv_IdCategory="" then
		pcv_IdCategory=0
	end if
	pcv_IdCategory=trim(pcv_IdCategory)	
	pcList1=split(pcv_IdCategory,",")
	if pcList1(0)="0" then
		pcList1=""
	Else
		pcList1=pcv_IdCategory
	End if
	
	if pcList1<>"" then
		query1=" Categories_Products.idCategory IN (" & pcv_IdCategory & ") AND "
	else
		query1=""
	end if
	
	if statusAPP="1" then
		query2=" OR (Categories_Products.idProduct=Products.pcProd_ParentPrd) "
	else
		query2=""
	end if
	
	if ieqtype="1" then
		query3=" products.iRewardPoints=Round(products.price,0) "
	else
		query3=" products.iRewardPoints=Round(products.price*" & imultiplier & ",0) "
	end if
	
	query="UPDATE Products SET " & query3 & " FROM Products, Categories_Products WHERE " & query1 & " ((Categories_Products.idProduct=Products.idProduct) " & query2 & ");"
	set rs = Server.CreateObject("ADODB.Recordset")     
	set rs=conntemp.execute(query)
	set rs=nothing
	
	call closeDb()
	response.redirect "RpProducts.asp?mode=3"
End If %>
<!--#include file="AdminHeader.asp"-->
<% if request.QueryString("mode")="3" then %>
	<table class="pcCPcontent">
		<tr> 
			<td align="center"><div class="pcCPmessageSuccess">You have successfully assigned points to the selected products. <a href="rpProducts.asp">Back</a>.</div></td>
		</tr>
	</table>
<% else %>
	<script type=text/javascript>
	function Form1_Validator(theForm)
	{
	
	if (theForm.buttoncheck.value == "0")
			{
				alert("Please select an option before submitting!");
					theForm.multiplier.focus();
					return (false);
		}
		
	return (true);
	}
	</script>

	<form name="form1" method="post" action="RpProducts.asp" onSubmit="return Form1_Validator(this)" class="pcForms">
		<table class="pcCPcontent">     
			<tr>       
				<td>Use the settings below to assign Reward Points to multiple products. Note that this action cannot be undone, but you can still edit the number of points on a product by product basis (or import that information using the Import Wizard).</td>
			</tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
            <tr>
                <th>Filter products by category</th>
            </tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
            <tr>
            <td>
                Please <strong>select one or more categories</strong>: all products assigned to those categories will be affected, regardless of whether they are also assigned to other categories that you do not select. Press down the CTRL key on your keyboard to select multiple categories.<br>
                <br>
                <%
                cat_DropDownName="idcategory"
                cat_Type="1"
                cat_DropDownSize="10"
                cat_MultiSelect="1"
                cat_ExcBTOHide="1"
                cat_StoreFront="0"
                cat_ShowParent="1"
                cat_DefaultItem="All categories"
                cat_SelectedItems="0,"
                cat_ExcItems=""
                cat_ExcSubs="0"
                cat_ExcBTOItems="1"
                cat_EventAction=""
                %>
               
                <%call pcs_CatList()%>							
                </td>
            </tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
			<tr> 
				<th>Indicate how points will be calculated</th>
			</tr>
            <tr>
                <td class="pcCPspacer"></td>
            </tr>
			<tr> 
				<td><input type="radio" name="eqtype" value="1" onClick="document.form1.buttoncheck.value='1';" class="clearBorder"> Points = Price</td>
			</tr>
			<tr> 
				<td>For example, if the price for a product is $28, that product will be assigned 28 points.</td>
			</tr>
			<tr> 
				<td class="pcCPspacer"></td>
			</tr>
			<tr> 
				<td><input type="radio" name="eqtype" value="2" onClick="document.form1.buttoncheck.value='1';" class="clearBorder"> Points = Price * Multiplier <span style="padding-left: 25px;">Multiplier: <input name="multiplier" type="text" size="10" value="10"> (e.g. 10)</span></td>
			</tr>
			<tr> 									
				<td>For example, if the price for a product is $28 and the multiplier is 10, that product will be assigned 280 points.</td>
			</tr>	
			<tr> 
				<td><hr></td>
			</tr>
			<tr> 
				<td align="center">
				<input type="hidden" name="buttoncheck" value="0">
				<input type="submit" name="Submit" value="Assign Reward Points to Products" class="btn btn-primary">&nbsp;                
				<input type="button" class="btn btn-default"  name="back" value="Back" onClick="document.location.href='rpstart.asp';">
				</td>
			</tr>
		</table>
	</form>
<% end if %>
<!--#include file="Adminfooter.asp"-->