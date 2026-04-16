<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% pageTitle="Custom Product Fields - Summary" %>
<% nav=request("nav")
if nav="bto" then
	Section="services"
else
	Section="products"
end if %>
<%PmAdmin=2%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="inc_UpdateDates.asp" -->
<!--#include file="../includes/utilities.asp"-->
<!--#include file="AdminHeader.asp"-->
<%
idproduct=request("idproduct")

IF request("action")="updvalue" THEN
	SFData=request("SFData")
	query="DELETE FROM pcSearchFields_Products WHERE idproduct=" & idproduct & ";"
	set rsQ=connTemp.execute(query)
	set rsQ=nothing
	if SFData<>"" then
		tmp1=split(SFData,"||")
		For i=0 to ubound(tmp1)
			if tmp1(i)<>"" then
				tmp2=split(tmp1(i),"^^^")
				if tmp2(2)="-1" then
					query="SELECT idSearchData FROM pcSearchData WHERE idSearchField=" & tmp2(1) & " AND pcSearchDataName like '" & tmp2(3) & "';"
					set rsQ=connTemp.execute(query)
					if not rsQ.eof then
						query="UPDATE pcSearchData SET idSearchField=" & tmp2(1) & ",pcSearchDataName=N'" & tmp2(3) & "',pcSearchDataOrder=" & tmp2(4) & " WHERE idSearchField=" & tmp2(1) & " AND pcSearchDataName like '" & tmp2(3) & "';"
						set rsQ=connTemp.execute(query)
					else
						query="INSERT INTO pcSearchData (idSearchField,pcSearchDataName,pcSearchDataOrder) VALUES (" & tmp2(1) & ",N'" & tmp2(3) & "'," & tmp2(4) & ");"
						set rsQ=connTemp.execute(query)
					end if
					set rsQ=nothing

					query="SELECT idSearchData FROM pcSearchData WHERE pcSearchDataName like '" & tmp2(3) & "';"
					set rsQ=connTemp.execute(query)
					idSearchData=rsQ("idSearchData")
					set rsQ=nothing
				else
					idSearchData=tmp2(2)
				end if
				query="INSERT INTO pcSearchFields_Products (idproduct,idSearchData) VALUES (" & idproduct & "," & idSearchData & ");"
				set rsQ=connTemp.execute(query)
				set rsQ=nothing
			end if
		Next
	end if
	
	call updPrdEditedDate(idproduct)
	
	msg="1"
END IF

%>
<table class="pcCPcontent">
	<tr>
		<td colspan="4">
        <div style="float: right; padding-top: 8px;"><span class="pcSmallText"><a href="FindProductType.asp?id=<%=idproduct%>">Edit</a> | <a href="../pc/viewPrd.asp?idproduct=<%=idproduct%>&adminPreview=1" target="_blank">Preview</a></span></div>
        <h2>Product: <strong><%=productName%></strong></h2>
		<p>You can add custom fields to a product to collect or display additional product information. ProductCart supports two types of custom fields (consult the ProductCart User Guide for more details):</p>
		<ul>
			<li><u>Input fields</u> allow you to collect information from the customer (e.g. name to be embroidered on the front of a polo shirt)&nbsp;<a href="http://wiki.productcart.com/productcart/input_fields_manage" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this topic" width="16" height="16" border="0"></a></li>
			<li><u>Search fields</u> allow you to add searchable properties to products (e.g. wine store: year, wine region, wine type, etc.)&nbsp;<a href="http://wiki.productcart.com/productcart/managing_search_fields" target="_blank"><img src="images/pcv3_infoIcon.gif" alt="More information on this topic" width="16" height="16" border="0"></a></li>
		</ul>
		</td>
	</tr>
	<%if msg="1" then%>
	<tr>
		<td colspan="4">
			<div class="pcCPmessageSuccess">Search fields were updated successfully!</div>
		</td>
	</tr>
	<%end if%>
	<tr>
		<th colspan="4">Custom Search Fields</th>
	</tr>
	<tr>
		<td colspan="4" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="4">
			<%query="SELECT pcSearchFields.idSearchField,pcSearchFields.pcSearchFieldName,pcSearchData.idSearchData,pcSearchData.pcSearchDataName,pcSearchData.pcSearchDataOrder FROM pcSearchFields INNER JOIN (pcSearchData INNER JOIN pcSearchFields_Products ON pcSearchData.idSearchData=pcSearchFields_Products.idSearchData) ON pcSearchFields.idSearchField=pcSearchData.idSearchField WHERE pcSearchFields_Products.idproduct=" & idproduct & ";"
			set rs=connTemp.execute(query)
			tmpJSStr=""
			tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
			tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
			tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
			tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
			tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
			intCount=-1
			IF not rs.eof THEN
				pcArr=rs.getRows()
				set rs=nothing
				intCount=ubound(pcArr,2)
				For i=0 to intCount
					tmpJSStr=tmpJSStr & "SFID[" & i & "]=" & pcArr(0,i) & ";" & vbcrlf
					tmpJSStr=tmpJSStr & "SFNAME[" & i & "]='" & replace(pcArr(1,i),"'","\'") & "';" & vbcrlf
					tmpJSStr=tmpJSStr & "SFVID[" & i & "]=" & pcArr(2,i) & ";" & vbcrlf
					tmpJSStr=tmpJSStr & "SFVALUE[" & i & "]='" & replace(pcArr(3,i),"'","\'") & "';" & vbcrlf
					tmpJSStr=tmpJSStr & "SFVORDER[" & i & "]=" & pcArr(4,i) & ";" & vbcrlf
				Next
			END IF
			set rs=nothing
			tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>
				<script type=text/javascript>
					<%=tmpJSStr%>
					function CreateTable()
					{
						var tmp1="";
						var tmp2="";
						var i=0;
						var found=0;
						tmp1='<table class="pcCPcontent"><tr><td colspan=2 nowrap><strong>Current Search Fields</strong></td><td nowrap><strong>Current Value</strong></td></tr>';
						for (var i=0;i<=SFCount;i++)
						{
							found=1;
							tmp1=tmp1 + '<tr><td align="right"><a href="javascript:ClearSF(SFID['+i+']);"><img src="../pc/images/minus.jpg" alt="Remove" border="0"></a></td><td width="275" nowrap>'+SFNAME[i]+'</td><td width="100%">'+SFVALUE[i]+'</td></tr>';
							if (tmp2=="") tmp2=tmp2 + "||";
							tmp2=tmp2 + "^^^" + SFID[i] + "^^^" + SFVID[i] + "^^^" + SFVALUE[i] + "^^^" + SFVORDER[i] + "^^^||"
						}
						tmp1=tmp1+'</table>';
						if (found==0) tmp1="<br><b>No search fields are assigned to this product</b><br><br>";
						document.getElementById("stable").innerHTML=tmp1;
						document.ajaxSearch.SFData.value=tmp2;
					}
					function ClearSF(tmpSFID)
					{
						var i=0;
						for (var i=0;i<=SFCount;i++)
						{
							if (SFID[i]==tmpSFID)
							{
								removedArr = SFID.splice(i,1);
								removedArr = SFNAME.splice(i,1);
								removedArr = SFVID.splice(i,1);
								removedArr = SFVALUE.splice(i,1);
								removedArr = SFVORDER.splice(i,1);
								SFCount--;
								break;
							}
						}
						CreateTable();
					}
					
					function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
					{
						if (tmpSValue!="")
						{
							var i=0;
							var found=0;
							for (var i=0;i<=SFCount;i++)
							{
								if (SFID[i]==tmpSFID)
								{
									SFVID[i]=tmpSVID;
									SFVALUE[i]=tmpSValue;
									SFVORDER[i]=tmpSOrder;
									found=1;
									break;
								}
							}
							if (found==0)
							{
								SFCount++;
								SFID[SFCount]=tmpSFID;
								SFNAME[SFCount]=tmpSFName;
								SFVID[SFCount]=tmpSVID;
								SFVALUE[SFCount]=tmpSValue;
								SFVORDER[SFCount]=tmpSOrder;
							}
							CreateTable();
						}
					}
				</script>
				<span id="stable" name="stable"></span>
				<%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
				set rs=Server.CreateObject("ADODB.Recordset")
				set rs=conntemp.execute(query)
				if not rs.eof then
					set pcv_tempFunc = new StringBuilder
					pcv_tempFunc.append "<script type=text/javascript>" & vbcrlf
					pcv_tempFunc.append "function CheckList(cvalue) {" & vbcrlf
					pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
					pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
					pcv_tempFunc.append "SelectA.options.length = 0; }" & vbcrlf
					
					set pcv_tempList = new StringBuilder
					pcv_tempList.append "<select name=""customfield"" onchange=""javascript:document.ajaxSearch.newvalue.value='';document.ajaxSearch.neworder.value='0';CheckList(document.ajaxSearch.customfield.value);"">" & vbcrlf
					
					pcArray=rs.getRows()
					intCount=ubound(pcArray,2)
					set rs=nothing
					
					For i=0 to intCount
						pcv_tempList.append "<option value=""" & pcArray(0,i) & """>" & replace(pcArray(1,i),"""","&quot;") & "</option>" & vbcrlf
						query="SELECT idSearchData,pcSearchDataName FROM pcSearchData WHERE idSearchField=" & pcArray(0,i) & " ORDER BY pcSearchDataOrder ASC,pcSearchDataName ASC;"
						set rs=connTemp.execute(query)
						if not rs.eof then
							tmpArr=rs.getRows()
							LCount=ubound(tmpArr,2)
							pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
							pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
							pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
							For j=0 to LCount
								pcv_tempFunc.append "SelectA.options[" & j & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
							Next
							pcv_tempFunc.append "}" & vbcrlf
						else
							pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
							pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues;" & vbcrlf
							pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
							pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option("""",""""); }" & vbcrlf
						end if
					Next
			
					pcv_tempList.append "</select>" & vbcrlf
					pcv_tempFunc.append "}" & vbcrlf
					pcv_tempFunc.append "</script>" & vbcrlf
					
					pcv_tempList=pcv_tempList.toString
					pcv_tempFunc=pcv_tempFunc.toString
					%>
					<hr>
					<Form action="AdminCustom.asp?action=updvalue" method="post" name="ajaxSearch" class="pcForms">
					<table class="pcCPcontent" style="width:auto;">
						<tr>
							<td colspan="2"><a name="2"></a><b>Add new search field values to this product</b></td>
						</tr>
						<tr>
							<td width="20%">Custom Field:</td>
							<td width="80%">
							<%=pcv_tempList%>&nbsp;Value:&nbsp;
							<select name="SearchValues">
							</select>
							<%=pcv_tempFunc%>
							<script type=text/javascript>
								CheckList(document.ajaxSearch.customfield.value);
							</script>
							&nbsp;<a href="javascript:AddSF(document.ajaxSearch.customfield.value,document.ajaxSearch.customfield.options[document.ajaxSearch.customfield.selectedIndex].text,document.ajaxSearch.SearchValues.value,document.ajaxSearch.SearchValues.options[document.ajaxSearch.SearchValues.selectedIndex].text,0);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
							</td>
						</tr>
						<tr>
							<td>New Value:</td>
							<td>
								<input type="text" value="" name="newvalue" size="30">&nbsp;&nbsp;Order: <input type="text" value="0" name="neworder" size="3">
								&nbsp;<a href="javascript:AddSF(document.ajaxSearch.customfield.value,document.ajaxSearch.customfield.options[document.ajaxSearch.customfield.selectedIndex].text,-1,document.ajaxSearch.newvalue.value,document.ajaxSearch.neworder.value);"><img src="../pc/images/plus.jpg" alt="Add" border="0"></a>
							</td>
						</tr>
						<tr>
							<td colspan="2">
							<em><b><u>Note:</u></b> All adjustments will be affected only when you click on the "Update Product" button below.</em>
							</td>
						</tr>
                        <tr>
                        	<td colspan="2" class="pcCPspacer"><hr></td>
                        </tr>
						<tr>
							<td colspan="2">
							<input type="hidden" name="SFData" value="">
							<input type="hidden" name="idproduct" value="<%=idproduct%>">
							<input type="submit" name="submit" value="Update Product &amp; Save Changes" class="btn btn-primary">
							</td>
						</tr>
						</table>
					</Form>
				<%else%>
					<a href="ManageSearchFields.asp">Click here</a> to add new product custom search field.</a>
				<%end if%>
				
				<script type=text/javascript>CreateTable();</script>
		
		</td>
	</tr>
</table>
<br />
<table class="pcCPcontent">
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>                   
		<th colspan="3">Custom Input Fields</th>
	</tr>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<%query="SELECT XFields.IdXField,XFields.XField FROM XFields INNER JOIN pcPrdXFields ON XFields.IdXField=pcPrdXFields.IdXField WHERE pcPrdXFields.IdProduct=" & idproduct & ";"
	set rs=ConnTemp.execute(query)
	IF rs.eof THEN
	set rs=nothing%>
	<tr> 
		<td colspan="2">
			<strong>No custom input fields are assigned to this product</strong><br>
			<a href="addCFtoPrds.asp?idproduct=<%=idproduct%>">Click here</a> to add new product custom input field.
		</td>
		<td></td>
	</tr>
	<%ELSE
		pcArr=rs.getRows()
		set rs=nothing
		intCount=ubound(pcArr,2)
		For i=0 to intCount%>
		<tr> 
			<td width="75%" colspan="2"><%=pcArr(1,i)%></td>
			<td width="25%" align="right"> 
				<a href="modCustomFields.asp?nav=<%=nav%>&idxfield=<%=pcArr(0,i)%>&idproduct=<%=idproduct%>">Edit</a> | <a href="JavaScript:if(confirm('Are you sure that you want to remove this custom input field from this product?')) location='removecustomfields.asp?nav=<%=nav%>&type=2&idxfield=<%=pcArr(0,i)%>&idproduct=<%=idproduct%>'">Remove</a>
			</td>
		</tr>
		<%Next%>
		<tr>
			<td colspan="3" class="pcCPspacer"></td>
		</tr>
		<tr>
			<td colspan="3" class="pcCPspacer"><a href="addCFtoPrds.asp?idproduct=<%=idproduct%>">Click here</a> to add new product custom input field.</td>
		</tr>
	<%END IF%>
	<tr>
		<td colspan="3" class="pcCPspacer"></td>
	</tr>
	<tr>
		<td colspan="3"><hr></td>
	</tr>
	<tr>
		<td colspan="3">
			<p align="center">
			<form class="pcForms">
			<input type="button" class="btn btn-default"  value="Copy to Another Product" onClick="location.href='ApplyCustomFields1.asp?nav=&idproduct=<%=idproduct%>'">&nbsp;
			<input type="button" class="btn btn-default"  value="Locate Another Product" onClick="location.href='LocateProducts.asp?cptype=0'">
			</form>
			</p>
		</td>
	</tr>
</table>

<%
set rs=nothing
%><!--#include file="Adminfooter.asp"-->
