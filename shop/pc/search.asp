<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/validation.asp" -->
<!--#include file="../includes/SearchConstants.asp"-->
<!--#include file="../includes/utilities.asp"-->
<%
'*******************************
' Page Name
'*******************************
Dim pcStrPageName
pcStrPageName = "search.asp"

%>
<!--#include file="pcStartSession.asp"-->
<!--#include file="header_wrapper.asp"-->
<!--#include file="pcValidateHeader.asp"-->

<!--Validate Form-->
<script type=text/javascript>

$pc(document).ready(function() {
	$pc("form[name='ajaxSearch']").find("select").change(function() {
		srcPrdsCount();
	});
	
	$pc("form[name='ajaxSearch']").find("input[type='text'],input[type='number']").blur(function() {
		srcPrdsCount();
	});
	
	$pc("form[name='ajaxSearch']").find("input[type='checkbox'],input[type='radio']").click(function() {
		srcPrdsCount();
	});
});
	
function isDigitA(s)
{
var test=""+s;
if(test==","||test=="."||test=="0"||test=="1"||test=="2"||test=="3"||test=="4"||test=="5"||test=="6"||test=="7"||test=="8"||test=="9")
		{
		return(true) ;
		}
		return(false);
	}
	
function allDigitA(s)
	{
		var test=""+s ;
		for (var k=0; k <test.length; k++)
		{
			var c=test.substring(k,k+1);
			if (isDigitA(c)==false)
			{
				return (false);
			}
		}
		return (true);
	}

function FormValidator(theForm)
{
 	qtt= document.ajaxSearch.priceFrom;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{		    
		    alert("<% = dictLanguage.Item(Session("language")&"_advSrca_26")%>");
		    qtt.focus();
			<% If SRCH_WAITBOX="1" Then %>
				CloseHS();
			<% End If %>
		    return (false);
		    }
	    }
	    
	qtt= document.ajaxSearch.priceUntil;
		if (qtt.value != "")
		{
			if (allDigitA(qtt.value) == false)
			{
		    alert("<% = dictLanguage.Item(Session("language")&"_advSrca_26")%>");
		    qtt.focus();
			<% If SRCH_WAITBOX="1" Then %>
				CloseHS();
			<% End If %>
		    return (false);
		    }
	    }
		return (true);
}
<% If SRCH_WAITBOX="1" Then %>
function CloseHS() 
{
	var t=setTimeout("hs.close('pcMainSearch')",50)
}
function OpenHS() 
{
	document.getElementById('pcMainSearch').onclick()
}
<% End If %>
</script>

<div id="pcMain" class="container-fluid pcSearch">
    <div class="row">
        <div class="col-xs-12">
  
        <h1><%=dictLanguage.Item(Session("language")&"_advSrca_1")%></h1>
        
        <% ' Show search page description, if any
        pcStrSearchDesc = dictLanguage.Item(Session("language")&"_search_1")
        if trim(pcStrSearchDesc) <> "" then %>
            <div class="pcPageDesc"><%=pcStrSearchDesc%></div>
        <% end if %>
        
        <%
        '// Set Submit Action
        Dim pcv_strSubmitAction
        If SRCH_WAITBOX="1" Then
            pcv_strSubmitAction = "OpenHS(); return FormValidator(this);"
        Else
            pcv_strSubmitAction = "return FormValidator(this);"
        End If
        %>       
        <form class="form-horizontal" role="form" name="ajaxSearch" method="get" action="showsearchresults.asp" onSubmit="<%=pcv_strSubmitAction%>">

            <div class="pcFormItem">
          	    <span id="totalresults" class="pcTextMessage">&nbsp;</span>
            </div>
        
            <!--Category Dropdown -->
            <div class="form-group">
            
                <% '// CATEGORY DROP DOWN - START
                select case schideCategory
                    case "0" ' // FIRST scenario: Category drop-down fully shown %>
                        <label for="idcategory" class="col-md-3 control-label"><%=dictLanguage.Item(Session("language")&"_advSrca_2") %></label>
                        <div class="col-md-8">
                                <%
                                cat_DropDownName="idcategory"
                                cat_Type="1"
                                cat_DropDownSize="1"
                                cat_MultiSelect="0"
                                cat_ExcBTOHide="1"
                                cat_StoreFront="1"
                                cat_ShowParent="1"
                                cat_DefaultItem=dictLanguage.Item(Session("language")&"_advSrca_4")
                                cat_SelectedItems="0,"
                                cat_ExcItems=""
                                cat_ExcSubs="0"
                                cat_EventAction=""
                                %>
                                
                                <% call pcs_CatList()%>
                        </div><% 
                    case "1" '// Only top-level categories are shown
                        query="SELECT DISTINCT categories.idCategory,categories.categoryDesc,categories.idParentCategory "
                        query=query&"FROM categories "
                        query=query&"WHERE categories.iBTOhide=0 AND categories.pccats_RetailHide=0 AND idParentCategory=1 AND idCategory<>1 "
                        query=query&"ORDER BY categories.categoryDesc ASC;"
                        
                        set rs=Server.CreateObject("ADODB.Recordset")
                        set rs=connTemp.execute(query)
                        if not rs.eof then
                                Dim categoryArray, categoryCount, categoryTotal
                                categoryArray = rs.getRows()
                                categoryCount = 0
                                categoryTotal = ubound(categoryArray,2)
                                %>
                                <label for="idcategory" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_2")%></label>
                                <div class="col-md-8">
                                        <%
                                            cat_EventAction=""
                                        %>
                                        <select class="form-control" name="idcategory" <%=cat_EventAction%>>
                                        <option value="0" selected><%=dictLanguage.Item(Session("language")&"_advSrca_4")%></option>
                                        <%
                                        do while (categoryCount <= categoryTotal)
                                        %>
                                        <option value="<%=categoryArray(0, categoryCount)%>"><%=categoryArray(1, categoryCount)%></option>
                                        <%
                                        categoryCount = categoryCount + 1
                                        loop
                                        %>
                                        </select>
                                </div>
                        <% end if
                        set rs = nothing	
                    case "-1" '// The category drop-down is hidden %>
                        <input type="hidden" name="idcategory" value="0">
                <% End Select
                '// CATEGORY DROP DOWN - END 
                %>  
                          
            </div>
            <!--End Category Dropdown -->
    
    
    
            <!--Price Range-->
            <div class="form-group">
                <span class="col-md-3 control-label"><%=dictLanguage.Item(Session("language")&"_advSrca_5") %></span>
                <div class="col-md-8">
                    <div class="form-group form-group-nomargin">
                        <label for="priceFrom" class="col-md-2 control-label"><%=dictLanguage.Item(Session("language")&"_advSrca_6") %></label>
                        <div class="col-md-4">
                            <input type="number" class="form-control" min="0" max="999999999" step="1" name="priceFrom" value="0" />
                        </div>
                        <label for="priceUntil" class="col-md-2 control-label"><%=dictLanguage.Item(Session("language")&"_advSrca_7") %></label>
                        <div class="col-md-4">
                            <input name="priceUntil" class="form-control" type="number" min="0" max="999999999" step="1" autocomplete="off" data-hint="" />
                        </div>
                    </div>
                </div>
            </div>
            <!--End Price Range-->
    
    
    
            <!--In Stock Checkbox -->
            <div class="form-group">
                <label for="withstock" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_8")%></label>
                <div class="col-md-8">
                    <input type="checkbox" name="withstock" value="-1" />
                </div>
            </div>
            <!-- -->
    
    
    
            <!--SKU Input-->
            <div class="form-group">
                <label for="SKU" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_11")%></label>
                <div class="col-md-8">
                    <input class="form-control" name="SKU" type="text" maxlength="150" placeholder="" autocomplete="off" data-hint="" />
                </div>
            </div>
            <!-- -->
    
    
    
            <!--Brand Dropdown -->
            <%
            'Show brands, if any
            query="Select IDBrand, BrandName from Brands order by BrandName asc"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=connTemp.execute(query)
            if not rs.eof then
                Dim brandArray, brandCount, brandTotal
                brandArray = rs.getRows()
                brandCount = 0
                brandTotal = ubound(brandArray,2)
                %>
                <div class="form-group">
                    <label for="IDBrand" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_13")%></label>
                    <div class="col-md-8">
                        <select name="IDBrand" class="form-control">
                            <option value="0" selected><%=dictLanguage.Item(Session("language")&"_advSrca_4")%></option>
                            <% do while (brandCount <= brandTotal) %>
                                <option value="<%=brandArray(0, brandCount)%>"><%=brandArray(1, brandCount)%></option>
                                <% brandCount = brandCount + 1
                            loop %>
                        </select>
                    </div>
                </div>
            <% end if
            set rs = nothing %>
            <!--End Brand Dropdown -->
    
    
    
            <!--Keyword-->
            <div class="form-group">
                <label for="keyWord" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_9")%></label>
                <div class="col-md-8">
                    <input type="text" class="form-control" name="keyWord" />
                    <span class="help-block"><input type="checkbox" name="exact" value="1" class="clearBorder" />&nbsp;<% = dictLanguage.Item(Session("language")&"_advSrca_14")%></span>
                </div>
            </div>
            <!--End Keyword-->
    
    
    
            <!-- search custom fields if any are defined -->
            <%tmpJSStr=""
            tmpJSStr=tmpJSStr & "var SFID=new Array();" & vbcrlf
            tmpJSStr=tmpJSStr & "var SFNAME=new Array();" & vbcrlf
            tmpJSStr=tmpJSStr & "var SFVID=new Array();" & vbcrlf
            tmpJSStr=tmpJSStr & "var SFVALUE=new Array();" & vbcrlf
            tmpJSStr=tmpJSStr & "var SFVORDER=new Array();" & vbcrlf
            intCount=-1
            tmpJSStr=tmpJSStr & "var SFCount=" & intCount & ";" & vbcrlf%>
            
            <%query="SELECT idSearchField,pcSearchFieldName,pcSearchFieldShow,pcSearchFieldOrder FROM pcSearchFields WHERE pcSearchFieldSearch=1 ORDER BY pcSearchFieldOrder ASC,pcSearchFieldName ASC;"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=conntemp.execute(query)
            if not rs.eof then
                set pcv_tempFunc = new StringBuilder
                pcv_tempFunc.append "<script type=text/javascript>" & vbcrlf
                pcv_tempFunc.append "function CheckList(cvalue,tmpvalue) {" & vbcrlf
                pcv_tempFunc.append "if (cvalue==0) {" & vbcrlf
                pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
                pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
                pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0"");" & vbcrlf
                pcv_tempFunc.append "SFID=new Array();" & vbcrlf
                pcv_tempFunc.append "SFNAME=new Array();" & vbcrlf
                pcv_tempFunc.append "SFVID=new Array();" & vbcrlf
                pcv_tempFunc.append "SFVALUE=new Array();" & vbcrlf
                pcv_tempFunc.append "SFVORDER=new Array();" & vbcrlf
                intCount=-1
                pcv_tempFunc.append "SFCount=" & intCount & ";" & vbcrlf
                pcv_tempFunc.append "CreateTable(tmpvalue);" & vbcrlf
                pcv_tempFunc.append "}" & vbcrlf
                
                set pcv_tempList = new StringBuilder
                pcv_tempList.append "<select name=""customfield1"" class=""form-control"" onchange=""javascript:CheckList(document.ajaxSearch.customfield1.value,0);"">" & vbcrlf
                pcv_tempList.append "<option value=""0"">All</option>" & vbcrlf
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
                    pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
                    pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
                    pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0"");" & vbcrlf
                    For j=0 to LCount
                        pcv_tempFunc.append "SelectA.options[" & j+1 & "]=new Option(""" & replace(tmpArr(1,j),"""","\""") & """,""" & tmpArr(0,j) & """);" & vbcrlf
                    Next
                    pcv_tempFunc.append "}" & vbcrlf
                else
                    pcv_tempFunc.append "if (cvalue==" & pcArray(0,i) & ") {" & vbcrlf
                    pcv_tempFunc.append "var SelectA = document.ajaxSearch.SearchValues1;" & vbcrlf
                    pcv_tempFunc.append "SelectA.options.length = 0;" & vbcrlf
                    pcv_tempFunc.append "SelectA.options[" & 0 & "]=new Option(""All"",""0""); }" & vbcrlf
                end if
                Next
                
                pcv_tempList.append "</select>" & vbcrlf
                pcv_tempFunc.append "}" & vbcrlf
                pcv_tempFunc.append "</script>" & vbcrlf
                
                pcv_tempList=pcv_tempList.toString
                pcv_tempFunc=pcv_tempFunc.toString
                %>
    
            <div class="form-group">
                <span class="col-md-3 control-label"></span>
                <div class="col-md-8"><% = dictLanguage.Item(Session("language")&"_advSrca_22")%></div>
            </div>
            
            <div class="form-group">
                <span class="col-md-3 control-label"><%=dictLanguage.Item(Session("language")&"_advSrca_12") %></span>
                <div class="col-md-8">
                    <div class="form-group form-group-nomargin">
                        <label for="SearchValues1" class="col-md-4 control-label"><%=pcv_tempList %></label>
                        <div class="col-md-4 control-label">
                            <select name="SearchValues1" class="form-control" onChange="javascript:var testvalue=document.ajaxSearch.customfield.value; if ((testvalue.indexOf('||')==-1) && (document.ajaxSearch.customfield1.value!='0')) {document.ajaxSearch.customfield.value=document.ajaxSearch.customfield1.value;document.ajaxSearch.SearchValues.value=this.value;srcPrdsCount()}">
                            </select>
                        </div>
                        <div class="col-md-2 control-label">
                            <%=pcv_tempFunc %>
                            <a class="pull-left" href="javascript:AddSF(document.ajaxSearch.customfield1.value,document.ajaxSearch.customfield1.options[document.ajaxSearch.customfield1.selectedIndex].text,document.ajaxSearch.SearchValues1.value,document.ajaxSearch.SearchValues1.options[document.ajaxSearch.SearchValues1.selectedIndex].text,0);"><img src="<%=pcf_getImagePath("../pc/images","plus.jpg")%>" alt="Add" border="0"></a>
                        </div>
                    </div>
                </div>
            </div>
    
    
            <input type="hidden" name="customfield" value="0">
            <input type="hidden" name="SearchValues" value="">
                        
                        
            <!--Dynamic Section for Search Filters -->
            <span id="stable" name="stable"></span>
                        
            <script type=text/javascript>
                <%=tmpJSStr%>
                function CreateTable(tmpRun)
                {
                    var tmp1="";
                    var tmp2="";
                    var tmp3="";
                    var i=0;
                    var found=0;
                    tmp1='<div class="pcFormField">';
                    for (var i=0;i<=SFCount;i++)
                    {
                        found=1;
                        tmp1=tmp1 + '<div class="pcFormLabel"><label>&nbsp;</label></div><div class="nsFormCheckbox"><label><a href="javascript:ClearSF(SFID['+i+']);"><img src="<%=pcf_getImagePath("../pc/images","minus.jpg")%>" alt="" border="0"></a><span class="nsFormFieldLabel">&nbsp;'+SFNAME[i]+': '+SFVALUE[i]+'</span></label></div>';
                        if (tmp2=="") tmp2=tmp2 + "||";
                        tmp2=tmp2 + SFID[i] + "||";
                        if (tmp3=="") tmp3=tmp3 + "||";
                        tmp3=tmp3 + SFVID[i] + "||";
                    }
                    tmp1=tmp1+'</div>';
                    if (found==0) tmp1="";
                    document.getElementById("stable").innerHTML=tmp1;
                    if (tmp2=="") tmp2=0;
                    document.ajaxSearch.customfield.value=tmp2;
                    document.ajaxSearch.SearchValues.value=tmp3;
                    if (tmp2==0)
                    {
                        document.ajaxSearch.customfield.value=document.ajaxSearch.customfield1.value;
                        document.ajaxSearch.SearchValues.value=document.ajaxSearch.SearchValues1.value;
                    }
                    if (tmpRun!=1) srcPrdsCount();
                }
                
                CheckList(document.ajaxSearch.customfield1.value,1);
                
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
                    CreateTable(0);
                }
                
                function AddSF(tmpSFID,tmpSFName,tmpSVID,tmpSValue,tmpSOrder)
                {
                    if ((tmpSVID!="") && (tmpSFID!="") && (tmpSVID!="0") && (tmpSFID!="0"))
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
                        CreateTable(0);
                    }
                }
        </script>
        <!--End Dynamic Section for Search Filters -->
    
        <%	pcv_HaveSearchFields=1
        else %>
            <input type="hidden" name="customfield" value="0">
        <% end if %>
        <!-- end of custom fields -->			
    
    
    
        <!--Currently On Sale-->
        <% 'SM-Start
        if UCase(scDB)="SQL" then
            tmpTargetType=0
            if session("customerCategory")<>"" AND session("customerCategory")<>"0" then
                tmpTargetType=session("customerCategory")
            else
                if session("customerType")="1" then
                    tmpTargetType="-1"
                end if
            end if
                                    
            query="SELECT pcSales_Completed.pcSC_ID ,pcSales_Completed.pcSC_SaveName FROM pcSales_Completed INNER JOIN pcSales ON pcSales_Completed.pcSales_ID=pcSales.pcSales_ID WHERE pcSales_Completed.pcSC_Status=2 AND pcSales.pcSales_TargetPrice=" & tmpTargetType & " AND pcSales_Completed.pcSC_Archived=0 ORDER BY pcSC_SaveName ASC;"
            set rs=Server.CreateObject("ADODB.Recordset")
            set rs=connTemp.execute(query)
            if not rs.eof then
                saleArr=rs.getRows()
                intSale=ubound(saleArr,2)
                %>		
                <div class="form-group">
                    <label for="incSale" class="col-md-3 control-label"><%=dictLanguage.Item(Session("language")&"_SaleSearch_1") %></label>
                    <div class="col-md-1">
                        <input type="checkbox" name="incSale" value="1" />
                    </div>
                </div>
                <div class="form-group">
                    <label for="IDSale" class="col-md-3 control-label"><%=dictLanguage.Item(Session("language")&"_SaleSearch_2") %></label>
                    <div class="col-md-8">
                        <select class="form-control" name="IDSale">
                            <option value="0" selected><%=dictLanguage.Item(Session("language")&"_SaleSearch_3")%></option>
                            <% For k=0 to intSale %>
                                <option value="<%=saleArr(0,k)%>"><%=saleArr(1,k)%></option>
                            <% Next %>
                        </select>
                    </div>
                </div>
    
            <% end if
            set rs = nothing
        end if
        'SM-End %>
        <!--End Currently On Sale-->
    
    
    
        <!--Results Per Page Dropdown -->
        <% '// Locate preferred results count and load as default
        Dim pcIntPreferredCount
        pcIntPreferredCount =(scPrdRow*scPrdRowsPerPage)
        if validNum(pcIntPreferredCount) then %>				
            <div class="form-group">
                <label for="resultCnt" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_15")%></label>
                <div class="col-md-2">
                    <select class="form-control" name="resultCnt" id="resultCnt">
                        <option value="<%=pcIntPreferredCount%>" selected><%=pcIntPreferredCount%></option>
                        <option value="<%=pcIntPreferredCount*2%>"><%=pcIntPreferredCount*2%></option>
                        <option value="<%=pcIntPreferredCount*3%>"><%=pcIntPreferredCount*3%></option>
                        <option value="<%=pcIntPreferredCount*4%>"><%=pcIntPreferredCount*4%></option>
                        <option value="<%=pcIntPreferredCount*5%>"><%=pcIntPreferredCount*5%></option>
                        <option value="<%=pcIntPreferredCount*10%>"><%=pcIntPreferredCount*10%></option>
                    </select>
                </div>
            </div>
        <% end if %>
        <!--End Results Per Page Dropdown -->
                    
    
    
        <!--Sort by Dropdown -->
        <div class="form-group">
            <label for="resultCnt" class="col-md-3 control-label"><% = dictLanguage.Item(Session("language")&"_advSrca_16")%></label>
            <div class="col-md-2">
                <select class="form-control" name="order">
                    <option value="0"<% if PCOrd=0 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_18")%></option>
                    <option value="1"<% if PCOrd=1 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_19")%></option>
                    <option value="3"<% if PCOrd=3 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_20")%></option>
                    <option value="2"<% if PCOrd=2 then %> selected<% end if %>><%=dictLanguage.Item(Session("language")&"_advSrca_21")%></option>
                </select>
            </div>
        </div>
        <!--End Sort by Dropdown -->
    
        <div class="form-group">
            <div class="col-md-offset-3 col-md-9">
                <button class="pcButton pcButtonSearch" id="Submit" name="Submit">
                    <img src="<%=pcf_getImagePath("",rslayout("submit"))%>" alt="<% = dictLanguage.Item(Session("language")&"_advSrca_10")%>" />
                    <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_pcLO_update") %></span>
                </button>
            </div>
        </div>

    </form>   
    
    <% If SRCH_WAITBOX="1" Then
        '// Loading Window
        '	>> Call Method with OpenHS();
        response.Write(pcf_ModalWindow(dictLanguage.Item(Session("language")&"_advSrca_23"), "pcMainSearch", 200))
    End If %>
        </div>
    </div>
</div>
<%if pcv_HaveSearchFields=1 then%>
	<script type=text/javascript>CreateTable(1);</script>
<%end if%>
<% set rstemp= nothing %>

<!--#include file="footer_wrapper.asp"-->
