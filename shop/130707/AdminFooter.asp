<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, all of its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit http://www.productcart.com.
%>
        </div>
        
        <% If pcv_strDisplayType <> "1" Then %>
        
        <div id="pcCPmainRight">
        
			<%
            if pcInt_ShowOrderLegend = 1 then
            %>
                <div id="cp5" class="panel panel-default">
                  <div class="panel-heading">Order Status Legend</div>
				  	<div class="panel-body">
                    <ul>
                    	<li><img src="images/bluedot.gif"> Pending</li>
                        <li><img src="images/yellowdot.gif"> Processed</li>
                        <li><img src="images/7dot.gif"> Partially Shipped</li>
                        <li><img src="images/8dot.gif"> Shipping</li>
                        <li><img src="images/greendot.gif"> Shipped</li>
                        <li><img src="images/9dot.gif"> Partially Returned</li>
                        <li><img src="images/orangedot.gif"> Returned</li>
                        <li><img src="images/purpledot.gif"> Incomplete</li>
                        <li><img src="images/reddot.gif"> Cancelled</li>
                    </ul>
                	</div>
				</div>
            
                <div id="cp6" class="panel panel-default">
                  <div class="panel-heading">Payment Status Legend</div>
				  <div class="panel-body">
                    <ul>
                        <li><img src="images/blueflag.gif"> Pending</li>
                        <li><img src="images/yellowflag.gif"> Authorized</li>
                        <li><img src="images/greenflag.gif"> Paid</li>
                        <li><img src="images/darkgreenflag.gif"> Refunded</li>
                        <li><img src="images/redflag.gif"> Voided</li>
                    </ul>
				  </div>
                </div>

			 <%
             end if
             
            IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN
     
             if session("admin")<>"0" and session("admin")<>"" then
             %>
         

            
            	<form name="searchOrdersFooter" action="resultsAdvanced.asp?" class="pcForms">
                    <div id="cp3" class="panel panel-default">
                        <div class="panel-heading">Find an <strong>order</strong> by...</div>
                        <div class="panel-body">
                        <p><select name="TypeSearch" size="1" class="form-control input-sm">
                            <option value="idOrder">Order ID</option>
                            <option value="orderCode">Order Code</option>
                            <% if GOOGLEACTIVE=-1 then %>
                            <option value="GoogleOrderID">Google Order ID</option>
                          <% end if %>
                            <option value="details">Product</option>
                            <option value="shipmentDetails">Shipping Type</option>
                            <option value="stateCode">State/Province Code</option>
                            <option value="CountryCode">Country Code</option>
                        </select></p>
                        <p><input type="text" class="form-control input-sm" name="advquery" placeholder="Enter Value" onFocus="clearText(this)"></p>
                        <p><input type="submit" name="B1" value="Find Orders" class="btn btn-info">
                        <input type="button" class="btn btn-default"  class="btn btn-default"  value="More" onClick="location.href='invoicing.asp'"></p>
                        </div>
                     </div>
				</form>
            	
                <div id="cp1" class="panel panel-default">
                    <div class="panel-heading">Find a <strong>product</strong> by...</div>
                    <div class="panel-body">
                        <form name="ajaxSearchFooter" method="post" action="srcPrds.asp?action=newsrc" class="pcForms">
                            <input type="hidden" name="referpage" value="NewSearch">
                            <input type="hidden" name="src_FormTitle1" value="Find Products">
                            <input type="hidden" name="src_FormTitle2" value="Product Search Results">
                            <input type="hidden" name="src_FormTips1" value="Use the following filters to look for products in your store.">
                            <input type="hidden" name="src_FormTips2" value="">
                            <input type="hidden" name="src_IncNormal" value="0">
                            <input type="hidden" name="src_IncBTO" value="0">
                            <input type="hidden" name="src_IncItem" value="0">
                            <input type="hidden" name="src_DisplayType" value="0">
                            <input type="hidden" name="pinactive" value="-1">
                            <input type="hidden" name="src_ShowLinks" value="1">
                            <input type="hidden" name="src_FromPage" value="LocateProducts.asp">
                            <input type="hidden" name="src_ToPage" value="">
                            <input type="hidden" name="src_Button2" value="Continue">
                            <input type="hidden" name="src_Button3" value="New Search">
                            <p><input name="sku" type="text" class="form-control input-sm" size="6" maxlength="150" placeholder="SKU"></p>
                            <p><input type="text" class="form-control input-sm" name="keyWord" size="10" placeholder="Keyword(s)"></p>
                            <p>
                            <select name="resultCnt" id="resultCnt" class="form-control input-sm">
                                <option value="5" selected>5</option>
                                <option value="10">10</option>
                                <option value="15">15</option>
                                <option value="20">20</option>
                                <option value="25">25</option>
                                <option value="50">50</option>
                                <option value="100">100</option>
                            </select>
                            </p>
                            <p>
                            <input type="hidden" name="act" value="newsrc">
                            <input name="Submit" type="submit" value="Go" class="btn btn-info">
                            <input type="button" class="btn btn-default"  class="btn btn-default"  value="More" onClick="javascript:location.href='LocateProducts.asp';">
                            </p>
                        </form>
                    </div>
                </div>
        
                <!--#include file="smallRecentProducts.asp"--> 
        
            
                <div id="cp4" class="panel panel-default">
                    <div class="panel-heading">Find a <strong>customer</strong> by...</div>
                    <div class="panel-body">
                        <form name="listCustFooter" action="viewCustb.asp" class="pcForms">
                        <p><input type="text" class="form-control input-sm" name="key2" size="14" value="" placeholder="Last Name"></p>
                        <p><input type="text" class="form-control input-sm" name="key3" size="14" value="" placeholder="Company"></p>
                        <p><input type="text" class="form-control input-sm" name="key4" size="14" value="" placeholder="Email"></p>
                        <p><input type="submit" name="srcView" value="Search" class="btn btn-info">
                        <input type="hidden" name="key5" value="">
                        <input type="hidden" name="key6" value="">
                        </p>
                    </form>
                    </div>
                </div>
			<% 
                end if
            END IF
            %>
           <div class="pcCPSpacer"></div>  
		</div>
        
        <% End If %>
        
	</div>
        
    <div id="pcFooter">
        <a href="about_terms.asp"><div style="float: left"><img src="images/pc_logo_100.gif" width="100" height="30" alt="ProductCart shopping cart software" border="0" /></div>Use of this software indicates acceptance of the End User License Agreement</a><br /><a href="http://www.productcart.com">Copyright&copy; 2001-<%=Year(now)%> NetSource Commerce. All Rights Reserved. ProductCart&reg; is a registered trademark of NetSource Commerce</a>.
    </div>

	<script type=text/javascript>
	<%
	tmpStr=""
	IF lcase(section)<>"quickbooks" AND lcase(section)<>"ebay" AND lcase(pageTitle)<>"productcart ebay add-on" THEN
 
	if session("admin")<>"0" and session("admin")<>"" then
	'tmpStr=tmpStr & "$pc( ""#cp1"" ).accordion( ""option"", ""active"", 0 );"
	'tmpStr=tmpStr & "$pc( ""#cp3"" ).accordion( ""option"", ""active"", 0 );"
	'tmpStr=tmpStr & "$pc( ""#cp4"" ).accordion( ""option"", ""active"", 0 );"
	%>
	//$pc( "#cp1" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp1 span").removeClass('ui-icon');
	//$pc( "#cp3" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp3 span").removeClass('ui-icon');
	//$pc( "#cp4" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp4 span").removeClass('ui-icon');
	<%end if
	END IF%>
	<% 
	if pcv_ShowSmallRecentProducts=1 then 
	'tmpStr=tmpStr & "$pc( ""#cp2"" ).accordion( ""option"", ""active"", 0 );"%>
	//$pc( "#cp2" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp2 span").removeClass('ui-icon');
	<% 
	end if 
	%>
	<%
	if pcInt_ShowOrderLegend = 1 then
	'tmpStr=tmpStr & "$pc( ""#cp5"" ).accordion( ""option"", ""active"", 0 );"
	'tmpStr=tmpStr & "$pc( ""#cp6"" ).accordion( ""option"", ""active"", 0 );"
	%>
	//$pc( "#cp5" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp5 span").removeClass('ui-icon');
	//$pc( "#cp6" ).accordion({collapsible: true, header: "h5", active:false});
	//$pc( "#cp6 span").removeClass('ui-icon');
	<%
	end if
	%>
</script>
<script src="https://ajax.googleapis.com/ajax/libs/angularjs/1.2.4/angular.min.js"></script>
<script src="service/app/service.js"></script>
<script src="service/app/apps.js"></script>
<script src="service/app/api.js"></script>
<script src="../includes/javascripts/jquery.blockUI.js"></script>
</body>
</html>