<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%
if len(pcv_intIDMain)>0 then
  query="SELECT * FROM SB_Packages WHERE idProduct="& pcv_intIDMain 
  Set rstemp=conntemp.execute(query)
  If NOT rstemp.eof Then
	  pcv_strIsLinked=rstemp("SB_IsLinked")
	  pcv_CurrencyCode=rstemp("SB_CurrencyCode")
	  pcv_IsTrial=rstemp("SB_IsTrial")
	  pcv_StartsImmediately=rstemp("SB_StartsImmediately")
	  pcv_StartDate=rstemp("SB_StartDate")
	  pcv_Price=rstemp("SB_Amount")
	  pcv_TrialPrice=rstemp("SB_TrialAmount")
	  pcv_RefName=rstemp("SB_RefName")
	  pcv_BillingPeriod=rstemp("SB_BillingPeriod")	  
	  pcv_BillingFrequency=rstemp("SB_BillingFrequency")
	  pcv_LinkID=rstemp("SB_LinkID")
	  pcv_BillingCycles=rstemp("SB_BillingCycles")
	  pcv_TrialBillingPeriod=rstemp("SB_TrialBillingPeriod")
	  pcv_TrialBillingFrequency=rstemp("SB_TrialBillingFrequency")
	  pcv_TrialBillingCycles=rstemp("SB_TrialBillingCycles")  	  
  End If
  Set rstemp=nothing
end if
%>
<table class="pcCPcontent">
	<tr>
		<td>
        	<input name="IsLinked" type="radio" value="1" onclick="openLinked();" <%if pcv_strIsLinked="1" then response.write("checked") %>  style="visibility:hidden">
            Select the SubscriptionBridge &quot;package&quot; to link your product/service to: 
            &nbsp;&nbsp;
            <%			
			'// API Credentials
			query="SELECT Setting_APIUser,Setting_APIPassword,Setting_APIKey,Setting_RegSuccess FROM SB_Settings;"
			set rs=connTemp.execute(query)
			if not rs.eof then
				Setting_APIUser=rs("Setting_APIUser")
				Setting_APIPassword=enDeCrypt(rs("Setting_APIPassword"), scCrypPass)
				Setting_APIKey=enDeCrypt(rs("Setting_APIKey"), scCrypPass)
			end if
			set rs=nothing

			Dim objSB 
			Set objSB = NEW pcARBClass
			result=objSB.GetPackages(Setting_APIUser, Setting_APIPassword, Setting_APIKey)	

			if result<>"0" AND SB_ErrMsg="" then
				Set myXmlDoc = Server.CreateObject("Msxml2.DOMDocument"&scXML)				 
				myXmlDoc.loadXml(result)
				%>
                <select name="LinkedPackage">
                <option value=""> -- select a package -- </option>
				<%
				Set Nodes = myXmlDoc.selectnodes("//GetPackagesResponse/Packages/Package")	
				For Each Node In Nodes	
					strLinkID = objSB.pcf_CheckNode(Node,"LinkID","")
					strPackageName = objSB.pcf_CheckNode(Node,"PackageName","")
					%><option value="<%=strLinkID%>" <%if cstr(pcv_LinkID)=cstr(strLinkID) then response.Write("selected")%>><%=strPackageName%> (<%=strLinkID%>)</option><%			
				Next	
				%>
                </select>
            <% else %>
            	<!--
                <span class="pcCPnotes">
                	We could not locate any SB packages.
                </span> 
                -->
            <% end if %>
            <% If strLinkID="" Then %>
				<div class="pcCPnotes"><strong>NOTE:</strong> You have no available packages. Make sure that you have packages in your SubscriptionBridge Merchant Center that are available for purchase (active, descriptions entered).</div>
			<% End If %>
        </td>
	</tr>
	<tr>
		<td class="pcCPspacer"></td>
	</tr>
 	<tr>
		<td>
            <div id="tLinkedProduct" style="display:none">
                <% if result<>"0" AND SB_ErrMsg="" then %>
                  
                  <div style="padding-bottom: 10px;">
                      ProductCart will use the Online Price for the package price and the plan details saved in your Subscription Bridge account.
                      If your package has a trial please fill out the following details:
                  </div>
                  Trial Settings
               
                    <div id="pIsTrial" style="margin-top: 8px;"> 
                        This package includes a trial:&nbsp;
                        <input id="LinkedIsTrial" name="LinkedIsTrial" type="checkbox" value="1" <%if pcv_IsTrial="1" then response.Write("checked")%>>
                    </div>
                    
                    <div id="pLinkedTrialPrice" style="margin-top: 6px;"> 
                    	The Trial Price is:&nbsp;
                        <input id="LinkedTrialPrice" name="LinkedTrialPrice" type="input" value="<%=pcv_TrialPrice%>">
                        &nbsp;
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                                     

                <% else %>
                
                  <div class="pcCPmessage">
               	  		This option is only available if you have pre configured at least one package in the SB Merchant Center. Please <a href="https://www.subscriptionbridge.com/merchantcenter" target="_blank">login to your SB account</a> to complete this task. 
                        <br />
                        <br />
                        If you have already configured a package there may be a problem:
                       	<ul>
                        	<li>Your activation credentials may be incorrect.  <a href="sb_manageAcc.asp">Review your credentials</a>.</li>
                            <li>Subscription Bridge API may be down for maintenance. Try again later.</li>
                        </ul>
                  </div>
                  
                <% end if %>
            </div>	
        </td>
	</tr> 
	<!--	
    <tr>
		<td class="pcCPspacer"><img src="images/pc_admin.gif" width="85" height="19"></td>
	</tr>
	<tr>
		<td>
        	<input name="IsLinked" type="radio" value="0" onclick="closeLinked();" <%if pcv_strIsLinked="0" then response.write("checked") %>>
        	<strong>No Thanks.</strong>  I will type my details in the form below. I understand that feature configuration will be disabled in the Customer Center:  
        </td>
	</tr> 
    -->
    <input name="IsLinked" type="radio" value="0" onclick="closeLinked();" <%if pcv_strIsLinked="0" then response.write("checked") %> style="visibility:hidden">
	<tr>
		<td>
            <div id="tNOTLinkedProduct" style="display:none">
                
					Package Settings

                    <!--
                    strCurrencyCode=request("CurrencyCode")
                    intIsTrial=request("IsTrial")
                    intStartsImmediately=request("StartsImmediately")
                    strStartDate=request("StartDate")  
                   
                    
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Price:  </b>  
                    <div id="pPrice"> 
                        <input id="Price" name="Price" type="input" style="width: 350px;" value="<%=pcv_Price%>">
                        &nbsp;
                        <img src="SubscriptionBridge/images/ok.gif" title="Valid" alt="Valid" class="validMsg" border="0"/>
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                    -->

                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Reference Name:  </b>  
                    <div id="pRefName"> 
                        <input id="RefName" name="RefName" type="input" style="width: 350px;" value="<%=pcv_RefName%>">
                        <span class="textfieldRequiredMsg">
                            *												
						</span>
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Billing Period:  </b>  
                    <div id="pBillingPeriod"> 
                        <select id="BillingPeriod" name="BillingPeriod">
                        	<option value="">Please Select One</option>
                            <option value="Day" <%if pcv_TrialBillingPeriod="Day" then response.Write("selected")%>>Day</option>
                            <option value="Week" <%if pcv_TrialBillingPeriod="Week" then response.Write("selected")%>>Week</option>
                            <option value="Month" <%if pcv_TrialBillingPeriod="Month" then response.Write("selected")%>>Month</option>
                            <option value="Year" <%if pcv_TrialBillingPeriod="Year" then response.Write("selected")%>>Year</option>
                        </select>
                        <span class="selectRequiredMsg">
                            *												
                        </span>	
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Billing Frequency:  </b>  
                    <div id="pBillingFrequency"> 
                        <input id="BillingFrequency" name="BillingFrequency" type="input" style="width: 350px;" value="<%=pcv_BillingFrequency%>">
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Billing Cycles:  </b>  
                    <div id="pBillingCycles"> 
                        <input id="BillingCycles" name="BillingCycles" type="input" style="width: 350px;" value="<%=pcv_BillingCycles%>">
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                    
                    <br />                    
                    <div id="pIsTrial"> 
                        <b>Check the box if this is a trial:  </b>
                        <input id="IsTrial2" name="IsTrial2" type="checkbox" value="1" <%if pcv_IsTrial="1" then response.Write("checked")%>>
                    </div>
                    
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Trial Price:  </b>  
                    <div id="pTrialPrice"> 
                        <input id="TrialPrice" name="TrialPrice" type="input" style="width: 350px;" value="<%=pcv_TrialPrice%>">
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Trial Billing Period:  </b>  
                    <div id="pTrialBillingPeriod"> 
                        <select id="TrialBillingPeriod" name="TrialBillingPeriod">
                        	<option value="">Please Select One</option>
                            <option value="Day" <%if pcv_BillingPeriod="Day" then response.Write("selected")%>>Day</option>
                            <option value="Week" <%if pcv_BillingPeriod="Week" then response.Write("selected")%>>Week</option>
                            <option value="Month" <%if pcv_BillingPeriod="Month" then response.Write("selected")%>>Month</option>
                            <option value="Year" <%if pcv_BillingPeriod="Year" then response.Write("selected")%>>Year</option>
                        </select>
                        <span class="selectRequiredMsg">
                            *												
                        </span>	
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Trial Billing Frequency:  </b>  
                    <div id="pTrialBillingFrequency"> 
                        <input id="TrialBillingFrequency" name="TrialBillingFrequency" type="input" style="width: 350px;" value="<%=pcv_TrialBillingFrequency%>">
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                           
                    <br />
                	<img src="SubscriptionBridge/images/imgRequired.gif" alt="Required" width="10" height="10" border="0">
                    <b>Trial Billing Cycles:  </b>  
                    <div id="pTrialBillingCycles"> 
                        <input id="TrialBillingCycles" name="TrialBillingCycles" type="input" style="width: 350px;" value="<%=pcv_TrialBillingCycles%>">
                        <span class="textfieldRequiredMsg">
                            *												
                        </span>	
                    </div>
                    
                    <br />
                    
                </fieldset>

            </div>	
			<script type=text/javascript>
				var LinkedTrialPrice;				
                var Price;
				var TrialPrice;
				var RefName;
				var BillingPeriod 
				var BillingFrequency 
				var BillingCycles 
				var TrialBillingPeriod
				var TrialBillingFrequency
				var TrialBillingCycles 
				
                function optionGroupA() {
                
                    document.getElementById("tProdCannotMatch").style.display = '';
                    document.getElementById("tLinkedProduct").style.display = 'none';
                    
                    // Clear Sub Selections
                    document.form1.ProdCannotMatch[0].checked = false;
                    document.form1.ProdCannotMatch[1].checked = false;


                }
                
                function optionGroupB() {
                
                    document.getElementById("tLinkedProduct").style.display = '';
                    document.getElementById("tProdCannotMatch").style.display = 'none';
                    
                    // Clear Sub Selections
                    document.form1.ProdCannotMatch[0].checked = false;
                    document.form1.ProdCannotMatch[1].checked = false;

                    
                    
                }
				
				
				if (document.form1.IsLinked[1].checked == false) {
					document.form1.IsLinked[0].checked = true
					document.form1.IsLinked[1].checked = false
					openLinked();
				} else {
					document.form1.IsLinked[1].checked = true
					document.form1.IsLinked[0].checked = false
					//closeLinked();
					openLinked();
				}
				
				function Form1_Validator(theForm)
				{
					<%if pcv_strIsLinked="1" OR pcv_strIsLinked="" then%>
						if (theForm.LinkedIsTrial.checked == true)
						{
							if (theForm.LinkedTrialPrice.value == "")
							{
								alert("Please enter value for the field 'Trial Price'");
								return (false);
							}
						}
					<%else%>
						if (theForm.TrialPrice.value == "")
						{
							alert("Please enter value for the field 'Trial Price'");
							return (false);
						}
						if (theForm.RefName.value == "")
						{
							alert("Please enter value for the field 'Reference Name'");
							return (false);
						}
						if (theForm.BillingPeriod.value == "")
						{
							alert("Please select value for the field 'Billing Period'");
							return (false);
						}
						if (theForm.BillingFrequency.value == "")
						{
							alert("Please enter value for the field 'Billing Frequency'");
							return (false);
						}
						if (theForm.BillingCycles.value == "")
						{
							alert("Please enter value for the field 'Billing Cycles'");
							return (false);
						}
						if (theForm.TrialBillingPeriod.value == "")
						{
							alert("Please select value for the field 'Trial Billing Period'");
							return (false);
						}
						if (theForm.TrialBillingFrequency.value == "")
						{
							alert("Please enter value for the field 'Trial Billing Frequency'");
							return (false);
						}
						if (theForm.TrialBillingCycles.value == "")
						{
							alert("Please enter value for the field 'Trial Billing Cycles'");
							return (false);
						}
					<%end if%>
					return(true);
				}


                function openLinked() {
                
                    document.getElementById("tLinkedProduct").style.display = '';
					document.getElementById("tNOTLinkedProduct").style.display = 'none';
                    
                }
                
                function closeLinked() {
                
                    document.getElementById("tLinkedProduct").style.display = 'none';
					document.getElementById("tNOTLinkedProduct").style.display = '';
                    
                }

            </script>
        </td>
	</tr>
</table>
