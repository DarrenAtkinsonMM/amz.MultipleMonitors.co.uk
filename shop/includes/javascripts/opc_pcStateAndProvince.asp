<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.

'/////////////////////////////////////////////////////////////////////////////////
'// START: Countries Array
'/////////////////////////////////////////////////////////////////////////////////
query="SELECT CountryCode,countryName,pcSubDivisionID FROM countries ORDER BY countryName ASC"
set rsCountries=server.CreateObject("ADODB.RecordSet")
set rsCountries=conntemp.execute(query)
if err.number<>0 then
	call LogErrorToDatabase()
	set rsCountries=nothing
	call closedb()
	response.redirect "techErr.asp?err="&pcStrCustRefID
end if
xCountryArrayCount = 0
do while not rsCountries.eof
	'// We need to form our Array
	'xCountryArrayCount = xCountryArrayCount + 1
	pcv_strTmpCountryName = rsCountries("countryName") '// country's name
	pcv_strTmpCountryCode = rsCountries("countryCode") '// iso code
	pcv_strTmpCountryFlag = rsCountries("pcSubDivisionID")&"" '// state or province
	if len(pcv_strTmpCountryFlag)<0 OR len(pcv_strTmpCountryFlag)=NULL OR pcv_strTmpCountryFlag <> "1" then
		pcv_strTmpCountryFlag = "2"
	end if
	tmpData =  pcv_strTmpCountryName& "#" &pcv_strTmpCountryCode & "#" & pcv_strTmpCountryFlag
	if pcv_strTmpCountryName <> "" AND pcv_strTmpCountryCode <> "" AND len(pcv_strTmpCountryCode)<3 then			
		pcv_strCountriesArray = pcv_strCountriesArray & tmpData & chr(124)	
	end if	
rsCountries.movenext
loop
set rsCountries=nothing
'// Trim the last pipe if there is one
xStringLength = len(pcv_strCountriesArray)
if xStringLength>0 then
	pcv_strCountriesArray = left(pcv_strCountriesArray,(xStringLength-1))
end if
'response.write pcv_strCountriesArray
'response.end
'// Set Up Our Array
pcArrayCountries = split(pcv_strCountriesArray,chr(124))
pcv_intLBound = 0		
pcv_intLBound = LBound(pcArrayCountries)
pcv_intUBound = 0
pcv_intUBound = UBound(pcArrayCountries)
'/////////////////////////////////////////////////////////////////////////////////
'// END: Countries Array
'/////////////////////////////////////////////////////////////////////////////////
%>
<script type=text/javascript>
// The states default relationships are defined at load time via asp script.

// START: Dynamic Array Creation
var States = new Array();
<% 
For i = pcv_intLBound To pcv_intUBound
	'// get nested array
	pcArrayCountriesValues = split(pcArrayCountries(i),"#")
	response.write "States["&i&"] = new Array('"&pcArrayCountriesValues(1)&"','"&pcArrayCountriesValues(2)&"');" '& chr(10)
next 
%>
// END: Dynamic Array Creation

// START: State Lists
<% 
For i = pcv_intLBound To pcv_intUBound
	'// get nested array
	pcArrayCountriesValues = split(pcArrayCountries(i),"#")
	if pcArrayCountriesValues(2) = "1" then
		query="SELECT stateCode,stateName FROM states WHERE pcCountryCode = '"&pcArrayCountriesValues(1)&"' ORDER BY stateName ASC"
		set rsStates=server.CreateObject("ADODB.RecordSet")
		set rsStates=conntemp.execute(query)
		if err.number<>0 then
			call LogErrorToDatabase()
			set rsStates=nothing
			call closedb()
			response.redirect "techErr.asp?err="&pcStrCustRefID
		end if
		if not rsStates.eof then
		response.write ""&pcArrayCountriesValues(1)&" = new Array();" '& chr(10)
		response.write ""&pcArrayCountriesValues(1)&"[0] = new Array('','');" '& chr(10)
			x = 1
			do while not rsStates.eof
				a=rsStates("stateCode")
				b=replace(rsStates("stateName"), "'", "")	
				b=replace(b, ",", "")	
				b=replace(b, """", "")						
				response.write ""&pcArrayCountriesValues(1)&"["&x&"] = new Array('"&a&"','"&b&"');" '& chr(10)
				x=x+1
			rsStates.movenext
			loop
		end if
		set rsStates=nothing
	end if
next 
%>
// END: State Lists

// START: Switch Zones
function SwitchStates(TargetForm,a,countrybox,statebox,provincebox,SelectedSession,FormInstance) {
    var targetBox = eval("document." + TargetForm + "." + countrybox + "").options;
    var totalNumStates = <%=(pcv_intUBound+1)%>;
	var b;
	if(targetBox[a]) {
    	b = targetBox[a].value;
	}
    var zoning = countrybox;
	ClearFields(zoning);
	// clear the state / province box
    //$pc("#" + statebox).empty();
    $pc("#" + provincebox).val('');					
	for (var i=0; i<totalNumStates; i++) {     
		if (States[i][1] == 1) { // if the zone is "1"
			if (States[i][0] == b) { // if we have a relationship, populate the states					
				document.getElementById('zone1' + zoning).style.display='';
				document.getElementById('zone2' + zoning).style.display='none';
				OverRideReqField(TargetForm,'false','<%=pcv_isStateCodeRequired%>', FormInstance);
				LabelField(States[i][0],countrybox);
				FilterStates(TargetForm,States[i][0],statebox,SelectedSession);
				break;
			}
		} else {
			if (i == 0) { // set the field on the first loop. A state will override this.
				LabelField(States[i][0],countrybox);
			}
			document.getElementById('zone2' + zoning).style.display='';
			OverRideReqField(TargetForm,'<%=pcv_isProvinceCodeRequired%>','false', FormInstance);
		}
	}
}
// END: Switch Zones

// START: Filter States
function FilterStates(TargetForm,array_name,statebox, SelectedSession) {
	var targetBox = eval("document." + TargetForm + "." + statebox + "").options;
	targetBox.length = 0;
	var array_name = eval(array_name);
	for (i=0; i<array_name.length; i++){
		targetBox[i] = new Option(array_name[i][1], array_name[i][0]);
		// If we have match
		var SelectedState = SelectedSession;
		if (array_name[i][0] == SelectedState) {
			targetBox[i].selected = true;
		}
	}
}
// END: Filter States

// START: Clear Fields
function ClearFields(zoning) {
	document.getElementById('zone1' + zoning).style.display='none';	
	document.getElementById('zone2' + zoning).style.display='none';
}
// END: Clear Fields

// START: Check Selected State
function SelectState(TargetForm,countrybox,statebox,provincebox,SelectedSession,FormInstance) {
	ClearFields(countrybox);
	// clear the state / province box
    //$pc("#" + statebox).empty();
    $pc("#" + provincebox).val('');
    // get target index
    var targetBoxIndex = $pc("#" + countrybox)[0].selectedIndex;
	SwitchStates(TargetForm,targetBoxIndex,countrybox,statebox,provincebox,SelectedSession,FormInstance);
}
// END: Check Selected State

// START: Fill hidden text param
function OverRideReqField(TargetForm,a, b, FormInstance) {
	var prov = eval("document." + TargetForm + ".pcv_isProvinceCodeRequired"+FormInstance+"");
	var state = eval("document." + TargetForm + ".pcv_isStateCodeRequired"+FormInstance+"");
	prov.value = a;
	state.value = b;
}									
// END: Check Selected State

// START: Fill Label text param
function LabelField(CountryCode, countrybox) {
	var LabelField = document.getElementById('Label' + countrybox);
	var LabelField2 = document.getElementById('Label2' + countrybox);
	var CountryCode = CountryCode;
	if (CountryCode == 'US') {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_5")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_5")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
	} else {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_19")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_19")%>';
	}
	if (CountryCode == 'CA') {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';	
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
	}
	if (CountryCode == 'GB') {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';	
	}
	if (CountryCode == 'NZ') {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';	
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
	}
	if (CountryCode == 'AU') {
		LabelField.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';
		LabelField2.innerHTML = '<%=dictLanguage.Item(Session("language")&"_CustAddModShip_6")%><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %>';	
	}
}									
// END: Fill Label text param
</script>

<% 
Dim pcv_strSPFieldLabel
select case Session(pcv_strSessionPrefix&pcv_strCountryBox)
case "US"
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_5")
case "CA"
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_6")
case "GB"
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_6")
case "NZ"
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_6")
case "AU"
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_6")
case else
	pcv_strSPFieldLabel = dictLanguage.Item(Session("language")&"_CustAddModShip_19")
end select

Dim pcv_strTargetBox, pcv_strTargetForm, pcv_strCountryBox, pcv_strProvinceBox, pcv_strContextName
Dim pcv_strStateCode, pcv_strStateName, pcv_isStateCodeRequired, pcv_isProvinceCodeRequired
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: State/ Province
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Public Sub pcs_StateProvince
    %>
    <div id="opc<%= pcv_strContextName %>State">
        <div class="form-group" id="zone1<%=pcv_strCountryBox%>" style="display: none;">
            <label ID="Label<%=pcv_strCountryBox%>" for="<%=pcv_strTargetBox%>"><%=pcv_strSPFieldLabel %><% If pcv_isStateCodeRequired Then %><div class="pcRequiredIcon"><img src="<%=pcv_strRequiredIcon%>" alt="Required"></div><% End If %></label>
            <% 
            query="SELECT stateCode,stateName FROM states ORDER BY stateName ASC"
            set rs=server.CreateObject("ADODB.RecordSet")
            set rs=conntemp.execute(query)
            if err.number<>0 then
                call LogErrorToDatabase()
                set rs=nothing
                call closedb()
                response.redirect "techErr.asp?err="&pcStrCustRefID
            end if
            %>
            <select autocomplete="off" class="form-control" <% If len(pcv_strStateModel)>0 Then %>ng-model='<%=pcv_strStateModel%>'<% End If %> name="<%=pcv_strTargetBox%>" id="<%=pcv_strTargetBox%>">
                <option value=""></option>
                <% 
                do while not rs.eof
                    pshippingStateCode=rs("stateCode") 
                    pshippingStateName=rs("stateName") 
                    %>
                    <option value="<%=pshippingStateCode%>" <%=pcf_SelectOption(pcv_strTargetBox,pshippingStateCode)%>><%=pshippingStateName%></option>
                    <% 
                    rs.movenext
                loop
                set rs=nothing
                %>
            </select>
            <% pcs_RequiredImageTagHorizontal pcv_strTargetBox, pcv_isStateCodeRequired %>            
        </div>
    </div>
    <div id="opc<%= pcv_strContextName %>Province">
        <div class="form-group" id="zone2<%=pcv_strCountryBox%>" style="display: none;">
        <div class="row specb-row"><div class="col-sm-3 daOpcBillingAdd">
								<p>County/Region:</p>
                                </div><div class="col-sm-9">
            <input autocomplete="off" class="form-control" type="text" name="<%=pcv_strProvinceBox%>" id="<%=pcv_strProvinceBox%>" <% If len(pcv_strProvinceModel)>0 Then %>ng-model='<%=pcv_strProvinceModel%>'<% End If %>>
            <label style="display:none;" ID="Label2<%=pcv_strCountryBox%>" for="<%=pcv_strProvinceBox%>"><%=pcv_strSPFieldLabel %></label>            
            <% pcs_RequiredImageTag pcv_strProvinceBox, pcv_isProvinceCodeRequired %>			
            <input name="pcv_isStateCodeRequired<%=pcv_strFormInstance%>" type="hidden" value="<%=pcv_isStateCodeRequired%>" />	
            <input name="pcv_isProvinceCodeRequired<%=pcv_strFormInstance%>" type="hidden" value="<%=pcv_isProvinceCodeRequired%>" />          
        </div></div></div>
    </div>
    <script type=text/javascript>
        $pc(function() {
            SelectState('<%=pcv_strTargetForm%>','<%=pcv_strCountryBox%>', '<%=pcv_strTargetBox%>', '<%=pcv_strProvinceBox%>', '<%=pcv_strTargetBoxValue%>', '<%=pcv_strFormInstance%>');
        });
    </script>  
    <%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: State/ Province
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Display the Country Dropdown
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~	
Dim pcv_strCountryCode, pcv_strTargetBoxValue, pcv_strFormInstance
    Public Sub pcs_CountryDropdown

	If Session(pcv_strSessionPrefix & pcv_strCountryBox) = "" Then
		Session(pcv_strSessionPrefix & pcv_strCountryBox) = scShipFromPostalCountry
	End If
	%>
    <div id="opc<%= pcv_strContextName %>Country">    
        <div id="opcBillingZip" class="form-group">
        <div class="row specb-row">
                                <div class="col-sm-3 daOpcBillingAdd">
								<p>Country: <span style="color:red;">*</span></p>
                                </div>
                                <div class="col-sm-9">
            <%
            pcv_strTargetBoxValue=Session(pcv_strSessionPrefix&pcv_strTargetBox)
            %>
            <select autocomplete="off" class="form-control" <% If len(pcv_strCountryModel)>0 Then %>ng-model='<%=pcv_strCountryModel%>'<% End If %> name="<%=pcv_strCountryBox%>" id="<%=pcv_strCountryBox%>" onchange="SwitchStates('<%=pcv_strTargetForm%>',this.options.selectedIndex, '<%=pcv_strCountryBox%>', '<%=pcv_strTargetBox%>', '<%=pcv_strProvinceBox%>', '<%=pcv_strTargetBoxValue%>', '<%=pcv_strFormInstance%>');<%=tmp_CountryBoxFunc%>">
                <% If pcStrPageName <> "estimateShipCost.asp" Then %>
                    <option value=""></option>
                <% End If %>
                <% 
                For i = pcv_intLBound To pcv_intUBound
                    pcArrayCountriesValues = split(pcArrayCountries(i),"#")
                    pcv_strCountryCode = pcArrayCountriesValues(1)
                    pcv_strCountryName = pcArrayCountriesValues(0)
                    %>
                    <option value="<%=pcv_strCountryCode%>" <%=pcf_SelectOption(pcv_strCountryBox,pcv_strCountryCode)%>><%=pcv_strCountryName%></option>
                    <%
                Next
                %>
            </select>
            <% pcs_RequiredImageTagHorizontal pcv_strCountryBox, pcv_isCountryCodeRequired %> 
            </div></div>
        </div>
	</div> 
    <%
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Display the Country Dropdown
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>
