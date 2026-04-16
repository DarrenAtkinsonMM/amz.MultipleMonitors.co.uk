<%
'This file is part of ProductCart, an ecommerce application developed and sold by Early Impact LLC. ProductCart, its source code, the ProductCart name and logo are property of Early Impact, LLC. Copyright 2001-2003. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of Early Impact. To contact Early Impact, please visit www.earlyimpact.com.
%>
<% response.Buffer=true %>
<!--#include file="../includes/common.asp"-->

<!--#include file="header_wrapper.asp"-->
<script src="/shop/includes/jquery/jquery.validate.min.js"></script>	
<script type="text/javascript">
	// Wait for the DOM to be ready
$(function() {
  // Initialize form validation on the registration form.
  // It has the name attribute "registration"
  $("form[name='PaymentForm']").validate({
    // Specify validation rules
    rules: {
      // The key name on the left side is the name attribute
      // of an input field. Validation rules are defined
      // on the right side
      mpBillFirstName: "required",
      mpBillLastName: "required",
	  mpBillTel: "required",
	  mpBillAddress1: "required",
	  mpBillCity: "required",
	  mpBillPCode: "required",
	  mpDelFirstName: "required",
      mpDelLastName: "required",
	  mpDelTel: "required",
	  mpDelAddress1: "required",
	  mpDelCity: "required",
	  mpDelPCode: "required",
		
      mpBillEmail: {
        required: true,
        // Specify that email should be validated
        // by the built-in "email" rule
        email: true
      },
    },
    // Specify validation error messages
    messages: {
      mpBillFirstName: "&nbsp;&nbsp;*",
      mpBillLastName: "&nbsp;&nbsp;*",
      mpBillEmail: "&nbsp;&nbsp;*",
	  mpBillTel: "&nbsp;&nbsp;*",
	  mpBillAddress1: "&nbsp;&nbsp;*",
	  mpBillCity: "&nbsp;&nbsp;*",
	  mpBillPCode: "&nbsp;&nbsp;*",
	  mpDelFirstName: "&nbsp;&nbsp;*",
      mpDelLastName: "&nbsp;&nbsp;*",
	  mpDelTel: "&nbsp;&nbsp;*",
	  mpDelAddress1: "&nbsp;&nbsp;*",
	  mpDelCity: "&nbsp;&nbsp;*",
	  mpDelPCode: "&nbsp;&nbsp;*",
    },
    // Make sure the form is submitted to the destination defined
    // in the "action" attribute of the form when valid
    submitHandler: function(form) {
      form.submit();
    }
  });
});
</script>	
<script type="text/javascript">
function sameDel() {
	if (document.getElementById('mpDelChk').checked == true) {
		document.getElementById('mpDelFirstName').value = document.getElementById('mpBillFirstName').value;
		document.getElementById('mpDelLastName').value = document.getElementById('mpBillLastName').value;
		document.getElementById('mpDelTel').value = document.getElementById('mpBillTel').value;
		document.getElementById('mpDelAddress1').value = document.getElementById('mpBillAddress1').value;
		document.getElementById('mpDelAddress2').value = document.getElementById('mpBillAddress2').value;
		document.getElementById('mpDelAddress3').value = document.getElementById('mpBillAddress3').value;
		document.getElementById('mpDelCity').value = document.getElementById('mpBillCity').value;
		document.getElementById('mpDelPCode').value = document.getElementById('mpBillPCode').value;
		document.getElementById('mpDelCountry').value = document.getElementById('mpBillCountry').value;
	} else {
		document.getElementById('mpDelFirstName').value = '';
		document.getElementById('mpDelLastName').value = '';
		document.getElementById('mpDelTel').value = '';
		document.getElementById('mpDelAddress1').value = '';
		document.getElementById('mpDelAddress2').value = '';
		document.getElementById('mpDelCity').value = '';
		document.getElementById('mpDelPCode').value = '';
		document.getElementById('mpDelCountry').value = 'GB';
	}
}
</script>
<% 
mpAmount = request.querystring("amount")   
   %>
<!-- Header: pagetitle -->
    <header id="pagetitle" class="pagetitle">
		<div class="pt-content">
			<div class="container">
				<div class="row">
					<div class="col-xs-12 pagetitle">
						<div class="wow fadeInDown" data-wow-offset="0" data-wow-delay="0.1s">
							<h3 class="color marginbot-0 h-semi" title="">Quick Payment Page - Step: 1 / 3</h3>
						</div>
					</div>				
				</div>		
			</div>		
		</div>	
    </header>
	<!-- /Header: pagetitle -->

	<section id="intWarranties" class="intWarranties paddingtop-30 paddingbot-70">	
           <div class="container">
				<div class="row">
                	<div class="col-sm-12 warrantyHeading wow fadeInUp" data-wow-offset="0" data-wow-delay="0.1s"><div id="pcMain">
						
						<p><strong>Payment For: <%=request.QueryString("pay")%></strong></p>
							<p><strong>Amount: &pound;<%=money(mpAmount)%></strong></p>
							<form method="POST" action="manpay2.asp" name="PaymentForm" class="pcForms">
							<input type="hidden" name="mpAmount" value="<%=request.querystring("amount") %>">
							<input type="hidden" name="mpDesc" value="<%=request.querystring("pay") %>">
						<div class="damanpayHolder">
							<div class="damanpaylabel">
								<label>First Name:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillFirstName" id="mpBillFirstName" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Last Name:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillLastName" id="mpBillLastName" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Phone:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillTel" id="mpBillTel" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Email:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillEmail" id="mpBillEmail" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 1:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillAddress1" id="mpBillAddress1" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 2:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillAddress2" id="mpBillAddress2" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 3:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillAddress3" id="mpBillAddress3" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Town / City:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillCity" id="mpBillCity" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Post Code:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpBillPCode" id="mpBillPCode" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Country:</label>	
							</div>
							<div class="damanpayinput">
								<select name="mpBillCountry" id="mpBillCountry" tabindex="10">
						  <option value="AF" >Afghanistan</option>
						  <option value="AX" >Åland Islands</option>
						  <option value="AL" >Albania</option>
						  <option value="DZ" >Algeria</option>
						  <option value="AS" >American Samoa</option>
						  <option value="AD" >Andorra</option>
						  <option value="AO" >Angola</option>
						  <option value="AI" >Anguilla</option>
						  <option value="AQ" >Antarctica</option>
						  <option value="AG" >Antigua and Barbuda</option>
						  <option value="AR" >Argentina</option>
						  <option value="AM" >Armenia</option>
						  <option value="AW" >Aruba</option>
						  <option value="AU" >Australia</option>
						  <option value="AT" >Austria</option>
						  <option value="AZ" >Azerbaijan</option>
						  <option value="BS" >Bahamas</option>
						  <option value="BH" >Bahrain</option>
						  <option value="BD" >Bangladesh</option>
						  <option value="BB" >Barbados</option>
						  <option value="BY" >Belarus</option>
						  <option value="BE" >Belgium</option>
						  <option value="BZ" >Belize</option>
						  <option value="BJ" >Benin</option>
						  <option value="BM" >Bermuda</option>
						  <option value="BT" >Bhutan</option>
						  <option value="BO" >Bolivia</option>
						  <option value="BA" >Bosnia and Herzegovina</option>
						  <option value="BW" >Botswana</option>
						  <option value="BV" >Bouvet Island</option>
						  <option value="BR" >Brazil</option>
						  <option value="IO" >British Indian Ocean Territory</option>
						  <option value="BN" >Brunei Darussalam</option>
						  <option value="BG" >Bulgaria</option>
						  <option value="BF" >Burkina Faso</option>
						  <option value="BI" >Burundi</option>
						  <option value="CI" >C?te D'Ivoire</option>
						  <option value="KH" >Cambodia</option>
						  <option value="CM" >Cameroon</option>
						  <option value="CA" >Canada</option>
						  <option value="CV" >Cape Verde</option>
						  <option value="KY" >Cayman Islands</option>
						  <option value="CF" >Central African Republic</option>
						  <option value="TD" >Chad</option>
						  <option value="CL" >Chile</option>
						  <option value="CN" >China - Peoples Republic of</option>
						  <option value="CX" >Christmas Island</option>
						  <option value="CC" >Cocos (Keeling) Islands</option>
						  <option value="CO" >Colombia</option>
						  <option value="KM" >Comoros</option>
						  <option value="CG" >Congo</option>
						  <option value="CK" >Cook Islands</option>
						  <option value="CR" >Costa Rica</option>
						  <option value="HR" >Croatia</option>
						  <option value="CU" >Cuba</option>
						  <option value="CY" >Cyprus</option>
						  <option value="CZ" >Czech Republic</option>
						  <option value="DK" >Denmark</option>
						  <option value="DJ" >Djibouti</option>
						  <option value="DM" >Dominica</option>
						  <option value="DO" >Dominican Republic</option>
						  <option value="EC" >Ecuador</option>
						  <option value="EG" >Egypt</option>
						  <option value="SV" >El Salvador</option>
						  <option value="GQ" >Equatorial Guinea</option>
						  <option value="ER" >Eritrea</option>
						  <option value="EE" >Estonia</option>
						  <option value="ET" >Ethiopia</option>
						  <option value="FK" >Falkland Islands (Malvinas)</option>
						  <option value="FO" >Faroe Islands</option>
						  <option value="FJ" >Fiji</option>
						  <option value="FI" >Finland</option>
						  <option value="FR" >France</option>
						  <option value="GF" >French Guiana</option>
						  <option value="PF" >French Polynesia</option>
						  <option value="TF" >French Southern Territories</option>
						  <option value="GA" >Gabon</option>
						  <option value="GM" >Gambia</option>
						  <option value="GE" >Georgia</option>

						  <option value="DE" >Germany</option>
						  <option value="GH" >Ghana</option>
						  <option value="GI" >Gibraltar</option>

						  <option value="GR" >Greece</option>
						  <option value="GL" >Greenland</option>
						  <option value="GD" >Grenada</option>
						  <option value="GP" >Guadeloupe</option>
						  <option value="GU" >Guam</option>
						  <option value="GT" >Guatemala</option>
						  <option value="GG" >Guernsey</option>
						  <option value="GN" >Guinea</option>
						  <option value="GW" >Guinea-Bissau</option>
						  <option value="GY" >Guyana</option>
						  <option value="HT" >Haiti</option>
						  <option value="HN" >Honduras</option>
						  <option value="HK" >Hong Kong</option>
						  <option value="HU" >Hungary</option>
						  <option value="IS" >Iceland</option>
						  <option value="IN" >India</option>
						  <option value="ID" >Indonesia</option>
						  <option value="IR" >Iran - Islamic Republic Of</option>
						  <option value="IQ" >Iraq</option>
						  <option value="IE" >Ireland</option>
						  <option value="IM" >Isle of Man</option>
						  <option value="IL" >Israel</option>
						  <option value="IT" >Italy</option>
						  <option value="JM" >Jamaica</option>
						  <option value="JP" >Japan</option>
						  <option value="JE" >Jersey</option>
						  <option value="JO" >Jordan</option>
						  <option value="KZ" >Kazakhstan</option>
						  <option value="KE" >Kenya</option>
						  <option value="KI" >Kiribati</option>
						  <option value="KR" >Korea - Republic of</option>
						  <option value="KW" >Kuwait</option>
						  <option value="KG" >Kyrgyzstan</option>
						  <option value="LV" >Latvia</option>
						  <option value="LB" >Lebanon</option>
						  <option value="LS" >Lesotho</option>
						  <option value="LR" >Liberia</option>
						  <option value="LY" >Libyan Arab Jamahiriya</option>
						  <option value="LI" >Liechtenstein</option>
						  <option value="LT" >Lithuania</option>
						  <option value="LU" >Luxembourg</option>
						  <option value="MO" >Macao</option>
						  <option value="MK" >Macedonia</option>
						  <option value="MG" >Madagascar</option>
						  <option value="MW" >Malawi</option>
						  <option value="MY" >Malaysia</option>
						  <option value="MV" >Maldives</option>
						  <option value="ML" >Mali</option>
						  <option value="MT" >Malta</option>
						  <option value="MH" >Marshall Islands</option>
						  <option value="MQ" >Martinique</option>
						  <option value="MR" >Mauritania</option>
						  <option value="MU" >Mauritius</option>
						  <option value="YT" >Mayotte</option>
						  <option value="MX" >Mexico</option>
						  <option value="FM" >Micronesia - Federated States of</option>
						  <option value="MD" >Moldova - Republic of</option>
						  <option value="MC" >Monaco</option>
						  <option value="MN" >Mongolia</option>
						  <option value="ME" >Montenegro</option>
						  <option value="MS" >Montserrat</option>
						  <option value="MA" >Morocco</option>
						  <option value="MZ" >Mozambique</option>
						  <option value="MM" >Myanmar</option>
						  <option value="NA" >Namibia</option>
						  <option value="NR" >Nauru</option>
						  <option value="NP" >Nepal</option>
						  <option value="NL" >Netherlands</option>
						  <option value="AN" >Netherlands Antilles</option>
						  <option value="NC" >New Caledonia</option>
						  <option value="NZ" >New Zealand</option>
						  <option value="NI" >Nicaragua</option>
						  <option value="NE" >Niger</option>
						  <option value="NG" >Nigeria</option>
						  <option value="NU" >Niue</option>
						  <option value="NF" >Norfolk Island</option>
						  <option value="NO" >Norway</option>
						  <option value="MP" >Nothern Mariana Islands</option>
						  <option value="OM" >Oman</option>
						  <option value="PK" >Pakistan</option>
						  <option value="PW" >Palau</option>
						  <option value="PA" >Panama</option>
						  <option value="PG" >Papua New Guinea</option>
						  <option value="PY" >Paraguay</option>
						  <option value="PE" >Peru</option>
						  <option value="PH" >Philippines</option>
						  <option value="PN" >Pitcairn</option>
						  <option value="PL" >Poland</option>
						  <option value="PT" >Portugal</option>
						  <option value="PR" >Puerto Rico</option>
						  <option value="QA" >Qatar</option>
						  <option value="RE" >Réunion</option>
						  <option value="RO" >Romania</option>
						  <option value="RU" >Russian Federation</option>
						  <option value="RW" >Rwanda</option>
						  <option value="SH" >Saint Helena</option>
						  <option value="KN" >Saint Kitts and Nevis</option>
						  <option value="LC" >Saint Lucia</option>
						  <option value="PM" >Saint Pierre and Miquelon</option>
						  <option value="VC" >Saint Vincent and the Grenadines</option>
						  <option value="WS" >Samoa</option>
						  <option value="SM" >San Marino</option>
						  <option value="ST" >Sao Tome and Principe</option>
						  <option value="SA" >Saudi Arabia</option>
						  <option value="SN" >Senegal</option>
						  <option value="RS" >Serbia</option>
						  <option value="SC" >Seychelles</option>
						  <option value="SL" >Sierra Leone</option>
						  <option value="SG" >Singapore</option>
						  <option value="SK" >Slovakia</option>
						  <option value="SI" >Slovenia</option>
						  <option value="SB" >Solomon Islands</option>
						  <option value="SO" >Somalia</option>
						  <option value="ZA" >South Africa</option>
						  <option value="ES" >Spain</option>
						  <option value="LK" >Sri Lanka</option>
						  <option value="SD" >Sudan</option>
						  <option value="SR" >Suriname</option>
						  <option value="SJ" >Svalbard and Jan Mayen</option>
						  <option value="SZ" >Swaziland</option>
						  <option value="SE" >Sweden</option>
						  <option value="CH" >Switzerland</option>
						  <option value="SY" >Syrian Arab Republic</option>
						  <option value="TW" >Taiwan - Province Of China</option>
						  <option value="TJ" >Tajikistan</option>
						  <option value="TZ" >Tanzania - United Republic Of</option>
						  <option value="TH" >Thailand</option>
						  <option value="TL" >Timor-Leste</option>
						  <option value="TG" >Togo</option>
						  <option value="TK" >Tokelau</option>
						  <option value="TO" >Tonga</option>
						  <option value="TT" >Trinidad And Tobago</option>
						  <option value="TN" >Tunisia</option>
						  <option value="TR" >Turkey</option>
						  <option value="TM" >Turkmenistan</option>
						  <option value="TC" >Turks and Caicos Islands</option>
						  <option value="TV" >Tuvalu</option>
						  <option value="UG" >Uganda</option>
						  <option value="UA" >Ukraine</option>
						  <option value="AE" >United Arab Emirates</option>
						  <option value="GB" selected>United Kingdom</option>
						  <option value="US" >United States</option>
						  <option value="UY" >Uruguay</option>
						  <option value="UM" >US - Minor Outlying Islands</option>
						  <option value="UZ" >Uzbekistan</option>
						  <option value="VU" >Vanuatu</option>
						  <option value="VA" >Vatican City</option>
						  <option value="VE" >Venezuela</option>
						  <option value="VN" >VietNam</option>
						  <option value="VG" >Virgin Islands - British</option>
						  <option value="VI" >Virgin Islands - U.S.</option>
						  <option value="WF" >Wallis and Futuna Islands</option>
						  <option value="EH" >Western Sahara</option>
						  <option value="YE" >Yemen</option>
						  <option value="ZM" >Zambia</option>
						  <option value="ZW" >Zimbabwe</option>
						  </select>
							</div>
						</div>
						<div class="damanpayHolder">
							<div class="damanpaysamedel">
								<label>Billing &amp; Delivery Addresses Are The Same:&nbsp;&nbsp;&nbsp; </label><input name="mpDelChk" id="mpDelChk" type="checkbox" value="1" onchange="sameDel();" tabindex="11" />
							</div>
							<div class="damanpaylabel">
								<label>First Name:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelFirstName" id="mpDelFirstName" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Last Name:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelLastName" id="mpDelLastName" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Phone:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelTel" id="mpDelTel" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 1:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelAddress1" id="mpDelAddress1" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 2:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelAddress2" id="mpDelAddress2" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Address Line 3:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelAddress3" id="mpDelAddress3" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Town / City:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelCity" id="mpDelCity" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Post Code:</label>	
							</div>
							<div class="damanpayinput">
								<input name="mpDelPCode" id="mpDelPCode" type="text" tabindex="1" />
							</div>
							<div class="damanpaylabel">
								<label>Country:</label>	
							</div>
							<div class="damanpayinput">
								<select name="mpDelCountry" id="mpDelCountry" tabindex="10">
						  <option value="AF" >Afghanistan</option>
						  <option value="AX" >Åland Islands</option>
						  <option value="AL" >Albania</option>
						  <option value="DZ" >Algeria</option>
						  <option value="AS" >American Samoa</option>
						  <option value="AD" >Andorra</option>
						  <option value="AO" >Angola</option>
						  <option value="AI" >Anguilla</option>
						  <option value="AQ" >Antarctica</option>
						  <option value="AG" >Antigua and Barbuda</option>
						  <option value="AR" >Argentina</option>
						  <option value="AM" >Armenia</option>
						  <option value="AW" >Aruba</option>
						  <option value="AU" >Australia</option>
						  <option value="AT" >Austria</option>
						  <option value="AZ" >Azerbaijan</option>
						  <option value="BS" >Bahamas</option>
						  <option value="BH" >Bahrain</option>
						  <option value="BD" >Bangladesh</option>
						  <option value="BB" >Barbados</option>
						  <option value="BY" >Belarus</option>
						  <option value="BE" >Belgium</option>
						  <option value="BZ" >Belize</option>
						  <option value="BJ" >Benin</option>
						  <option value="BM" >Bermuda</option>
						  <option value="BT" >Bhutan</option>
						  <option value="BO" >Bolivia</option>
						  <option value="BA" >Bosnia and Herzegovina</option>
						  <option value="BW" >Botswana</option>
						  <option value="BV" >Bouvet Island</option>
						  <option value="BR" >Brazil</option>
						  <option value="IO" >British Indian Ocean Territory</option>
						  <option value="BN" >Brunei Darussalam</option>
						  <option value="BG" >Bulgaria</option>
						  <option value="BF" >Burkina Faso</option>
						  <option value="BI" >Burundi</option>
						  <option value="CI" >C?te D'Ivoire</option>
						  <option value="KH" >Cambodia</option>
						  <option value="CM" >Cameroon</option>
						  <option value="CA" >Canada</option>
						  <option value="CV" >Cape Verde</option>
						  <option value="KY" >Cayman Islands</option>
						  <option value="CF" >Central African Republic</option>
						  <option value="TD" >Chad</option>
						  <option value="CL" >Chile</option>
						  <option value="CN" >China - Peoples Republic of</option>
						  <option value="CX" >Christmas Island</option>
						  <option value="CC" >Cocos (Keeling) Islands</option>
						  <option value="CO" >Colombia</option>
						  <option value="KM" >Comoros</option>
						  <option value="CG" >Congo</option>
						  <option value="CK" >Cook Islands</option>
						  <option value="CR" >Costa Rica</option>
						  <option value="HR" >Croatia</option>
						  <option value="CU" >Cuba</option>
						  <option value="CY" >Cyprus</option>
						  <option value="CZ" >Czech Republic</option>
						  <option value="DK" >Denmark</option>
						  <option value="DJ" >Djibouti</option>
						  <option value="DM" >Dominica</option>
						  <option value="DO" >Dominican Republic</option>
						  <option value="EC" >Ecuador</option>
						  <option value="EG" >Egypt</option>
						  <option value="SV" >El Salvador</option>
						  <option value="GQ" >Equatorial Guinea</option>
						  <option value="ER" >Eritrea</option>
						  <option value="EE" >Estonia</option>
						  <option value="ET" >Ethiopia</option>
						  <option value="FK" >Falkland Islands (Malvinas)</option>
						  <option value="FO" >Faroe Islands</option>
						  <option value="FJ" >Fiji</option>
						  <option value="FI" >Finland</option>
						  <option value="FR" >France</option>
						  <option value="GF" >French Guiana</option>
						  <option value="PF" >French Polynesia</option>
						  <option value="TF" >French Southern Territories</option>
						  <option value="GA" >Gabon</option>
						  <option value="GM" >Gambia</option>
						  <option value="GE" >Georgia</option>

						  <option value="DE" >Germany</option>
						  <option value="GH" >Ghana</option>
						  <option value="GI" >Gibraltar</option>

						  <option value="GR" >Greece</option>
						  <option value="GL" >Greenland</option>
						  <option value="GD" >Grenada</option>
						  <option value="GP" >Guadeloupe</option>
						  <option value="GU" >Guam</option>
						  <option value="GT" >Guatemala</option>
						  <option value="GG" >Guernsey</option>
						  <option value="GN" >Guinea</option>
						  <option value="GW" >Guinea-Bissau</option>
						  <option value="GY" >Guyana</option>
						  <option value="HT" >Haiti</option>
						  <option value="HN" >Honduras</option>
						  <option value="HK" >Hong Kong</option>
						  <option value="HU" >Hungary</option>
						  <option value="IS" >Iceland</option>
						  <option value="IN" >India</option>
						  <option value="ID" >Indonesia</option>
						  <option value="IR" >Iran - Islamic Republic Of</option>
						  <option value="IQ" >Iraq</option>
						  <option value="IE" >Ireland</option>
						  <option value="IM" >Isle of Man</option>
						  <option value="IL" >Israel</option>
						  <option value="IT" >Italy</option>
						  <option value="JM" >Jamaica</option>
						  <option value="JP" >Japan</option>
						  <option value="JE" >Jersey</option>
						  <option value="JO" >Jordan</option>
						  <option value="KZ" >Kazakhstan</option>
						  <option value="KE" >Kenya</option>
						  <option value="KI" >Kiribati</option>
						  <option value="KR" >Korea - Republic of</option>
						  <option value="KW" >Kuwait</option>
						  <option value="KG" >Kyrgyzstan</option>
						  <option value="LV" >Latvia</option>
						  <option value="LB" >Lebanon</option>
						  <option value="LS" >Lesotho</option>
						  <option value="LR" >Liberia</option>
						  <option value="LY" >Libyan Arab Jamahiriya</option>
						  <option value="LI" >Liechtenstein</option>
						  <option value="LT" >Lithuania</option>
						  <option value="LU" >Luxembourg</option>
						  <option value="MO" >Macao</option>
						  <option value="MK" >Macedonia</option>
						  <option value="MG" >Madagascar</option>
						  <option value="MW" >Malawi</option>
						  <option value="MY" >Malaysia</option>
						  <option value="MV" >Maldives</option>
						  <option value="ML" >Mali</option>
						  <option value="MT" >Malta</option>
						  <option value="MH" >Marshall Islands</option>
						  <option value="MQ" >Martinique</option>
						  <option value="MR" >Mauritania</option>
						  <option value="MU" >Mauritius</option>
						  <option value="YT" >Mayotte</option>
						  <option value="MX" >Mexico</option>
						  <option value="FM" >Micronesia - Federated States of</option>
						  <option value="MD" >Moldova - Republic of</option>
						  <option value="MC" >Monaco</option>
						  <option value="MN" >Mongolia</option>
						  <option value="ME" >Montenegro</option>
						  <option value="MS" >Montserrat</option>
						  <option value="MA" >Morocco</option>
						  <option value="MZ" >Mozambique</option>
						  <option value="MM" >Myanmar</option>
						  <option value="NA" >Namibia</option>
						  <option value="NR" >Nauru</option>
						  <option value="NP" >Nepal</option>
						  <option value="NL" >Netherlands</option>
						  <option value="AN" >Netherlands Antilles</option>
						  <option value="NC" >New Caledonia</option>
						  <option value="NZ" >New Zealand</option>
						  <option value="NI" >Nicaragua</option>
						  <option value="NE" >Niger</option>
						  <option value="NG" >Nigeria</option>
						  <option value="NU" >Niue</option>
						  <option value="NF" >Norfolk Island</option>
						  <option value="NO" >Norway</option>
						  <option value="MP" >Nothern Mariana Islands</option>
						  <option value="OM" >Oman</option>
						  <option value="PK" >Pakistan</option>
						  <option value="PW" >Palau</option>
						  <option value="PA" >Panama</option>
						  <option value="PG" >Papua New Guinea</option>
						  <option value="PY" >Paraguay</option>
						  <option value="PE" >Peru</option>
						  <option value="PH" >Philippines</option>
						  <option value="PN" >Pitcairn</option>
						  <option value="PL" >Poland</option>
						  <option value="PT" >Portugal</option>
						  <option value="PR" >Puerto Rico</option>
						  <option value="QA" >Qatar</option>
						  <option value="RE" >Réunion</option>
						  <option value="RO" >Romania</option>
						  <option value="RU" >Russian Federation</option>
						  <option value="RW" >Rwanda</option>
						  <option value="SH" >Saint Helena</option>
						  <option value="KN" >Saint Kitts and Nevis</option>
						  <option value="LC" >Saint Lucia</option>
						  <option value="PM" >Saint Pierre and Miquelon</option>
						  <option value="VC" >Saint Vincent and the Grenadines</option>
						  <option value="WS" >Samoa</option>
						  <option value="SM" >San Marino</option>
						  <option value="ST" >Sao Tome and Principe</option>
						  <option value="SA" >Saudi Arabia</option>
						  <option value="SN" >Senegal</option>
						  <option value="RS" >Serbia</option>
						  <option value="SC" >Seychelles</option>
						  <option value="SL" >Sierra Leone</option>
						  <option value="SG" >Singapore</option>
						  <option value="SK" >Slovakia</option>
						  <option value="SI" >Slovenia</option>
						  <option value="SB" >Solomon Islands</option>
						  <option value="SO" >Somalia</option>
						  <option value="ZA" >South Africa</option>
						  <option value="ES" >Spain</option>
						  <option value="LK" >Sri Lanka</option>
						  <option value="SD" >Sudan</option>
						  <option value="SR" >Suriname</option>
						  <option value="SJ" >Svalbard and Jan Mayen</option>
						  <option value="SZ" >Swaziland</option>
						  <option value="SE" >Sweden</option>
						  <option value="CH" >Switzerland</option>
						  <option value="SY" >Syrian Arab Republic</option>
						  <option value="TW" >Taiwan - Province Of China</option>
						  <option value="TJ" >Tajikistan</option>
						  <option value="TZ" >Tanzania - United Republic Of</option>
						  <option value="TH" >Thailand</option>
						  <option value="TL" >Timor-Leste</option>
						  <option value="TG" >Togo</option>
						  <option value="TK" >Tokelau</option>
						  <option value="TO" >Tonga</option>
						  <option value="TT" >Trinidad And Tobago</option>
						  <option value="TN" >Tunisia</option>
						  <option value="TR" >Turkey</option>
						  <option value="TM" >Turkmenistan</option>
						  <option value="TC" >Turks and Caicos Islands</option>
						  <option value="TV" >Tuvalu</option>
						  <option value="UG" >Uganda</option>
						  <option value="UA" >Ukraine</option>
						  <option value="AE" >United Arab Emirates</option>
						  <option value="GB" selected>United Kingdom</option>
						  <option value="US" >United States</option>
						  <option value="UY" >Uruguay</option>
						  <option value="UM" >US - Minor Outlying Islands</option>
						  <option value="UZ" >Uzbekistan</option>
						  <option value="VU" >Vanuatu</option>
						  <option value="VA" >Vatican City</option>
						  <option value="VE" >Venezuela</option>
						  <option value="VN" >VietNam</option>
						  <option value="VG" >Virgin Islands - British</option>
						  <option value="VI" >Virgin Islands - U.S.</option>
						  <option value="WF" >Wallis and Futuna Islands</option>
						  <option value="EH" >Western Sahara</option>
						  <option value="YE" >Yemen</option>
						  <option value="ZM" >Zambia</option>
						  <option value="ZW" >Zimbabwe</option>
						  </select>
							</div>
						</div>

						<div class="damanpaysubmit">
						<input type="submit" value="Continue To Payment" name="Continue" class="btn product-action pg-green-btn" id="submit" tabindex="21">
						
						</div>
					</div>
				</div>
		    </div>
					</form>
    </section>	
    <!-- /Section: Welcome -->
<% 
'***********************************************
' Useful methods
'***********************************************

function findField( fieldName, postResponse )
  items = split( postResponse, chr( 13 ) )
  for idx = LBound( items ) to UBound( items )
    item = replace( items( idx ), chr( 10 ), "" )
    if InStr( item, fieldName & "=" ) = 1 then
      ' found
      findField = right( item, len( item ) - len( fieldName ) - 1 )
      Exit For
    end if
  next 
end function
%>
<!--#include file="footer_wrapper.asp"-->