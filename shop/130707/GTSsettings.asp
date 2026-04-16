<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce, Icon. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="../includes/SearchConstants.asp"-->
<%
pageTitle="Google Trusted Store Settings"
pageIcon="pcv4_icon_search.png"
section="layout"
%>
<%
Dim pcv_strPageName
pcv_strPageName="GTSsettings.asp"

msg=""

If request("action")="upd" Then
	gtsTurnOn=request("TurnOn")
	if gtsTurnOn="" then
		gtsTurnOn="0"
	end if
	gtsAccNo=request("AccNo")
	gtsShopAccID=request("shopAccID")
	gtsCountry=request("Country")
	gtsLang=request("Lang")
	gtsCur=request("Cur")
	gtsShipDays=request("ShipDays")
	if gtsShipDays="" then
		gtsShipDays="0"
	end if
	gtsDeDays=request("DeDays")
	if gtsDeDays="" then
		gtsDeDays="0"
	end if
	gtsPageLang=gtsLang & "_" & gtsCountry
	
	query="SELECT pcGTS_ID FROM pcGoogleTS;"
	set rs=connTemp.execute(query)
	if not rs.eof then
		query="UPDATE pcGoogleTS SET pcGTS_TurnOn=" & gtsTurnOn & ",pcGTS_AccNo='" & gtsAccNo & "',pcGTS_PageLang='" & gtsPageLang & "',pcGTS_ShopAccID='" & gtsShopAccID & "',pcGTS_ShopCountry='" & gtsCountry & "',pcGTS_ShopLang='" & gtsLang & "',pcGTS_Currency='" & gtsCur & "',pcGTS_ShipDays=" & gtsShipDays & ",pcGTS_DeDays=" & gtsDeDays & ";"
	else
		query="INSERT INTO pcGoogleTS (pcGTS_TurnOn,pcGTS_AccNo,pcGTS_PageLang,pcGTS_ShopAccID,pcGTS_ShopCountry,pcGTS_ShopLang,pcGTS_Currency,pcGTS_ShipDays,pcGTS_DeDays) VALUES (" & gtsTurnOn & ",'" & gtsAccNo & "','" & gtsPageLang & "','" & gtsShopAccID & "','" & gtsCountry & "','" & gtsLang & "','" & gtsCur & "'," & gtsShipDays & "," & gtsDeDays & ");"
	end if
	set rs=connTemp.execute(query)
	set rs=nothing
		
	msg="success"
End If

gtsTurnOn=0
gtsAccNo=""
gtsShopAccID=""
gtsCountry=""
gtsLang=""
gtsCur=""
gtsShipDays=1
gtsDeDays=1

query="SELECT pcGTS_TurnOn,pcGTS_AccNo,pcGTS_ShopAccID,pcGTS_ShopCountry,pcGTS_ShopLang,pcGTS_Currency,pcGTS_ShipDays,pcGTS_DeDays FROM pcGoogleTS;"
set rs=connTemp.execute(query)
if not rs.eof then
	gtsTurnOn=rs("pcGTS_TurnOn")
	gtsAccNo=rs("pcGTS_AccNo")
	gtsShopAccID=rs("pcGTS_ShopAccID")
	gtsCountry=rs("pcGTS_ShopCountry")
	gtsLang=rs("pcGTS_ShopLang")
	gtsCur=rs("pcGTS_Currency")
	gtsShipDays=rs("pcGTS_ShipDays")
	gtsDeDays=rs("pcGTS_DeDays")
end if
set rs=nothing
%>
<!--#include file="AdminHeader.asp"-->
<table class="pcCPcontent">
<% If msg="success" Then %>
<tr>
	<td align="center" colspan="2">
		<div class="pcCPmessageSuccess">
			<p>Google Trusted Store Settings Saved Successfully!</p>
		</div>
	</td>
</tr>
<tr>
	<td class="pcCPspacer" colspan="2"></td>
</tr>
<%End if%>

<form method="post" name="form1" action="<%=pcv_strPageName%>?action=upd" onSubmit="return Form1_Validator(this)" class="pcForms">
	<tr>
		<td class="pcCPspacer" colspan="2"></td>
	</tr>
	<tr>
		<td colspan="2">
			<p>This page allows you to update Google Trusted Store settings</p>
      </td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"><hr></td>
	</tr>
	<tr>
		<td>
        	Turn "Google Trusted Store" On:
    	</td>
		<td>
			<input name="TurnOn" type="checkbox" value="1" <%if gtsTurnOn="1" then%>checked<%end if%> class="clearBorder">
		</td>
	</tr>
	<tr>
		<td>
        	Google Trusted Store ID:
    	</td>
		<td>
			<input name="AccNo" type="text" size="30" value="<%=gtsAccNo%>">
		</td>
	</tr>
	<tr>
		<td nowrap>
        	Shopping Account ID:
    	</td>
		<td>
			<input name="shopAccID" type="text" size="30" value="<%=gtsShopAccID%>">&nbsp;(<em>optional</em>)
		</td>
	</tr>
	<tr>
		<td>
        	Shopping Country:
    	</td>
		<td>
			<select name="country">
				<option value=""></option>
				<option value="AD">Andorra</option>
				<option value="AE">United Arab Emirates</option>
				<option value="AF">Afghanistan</option>
				<option value="AG">Antigua and Barbuda</option>
				<option value="AI">Anguilla</option>
				<option value="AL">Albania</option>
				<option value="AM">Armenia</option>
				<option value="AO">Angola</option>
				<option value="AQ">Antarctica</option>
				<option value="AR">Argentina</option>
				<option value="AS">American Samoa</option>
				<option value="AT">Austria</option>
				<option value="AU">Australia</option>
				<option value="AW">Aruba</option>
				<option value="AX">Åland Islands</option>
				<option value="AZ">Azerbaijan</option>
				<option value="BA">Bosnia and Herzegovina</option>
				<option value="BB">Barbados</option>
				<option value="BD">Bangladesh</option>
				<option value="BE">Belgium</option>
				<option value="BF">Burkina Faso</option>
				<option value="BG">Bulgaria</option>
				<option value="BH">Bahrain</option>
				<option value="BI">Burundi</option>
				<option value="BJ">Benin</option>
				<option value="BL">Saint Barthélemy</option>
				<option value="BM">Bermuda</option>
				<option value="BN">Brunei Darussalam</option>
				<option value="BO">Bolivia, Plurinational State of</option>
				<option value="BQ">Bonaire, Sint Eustatius and Saba</option>
				<option value="BR">Brazil</option>
				<option value="BS">Bahamas</option>
				<option value="BT">Bhutan</option>
				<option value="BV">Bouvet Island</option>
				<option value="BW">Botswana</option>
				<option value="BY">Belarus</option>
				<option value="BZ">Belize</option>
				<option value="CA">Canada</option>
				<option value="CC">Cocos (Keeling) Islands</option>
				<option value="CD">Congo, the Democratic Republic of the</option>
				<option value="CF">Central African Republic</option>
				<option value="CG">Congo</option>
				<option value="CH">Switzerland</option>
				<option value="CI">Côte d'Ivoire</option>
				<option value="CK">Cook Islands</option>
				<option value="CL">Chile</option>
				<option value="CM">Cameroon</option>
				<option value="CN">China</option>
				<option value="CO">Colombia</option>
				<option value="CR">Costa Rica</option>
				<option value="CU">Cuba</option>
				<option value="CV">Cabo Verde</option>
				<option value="CW">Curaçao</option>
				<option value="CX">Christmas Island</option>
				<option value="CY">Cyprus</option>
				<option value="CZ">Czech Republic</option>
				<option value="DE">Germany</option>
				<option value="DJ">Djibouti</option>
				<option value="DK">Denmark</option>
				<option value="DM">Dominica</option>
				<option value="DO">Dominican Republic</option>
				<option value="DZ">Algeria</option>
				<option value="EC">Ecuador</option>
				<option value="EE">Estonia</option>
				<option value="EG">Egypt</option>
				<option value="EH">Western Sahara</option>
				<option value="ER">Eritrea</option>
				<option value="ES">Spain</option>
				<option value="ET">Ethiopia</option>
				<option value="FI">Finland</option>
				<option value="FJ">Fiji</option>
				<option value="FK">Falkland Islands (Malvinas)</option>
				<option value="FM">Micronesia, Federated States of</option>
				<option value="FO">Faroe Islands</option>
				<option value="FR">France</option>
				<option value="GA">Gabon</option>
				<option value="GB">United Kingdom</option>
				<option value="GD">Grenada</option>
				<option value="GE">Georgia</option>
				<option value="GF">French Guiana</option>
				<option value="GG">Guernsey</option>
				<option value="GH">Ghana</option>
				<option value="GI">Gibraltar</option>
				<option value="GL">Greenland</option>
				<option value="GM">Gambia</option>
				<option value="GN">Guinea</option>
				<option value="GP">Guadeloupe</option>
				<option value="GQ">Equatorial Guinea</option>
				<option value="GR">Greece</option>
				<option value="GS">South Georgia and the South Sandwich Islands</option>
				<option value="GT">Guatemala</option>
				<option value="GU">Guam</option>
				<option value="GW">Guinea-Bissau</option>
				<option value="GY">Guyana</option>
				<option value="HK">Hong Kong</option>
				<option value="HM">Heard Island and McDonald Islands</option>
				<option value="HN">Honduras</option>
				<option value="HR">Croatia</option>
				<option value="HT">Haiti</option>
				<option value="HU">Hungary</option>
				<option value="ID">Indonesia</option>
				<option value="IE">Ireland</option>
				<option value="IL">Israel</option>
				<option value="IM">Isle of Man</option>
				<option value="IN">India</option>
				<option value="IO">British Indian Ocean Territory</option>
				<option value="IQ">Iraq</option>
				<option value="IR">Iran, Islamic Republic of</option>
				<option value="IS">Iceland</option>
				<option value="IT">Italy</option>
				<option value="JE">Jersey</option>
				<option value="JM">Jamaica</option>
				<option value="JO">Jordan</option>
				<option value="JP">Japan</option>
				<option value="KE">Kenya</option>
				<option value="KG">Kyrgyzstan</option>
				<option value="KH">Cambodia</option>
				<option value="KI">Kiribati</option>
				<option value="KM">Comoros</option>
				<option value="KN">Saint Kitts and Nevis</option>
				<option value="KP">Korea, Democratic People's Republic of</option>
				<option value="KR">Korea, Republic of</option>
				<option value="KW">Kuwait</option>
				<option value="KY">Cayman Islands</option>
				<option value="KZ">Kazakhstan</option>
				<option value="LA">Lao People's Democratic Republic</option>
				<option value="LB">Lebanon</option>
				<option value="LC">Saint Lucia</option>
				<option value="LI">Liechtenstein</option>
				<option value="LK">Sri Lanka</option>
				<option value="LR">Liberia</option>
				<option value="LS">Lesotho</option>
				<option value="LT">Lithuania</option>
				<option value="LU">Luxembourg</option>
				<option value="LV">Latvia</option>
				<option value="LY">Libya</option>
				<option value="MA">Morocco</option>
				<option value="MC">Monaco</option>
				<option value="MD">Moldova, Republic of</option>
				<option value="ME">Montenegro</option>
				<option value="MF">Saint Martin (French part)</option>
				<option value="MG">Madagascar</option>
				<option value="MH">Marshall Islands</option>
				<option value="MK">Macedonia, the former Yugoslav Republic of</option>
				<option value="ML">Mali</option>
				<option value="MM">Myanmar</option>
				<option value="MN">Mongolia</option>
				<option value="MO">Macao</option>
				<option value="MP">Northern Mariana Islands</option>
				<option value="MQ">Martinique</option>
				<option value="MR">Mauritania</option>
				<option value="MS">Montserrat</option>
				<option value="MT">Malta</option>
				<option value="MU">Mauritius</option>
				<option value="MV">Maldives</option>
				<option value="MW">Malawi</option>
				<option value="MX">Mexico</option>
				<option value="MY">Malaysia</option>
				<option value="MZ">Mozambique</option>
				<option value="NA">Namibia</option>
				<option value="NC">New Caledonia</option>
				<option value="NE">Niger</option>
				<option value="NF">Norfolk Island</option>
				<option value="NG">Nigeria</option>
				<option value="NI">Nicaragua</option>
				<option value="NL">Netherlands</option>
				<option value="NO">Norway</option>
				<option value="NP">Nepal</option>
				<option value="NR">Nauru</option>
				<option value="NU">Niue</option>
				<option value="NZ">New Zealand</option>
				<option value="OM">Oman</option>
				<option value="PA">Panama</option>
				<option value="PE">Peru</option>
				<option value="PF">French Polynesia</option>
				<option value="PG">Papua New Guinea</option>
				<option value="PH">Philippines</option>
				<option value="PK">Pakistan</option>
				<option value="PL">Poland</option>
				<option value="PM">Saint Pierre and Miquelon</option>
				<option value="PN">Pitcairn</option>
				<option value="PR">Puerto Rico</option>
				<option value="PS">Palestine, State of</option>
				<option value="PT">Portugal</option>
				<option value="PW">Palau</option>
				<option value="PY">Paraguay</option>
				<option value="QA">Qatar</option>
				<option value="RE">Réunion</option>
				<option value="RO">Romania</option>
				<option value="RS">Serbia</option>
				<option value="RU">Russian Federation</option>
				<option value="RW">Rwanda</option>
				<option value="SA">Saudi Arabia</option>
				<option value="SB">Solomon Islands</option>
				<option value="SC">Seychelles</option>
				<option value="SD">Sudan</option>
				<option value="SE">Sweden</option>
				<option value="SG">Singapore</option>
				<option value="SH">Saint Helena, Ascension and Tristan da Cunha</option>
				<option value="SI">Slovenia</option>
				<option value="SJ">Svalbard and Jan Mayen</option>
				<option value="SK">Slovakia</option>
				<option value="SL">Sierra Leone</option>
				<option value="SM">San Marino</option>
				<option value="SN">Senegal</option>
				<option value="SO">Somalia</option>
				<option value="SR">Suriname</option>
				<option value="SS">South Sudan</option>
				<option value="ST">Sao Tome and Principe</option>
				<option value="SV">El Salvador</option>
				<option value="SX">Sint Maarten (Dutch part)</option>
				<option value="SY">Syrian Arab Republic</option>
				<option value="SZ">Swaziland</option>
				<option value="TC">Turks and Caicos Islands</option>
				<option value="TD">Chad</option>
				<option value="TF">French Southern Territories</option>
				<option value="TG">Togo</option>
				<option value="TH">Thailand</option>
				<option value="TJ">Tajikistan</option>
				<option value="TK">Tokelau</option>
				<option value="TL">Timor-Leste</option>
				<option value="TM">Turkmenistan</option>
				<option value="TN">Tunisia</option>
				<option value="TO">Tonga</option>
				<option value="TR">Turkey</option>
				<option value="TT">Trinidad and Tobago</option>
				<option value="TV">Tuvalu</option>
				<option value="TW">Taiwan, Province of China</option>
				<option value="TZ">Tanzania, United Republic of</option>
				<option value="UA">Ukraine</option>
				<option value="UG">Uganda</option>
				<option value="UM">United States Minor Outlying Islands</option>
				<option value="US">United States</option>
				<option value="UY">Uruguay</option>
				<option value="UZ">Uzbekistan</option>
				<option value="VA">Holy See (Vatican City State)</option>
				<option value="VC">Saint Vincent and the Grenadines</option>
				<option value="VE">Venezuela, Bolivarian Republic of</option>
				<option value="VG">Virgin Islands, British</option>
				<option value="VI">Virgin Islands, U.S.</option>
				<option value="VN">Viet Nam</option>
				<option value="VU">Vanuatu</option>
				<option value="WF">Wallis and Futuna</option>
				<option value="WS">Samoa</option>
				<option value="YE">Yemen</option>
				<option value="YT">Mayotte</option>
				<option value="ZA">South Africa</option>
				<option value="ZM">Zambia</option>
				<option value="ZW">Zimbabwe</option>			
			</select>
		</td>
	</tr>
	<tr>
		<td>
        	Shopping Language:
    	</td>
		<td>
			<select name="lang">
				<option value=""></option>
				<option value="ab">Abkhaz</option>
				<option value="aa">Afar</option>
				<option value="af">Afrikaans</option>
				<option value="ak">Akan</option>
				<option value="sq">Albanian</option>
				<option value="am">Amharic</option>
				<option value="ar">Arabic</option>
				<option value="an">Aragonese</option>
				<option value="hy">Armenian</option>
				<option value="as">Assamese</option>
				<option value="av">Avaric</option>
				<option value="ae">Avestan</option>
				<option value="ay">Aymara</option>
				<option value="az">Azerbaijani</option>
				<option value="bm">Bambara</option>
				<option value="ba">Bashkir</option>
				<option value="eu">Basque</option>
				<option value="be">Belarusian</option>
				<option value="bn">Bengali, Bangla</option>
				<option value="bh">Bihari</option>
				<option value="bi">Bislama</option>
				<option value="bs">Bosnian</option>
				<option value="br">Breton</option>
				<option value="bg">Bulgarian</option>
				<option value="my">Burmese</option>
				<option value="ca">Catalan, Valencian</option>
				<option value="ch">Chamorro</option>
				<option value="ce">Chechen</option>
				<option value="ny">Chichewa, Chewa, Nyanja</option>
				<option value="zh">Chinese</option>
				<option value="cv">Chuvash</option>
				<option value="kw">Cornish</option>
				<option value="co">Corsican</option>
				<option value="cr">Cree</option>
				<option value="hr">Croatian</option>
				<option value="cs">Czech</option>
				<option value="da">Danish</option>
				<option value="dv">Divehi, Dhivehi, Maldivian</option>
				<option value="nl">Dutch</option>
				<option value="dz">Dzongkha</option>
				<option value="en">English</option>
				<option value="eo">Esperanto</option>
				<option value="et">Estonian</option>
				<option value="ee">Ewe</option>
				<option value="fo">Faroese</option>
				<option value="fj">Fijian</option>
				<option value="fi">Finnish</option>
				<option value="fr">French</option>
				<option value="ff">Fula, Fulah, Pulaar, Pular</option>
				<option value="gl">Galician</option>
				<option value="ka">Georgian</option>
				<option value="de">German</option>
				<option value="el">Greek (modern)</option>
				<option value="gn">Guaraní</option>
				<option value="gu">Gujarati</option>
				<option value="ht">Haitian, Haitian Creole</option>
				<option value="ha">Hausa</option>
				<option value="he">Hebrew (modern)</option>
				<option value="hz">Herero</option>
				<option value="hi">Hindi</option>
				<option value="ho">Hiri Motu</option>
				<option value="hu">Hungarian</option>
				<option value="ia">Interlingua</option>
				<option value="id">Indonesian</option>
				<option value="ie">Interlingue</option>
				<option value="ga">Irish</option>
				<option value="ig">Igbo</option>
				<option value="ik">Inupiaq</option>
				<option value="io">Ido</option>
				<option value="is">Icelandic</option>
				<option value="it">Italian</option>
				<option value="iu">Inuktitut</option>
				<option value="ja">Japanese</option>
				<option value="jv">Javanese</option>
				<option value="kl">Kalaallisut, Greenlandic</option>
				<option value="kn">Kannada</option>
				<option value="kr">Kanuri</option>
				<option value="ks">Kashmiri</option>
				<option value="kk">Kazakh</option>
				<option value="km">Khmer</option>
				<option value="ki">Kikuyu, Gikuyu</option>
				<option value="rw">Kinyarwanda</option>
				<option value="ky">Kyrgyz</option>
				<option value="kv">Komi</option>
				<option value="kg">Kongo</option>
				<option value="ko">Korean</option>
				<option value="ku">Kurdish</option>
				<option value="kj">Kwanyama, Kuanyama</option>
				<option value="la">Latin</option>
				<option value="lb">Luxembourgish, Letzeburgesch</option>
				<option value="lg">Ganda</option>
				<option value="li">Limburgish, Limburgan, Limburger</option>
				<option value="ln">Lingala</option>
				<option value="lo">Lao</option>
				<option value="lt">Lithuanian</option>
				<option value="lu">Luba-Katanga</option>
				<option value="lv">Latvian</option>
				<option value="gv">Manx</option>
				<option value="mk">Macedonian</option>
				<option value="mg">Malagasy</option>
				<option value="ms">Malay</option>
				<option value="ml">Malayalam</option>
				<option value="mt">Maltese</option>
				<option value="mi">Maori</option>
				<option value="mr">Marathi (Mara?hi)</option>
				<option value="mh">Marshallese</option>
				<option value="mn">Mongolian</option>
				<option value="na">Nauru</option>
				<option value="nv">Navajo, Navaho</option>
				<option value="nd">Northern Ndebele</option>
				<option value="ne">Nepali</option>
				<option value="ng">Ndonga</option>
				<option value="nb">Norwegian Bokmål</option>
				<option value="nn">Norwegian Nynorsk</option>
				<option value="no">Norwegian</option>
				<option value="ii">Nuosu</option>
				<option value="nr">Southern Ndebele</option>
				<option value="oc">Occitan</option>
				<option value="oj">Ojibwe, Ojibwa</option>
				<option value="cu">Old Church Slavonic, Church Slavonic, Old Bulgarian</option>
				<option value="om">Oromo</option>
				<option value="or">Oriya</option>
				<option value="os">Ossetian, Ossetic</option>
				<option value="pa">Panjabi, Punjabi</option>
				<option value="pi">Pali</option>
				<option value="fa">Persian (Farsi)</option>
				<option value="pl">Polish</option>
				<option value="ps">Pashto, Pushto</option>
				<option value="pt">Portuguese</option>
				<option value="qu">Quechua</option>
				<option value="rm">Romansh</option>
				<option value="rn">Kirundi</option>
				<option value="ro">Romanian</option>
				<option value="ru">Russian</option>
				<option value="sa">Sanskrit (Sa?sk?ta)</option>
				<option value="sc">Sardinian</option>
				<option value="sd">Sindhi</option>
				<option value="se">Northern Sami</option>
				<option value="sm">Samoan</option>
				<option value="sg">Sango</option>
				<option value="sr">Serbian</option>
				<option value="gd">Scottish Gaelic, Gaelic</option>
				<option value="sn">Shona</option>
				<option value="si">Sinhala, Sinhalese</option>
				<option value="sk">Slovak</option>
				<option value="sl">Slovene</option>
				<option value="so">Somali</option>
				<option value="st">Southern Sotho</option>
				<option value="es">Spanish, Castilian</option>
				<option value="su">Sundanese</option>
				<option value="sw">Swahili</option>
				<option value="ss">Swati</option>
				<option value="sv">Swedish</option>
				<option value="ta">Tamil</option>
				<option value="te">Telugu</option>
				<option value="tg">Tajik</option>
				<option value="th">Thai</option>
				<option value="ti">Tigrinya</option>
				<option value="bo">Tibetan Standard, Tibetan, Central</option>
				<option value="tk">Turkmen</option>
				<option value="tl">Tagalog</option>
				<option value="tn">Tswana</option>
				<option value="to">Tonga (Tonga Islands)</option>
				<option value="tr">Turkish</option>
				<option value="ts">Tsonga</option>
				<option value="tt">Tatar</option>
				<option value="tw">Twi</option>
				<option value="ty">Tahitian</option>
				<option value="ug">Uyghur, Uighur</option>
				<option value="uk">Ukrainian</option>
				<option value="ur">Urdu</option>
				<option value="uz">Uzbek</option>
				<option value="ve">Venda</option>
				<option value="vi">Vietnamese</option>
				<option value="vo">Volapük</option>
				<option value="wa">Walloon</option>
				<option value="cy">Welsh</option>
				<option value="wo">Wolof</option>
				<option value="fy">Western Frisian</option>
				<option value="xh">Xhosa</option>
				<option value="yi">Yiddish</option>
				<option value="yo">Yoruba</option>
				<option value="za">Zhuang, Chuang</option>
				<option value="zu">Zulu</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>
        	Shopping Currency:
    	</td>
		<td>
			<select name="Cur">
				<option value=""></option>
				<option value="AED">United Arab Emirates dirham</option>
				<option value="AFN">Afghan afghani</option>
				<option value="ALL">Albanian lek</option>
				<option value="AMD">Armenian dram</option>
				<option value="ANG">Netherlands Antillean guilder</option>
				<option value="AOA">Angolan kwanza</option>
				<option value="ARS">Argentine peso</option>
				<option value="AUD">Australian dollar</option>
				<option value="AWG">Aruban florin</option>
				<option value="AZN">Azerbaijani manat</option>
				<option value="BAM">Bosnia and Herzegovina convertible mark</option>
				<option value="BBD">Barbados dollar</option>
				<option value="BDT">Bangladeshi taka</option>
				<option value="BGN">Bulgarian lev</option>
				<option value="BHD">Bahraini dinar</option>
				<option value="BIF">Burundian franc</option>
				<option value="BMD">Bermudian dollar</option>
				<option value="BND">Brunei dollar</option>
				<option value="BOB">Boliviano</option>
				<option value="BRL">Brazilian real</option>
				<option value="BSD">Bahamian dollar</option>
				<option value="BTN">Bhutanese ngultrum</option>
				<option value="BWP">Botswana pula</option>
				<option value="BYR">Belarusian ruble</option>
				<option value="BZD">Belize dollar</option>
				<option value="CAD">Canadian dollar</option>
				<option value="CDF">Congolese franc</option>
				<option value="CHF">Swiss franc</option>
				<option value="CLP">Chilean peso</option>
				<option value="CNY">Chinese yuan</option>
				<option value="COP">Colombian peso</option>
				<option value="CRC">Costa Rican colon</option>
				<option value="CUC">Cuban convertible peso</option>
				<option value="CUP">Cuban peso</option>
				<option value="CVE">Cape Verde escudo</option>
				<option value="CZK">Czech koruna</option>
				<option value="DJF">Djiboutian franc</option>
				<option value="DKK">Danish krone</option>
				<option value="DOP">Dominican peso</option>
				<option value="DZD">Algerian dinar</option>
				<option value="EGP">Egyptian pound</option>
				<option value="ERN">Eritrean nakfa</option>
				<option value="ETB">Ethiopian birr</option>
				<option value="EUR">Euro</option>
				<option value="FJD">Fiji dollar</option>
				<option value="FKP">Falkland Islands pound</option>
				<option value="GBP">Pound sterling</option>
				<option value="GEL">Georgian lari</option>
				<option value="GHS">Ghanaian cedi</option>
				<option value="GIP">Gibraltar pound</option>
				<option value="GMD">Gambian dalasi</option>
				<option value="GNF">Guinean franc</option>
				<option value="GTQ">Guatemalan quetzal</option>
				<option value="GYD">Guyanese dollar</option>
				<option value="HKD">Hong Kong dollar</option>
				<option value="HNL">Honduran lempira</option>
				<option value="HRK">Croatian kuna</option>
				<option value="HTG">Haitian gourde</option>
				<option value="HUF">Hungarian forint</option>
				<option value="IDR">Indonesian rupiah</option>
				<option value="ILS">Israeli new shekel</option>
				<option value="INR">Indian rupee</option>
				<option value="IQD">Iraqi dinar</option>
				<option value="IRR">Iranian rial</option>
				<option value="ISK">Icelandic króna</option>
				<option value="JMD">Jamaican dollar</option>
				<option value="JOD">Jordanian dinar</option>
				<option value="JPY">Japanese yen</option>
				<option value="KES">Kenyan shilling</option>
				<option value="KGS">Kyrgyzstani som</option>
				<option value="KHR">Cambodian riel</option>
				<option value="KMF">Comoro franc</option>
				<option value="KPW">North Korean won</option>
				<option value="KRW">South Korean won</option>
				<option value="KWD">Kuwaiti dinar</option>
				<option value="KYD">Cayman Islands dollar</option>
				<option value="KZT">Kazakhstani tenge</option>
				<option value="LAK">Lao kip</option>
				<option value="LBP">Lebanese pound</option>
				<option value="LKR">Sri Lankan rupee</option>
				<option value="LRD">Liberian dollar</option>
				<option value="LSL">Lesotho loti</option>
				<option value="LTL">Lithuanian litas</option>
				<option value="LYD">Libyan dinar</option>
				<option value="MAD">Moroccan dirham</option>
				<option value="MDL">Moldovan leu</option>
				<option value="MGA">Malagasy ariary</option>
				<option value="MKD">Macedonian denar</option>
				<option value="MMK">Myanmar kyat</option>
				<option value="MNT">Mongolian tugrik</option>
				<option value="MOP">Macanese pataca</option>
				<option value="MRO">Mauritanian ouguiya</option>
				<option value="MUR">Mauritian rupee</option>
				<option value="MVR">Maldivian rufiyaa</option>
				<option value="MWK">Malawian kwacha</option>
				<option value="MXN">Mexican peso</option>
				<option value="MYR">Malaysian ringgit</option>
				<option value="MZN">Mozambican metical</option>
				<option value="NAD">Namibian dollar</option>
				<option value="NGN">Nigerian naira</option>
				<option value="NIO">Nicaraguan córdoba</option>
				<option value="NOK">Norwegian krone</option>
				<option value="NPR">Nepalese rupee</option>
				<option value="NZD">New Zealand dollar</option>
				<option value="OMR">Omani rial</option>
				<option value="PAB">Panamanian balboa</option>
				<option value="PEN">Peruvian nuevo sol</option>
				<option value="PGK">Papua New Guinean kina</option>
				<option value="PHP">Philippine peso</option>
				<option value="PKR">Pakistani rupee</option>
				<option value="PLN">Polish złoty</option>
				<option value="PYG">Paraguayan guaraní</option>
				<option value="QAR">Qatari riyal</option>
				<option value="RON">Romanian new leu</option>
				<option value="RSD">Serbian dinar</option>
				<option value="RUB">Russian ruble</option>
				<option value="RWF">Rwandan franc</option>
				<option value="SAR">Saudi riyal</option>
				<option value="SBD">Solomon Islands dollar</option>
				<option value="SCR">Seychelles rupee</option>
				<option value="SDG">Sudanese pound</option>
				<option value="SEK">Swedish krona/kronor</option>
				<option value="SGD">Singapore dollar</option>
				<option value="SHP">Saint Helena pound</option>
				<option value="SLL">Sierra Leonean leone</option>
				<option value="SOS">Somali shilling</option>
				<option value="SRD">Surinamese dollar</option>
				<option value="SSP">South Sudanese pound</option>
				<option value="STD">São Tomé and Príncipe dobra</option>
				<option value="SYP">Syrian pound</option>
				<option value="SZL">Swazi lilangeni</option>
				<option value="THB">Thai baht</option>
				<option value="TJS">Tajikistani somoni</option>
				<option value="TMT">Turkmenistani manat</option>
				<option value="TND">Tunisian dinar</option>
				<option value="TOP">Tongan paʻanga</option>
				<option value="TRY">Turkish lira</option>
				<option value="TTD">Trinidad and Tobago dollar</option>
				<option value="TWD">New Taiwan dollar</option>
				<option value="TZS">Tanzanian shilling</option>
				<option value="UAH">Ukrainian hryvnia</option>
				<option value="UGX">Ugandan shilling</option>
				<option value="USD">United States dollar</option>
				<option value="UYU">Uruguayan peso</option>
				<option value="UZS">Uzbekistan som</option>
				<option value="VEF">Venezuelan bolívar</option>
				<option value="VND">Vietnamese dong</option>
				<option value="VUV">Vanuatu vatu</option>
				<option value="WST">Samoan tala</option>
				<option value="XCD">East Caribbean dollar</option>
				<option value="YER">Yemeni rial</option>
				<option value="ZAR">South African rand</option>
				<option value="ZMW">Zambian kwacha</option>
				<option value="ZWD">Zimbabwe dollar</option>
			</select>
		</td>
	</tr>
	<tr>
		<td>
        	Order(s) will be shipped after:
    	</td>
		<td>
			<select name="ShipDays">
				<option value="1">1</option>
				<option value="2">2</option>
				<option value="3">3</option>
				<option value="4">4</option>
				<option value="5">5</option>
				<option value="6">6</option>
				<option value="7">7</option>
				<option value="8">8</option>
				<option value="9">9</option>
				<option value="10">10</option>
				<option value="11">11</option>
				<option value="12">12</option>
				<option value="13">13</option>
				<option value="14">14</option>
				<option value="15">15</option>
			</select>&nbsp;days
		</td>
	</tr>
	<tr>
		<td>
        	Estimated Delivery Date:
    	</td>
		<td>
			<select name="DeDays">
				<option value="1">1</option>
				<option value="2">2</option>
				<option value="3">3</option>
				<option value="4">4</option>
				<option value="5">5</option>
				<option value="6">6</option>
				<option value="7">7</option>
				<option value="8">8</option>
				<option value="9">9</option>
				<option value="10">10</option>
				<option value="11">11</option>
				<option value="12">12</option>
				<option value="13">13</option>
				<option value="14">14</option>
				<option value="15">15</option>
			</select>&nbsp;days after shipping
		</td>
	</tr>
	<tr>
		<td class="pcCPspacer" colspan="2"><hr></td>
	</tr>
	<tr> 
		<td style="text-align: center;" colspan="2">
			<input name="submit" type="submit" class="btn btn-primary" value="Update Settings">&nbsp;
    </td>
	</tr>
	</table>
</form>
<script type="text/javascript">
function Form1_Validator(theForm)
{
	if (theForm.AccNo.value == "")
 	{
		    alert("Please enter a value for the 'Google Trusted Store ID'.");
		    theForm.AccNo.focus();
		    return (false);
	}
	if (theForm.country.value == "")
 	{
		    alert("Please select a value for the 'Country'.");
		    theForm.country.focus();
		    return (false);
	}
	if (theForm.lang.value == "")
 	{
		    alert("Please select a value for the 'Language'.");
		    theForm.lang.focus();
		    return (false);
	}
	if (theForm.Cur.value == "")
 	{
		    alert("Please select a value for the 'Currency'.");
		    theForm.Cur.focus();
		    return (false);
	}
	if (theForm.ShipDays.value == "")
 	{
		    alert("Please select a value for the 'Shipping Days'.");
		    theForm.ShipDays.focus();
		    return (false);
	}
	if (theForm.DeDays.value == "")
 	{
		    alert("Please select a value for the 'Shipping Days'.");
		    theForm.ShipDays.focus();
		    return (false);
	}
	return (true);
}

<%if gtsCountry<>"" then%>
	document.form1.country.value="<%=gtsCountry%>";
<%end if%>
<%if gtsLang<>"" then%>
	document.form1.lang.value="<%=gtsLang%>";
<%end if%>
<%if gtsCur<>"" then%>
	document.form1.Cur.value="<%=gtsCur%>";
<%end if%>
<%if gtsShipDays<>"" then%>
	document.form1.ShipDays.value="<%=gtsShipDays%>";
<%end if%>
<%if gtsDeDays<>"" then%>
	document.form1.DeDays.value="<%=gtsDeDays%>";
<%end if%>
</script>
<!--#include file="AdminFooter.asp"-->
