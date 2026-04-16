<%
dim dictLanguageMobile  
set dictLanguageMobile=CreateObject("Scripting.Dictionary")

' english
dictLanguageMobile.Add "english_header_1", "Specials"  
dictLanguageMobile.Add "english_header_2", "Featured"
dictLanguageMobile.Add "english_header_3", "New"
dictLanguageMobile.Add "english_header_4", "Best Sellers"  

dictLanguageMobile.Add "english_footer_1", "Browse Catalog"  
dictLanguageMobile.Add "english_footer_2", "Use the full site to register a new account"
dictLanguageMobile.Add "english_footer_3", "Advanced search"
dictLanguageMobile.Add "english_footer_4", "Contact Us" 
dictLanguageMobile.Add "english_footer_5", "Full Web site" 

dictLanguageMobile.Add "english_opc_checkorv_1", "Enter both your E-mail and your Order Code."
dictLanguageMobile.Add "english_opc_checkorv_2", "Either the E-mail or the Order Code are invalid. "
dictLanguageMobile.Add "english_opc_checkorv_3", "Your account was suspended."
dictLanguageMobile.Add "english_opc_checkorv_4", "Your account was locked."

dictLanguageMobile.Add "english_opc_mobile1", "Coupons/Gift Certificates:"
dictLanguageMobile.Add "english_opc_mobile2", "Separate multiple codes with a comma."
dictLanguageMobile.Add "english_opc_mobile3", "Apply"
dictLanguageMobile.Add "english_opc_mobile4", "Place Order"


dictLanguageMobile.Add "english_Mobile_Back", "Return to Details"
dictLanguageMobile.Add "english_Mobile_MoreImgs",   "View more pictures"
dictLanguageMobile.Add "english_Mobile_reCalAlert", "We have detected that the discount information has changed. Please click the 'Recalculate' button if you would like to recalculate the order total using the new information."

' end language definitions
function clearLanguageMobile()
	' clear the dictionary.
	on error resume next
	clearLanguageMobile=dictLanguageMobile.removeAll   
	set clearLanguageMobile=nothing
end function
%>