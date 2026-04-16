						<% 
						' This file is included in most payment gateway files
						' It provides the code for the "Back" and "Place Order" buttons shown at the
						' bottom of the payment form.
						 
						 
						
							If scSSL="1" And scIntSSLPage="1" Then
								tempURL=replace((scSslURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
								tempURL=replace(tempURL,"https:/","https://")
								tempURL=replace(tempURL,"http:/","http://")
							Else
								tempURL=replace((scStoreURL&"/"&scPcFolder&"/pc/OnePageCheckout.asp?PayPanel=1"),"//","/")
								tempURL=replace(tempURL,"https:/","https://")
								tempURL=replace(tempURL,"http:/","http://")
							End If
						 
						 if pcStrPageName = "manpay2" then
						 	tempURL="https://www.multiplemonitors.co.uk/shop/pc/manpayForm.asp?amount="&pcBillingTotal&"&pay="&pcPaymentDesc
						 	strButtonText="Continue To Payment"
						 Else
						 	strButtonText="Submit Payment & Place Order"
						 end if
						%> 

            <%
              if (Session("SBEditOrder")<>"") AND (Session("SBEditOrderID")<>"") then
                buttonText = dictLanguage.Item(Session("language")&"_css_pcLO_update")
                buttonImage = rslayout("pcLO_Update")
                buttonClass = "pcButtonContinue"
              else
                buttonText = dictLanguage.Item(Session("language")&"_css_pcLO_placeorder")
                buttonImage = rslayout("pcLO_placeOrder")
                buttonClass = "pcButtonPlaceOrder"
              end if
            %>

            <button class="<%= "pcButton " & buttonClass %> btn btn-skin btn-wc btn-pay" id="submit">
              <span class="pcButtonText"><%=strButtonText%></span>
            </button>

						<!--<input type="image" name="Continue">-->

						<script type=text/javascript>
              $pc(document).ready(function() {
                  $pc('#submit', this).attr('disabled', false);
                  $pc('form').submit(function(){
                      $pc('#submit', this).attr('disabled', true);
                      return 
                  });
              });
            </script> 
            <a class="daOPCCCBackLink" href="<%=tempURL%>">
              <span class="pcButtonText">Cancel and return to checkout page</span>
            </a>
