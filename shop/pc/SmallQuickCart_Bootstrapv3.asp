<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>   
<% If Instr(Ucase(Request.ServerVariables("SCRIPT_NAME")), "GW") = 0 Then %> 
  
<div class="dropdown ng-cloak dropblock" data-ng-controller="QuickCartCtrl" data-ng-cloak>

    <a data-ng-hide="shoppingcart.totalQuantity>0" href="#">
        <%=dictLanguage.Item(Session("language")&"_smallcart_11")%>
    </a>  

    <a href="#" class="dropdown-toggle" data-toggle="dropdown" data-ng-show="shoppingcart.totalQuantity>0">
        <span class="cartbox">
            <span class="carboxCount">{{shoppingcart.totalQuantityDisplay}}</span>
        </span>
        <%=dictLanguage.Item(Session("language")&"_showcart_12")%>
        <span data-ng-show="!Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.subtotal}}</span>
        <span data-ng-show="Evaluate(shoppingcart.checkoutStage)">{{shoppingcart.total}}</span> <b class="caret"></b>
    </a>
    <div class="dropdown-menu ng-cloak" data-ng-show="shoppingcart.totalQuantity>0" data-ng-cloak>	


        <div class="pcCartDropDown pcCartLayout pcTable" data-ng-show="shoppingcart.totalQuantity>0">

            <div class="pcTableRow" data-ng-repeat="shoppingcartitem in shoppingcart.shoppingcartrow | limitTo: 10 | filter:{productID: '!!'}">                
        
                <% '// START Main Product Data %>                        
                <div class="pcTableRow pcCartRowMain">
                    <div class="pcQuickCartImage">
                    
                        <div data-ng-show="!Evaluate(shoppingcart.IsBuyGift)">
                            <a data-ng-show="Evaluate(shoppingcartitem.ShowImage);" rel="nofollow" data-ng-href="{{shoppingcartitem.productURL}}"><img src="<%=pcf_getImagePath("catalog","no_image.gif")%>" data-ng-src="catalog/{{shoppingcartitem.ImageURL}}" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")%>{{shoppingcartitem.description}}"></a>                                    
                        </div>

                        <div data-ng-show="Evaluate(shoppingcart.IsBuyGift)">
                            <a data-ng-show="Evaluate(shoppingcartitem.ShowImage);" data-ng-href="ggg_viewEP.asp?grCode={{shoppingcart.grCode}}&amp;geID={{shoppingcartitem.geID}}"><img src="<%=pcf_getImagePath("catalog","no_image.gif")%>" data-ng-src="catalog/{{shoppingcartitem.ImageURL}}" alt="<%=dictLanguage.Item(Session("language")&"_altTag_1")%>{{shoppingcartitem.description}}"></a>                                    
                        </div>
                        
                    </div>
                    <div class="pcQuickCartDescription">
                        <a class="pcItemDescription title bold" rel="nofollow" data-ng-href="{{shoppingcartitem.productURL}}"><span data-ng-bind-html="shoppingcartitem.description|unsafe">{{shoppingcartitem.description}}</span></a>
                        <br />
                        <span class="pcQuickCartQtyText"><%=dictLanguage.Item(Session("language")&"_showcart_4")%> {{shoppingcartitem.quantity}}</span>
                        
                        <% '// START Product Options %>    
                        <div class="pcViewCartOptions" data-ng-repeat="productoption in shoppingcartitem.productoptions">                            
                            <span class="small">{{productoption.name}}</span>
                        </div>    
                        <% '// END Product Options %>
                        
                    </div>  
                                               
                </div>
                <% '// END Main Product Data %>

                <div class="pcTableRow row-divider"></div>
                
            </div>  

                   
            <div class="pcQuickCartButtons">
            
                <div class="pcButton pcButtonViewCart" data-ng-click="viewCart()">
                  <img src="<%=viewcartbtn%>" alt="<%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %>">
                  <span class="pcButtonText"><%= dictLanguage.Item(Session("language")&"_css_viewcartbtn") %> {{shoppingcart.totalQuantityDisplay}} <%= dictLanguage.Item(Session("language")&"_smallcart_2") %></span>
                </div>
            
            </div>  

        </div> 

    </div> 

</div>
<% End If %>
