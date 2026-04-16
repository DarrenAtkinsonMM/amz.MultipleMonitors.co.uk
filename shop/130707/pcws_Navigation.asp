<div class="container-fluid">

    <% If pcPageName = "pcws_Settings.asp" Then %>    
        <div class="row">
            <div class="col-xs-4">
                
                <a href="pcws_MyApps.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-th-large" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>
                </a>
                <a href="pcws_Market.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-shopping-cart" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>
                </a>  
                
            </div>
            <div class="col-xs-8">
                <div ng-bind-html="error"></div>
            </div>
        </div>    
    <% End If %>

    <% If pcPageName = "pcws_MyAccount.asp" Then %>    
        <div class="row">
            <div class="col-xs-4">
                
                <a href="pcws_MyApps.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-th-large" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>
                </a>
                <a href="pcws_Market.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-shopping-cart" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>
                </a>  
                
            </div>
            <div class="col-xs-8">
                <div ng-bind-html="error"></div>
            </div>
        </div>    
    <% End If %>
    
    <% If pcPageName = "pcws_Market.asp" Then %>
        <div class="row">
            <div class="col-xs-4">
                
                <a href="pcws_MyAccount.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnMyAccount") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-home" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnMyAccount") %>
                </a>
                <a href="pcws_MyApps.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-th-large" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnManage") %>
                </a>
                
            </div>
            <div class="col-xs-8">
                <div ng-bind-html="error"></div>
            </div>
        </div>  
    <% End If %>
    
    <% If pcPageName = "pcws_MyApps.asp" Then %>
        <div class="row">
            <div class="col-xs-4">
                
                <a href="pcws_MyAccount.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnMyAccount") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-home" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnMyAccount") %>
                </a>
                <a href="pcws_Market.asp" title="<%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>" class="btn btn-default">
                    <span class="glyphicon glyphicon-shopping-cart" aria-hidden="true"></span>
                    <%=dictLanguage.Item(Session("language")&"_pcAppBtnMarket") %>
                </a>  
                
            </div>
            <div class="col-xs-8">
                <div ng-bind-html="error"></div>
            </div>
        </div>
    <% End If %>

</div>