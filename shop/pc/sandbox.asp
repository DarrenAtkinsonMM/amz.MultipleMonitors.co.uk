<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "sandbox.asp"
' This page displays the items in the cart.
'
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="header_wrapper.asp"-->

<style>
/*
 * Examples
 *
 * Isolated sections of example content for each component or feature. Usually
 * followed by a code snippet.
 */

.bs-example {
  position: relative;
  padding: 45px 15px 15px;
  margin-top: 10px;
  margin-bottom: 20px;
  background-color: #fafafa;
  box-shadow: inset 0 3px 6px rgba(0,0,0,.05);
  border-color: #e5e5e5 #eee #eee;
  border-style: solid;
  border-width: 1px 0;
}
/* Echo out a label for the example */
.bs-example:after {
  content: "Example";
  position: absolute;
  top:  15px;
  left: 15px;
  font-size: 12px;
  font-weight: bold;
  color: #bbb;
  text-transform: uppercase;
  letter-spacing: 1px;
}

/* Tweak display of the code snippets when following an example */
.bs-example + .highlight {
  margin: -15px -15px 15px;
  border-radius: 0;
  border-width: 0 0 1px;
}
</style>

<div id="pcMain">

    <div class="pcMainContent">

        <h1>Typography</h1>

        <h2>Headings</h2>
        <div class="bs-example">
            <h1>Heading 1</h1>
            <h2>Heading 2</h2>
            <h3>Heading 3</h3>
            <h4>Heading 4</h4>       
        </div>
        
        <h2>Font Sizes</h2>
        <div class="bs-example">
            <div class="pcLargerText">large text (class=pcLargerText)</div>
            <div class="pcSpacer"></div>
            <div class="pcSmallText">small text (class=pcSmallText)</div>     
        </div>
        
        <h2>Lists</h2>
        <div class="bs-example">
            <ul>
                <li>List Item 1</li>
                <li>List Item 2</li>
            </ul>     
        </div>

        <div class="pcSpacer"></div>

        <h1>Design</h1>

        <h2>Messages Styles</h2>
        <div class="bs-example">
            <div class="pcInfoMessage">This is information. class=pcInfoMessage</div>
            <div class="pcAttention">This is a message to get your attention. class=pcAttention</div>
            <div class="pcSuccessMessage">This is a message for successful actions. class=pcSuccessMessage</div>
            <div class="pcErrorMessage">This is a message for errors. class=pcErrorMessage</div>
            <div class="pcPromoMessage">This is a message for promotions. class=pcPromoMessage</div>
        </div>
        
        <h2>Button Styles</h2>
        <div class="bs-example">
            
            <h3>Normal Button</h3>
            <div class="pcFormButtons">
                <a class="pcButton pcButtonContinue">
                    <img src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" />
                    <span class="pcButtonText">My Button</span>        
                </a>
            </div>
            
            <h3>Secondary Button</h3>
            <div class="pcFormButtons">
                <a class="pcButton secondary">
                    <img src="<%=pcf_getImagePath("",RSlayout("pcLO_Update"))%>" alt="Submit" />
                    <span class="pcButtonText">My Button</span>        
                </a>
            </div>            
            
          
        </div>
        
        <div class="pcSpacer"></div>

        <h1>Layout</h1>
        
        <h2>Basic Table</h2>
        <div class="bs-example">
        
            <div class="pcSpacer"></div>
            
            <div class="pcTable">
              <div class="pcTableHeader">
                <div class="pcTableColumnLeft">Column 1</div>
                <div class="pcTableColumnRight">Column 2</div>
              </div>
              <div class="pcTableRow">
                <div class="pcTableColumnLeft">Row 1 Column 1</div>
                <div class="pcTableColumnRight">Row 1 Column 2</div>
              </div>
              <div class="pcTableRow">
                <div class="pcTableRowFull"> This is a full row...................................................................................</div>   
              </div>
            </div>
            
             <div class="pcSpacer"></div>

            
        </div>
        
        
        <h2>Forms</h2>
        <div class="bs-example">
        
            <div class="pcFormItem"> 
                <div class="pcFormLabel">Label</div>
                <div class="pcFormField">
                    <input class="form-control input-sm" type="text" id="email" name="email" size="25">
                </div> 
            </div> 
            
        </div>     
        
        
   
        
        
        
    </div>

</div>
<!--#include file="footer_wrapper.asp"-->
