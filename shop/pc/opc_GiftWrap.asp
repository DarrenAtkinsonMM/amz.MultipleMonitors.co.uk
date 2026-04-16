<%@  language="VBSCRIPT" %>
<%
'--------------------------------------------------------------
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce.
'ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce.
'Copyright 2001-2015. All rights reserved. You are not allowed to use, alter,
'distribute and/or resell any parts of ProductCart's source code without the written consent of 
'NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
'--------------------------------------------------------------
Dim pcStrPageName
pcStrPageName = "opc_GiftWrap.asp"
%>
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<!--#include file="opc_contentType.asp" -->
<%
HaveSecurity=0
if session("idCustomer")=0 OR session("idCustomer")="" then
	HaveSecurity=1
end if

dim pcCartArray, ppcCartIndex, f, cont

Call SetContentType()

IF HaveSecurity=0 THEN

	'*****************************************************************************************************
	'// START: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
    %><!--#include file="pcVerifySession.asp"--><%
	pcs_VerifySession
	'*****************************************************************************************************
	'// END: Validate AND Set "pcCartArray" AND "pcCartIndex"
	'*****************************************************************************************************
	
	pIdCustomer=session("idCustomer")
	ppcCartIndex=Session("pcCartIndex")
	pcCartArray=Session("pcCartSession")
	
	f=getUserInput(request("index"),0)
	
	if pcCartArray(f,34)="" then
		pcCartArray(f,34)="0"
	end if

	UpdateSuccess="0"
	tmpPrdList=""
	if request("action")="add" then
		
		if pcCartArray(f,10)=0 then
			GW=getUserInput(request("GW" & f),0)
			pcCartArray(f,34)=GW
			if GW<>"" AND GW<>"0" then
				if tmpPrdList<>"" then
					tmpPrdList=tmpPrdList & ","
				end if
				tmpPrdList=tmpPrdList & pcCartArray(f,0)
			end if
			GWText=URLDecode(getUserInput(request("GWText" & f),240))
			if GWText<>"" then
				GWText=replace(GWText,"''","'")
			end if
			pcCartArray(f,35)=GWText
		end if

		session("pcCartSession")=pcCartArray
		UpdateSuccess="1"
	end if

END IF
%>
<script type=text/javascript>
    function Form1_Validator(theForm)
    {
        return (true);
    }
    function testchars(tmpfield,idx)
    {
        var tmp1=tmpfield.value;
        if (tmp1.length>240)
        {
            alert("<%response.write FixLang(dictLanguage.Item(Session("language")&"_GiftWrap_9"))%>");
            tmp1=tmp1.substr(0,240);
            tmpfield.value=tmp1;
            document.getElementById("countchar" + idx).innerHTML=240-tmp1.length;
            tmpfield.focus();
        }
        document.getElementById("countchar" + idx).innerHTML=240-tmp1.length;
    }
    jQuery(function($) {
        $pc("#GWASubmit").click(function() {

            var $form = $pc("#Form1");
            var $target = $pc($form.attr('data-target'));
    
            $pc.ajax({
                type: $form.attr('method'),
                url: $form.attr('action'),
                data: $form.serialize(),
    
                success: function(data, status) {
                    parent.recalculate("", "#GWframeloader", 0, '');
                    $target.modal('hide');
                }
            });
    
            return false;
        });
    });
</script>
<form id="Form1" class="pcForms" data-target="#QuickViewDialog" action="opc_GiftWrap.asp?action=add" onsubmit="return Form1_Validator(this)" method="POST">

    <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
        <h3 class="modal-title" id="pcDialogTitle"><%=dictLanguage.Item(Session("language")&"_opc_gwa_title")%></h3>
    </div>
    <div class="modal-body">
        <div class="pcClear"></div>
        <input name="index" type="hidden" value="<%=f%>">

        <div class="pcTable">

            <% IF HaveSecurity=1 THEN %>

                    <div class="pcTableRow">
                        <div class="pcTableRowFull">
                            <div class="pcErrorMessage"><%=dictLanguage.Item(Session("language")&"_opc_gwa_1")%></div>
                        </div>
                    </div>

            <% ELSE %>

                    <% IF UpdateSuccess="1" THEN %>
        
                        <div class="pcTableRow">
                            <div class="pcTableRowFull">
                                <div class="pcSuccessMessage"><%=dictLanguage.Item(Session("language")&"_opc_gwa_2")%></div>
                            </div>
                        </div>
        
                    <% ELSE %>
        
                        <%
                        if pcCartArray(f,10)=0 then
                            pIDProduct=pcCartArray(f,0)
                            pName=pcCartArray(f,1)
                            %>
                            <div class="pcTableRow">
                                <div class="pcTableRowFull">
                                    <h2><%response.write dictLanguage.Item(Session("language")&"_opc_giftWrap_3") & pName %></h2>
                                </div>
                            </div>
                
                            <div class="pcTableRow">
                
                                <div class="pcTableRowFull">
                
                
                                    <div class="pcTable">
                                        <div class="pcTableRow">
                                            <div class="pcTableColumnLeft">
                
                                                <div class="pcTableRow">
                                                    <input type="radio" name="GW<%=f%>" value="" class="clearBorder">
                                                    <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_3")%>
                                                </div>
                
                                                <%
                                                query="SELECT pcGW_IDOpt,pcGW_OptName,pcGW_OptImg,pcGW_OptPrice from pcGWOptions WHERE pcGW_removed=0 AND pcGW_OptActive=1 ORDER BY pcGW_OptOrder ASC,pcGW_OptName ASC;"
                                                set rstemp=connTemp.execute(query)
                                                Count=0
                                                do while not rstemp.eof
                                                    IDOpt=rstemp("pcGW_IDOpt")
                                                    OptName=rstemp("pcGW_OptName")
                                                    OptImg=rstemp("pcGW_OptImg")
                                                    OptPrice=rstemp("pcGW_OptPrice")
                                                    %>
                                                    <div class="pcTableRow">
                                                        <span>
                                                            <input type="radio" name="GW<%=f%>" value="<%=IDOpt%>" <%if (Count=0) or (cdbl(IDOpt)=cdbl(pcCartArray(f,34))) then%>checked<%Count=1%><%end if%> class="clearBorder">
                                                        </span>
                                                        <span>
                                                            <%=OptName%>
                
                                                            <%if cdbl(OptPrice)=0 then%>
                                                                <b><%response.write dictLanguage.Item(Session("language")&"_GiftWrap_6")%></b>
                                                            <%else%>
                                                                &nbsp;-&nbsp;<%=scCurSign & money(OptPrice)%>
                                                            <%end if%>
                                                        </span>
                                                        <span>
                                                            <%if OptImg<>"" then%><img src="<%=pcf_getImagePath("catalog",OptImg)%>" border="0" align="top"><%end if%>
                                                        </span>
                                                    </div>
                                                    <%
                                                    rstemp.MoveNext
                                                loop 
                                                %>
                                            </div>
                                            <div class="pcTableColumnRight">
                
                                                <%= dictLanguage.Item(Session("language")&"_GiftWrap_4")%><br>
                                                <textarea name="GWText<%=f%>" rows="6" cols="35" onkeyup="javascript:testchars(this,'<%=f%>');" maxlength="240"><%=pcCartArray(f,35)%></textarea>
                                                <br>
                                                <%= dictLanguage.Item(Session("language")&"_GiftWrap_5a")%>
                                                <span id="countchar<%=f%>" name="countchar<%=f%>" style="font-weight: bold">240</span>
                                                <%response.write dictLanguage.Item(Session("language")&"_GiftWrap_5b")%><br>
                                            </div>
                                        </div>
                                    </div>
                
                                </div>
                
                            </div>
            
            
                        <% end if %>
        
                    <% END IF %>

            <% END IF %>
            
        </div>
        <div class="pcClear"></div>
    </div>
    <div class="modal-footer">
        <button id="GWASubmit" form="ratting-form" class="btn btn-default" type="submit"><%response.write dictLanguage.Item(Session("language")&"_GiftWrap_7")%></button>
    </div>
</form>
<% 
call closedb()
%>

