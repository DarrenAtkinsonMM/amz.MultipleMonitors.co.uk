<%if enableBulkAdd=1 then %><div id="bulk-panel">
    <div id="bulk-tab">Bulk<br />Add</div>
    <div id="cat-bulk-add-wrapper">
        <form name="categorybulkadd" id="categorybulkadd">
        <div id="cat-bulk-add" >
            <h2>Bulk Add To Cart</h2>
            <p class="ba1">If you know the SKUs enter them below. Values remain until you add to cart or reset the form.</p>
   
            <div style="width:120px; display:inline-block">SKU:</div><div style="display:inline">Qty:</div>
                 <%
                    strAllAdded=session("bulkcategoryadd")
                    arrAllAdded=split(strAllAdded,"||")
                    intAllAddedCount=ubound(arrAllAdded)-1
                    y=0
                    for x=0 to 6
                         if y<=intAllAddedCount then 
                            strCurSkuVal=arrAllAdded(y) 
                            strCurSkuQty=arrAllAdded(y+1) 
                            y=y+1
                            if strCurSkuQty="" then
                                strCurSkuQty=1
                            end if
                        else 
                            strCurSkuVal=""
                            strCurSkuQty="1"
                        end if 
                 %>
            <div class="bulksku"><input type="text" class="form-control sku" name="sku<%=x %>" id="sku<%=x %>" placeholder="SKU:"  value="<%=strCurSkuVal%>" /><input type="text" class="form-control qty" name="sku<%=x %>qty" id="sku<%=x %>qty" value="<%= strCurSkuQty%>" /></div>
                <%
                    y=y+1
                next 
                %>
            <div class="btn bulk">
                <button id="bulkaddtocart" class="btn btn-primary bulk">Add to Cart</button>
                <button id="bulkreset" class="btn btn-reset bulk">Reset</button>
            </div>
        </div>
    </form>
</div>
</div>
<%end if %>
