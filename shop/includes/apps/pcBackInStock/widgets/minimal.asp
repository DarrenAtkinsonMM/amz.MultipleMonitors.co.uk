<div class="BackInStockWrapper">
    <div class="BackInStockButtonContainer">
        <div class="input-group">
            <input id="nmEmail" name="nmEmail" class="form-control" maxlength="254" placeholder="you@domain.com" value="<%=Session("pcSFFromEmail")%>" type="email" />
            <span class="input-group-btn" style="width:0;">
                <button onclick="javascript:sendBackInStock()" name="nmButton" class="btn btn-default" type="button"><%=nmBText %></button>
            </span>
        </div>
    </div>
</div>