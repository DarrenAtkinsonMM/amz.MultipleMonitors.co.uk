<div id="bis_modal" class="modal fade" tabindex="-1" role="dialog">
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title">&nbsp;<%'=nmBText %></h4>
      </div>
      <div class="modal-body">
        <div class="pcBackInStockSection">
            <div>
                <input id="nmEmail" name="nmEmail" class="form-control" placeholder="you@domain.com" value="<%=Session("pcSFFromEmail")%>" type="email" />
            </div>
        </div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
        <button onclick="javascript:sendBackInStock()" name="nmButton" class="btn btn-primary" type="button"><%=nmBText %></button>
      </div>
    </div>
  </div>
</div>

<div class="BackInStockWrapper">
    <div class="BackInStockButtonContainer">
        <button type="button" class="pcButton pcBackInStockButton" data-toggle="modal" data-target="#bis_modal">
            <span class="pcButtonText"><%=nmBText %></span>
        </button>
    </div>
</div>