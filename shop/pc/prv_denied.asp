
<!--#include file="../includes/common.asp"-->

<%
  msg = request("message")
  rvd = request("rvd")
%>

<div id="prv_postDenied">
  <div class="modal-header">
    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
    <h3 class="modal-title"><%=dictLanguage.Item(Session("language")&"_prv_9")%></h3>
  </div>
  <div class="modal-body">
    <p>&nbsp;</p>
    <% If msg&""<>"" Then %>
      <p><%=pcf_getStoreMsg(msg)%></p>
    <% ElseIf rvd = "1" Then %>
      <p><%=dictLanguage.Item(Session("language")&"_prv_46")%></p>
    <% Else %>
      <p><%=dictLanguage.Item(Session("language")&"_prv_8")%></p>
    <% End If %>
    <p>&nbsp;</p>
  </div>
  <div class="modal-footer">
    <button class="btn btn-default pcButtonCloseWindow" data-dismiss="modal" type="submit">Close Window</button>
  </div>
</div>