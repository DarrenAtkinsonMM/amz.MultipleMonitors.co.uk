<!--#include file="../../common.asp"-->
<!--#include file="../../../pc/pcCheckPricingCats.asp"-->
<%
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' START: Implement Interface
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
pcv_strMethod = Session("SFMethod")

Select Case pcv_strMethod

    Case "BackInStockWidget" : Call pcs_PrdBackInStock("minimal")
    
    Case "BackInStockWidgetModal" : Call pcs_PrdBackInStock("modal")

    Case "BackInStockMenu" : Call pcs_BISMenu()
    
    Case "BackInStockCPanelJS" : Call pcs_BISCPanelJS()
    
    Case "PrdBackInStockWaitList" : Call pcs_AddWaitList()
    
End Select
Session("SFMethod") = ""
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' END: Implement Interface
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
%>