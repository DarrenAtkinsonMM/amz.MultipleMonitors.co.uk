<%
Public Function calculateTotalPromotions(pcPromoIndex, pcPromoSession)

    TotalPromotions=0
    If pcPromoIndex<>"" And pcPromoIndex>"0" Then

        PromoArr1 = pcPromoSession
        PromoIndex = pcPromoIndex

        For m=1 To PromoIndex
            TotalPromotions = TotalPromotions + cdbl(PromoArr1(m,2))
        Next
        
    End If
				
    Session("PromotionTotal") = TotalPromotions
                
    calculateTotalPromotions = TotalPromotions
                
End Function   
%>