<%
Function calculateRewardsTotal(RewardsActive, pcSFUseRewards, RewardsPercent, pSubTotal)

    iDollarValue = 0
    If RewardsActive=1 And pcSFUseRewards<>"" Then 
        iDollarValue = pcSFUseRewards * (RewardsPercent / 100)
        
        pSubTotal = pSubTotal - iDollarValue        
        If pSubTotal<0 Then
            xVar = (pSubTotal + iDollarValue) / (RewardsPercent / 100)
            pcIntUseRewards = Round(xVar)
            pcSFUseRewards = pcIntUseRewards
            iDollarValue = pcSFUseRewards * (RewardsPercent / 100)
            pSubTotal = 0
        End If
        
    End If
    
    calculateRewardsTotal = iDollarValue

End Function



'// IN
' RewardsActive  (Where does this come from... it seems like a constant from the settings file)
' pcIntRewardPointsAccrued  (This input comes from the db and is only used in the following block)
' pcIntRewardPointsUsed  (This input comes from the db and is only used in the following block)
' RewardsPercent  (This comes from the db and is used in multiple calculations on this page)

'// OUT
' pcIntBalance  (this is an important return value that is saved... BUT it is the same as pcSFIntBalance)
' pcIntDollarValue  (this value is NEVER used and has absolutely no purpose)
' pcSFIntBalance = (this is a duplicate of the value pcIntBalance and it has been removed)

Function calculateRewardsBalance(RewardsActive, pcIntRewardPointsAccrued, pcIntRewardPointsUsed, RewardsPercent)

    pcIntBalance = 0
    If RewardsActive = 1 Then
    
        If IsNull(pcIntRewardPointsAccrued) or pcIntRewardPointsAccrued="" Then 
            pcIntRewardPointsAccrued = 0
        End if
        If IsNull(pcIntRewardPointsUsed) or pcIntRewardPointsUsed="" Then 
            pcIntRewardPointsUsed = 0
        End if
        pcIntBalance = pcIntRewardPointsAccrued - pcIntRewardPointsUsed

    End If
    calculateRewardsBalance = pcIntBalance
    
End Function




' pcIntRewardPointsUsed
' pcUseRewards
' pcSFUseRewards
' pcIntCartRewards (this is a duplicate value and needs removed)
' pcSFCartRewards

Function IsUsingRewards(rtype, pcUseRewards, savUseRewards, pcIntBalance, RewardsIncludeWholesale)

    If rtype="1" Then

        If IsNull(pcUseRewards) Or pcUseRewards="" Then
            pcUseRewards=0
        End If
        
    Else

        pcUseRewards=savUseRewards
        If IsNull(pcUseRewards) Or pcUseRewards="" Then
            pcUseRewards=0
        End If

    End If

    If validNum(pcUseRewards) Then
    
        pcSFCartRewards = 0
        
        If (CLng(pcUseRewards) > pcIntBalance) AND (pcIntBalance>0) Then
            pcUseRewards = pcIntBalance
        End If
        
        If pcIntBalance=0 Then
            pcUseRewards = 0
        End If
        
        pcSFUseRewards = pcUseRewards
        
        If session("customerType")="1" And RewardsIncludeWholesale=0 Then 
            pcSFUseRewards = 0
        End If
            
    Else '// If validNum(pcUseRewards) Then
    
        If rtype<>"1" Then
            If Not validNum(pcSFUseRewards) Then
                pcSFUseRewards = 0
            End If                
        Else
            pcSFUseRewards = 0
        End If

    End If
    
    '// RP ADDON-S
    If session("customerType")="1" And RewardsIncludeWholesale=0 Then 
        pcSFUseRewards=""
    End If
    
    IsUsingRewards = pcSFUseRewards

End Function


Function calculateSFCartRewards(pcSFUseRewards, pcCartArray, ppcCartIndex, RewardsIncludeWholesale)

    pcSFCartRewards = 0
    
    '// This customer will accrue the points since they are not using any for the purchase
    If pcSFUseRewards="" Or pcSFUseRewards="0" Then

        pcSFUseRewards=""

        pcSFCartRewards = Int(calculateCartRewards(pcCartArray, ppcCartIndex))

        '// If customer is wholesale and wholesale is not included in rewards
        If session("customerType")="1" And RewardsIncludeWholesale=0 Then 
            pcSFCartRewards=0
        End If
        
    End If
    
    calculateSFCartRewards = pcSFCartRewards

End Function



'// Cart Points Rewarded Quantity
Function calculateCartRewards(pcCartArray, indexCart)

    Dim f, totalQuantity
	totalQuantity = 0
	for f=1 to indexCart
		if pcCartArray(f,10) = 0 then  
			if (pcCartArray(f,22)="") OR IsNull(pcCartArray(f,22)) then
				pcCartArray(f,22)=0
			end if
			totalQuantity = totalQuantity + (pcCartArray(f,2) * cdbl(pcCartArray(f,22)))
			'BTO Additional Charges Reward Points
			if (pcCartArray(f,29)<>"") and (pcCartArray(f,29)<>"0") then
				totalQuantity = totalQuantity + pcCartArray(f,29)
			end if
		end if
	next   
	calculateCartRewards = totalQuantity  
	set f			= nothing
	set totalQuantity	= nothing

End Function
%>