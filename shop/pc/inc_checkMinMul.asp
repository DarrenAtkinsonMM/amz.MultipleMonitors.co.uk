<%
Sub CheckMinMulQty(tmpID,tmpQty)
Dim queryQ,rsQ
Dim tmpValid,tmpMin,tmpMul,tmpDesc


pcv_intIdProduct = pcf_GetParentId(tmpID)

queryQ="SELECT pcprod_qtyvalidate,pcprod_minimumqty,pcProd_multiQty,Description FROM Products WHERE idProduct=" & pcv_intIdProduct & ";"
set rsQ=connTemp.execute(queryQ)

if not rsQ.eof then
	tmpValid=rsQ("pcprod_qtyvalidate")
	if IsNull(tmpValid) OR tmpValid="" then
		tmpValid=0
	end if
	tmpMin=rsQ("pcprod_minimumqty")
	if IsNull(tmpMin) OR tmpMin="" then
		tmpMin=0
	end if
	tmpMul=rsQ("pcProd_multiQty")
	if IsNull(tmpMul) OR tmpMul="" then
		tmpMul=0
	end if
	tmpDesc=rsQ("Description")
	set rsQ=nothing
	
	if Clng(tmpMin)>0 then
		if Clng(tmpQty)<Clng(tmpMin) then
			call closedb()
			Session("message") = "The quantity of "&tmpDesc&" that you are trying to order is less than the minimum quantity customers can buy. You need to buy at least "&tmpMin&" unit(s)."
			response.redirect "msgb.asp?back=1"
		end if
	end if
	
	if (Clng(tmpMul)>0) AND (Clng(tmpValid)=1) then
		if (Clng(tmpQty) Mod Clng(tmpMul))>0 then
			call closedb()
			Session("message") = "The product "&tmpDesc&" can only be ordered in multiples of "&tmpMul&"."
			response.redirect "msgb.asp?back=1"
		end if
	end if
end if
set rsQ=nothing
End Sub
%>
	
	