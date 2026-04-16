var BTOConflict=0;
var BTONameConflict="";
var BTOValueConflict="";
var TempUnchecked="";
var ListUnchecked="";
var pcv_Preset=0;
var BTOCMMsgBox_CSSClass="pcAttention";
var BoxHadMsg=0;

//Find Item Location in the Saved Array - MULTIPLE TIMES
function New_FindItemLocationMulti(itemID,tmpIdx)
{
	var tmpIndex=0
	var i=FormItemCount+1;
	var j=Math.round((FormItemCount+1)/2);
	var m=-1;
	do
	{
		i--;
		if (eval(FormItem1[i])==eval(itemID))
		{
			tmpIndex++;
			if (tmpIndex>tmpIdx)
			{
				return(FormItem2[i]);
				break;
			}
		}
		m++;
		if (FormItem1[m]==itemID)
		{
			tmpIndex++;
			if (tmpIndex>tmpIdx)
			{
				return(FormItem2[m]);
				break;
			}
		}
	}
	while (--j);
	return('');
}

function CheckSameGroup(idvalue)
{
	var tmpStr=New_FindItemLocation(idvalue);
	var tmpStr1=tmpStr.split("_");
	var groupid=tmpStr1[3];	
	if (tmpStr1[1]==1)
	{
		return(FormDrop3[groupid]);
	}
	else
	{
		if (tmpStr1[1]==0)
		{
			return(FormRadio4[groupid]);
		}
	}
	return(0);
}

function GetCatName(iditem)
{
var k=0;

	for (k=0;k<=CatCount;k++)
	{
		if (CatID[k]=="" + iditem)
		{
			return(CatName[k]);
		}
	}
	return("");
}

function GetCatNameI(iditem)
{
var iditem1=iditem + "_0";
var k=0;
var tmp2=iditem1.split("_");
var tmp3=tmp2[0];
	for (k=0;k<=CatCount;k++)
	{
		var tmp1=CatPrds[k];
		var pos=tmp1.indexOf(tmp3 +',');
		if (pos==0)
		{
			pos=1;
		}
		else
		{
			pos=tmp1.indexOf(','+ tmp3 +',');
		}
		if (pos>0)
		{
			return(CatName[k]);
		}
	}
	return("");
}

function GetCatID(iditem)
{
var k=0;
var tmp3=iditem;
	for (k=0;k<=CatCount;k++)
	{
		var tmp1=CatPrds[k];
		var pos=tmp1.indexOf(tmp3 +',');
		if (pos==0)
		{
			pos=1;
		}
		else
		{
			pos=tmp1.indexOf(','+ tmp3 +',');
		}
		if (pos>0)
		{
			return(CatID[k]);
		}
	}
	return(0);
}

function DisplayAlert1(catlocation,distype,prdname,olditem,newitem,catname)
{
	var tmpStr="";
	if (pcv_Preset==0)
	{
		if (eval(distype)==1)
		{
			tmpStr="'" + prdname + "'" + pcv_msg_btocm_8a + "'" + olditem + "'" + pcv_msg_btocm_8f + "'" + catname + "'" + pcv_msg_btocm_8g;
		}
		if (eval(distype)==0)
		{
			tmpStr="'" + prdname + "'" + pcv_msg_btocm_8a + "'" + olditem + "'" + pcv_msg_btocm_8b + "'" + newitem + "'" + pcv_msg_btocm_8c + "'" + catname + "'" + pcv_msg_btocm_8d + "'" + prdname + "'" + pcv_msg_btocm_8e;
		}
		if (eval(distype)==2)
		{
			tmpStr="'" + prdname + "'" + pcv_msg_btocm_8a + pcv_msg_btocm_11a + "'" + catname + "'" + pcv_msg_btocm_11;
		}
		if (eval(distype)==3)
		{
			tmpStr=pcv_msg_btocm_4 + "'" + newitem + "'" + pcv_msg_btocm_2 + "'" + prdname + "'" + pcv_msg_btocm_4a + "'" + olditem + "'" + pcv_msg_btocm_4c + "'" + newitem + "'" + pcv_msg_btocm_4d + "'" + catname + "'";
		}
		if (eval(distype)==4)
		{
			tmpStr=pcv_msg_btocm_4 + "'" + newitem + "'" + pcv_msg_btocm_2 + "'" + prdname + "'" + pcv_msg_btocm_4b + "'" + newitem + "'" + pcv_msg_btocm_4d + "'" + catname + "'";
		}
		//Display Configurator Plus Messages
		if (ShowBTOCMMsg==1)
		{
			var tmpSource=document.getElementById("CMMsg" + catlocation).innerHTML;
			if (tmpSource=="")
			{
				document.getElementById("CMMsg" + catlocation).innerHTML="<div class='" + BTOCMMsgBox_CSSClass + "'>"+tmpStr+"</div>";
			}
			else
			{
				tmpSource=tmpSource.replace(/<\/DIV>/gi, "<br>"+tmpStr+"</DIV>");
				document.getElementById("CMMsg" + catlocation).innerHTML=tmpSource;
			}
			if (BoxHadMsg==1)
			{
				var tmpSource=document.getElementById("modal_cmmsgbox").innerHTML;
				tmpSource=tmpSource.replace(/<\/DIV>/gi, "<br>"+tmpStr+"</DIV>");
				document.getElementById("modal_cmmsgbox").innerHTML=tmpSource;
			}
			else
			{
				document.getElementById("modal_cmmsgbox").innerHTML="<div align='center'>"+tmpStr+"</div>";
				BoxHadMsg=1;
			}
			OpenHS();
		}
	}
}

function DisplayAlert2(catlocation,tmpStr)
{
	//Display Configurator Plus Messages
	if (ShowBTOCMMsg==1)
	{
		var tmpSource=document.getElementById("CMMsg" + catlocation).innerHTML;
		if (tmpSource=="")
		{
			document.getElementById("CMMsg" + catlocation).innerHTML="<div class='" + BTOCMMsgBox_CSSClass + "'>"+tmpStr+"</div>";
		}
		else
		{
			tmpSource=tmpSource.replace(/<\/DIV>/gi, "<br>"+tmpStr+"</DIV>");
			document.getElementById("CMMsg" + catlocation).innerHTML=tmpSource;
		}
		if (BoxHadMsg==1)
		{
			var tmpSource=document.getElementById("modal_cmmsgbox").innerHTML;
			tmpSource=tmpSource.replace(/<\/DIV>/gi, "<br>"+tmpStr+"</DIV>");
			document.getElementById("modal_cmmsgbox").innerHTML=tmpSource;
		}
		else
		{
			document.getElementById("modal_cmmsgbox").innerHTML="<div align='center'>"+tmpStr+"</div>";
			BoxHadMsg=1;
		}
		OpenHS();
	}
}

function GetIName(dropmenu,itemvalue)
{
var j=0;
var oSelect = dropmenu;
	for (j=0;j<oSelect.options.length;j++)
	{
		var tmpStr1=oSelect.options[j].value;
		var tmpstr=tmpStr1.split('_');
		
		if (tmpstr[4]==itemvalue)
		{
			var tmpStr2=oSelect.options[j].text;
			var str_array=tmpStr2.split(' - ' + pcv_dicProdOpt1);
			var str1_array=str_array[0].split(' - ' + pcv_dicProdOpt2);
			return(str1_array[0]);
		}
	}
	return('None');
}

function GetItemName(itemvalue)
{
	var provalue=itemvalue + "_0";
	var tmpstr=provalue.split('_');
	var idpro=eval(tmpstr[0]);
	if (idpro!=0)
	{
		var proname=document.getElementsByName("TXT" + idpro).item(0).value;
	}
	else
	{
		var proname='None';
	}
	return(proname);
}

function PreTestConflict(idpro1,preidpro1,prorule1,prorule2,catrule1,catrule2,precat)
{
var i=0;
var j=0;

	for (i=0;i<=RuleCount;i++)
	{
		if ((Rule5[i]==1) && (eval(Rule1[i])!=preidpro1))
		{
		if ((prorule2.indexOf(Rule1[i]+",")==0) || (prorule2.indexOf(","+Rule1[i]+",")>0))
		{
			TempUnchecked=TempUnchecked + Rule1[i]+ ",";
			Rule5[i]=0;
			var tmp1=GetCatID(Rule5[i] + "_0");
			for (j=0;j<=RuleCount;j++)
			{
				if ((Rule5[j]==1) && (eval(Rule1[j])==eval(tmp1)) && (Rule8[j]==1))
				{
					ListUnchecked=ListUnchecked + Rule1[j]+ ",";
					Rule5[j]=0;
					break;
				}
			}
		}
		else
		{
			var tmpcat=CheckCatRules(Rule1[i],catrule2)
			if (tmpcat>-1)
			{
				TempUnchecked=TempUnchecked + Rule1[i]+ ",";
				Rule5[i]=0;
				var tmp1=GetCatID(Rule5[i] + "_0");
				for (j=0;j<=RuleCount;j++)
				{
					if ((Rule5[j]==1) && (eval(Rule1[j])==eval(tmp1)) && (Rule8[j]==1))
					{
						ListUnchecked=ListUnchecked + Rule1[j]+ ",";
						Rule5[j]=0;
						break;
					}
				}
			}
		}
		} //End of Rule5[i]=1
	}
	
	var tmp1=prorule1.split(",");
	
	for (i=0;i<tmp1.length;i++)
	{
		for (j=0;j<=RuleCount;j++)
		{
			if ((eval(Rule1[j])==eval(tmp1[i])) && (Rule5[j]==0) && (Rule8[j]==0))
			{
				var tmp2=CheckSameGroup(tmp1[i]);
				PreTestConflict(tmp1[i],tmp2,Rule2[j],Rule3[j],Rule6[j],Rule7[j]);
				break;
			}
			
		}
	}
	
	//isCAT
	
	var tmp1=catrule1.split(",");
	
	for (i=0;i<tmp1.length;i++)
	{
		for (j=0;j<=RuleCount;j++)
		{
			if (Rule6[j].indexOf(precat+",")>=0)
			{
				break;
			}
			else
			{
				if ((eval(Rule1[j])==eval(tmp1[i])) && (Rule5[j]==0) && (Rule8[j]==1))
				{
					PreTestConflict(0,0,Rule2[j],Rule3[j],Rule6[j],Rule7[j]);
					break;
				}
			}
		}
	}

}

function CheckCatRules(iditem,RuleList)
{
var i=0;
var k=0;
var pos=0;
var tmpstr=RuleList;
var tmpstr1=tmpstr.split(",");
	for (i=0; i<tmpstr1.length;i++)
	{
	if (tmpstr1[i]!="")
	{
		for (k=0;k<=CatCount;k++)
		{
			if (CatID[k]==tmpstr1[i])
			{
				var tmp1=CatPrds[k];
				pos=tmp1.indexOf(iditem +',');
				if (pos==0)
				{
					pos=1;
				}
				else
				{
					pos=tmp1.indexOf(','+ iditem +',');
				}
				if (pos>0)
				{
				return(k);
				}
				break;
			}
		}
	}
	}
	return(-1);
}

function special_CheckCatRules(iditem,parentCat,RuleList)
{
var i=0;
var k=0;
var pos=0;
var tmpstr=RuleList;
var tmpstr1=tmpstr.split(",");
	for (i=0; i<tmpstr1.length;i++)
	{
	if (tmpstr1[i]!="")
	{
		for (k=0;k<=CatCount;k++)
		{
			if ((CatID[k]==tmpstr1[i]) && (parentCat==CatID[k]))
			{
				var tmp1=CatPrds[k];
				pos=tmp1.indexOf(iditem +',');
				if (pos==0)
				{
					pos=1;
				}
				else
				{
					pos=tmp1.indexOf(','+ iditem +',');
				}
				if (pos>0)
				{
				return(k);
				}
				break;
			}
		}
	}
	}
	return(-1);
}

function TestConflict(idpro,preidpro,prorule1,prorule2,catrule1,catrule2,precat)
{
var i=0;
var j=0;

	for (i=0;i<=RuleCount;i++)
	{
		if ((Rule5[i]==1) && ((eval(Rule1[i])!=preidpro) || (preidpro==0)))
		{
			var tmprule1=Rule2[i];
			var tmprule2=Rule3[i];
			var tmprule3=Rule6[i];
			var tmprule4=Rule7[i];
						
			var tmpArr1=new Array();
			tmpArr1=tmprule1.split(",");
			var tmpArr2=new Array();
			tmpArr2=tmprule2.split(",");
			var haveconflict=0;
			for (j=0; j<tmpArr2.length; j++)
			{
			if (tmpArr2[j]!="")
			{
				if ((prorule1.indexOf(tmpArr2[j]+",")==0) || (prorule1.indexOf(","+tmpArr2[j]+",")>0))
				{
					haveconflict=1;
					break;
				}
			}
			}
			if (haveconflict==0)
			{
			for (j=0; j<tmpArr1.length; j++)
			{
			if (tmpArr1[j]!="")
			{
				if ((prorule2.indexOf(tmpArr1[j]+",")==0) || (prorule2.indexOf(","+tmpArr1[j]+",")>0))
				{
					haveconflict=1;
					break;
				}
				else
				{
					var tmpcat=CheckCatRules(tmpArr1[j],catrule2)
					if (tmpcat>-1)
					{
						haveconflict=1;
						break;
					}
				}

			}
			}
			}
			
			if (haveconflict==0)
			{
			var tmpArr4=new Array();
			tmpArr4=prorule1.split(",");
			for (j=0; j<tmpArr4.length; j++)
			{
			if (tmpArr4[j]!="")
			{
				var tmpcat=CheckCatRules(tmpArr4[j],tmprule4)
				if (tmpcat>-1)
				{
					haveconflict=1;
					break;
				}
			}
			}
			}
			
			
			if (haveconflict==1)
			{
				BTOConflict=BTOConflict+1;
				BTONameConflict=Rule4[i];
				BTOValueConflict=Rule1[i];
				return(false);
			}
			
		} //End of Rule5[i]=1
	}
		
	var tmp1=prorule1.split(",");

	for (i=0;i<tmp1.length;i++)
	{
		for (j=0;j<=RuleCount;j++)
		{
			if ((eval(Rule1[j])==eval(tmp1[i])) && (Rule5[j]==0)  && (Rule8[j]==0))
			{
				TestConflict(tmp1[i],0,Rule2[j],Rule3[j],Rule6[j],Rule7[j]);
				break;
			}
			
		}
	}
	
	//isCAT
	
	var tmp1=catrule1.split(",");

	for (i=0;i<tmp1.length;i++)
	{
		for (j=0;j<=RuleCount;j++)
		{
			if ((eval(Rule1[j])==eval(tmp1[i])) && (Rule5[j]==0)  && (Rule8[j]==1))
			{
				if (Rule6[j].indexOf(precat+",")>=0)
				{
					break;
				}
				else
				{
					TestConflict(0,0,Rule2[j],Rule3[j],Rule6[j],Rule7[j],Rule1[j]);
					break;
				}
			}
			
		}
	}

}

function TestItem(xfield,fieldtype,ctype)
{
	var new_provalue=xfield.value;
	var tmpstr=new_provalue.split('_');
	var tmp1=xfield.name;
	var tmp2=tmp1.split("CAG")
	if (parseInt(fieldtype)!=3)
	{
		var parentcat=tmp2[1];
	}
	else
	{
		var tmp3=tmp2[1];
		var parentcat=tmp3.substr(0,tmp3.length-tmpstr[4].length);
	}
	var new_idpro=eval(tmpstr[4]);
	if (new_idpro!=0)
	{
		if (fieldtype==1)
		{
				var oSelect = eval("document.additem." + xfield.name);
				var tempStr1=oSelect.options[oSelect.selectedIndex].text;
			  	var str_array=tempStr1.split(' - ' + pcv_dicProdOpt1);
				var str1_array=str_array[0].split(' - ' + pcv_dicProdOpt2);
				var new_proname=str1_array[0];
		}
		else
		{
			var new_proname=eval("document.additem.TXT" + new_idpro + ".value");
		}
	}
	else
	{
		var new_proname='';
	}

	var i=0;
	for (i=0;i<=RuleCount;i++)
	{
		var tmpstr=Rule3[i];
		var pos=tmpstr.indexOf(new_idpro +',');
		if (pos==0)
		{
			pos=1;
		}
		else
		{
			pos=tmpstr.indexOf(','+new_idpro +',');
		}
		if ((Rule7[i]!="") && (Rule5[i]==1))
		{
			var tmpcat=special_CheckCatRules(new_idpro,parentcat,Rule7[i]);
		}
		if (((pos>0) || (tmpcat>-1)) && (Rule5[i]==1))
		{
			if (pos>0)
			{
				if (Rule8[i]==0)
				{
					if (ctype==0) DisplayAlert2(GetCatID(new_idpro),pcv_msg_btocm_1 + "'" + new_proname + "'" + pcv_msg_btocm_9 + "'" + GetCatNameI(new_idpro + "_0")+"'" + pcv_msg_btocm_2 + "'" + Rule4[i]  + "'" + pcv_msg_btocm_9 + "'" + GetCatNameI(Rule1[i] + "_0")+"'.");
				}
				else
				{
					if (ctype==0) DisplayAlert2(GetCatID(new_idpro),pcv_msg_btocm_1 + "'" + new_proname + "'" + pcv_msg_btocm_9 + "'" + GetCatNameI(new_idpro + "_0")+"'" + pcv_msg_btocm_2 + pcv_msg_btocm_10 + "'" + Rule4[i] +"'.");
				}
			}
			else
			{
				if (Rule8[i]==0)
				{
					if (ctype==0) DisplayAlert2(CatID[tmpcat],pcv_msg_btocm_3 + "'" + CatName[tmpcat] + "'" + pcv_msg_btocm_2 + "'" + Rule4[i]  + "'.");
				}
				else
				{
					if (ctype==0) DisplayAlert2(CatID[tmpcat],pcv_msg_btocm_3 + "'" + CatName[tmpcat] + "'" + pcv_msg_btocm_2 + pcv_msg_btocm_10 + "'" + Rule4[i]  + "'.");
				}
			}
			return(false);
		}
	}
	
	prohaverules=0;
	prorule1="";
	prorule2="";
	catrule1="";
	catrule2="";
	prdname1="";
	
	for (i=0;i<=RuleCount;i++)
	{
		if (eval(Rule1[i])==new_idpro)
		{
			prdname1=Rule4[i];
			prorule1=Rule2[i];
			prorule2=Rule3[i];
			catrule1=Rule6[i];
			catrule2=Rule7[i];
			prohaverules=1;
			break;
		}
	}
	
	BTOConflict=0;
	BTONameConflict="";
	TempUnchecked="";
	ListUnchecked="";
	var idpro=CheckSameGroup(new_idpro);
	if (prohaverules==1)
	{
		PreTestConflict(new_idpro,idpro,prorule1,prorule2,catrule1,catrule2);
		TestConflict(new_idpro,idpro,prorule1,prorule2,catrule1,catrule2);
	}
	
	//isCat
	if (BTOConflict==0)
	{
		//var parentcat=GetCatID(new_idpro)
		if (eval(parentcat)!=0)
		{
			var t1=0;
			for (t1=0;t1<=RuleCount;t1++)
			{
				if ((eval(Rule1[t1])==eval(parentcat)) && (Rule8[t1]==1))
				{
					TestConflict(0,0,Rule2[t1],Rule3[t1],Rule6[t1],Rule7[t1]);
				}
			}
		}
	
	}
	
	if (BTOConflict!=0)
	{
		if (TempUnchecked!="")
		{
			var tmpC=TempUnchecked.split(",");
			for (k=0; k<tmpC.length;k++)
			{
			for (j=0;j<=RuleCount;j++)
			{
				if (tmpC[k]!="")
				{
				if ((eval(tmpC[k])==eval(Rule1[j])) && (Rule8[j]==0))
				{
					Rule5[j]=1;
					break;
				}
				}
			}
			}
			TempUnchecked="";
		}
		if (ListUnchecked!="")
		{
			var tmpC=ListUnchecked.split(",");
			for (k=0; k<tmpC.length;k++)
			{
			for (j=0;j<=RuleCount;j++)
			{
				if (tmpC[k]!="")
				{
				if ((eval(tmpC[k])==eval(Rule1[j])) && (Rule8[j]==1))
				{
					Rule5[j]=1;
					break;
				}
				}
			}
			}
			ListUnchecked="";
		}
		if (ctype==0) DisplayAlert2(GetCatID(BTOValueConflict),pcv_msg_btocm_5 + "'" + BTONameConflict + "'" + pcv_msg_btocm_8f + "'" + GetCatNameI(BTOValueConflict + "_0") +"'" + pcv_msg_btocm_5a + "'" + new_proname + "'" + pcv_msg_btocm_9 + "'" + GetCatNameI(new_idpro+"_0") +"'" + pcv_msg_btocm_5b);
		return(false);
	}
	return(true);
}

//Unlock Drop-down box Function
function New_UnlockDropDown(tmpindex)
{
	var k=0;
	var m=0;
	var objElems = document.additem.elements;
	objElems[tmpindex].disabled=false;
}

//Lock Drop-down box Function
function New_LockDropDown(tmpindex,ctype)
{
	document.additem.elements[tmpindex].disabled=true;
	try
	{
		eval("test"+document.additem.elements[tmpindex].name+"()");
	}catch(e){}

	if (ctype==1)
	{
		var j=document.additem.elements[tmpindex].length;
		var i=0;
		do
		{
			i=j-1;
			var tmpStr=document.additem.elements[tmpindex].options[i].value;
			var tmpStr1=tmpStr.split("_");
			if (eval(tmpStr1[4])!=0)
			{
				New_ReverseItem(tmpStr1[4]);
			}
		}
		while (--j);	
	}	
}

//Lock Drop-down item Function
function New_LockDropDownItem(tmpindex,idopt)
{
	document.additem.elements[tmpindex].options[idopt].style.color="gray";
}

//Unlock Drop-down item Function
function New_UnLockDropDownItem(tmpindex,idopt)
{
	document.additem.elements[tmpindex].options[idopt].style.color="";
}

//Lock Radio items List Function
function New_LockRadioList(tmpindex,ctype)
{
	var m=FormRadio2[tmpindex];
	var k=FormRadio3[tmpindex];
	var tmp1=0;
	var OptC=0;
	var ChkO=-1;
	var RName="";
	for (tmp1=k; tmp1>=m;tmp1--)
	{
		if (document.additem.elements[tmp1].type=="radio")
		{
			OptC++;
			RName=document.additem.elements[tmp1].name;
			document.additem.elements[tmp1].disabled=true;
			var tmpStr=document.additem.elements[tmp1].value;
			var tmpStr1=tmpStr.split("_");
			if (eval(tmpStr1[4])!=0)
			{
				if (document.additem.elements[tmp1].checked==true)
				{
					ChkO=OptC;
					try
					{
					document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='';
					}catch(e){}
				}
				else
				{
					try
					{
					document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='none';
					}catch(e){}
				}
			}
			if (ctype==1)
			{
				var tmpStr=document.additem.elements[tmp1].value;
				var tmpStr1=tmpStr.split("_");
				if (eval(tmpStr1[4])!=0)
				{
					New_ReverseItem(tmpStr1[4]);
				}
			}
		}
	}
	var OptC=0;
	var tQtyF=-1;
	for (tmp1=k; tmp1>=m;tmp1--)
	{
		var tname=document.additem.elements[tmp1].name;
		if ((document.additem.elements[tmp1].type=="text") && (tname.indexOf(RName + "QF")>=0))
		{
			tQtyF=tmp1;
		}
		if (document.additem.elements[tmp1].type=="radio")
		{
			OptC++;
			if (tQtyF!=-1)
				{
				if (OptC==ChkO)
				{
					try {
					document.additem.elements[tQtyF].style.display='';
					} catch(e){}
				}
				else
				{
					try {
					document.additem.elements[tQtyF].style.display='none';
					} catch(e){}
				}
			}
		}
	}
}

//Unlock Radio items List Function
function New_UnLockRadioList(tmpindex)
{
	var m=FormRadio2[tmpindex];
	var k=FormRadio3[tmpindex];
	var tmp1=0;
	var RName="";
	for (tmp1=k; tmp1>=m;tmp1--)
	{
		if (document.additem.elements[tmp1].type=="radio")
		{
			RName=document.additem.elements[tmp1].name;
			document.additem.elements[tmp1].disabled=false;
			var tmpStr=document.additem.elements[tmp1].value;
			var tmpStr1=tmpStr.split("_");
			if (eval(tmpStr1[4])!=0)
			{
				try
				{
				document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='';
				}catch(e){}
			}
		}
		if (RName!="")
		{
			var tname=document.additem.elements[tmp1].name;
			if ((document.additem.elements[tmp1].type=="text") && (tname.indexOf(RName + "QF")>=0))
			{
				try
				{
				document.additem.elements[tmp1].style.display='';
				} catch(e){}
			}
		}
	}
}

//Lock a Radio or Checkbox Item Function
function New_LockItem(tmpindex)
{
	document.additem.elements[tmpindex].disabled=true;
	var tmpStr=document.additem.elements[tmpindex].value;
	var tmpStr1=tmpStr.split("_");
	if (eval(tmpStr1[4])!=0)
	{
		if (document.additem.elements[tmpindex].checked==true)
		{
			try
			{
			document.getElementById("show_"+document.additem.elements[tmpindex].name+"P"+tmpStr1[4]).style.display='';
			}catch(e){}
			try {
			document.additem.elements[eval(tmpindex)+1].style.display='';
			} catch(e){}
		}
		else
		{
			try
			{
			document.getElementById("show_"+document.additem.elements[tmpindex].name+"P"+tmpStr1[4]).style.display='none';
			}catch(e){}
			try {
			document.additem.elements[eval(tmpindex)+1].style.display='none';
			} catch(e){}
		}
	}
}

//Unlock a Radio or Checkbox Item Function
function New_UnLockItem(tmpindex)
{
	document.additem.elements[tmpindex].disabled=false;
	var tmpStr=document.additem.elements[tmpindex].value;
	var tmpStr1=tmpStr.split("_");
	if (eval(tmpStr1[4])!=0)
	{
			try
			{
			document.getElementById("show_"+document.additem.elements[tmpindex].name+"P"+tmpStr1[4]).style.display='';
			}catch(e){}
			try {
			document.additem.elements[eval(tmpindex)+1].style.display='';
			} catch(e){}
	}
}

//Lock Checkbox items List
function New_LockCBList(tmpindex,ltype)
{
	var haveselected=0;
	var m=FormCB2[tmpindex];
	var k=FormCB3[tmpindex];
	var tmp1=0;
	for (tmp1=k; tmp1>=m;tmp1--)
	{
		if (document.additem.elements[tmp1].type=="checkbox")
		{
			document.additem.elements[tmp1].disabled=true;
			var tmpStr=document.additem.elements[tmp1].value;
			var tmpStr1=tmpStr.split("_");
			if (eval(tmpStr1[4])!=0)
			{
				if ((document.additem.elements[tmp1].checked==true) && (ltype!=1))
				{
					try
					{
					document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='';
					}catch(e){}
					try {
					document.additem.elements[eval(tmp1)+1].style.display='';
					} catch(e){}
				}
				else
				{
					try
					{
					document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='none';
					}catch(e){}
					try {
					document.additem.elements[eval(tmp1)+1].style.display='none';
					} catch(e){}
				}
			}
			if (ltype==1)
			{
				if (document.additem.elements[tmp1].checked==true) haveselected=1;
				document.additem.elements[tmp1].checked=false;
				var tmpStr=document.additem.elements[tmp1].value;
				var tmpStr1=tmpStr.split("_");
				if (eval(tmpStr1[4])!=0)
				{
					New_ReverseItem(tmpStr1[4]);
				}
			}
			calculate(document.additem.elements[tmp1],1);
		}
	}
	return(haveselected);
}

//Unlock Checkbox items List
function New_UnLockCBList(tmpindex)
{
	var m=FormCB2[tmpindex];
	var k=FormCB3[tmpindex];
	var tmp1=0;
	for (tmp1=k; tmp1>=m;tmp1--)
	{
		if (document.additem.elements[tmp1].type=="checkbox") document.additem.elements[tmp1].disabled=false;
		var tmpStr=document.additem.elements[tmp1].value;
		var tmpStr1=tmpStr.split("_");
		if (eval(tmpStr1[4])!=0)
		{
				try
				{
				document.getElementById("show_"+document.additem.elements[tmp1].name+"P"+tmpStr1[4]).style.display='';
				}catch(e){}
				try {
				document.additem.elements[eval(tmp1)+1].style.display='';
				} catch(e){}
		}
	}
}

//Find Option Group Location in the Saved Array
function New_FindGroupLocation(itemID)
{
	var i=0;
	var j=0;
	if (FormDropCount!=-1)
	{
		var j=FormDropCount+1;
		do
		{
			i=j-1;
			if (FormDrop2[i]=="CAG"+itemID)
			{
				return("" + i + "_1_" + FormDrop1[i]);
				break; 
			}
		}
		while (--j);
	}
	
	if (FormRadioCount!=-1)
	{
		var j=FormRadioCount+1;
		do
		{
			i=j-1;
			if (FormRadio1[i]=="CAG"+itemID)
			{
				return("" + i + "_0_0");
				break; 
			}
		}
		while (--j);
	}
	
	if (FormCBCount!=-1)
	{
		var j=FormCBCount+1;
		do
		{
			i=j-1;
			if (FormCB1[i]=="CAG"+itemID)
			{
				return("" + i + "_2_0");
				break; 
			}
		}
		while (--j);
	}
	return("notfound");
}

//Check value of Drop-down box
function New_TestDropDownValue(Dvalue)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		var tmpstr=Rule3[i];
		var pos=tmpstr.indexOf(Dvalue +',');
		if (pos==0)
		{
			pos=1;
		}
		else
		{
			pos=tmpstr.indexOf(','+ Dvalue +',');
		}
		if ((pos>0) && (Rule5[i]==1))
		{
			return(false);
			break;
		}
	}
	while (--j);		

	return(true);
}

//Select able value for Drop-down box
function New_SelectDropDownValue(tmpindex,ctype)
{
	var i=0;
	var j=0;
	var m=0;
	var oSelect = document.additem.elements[tmpindex];
	m=oSelect.options.length;
	
	if (ctype==0)
	{
		//Check default value of this drop-down
		var tmpStr1=oSelect.value;
		var tmpStr2=tmpStr1.split("_");
		tmpid=tmpStr2[4];
		if (parseInt(tmpid)!=0)
		{
			var testresult=New_TestDropDownValue(tmpid);
			if (testresult==true)
			{
				return(tmpStr1);
			}
		}
	}
	
	j=m;
	do
	{
		i=j-1;
		var tmpStr1=oSelect.options[i].value;
		if (tmpStr1.indexOf("0_")==0)
		{
			return(oSelect.options[i].value);
			break;
		}
	}
	while (--j);
	
	var tpos=-2;
	var tval=100000000;
	var tstrval="";
	
	j=m;
	do
	{
		i=j-1;
		var tmpStr1=oSelect.options[i].value;
		var tmpStr2=tmpStr1.split("_");
		tmpid=tmpStr2[4];
		var tmpvalue=tmpStr2[1];
		var testresult=New_TestDropDownValue(tmpid);
		if (testresult==true)
		{
			if (parseFloatEx(tmpvalue)<parseFloatEx(tval))
			{
				tpos=i;
				tval=tmpvalue;
				tstrval=oSelect.options[i].value;
			}
		}
	}
	while (--j);
	if (tpos>-2)
	{
		return(tstrval);	
	}

	return("0_0.00_0_0_0");
}

//Check value of Radio List
function New_TestRadioValue(Dvalue)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		var tmpstr=Rule3[i];
		var pos=tmpstr.indexOf(Dvalue +',');
		if (pos==0)
		{
			pos=1;
		}
		else
		{
			pos=tmpstr.indexOf(','+ Dvalue +',');
		}
		if ((pos>0) && (Rule5[i]==1))
		{
			return(false);
		}
	}
	while (--j);
	return(true);
}

//Select able value for Radio List
function New_SelectRadioValue(tmpindex,ctype)
{
	var k=FormRadio2[tmpindex];
	var m=FormRadio3[tmpindex];
	var j=0;
	var objElems = document.additem.elements;
	
	if (ctype==0)
	{
		//Check default value of this drop-down
		for(j=m;j>=k;j--)
		{
			if (objElems[j].type=="radio")
			{
				if (objElems[j].checked==true)
				{
					var tmpStr1=objElems[j].value;
					var tmpStr2=tmpStr1.split("_");
					tmpid=tmpStr2[4];
					var testresult=New_TestRadioValue(tmpid);
					if (testresult==true)
					{
						objElems[j].checked=true;
						return(objElems[j].value);
						break;
					}
				}
			}
		}
	}
	
	var tpos=-2;
	var tval=100000000;
	var tstrval="";
	
	for(j=m;j>=k;j--)
	{
		if (objElems[j].type=="radio")
		{
			var tmpStr1=objElems[j].value;
			var tmpStr2=tmpStr1.split("_");
			tmpid=tmpStr2[4];
			var tmpvalue=tmpStr2[1];
			if (eval(tmpid)==0)
			{
				objElems[j].checked=true;
				return(objElems[j].value);
				break;
			}
			else
			{
				var testresult=New_TestRadioValue(tmpid);
				if (testresult==true)
				{
					if (parseFloatEx(tmpvalue)<parseFloatEx(tval))
					{
						tpos=j;
						tval=tmpvalue;
						tstrval=objElems[j].value;
					}
				}
			}
		}
	}
	if (tpos>-2)
	{
		objElems[tpos].checked=true;
		return(tstrval);
	}
	return("0_0_0_0_0");
}

//Check able value of Checkbox List
function New_CheckCBHaveItem(tmpindex)
{
	var k=FormCB2[tmpindex];
	var m=FormCB3[tmpindex];
	var j=0;
	var objElems = document.additem.elements;
	
	for(j=m;j>=k;j--)
	{
		if (objElems[j].type=="checkbox")
		{
			if (objElems[j].checked==true)
			{
				return(true);
				break;
			}
		}
	}
	return(false);
}

//Uncheck CAT if no items were selected
function New_UnselectCat(tmpindex)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		if ((eval(Rule1[i])==eval(tmpindex)) && (Rule8[i]==1))
		{
			if ((Rule5[i]==1) || (pcv_Preset==1))
			{
				Rule5[i]=0;
				New_RevFromValues(Rule4[i],Rule2[i],Rule3[i],Rule6[i],Rule7[i]);
			}
			break;
		}
	}
	while (--j);
}

//Check CAT if any items were selected
function New_SelectCat(tmpindex)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		if ((eval(Rule1[i])==eval(tmpindex)) && (Rule8[i]==1))
		{
			if ((Rule5[i]==0) || (pcv_Preset==1))
			{
				Rule5[i]=1;
				New_SetFromValues(Rule4[i],Rule2[i],Rule3[i],Rule6[i],Rule7[i]);
			}
			break;
		}
	}
	while (--j);
}

//Set Field status Function
function New_SetFromValues(prdname,prorule1,prorule2,catrule1,catrule2)
{
var i=0;
var j=0;
var k=0;
var m=0;
var tmpindex=0;

	//Disable "CAN NOT" items
	var tmp1=prorule2.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			//Reverse Items: Rule5=0
			New_ReverseItem(tmp1[i]);
			tmpindex=0;
			do
			{
			var tmpLoc=New_FindItemLocationMulti(tmp1[i],tmpindex);
			tmpindex++;
			
			if (tmpLoc!='')
			{
			var tmpStr1=tmpLoc.split("_");
			
			//Drop-down box
			if ((eval(tmpStr1[0])!=0) && (eval(tmpStr1[1])==1))
			{
				New_LockDropDownItem(tmpStr1[0],tmpStr1[2]);
				var oldvalue=tmp1[i];
				var olditem=GetIName(document.additem.elements[tmpStr1[0]],tmp1[i]);
				var haveselected=0;
				var tmpItem=document.additem.elements[tmpStr1[0]].value;
				var tmpItem1=tmpItem.split("_");
				if (tmpItem1[0]==tmp1[i]) haveselected=1;
				if (haveselected==0)
				{
					var tmpoldvalue1=document.additem.elements[tmpStr1[0]].value;
					var tmpoldvalue=tmpoldvalue1.split("_");
					var oldvalue=tmpoldvalue[4];
					var olditem=GetIName(document.additem.elements[tmpStr1[0]],oldvalue);
				}
				document.additem.elements[tmpStr1[0]].value=New_SelectDropDownValue(tmpStr1[0],0);
				try
				{
				eval("test"+document.additem.elements[tmpStr1[0]].name+"()");
				}catch(e){}
				var tmpStr=document.additem.elements[tmpStr1[0]].value;
				if (haveselected==0)
				{
					var tmpnewvalue=tmpStr.split("_");
					if ((parseInt(tmpnewvalue[4])!=parseInt(oldvalue)) && (parseInt(oldvalue)!=0)) haveselected=1;
				}
				if (haveselected==1)
				{
					var tmpItem=tmpStr.split("_");
					var newitem=GetIName(document.additem.elements[tmpStr1[0]],tmpItem[4]);
					DisplayAlert1(GetCatID(oldvalue),0,prdname,olditem,newitem,GetCatNameI(oldvalue));
				}
				//CAT Rules
				if (tmpStr.indexOf("0_")==0)
				{
					var tmpcat=FormDrop2[tmpStr1[3]].replace("CAG","");
					New_UnselectCat(tmpcat);
				}
				//Save new value to Array
				var tmpStr2=tmpStr.split("_");
				FormDrop3[tmpStr1[3]]=tmpStr2[4];
				calculate(document.additem.elements[tmpStr1[0]],1);
			}
			else
			{
				//Radio List
				if ((eval(tmpStr1[0])==0) && (eval(tmpStr1[1])==0))
				{
					var oldvalue=tmp1[i];
					var olditem=GetItemName(tmp1[i] + "_0");
					var haveselected=0;
					if (document.additem.elements[tmpStr1[2]].checked==true) haveselected=1;
					New_LockItem(tmpStr1[2]);
					if (haveselected==0)
					{
						var oldvalue=FormRadio4[tmpStr1[3]];
						var olditem=GetItemName(oldvalue + "_0");
					}
					var tmpStr=New_SelectRadioValue(tmpStr1[3],0);
					if (haveselected==0)
					{
						var tmpnewvalue=tmpStr.split("_");
						if ((parseInt(tmpnewvalue[4])!=parseInt(oldvalue)) && (parseInt(oldvalue)!=0)) haveselected=1;
					}
					if (haveselected==1)
					{
						var tmpnewvalue=tmpStr.split("_");
						var newitem=GetItemName(tmpnewvalue[4] + "_0");
						DisplayAlert1(GetCatID(oldvalue),0,prdname,olditem,newitem,GetCatNameI(oldvalue));
					}
					//CAT Rules
					if (tmpStr.indexOf("0_")==0)
					{
						var tmpcat=FormRadio1[tmpStr1[3]].replace("CAG","");
						New_UnselectCat(tmpcat);
					}
					//Save new value to Array
					var tmpStr2=tmpStr.split("_");
					FormRadio4[tmpStr1[3]]=tmpStr2[4];
					calculate(eval("document.additem." + FormRadio1[tmpStr1[3]]),3);
				}
				else
				//Checkbox List
				{
					var haveselected=0;
					if (document.additem.elements[tmpStr1[2]].checked==true) haveselected=1;
					document.additem.elements[tmpStr1[2]].checked=false;
					New_LockItem(tmpStr1[2]);
					if (haveselected==1)
					{
						var tmpItem=document.additem.elements[tmpStr1[2]].value;
						var tmpItem1=tmpItem.split("_");
						DisplayAlert1(GetCatID(tmpItem1[4]),1,prdname,GetItemName(tmpItem1[4] + "_0"),"",GetCatNameI(tmpItem1[4]));
					}
					//CAT Rules
					var tmpvalue=New_CheckCBHaveItem(tmpStr1[3]);
					if (tmpvalue==false)
					{
						var tmpcat=FormCB1[tmpStr1[3]].replace("CAG","");
						New_UnselectCat(tmpcat);
					}
					calculate(document.additem.elements[tmpStr1[2]],1);
				}
				
			} //It isnt Drop-down item
			} //tmpLoc<>''
			}
			while (tmpLoc!='');
		} //Have item
	}
	while (--j);
	
	//Disable "CAN NOT" CATs
	var tmp1=catrule2.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			New_UnselectCat(tmp1[i]);
			var haveselected=0;
						
			var tmpStr1=New_FindGroupLocation(tmp1[i]);
			if (tmpStr1!="notfound")
			{
			var tmpStr2=tmpStr1.split("_");
			//Drop-down box
			if (eval(tmpStr2[1])==1)
			{
				var tmpItem=document.additem.elements[tmpStr2[2]].value;
				var tmpItem1=tmpItem.split("_");
				if (tmpItem1[4]!=0) haveselected=1;
				document.additem.elements[tmpStr2[2]].value=New_SelectDropDownValue(tmpStr2[2],1);
				try
				{
				eval("test"+document.additem.elements[tmpStr2[2]].name+"()");
				}catch(e){}
				FormDrop3[tmpStr2[0]]=0;
				New_LockDropDown(tmpStr2[2],1);
				calculate(document.additem.elements[tmpStr2[2]],1);
			}
			else
			{
				//Radio List
				if (eval(tmpStr2[1])==0)
				{
					var tmpStr=New_SelectRadioValue(tmpStr2[0],1);
					if (FormRadio4[tmpStr2[0]]!=0) haveselected=1;
					FormRadio4[tmpStr2[0]]=0;
					New_LockRadioList(tmpStr2[0],1);
					calculate(eval("document.additem." + FormRadio1[tmpStr2[0]]),3);
				}
				else
				//Checkbox List
				{
					haveselected=New_LockCBList(tmpStr2[0],1);
					//Already re-calculate in the New_LockCBList function
				}
			
			} //It isnt Drop-down box
			if (haveselected==1) DisplayAlert1(tmp1[i],2,prdname,"","",GetCatName(tmp1[i]));
			}//notfound
		} //Have Cat
	}
	while (--j);
	
	//Lock "MUST" items
	var tmp1=prorule1.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			tmpindex=0;
			do
			{
			var tmpLoc=New_FindItemLocationMulti(tmp1[i],tmpindex);
			tmpindex++;
			
			if (tmpLoc!='')
			{
			var tmpStr1=tmpLoc.split("_");
			
			//Drop-down box
			if ((eval(tmpStr1[0])!=0) && (eval(tmpStr1[1])==1))
			{
				var tmpItem=document.additem.elements[tmpStr1[0]].value;
				var tmpItem1=tmpItem.split("_");
				var olditem=GetIName(document.additem.elements[tmpStr1[0]],tmpItem1[0]);
				
				document.additem.elements[tmpStr1[0]].options[tmpStr1[2]].selected=true;
				var tmpStr=document.additem.elements[tmpStr1[0]].options[tmpStr1[2]].value;
				var tmpItem1=tmpStr.split("_");
				var newitem=GetIName(document.additem.elements[tmpStr1[0]],tmpItem1[0]);
				
				New_LockDropDown(tmpStr1[0],0);
				//CAT Rules
				var tmpcat=FormDrop2[tmpStr1[3]].replace("CAG","");
				New_SelectCat(tmpcat);
				//Save new value to Array
				var tmpStr2=tmpStr.split("_");
				FormDrop3[tmpStr1[3]]=tmpStr2[4];
				calculate(document.additem.elements[tmpStr1[0]],1);
				if (olditem!=newitem) DisplayAlert1(tmpcat,3,prdname,olditem,newitem,GetCatName(tmpcat));
			}
			else
			{
				//Radio List
				if ((eval(tmpStr1[0])==0) && (eval(tmpStr1[1])==0))
				{
					
					var oldvalue=FormRadio4[tmpStr1[3]];
					var olditem=GetItemName(oldvalue + "_0");
					
					document.additem.elements[tmpStr1[2]].checked=true;
					New_LockRadioList(tmpStr1[3],0);
					var tmpStr=document.additem.elements[tmpStr1[2]].value;
					
					//CAT Rules
					var tmpcat=FormRadio1[tmpStr1[3]].replace("CAG","");
					New_SelectCat(tmpcat);
					//Save new value to Array
					var tmpStr2=tmpStr.split("_");
					FormRadio4[tmpStr1[3]]=tmpStr2[4];
					calculate(eval("document.additem." + FormRadio1[tmpStr1[3]]),3);
					
					var newvalue=FormRadio4[tmpStr1[3]];
					var newitem=GetItemName(newvalue + "_0");
					
					if (olditem!=newitem) DisplayAlert1(tmpcat,3,prdname,olditem,newitem,GetCatName(tmpcat));
				}
				else
				//Checkbox List
				{
					olditem=0;
					
					if (document.additem.elements[tmpStr1[2]].checked==true) olditem=1;
					
					document.additem.elements[tmpStr1[2]].checked=true;
					New_LockItem(tmpStr1[2]);
					var tmpvalue=document.additem.elements[tmpStr1[2]].value;
					var tmpItem1=tmpvalue.split("_");
					var newitem=GetItemName(tmpItem1[0] + "_0");
					
					if (olditem==1) olditem=newitem;
					
					//CAT Rules
					var tmpcat=FormCB1[tmpStr1[3]].replace("CAG","")
					New_SelectCat(tmpcat);
					calculate(document.additem.elements[tmpStr1[2]],1);
					if (olditem!=newitem) DisplayAlert1(tmpcat,4,prdname,olditem,newitem,GetCatName(tmpcat));
				}
				
			} //It isnt Drop-down item
			} //tmpLoc<>''
			}
			while (tmpLoc!='');
		} //Have item
	}
	while (--j);
	
	//Follow the RULE TREE of "MUST" items	
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			var m=RuleCount+1;
			do
			{
				k=m-1;
				if ((eval(Rule1[k])==eval(tmp1[i])) && (Rule5[k]==0) && (Rule8[k]==0))
				{
					Rule5[k]=1;
					New_SetFromValues(Rule4[k],Rule2[k],Rule3[k],Rule6[k],Rule7[k]);
					break;
				}
			}
			while (--m);			
		}
	}
	while (--j);
	
	//Follow the RULE TREE of "MUST" CATs
	var tmp1=catrule1.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			var m=RuleCount+1;
			do
			{
				k=m-1;
				if ((eval(Rule1[k])==eval(tmp1[i])) && (Rule5[k]==0) && (Rule8[k]==1))
				{
					Rule5[k]=1;
					New_SetFromValues(Rule4[k],Rule2[k],Rule3[k],Rule6[k],Rule7[k]);
					break;
				}
			}
			while (--m);			
		}
	}
	while (--j);

}

//Reverse Field status Function
function New_RevFromValues(prdname,prorule1,prorule2,catrule1,catrule2)
{
var i=0;
var j=0;
var k=0;
var m=0;
var tmpindex=0;

	//Enable "CAN NOT" items
	var tmp1=prorule2.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			tmpindex=0;
			do
			{
			var tmpLoc=New_FindItemLocationMulti(tmp1[i],tmpindex);
			tmpindex++;
			
			if (tmpLoc!='')
			{
			var tmpStr1=tmpLoc.split("_");
			
			//Drop-down box
			if ((eval(tmpStr1[0])!=0) && (eval(tmpStr1[1])==1))
			{
				New_UnLockDropDownItem(tmpStr1[0],tmpStr1[2]);
			}
			else
			{
				//Radio List
				if ((eval(tmpStr1[0])==0) && (eval(tmpStr1[1])==0))
				{
					New_UnLockItem(tmpStr1[2]);
				}
				else
				//Checkbox List
				{
					New_UnLockItem(tmpStr1[2]);
				}
				
			} //It isnt Drop-down item
			} // tmpLoc<>''
			}
			while (tmpLoc!='');
		} //Have item
	}
	while (--j);
	
	//Enable "CAN NOT" CATs
	var tmp1=catrule2.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			var tmpStr1=New_FindGroupLocation(tmp1[i]);
			if (tmpStr1!="notfound")
			{
			var tmpStr2=tmpStr1.split("_");
			//Drop-down box
			if (eval(tmpStr2[1])==1)
			{
				New_UnlockDropDown(tmpStr2[2]);
			}
			else
			{
				//Radio List
				if (eval(tmpStr2[1])==0)
				{
					New_UnLockRadioList(tmpStr2[0]);
				}
				else
				//Checkbox List
				{
					New_UnLockCBList(tmpStr2[0],1);
				}
			
			} //It isnt Drop-down box
			}//notfound
		} //Have Cat		
	}
	while (--j);
	
	//Unlock "MUST" items
	var tmp1=prorule1.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			tmpindex=0;
			do
			{
			var tmpLoc=New_FindItemLocationMulti(tmp1[i],tmpindex);
			tmpindex++;
			
			if (tmpLoc!='')
			{
			var tmpStr1=tmpLoc.split("_");
			
			//Drop-down box
			if ((eval(tmpStr1[0])!=0) && (eval(tmpStr1[1])==1))
			{
				New_UnlockDropDown(tmpStr1[0]);
			}
			else
			{
				//Radio List
				if ((eval(tmpStr1[0])==0) && (eval(tmpStr1[1])==0))
				{
					New_UnLockRadioList(tmpStr1[3]);
				}
				else
				//Checkbox List
				{
					New_UnLockItem(tmpStr1[2]);
				}
				
			} //It isnt Drop-down item
			} // tmpLoc<>''
			}
			while (tmpLoc!='');
		} //Have item
	}
	while (--j);
	
	//No need to follow the RULE TREE of "MUST" items because they are still selected
	
	//Follow the RULE TREE of "MUST" CATs
	var tmp1=catrule1.split(",");
	j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			var m=RuleCount+1;
			do
			{
				k=m-1;
				if ((eval(Rule1[k])==eval(tmp1[i])) && (Rule5[k]==1) && (Rule8[k]==1))
				{
					New_RevFromValues(Rule4[k],Rule2[k],Rule3[k],Rule6[k],Rule7[k]);
					if (!New_CheckSelectedCat(Rule1[k])) Rule5[k]=0;
					break;
				}
			}
			while (--m);			
		}
	}
	while (--j);

}

//Uncheck Rule of The Item
function New_ReverseItem(tmpindex)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		if ((eval(Rule1[i])==eval(tmpindex)) && (Rule8[i]==0))
		{
			if ((Rule5[i]==1) || (pcv_Preset==1))
			{
				Rule5[i]=0;
				New_RevFromValues(Rule4[i],Rule2[i],Rule3[i],Rule6[i],Rule7[i]);
			}
			break;
		}
	}
	while (--j);
}

//Check Rule of The Item
function New_CheckItem(tmpindex)
{
	var i=0;
	var j=0;
	j=RuleCount+1;
	do
	{
		i=j-1;
		if ((eval(Rule1[i])==eval(tmpindex)) && (Rule8[i]==0))
		{
			if ((Rule5[i]==0) || (pcv_Preset==1))
			{
				Rule5[i]=1;
				New_SetFromValues(Rule4[i],Rule2[i],Rule3[i],Rule6[i],Rule7[i]);
			}
			break;
		}
	}
	while (--j);
}

//Get CAT Name Function
function New_GetCatName(iditem)
{
	var i=0;
	var j=CatCount+1;
	do
	{
		i=j-1;
		if (CatID[i]=="" + iditem)
		{
			return(CatName[i]);
			break;
		}
	}
	while (--j);
	return("");
}

//Check "MUST" and "CAN NOT" CAT List
function New_CheckCatBS(catrule1,catrule2)
{
	var i=0;
	var j=0;
	var k=0;
	var haveit=0;

	//Check "MUST" CATs
	var tmp1=catrule1.split(",");
	var j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			haveit=0;
			var tmpStr=New_FindGroupLocation(tmp1[i]);
			if (tmpStr!="notfound")
			{
			var tmpStr2=tmpStr.split("_")
			if (tmpStr2[1]==1)
			{
				if (eval(FormDrop3[tmpStr2[0]])>0) haveit=1;
			}
			else
			{
				if (tmpStr2[1]==0)
				{
					if (eval(FormRadio4[tmpStr2[0]])>0) haveit=1;
				}
				else
				{
					if (New_CheckCBHaveItem(tmpStr2[0])==true) haveit=1;
				}
			}
			if (haveit==0)
			{
				DisplayAlert2(tmp1[i],pcv_msg_btocm_6 + "'" + New_GetCatName(tmp1[i]) + "'");
				return("error");
				break;
			}
			}//notfound
		}
	}
	while (--j);
	
	//Check "CAN NOT" CATs
	var tmp1=catrule2.split(",");
	var j=tmp1.length;
	do
	{
		i=j-1;
		if (tmp1[i]!="")
		{
			haveit=0;
			var tmpStr=New_FindGroupLocation(tmp1[i]);
			if (tmpStr!="notfound")
			{
			var tmpStr2=tmpStr.split("_")
			if (tmpStr2[1]==1)
			{
				if (eval(FormDrop3[tmpStr2[0]])>0) haveit=1;
			}
			else
			{
				if (tmpStr2[1]==0)
				{
					if (eval(FormRadio4[tmpStr2[0]])>0) haveit=1;
				}
				else
				{
					if (New_CheckCBHaveItem(tmpStr2[0])==true) haveit=1;
				}
			}
			if (haveit==1)
			{
				DisplayAlert2(tmp1[i],pcv_msg_btocm_7 + "'" + New_GetCatName(tmp1[i]) + "'");
				return("error");
				break;
			}
			}//notfound
		}
	}
	while (--j);
	return("");
}
	
function CheckCatBeforeSubmit()
{
	var i=0;
	var j=RuleCount+1;
	var tmp1="";
	ClearCMMsgs();
	do
	{
		i=j-1;
		tmp1="";
		if (Rule5[i]==1) 
		{
			tmp1=New_CheckCatBS(Rule6[i],Rule7[i]);
			if (tmp1!="")
			{
				return("no");
				break;
			}
		}
	}
	while (--j);
	return("ok");
}

//Reset Rules Array
function new_ResetRules()
{
	var i=0;
	var j=RuleCount+1;
	var tmp1="";
	
	do
	{
		i=j-1;
		Rule5[i]=0;
	}
	while (--j);
}

//Find Item Location in the Form
function PresetValues()
{
	pcv_Preset=1;
	new_ResetRules();
	var tmpArr=new Array();
	var tmpArrGrp=new Array();
	var tmpCount=0;
	var i=0;
	var k=0;
	var m=0;
	var objElems = document.additem.elements;
	var j=objElems.length;
	do
	{
		i=j-1;
		var tmptype=objElems[i].type;
		
		if (tmptype=="select-one")
		{
			var tmpStr1=objElems[i].value;
			var tmpStr2=tmpStr1.split("_");
			tmpCount++;
			tmpArr[tmpCount-1]=tmpStr2[4];
			
			var tmpStr2=objElems[i].name;
			tmpStr2=tmpStr2.replace("CAG","");
			var tmpStr=New_FindGroupLocation(tmpStr2);
			if (tmpStr!="notfound")
			{
				tmpStr1=tmpStr.split("_");
				tmpArrGrp[tmpCount-1]=tmpStr1[2];
			}//notfound
		}
		else
		{
			if (tmptype=="radio")
			{
				if (objElems[i].checked==true)
				{
					var tmpStr1=objElems[i].value;
					var tmpStr2=tmpStr1.split("_");
					tmpCount++;
					tmpArr[tmpCount-1]=tmpStr2[4];
				}
			}
			else
			{
				if (tmptype=="checkbox")
				{
					if (objElems[i].checked==true)
					{
						var tmpStr1=objElems[i].value;
						var tmpStr2=tmpStr1.split("_");
						tmpCount++;
						tmpArr[tmpCount-1]=tmpStr2[4];
					}
				}
			}
		}
	}
	while (--j);
	
	j=tmpCount;
	if (tmpCount!=0)
	{
	do
	{
		i=j-1;
		if (eval(tmpArr[i])!=0)
		{
			if (tmpArrGrp[i]==undefined)
			{
				var tmpindex=0;
				do
				{
					var tmpStr1=New_FindItemLocationMulti(tmpArr[i],tmpindex);
					var tmpStr2=tmpStr1.split("_");
					tmpindex++;
				}
				while (eval(tmpStr2[1])==1);
			}
			else
			{
				var tmpindex=0;
				do
				{
					var tmpStr1=New_FindItemLocationMulti(tmpArr[i],tmpindex);
					var tmpStr2=tmpStr1.split("_");
					tmpindex++;
				}
				while (eval(tmpStr2[1])!=1);
			}
				
			//Drop-down
			if (eval(tmpStr2[1])==1)
			{
				if ((document.additem.elements[tmpArrGrp[i]].disabled==false) && (document.additem.elements[tmpArrGrp[i]].options[tmpStr2[2]].style.color!="gray") && ((document.additem.elements[tmpArrGrp[i]].options[tmpStr2[2]].selected==true) || (document.additem.elements[tmpArrGrp[i]].value==document.additem.elements[tmpArrGrp[i]].options[tmpStr2[2]].value)))
				{
					CheckPreValue(document.additem.elements[tmpArrGrp[i]],1,1);
				}
			}
			else
			{
				//Radio
				if (eval(tmpStr2[1])==0)
				{
					if ((document.additem.elements[tmpStr2[2]].disabled==false) && (document.additem.elements[tmpStr2[2]].checked==true))
					{
						CheckPreValue(document.additem.elements[tmpStr2[2]],2,1);
					}
				}
				//Checkbox
				else
				{
					if ((document.additem.elements[tmpStr2[2]].disabled==false) && (document.additem.elements[tmpStr2[2]].checked==true))
					{
						CheckBoxPreValue(document.additem.elements[tmpStr2[2]],1)
					}
				}
			}
		} //IDItem <> 0
	}
	while (--j);
	}
	pcv_Preset=0;
	return(true);
}

function New_CheckSelectedCat(tmpcatid)
{
	var i=0;
	var j=0;
	var k=0;
	var haveit=0;

	//Check "MUST" CATs
	var tmp1=tmpcatid;

	if (tmp1!="")
	{
		haveit=0;
		var tmpStr=New_FindGroupLocation(tmp1);
		if (tmpStr!="notfound")
		{
			var tmpStr2=tmpStr.split("_")
			if (tmpStr2[1]==1)
			{
				if (eval(FormDrop3[tmpStr2[0]])>0) haveit=1;
			}
			else
			{
				if (tmpStr2[1]==0)
				{
					if (eval(FormRadio4[tmpStr2[0]])>0) haveit=1;
				}
				else
				{
					if (New_CheckCBHaveItem(tmpStr2[0])==true) haveit=1;
				}
			}
		}//notfound
	}
	
	if (haveit==1)
	{
		return(true);
	}
	else
	{
		return(false);
	}
}

function parseFloatEx(tmpvalue)
{
	var tmp1=""+tmpvalue;
	if (scDecSign==",")	tmp1=tmp1.replace(",",".");
	return(parseFloat(tmp1));
}
