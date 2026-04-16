<%@ LANGUAGE = VBScript.Encode %>
<%
'This file is part of ProductCart, an ecommerce application developed and sold by NetSource Commerce. ProductCart, its source code, the ProductCart name and logo are property of NetSource Commerce. Copyright 2001-2015. All rights reserved. You are not allowed to use, alter, distribute and/or resell any parts of ProductCart's source code without the written consent of NetSource Commerce. To contact NetSource Commerce, please visit www.productcart.com.
%>
<% Section="layout" %>
<%PmAdmin=1%>
<!--#include file="adminv.asp"--> 
<!--#include file="../includes/common.asp"-->
<!--#include file="../includes/languagesCP.asp"-->
<!--#include file="../includes/common_checkout.asp"-->
<%
pageTitle=dictLanguage.Item(Session("language")&"_pcAppBtnManage")
pageIcon="pcv4_icon_settings.png"
%>
<!--#include file="AdminHeader.asp"-->
<%
pcPageName="pcws_MyApps.asp"
%>
<%
'// START: Refresh Token
If request("refresh")<>"" Then
    pcf_UpdateToken()
End If
'// END: Refresh Token
%>
<style type="text/css">
.panel-title a {
 	display: block; 
}
/*! ========================================================================
 * Bootstrap Toggle: bootstrap-toggle.css v2.2.0
 * http://www.bootstraptoggle.com
 * ========================================================================
 * Copyright 2014 Min Hur, The New York Times Company
 * Licensed under MIT
 * ======================================================================== */
.checkbox label .toggle,.checkbox-inline .toggle{margin-left:-20px;margin-right:5px}
.toggle{position:relative;overflow:hidden}
.toggle input[type=checkbox]{display:none}
.toggle-group{position:absolute;width:200%;top:0;bottom:0;left:0;transition:left .35s;-webkit-transition:left .35s;-moz-user-select:none;-webkit-user-select:none}
.toggle.off .toggle-group{left:-100%}
.toggle-on{position:absolute;top:0;bottom:0;left:0;right:50%;margin:0;border:0;border-radius:0}
.toggle-off{position:absolute;top:0;bottom:0;left:50%;right:0;margin:0;border:0;border-radius:0}
.toggle-handle{position:relative;margin:0 auto;padding-top:0;padding-bottom:0;height:100%;width:0;border-width:0 1px}
.toggle.btn{min-width:59px;min-height:34px}
.toggle-on.btn{padding-right:24px}
.toggle-off.btn{padding-left:24px}
.toggle.btn-lg{min-width:79px;min-height:45px}
.toggle-on.btn-lg{padding-right:31px}
.toggle-off.btn-lg{padding-left:31px}
.toggle-handle.btn-lg{width:40px}
.toggle.btn-sm{min-width:50px;min-height:30px}
.toggle-on.btn-sm{padding-right:20px}
.toggle-off.btn-sm{padding-left:20px}
.toggle.btn-xs{min-width:35px;min-height:22px}
.toggle-on.btn-xs{padding-right:12px}
.toggle-off.btn-xs{padding-left:12px}
</style>

<% ' START show message, if any %>
	<!--#include file="pcv4_showMessage.asp"-->
<% 	' END show message %>




<%
query="SELECT pcPCWS_Uid, pcPCWS_AuthToken, pcPCWS_Username, pcPCWS_Password FROM pcWebServiceSettings;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    pcv_strUid = rs("pcPCWS_Uid")
    pcv_AuthToken = rs("pcPCWS_AuthToken")  
    pcv_strUsername = rs("pcPCWS_Username")  
    pcv_strPassword = enDeCrypt(rs("pcPCWS_Password"), scCrypPass)          
End If
Set rs=nothing
%>

<div class="apps" data-ng-controller="appsCtrl">

    <!--#include file="pcws_Navigation.asp"--> 

    <div class="container-fluid">
        <div class="row">
            <div class="col-md-12">
                <htmldiv content="myhtml">
                    <%= pcf_displayMarket("/MyApps/" & pcv_strUid, pcv_AuthToken) %>
                </htmldiv>
            </div>
        </div>
    </div>

</div>

<script>
// Local Variable Overrides
<%
query="SELECT [pcPCWS_FeatureCode], pcPCWS_IsActive FROM pcWebServiceFeatures;"
Set rs=connTemp.execute(query)
If Not rs.eof Then
    Do While Not rs.eof
        pcv_strFeatureCode = rs("pcPCWS_FeatureCode")
        pcv_intIsActive = rs("pcPCWS_IsActive")
        If pcv_intIsActive = 1 Then
            pcv_boolIsActive = true
        Else
            pcv_boolIsActive = false
        End If
        %>
        <%=pcv_strFeatureCode %> = <%=lcase(pcv_boolIsActive) %>;
        <%
        rs.movenext
    Loop
End If
Set rs=nothing
%>
</script>
<script>
/*! ========================================================================
 * Bootstrap Toggle: bootstrap-toggle.js v2.2.0
 * http://www.bootstraptoggle.com
 * ========================================================================
 * Copyright 2014 Min Hur, The New York Times Company
 * Licensed under MIT
 * ======================================================================== */
+function(a){"use strict";function b(b){return this.each(function(){var d=a(this),e=d.data("bs.toggle"),f="object"==typeof b&&b;e||d.data("bs.toggle",e=new c(this,f)),"string"==typeof b&&e[b]&&e[b]()})}var c=function(b,c){this.$element=a(b),this.options=a.extend({},this.defaults(),c),this.render()};c.VERSION="2.2.0",c.DEFAULTS={on:"On",off:"Off",onstyle:"primary",offstyle:"default",size:"normal",style:"",width:null,height:null},c.prototype.defaults=function(){return{on:this.$element.attr("data-on")||c.DEFAULTS.on,off:this.$element.attr("data-off")||c.DEFAULTS.off,onstyle:this.$element.attr("data-onstyle")||c.DEFAULTS.onstyle,offstyle:this.$element.attr("data-offstyle")||c.DEFAULTS.offstyle,size:this.$element.attr("data-size")||c.DEFAULTS.size,style:this.$element.attr("data-style")||c.DEFAULTS.style,width:this.$element.attr("data-width")||c.DEFAULTS.width,height:this.$element.attr("data-height")||c.DEFAULTS.height}},c.prototype.render=function(){this._onstyle="btn-"+this.options.onstyle,this._offstyle="btn-"+this.options.offstyle;var b="large"===this.options.size?"btn-lg":"small"===this.options.size?"btn-sm":"mini"===this.options.size?"btn-xs":"",c=a('<label class="btn">').html(this.options.on).addClass(this._onstyle+" "+b),d=a('<label class="btn">').html(this.options.off).addClass(this._offstyle+" "+b+" active"),e=a('<span class="toggle-handle btn btn-default">').addClass(b),f=a('<div class="toggle-group">').append(c,d,e),g=a('<div class="toggle btn" data-toggle="toggle">').addClass(this.$element.prop("checked")?this._onstyle:this._offstyle+" off").addClass(b).addClass(this.options.style);this.$element.wrap(g),a.extend(this,{$toggle:this.$element.parent(),$toggleOn:c,$toggleOff:d,$toggleGroup:f}),this.$toggle.append(f);var h=this.options.width||Math.max(c.outerWidth(),d.outerWidth())+e.outerWidth()/2,i=this.options.height||Math.max(c.outerHeight(),d.outerHeight());c.addClass("toggle-on"),d.addClass("toggle-off"),this.$toggle.css({width:h,height:i}),this.options.height&&(c.css("line-height",c.height()+"px"),d.css("line-height",d.height()+"px")),this.update(!0),this.trigger(!0)},c.prototype.toggle=function(){this.$element.prop("checked")?this.off():this.on()},c.prototype.on=function(a){return this.$element.prop("disabled")?!1:(this.$toggle.removeClass(this._offstyle+" off").addClass(this._onstyle),this.$element.prop("checked",!0),void(a||this.trigger()))},c.prototype.off=function(a){return this.$element.prop("disabled")?!1:(this.$toggle.removeClass(this._onstyle).addClass(this._offstyle+" off"),this.$element.prop("checked",!1),void(a||this.trigger()))},c.prototype.enable=function(){this.$toggle.removeAttr("disabled"),this.$element.prop("disabled",!1)},c.prototype.disable=function(){this.$toggle.attr("disabled","disabled"),this.$element.prop("disabled",!0)},c.prototype.update=function(a){this.$element.prop("disabled")?this.disable():this.enable(),this.$element.prop("checked")?this.on(a):this.off(a)},c.prototype.trigger=function(b){this.$element.off("change.bs.toggle"),b||this.$element.change(),this.$element.on("change.bs.toggle",a.proxy(function(){this.update()},this))},c.prototype.destroy=function(){this.$element.off("change.bs.toggle"),this.$toggleGroup.remove(),this.$element.removeData("bs.toggle"),this.$element.unwrap()};var d=a.fn.bootstrapToggle;a.fn.bootstrapToggle=b,a.fn.bootstrapToggle.Constructor=c,a.fn.toggle.noConflict=function(){return a.fn.bootstrapToggle=d,this},a(function(){a("input[type=checkbox][data-toggle^=toggle]").bootstrapToggle()}),a(document).on("click.bs.toggle","div[data-toggle^=toggle]",function(b){var c=a(this).find("input[type=checkbox]");c.bootstrapToggle("toggle"),b.preventDefault()})}(jQuery);
</script>
<!--#include file="AdminFooter.asp"-->