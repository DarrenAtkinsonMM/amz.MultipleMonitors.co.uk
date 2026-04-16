/***********************************************************
InnovaStudio Live Editor 1.7.3 - IE
© 2014, InnovaStudio (www.innovastudio.com). All rights reserved.
************************************************************/
var UA = navigator.userAgent.toLowerCase();
var isIE = (UA.indexOf('msie') >= 0) ? true : false;
var isNS = (UA.indexOf('mozilla') >= 0) ? true : false;
var isIE9 = (UA.indexOf('msie 9.0') >= 0) ? true : false;
var oUtil = new InnovaEditorUtil();

function InnovaEditorUtil() {
    this.langDir = "en-US";
    try {
        if (LanguageDirectory) this.langDir = LanguageDirectory
    } catch (e) {}
    var oScripts = document.getElementsByTagName("script");
    for (var i = 0; i < oScripts.length; i++) {
        var sSrc = oScripts[i].src.toLowerCase();
        if (sSrc.indexOf("scripts/editor.js") != -1) this.scriptPath = oScripts[i].src.replace(/editor.js/ig, "")
    }
    this.scriptPathLang = this.scriptPath + "language/" + this.langDir + "/";
    if (this.langDir == "en-US") document.write("<scr" + "ipt src='" + this.scriptPathLang + "editor_lang.js'></scr" + "ipt>");
    this.oName;
    this.oEditor;
    this.obj;
    this.oSel;
    this.sType;
    this.bInside = bInside;
    this.useSelection = true;
    this.arrEditor = [];
    this.onSelectionChanged = function () {
        return true
    };
    this.activeElement;
    this.activeModalWin;
    this.Table;
    this.protocol = window.location.protocol;
    this.spcCharCode = [
        [169, "&copy;"],
        [163, "&pound;"],
        [174, "&reg;"],
        [233, "&eacute;"],
        [201, "&Eacute;"],
        [8364, "&euro;"],
        [8220, "\""]
    ];
    this.spcChar = [];
    this.loadSpecialCharCode = function (spcCharCodes) {
        if (spcCharCodes != null) {
            this.spcCharCode = spcCharCodes;
            this.spcChar = []
        }
        for (var i = 0; i < this.spcCharCode.length; i++) {
            this.spcChar[i] = [new RegExp(String.fromCharCode(this.spcCharCode[i][0]), "g"), this.spcCharCode[i][1]]
        }
    };
    for (var i = 0; i < this.spcCharCode.length; i++) {
        this.spcChar[i] = [new RegExp(String.fromCharCode(this.spcCharCode[i][0]), "g"), this.spcCharCode[i][1]]
    }
    this.replaceSpecialChar = function (sHTML) {
        for (var i = 0; i < this.spcChar.length; i++) {
            sHTML = sHTML.replace(this.spcChar[i][0], this.spcChar[i][1])
        }
        return sHTML
    };
    this.initializeEditor = function (tselector, opt) {
        var allText = [],
            txt, edtCnt;
        if (typeof (tselector) == "object" && tselector.tagName && tselector.tagName == "TEXTAREA") {
            allText[0] = tselector
        } else if (tselector.substr(0, 1) == "#") {
            txt = document.getElementById(tselector.substr(1));
            if (!txt) return;
            allText[0] = txt
        } else {
            var all = document.getElementsByTagName("TEXTAREA");
            for (var i = 0; i < all.length; i++) {
                if (all[i].className == tselector) {
                    allText[allText.length] = all[i]
                }
            }
        }
        for (var i = 0; i < allText.length; i++) {
            txt = allText[i];
            if (txt.id || txt.id == "") txt.id = "editorarea" + i;
            edtCnt = document.createElement("DIV");
            edtCnt.id = "innovaeditor" + i;
            txt.parentNode.insertBefore(edtCnt, txt);
            window["oEdit" + i] = new InnovaEditor("oEdit" + i);
            var objStyle;
            if (window.getComputedStyle) {
                objStyle = window.getComputedStyle(txt, null)
            } else if (txt.currentStyle) {
                objStyle = txt.currentStyle
            } else {
                objStyle = {
                    width: window["oEdit" + i].width,
                    height: window["oEdit" + i].height
                }
            }
            window["oEdit" + i].width = objStyle.width;
            window["oEdit" + i].height = objStyle.height;
            if (opt) {
                for (var it in opt) {
                    window["oEdit" + i][it] = opt[it]
                }
            }
            window["oEdit" + i].REPLACE(txt.id, "innovaeditor" + i)
        }
    }
};

function bInside(oElement) {
    while (oElement != null) {
        if (oElement.contentEditable == "true") return true;
        oElement = oElement.parentElement
    }
    return false
};

function checkFocus() {
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (oSel.parentElement != null) {
        if (!bInside(oSel.parentElement())) return false
    } else {
        if (!bInside(oSel.item(0))) return false
    }
    return true
};

function iwe_focus() {
    var oEditor = eval("idContent" + this.oName);
    oEditor.focus()
};

function set_focus(type) {
    var oEditor = eval("idContent" + this.oName);
    oEditor.focus();
    try {
        if (this.rangeBookmark != null) {
            var oSel = oEditor.document.selection;
            var oRange = oSel.createRange();
            var bmRange = this.rangeBookmark;
            if (bmRange.parentElement()) {
                oRange.moveToElementText(bmRange.parentElement());
                oRange.setEndPoint("StarttoStart", bmRange);
                oRange.setEndPoint("EndToEnd", bmRange);
                oRange.select()
            }
            if (type) {
                oRange.moveToElementText(oEditor.document.body);
                oRange.select();
                if (type == "start") {
                    oRange.collapse()
                } else if (type == "end") {
                    oRange.collapse(false)
                }
                oRange.select()
            }
        } else if (this.controlBookmark != null) {
            var oSel = oEditor.document.body.createControlRange();
            oSel.add(this.controlBookmark);
            oSel.select()
        }
    } catch (e) {}
};

function bookmark_selection() {
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection;
    var oRange = oSel.createRange();
    if (oSel.type == "None" || oSel.type == "Text") {
        this.rangeBookmark = oRange;
        this.controlBookmark = null
    } else {
        this.controlBookmark = oRange.item(0);
        this.rangeBookmark = null
    }
};

function InnovaEditor(oName) {
    this.oName = oName;
    this.RENDER = RENDER;
    this.onRender = function () {};
    this.init = initISEditor;
    this.IsSecurityRestricted = false;
    this.loadHTML = loadHTML;
    this.putHTML = putHTML;
    this.getHTMLBody = getHTMLBody;
    this.getXHTMLBody = getXHTMLBody;
    this.getHTML = getHTML;
    this.getXHTML = getXHTML;
    this.getTextBody = getTextBody;
    this.initialRefresh = false;
    this.preserveSpace = false;
    this.bInside = bInside;
    this.checkFocus = checkFocus;
    this.focus = iwe_focus;
    this.setFocus = set_focus;
    this.bookmarkSelection = bookmark_selection;
    this.disableFocusOnLoad = false;
    this.onKeyPress = function (evt) {
        return true
    };
    this.styleSelectionHoverBg = "#cccccc";
    this.styleSelectionHoverFg = "white";
    this.styleSelectorPrefix = "";
    this.cleanEmptySpan = cleanEmptySpan;
    this.cleanFonts = cleanFonts;
    this.cleanTags = cleanTags;
    this.replaceTags = replaceTags;
    this.cleanDeprecated = cleanDeprecated;
    this.doClean = doClean;
    this.applySpanStyle = applySpanStyle;
    this.applyLine = applyLine;
    this.applyBold = applyBold;
    this.applyItalic = applyItalic;
    this.doOnPaste = doOnPaste;
    this.isAfterPaste = false;
    this.doCmd = doCmd;
    this.applyParagraph = applyParagraph;
    this.applyFontName = applyFontName;
    this.applyFontSize = applyFontSize;
    this.applyBullets = applyBullets;
    this.applyNumbering = applyNumbering;
    this.applyJustifyLeft = applyJustifyLeft;
    this.applyJustifyCenter = applyJustifyCenter;
    this.applyJustifyRight = applyJustifyRight;
    this.applyJustifyFull = applyJustifyFull;
    this.applyBlockDirLTR = applyBlockDirLTR;
    this.applyBlockDirRTL = applyBlockDirRTL;
    this.doPaste = doPaste;
    this.doPasteText = doPasteText;
    this.applySpan = applySpan;
    this.makeAbsolute = makeAbsolute;
    this.insertHTML = insertHTML;
    this.clearAll = clearAll;
    this.insertCustomTag = insertCustomTag;
    this.selectParagraph = selectParagraph;
    this.hide = hide;
    this.width = "700";
    this.height = "350";
    this.publishingPath = "";
    var oScripts = document.getElementsByTagName("script");
    for (var i = 0; i < oScripts.length; i++) {
        var sSrc = oScripts[i].src.toLowerCase();
        if (sSrc.indexOf("scripts/editor.js") != -1) this.scriptPath = oScripts[i].src.replace(/editor.js/, "")
    }
    this.dialogPath = this.scriptPath + "common/";
    this.GetEmoticons = GetEmoticons;
    this.insertEmoticon = insertEmoticon;
    this.applyQuote = applyQuote;
    this.iconPath = "icons/";
    this.iconWidth = 29;
    this.iconHeight = 25;
    this.iconOffsetTop;
    this.dropTopAdjustment = -1;
    this.dropLeftAdjustment = 0;
    this.heightAdjustment = -70;
    this.runtimeBorder = runtimeBorder;
    this.runtimeBorderOn = runtimeBorderOn;
    this.runtimeBorderOff = runtimeBorderOff;
    this.IsRuntimeBorderOn = true;
    this.runtimeStyles = runtimeStyles;
    this.applyColor = applyColor;
    this.oColor1 = new ColorPicker("oColor1", this.oName);
    this.oColor2 = new ColorPicker("oColor2", this.oName);
    this.expandSelection = expandSelection;
    this.fullScreen = fullScreen;
    this.stateFullScreen = false;
    this.onFullScreen = function () {
        return true
    };
    this.onNormalScreen = function () {
        return true
    };
    this.arrElm = new Array(300);
    this.getElm = iwe_getElm;
    this.features = [];
    this.btnParagraph = false;
    this.btnFontName = false;
    this.btnFontSize = false;
    this.btnCut = false;
    this.btnCopy = false;
    this.btnPaste = false;
    this.btnPasteText = false;
    this.btnUndo = false;
    this.btnRedo = false;
    this.btnBold = false;
    this.btnItalic = false;
    this.btnUnderline = false;
    this.btnStrikethrough = false;
    this.btnSuperscript = false;
    this.btnSubscript = false;
    this.btnJustifyLeft = false;
    this.btnJustifyCenter = false;
    this.btnJustifyRight = false;
    this.btnJustifyFull = false;
    this.btnNumbering = false;
    this.btnBullets = false;
    this.btnIndent = false;
    this.btnOutdent = false;
    this.btnLTR = false;
    this.btnRTL = false;
    this.btnForeColor = false;
    this.btnBackColor = false;
    this.btnTable = false;
    this.btnLine = false;
    this.tabs = [
        ["tabHome", "Home", ["group1", "group2", "group4"]],
        ["tabStyle", "Insert", ["group3"]]
    ];
    this.groups = [
        ["group1", "", ["Bold", "Italic", "Underline", "FontDialog", "ForeColor", "TextDialog", "RemoveFormat"]],
        ["group2", "", ["Bullets", "Numbering", "JustifyLeft", "JustifyCenter", "JustifyRight"]],
        ["group3", "", ["LinkDialog", "ImageDialog", "Table", "TableDialog"]],
        ["group4", "", ["Undo", "Redo", "FullScreen", "SourceDialog"]]
    ];
    this.toolbarMode = 2;
    this.showResizeBar = true;
    this.pasteTextOnCtrlV = false;
    this.dialogSize = {
        "Preview": {
            w: 900,
            h: 600
        },
        "TableDialog": {
            w: 785,
            h: 500
        },
        "ImageDialog": {
            w: 755,
            h: 545
        },
        "TextDialog": {
            w: 375,
            h: 475
        },
        "YoutubeDialog": {
            w: 421,
            h: 545
        },
        "LinkDialog": {
            w: 605,
            h: 475
        },
        "SourceDialog": {
            w: 700,
            h: 450
        },
        "CompleteTextDialog": {
            w: 815,
            h: 470
        },
        "FontDialog": {
            w: 500,
            h: 470
        },
        "FlashDialog": {
            w: 390,
            h: 195
        },
        "BookmarkDialog": {
            w: 360,
            h: 240
        },
        "CharsDialog": {
            w: 700,
            h: 122
        },
        "SearchDialog": {
            w: 370,
            h: 140
        }
    };
    this.setDialogSize = function (name, dim) {
        this.dialogSize[name] = dim
    };
    this.fileBrowser = "";
    this.enableFlickr = false;
    this.enableImageStyles = true;
    this.enableYTVideoStyles = true;
    this.enableCssButtons = false;
    this.enableLightbox = false;
    this.enableTableAutoformat = false;
    this.flickrUser = "ysw.insite";
    this.cmdAssetManager = "";
    this.cmdFileManager = "";
    this.cmdImageManager = "";
    this.cmdMediaManager = "";
    this.cmdFlashManager = "";
    this.btnContentBlock = false;
    this.cmdContentBlock = ";";
    this.btnInternalLink = false;
    this.cmdInternalLink = ";";
    this.insertLink = insertLink;
    this.btnCustomObject = false;
    this.cmdCustomObject = ";";
    this.btnInternalImage = false;
    this.cmdInternalImage = ";";
    this.css = "";
    this.arrStyle = [];
    this.isCssLoaded = false;
    this.openStyleSelect = openStyleSelect;
    this.arrParagraph = [
        [getTxt("Heading 1"), "H1"],
        [getTxt("Heading 2"), "H2"],
        [getTxt("Heading 3"), "H3"],
        [getTxt("Heading 4"), "H4"],
        [getTxt("Heading 5"), "H5"],
        [getTxt("Heading 6"), "H6"],
        [getTxt("Preformatted"), "PRE"],
        [getTxt("Normal (P)"), "P"],
        [getTxt("Normal (DIV)"), "DIV"]
    ];
    this.arrFontName = ["Impact, Charcoal, sans-serif", "Palatino Linotype, Book Antiqua, Palatino, serif", "Tahoma, Geneva, sans-serif", "Century Gothic, sans-serif", "Lucida Sans Unicode, Lucida Grande, sans-serif", "Times New Roman, Times, serif", "Arial Narrow, sans-serif", "Verdana, Geneva, sans-serif", "Copperplate Gothic Light, sans-serif", "Lucida Console, Monaco, monospace", "Gill Sans MT, sans-serif", "Trebuchet MS, Helvetica, sans-serif", "Courier New, Courier, monospace", "Arial, Helvetica, sans-serif", "Georgia, Serif", "Garamond, Serif"];
    this.arrFontSize = [
        [getTxt("Size 1"), "8pt"],
        [getTxt("Size 2"), "10pt"],
        [getTxt("Size 3"), "12pt"],
        [getTxt("Size 4"), "14pt"],
        [getTxt("Size 5"), "18pt"],
        [getTxt("Size 6"), "24pt"],
        [getTxt("Size 7"), "36pt"]
    ];
    this.arrCustomTag = [];
    this.docType = "<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">";
    this.html = "<html>";
    this.headContent = "";
    this.preloadHTML = "";
    this.originalContent = "";
    this.isContentChanged = isContentChanged;
    this.onSave = function () {
        document.getElementById("iwe_btnSubmit" + this.oName).click()
    };
    this.useBR = false;
    this.useDIV = true;
    this.returnKeyMode = -1;
    this.doUndo = doUndo;
    this.doRedo = doRedo;
    this.saveForUndo = saveForUndo;
    this.arrUndoList = [];
    this.arrRedoList = [];
    this.useTagSelector = true;
    this.TagSelectorPosition = "bottom";
    this.moveTagSelector = moveTagSelector;
    this.selectElement = selectElement;
    this.removeTag = removeTag;
    this.doClick_TabCreate = doClick_TabCreate;
    this.doRefresh_TabCreate = doRefresh_TabCreate;
    this.arrCustomButtons = [
        ["CustomName1", "alert(0)", "caption here", "btnSave.gif"],
        ["CustomName2", "alert(0)", "caption here", "btnSave.gif"]
    ];
    this.customDialogShow = customDialogShow;
    this.customDialog = [];
    this.onSelectionChanged = function () {
        return true
    };
    this.spellCheckMode = "ieSpell";
    this.encodeIO = false;

    this.changeHeight = changeHeight;
    this.fixWord = fixWord;
    this.rangeBookmark = null;
    this.controlBookmark = null;
    this.REPLACE = REPLACE;
    this.idTextArea;
    this.mode = "XHTMLBody";
    var me = this;
    this.tbar = new ISToolbarManager(this.oName);
    this.tbar.iconPath = this.scriptPath + this.iconPath
};

function changeActiveEditor(oName) {
    var edtObj = eval(oName);
    var edtFrm = document.getElementById("idContent" + oName);
    oUtil.activeElement = null;
    oUtil.oName = oName;
    oUtil.oEditor = edtFrm.contentWindow;
    oUtil.obj = edtObj;
    edtObj.hide()
}

function saveForUndo() {
    var oEditor = eval("idContent" + this.oName);
    var obj = eval(this.oName);
    if (obj.arrUndoList[0])
        if (oEditor.document.body.innerHTML == obj.arrUndoList[0][0]) return;
    for (var i = 20; i > 1; i--) obj.arrUndoList[i - 1] = obj.arrUndoList[i - 2];
    obj.focus();
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (sType == "None") obj.arrUndoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "None"];
    else if (sType == "Text") obj.arrUndoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "Text"];
    else if (sType == "Control") {
        oSel.item(0).selThis = "selThis";
        obj.arrUndoList[0] = [oEditor.document.body.innerHTML, null, "Control"];
        oSel.item(0).removeAttribute("selThis", 0)
    }
    this.arrRedoList = [];
    if (this.btnUndo) this.tbar.btns["btnUndo" + this.oName].setState(1);
    if (this.btnRedo) this.tbar.btns["btnRedo" + this.oName].setState(5)
};

function doUndo() {
    var oEditor = eval("idContent" + this.oName);
    var obj = eval(this.oName);
    if (!obj.arrUndoList[0]) return;
    for (var i = 20; i > 1; i--) obj.arrRedoList[i - 1] = obj.arrRedoList[i - 2];
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (sType == "None") this.arrRedoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "None"];
    else if (sType == "Text") this.arrRedoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "Text"];
    else if (sType == "Control") {
        oSel.item(0).selThis = "selThis";
        this.arrRedoList[0] = [oEditor.document.body.innerHTML, null, "Control"];
        oSel.item(0).removeAttribute("selThis", 0)
    }
    sHTML = obj.arrUndoList[0][0];
    sHTML = fixPathEncode(sHTML);
    oEditor.document.body.innerHTML = sHTML;
    fixPathDecode(oEditor);
    this.runtimeBorder(false);
    this.runtimeStyles();
    var oRange = oEditor.document.body.createTextRange();
    if (obj.arrUndoList[0][2] == "None") {
        oRange.moveToBookmark(obj.arrUndoList[0][1]);
        oRange.select()
    } else if (obj.arrUndoList[0][2] == "Text") {
        oRange.moveToBookmark(obj.arrUndoList[0][1]);
        oRange.select()
    } else if (obj.arrUndoList[0][2] == "Control") {
        for (var i = 0; i < oEditor.document.all.length; i++) {
            if (oEditor.document.all[i].selThis == "selThis") {
                var oSelRange = oEditor.document.body.createControlRange();
                oSelRange.add(oEditor.document.all[i]);
                oSelRange.select();
                oEditor.document.all[i].removeAttribute("selThis", 0)
            }
        }
    }
    for (var i = 0; i < 19; i++) obj.arrUndoList[i] = obj.arrUndoList[i + 1];
    obj.arrUndoList[19] = null;
    realTime(this.oName)
};

function doRedo() {
    var oEditor = eval("idContent" + this.oName);
    var obj = eval(this.oName);
    if (!obj.arrRedoList[0]) return;
    for (var i = 20; i > 1; i--) obj.arrUndoList[i - 1] = obj.arrUndoList[i - 2];
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (sType == "None") obj.arrUndoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "None"];
    else if (sType == "Text") obj.arrUndoList[0] = [oEditor.document.body.innerHTML, oEditor.document.selection.createRange().getBookmark(), "Text"];
    else if (sType == "Control") {
        oSel.item(0).selThis = "selThis";
        this.arrUndoList[0] = [oEditor.document.body.innerHTML, null, "Control"];
        oSel.item(0).removeAttribute("selThis", 0)
    }
    sHTML = obj.arrRedoList[0][0];
    sHTML = fixPathEncode(sHTML);
    oEditor.document.body.innerHTML = sHTML;
    fixPathDecode(oEditor);
    this.runtimeBorder(false);
    this.runtimeStyles();
    var oRange = oEditor.document.body.createTextRange();
    if (obj.arrRedoList[0][2] == "None") {
        oRange.moveToBookmark(obj.arrRedoList[0][1])
    } else if (obj.arrRedoList[0][2] == "Text") {
        oRange.moveToBookmark(obj.arrRedoList[0][1]);
        oRange.select()
    } else if (obj.arrRedoList[0][2] == "Control") {
        for (var i = 0; i < oEditor.document.all.length; i++) {
            if (oEditor.document.all[i].selThis == "selThis") {
                var oSelRange = oEditor.document.body.createControlRange();
                oSelRange.add(oEditor.document.all[i]);
                oSelRange.select();
                oEditor.document.all[i].removeAttribute("selThis", 0)
            }
        }
    }
    for (var i = 0; i < 19; i++) obj.arrRedoList[i] = obj.arrRedoList[i + 1];
    obj.arrRedoList[19] = null;
    realTime(this.oName)
};
var bOnSubmitOriginalSaved = false;

function REPLACE(idTextArea, dvId) {
    this.idTextArea = idTextArea;
    var oTextArea = document.getElementById(idTextArea);
    oTextArea.style.display = "none";
    var oForm = oTextArea.form;
    if (oForm) {
        if (!bOnSubmitOriginalSaved) {
            onsubmit_original = oForm.onsubmit;
            bOnSubmitOriginalSaved = true
        }
        oForm.onsubmit = new Function("return onsubmit_new()")
    }
    var sContent = document.getElementById(idTextArea).value;
    sContent = sContent.replace(/&/g, "&amp;");
    sContent = sContent.replace(/</g, "&lt;");
    sContent = sContent.replace(/>/g, "&gt;");
    this.RENDER(sContent, dvId)
};

function isContentChanged() {
    var sContent = "";
    if (this.mode == "HTMLBody") {
        sContent = this.getHTMLBody()
    } else if (this.mode == "HTML") {
        sContent = this.getHTML()
    } else if (this.mode == "XHTMLBody") {
        sContent = this.getXHTMLBody()
    } else if (this.mode == "XHTML") {
        sContent = this.getXHTML()
    }
    if (sContent != this.originalContent) {
        return true
    }
    return false
};

function onsubmit_new() {
    var sContent;
    for (var i = 0; i < oUtil.arrEditor.length; i++) {
        if (!document.getElementById("idContent" + oUtil.arrEditor[i])) continue;
        var oEdit = eval(oUtil.arrEditor[i]);
        var oEditor = eval("idContent" + oEdit.oName);
        var allSpans = oEditor.document.getElementsByTagName("SPAN");
        for (var j = 0; j < allSpans.length; j++) {
            if ((allSpans[j].innerHTML == "") && (allSpans[j].parentElement.children.length == 1)) {
                allSpans[j].innerHTML = "&nbsp;"
            }
        }
        if (oEdit.mode == "HTMLBody") sContent = oEdit.getHTMLBody();
        if (oEdit.mode == "HTML") sContent = oEdit.getHTML();
        if (oEdit.mode == "XHTMLBody") sContent = oEdit.getXHTMLBody();
        if (oEdit.mode == "XHTML") sContent = oEdit.getXHTML();
        document.getElementById(oEdit.idTextArea).value = sContent
    }
    if (onsubmit_original) return onsubmit_original()
};

function onsubmit_original() {};
var iconHeight;

function RENDER(sPreloadHTML, dvId) {
    if (document.compatMode && document.compatMode != "BackCompat") {
        if (String(this.height).indexOf("%") == -1) {
            var eh = parseInt(this.height, 10);
            eh += this.heightAdjustment;
            this.height = eh + "px"
        }
    }
    iconHeight = this.iconHeight;
    if (sPreloadHTML.substring(0, 4) == "<!--" && sPreloadHTML.substring(sPreloadHTML.length - 3) == "-->") sPreloadHTML = sPreloadHTML.substring(4, sPreloadHTML.length - 3);
    if (sPreloadHTML.substring(0, 4) == "<!--" && sPreloadHTML.substring(sPreloadHTML.length - 6) == "--&gt;") sPreloadHTML = sPreloadHTML.substring(4, sPreloadHTML.length - 6);
    sPreloadHTML = sPreloadHTML.replace(/&lt;/g, "<");
    sPreloadHTML = sPreloadHTML.replace(/&gt;/g, ">");
    sPreloadHTML = sPreloadHTML.replace(/&amp;/g, "&");
    this.preloadHTML = sPreloadHTML;
    var sHTMLDropMenus = "";
    var sHTMLIcons = "";
    var sTmp = "";
    this.oColor1.onShow = new Function(this.oName + ".hide()");
    this.oColor1.onPickColor = new Function(this.oName + ".applyColor('ForeColor',eval('" + this.oName + "').oColor1.color)");
    this.oColor1.onRemoveColor = new Function(this.oName + ".applyColor('ForeColor','')");
    this.oColor2.onShow = new Function(this.oName + ".hide()");
    this.oColor2.onPickColor = new Function(this.oName + ".applyColor('BackColor',eval('" + this.oName + "').oColor2.color)");
    this.oColor2.onRemoveColor = new Function(this.oName + ".applyColor('BackColor','')");
    var me = this;
    var tmp = null,
        tmpTb, grpMap = new Object();
    for (var i = 0; i < this.groups.length; i++) {
        tmp = this.groups[i];
        tmpTb = this.tbar.createToolbar(this.oName + "tbar" + tmp[0]);
        tmpTb.onClick = function (id) {
            tbAction(tmpTb, id, me, me.oName)
        };
        tmpTb.style.toolbar = "main_istoolbar";
        tmpTb.iconPath = this.scriptPath + this.iconPath;
        tmpTb.btnWidth = this.iconWidth;
        tmpTb.btnHeight = this.iconHeight;
        for (var j = 0; j < tmp[2].length; j++) {
            eval(this.oName + ".btn" + tmp[2][j] + "=true")
        }
        buildToolbar(tmpTb, this, tmp[2]);
        grpMap[tmp[0]] = tmp[1]
    }
    if (this.toolbarMode == 1) {
        var eTab = this.tbar.createTbTab("tabCtl" + this.oName),
            tmpGrp;
        for (var i = 0; i < this.tabs.length; i++) {
            tmp = this.tabs[i];
            tmpGrp = this.tbar.createTbGroup(this.oName + "grp" + tmp[0]);
            for (var j = 0; j < tmp[2].length; j++) {
                tmpGrp.addGroup(this.oName + tmp[2][j], grpMap[tmp[2][j]], this.oName + "tbar" + tmp[2][j])
            }
            eTab.addTab(this.oName + tmp[0], tmp[1], tmpGrp)
        }
    } else if (this.groups.length > 0) {
        var tmpGrp;
        tmpGrp = this.tbar.createTbGroup(this.oName + "grp");
        for (var i = 0; i < this.groups.length; i++) {
            tmp = this.groups[i];
            tmpGrp.addGroup(this.oName + tmp[0], grpMap[tmp[0]], this.oName + "tbar" + tmp[0])
        }
        if (this.toolbarMode == 3) {
            tmpGrp.groupFlow = true
        }
        if (this.toolbarMode == 4) {
            tmpGrp.groupFlow = true;
            tmpGrp.draggable = true
        }
    }
    var sHTML = "";
    sHTML += "<iframe name=idFixZIndex" + this.oName + " id=idFixZIndex" + this.oName + "  frameBorder=0 style='display:none;position:fixed;filter:progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=0)' src='" + this.scriptPath + "blank.gif' ></iframe>";
    sHTML += "<table id=idArea" + this.oName + " name=idArea" + this.oName + " border=0 " + "cellpadding=0 cellspacing=0 width='" + this.width + "' height='" + this.height + "' style='width:" + this.width + ";height:" + this.height + ";border:none;border-bottom:#cfcfcf 1px solid'>" + "<tr><td colspan=2 style=\"position:relative;padding:0px;padding-left:0px;border:#cfcfcf 0px solid;border-bottom:0;background:url('" + this.scriptPath + "icons/bg.gif')\">" + "<table cellpadding=0 cellspacing=0 width='100%' id='idToolbar" + this.oName + "'  style='border:none;margin:0px'><tr><td dir=ltr style='padding:0px'>" + this.tbar.render() + "</td></tr></table>" + "</td></tr>" + "<tr id=idTagSelTopRow" + this.oName + "><td colspan=2 id=idTagSelTop" + this.oName + " height=0 style='padding:0px;color:#000000'></td></tr>";
    sHTML += "<tr style='width:100%;height:100%'><td colspan=2 valign=top height=100% style='padding:0px;background:white;'>";
    sHTML += "<table cellpadding=0 cellspacing=0 width='100%' height=100% style='margin:0px;border:none;'><tr><td width=100% height=100% style='border:none;padding:0px;border:#cfcfcf 1px solid;border-bottom:none;'>";
    if (this.IsSecurityRestricted) sHTML += "<iframe security='restricted' style='width:100%;height:100%;' frameborder='no' src='" + this.scriptPath + "blank.gif'" + " name=idContent" + this.oName + " id=idContent" + this.oName + " contentEditable=true></iframe>";
    else sHTML += "<iframe style='width:100%;height:100%;' frameborder='no' src='" + this.scriptPath + "blank.gif'" + " name=idContent" + this.oName + " id=idContent" + this.oName + " contentEditable=true></iframe>";
    sHTML += "<iframe style='position:fixed;left:-1000px;top:-1000px;width:1px;height:1px;overflow:auto;' src='" + this.scriptPath + "blank.gif'" + " name=idContentWord" + this.oName + " id=idContentWord" + this.oName + " contentEditable=true onfocus='" + this.oName + ".hide()'></iframe>";
    sHTML += "</td><td id=idStyles" + this.oName + " style='padding:0px;background:#f4f4f4' valign=top></td></tr></table>";
    sHTML += "</td></tr>";
    sHTML += "<tr id=idTagSelBottomRow" + this.oName + "><td colspan=2 id=idTagSelBottom" + this.oName + " style='width:100%;padding:0px;color:#000000'></td></tr>";
    if (this.showResizeBar) {
        sHTML += "<tr id=idResizeBar" + this.oName + "><td colspan=2 style='padding:0px;'><div class='resize_bar' style='cursor:n-resize;' onmousedown=\"onEditorStartResize(event, this, '" + this.oName + "')\" ></div></td></tr>"
    }
    sHTML += "</table>";
    sHTML += sHTMLDropMenus;
    sHTML += "<input type=submit name=iwe_btnSubmit" + this.oName + " id=iwe_btnSubmit" + this.oName + " value=SUBMIT style='display:none' >";
    if (dvId) {
        var edtStr = [];
        edtStr[0] = sHTML;
        document.getElementById(dvId).innerHTML = edtStr.join("")
    } else {
        document.write(sHTML)
    }
    this.init()
};

function onEditorStartResize(ev, elm, oName) {
    document.onmousemove = onEditorResize;
    document.onmouseup = onEditorStopResize;
    document.onselectstart = function () {
        return false
    };
    document.ondragstart = function () {
        return false
    };
    document.body.style.cursor = "n-resize";
    oUtil.currentResized = eval(oName);
    oUtil.resizeInit = {
        x: ev.screenX,
        y: ev.screenY
    };
    if (!oUtil.isWindow) oUtil.isWindow = new ISWindow(oName);
    oUtil.isWindow.showOverlay()
};

function onEditorStopResize() {
    oUtil.resizeOffset = {
        dx: event.screenX - oUtil.resizeInit.x,
        dy: event.screenY - oUtil.resizeInit.y
    };
    oUtil.currentResized.changeHeight(oUtil.resizeOffset.dy);
    oUtil.isWindow.hideOverlay();
    document.onmousemove = null;
    document.onmouseup = null;
    document.body.style.cursor = "default"
};

function onEditorResize() {
    oUtil.resizeOffset = {
        dx: event.screenX - oUtil.resizeInit.x,
        dy: event.screenY - oUtil.resizeInit.y
    };
    oUtil.currentResized.changeHeight(oUtil.resizeOffset.dy);
    oUtil.resizeInit = {
        x: event.screenX,
        y: event.screenY
    }
};

function initISEditor() {
    if (this.useTagSelector) {
        if (this.TagSelectorPosition == "bottom") this.TagSelectorPosition = "top";
        else this.TagSelectorPosition = "bottom";
        this.moveTagSelector()
    }
    oUtil.oName = this.oName;
    oUtil.oEditor = eval("idContent" + this.oName);
    oUtil.obj = eval(this.oName);
    oUtil.arrEditor.push(this.oName);
    eval("idArea" + this.oName).style.position = "fixed";
    window.setTimeout("eval('idArea" + this.oName + "').style.position='';", 1);
    var arrA = String(this.preloadHTML).match(/<HTML[^>]*>/ig);
    if (arrA) {
        this.loadHTML("");
        window.setTimeout(this.oName + ".putHTML(" + this.oName + ".preloadHTML)", 0)
    } else {
        this.loadHTML(this.preloadHTML)
    }
    switch (this.mode) {
    case "HTMLBody":
        this.originalContent = this.getHTMLBody();
        break;
    case "HTML":
        this.originalContent = this.getHTML();
        break;
    case "XHTMLBody":
        this.originalContent = this.getXHTMLBody();
        break;
    case "XHTML":
        this.originalContent = this.getXHTML();
        break
    };
    if (this.returnKeyMode == -1) {
        if (this.useDIV) this.returnKeyMode = 1;
        else if (this.useBR) this.returnKeyMode = 2;
        else if (!this.useDIV && !this.useBR) this.returnKeyMode = 3
    }
    this.onRender();
    if (this.disableFocusOnLoad == false) {
        this.focus()
    }
};

function buildToolbar(tb, oEdt, btnMap) {
    var oName = oEdt.oName;
    for (var i = 0; i < btnMap.length; i++) {
        sButtonName = btnMap[i];
        switch (sButtonName) {
        case "|":
            tb.addSeparator();
            break;
        case "BRK":
            tb.addBreak();
            break;
        case "FullScreen":
            tb.addButton("btnFullScreen" + oName, "btnFullScreen.gif", getTxt("Full Screen"));
            break;
        case "Print":
            tb.addButton("btnPrint" + oName, "btnPrint.gif", getTxt("Print"));
            break;
        case "Search":
            tb.addButton("btnSearch" + oName, "btnSearch.gif", getTxt("Search"));
            break;
        case "SpellCheck":
            tb.addButton("btnSpellCheck" + oName, "btnSpellCheck.gif", getTxt("Check Spelling"));
            break;
        case "Styles":
            tb.addButton("btnStyles" + oName, "btnStyleSelect.gif", getTxt("Style Selection"));
            break;
        case "Paragraph":
            tb.addDropdownButton("btnParagraph" + oName, "ddParagraph" + oName, "btnParagraph.gif", getTxt("Paragraph"), 37);
            var ddPar = new ISDropdown("ddParagraph" + oName);
            ddPar.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrParagraph.length; j++) {
                ddPar.addItem("btnParagraph_" + j + oName, "<" + oEdt.arrParagraph[j][1] + " style=\"\margin-bottom:4px\"  unselectable=on> " + oEdt.arrParagraph[j][0] + "</" + oEdt.arrParagraph[j][1] + ">")
            }
            break;
        case "FontName":
            tb.addDropdownButton("btnFontName" + oName, "ddFontName" + oName, "btnFontName.gif", getTxt("Font Name"), 37);
            var ddFont = new ISDropdown("ddFontName" + oName);
            ddFont.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrFontName.length; j++) {
                ddFont.addItem("btnFontName_" + j + oName, "<span style='font-family:" + oEdt.arrFontName[j] + ";font-size:12px' unselectable=on>" + oEdt.arrFontName[j] + "</span>")
            }
            break;
        case "FontSize":
            tb.addDropdownButton("btnFontSize" + oName, "ddFontSize" + oName, "btnFontSize.gif", getTxt("Font Size"), 37);
            var ddFs = new ISDropdown("ddFontSize" + oName);
            ddFs.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrFontSize.length; j++) {
                ddFs.addItem("btnFontSize_" + j + oName, "<div unselectable=on style=\"font-size:" + oEdt.arrFontSize[j][1] + "\">" + oEdt.arrFontSize[j][0] + "</div>")
            }
            break;
        case "Cut":
            tb.addButton("btnCut" + oName, "btnCut.gif", getTxt("Cut"));
            break;
        case "Copy":
            tb.addButton("btnCopy" + oName, "btnCopy.gif", getTxt("Copy"));
            break;
        case "Paste":
            tb.addDropdownButton("btnPaste" + oName, "ddPaste" + oName, "btnPaste.gif", getTxt("Paste"));
            var pvDD = new ISDropdown("ddPaste" + oName);
            pvDD.iconPath = tb.iconPath;
            pvDD.addItem("btnPasteClip" + oName, getTxt("Paste"), "btnPasteClip.gif");
            pvDD.addItem("btnPasteText" + oName, getTxt("Paste Text"), "btnPasteText.gif");
            pvDD.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            break;
        case "Emoticons":
            tb.addDropdownButton("btnEmoticons" + oName, "ddEmoticons" + oName, "btnEmoticons.gif", getTxt("Emoticons"));
            var ddTable = new ISDropdown("ddEmoticons" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertEmoticons", oEdt.GetEmoticons()));
            break;
        case "Quote":
            tb.addToggleButton("btnQuote" + oName, "", false, "btnQuote.gif", getTxt("Quote"));
            break;
        case "Undo":
            tb.addButton("btnUndo" + oName, "btnUndo.gif", getTxt("Undo"));
            break;
        case "Redo":
            tb.addButton("btnRedo" + oName, "btnRedo.gif", getTxt("Redo"));
            break;
        case "Bold":
            tb.addToggleButton("btnBold" + oName, "", false, "btnBold.gif", getTxt("Bold"));
            break;
        case "Italic":
            tb.addToggleButton("btnItalic" + oName, "", false, "btnItalic.gif", getTxt("Italic"));
            break;
        case "Underline":
            tb.addToggleButton("btnUnderline" + oName, "", false, "btnUnderline.gif", getTxt("Underline"));
            break;
        case "Strikethrough":
            tb.addToggleButton("btnStrikethrough" + oName, "", false, "btnStrikethrough.gif", getTxt("Strikethrough"));
            break;
        case "Superscript":
            tb.addToggleButton("btnSuperscript" + oName, "", false, "btnSuperscript.gif", getTxt("Superscript"));
            break;
        case "Subscript":
            tb.addToggleButton("btnSubscript" + oName, "", false, "btnSubscript.gif", getTxt("Subscript"));
            break;
        case "JustifyLeft":
            tb.addToggleButton("btnJustifyLeft" + oName, "align", false, "btnLeft.gif", getTxt("Justify Left"));
            break;
        case "JustifyCenter":
            tb.addToggleButton("btnJustifyCenter" + oName, "align", false, "btnCenter.gif", getTxt("Justify Center"));
            break;
        case "JustifyRight":
            tb.addToggleButton("btnJustifyRight" + oName, "align", false, "btnRight.gif", getTxt("Justify Right"));
            break;
        case "JustifyFull":
            tb.addToggleButton("btnJustifyFull" + oName, "align", false, "btnFull.gif", getTxt("Justify Full"));
            break;
        case "Numbering":
            tb.addToggleButton("btnNumbering" + oName, "bullet", false, "btnNumber.gif", getTxt("Numbering"));
            break;
        case "Bullets":
            tb.addToggleButton("btnBullets" + oName, "bullet", false, "btnList.gif", getTxt("Bullets"));
            break;
        case "Indent":
            tb.addButton("btnIndent" + oName, "btnIndent.gif", getTxt("Indent"));
            break;
        case "Outdent":
            tb.addButton("btnOutdent" + oName, "btnOutdent.gif", getTxt("Outdent"));
            break;
        case "LTR":
            tb.addToggleButton("btnLTR" + oName, "dir", false, "btnLTR.gif", getTxt("Left To Right"));
            break;
        case "RTL":
            tb.addToggleButton("btnRTL" + oName, "dir", false, "btnRTL.gif", getTxt("Right To Left"));
            break;
        case "ForeColor":
            tb.addDropdownButton("btnForeColor" + oName, "ddForeColor" + oName, "btnForeColor.gif", getTxt("Foreground Color"));
            var ddTable = new ISDropdown("ddForeColor" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertForeColor", oEdt.oColor1.generateHTML()));
            break;
        case "BackColor":
            tb.addDropdownButton("btnBackColor" + oName, "ddBackColor" + oName, "btnBackColor.gif", getTxt("Background Color"));
            var ddTable = new ISDropdown("ddBackColor" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertBackColor", oEdt.oColor2.generateHTML()));
            break;
        case "FontDialog":
            tb.addButton("btnFontDialog" + oName, "btnFont.gif", getTxt("Fonts"));
            break;
        case "TextDialog":
            tb.addButton("btnTextDialog" + oName, "btnText.gif", getTxt("Text"));
            break;
        case "CompleteTextDialog":
            tb.addButton("btnCompleteTextDialog" + oName, "btnText.gif", getTxt("Text"));
            break;
        case "LinkDialog":
            tb.addButton("btnLinkDialog" + oName, "btnHyperlink.gif", getTxt("Link"));
            break;
        case "ImageDialog":
            tb.addButton("btnImageDialog" + oName, "btnImage.gif", getTxt("Image"));
            break;
        case "YoutubeDialog":
            //tb.addButton("btnYoutubeDialog" + oName, "btnYoutubeVideo.gif", getTxt("YoutubeVideo"));
            break;
        case "TableDialog":
            tb.addButton("btnTableDialog" + oName, "btnTableEdit.gif", getTxt("Table"));
            break;
        case "FlashDialog":
            tb.addButton("btnFlashDialog" + oName, "btnFlash.gif", getTxt("Flash"));
            break;
        case "CharsDialog":
            tb.addButton("btnCharsDialog" + oName, "btnSymbol.gif", getTxt("Special Characters"));
            break;
        case "SearchDialog":
            tb.addButton("btnSearchDialog" + oName, "btnSearch.gif", getTxt("Search & Replace"));
            break;
        case "SourceDialog":
            tb.addButton("btnSourceDialog" + oName, "btnSource.gif", getTxt("HTML Editor"));
            break;
        case "BookmarkDialog":
            tb.addButton("btnBookmarkDialog" + oName, "btnBookmark.gif", getTxt("Bookmark"));
            break;
        case "Preview":
            tb.addButton("btnPreview" + oName, "btnPreview.gif", getTxt("Preview"));
            break;
        case "CustomTag":
            tb.addDropdownButton("btnCustomTag" + oName, "ddCustomTag" + oName, "btnCustomTag.gif", getTxt("Tags"), 37);
            var ddCustomTag = new ISDropdown("ddCustomTag" + oName);
            ddCustomTag.onClick = function (id) {
                ddAction(tb, id, oEdt, oEdt.oName)
            };
            for (var j = 0; j < oEdt.arrCustomTag.length; j++) {
                ddCustomTag.addItem("btnCustomTag_" + j + oName, oEdt.arrCustomTag[j][0])
            }
            break;
        case "ContentBlock":
            if (oEdt.btnContentBlock) tb.addButton("btnContentBlock" + oName, "btnContentBlock.gif", getTxt("Content Block"));
            break;
        case "InternalLink":
            tb.addButton("btnInternalLink" + oName, "btnInternalLink.gif", getTxt("Internal Link"));
            break;
        case "InternalImage":
            tb.addButton("btnInternalImage" + oName, "btnInternalImage.gif", getTxt("Internal Image"));
            break;
        case "CustomObject":
            tb.addButton("btnCustomObject" + oName, "btnCustomObject.gif", getTxt("Object"));
            break;
        case "Table":
            var sdd = [],
                sZ = 0;
            sdd[sZ++] = "<table width=195 id=dropTableCreate" + oName + " onmouseout='doOut_TabCreate();event.cancelBubble=true' style='cursor:default;background:#f3f3f8;border:#8a867a 0px solid;' cellpadding=0 cellspacing=2 border=0 unselectable=on>";
            for (var m = 0; m < 8; m++) {
                sdd[sZ++] = "<tr>";
                for (var n = 0; n < 8; n++) {
                    sdd[sZ++] = "<td onclick='" + oName + ".doClick_TabCreate()' onmouseover='doOver_TabCreate()' style='background:#ffffff;font-size:1px;border:#8a867a 1px solid;width:20px;height:20px;' unselectable=on>&nbsp;</td>"
                }
                sdd[sZ++] = "</tr>"
            }
            sdd[sZ++] = "<tr><td colspan=8 onclick=\"" + oName + ".hide();modelessDialogShow('" + oEdt.dialogPath + "webtable.htm',785, 500);\" onmouseover=\"this.innerText='" + getTxt("Advanced Table Insert") + "';this.style.border='#777777 1px solid';this.style.backgroundColor='#444444';this.style.color='#ffffff'\" onmouseout=\"this.style.border='#f3f3f8 1px solid';this.style.backgroundColor='#f3f3f8';this.style.color='#000000'\" align=center style='font-family:verdana;font-size:10px;font-color:black;border:#f3f3f8 1px solid;' unselectable=on>" + getTxt("Advanced Table Insert") + "</td></tr>";
            sdd[sZ++] = "</table>";
            tb.addDropdownButton("btnTable" + oName, "ddTable" + oName, "btnTable.gif", getTxt("Insert Table"));
            var ddTable = new ISDropdown("ddTable" + oName);
            ddTable.add(new ISCustomDDItem("btnInsertTable", sdd.join("")));
            break;
        case "Guidelines":
            tb.addButton("btnGuidelines" + oName, "btnGuideline.gif", getTxt("Show/Hide Guidelines"));
            break;
        case "Absolute":
            tb.addButton("btnAbsolute" + oName, "btnAbsolute.gif", getTxt("Absolute"));
            break;
        case "Line":
            tb.addButton("btnLine" + oName, "btnLine.gif", getTxt("Line"));
            break;
        case "RemoveFormat":
            tb.addButton("btnRemoveFormat" + oName, "btnRemoveFormat.gif", getTxt("Remove Formatting"));
            break;
        case "ClearAll":
            tb.addButton("btnClearAll" + oName, "btnDelete.gif", getTxt("Clear All"));
            break;
        default:
            for (var j = 0; j < oEdt.arrCustomButtons.length; j++) {
                if (sButtonName == oEdt.arrCustomButtons[j][0]) {
                    sCbName = oEdt.arrCustomButtons[j][0];
                    sCbCaption = oEdt.arrCustomButtons[j][2];
                    sCbImage = oEdt.arrCustomButtons[j][3];
                    if (oEdt.arrCustomButtons[j].length < 5) {
                        if (oEdt.arrCustomButtons[j][4]) tb.addButton(sCbName + oName, sCbImage, sCbCaption, oEdt.arrCustomButtons[j][4]);
                        else tb.addButton(sCbName + oName, sCbImage, sCbCaption)
                    } else {
                        if (oEdt.arrCustomButtons[j][4] != 0) tb.addDropdownButton(sCbName + oName, "dd" + sCbName + oName, sCbImage, sCbCaption, oEdt.arrCustomButtons[j][4]);
                        else tb.addDropdownButton(sCbName + oName, "dd" + sCbName + oName, sCbImage, sCbCaption);
                        var ddTable = new ISDropdown("dd" + sCbName + oName);
                        var arrItems = oEdt.arrCustomButtons[j][5];
                        for (var k = 0; k < arrItems.length; k++) {
                            tmp = arrItems[k];
                            ddTable.addItem('item_' + sCbName + k + oName, '<div style="margin:4px 0;padding:2px 0;font-size:13px;"  unselectable=on onclick="' + tmp[1] + '" > ' + tmp[2] + '</div>')
                        }
                    }
                }
            }
            break
        }
    }
};

function iwe_getElm(s) {
    return document.getElementById(s + this.oName)
};
var arrColorPickerObjects = [];

function ColorPicker(sName, sParent) {
    this.oParent = sParent;
    if (sParent) {
        this.oName = sParent + "." + sName;
        this.oRenderName = sName + sParent
    } else {
        this.oName = sName;
        this.oRenderName = sName
    }
    arrColorPickerObjects.push(this.oName);
    this.onShow = function () {
        return true
    };
    this.onHide = function () {
        return true
    };
    this.onPickColor = function () {
        return true
    };
    this.onRemoveColor = function () {
        return true
    };
    this.hide = hideColorPicker;
    this.hideAll = hideColorPickerAll;
    this.color;
    this.isActive = false;
    this.generateHTML = generateHTML
};

function generateHTML() {
    var arrColors = [
        ["#ef001b", "#cc0017", "#a60012", "#83000e", "#ef0078", "#ce0067", "#ad0057", "#8b0045", "#e301ed", "#c501ce", "#a401ab", "#88018e", "#6716ef", "#5913ce", "#4b10af", "#3e0d90"],
        ["#f67684", "#e36875", "#ca5965", "#b34e59", "#f563ac", "#de599b", "#cc5490", "#b24d7f", "#ee68f4", "#db5fe1", "#c759cc", "#b255b6", "#a779f5", "#976cdf", "#8d68cc", "#7f5eb7"],
        ["#fcc0c6", "#eea8af", "#dd959c", "#ce8c93", "#fec7e2", "#f4b8d6", "#e5a6c6", "#d495b4", "#fabffd", "#eeaff1", "#e19fe4", "#cf90d2", "#e0c3fd", "#d1b1f1", "#c1a0e2", "#b192d1"],
        ["#fef5f6", "#fdeced", "#f7dee0", "#eacedc", "#fef3f8", "#fbe8f1", "#efd0e0", "#e6c7d6", "#fef2fe", "#fae6fb", "#f1d3f2", "#e3c1e4", "#f5edfe", "#f0e5fb", "#e1d3ef", "#d9cbe7"],
        ["#028b6c", "#02775d", "#02644e", "#015441", "#1882ed", "#1574d4", "#115eab", "#0e4f90", "#0040eb", "#0039d0", "#0030b1", "#002892", "#50509e", "#46468b", "#3a3a73", "#303060"],
        ["#69baa7", "#61a898", "#57998a", "#508b7d", "#7bb8f5", "#6ea7e0", "#6195c9", "#5684b2", "#6d92f5", "#5f82e0", "#5675c9", "#4d68b2", "#9b9bc9", "#8b8bb6", "#7e7ea5", "#747496"],
        ["#d0eae4", "#b3d7cf", "#9bc4ba", "#8fb4ac", "#c3dffc", "#aacdf0", "#9bbde0", "#97b4d1", "#bdcdfb", "#a8bbef", "#96aae1", "#8a9bcb", "#d8d8eb", "#c7c7dc", "#b5b5cc", "#a5a5bc"],
        ["#f0f8f6", "#deedea", "#d7e6e2", "#ceddda", "#f1f7fe", "#e5f0fb", "#d8e5f2", "#cfdbe7", "#eff3fe", "#e5eafa", "#dde3f4", "#d2d8ea", "#f4f4f9", "#e5e5ef", "#dbdbe5", "#d6d6df"],
        ["#00a000", "#008d00", "#007700", "#006000", "#86d800", "#73ba00", "#629e00", "#528400", "#eded00", "#cece00", "#afaf00", "#909000", "#e3ab00", "#c79600", "#aa8000", "#856400"],
        ["#68c868", "#5cb65c", "#56a456", "#4b924b", "#b7e768", "#a8d45f", "#97c056", "#86aa4d", "#f1f164", "#e1e15d", "#caca58", "#b2b24d", "#eecc65", "#dabc5e", "#c7ac59", "#b09850"],
        ["#c6ecc6", "#addead", "#96cd96", "#87b987", "#e1f6c0", "#d0eba6", "#c1d99a", "#b1c88c", "#fbfbad", "#f1f194", "#e2e28e", "#cece8c", "#faeaba", "#f2dfa7", "#e6d090", "#cbbb8b"],
        ["#eef9ee", "#dff1df", "#d5e8d5", "#c6dbc6", "#f1fbe2", "#e9f5d5", "#dfebcd", "#d4e1c0", "#fefef0", "#fafae3", "#f0f0cb", "#e4e4c5", "#fdf8ea", "#f9f2de", "#eee4c7", "#dfd7bf"],
        ["#818181", "#676767", "#494949", "#000000", "#783c00", "#673300", "#562b00", "#472300", "#eb4600", "#cd3d00", "#ad3300", "#8f2a00", "#ed7700", "#d26900", "#af5800", "#904800"],
        ["#c9c9c9", "#a9a9a9", "#919191", "#787878", "#af8b68", "#a28264", "#917458", "#856d55", "#f19068", "#dd8561", "#c97654", "#b47053", "#f5ac63", "#e1a05f", "#ca9259", "#b78451"],
        ["#efefef", "#dcdcdc", "#c1c1c1", "#9d9d9d", "#dbcab9", "#ccb8a5", "#bda792", "#a3917f", "#fbcebc", "#f1bba5", "#e1aa93", "#ce9f8b", "#fcd7b3", "#f3caa2", "#e7b98c", "#c8a078"],
        ["#ffffff", "#f7f7f7", "#ededed", "#dddddd", "#f4efeb", "#efe8e1", "#e6ded6", "#dbd3cc", "#fef5f2", "#fae8e1", "#f0dbd3", "#e1cbc2", "#fef7f0", "#faecde", "#f1e2d3", "#e3d3c3"]
    ];
    var sHTMLColor = "<table id=dropColor" + this.oRenderName + " style=\"background-color:#fcfcfc;\" unselectable=on cellpadding=0 cellspacing=0 width=140px><tr><td unselectable=on style='padding:2px;'>";
    sHTMLColor += "<table align=center cellpadding=0 cellspacing=0 border=0 unselectable=on>";
    for (var i = 0; i < arrColors.length; i++) {
        sHTMLColor += "<tr>";
        for (var j = 0; j < arrColors[i].length; j++) sHTMLColor += "<td onclick=\"" + this.oName + ".color='" + arrColors[i][j] + "';" + this.oName + ".onPickColor();" + this.oName + ".hideAll()\" style=\"\" unselectable=on>" + "<table style='-moz-border-radius:0px;-webkit-border-radius: 0px;margin:0px;padding:0px;width:12px;height:12px;background:" + arrColors[i][j] + ";' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td unselectable=on style='cursor:pointer;margin:0px;padding:0px;line-height:normal;font-size:10px;'>&nbsp;</td></tr>" + "</table></td>";
        sHTMLColor += "</tr>"
    }
    sHTMLColor += "<tr>";
    sHTMLColor += "<td colspan=16 unselectable=on style='padding:0px;'>" + "<table style='padding:0px;margin-left:1px;width:100%;height:14px;background:#f4f4f4;' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td onclick=\"" + this.oName + ".onRemoveColor();" + this.oName + ".hideAll()\" onmouseover=\"this.style.border='#777777 1px solid'\" onmouseout=\"this.style.border='white 1px solid'\" style=\"cursor:default;padding:1px;border:white 1px solid;font-family:verdana;font-size:10px;color:#000000;line-height:9px;\" align=center valign=top unselectable=on>x</td></tr>" + "</table></td>";
    sHTMLColor += "</tr>";
    sHTMLColor += "</table>";
    sHTMLColor += "</td></tr></table>";
    return sHTMLColor
};

function hideColorPicker() {
    this.onHide();
    return;
    var box = eval("dropColor" + this.oRenderName);
    box.style.display = "none";
    this.isActive = false
};

function hideColorPickerAll() {
    return;
    for (var i = 0; i < arrColorPickerObjects.length; i++) {
        var box = eval("dropColor" + eval(arrColorPickerObjects[i]).oRenderName);
        box.style.display = "none";
        eval(arrColorPickerObjects[i]).isActive = false
    }
};

function loadHTML(sHTML) {
    var oEditor = eval("idContent" + this.oName);
    var sStyle = "";
    if (typeof this.css == 'object') {
        for (var j = 0; j < this.css.length; j++) {
            sStyle += "<link href='" + this.css[j] + "' rel='stylesheet' type='text/css'>"
        }
    } else {
        if (this.css != "") sStyle = "<link href='" + this.css + "' rel='stylesheet' type='text/css'>"
    }
    var oDoc = oEditor.document.open("text/html", "replace");
    if (this.publishingPath != "") {
        var arrA = String(this.preloadHTML).match(/<base[^>]*>/ig);
        if (!arrA) {
            sHTML = this.docType + "<html><head><base href=\"" + this.publishingPath + "\"/>" + this.headContent + sStyle + "</head><body contenteditable=true>" + sHTML + "</body></html>"
        }
    } else {
        sHTML = this.docType + "<html><head>" + this.headContent + sStyle + "</head><body contentEditable='true'>" + sHTML + "</body></html>"
    }
		// START ProductCart Mod: Fix YouTube and Vimeo URLs
		if (sHTML !== undefined && document.location.protocol == "https:") {
    	sHTML = sHTML.replace("http://www.youtube.com/", "https://www.youtube.com/");
    	sHTML = sHTML.replace("http://player.vimeo.com/", "https://player.vimeo.com/");
		}
		// END ProductCart Mod: Fix HTML content when using HTTPS.
    oDoc.write(sHTML);
    oDoc.close();
    oEditor.document.body.contentEditable = true;
    oEditor.document.execCommand("2D-Position", true, true);
    oEditor.document.execCommand("MultipleSelection", true, true);
    oEditor.document.execCommand("LiveResize", true, true);
    oEditor.document.body.onkeyup = new Function("editorDoc_onkeyup('" + this.oName + "')");
    oEditor.document.onmouseup = new Function("editorDoc_onmouseup('" + this.oName + "')");
    oEditor.document.body.onkeydown = new Function("doKeyPress(eval('idContent" + this.oName + "').event,'" + this.oName + "')");
    this.runtimeBorder(false);
    this.runtimeStyles();
    oEditor.document.body.onpaste = new Function("return " + this.oName + ".doOnPaste()");
    oEditor.document.body.oncut = new Function(this.oName + ".saveForUndo()");
    oEditor.document.body.style.lineHeight = "1.2";
    oEditor.document.body.style.lineHeight = "";
    if (this.initialRefresh) {
        oEditor.document.execCommand("SelectAll");
        window.setTimeout("eval('idContentWord" + this.oName + "').document.execCommand('SelectAll');", 0)
    }
    if (isIE9) {
        if (this.arrStyle.length > 0) {
            var oElement = oEditor.document.createElement("STYLE");
            oEditor.document.documentElement.childNodes[0].appendChild(oElement);
            var sCssText = "";
            for (var i = 0; i < this.arrStyle.length; i++) {
                selector = this.arrStyle[i][0];
                style = this.arrStyle[i][3];
                sCssText += selector + " { " + style + " } "
            }
            oElement.appendChild(oEditor.document.createTextNode(sCssText))
        }
    } else {
        if (this.arrStyle.length > 0) {
            var oElement = oEditor.document.createElement("STYLE");
            var n = oEditor.document.styleSheets.length;
            oEditor.document.documentElement.childNodes[0].appendChild(oElement);
            for (var i = 0; i < this.arrStyle.length; i++) {
                selector = this.arrStyle[i][0];
                style = this.arrStyle[i][3];
                oEditor.document.styleSheets[n].addRule(selector, style)
            }
        }
    }
    this.cleanDeprecated()
};

function putHTML(sHTML) {
    var oEditor = eval("idContent" + this.oName);
    var arrA = String(sHTML).match(/<!DOCTYPE[^>]*>/ig);
    if (arrA)
        for (var i = 0; i < arrA.length; i++) {
            this.docType = arrA[i]
        } else this.docType = "";
    var arrB = String(sHTML).match(/<HTML[^>]*>/ig);
    if (arrB)
        for (var i = 0; i < arrB.length; i++) {
            s = arrB[i];
            s = s.replace(/\"[^\"]*\"/ig, function (x) {
                x = x.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/'/g, "&apos;").replace(/\s+/ig, "#_#");
                return x
            });
            s = s.replace(/<([^ >]*)/ig, function (x) {
                return x.toLowerCase()
            });
            s = s.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
            s = s.replace(/ ([^=]+)=/ig, function (x) {
                return x.toLowerCase()
            });
            s = s.replace(/#_#/ig, " ");
            this.html = s
        } else this.html = "<html>";
    var oDoc = oEditor.document.open("text/html", "replace");
    oDoc.write(sHTML);
    oDoc.close();
    oEditor.document.body.contentEditable = true;
    oEditor.document.execCommand("2D-Position", true, true);
    oEditor.document.execCommand("MultipleSelection", true, true);
    oEditor.document.execCommand("LiveResize", true, true);
    oEditor.document.body.onkeyup = new Function("editorDoc_onkeyup('" + this.oName + "')");
    oEditor.document.onmouseup = new Function("editorDoc_onmouseup('" + this.oName + "')");
    oEditor.document.body.onkeydown = new Function("doKeyPress(eval('idContent" + this.oName + "').event,'" + this.oName + "')");
    this.runtimeBorder(false);
    this.runtimeStyles();
    this.cleanDeprecated()
};

function encodeHTMLCode(sHTML) {
    return sHTML.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
};

function getTextBody() {
    var oEditor = eval("idContent" + this.oName);
    return oEditor.document.body.innerText
};

function getHTML(bEdit) {
    var oEditor = eval("idContent" + this.oName);
    this.cleanDeprecated();
    sHTML = oEditor.document.documentElement.outerHTML;
    sHTML = String(sHTML).replace(/\<PARAM NAME=\"Play\" VALUE=\"0\">/ig, "<PARAM NAME=\"Play\" VALUE=\"-1\">");
    sHTML = this.docType + sHTML;
    sHTML = oUtil.replaceSpecialChar(sHTML);
    sHTML = sHTML.replace(/align=['"]*middle['"]*/gi, "align=\"center\"");
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    var text = this.getTextBody();
    if (!text || text.trim() == "") sHTML = "";
    return sHTML
};

function getHTMLBody(bEdit) {
    var oEditor = eval("idContent" + this.oName);
    this.cleanDeprecated();
    sHTML = oEditor.document.body.innerHTML;
    sHTML = String(sHTML).replace(/\<PARAM NAME=\"Play\" VALUE=\"0\">/ig, "<PARAM NAME=\"Play\" VALUE=\"-1\">");
    sHTML = oUtil.replaceSpecialChar(sHTML);
    sHTML = sHTML.replace(/align=['"]*middle['"]*/gi, "align=\"center\"");
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    var text = this.getTextBody();
    if (!text || text.trim() == "") sHTML = "";
    return sHTML
};
var sBaseHREF = "";

function getXHTML(bEdit) {
    var oEditor = eval("idContent" + this.oName);
    this.cleanDeprecated();
    sHTML = oEditor.document.documentElement.outerHTML;
    var arrTmp = sHTML.match(/<BASE([^>]*)>/ig);
    if (arrTmp != null) sBaseHREF = arrTmp[0];
    sBaseHREF = sBaseHREF.replace(/<([^ >]*)/ig, function (x) {
        return x.toLowerCase()
    });
    sBaseHREF = sBaseHREF.replace(/ [^=]+="[^"]+"/ig, function (x) {
        x = x.replace(/\s+/ig, "#_#");
        x = x.replace(/^#_#/, " ");
        return x
    });
    sBaseHREF = sBaseHREF.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
    sBaseHREF = sBaseHREF.replace(/ ([^=]+)=/ig, function (x) {
        return x.toLowerCase()
    });
    sBaseHREF = sBaseHREF.replace(/#_#/ig, " ");
    sBaseHREF = sBaseHREF.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
    sHTML = recur(oEditor.document.documentElement, "");
    sHTML = this.docType + this.html + sHTML + "\n</html>";
    sHTML = sHTML.replace(/<\/title>/, "<\/title>" + sBaseHREF);
    sHTML = oUtil.replaceSpecialChar(sHTML);
    sHTML = sHTML.replace(/align=['"]*middle['"]*/gi, "align=\"center\"");
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    var text = this.getTextBody();
    if (!text || text.trim() == "") sHTML = "";
    return sHTML
};

function getXHTMLBody(bEdit) {
    var oEditor = eval("idContent" + this.oName);
    this.cleanDeprecated();
    sHTML = oEditor.document.documentElement.outerHTML;
    var arrTmp = sHTML.match(/<BASE([^>]*)>/ig);
    if (arrTmp != null) sBaseHREF = arrTmp[0];
    sHTML = recur(oEditor.document.body, "");
    sHTML = oUtil.replaceSpecialChar(sHTML);
    sHTML = sHTML.replace(/align=['"]*middle['"]*/gi, "align=\"center\"");
    if (this.encodeIO) sHTML = encodeHTMLCode(sHTML);
    if (!bEdit) {
        sHTML = sHTML.replace(/ contenteditable=\"true\"/ig, "");
        sHTML = sHTML.replace(/ contenteditable=\"false\"/ig, "")
    }
    var text = this.getTextBody();
    if (!text || text.trim() == "") sHTML = "";
    return sHTML
};

function ApplyExternalStyle(oName) {
    var edtObj = eval(oName);
    var oEditor = eval("idContent" + oName);
    var sTmp = "";
    var prefixes = edtObj.styleSelectorPrefix.split(","),
        found = false;
    for (var j = 0; j < oEditor.document.styleSheets.length; j++) {
        var myStyle = oEditor.document.styleSheets[j];
        for (var i = 0; i < myStyle.rules.length; i++) {
            sSelector = myStyle.rules.item(i).selectorText;
            sCssText = myStyle.rules.item(i).style.cssText.replace(/"/g, "&quot;");
            if (sSelector.match(/table\./gi)) {
                continue
            }
            found = false;
            if (prefixes.length > 0) {
                for (var ix = 0; ix < prefixes.length; ix++) {
                    if (prefixes[ix] != "" && sSelector.indexOf("." + prefixes[ix]) >= 0) {
                        found = true;
                        break
                    }
                }
            }
            if (!found && edtObj.styleSelectorPrefix != "") continue;
            sSelector = sSelector.replace(/"/g, "\\\"");
            var itemCount = sSelector.split(".").length;
            if (itemCount > 1) {
                sCaption = sSelector.split(".")[1];
                sTmp += ",[\"" + sSelector + "\",true,\"" + sCaption + "\",\"" + sCssText + "\"]"
            } else sTmp += ",[\"" + sSelector + "\",false,\"\",\"" + sCssText + "\"]"
        }
    }
    var arrStyle = eval("[" + sTmp.substr(1) + "]");
    for (var i = 0; i < arrStyle.length; i++) {
        for (var j = 0; j < edtObj.arrStyle.length; j++) {
            if (arrStyle[i][0].toLowerCase() == edtObj.arrStyle[j][0].toLowerCase()) {
                arrStyle[i][1] = edtObj.arrStyle[j][1]
            }
        }
    }
    edtObj.arrStyle = arrStyle
};

function doApplyStyle(oName, sClassName) {
    if (!eval(oName).checkFocus()) return;
    var oEditor = eval("idContent" + oName);
    var oSel = oEditor.document.selection.createRange();
    eval(oName).saveForUndo();
    if (oUtil.activeElement) {
        oElement = oUtil.activeElement;
        oElement.className = sClassName
    } else if (oSel.parentElement) {
        if (oSel.text == "") {
            oElement = oSel.parentElement();
            if (oElement.tagName == "BODY") return;
            oElement.className = sClassName
        } else {
            eval(oName).applySpanStyle([], sClassName)
        }
    } else {
        oElement = oSel.item(0);
        oElement.className = sClassName
    }
    realTime(oName)
};

function openStyleSelect() {
    if (!this.isCssLoaded) ApplyExternalStyle(this.oName);
    this.isCssLoaded = true;
    var bShowStyles = false;
    var idStyles = document.getElementById("idStyles" + this.oName);
    if (idStyles.innerHTML != "") {
        if (idStyles.style.display == "") idStyles.style.display = "none";
        else idStyles.style.display = "";
        return
    }
    idStyles.style.display = "";
    var h = document.getElementById("idContent" + this.oName).offsetHeight - 27;
    var arrStyle = this.arrStyle;
    var sHTML = "";
    sHTML += "<div unselectable=on style='width:200px;margin:1px;margin-top:0;margin-right:2px;' align=right>";
    sHTML += "<table style='margin-right:1px;margin-bottom:3px;width:14px;height:14px;background:#f4f4f4;' cellpadding=0 cellspacing=0 unselectable=on>" + "<tr><td onclick=\"" + this.oName + ".openStyleSelect();\" onmouseover=\"this.style.border='#708090 1px solid';this.style.color='white';this.style.backgroundColor='9FA7BB'\" onmouseout=\"this.style.border='white 1px solid';this.style.color='black';this.style.backgroundColor=''\" style=\"cursor:default;padding:1px;border:white 1px solid;font-family:verdana;font-size:10px;font-color:black;line-height:9px;\" align=center valign=top unselectable=on>x</td></tr>" + "</table></div>";
    var sBody = "";
    for (var i = 0; i < arrStyle.length; i++) {
        sSelector = arrStyle[i][0];
        if (sSelector == "BODY") sBody = arrStyle[i][3]
    }
    sHTML += "<div unselectable=on style='overflow:auto;width:200px;height:" + h + "px;padding-left:3px;'>";
    sHTML += "<table name='tblStyles" + this.oName + "' id='tblStyles" + this.oName + "' cellpadding=0 cellspacing=0 style='background:#fcfcfc;" + sBody + ";width:100%;height:100%;margin:0;'>";
    for (var i = 0; i < arrStyle.length; i++) {
        sSelector = arrStyle[i][0];
        isOnSelection = arrStyle[i][1];
        sCssText = arrStyle[i][3];
        sCaption = arrStyle[i][2];
        if (isOnSelection) {
            if (sSelector.split(".").length > 1) {
                var tmpSelector = sSelector;
                if (sSelector.indexOf(":") > 0) tmpSelector = sSelector.substring(0, sSelector.indexOf(":"));
                bShowStyles = true;
                if (tmpSelector.indexOf("awesome") == -1) {
                    sHTML += "<tr style=\"cursor:default\" onmouseover=\"if(this.style.marginRight!='1px'){this.style.background='" + this.styleSelectionHoverBg + "';this.style.color='" + this.styleSelectionHoverFg + "'}\" onmouseout=\"if(this.style.marginRight!='1px'){this.style.background='';this.style.color=''}\">";
                    sHTML += "<td unselectable=on onclick=\"doApplyStyle('" + this.oName + "','" + tmpSelector.split(".")[1] + "')\" style='padding:2px;'>";
                    if (sSelector.split(".")[0] == "") sHTML += "<span unselectable=on style=\"" + sCssText + ";margin:0;\">" + sCaption + "</span>";
                    else sHTML += "<span unselectable=on style=\"" + sCssText + ";margin:0;\">" + sSelector + "</span>";
                    sHTML += "</td></tr>"
                }
            }
        }
    }
    sHTML += "<tr><td height=50%>&nbsp;</td></tr></table></div>";
    if (bShowStyles) document.getElementById("idStyles" + this.oName).innerHTML = sHTML;
    else {
        alert("No stylesheet found.")
    }
};

function editorDoc_onkeyup(oName) {
    if (eval(oName).isAfterPaste) {
        eval(oName).cleanDeprecated();
        eval(oName).runtimeBorder(false);
        eval(oName).runtimeStyles();
        eval(oName).isAfterPaste = false
    }
    var oEdt = eval(oName);
    if (oEdt.tmKeyup) {
        clearTimeout(oEdt.tmKeyup);
        oEdt.tmKeyup = null
    }
    if (!oEdt.tmKeyup) oEdt.tmKeyup = setTimeout(function () {
        realTime(oName)
    }, 1000);
    oEdt.bookmarkSelection()
};

function editorDoc_onmouseup(oName) {
    oUtil.activeElement = null;
    oUtil.oName = oName;
    oUtil.oEditor = eval("idContent" + oName);
    oUtil.obj = eval(oName);
    eval(oName).hide();
    if (oUtil.oEditor.document.selection.type != "Control") {
        oUtil.oEditor.focus()
    }
    realTime(oName);
    oUtil.obj.bookmarkSelection()
};

function setActiveEditor(oName) {
    oUtil.oName = oName;
    oUtil.oEditor = eval("idContent" + oName);
    oUtil.obj = eval(oName)
};
var arrTmp = [];

function GetElement(oElement, sMatchTag) {
    while (oElement != null && oElement.tagName != sMatchTag) {
        if (oElement.tagName == "BODY") return null;
        oElement = oElement.parentElement
    }
    return oElement
};
var arrTmp2 = [];

function realTime(oName, bTagSel) {
    if (!eval(oName).checkFocus()) return;
    var oEditor = eval("idContent" + oName);
    var oSel = oEditor.document.selection.createRange();
    var obj = eval(oName);
    var tbar = obj.tbar;
    var btn = null;
    var oTable = (oSel.parentElement != null ? GetElement(oSel.parentElement(), "TABLE") : GetElement(oSel.item(0), "TABLE"));
    if (oTable) obj.Table = oTable;
    var doc = oEditor.document;
    if (obj.btnParagraph) {
        btn = tbar.btns["btnParagraph" + oName];
        btn.setState(doc.queryCommandEnabled("FormatBlock") ? 1 : 5)
    }
    if (obj.btnFontName) {
        btn = tbar.btns["btnFontName" + oName];
        btn.setState(doc.queryCommandEnabled("FontName") ? 1 : 5)
    }
    if (obj.btnFontSize) {
        btn = tbar.btns["btnFontSize" + oName];
        btn.setState(doc.queryCommandEnabled("FontSize") ? 1 : 5)
    }
    if (obj.btnCut) {
        btn = tbar.btns["btnCut" + oName];
        btn.setState(doc.queryCommandEnabled("Cut") ? 1 : 5)
    }
    if (obj.btnCopy) {
        btn = tbar.btns["btnCopy" + oName];
        btn.setState(doc.queryCommandEnabled("Copy") ? 1 : 5)
    }
    if (obj.btnPaste) {
        btn = tbar.btns["btnPaste" + oName];
        btn.setState(doc.queryCommandEnabled("Paste") ? 1 : 5)
    }
    if (obj.btnUndo) {
        btn = tbar.btns["btnUndo" + oName];
        btn.setState(!obj.arrUndoList[0] ? 5 : 1)
    }
    if (obj.btnRedo) {
        btn = tbar.btns["btnRedo" + oName];
        btn.setState(!obj.arrRedoList[0] ? 5 : 1)
    }
    if (obj.btnBold) {
        btn = tbar.btns["btnBold" + oName];
        btn.setState(doc.queryCommandEnabled("Bold") ? (doc.queryCommandState("Bold") ? 4 : 1) : 5)
    }
    if (obj.btnItalic) {
        btn = tbar.btns["btnItalic" + oName];
        btn.setState(doc.queryCommandEnabled("Italic") ? (doc.queryCommandState("Italic") ? 4 : 1) : 5)
    }
    if (obj.btnUnderline) {
        btn = tbar.btns["btnUnderline" + oName];
        btn.setState(doc.queryCommandEnabled("Underline") ? (doc.queryCommandState("Underline") ? 4 : 1) : 5)
    }
    if (obj.btnStrikethrough) {
        btn = tbar.btns["btnStrikethrough" + oName];
        btn.setState(doc.queryCommandEnabled("Strikethrough") ? (doc.queryCommandState("Strikethrough") ? 4 : 1) : 5)
    }
    if (obj.btnSuperscript) {
        btn = tbar.btns["btnSuperscript" + oName];
        btn.setState(doc.queryCommandEnabled("Superscript") ? (doc.queryCommandState("Superscript") ? 4 : 1) : 5)
    }
    if (obj.btnSubscript) {
        btn = tbar.btns["btnSubscript" + oName];
        btn.setState(doc.queryCommandEnabled("Subscript") ? (doc.queryCommandState("Subscript") ? 4 : 1) : 5)
    }
    if (obj.btnNumbering) {
        btn = tbar.btns["btnNumbering" + oName];
        btn.setState(doc.queryCommandEnabled("InsertOrderedList") ? (doc.queryCommandState("InsertOrderedList") ? 4 : 1) : 5)
    }
    if (obj.btnBullets) {
        btn = tbar.btns["btnBullets" + oName];
        btn.setState(doc.queryCommandEnabled("InsertUnorderedList") ? (doc.queryCommandState("InsertUnorderedList") ? 4 : 1) : 5)
    }
    if (obj.btnJustifyLeft) {
        btn = tbar.btns["btnJustifyLeft" + oName];
        btn.setState(doc.queryCommandEnabled("JustifyLeft") ? (doc.queryCommandState("JustifyLeft") ? 4 : 1) : 5)
    }
    if (obj.btnJustifyCenter) {
        btn = tbar.btns["btnJustifyCenter" + oName];
        btn.setState(doc.queryCommandEnabled("JustifyCenter") ? (doc.queryCommandState("JustifyCenter") ? 4 : 1) : 5)
    }
    if (obj.btnJustifyRight) {
        btn = tbar.btns["btnJustifyRight" + oName];
        btn.setState(doc.queryCommandEnabled("JustifyRight") ? (doc.queryCommandState("JustifyRight") ? 4 : 1) : 5)
    }
    if (obj.btnJustifyFull) {
        btn = tbar.btns["btnJustifyFull" + oName];
        btn.setState(doc.queryCommandEnabled("JustifyFull") ? (doc.queryCommandState("JustifyFull") ? 4 : 1) : 5)
    }
    if (obj.btnIndent) {
        btn = tbar.btns["btnIndent" + oName];
        btn.setState(doc.queryCommandEnabled("Indent") ? 1 : 5)
    }
    if (obj.btnOutdent) {
        btn = tbar.btns["btnOutdent" + oName];
        btn.setState(doc.queryCommandEnabled("Outdent") ? 1 : 5)
    }
    if (obj.btnLTR) {
        btn = tbar.btns["btnLTR" + oName];
        btn.setState(doc.queryCommandEnabled("BlockDirLTR") ? (doc.queryCommandState("BlockDirLTR") ? 4 : 1) : 5)
    }
    if (obj.btnRTL) {
        btn = tbar.btns["btnRTL" + oName];
        btn.setState(doc.queryCommandEnabled("BlockDirRTL") ? (doc.queryCommandState("BlockDirRTL") ? 4 : 1) : 5)
    }
    var v = (oSel.parentElement ? 1 : 5);
    if (obj.btnForeColor) tbar.btns["btnForeColor" + oName].setState(v);
    if (obj.btnBackColor) tbar.btns["btnBackColor" + oName].setState(v);
    if (obj.btnLine) tbar.btns["btnLine" + oName].setState(v);
    try {
        oUtil.onSelectionChanged()
    } catch (e) {}
    var idStyles = document.getElementById("idStyles" + oName);
    if (idStyles.innerHTML != "") {
        var oElement;
        if (oUtil.activeElement) oElement = oUtil.activeElement;
        else {
            if (oSel.parentElement) oElement = oSel.parentElement();
            else oElement = oSel.item(0)
        }
        var sCurrClass = oElement.className;
        var oRows = document.getElementById("tblStyles" + oName).rows;
        for (var i = 0; i < oRows.length - 1; i++) {
            sClass = oRows[i].childNodes[0].innerText;
            if (sClass.split(".").length > 1 && sClass != "") sClass = sClass.split(".")[1];
            if (sCurrClass == sClass) {
                oRows[i].style.marginRight = "1px";
                oRows[i].style.backgroundColor = obj.styleSelectionHoverBg;
                oRows[i].style.color = obj.styleSelectionHoverFg
            } else {
                oRows[i].style.marginRight = "";
                oRows[i].style.backgroundColor = "";
                oRows[i].style.color = ""
            }
        }
    }
    if (obj.useTagSelector && !bTagSel) {
        if (oSel.parentElement) oElement = oSel.parentElement();
        else oElement = oSel.item(0);
        var sHTML = "";
        var i = 0;
        arrTmp2 = [];
        while (oElement != null && oElement.tagName != "BODY") {
            arrTmp2[i] = oElement;
            var sTagName = oElement.tagName;
            sHTML = "&nbsp; &lt;<span id=tag" + oName + i + " unselectable=on style='text-decoration:underline;cursor:pointer' onclick=\"" + oName + ".selectElement(" + i + ")\">" + sTagName + "</span>&gt;" + sHTML;
            oElement = oElement.parentElement;
            i++
        }
        sHTML = "&nbsp;&lt;BODY&gt;" + sHTML;
        eval("idElNavigate" + oName).innerHTML = sHTML;
        eval("idElCommand" + oName).style.display = "none"
    }
    if (obj.isAfterPaste) {
        obj.cleanDeprecated();
        obj.runtimeBorder(false);
        obj.runtimeStyles();
        obj.isAfterPaste = false
    }
    if (tbar.btns["btnQuote" + oName]) {
        var oQuote = (oSel.parentElement != null ? GetElement(oSel.parentElement(), "BLOCKQUOTE") : GetElement(oSel.item(0), "BLOCKQUOTE"));
        if (oQuote) tbar.btns["btnQuote" + oName].setState(4);
        else tbar.btns["btnQuote" + oName].setState(1)
    }
};

function realtimeFontSelect(oName) {
    var oEditor = eval("idContent" + oName);
    var sFontName = oEditor.document.queryCommandValue("FontName");
    var edt = eval(oName);
    var found = false;
    for (var i = 0; i < edt.arrFontName.length; i++) {
        if (sFontName == edt.arrFontName[i]) {
            found = true;
            break
        }
    }
    if (found) {
        isDDs["ddFontName" + oName].selectItem("btnFontName_" + i + oName, true)
    } else {
        isDDs["ddFontName" + oName].clearSelection()
    }
};

function realtimeSizeSelect(oName) {
    var oEditor = eval("idContent" + oName);
    var sFontSize = oEditor.document.queryCommandValue("FontSize");
    var edt = eval(oName);
    var found = false;
    for (var i = 0; i < edt.arrFontSize.length; i++) {
        if (sFontSize == edt.arrFontSize[i][1]) {
            found = true;
            break
        }
    }
    if (found) {
        isDDs["ddFontSize" + oName].selectItem("btnFontSize_" + i + oName, true)
    } else {
        isDDs["ddFontSize" + oName].clearSelection()
    }
};

function moveTagSelector() {
    var sTagSelTop = "<table unselectable=on ondblclick='" + this.oName + ".moveTagSelector()' width='100%' cellpadding=0 cellspacing=0 style='border:none;border:#cfcfcf 1px solid;border-bottom:none;margin:0px;'><tr style='background:#f8f8f8;font-family:arial;font-size:10px;color:black;'>" + "<td id=idElNavigate" + this.oName + " style='color:inherit;line-height:normal;font-size:inherit;padding:1px;width:100%;' valign=top>&nbsp;</td>" + "<td align=right valign='center' nowrap style='color:inherit;padding:0;margin;0;font-size:inherit;line-height:normal'>" + "<span id=idElCommand" + this.oName + " unselectable=on style='vertical-align:middle;display:none;text-decoration:underline;cursor:pointer;padding-right:5px;' onclick='" + this.oName + ".removeTag()'>" + getTxt("Remove Tag") + "</span>" + "</td></tr></table>";
    var sTagSelBottom = "<table unselectable=on ondblclick='" + this.oName + ".moveTagSelector()' width='100%' cellpadding=0 cellspacing=0 style='border:none;border-left:#cfcfcf 1px solid;border-right:#cfcfcf 1px solid;margin:0px;'><tr style='background-color:#f8f8f8;font-family:arial;font-size:10px;color:black;'>" + "<td id=idElNavigate" + this.oName + " style='color:inherit;line-height:normal;font-size:inherit;padding:1px;width:100%;' valign=top>&nbsp;</td>" + "<td align=right valign='center' nowrap style='color:inherit;padding:0;margin;0;font-size:inherit;line-height:normal'>" + "<span id=idElCommand" + this.oName + " unselectable=on style='vertical-align:middle;display:none;text-decoration:underline;cursor:pointer;padding-right:5px;' onclick='" + this.oName + ".removeTag()'>" + getTxt("Remove Tag") + "</span>" + "</td></tr></table>";
    if (this.TagSelectorPosition == "top") {
        eval("idTagSelTop" + this.oName).innerHTML = "";
        eval("idTagSelBottom" + this.oName).innerHTML = sTagSelBottom;
        eval("idTagSelTopRow" + this.oName).style.display = "none";
        eval("idTagSelBottomRow" + this.oName).style.display = "";
        this.TagSelectorPosition = "bottom"
    } else {
        eval("idTagSelTop" + this.oName).innerHTML = sTagSelTop;
        eval("idTagSelBottom" + this.oName).innerHTML = "";
        eval("idTagSelTopRow" + this.oName).style.display = "";
        eval("idTagSelBottomRow" + this.oName).style.display = "none";
        this.TagSelectorPosition = "top"
    }
};

function selectElement(i) {
    var oEditor = eval("idContent" + this.oName);
    var oSelRange = oEditor.document.body.createControlRange();
    var oActiveElement;
    try {
        oSelRange.add(arrTmp2[i]);
        oSelRange.select();
        realTime(this.oName, true);
        oActiveElement = arrTmp2[i];
        if (oActiveElement.tagName != "TD" && oActiveElement.tagName != "TR" && oActiveElement.tagName != "TBODY" && oActiveElement.tagName != "LI") eval("idElCommand" + this.oName).style.display = ""
    } catch (e) {
        try {
            var oSelRange = oEditor.document.body.createTextRange();
            oSelRange.moveToElementText(arrTmp2[i]);
            oSelRange.select();
            realTime(this.oName, true);
            oActiveElement = arrTmp2[i];
            if (oActiveElement.tagName != "TD" && oActiveElement.tagName != "TR" && oActiveElement.tagName != "TBODY" && oActiveElement.tagName != "LI") eval("idElCommand" + this.oName).style.display = ""
        } catch (e) {
            return
        }
    }
    for (var j = 0; j < arrTmp2.length; j++) eval("tag" + this.oName + j).style.background = "";
    eval("tag" + this.oName + i).style.background = "DarkGray";
    if (oActiveElement) oUtil.activeElement = oActiveElement
};

function removeTag() {
    if (!this.checkFocus()) return;
    eval(this.oName).saveForUndo();
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (sType == "Control") {
        oSel.item(0).outerHTML = "";
        this.focus();
        realTime(this.oName);
        return
    }
    var oActiveElement = oUtil.activeElement;
    var oSelRange = oEditor.document.body.createTextRange();
    oSelRange.moveToElementText(oActiveElement);
    oSel.setEndPoint("StartToStart", oSelRange);
    oSel.setEndPoint("EndToEnd", oSelRange);
    oSel.select();
    this.saveForUndo();
    sHTML = oActiveElement.innerHTML;
    sHTML = fixPathEncode(sHTML);
    var oTmp = oActiveElement.parentElement;
    if (oTmp.innerHTML == oActiveElement.outerHTML) {
        oTmp.innerHTML = sHTML;
        fixPathDecode(oEditor);
        var oSelRange = oEditor.document.body.createTextRange();
        oSelRange.moveToElementText(oTmp);
        oSel.setEndPoint("StartToStart", oSelRange);
        oSel.setEndPoint("EndToEnd", oSelRange);
        oSel.select();
        realTime(this.oName);
        this.selectElement(0);
        return
    } else {
        oActiveElement.outerHTML = "";
        oSel.pasteHTML(sHTML);
        fixPathDecode(oEditor);
        this.focus();
        realTime(this.oName)
    }
    this.runtimeBorder(false);
    this.runtimeStyles()
};

function runtimeBorderOn() {
    this.runtimeBorderOff();
    var oEditor = eval("idContent" + this.oName);
    var oTables = oEditor.document.getElementsByTagName("TABLE");
    for (i = 0; i < oTables.length; i++) {
        var oTable = oTables[i];
        if (oTable.border == 0) {
            var oCells = oTable.getElementsByTagName("TD");
            for (j = 0; j < oCells.length; j++) {
                if (oCells[j].style.borderLeftWidth == "0px" || oCells[j].style.borderLeftWidth == "" || oCells[j].style.borderLeftWidth == "medium") {
                    oCells[j].runtimeStyle.borderLeftWidth = "1px";
                    oCells[j].runtimeStyle.borderLeftColor = "#BCBCBC";
                    oCells[j].runtimeStyle.borderLeftStyle = "dotted"
                }
                if (oCells[j].style.borderRightWidth == "0px" || oCells[j].style.borderRightWidth == "" || oCells[j].style.borderRightWidth == "medium") {
                    oCells[j].runtimeStyle.borderRightWidth = "1px";
                    oCells[j].runtimeStyle.borderRightColor = "#BCBCBC";
                    oCells[j].runtimeStyle.borderRightStyle = "dotted"
                }
                if (oCells[j].style.borderTopWidth == "0px" || oCells[j].style.borderTopWidth == "" || oCells[j].style.borderTopWidth == "medium") {
                    oCells[j].runtimeStyle.borderTopWidth = "1px";
                    oCells[j].runtimeStyle.borderTopColor = "#BCBCBC";
                    oCells[j].runtimeStyle.borderTopStyle = "dotted"
                }
                if (oCells[j].style.borderBottomWidth == "0px" || oCells[j].style.borderBottomWidth == "" || oCells[j].style.borderBottomWidth == "medium") {
                    oCells[j].runtimeStyle.borderBottomWidth = "1px";
                    oCells[j].runtimeStyle.borderBottomColor = "#BCBCBC";
                    oCells[j].runtimeStyle.borderBottomStyle = "dotted"
                }
            }
        }
    }
};

function runtimeBorderOff() {
    var oEditor = eval("idContent" + this.oName);
    var oTables = oEditor.document.getElementsByTagName("TABLE");
    for (i = 0; i < oTables.length; i++) {
        var oTable = oTables[i];
        if (oTable.border == 0) {
            var oCells = oTable.getElementsByTagName("TD");
            for (j = 0; j < oCells.length; j++) {
                oCells[j].runtimeStyle.borderWidth = "";
                oCells[j].runtimeStyle.borderColor = "";
                oCells[j].runtimeStyle.borderStyle = ""
            }
        }
    }
};

function runtimeBorder(bToggle) {
    if (bToggle) {
        if (this.IsRuntimeBorderOn) {
            this.runtimeBorderOff();
            this.IsRuntimeBorderOn = false
        } else {
            this.runtimeBorderOn();
            this.IsRuntimeBorderOn = true
        }
    } else {
        if (this.IsRuntimeBorderOn) this.runtimeBorderOn();
        else this.runtimeBorderOff()
    }
};

function runtimeStyles() {
    var oEditor = eval("idContent" + this.oName);
    var oForms = oEditor.document.getElementsByTagName("FORM");
    for (i = 0; i < oForms.length; i++) oForms[i].runtimeStyle.border = "#7bd158 1px dotted";
    var oBookmarks = oEditor.document.getElementsByTagName("A");
    for (i = 0; i < oBookmarks.length; i++) {
        var oBookmark = oBookmarks[i];
        if (oBookmark.name || oBookmark.NAME) {
            if (oBookmark.innerHTML == "") oBookmark.runtimeStyle.width = "1px";
            oBookmark.runtimeStyle.padding = "0px";
            oBookmark.runtimeStyle.paddingLeft = "1px";
            oBookmark.runtimeStyle.paddingRight = "1px";
            oBookmark.runtimeStyle.border = "#888888 1px dotted";
            oBookmark.runtimeStyle.borderLeft = "#cccccc 10px solid"
        }
    }
};

function cleanFonts() {
    var oEditor = eval("idContent" + this.oName);
    var allFonts = oEditor.document.body.getElementsByTagName("FONT");
    if (allFonts.length == 0) return false;
    var f;
    while (allFonts.length > 0) {
        f = allFonts[0];
        if (f.hasChildNodes && f.childNodes.length == 1 && f.childNodes[0].nodeType == 1 && f.childNodes[0].nodeName == "SPAN") {
            copyAttribute(f.childNodes[0], f);
            f.removeNode(false)
        } else if (f.parentElement.nodeName == "SPAN" && f.parentElement.childNodes.length == 1) {
            copyAttribute(f.parentElement, f);
            f.removeNode(false)
        } else {
            var newSpan = oEditor.document.createElement("SPAN");
            copyAttribute(newSpan, f);
            var sHTML = f.innerHTML;
            sHTML = fixPathEncode(sHTML);
            newSpan.innerHTML = sHTML;
            f.replaceNode(newSpan);
            fixPathDecode(oEditor)
        }
    }
    return true
};

function cleanTags(elements, sVal) {
    var oEditor = eval("idContent" + this.oName);
    var f;
    while (elements.length > 0) {
        f = elements[0];
        if (f.hasChildNodes && f.childNodes.length == 1 && f.childNodes[0].nodeType == 1 && f.childNodes[0].nodeName == "SPAN") {
            if (sVal == "bold") f.childNodes[0].style.fontWeight = "bold";
            if (sVal == "italic") f.childNodes[0].style.fontStyle = "italic";
            if (sVal == "line-through") f.childNodes[0].style.textDecoration = "line-through";
            if (sVal == "underline") f.childNodes[0].style.textDecoration = "underline";
            f.removeNode(false)
        } else if (f.parentElement.nodeName == "SPAN" && f.parentElement.childNodes.length == 1) {
            if (sVal == "bold") f.parentElement.style.fontWeight = "bold";
            if (sVal == "italic") f.parentElement.style.fontStyle = "italic";
            if (sVal == "line-through") f.parentElement.style.textDecoration = "line-through";
            if (sVal == "underline") f.parentElement.style.textDecoration = "underline";
            f.removeNode(false)
        } else {
            var newSpan = oEditor.document.createElement("SPAN");
            if (sVal == "bold") newSpan.style.fontWeight = "bold";
            if (sVal == "italic") newSpan.style.fontStyle = "italic";
            if (sVal == "line-through") newSpan.style.textDecoration = "line-through";
            if (sVal == "underline") newSpan.style.textDecoration = "underline";
            var sHTML = f.innerHTML;
            sHTML = fixPathEncode(sHTML);
            newSpan.innerHTML = sHTML;
            f.replaceNode(newSpan);
            fixPathDecode(oEditor)
        }
    }
};

function replaceTags(sFrom, sTo) {
    var oEditor = eval("idContent" + this.oName);
    var elements = oEditor.document.getElementsByTagName(sFrom);
    var newSpan;
    var count = elements.length;
    while (count > 0) {
        f = elements[0];
        newSpan = oEditor.document.createElement(sTo);
        var sHTML = f.innerHTML;
        sHTML = fixPathEncode(sHTML);
        newSpan.innerHTML = sHTML;
        f.replaceNode(newSpan);
        fixPathDecode(oEditor);
        count--
    }
};

function cleanDeprecated() {
    var oEditor = eval("idContent" + this.oName);
    var elements;
    elements = oEditor.document.body.getElementsByTagName("STRIKE");
    this.cleanTags(elements, "line-through");
    elements = oEditor.document.body.getElementsByTagName("S");
    this.cleanTags(elements, "line-through");
    elements = oEditor.document.body.getElementsByTagName("U");
    this.cleanTags(elements, "underline");
    this.replaceTags("DIR", "DIV");
    this.replaceTags("MENU", "DIV");
    this.replaceTags("CENTER", "DIV");
    this.replaceTags("XMP", "PRE");
    this.replaceTags("BASEFONT", "SPAN");
    elements = oEditor.document.body.getElementsByTagName("APPLET");
    var count = elements.length;
    while (count > 0) {
        f = elements[0];
        f.removeNode(false);
        count--
    }
    this.cleanFonts();
    this.cleanEmptySpan();
    return true
};

function applyBold() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    this.doCmd("bold");
    return;
    var currState = oEditor.document.queryCommandState("Bold");
    if (oUtil.activeElement) oElement = oUtil.activeElement;
    else {
        if (oSel.parentElement) {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                if (oElement.tagName == "BODY") return
            } else {
                if (currState) {
                    this.applySpanStyle([
                        ["fontWeight", ""]
                    ]);
                    this.cleanEmptySpan()
                } else this.applySpanStyle([
                    ["fontWeight", "bold"]
                ]); if (currState == oEditor.document.queryCommandState("Bold") && currState == true) this.applySpanStyle([
                    ["fontWeight", "normal"]
                ]);
                return
            }
        } else oElement = oSel.item(0)
    } if (currState) oElement.style.fontWeight = "";
    else oElement.style.fontWeight = "bold"; if (currState == oEditor.document.queryCommandState("Bold") && currState == true) oElement.style.fontWeight = "normal"
};

function applyItalic() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    this.doCmd("italic");
    return;
    var currState = oEditor.document.queryCommandState("Italic");
    if (oUtil.activeElement) oElement = oUtil.activeElement;
    else {
        if (oSel.parentElement) {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                if (oElement.tagName == "BODY") return
            } else {
                if (currState) {
                    this.applySpanStyle([
                        ["fontStyle", ""]
                    ]);
                    this.cleanEmptySpan()
                } else this.applySpanStyle([
                    ["fontStyle", "italic"]
                ]); if (currState == oEditor.document.queryCommandState("Italic") && currState == true) this.applySpanStyle([
                    ["fontStyle", "normal"]
                ]);
                return
            }
        } else oElement = oSel.item(0)
    } if (currState) oElement.style.fontStyle = "";
    else oElement.style.fontStyle = "italic"; if (currState == oEditor.document.queryCommandState("Italic") && currState == true) oElement.style.fontStyle = "normal"
};

function GetUnderlinedTag(oElement) {
    while (oElement != null && oElement.style.textDecoration.indexOf("underline") == -1) {
        if (oElement.tagName == "BODY") return null;
        oElement = oElement.parentElement
    }
    return oElement
};

function GetOverlinedTag(oElement) {
    while (oElement != null && oElement.style.textDecoration.indexOf("line-through") == -1) {
        if (oElement.tagName == "BODY") return null;
        oElement = oElement.parentElement
    }
    return oElement
};

function applyLine(sCmd) {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    if (!oSel.parentElement) return;
    var bIsUnderlined = oEditor.document.queryCommandState("Underline");
    var bIsOverlined = oEditor.document.queryCommandState("Strikethrough");
    if (bIsUnderlined && !bIsOverlined) {
        if (sCmd == "underline") {
            oElement = GetUnderlinedTag(oSel.parentElement());
            if (oElement) oElement.style.textDecoration = oElement.style.textDecoration.replace("underline", "")
        } else {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                oElement.style.textDecoration = oElement.style.textDecoration + " line-through"
            } else {
                this.applySpanStyle([
                    ["textDecoration", "line-through"]
                ])
            }
        }
    } else if (bIsOverlined && !bIsUnderlined) {
        if (sCmd == "line-through") {
            oElement = GetOverlinedTag(oSel.parentElement());
            if (oElement) oElement.style.textDecoration = oElement.style.textDecoration.replace("line-through", "")
        } else {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                oElement.style.textDecoration = oElement.style.textDecoration + " underline"
            } else {
                this.applySpanStyle([
                    ["textDecoration", "underline"]
                ])
            }
        }
    } else if (bIsUnderlined && bIsOverlined) {
        if (sCmd == "underline") {
            oElement = GetUnderlinedTag(oSel.parentElement());
            if (oElement) oElement.style.textDecoration = oElement.style.textDecoration.replace("underline", "")
        } else {
            oElement = GetOverlinedTag(oSel.parentElement());
            if (oElement) oElement.style.textDecoration = oElement.style.textDecoration.replace("line-through", "")
        }
    } else {
        if (sCmd == "underline") {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                if (oElement.tagName == "BODY") return;
                oElement.style.textDecoration = "underline"
            } else this.applySpanStyle([
                ["textDecoration", "underline"]
            ])
        } else {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                if (oElement.tagName == "BODY") return;
                oElement.style.textDecoration = "line-through"
            } else this.applySpanStyle([
                ["textDecoration", "line-through"]
            ])
        }
    }
    return;
    var currState1 = oEditor.document.queryCommandState("Underline");
    var currState2 = oEditor.document.queryCommandState("Strikethrough");
    var sValue;
    if (sCmd == "underline") {
        if (currState1 && currState2) sValue = "line-through";
        else if (!currState1 && currState2) sValue = "underline line-through";
        else if (currState1 && !currState2) sValue = "";
        else if (!currState1 && !currState2) sValue = "underline"
    } else {
        if (currState1 && currState2) sValue = "underline";
        else if (!currState1 && currState2) sValue = "";
        else if (currState1 && !currState2) sValue = "underline line-through";
        else if (!currState1 && !currState2) sValue = "line-through"
    } if (oUtil.activeElement) oElement = oUtil.activeElement;
    else {
        if (oSel.parentElement) {
            if (oSel.text == "") {
                oElement = oSel.parentElement();
                if (oElement.tagName == "BODY") return
            } else {
                if (sValue == "") {
                    this.applySpanStyle([
                        ["textDecoration", ""]
                    ]);
                    this.cleanEmptySpan()
                } else this.applySpanStyle([
                    ["textDecoration", sValue]
                ]); if ((sCmd == "underline" && currState1 == oEditor.document.queryCommandState("Underline") && currState1 == true) || (sCmd == "line-through" && currState2 == oEditor.document.queryCommandState("Strikethrough") && currState2 == true)) {
                    this.applySpanStyle([
                        ["textDecoration", ""]
                    ]);
                    this.cleanEmptySpan()
                }
                return
            }
        } else oElement = oSel.item(0)
    }
    oElement.style.textDecoration = sValue;
    if ((sCmd == "underline" && currState1 == oEditor.document.queryCommandState("Underline") && currState1 == true) || (sCmd == "line-through" && currState2 == oEditor.document.queryCommandState("Strikethrough") && currState2 == true)) {
        this.applySpanStyle([
            ["textDecoration", ""]
        ]);
        this.cleanEmptySpan()
    }
};

function applyColor(sType, sColor) {
    if (!this.checkFocus()) return;
    this.hide();
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    if (oUtil.activeElement) {
        oElement = oUtil.activeElement;
        if (sType == "ForeColor") oElement.style.color = sColor;
        else oElement.style.backgroundColor = sColor
    } else if (oSel.parentElement) {
        if (oSel.text == "") {
            oElement = oSel.parentElement();
            if (oElement.tagName == "BODY") return;
            if (sType == "ForeColor") oElement.style.color = sColor;
            else oElement.style.backgroundColor = sColor
        } else {
            if (sType == "ForeColor") this.applySpanStyle([
                ["color", sColor]
            ]);
            else this.applySpanStyle([
                ["backgroundColor", sColor]
            ])
        }
    }
    if (sColor == "") {
        this.cleanEmptySpan();
        realTime(this.oName)
    }
};

function applyFontName(val) {
    this.hide();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.hide();
    oSel.select();
    this.saveForUndo();
    if (oSel.parentElement) {
        var tempRange = oEditor.document.body.createTextRange();
        var allSpans = oEditor.document.getElementsByTagName("SPAN");
        for (var i = 0; i < allSpans.length; i++) {
            tempRange.moveToElementText(allSpans[i]);
            if (oSel.inRange(tempRange)) allSpans[i].style.fontFamily = val
        }
    }
    this.doCmd("fontname", val);
    replaceWithSpan(oEditor);
    realTime(this.oName)
};

function applyFontSize(val) {
    this.hide();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.hide();
    oSel.select();
    this.saveForUndo();
    this.doCmd("fontsize", 2);
    replaceWithSpan(oEditor);
    if (oSel.parentElement) {
        var tempRange = oEditor.document.body.createTextRange();
        var allSpans = oSel.parentElement().getElementsByTagName("SPAN");
        if (allSpans.length == 0) {
            allSpans = oEditor.document.getElementsByTagName("SPAN");
            for (var i = 0; i < allSpans.length; i++) {
                tempRange.moveToElementText(allSpans[i]);
                if (oSel.inRange(tempRange)) {
                    allSpans[i].style.fontSize = val
                }
            }
        } else {
            for (var i = 0; i < allSpans.length; i++) {
                allSpans[i].style.fontSize = val
            }
        }
    }
    realTime(this.oName)
};

function applySpanStyle(arrStyles, sClassName) {
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.hide();
    oSel.select();
    this.saveForUndo();
    if (oSel.parentElement) {
        var tempRange = oEditor.document.body.createTextRange();
        var oEl = oSel.parentElement();
        var allSpans = oEditor.document.getElementsByTagName("SPAN");
        for (var i = 0; i < allSpans.length; i++) {
            tempRange.moveToElementText(allSpans[i]);
            if (oSel.inRange(tempRange)) copyStyleClass(allSpans[i], arrStyles, sClassName)
        }
    }
    this.doCmd("fontname", "default");
    replaceWithSpan(oEditor, arrStyles, sClassName);
    this.cleanEmptySpan();
    realTime(this.oName)
};

function doClean() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    this.saveForUndo();
    this.doCmd('RemoveFormat');
    if (oUtil.activeElement) {
        var oActiveElement = oUtil.activeElement;
        oActiveElement.removeAttribute("className", 0);
        oActiveElement.removeAttribute("style", 0);
        if (oActiveElement.tagName == "H1" || oActiveElement.tagName == "H2" || oActiveElement.tagName == "H3" || oActiveElement.tagName == "H4" || oActiveElement.tagName == "H5" || oActiveElement.tagName == "H6" || oActiveElement.tagName == "PRE" || oActiveElement.tagName == "P" || oActiveElement.tagName == "DIV") {
            if (this.returnKeyMode == 1) this.doCmd('FormatBlock', '<DIV>');
            else this.doCmd('FormatBlock', '<P>')
        }
    } else {
        var oSel = oEditor.document.selection.createRange();
        var sType = oEditor.document.selection.type;
        if (oSel.parentElement) {
            if (oSel.text == "") {
                oEl = oSel.parentElement();
                if (oEl.tagName == "BODY") return;
                else {
                    oEl.removeAttribute("className", 0);
                    oEl.removeAttribute("style", 0);
                    if (oEl.tagName == "H1" || oEl.tagName == "H2" || oEl.tagName == "H3" || oEl.tagName == "H4" || oEl.tagName == "H5" || oEl.tagName == "H6" || oEl.tagName == "PRE" || oEl.tagName == "P" || oEl.tagName == "DIV") {
                        if (this.returnKeyMode == 1) this.doCmd('FormatBlock', '<DIV>');
                        else this.doCmd('FormatBlock', '<P>')
                    }
                }
            } else {
                this.applySpanStyle([
                    ["backgroundColor", ""],
                    ["color", ""],
                    ["fontFamily", ""],
                    ["fontSize", ""],
                    ["fontWeight", ""],
                    ["fontStyle", ""],
                    ["textDecoration", ""],
                    ["letterSpacing", ""],
                    ["verticalAlign", ""],
                    ["textTransform", ""],
                    ["fontVariant", ""]
                ], "");
                return
            }
        } else {
            oEl = oSel.item(0);
            oEl.removeAttribute("className", 0);
            oEl.removeAttribute("style", 0)
        }
    }
    this.cleanEmptySpan();
    realTime(this.oName)
};

function cleanEmptySpan() {
    var bReturn = false;
    var oEditor = eval("idContent" + this.oName);
    var allSpans = oEditor.document.getElementsByTagName("SPAN");
    if (allSpans.length == 0) return false;
    var emptySpans = [];
    var reg = /<\s*SPAN\s*>/gi;
    for (var i = 0; i < allSpans.length; i++) {
        if (allSpans[i].outerHTML.search(reg) == 0) emptySpans[emptySpans.length] = allSpans[i]
    }
    var theSpan, theParent;
    for (var i = 0; i < emptySpans.length; i++) {
        theSpan = emptySpans[i];
        theSpan.removeNode(false);
        bReturn = true
    }
    return bReturn
};

function copyStyleClass(newSpan, arrStyles, sClassName) {
    if (arrStyles)
        for (var i = 0; i < arrStyles.length; i++) {
            newSpan.style[arrStyles[i][0]] = arrStyles[i][1]
        }
    if (newSpan.style.fontFamily == "") {
        newSpan.style.cssText = newSpan.style.cssText.replace("FONT-FAMILY: ; ", "");
        newSpan.style.cssText = newSpan.style.cssText.replace("FONT-FAMILY: ", "")
    }
    if (sClassName != null) {
        newSpan.className = sClassName;
        if (newSpan.className == "") newSpan.removeAttribute("className", 0)
    }
};

function copyAttribute(newSpan, f) {
    if ((f.face != null) && (f.face != "")) newSpan.style.fontFamily = f.face;
    if ((f.size != null) && (f.size != "")) {
        var nSize = "";
        if (f.size == 1) nSize = "8pt";
        else if (f.size == 2) nSize = "10pt";
        else if (f.size == 3) nSize = "12pt";
        else if (f.size == 4) nSize = "14pt";
        else if (f.size == 5) nSize = "18pt";
        else if (f.size == 6) nSize = "24pt";
        else if (f.size >= 7) nSize = "36pt";
        else if (f.size <= -2 || f.size == "0") nSize = "8pt";
        else if (f.size == "-1") nSize = "10pt";
        else if (f.size == 0) nSize = "12pt";
        else if (f.size == "+1") nSize = "14pt";
        else if (f.size == "+2") nSize = "18pt";
        else if (f.size == "+3") nSize = "24pt";
        else if (f.size == "+4" || f.size == "+5" || f.size == "+6") nSize = "36pt";
        else nSize = ""; if (nSize != "") newSpan.style.fontSize = nSize
    }
    if ((f.style.backgroundColor != null) && (f.style.backgroundColor != "")) newSpan.style.backgroundColor = f.style.backgroundColor;
    if ((f.color != null) && (f.color != "")) newSpan.style.color = f.color;
    if ((f.className != null) && (f.className != "")) newSpan.className = f.className
};

function replaceWithSpan(oEditor, arrStyles, sClassName) {
    var oSel = oEditor.document.selection.createRange();
    var oSpanStart;
    oSel.select();
    var nSelLength = oSel.text.length;
    var allFonts = new Array();
    if (oSel.parentElement().nodeName == "FONT" && oSel.parentElement().innerText == oSel.text) {
        oSel.moveToElementText(oSel.parentElement());
        allFonts[0] = oSel.parentElement()
    } else {
        allFonts = oEditor.document.getElementsByTagName("FONT")
    }
    var tempRange = oEditor.document.body.createTextRange();
    var newSpan, f;
    var count = allFonts.length;
    while (count > 0) {
        f = allFonts[0];
        if (f == null || f.parentElement == null) {
            count--;
            continue
        }
        tempRange.moveToElementText(f);
        var sTemp = "f";
        var nLevel = 0;
        while (eval(sTemp + ".parentElement")) {
            nLevel++;
            sTemp += ".parentElement"
        }
        var bBreak = false;
        for (var j = nLevel; j > 0; j--) {
            sTemp = "f";
            for (var k = 1; k <= j; k++) sTemp += ".parentElement";
            if (!bBreak)
                if (eval(sTemp).nodeName == "SPAN" && eval(sTemp).innerText == f.innerText) {
                    newSpan = eval(sTemp);
                    if (arrStyles || sClassName) copyStyleClass(newSpan, arrStyles, sClassName);
                    else copyAttribute(newSpan, f);
                    f.removeNode(false);
                    bBreak = true
                }
        }
        if (bBreak) {
            continue
        }
        newSpan = oEditor.document.createElement("SPAN");
        if (arrStyles || sClassName) copyStyleClass(newSpan, arrStyles, sClassName);
        else copyAttribute(newSpan, f);
        var sHTML = f.innerHTML;
        sHTML = fixPathEncode(sHTML);
        newSpan.innerHTML = sHTML;
        f.replaceNode(newSpan);
        fixPathDecode(oEditor);
        count--;
        if (!oSpanStart) oSpanStart = newSpan
    }
    var rng = oEditor.document.selection.createRange();
    if (oSpanStart) {
        rng.moveToElementText(oSpanStart);
        rng.select()
    }
    rng.moveEnd("character", nSelLength - rng.text.length);
    rng.select();
    rng.moveEnd("character", nSelLength - rng.text.length);
    rng.select();
    rng.moveEnd("character", nSelLength - rng.text.length);
    rng.select()
};

function doOnPaste() {
    this.isAfterPaste = true;
    this.saveForUndo();
    if (this.pasteTextOnCtrlV == true) {
        this.doPasteText();
        return false
    }
};

function doPaste() {
    this.saveForUndo();
    if (this.pasteTextOnCtrlV == true) {
        this.doOnPaste()
    } else {
        var edt = eval(this.oName);
        edt.bookmarkSelection();
        var pasteFrame = document.getElementById("idContentWord" + this.oName);
        var oDoc = pasteFrame.contentWindow.document;
        var oHTML = oDoc.open("text/html", "replace");
        oHTML.close();
        oDoc.body.contentEditable = true;
        pasteFrame.contentWindow.focus();
        var oEditorWord = eval("idContentWord" + this.oName);
        var oSelWord = oEditorWord.document.selection.createRange();
        var sTypeWord = oEditorWord.document.selection.type;
        var oTarget = (sTypeWord == "None" ? oEditorWord.document : sTypeWord);
        oTarget.execCommand("Paste", false);
        window.setTimeout(function () {
            edt.setFocus();
            edt.fixWord()
        }, 50)
    }
    this.runtimeBorder(false)
};

function doCmd(sCmd, sOption) {
    if (!this.checkFocus()) return;
    if (sCmd == "Cut" || sCmd == "Copy" || sCmd == "Superscript" || sCmd == "Subscript" || sCmd == "Indent" || sCmd == "Outdent" || sCmd == "InsertHorizontalRule" || sCmd == "BlockDirLTR" || sCmd == "BlockDirRTL") this.saveForUndo();
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    var oTarget = (sType == "None" ? oEditor.document : oSel);
    oTarget.execCommand(sCmd, false, sOption);
    realTime(this.oName)
};

function applyParagraph(val) {
    this.hide();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.hide();
    oSel.select();
    this.saveForUndo();
    this.doCmd("FormatBlock", val)
};

function applyBullets() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("InsertUnOrderedList");
    this.tbar.btns["btnNumbering" + this.oName].setState(1);
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var oElement = oSel.parentElement();
    while (oElement != null && oElement.tagName != "OL" && oElement.tagName != "UL") {
        if (oElement.tagName == "BODY") return;
        oElement = oElement.parentElement
    }
    oElement.removeAttribute("type", 0);
    oElement.style.listStyleImage = ""
};

function applyNumbering() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("InsertOrderedList");
    this.tbar.btns["btnBullets" + this.oName].setState(1);
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var oElement = oSel.parentElement();
    while (oElement != null && oElement.tagName != "OL" && oElement.tagName != "UL") {
        if (oElement.tagName == "BODY") return;
        oElement = oElement.parentElement
    }
    oElement.removeAttribute("type", 0);
    oElement.style.listStyleImage = ""
};

function applyJustifyLeft() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("JustifyLeft")
};

function applyJustifyCenter() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("JustifyCenter")
};

function applyJustifyRight() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("JustifyRight")
};

function applyJustifyFull() {
    if (!this.checkFocus()) return;
    this.saveForUndo();
    this.doCmd("JustifyFull")
};

function applyBlockDirLTR() {
    if (!this.checkFocus()) return;
    this.doCmd("BlockDirLTR")
};

function applyBlockDirRTL() {
    if (!this.checkFocus()) return;
    this.doCmd("BlockDirRTL")
};

function doPasteText() {
    if (!this.checkFocus()) return;
    var oWord = eval("idContentWord" + this.oName);
    oWord.document.open("text/html", "replace");
    oWord.document.write("<html><head></head><body></body></html>");
    oWord.document.close();
    oWord.document.body.contentEditable = true;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    var oWord = eval("idContentWord" + this.oName);
    oWord.focus();
    oWord.document.execCommand("SelectAll");
    oWord.document.execCommand("Paste");
    var sHTML = oWord.document.body.innerHTML;
    sHTML = sHTML.replace(/(<br>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/tr>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/div>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h1>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h2>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h3>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h4>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h5>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<\/h6>)/gi, "$1&lt;REPBR&gt;");
    sHTML = sHTML.replace(/(<p>)/gi, "$1&lt;REPBR&gt;");
    sHTML = fixPathEncode(sHTML);
    oWord.document.body.innerHTML = sHTML;
    fixPathDecode(oWord);
    sHTML = oWord.document.body.innerText.replace(/<REPBR>/gi, "<br />");
    sHTML = sHTML.replace(/[\n\t\r]/gi, "");
    sHTML = sHTML.replace(/>\s*</gi, "><");
    oSel.pasteHTML(sHTML)
};

function insertCustomTag(index) {
    this.hide();
    if (!this.checkFocus()) return;
    this.insertHTML(this.arrCustomTag[index][1]);
    this.hide();
    this.focus()
};

function insertHTML(sHTML) {
    this.setFocus();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    var arrA = String(sHTML).match(/<A[^>]*>/ig);
    if (arrA)
        for (var i = 0; i < arrA.length; i++) {
            sTmp = arrA[i].replace(/href=/, "href_iwe=");
            sHTML = String(sHTML).replace(arrA[i], sTmp)
        }
    var arrB = String(sHTML).match(/<IMG[^>]*>/ig);
    if (arrB)
        for (var i = 0; i < arrB.length; i++) {
            sTmp = arrB[i].replace(/src=/, "src_iwe=");
            sHTML = String(sHTML).replace(arrB[i], sTmp)
        }
    if (oSel.parentElement) {
        if (isIE9) {
            var oSel = oEditor.getSelection();
            var range = oSel.getRangeAt(0);
            var frag = oEditor.document.createDocumentFragment(),
                div = oEditor.document.createElement("div");
            frag.appendChild(div);
            div.outerHTML = sHTML;
            range.insertNode(frag)
        } else {
            oSel.pasteHTML("<span id='iwe_delete'>toberemoved</span>" + sHTML)
        }
    } else oSel.item(0).outerHTML = sHTML;
    var toRemoved = oEditor.document.getElementById("iwe_delete");
    if (toRemoved) toRemoved.removeNode(true);
    for (var i = 0; i < oEditor.document.all.length; i++) {
        var elm = oEditor.document.all[i];
        if (elm.nodeType == 1) {
            if (elm.getAttribute("href_iwe")) {
                elm.href = elm.getAttribute("href_iwe");
                elm.removeAttribute("href_iwe", 0)
            }
            if (elm.getAttribute("src_iwe")) {
                elm.src = elm.getAttribute("src_iwe");
                elm.removeAttribute("src_iwe", 0)
            }
        }
    }
    this.bookmarkSelection()
};

function insertLink(url, title, target, rel) {
    this.setFocus();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    if (oSel.parentElement) {
        if (oSel.text == "") {
            var oSelTmp = oSel.duplicate();
            if (title != "" && title != undefined) oSel.text = title;
            else oSel.text = url;
            oSel.setEndPoint("StartToStart", oSelTmp);
            oSel.select()
        }
    }
    oSel.execCommand("CreateLink", false, url);
    if (oSel.parentElement) oEl = GetElement(oSel.parentElement(), "A");
    else oEl = GetElement(oSel.item(0), "A"); if (oEl) {
        if (target != "" && target != undefined) oEl.target = target;
        if (rel != "" && rel != undefined) oEl.setAttribute("rel", rel)
    }
    this.bookmarkSelection()
};

function clearAll() {
    if (confirm(getTxt("Are you sure you wish to delete all contents?")) == true) {
        var oEditor = eval("idContent" + this.oName);
        this.saveForUndo();
        oEditor.document.body.innerHTML = ""
    }
};

function applySpan() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var sType = oEditor.document.selection.type;
    if (sType == "Control" || sType == "None") return;
    sHTML = oSel.htmlText;
    var oParent = oSel.parentElement();
    if (oParent)
        if (oParent.innerText == oSel.text) {
            if (oParent.tagName == "SPAN") {
                idSpan = oParent;
                return idSpan
            }
        }
    var arrA = String(sHTML).match(/<A[^>]*>/ig);
    if (arrA)
        for (var i = 0; i < arrA.length; i++) {
            sTmp = arrA[i].replace(/href=/, "href_iwe=");
            sHTML = String(sHTML).replace(arrA[i], sTmp)
        }
    var arrB = String(sHTML).match(/<IMG[^>]*>/ig);
    if (arrB)
        for (var i = 0; i < arrB.length; i++) {
            sTmp = arrB[i].replace(/src=/, "src_iwe=");
            sHTML = String(sHTML).replace(arrB[i], sTmp)
        }
    oSel.pasteHTML("<SPAN id='idSpan__abc'>" + sHTML + "</SPAN>");
    var idSpan = oEditor.document.all.idSpan__abc;
    var oSelRange = oEditor.document.body.createTextRange();
    oSelRange.moveToElementText(idSpan);
    oSel.setEndPoint("StartToStart", oSelRange);
    oSel.setEndPoint("EndToEnd", oSelRange);
    oSel.select();
    for (var i = 0; i < oEditor.document.all.length; i++) {
        if (oEditor.document.all[i].getAttribute("href_iwe")) {
            oEditor.document.all[i].href = oEditor.document.all[i].getAttribute("href_iwe");
            oEditor.document.all[i].removeAttribute("href_iwe", 0)
        }
        if (oEditor.document.all[i].getAttribute("src_iwe")) {
            oEditor.document.all[i].src = oEditor.document.all[i].getAttribute("src_iwe");
            oEditor.document.all[i].removeAttribute("src_iwe", 0)
        }
    }
    idSpan.removeAttribute("id", 0);
    return idSpan
};

function makeAbsolute() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    this.saveForUndo();
    if (oSel.parentElement) {
        var oElement = oSel.parentElement();
        oElement.style.position = "fixed"
    } else this.doCmd("AbsolutePosition")
};

function expandSelection() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    if (oSel.text != "") return;
    oSel.expand("word");
    oSel.select();
    if (oSel.text.substr(oSel.text.length * 1 - 1, oSel.text.length) == " ") {
        oSel.moveEnd("character", -1);
        oSel.select()
    }
};

function selectParagraph() {
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    if (oSel.parentElement) {
        if (oSel.text == "") {
            var oElement = oSel.parentElement();
            while (oElement != null && oElement.tagName != "H1" && oElement.tagName != "H2" && oElement.tagName != "H3" && oElement.tagName != "H4" && oElement.tagName != "H5" && oElement.tagName != "H6" && oElement.tagName != "PRE" && oElement.tagName != "P" && oElement.tagName != "DIV") {
                if (oElement.tagName == "BODY") return;
                oElement = oElement.parentElement
            }
            var oSelRange = oEditor.document.body.createControlRange();
            try {
                oSelRange.add(oElement);
                oSelRange.select()
            } catch (e) {
                var oSelRange = oEditor.document.body.createTextRange();
                try {
                    oSelRange.moveToElementText(oElement);
                    oSelRange.select()
                } catch (e) {}
            }
        }
    }
};

function doOver_TabCreate() {
    var oTD = event.srcElement;
    var oTable = oTD.parentElement.parentElement.parentElement;
    var nRow = oTD.parentElement.rowIndex;
    var nCol = oTD.cellIndex;
    var rows = oTable.rows;
    rows[rows.length - 1].childNodes[0].innerHTML = "<div>" + (nRow * 1 + 1) + " x " + (nCol * 1 + 1) + " " + getTxt("Table Dimension Text") + "</div>";
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            if (i <= nRow && j <= nCol) oCol.style.backgroundColor = "#316ac5";
            else oCol.style.backgroundColor = "#ffffff"
        }
    }
    event.cancelBubble = true
};

function doOut_TabCreate() {
    var oTable = event.srcElement;
    if (oTable.tagName != "TABLE") return;
    var rows = oTable.rows;
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            oCol.style.backgroundColor = "#ffffff"
        }
    }
    event.cancelBubble = true
};

function doRefresh_TabCreate() {
    if (!this.btnTable) return;
    var oTable = eval("dropTableCreate" + this.oName);
    var rows = oTable.rows;
    for (var i = 0; i < rows.length - 1; i++) {
        var oRow = rows[i];
        for (var j = 0; j < oRow.childNodes.length; j++) {
            var oCol = oRow.childNodes[j];
            oCol.style.backgroundColor = "#ffffff"
        }
    }
};

function doClick_TabCreate() {
    this.hide();
    if (!this.checkFocus()) return;
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var oTD = event.srcElement;
    var nRow = oTD.parentElement.rowIndex + 1;
    var nCol = oTD.cellIndex + 1;
    this.saveForUndo();
    var sHTML = "<table style='border-collapse:collapse;width:100%;'>";
    for (var i = 1; i <= nRow; i++) {
        sHTML += "<tr>";
        for (var j = 1; j <= nCol; j++) {
            sHTML += "<td></td>"
        }
        sHTML += "</tr>"
    }
    sHTML += "</table>";
    if (oSel.parentElement) {
        oSel.collapse();
        oSel.pasteHTML(sHTML)
    } else oSel.item(0).outerHTML = sHTML;
    realTime(this.oName);
    this.runtimeBorder(false);
    this.runtimeStyles()
};

function doKeyPress(evt, oName) {
    var edt = eval(oName);
    if (!edt.arrUndoList[0]) {
        edt.saveForUndo()
    }
    if (evt.ctrlKey) {
        if (evt.keyCode == 86) {
            edt.bookmarkSelection();
            if (edt.pasteTextOnCtrlV == true) {} else {
                var pasteFrame = document.getElementById("idContentWord" + oName);
                var oDoc = pasteFrame.contentWindow.document;
                var oHTML = oDoc.open("text/html", "replace");
                oHTML.close();
                oDoc.body.contentEditable = true;
                pasteFrame.contentWindow.focus();
                window.setTimeout(function () {
                    edt.setFocus();
                    edt.fixWord()
                }, 50)
            }
        }
        if (evt.keyCode == 89) {
            if (!evt.altKey) edt.doRedo()
        }
        if (evt.keyCode == 90) {
            if (!evt.altKey) edt.doUndo()
        }
        if (evt.keyCode == 65) {
            if (!evt.altKey) edt.doCmd("SelectAll")
        }
        if (evt.keyCode == 66 || evt.keyCode == 73 || evt.keyCode == 85) {
            if (!evt.altKey) edt.saveForUndo()
        }
    }
    if (evt.keyCode == 37 || evt.keyCode == 38 || evt.keyCode == 39 || evt.keyCode == 40) {
        edt.saveForUndo()
    }
    if (evt.keyCode == 13) {
        if (edt.returnKeyMode == 1) {
            var oSel = document.selection.createRange();
            if (oSel.parentElement) {
                edt.saveForUndo();
                if (GetElement(oSel.parentElement(), "LI")) {} else if (GetElement(oSel.parentElement(), "FORM")) {
                    var oSel = document.selection.createRange();
                    oSel.pasteHTML('<br>');
                    evt.cancelBubble = true;
                    evt.returnValue = false;
                    oSel.select();
                    oSel.moveEnd("character", 1);
                    oSel.moveStart("character", 1);
                    oSel.collapse(false);
                    return false
                } else {
                    var oEl = GetElement(oSel.parentElement(), "H1");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "H2");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "H3");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "H4");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "H5");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "H6");
                    if (!oEl) oEl = GetElement(oSel.parentElement(), "PRE");
                    if (!oEl) edt.doCmd("FormatBlock", "<div>");
                    return true
                }
            }
        } else if (edt.returnKeyMode == 2) {
            var oSel = document.selection.createRange();
            if (oSel.parentElement) {
                if (GetElement(oSel.parentElement(), "LI")) {} else {
                    oSel.pasteHTML('<br>');
                    evt.cancelBubble = true;
                    evt.returnValue = false;
                    oSel.select();
                    oSel.moveEnd("character", 1);
                    oSel.moveStart("character", 1);
                    oSel.collapse(false);
                    return false
                }
            }
        }
        edt.saveForUndo()
    }
    edt.onKeyPress(evt);
};

function fullScreen() {
    this.hide();
    var oEditor = eval("idContent" + this.oName);
    if (this.stateFullScreen) {
        this.onNormalScreen();
        this.stateFullScreen = false;
        document.body.style.overflow = "";
        var idArea = eval("idArea" + this.oName);
        idArea.style.position = "";
        idArea.style.top = 0;
        idArea.style.left = 0;
        var w = new String(this.width);
        if (w.indexOf("%") == -1) {
            w = parseInt(w, 10) + "px"
        }
        idArea.style.width = w;
        var h = new String(this.height);
        if (h.indexOf("%") == -1) {
            h = parseInt(h, 10) + "px"
        }
        idArea.style.height = this.height;
        var ifrm = document.getElementById("idFixZIndex" + this.oName);
        ifrm.style.top = 0;
        ifrm.style.left = 0;
        ifrm.style.width = 0;
        ifrm.style.height = 0;
        ifrm.style.display = "none";
        oEditor.document.body.style.lineHeight = "1.2";
        window.setTimeout("eval('idContent" + this.oName + "').document.body.style.lineHeight='';", 0);
        for (var i = 0; i < oUtil.arrEditor.length; i++) {
            if (oUtil.arrEditor[i] != this.oName) eval("idArea" + oUtil.arrEditor[i]).style.display = "block"
        }
    } else {
        this.onFullScreen();
        this.stateFullScreen = true;
        scroll(0, 0);
        var idArea = eval("idArea" + this.oName);
        idArea.style.position = "fixed";
        idArea.style.top = 0;
        idArea.style.left = 0;
        idArea.style.zIndex = 2000;
        var tbPc = document.getElementById("idToolbar" + this.oName);
        nToolbarHeight = tbPc.offsetHeight;
        if (this.useTagSelector) {
            nToolbarHeight += 13
        }
        if (this.showResizeBar) {
            nToolbarHeight += 8
        }
        var w = 0,
            h = 0;
        if (isIE9) {
            w = window.innerWidth;
            h = window.innerHeight;
            h = h - nToolbarHeight
        } else {
            if (document.compatMode && document.compatMode != "BackCompat") {
                var html = document.documentElement;
                try {
                    w = (document.body.offsetWidth);
                    document.body.style.height = "100%";
                    h = html.clientHeight - nToolbarHeight;
                    document.body.style.height = ""
                } catch (e) {
                    w = (document.body.offsetWidth + 20);
                    document.body.style.height = "100%";
                    h = html.clientHeight - nToolbarHeight;
                    document.body.style.height = ""
                }
            } else {
                if (document.body.style.overflow == "hidden") {
                    w = document.body.offsetWidth
                } else {
                    w = document.body.offsetWidth - 22
                }
                h = document.body.offsetHeight - 4
            }
        }
        idArea.style.width = w + "px";
        idArea.style.height = h + "px";
        var ifrm = document.getElementById("idFixZIndex" + this.oName);
        ifrm.style.top = 0;
        ifrm.style.left = 0;
        ifrm.style.width = w;
        ifrm.style.height = h;
        ifrm.style.display = "";
        ifrm.style.zIndex = 1900;
        for (var i = 0; i < oUtil.arrEditor.length; i++) {
            if (oUtil.arrEditor[i] != this.oName) eval("idArea" + oUtil.arrEditor[i]).style.display = "none"
        }
        oEditor.document.body.style.lineHeight = "1.2";
        window.setTimeout("eval('idContent" + this.oName + "').document.body.style.lineHeight='';", 0);
        oEditor.focus()
    }
    var idStyles = document.getElementById("idStyles" + this.oName);
    idStyles.innerHTML = ""
};

function hide() {
    hideAllDD();
    this.oColor1.hide();
    this.oColor2.hide();
    if (this.btnTable) this.doRefresh_TabCreate()
};

function convertBorderWidth(width) {
    return eval(width.substr(0, width.length - 2))
};

function modelessDialogShow(url, width, height, p, opt) {
    modalDialog(url, width, height)
};

function modalDialogShow(url, width, height, p, opt) {
    modalDialog(url, width, height)
};

function windowOpen(url, wd, hg, ov, p, opt) {
    var id = "ID" + (new Date()).getTime();
    var f = new ISWindow(id);
    f.iconPath = oUtil.scriptPath + "icons/";
    f.show({
        width: wd + "px",
        height: hg + "px",
        overlay: ov,
        center: true,
        url: url,
        openerWin: p,
        options: opt
    })
};

function lineBreak1(tag) {
    arrReturn = ["\n", "", ""];
    if (tag == "A" || tag == "B" || tag == "CITE" || tag == "CODE" || tag == "EM" || tag == "FONT" || tag == "I" || tag == "SMALL" || tag == "STRIKE" || tag == "BIG" || tag == "STRONG" || tag == "SUB" || tag == "SUP" || tag == "U" || tag == "SAMP" || tag == "S" || tag == "VAR" || tag == "BASEFONT" || tag == "KBD" || tag == "TT" || tag == "SPAN" || tag == "IMG") arrReturn = ["", "", ""];
    if (tag == "TEXTAREA" || tag == "TABLE" || tag == "THEAD" || tag == "TBODY" || tag == "TR" || tag == "OL" || tag == "UL" || tag == "DIR" || tag == "MENU" || tag == "FORM" || tag == "SELECT" || tag == "MAP" || tag == "DL" || tag == "HEAD" || tag == "BODY" || tag == "HTML") arrReturn = ["\n", "", "\n"];
    if (tag == "STYLE" || tag == "SCRIPT") arrReturn = ["\n", "", ""];
    if (tag == "BR" || tag == "HR") arrReturn = ["", "\n", ""];
    return arrReturn
};

function fixAttr(s) {
    s = String(s).replace(/&/g, "&amp;");
    s = String(s).replace(/</g, "&lt;");
    s = String(s).replace(/"/g, "&quot;");
    return s
};

function fixVal(s) {
    s = String(s).replace(/&/g, "&amp;");
    s = String(s).replace(/</g, "&lt;");
    var x = escape(s);
    x = unescape(x.replace(/\%A0/gi, "-*REPL*-"));
    s = x.replace(/-\*REPL\*-/gi, "&nbsp;");
    return s
};

function recur(oEl, sTab) {
    var sHTML = "";
    for (var i = 0; i < oEl.childNodes.length; i++) {
        var oNode = oEl.childNodes[i];
        if (oNode.parentNode != oEl) continue;
        if (oNode.nodeType == 1) {
            var sTagName = oNode.nodeName;
            var sCloseTag = oNode.outerHTML;
            if (sCloseTag.indexOf("<?xml:namespace") > -1) sCloseTag = sCloseTag.substr(sCloseTag.indexOf(">") + 1);
            sCloseTag = sCloseTag.substring(1, sCloseTag.indexOf(">"));
            if (sCloseTag.indexOf(" ") > -1) sCloseTag = sCloseTag.substring(0, sCloseTag.indexOf(" "));
            var bDoNotProcess = false;
            if (sTagName.substring(0, 1) == "/") {
                bDoNotProcess = true
            } else {
                var sT = sTab;
                sHTML += lineBreak1(sTagName)[0];
                if (lineBreak1(sTagName)[0] != "") sHTML += sT
            } if (bDoNotProcess) {} else if (sTagName == "OBJECT" || sTagName == "EMBED") {
                s = getOuterHTML(oNode);
                s = s.replace(/\"[^\"]*\"/ig, function (x) {
                    x = x.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/'/g, "&apos;").replace(/\s+/ig, "#_#").replace(/&amp;amp;/gi, "&amp;");
                    return x
                });
                s = s.replace(/<([^ >]*)/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/ ([^=]+)=([^"' >]+)/ig, " $1=\"$2\"");
                s = s.replace(/ ([^=]+)=/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/#_#/ig, " ");
                if (sTagName == "EMBED")
                    if (oNode.innerHTML == "") s = s.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
                s = s.replace(/<param name=\"Play\" value=\"0\" \/>/, "<param name=\"Play\" value=\"-1\" \/>");
                sHTML += s
            } else if (sTagName == "TITLE") {
                sHTML += "<title>" + oNode.innerHTML + "</title>"
            } else {
                if (sTagName == "AREA") {
                    var sCoords = oNode.coords;
                    var sShape = oNode.shape
                }
                if (sTagName == "BODY") {
                    var ht = oNode.outerHTML;
                    if (isIE9) {
                        ht = getOuterHTML(oNode)
                    }
                    s = ht.substring(0, ht.indexOf(">") + 1)
                } else {
                    var oNode2 = oNode.cloneNode();
                    if (oNode.checked) oNode2.checked = oNode.checked;
                    if (oNode.selected) oNode2.selected = oNode.selected;
                    s = oNode2.outerHTML.replace(/<\/[^>]*>/, "");
                    if (isIE9) {
                        s = getOuterHTML(oNode2).replace(/<\/[^>]*>/, "")
                    }
                } if (sTagName == "STYLE") {
                    var arrTmp = s.match(/<[^>]*>/ig);
                    s = arrTmp[0]
                }
                s = s.replace(/\"[^\"]*\"/ig, function (x) {
                    x = x.replace(/&/g, "&amp;").replace(/&amp;amp;/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\s+/ig, "#_#");
                    return x
                });
                s = s.replace(/<([^ >]*)/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/ ([^=]+)=([^" >]+)/ig, " $1=\"$2\"");
                s = s.replace(/ ([^=]+)=/ig, function (x) {
                    return x.toLowerCase()
                });
                s = s.replace(/#_#/ig, " ");
                s = s.replace(/(<hr[^>]*)(noshade=""|noshade )/ig, "$1noshade=\"noshade\" ");
                s = s.replace(/(<input[^>]*)(checked=""|checked )/ig, "$1checked=\"checked\" ");
                s = s.replace(/(<select[^>]*)(multiple=""|multiple )/ig, "$1multiple=\"multiple\" ");
                s = s.replace(/(<option[^>]*)(selected=""|selected )/ig, "$1selected=\"true\" ");
                s = s.replace(/(<input[^>]*)(readonly=""|readonly )/ig, "$1readonly=\"readonly\" ");
                s = s.replace(/(<input[^>]*)(disabled=""|disabled )/ig, "$1disabled=\"disabled\" ");
                s = s.replace(/(<td[^>]*)(nowrap=""|nowrap )/ig, "$1nowrap=\"nowrap\" ");
                s = s.replace(/(<td[^>]*)(nowrap=""\>|nowrap\>)/ig, "$1nowrap=\"nowrap\"\>");
                if (sTagName == "AREA") {
                    s = s.replace(/ coords=\"0,0,0,0\"/ig, " coords=\"" + sCoords + "\"");
                    s = s.replace(/ shape=\"RECT\"/ig, " shape=\"" + sShape + "\"")
                }
                var bClosingTag = true;
                if (sTagName == "IMG" || sTagName == "BR" || sTagName == "AREA" || sTagName == "HR" || sTagName == "INPUT" || sTagName == "BASE" || sTagName == "LINK") {
                    s = s.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>");
                    bClosingTag = false
                }
                sHTML += s;
                if (sTagName != "TEXTAREA") sHTML += lineBreak1(sTagName)[1];
                if (sTagName != "TEXTAREA")
                    if (lineBreak1(sTagName)[1] != "") sHTML += sT;
                if (bClosingTag) {
                    s = oNode.outerHTML;
                    if (sTagName == "SCRIPT") {
                        s = s.replace(/<script([^>]*)>[\n+\s+\t+]*/ig, "<script$1>");
                        s = s.replace(/[\n+\s+\t+]*<\/script>/ig, "<\/script>");
                        s = s.replace(/<script([^>]*)>\/\/<!\[CDATA\[/ig, "");
                        s = s.replace(/\/\/\]\]><\/script>/ig, "");
                        s = s.replace(/<script([^>]*)>/ig, "");
                        s = s.replace(/<\/script>/ig, "");
                        s = s.replace(/^\s+/, '').replace(/\s+$/, '');
                        sHTML += "\n" + sT + "//<![CDATA[\n" + sT + s + "\n" + sT + "//]]>\n" + sT
                    }
                    if (sTagName == "STYLE") {
                        s = s.replace(/<style([^>]*)>[\n+\s+\t+]*/ig, "<style$1>");
                        s = s.replace(/[\n+\s+\t+]*<\/style>/ig, "<\/style>");
                        s = s.replace(/<style([^>]*)><!--/ig, "");
                        s = s.replace(/--><\/style>/ig, "");
                        s = s.replace(/<style([^>]*)>/ig, "");
                        s = s.replace(/<\/style>/ig, "");
                        s = s.replace(/^\s+/, "").replace(/\s+$/, "");
                        sHTML += "\n" + sT + "<!--\n" + sT + s + "\n" + sT + "-->\n" + sT
                    }
                    if (sTagName == "DIV" || sTagName == "P") {
                        if (oNode.innerHTML == "" || oNode.innerHTML == "&nbsp;") {
                            sHTML += "&nbsp;"
                        } else sHTML += recur(oNode, sT + "\t")
                    } else if (sTagName == "STYLE" || sTagName == "SCRIPT") {} else {
                        sHTML += recur(oNode, sT + "\t")
                    } if (sTagName != "TEXTAREA") sHTML += lineBreak1(sTagName)[2];
                    if (sTagName != "TEXTAREA")
                        if (lineBreak1(sTagName)[2] != "") sHTML += sT;
                    if (sCloseTag.indexOf(":") >= 0) {
                        sHTML += "</" + sCloseTag.toLowerCase() + ">"
                    } else {
                        sHTML += "</" + sTagName.toLowerCase() + ">"
                    }
                }
            }
        } else if (oNode.nodeType == 3) {
            sHTML += fixVal(oNode.nodeValue).replace(/^[\t\r\n\v\f]*/, "").replace(/[\t\r\n\v\f]*$/, "")
        } else if (oNode.nodeType == 8) {
            var sTmp = oNode.nodeValue;
            sTmp = sTmp.replace(/^\s+/, "").replace(/\s+$/, "");
            var sT = "";
            sHTML += "\n" + sT + "<!--\n" + sT + sTmp + "\n" + sT + "-->\n" + sT
        } else {}
    }
    return sHTML
};

function fixPathEncode(sHTML) {
    var arrA = String(sHTML).match(/<A[^>]*>/g);
    if (arrA)
        for (var i = 0; i < arrA.length; i++) {
            sTmp = arrA[i].replace(/href=/, "href_iwe=");
            sHTML = String(sHTML).replace(arrA[i], sTmp)
        }
    var arrB = String(sHTML).match(/<IMG[^>]*>/g);
    if (arrB)
        for (var i = 0; i < arrB.length; i++) {
            sTmp = arrB[i].replace(/src=/, "src_iwe=");
            sHTML = String(sHTML).replace(arrB[i], sTmp)
        }
    var arrC = String(sHTML).match(/<AREA[^>]*>/ig);
    if (arrC)
        for (var i = 0; i < arrC.length; i++) {
            sTmp = arrC[i].replace(/href=/, "href_iwe=");
            sHTML = String(sHTML).replace(arrC[i], sTmp)
        }
    return sHTML
};

function fixPathDecode(oEditor) {
    for (var i = 0; i < oEditor.document.all.length; i++) {
        if (oEditor.document.all[i].getAttribute("href_iwe")) {
            oEditor.document.all[i].href = oEditor.document.all[i].getAttribute("href_iwe");
            oEditor.document.all[i].removeAttribute("href_iwe", 0)
        }
        if (oEditor.document.all[i].getAttribute("src_iwe")) {
            oEditor.document.all[i].src = oEditor.document.all[i].getAttribute("src_iwe");
            oEditor.document.all[i].removeAttribute("src_iwe", 0)
        }
    }
};

function tbAction(tb, id, edt, sfx) {
    var e = edt,
        oN = sfx,
        btn = id.substring(0, id.lastIndexOf(oN));
    switch (btn) {
    case "btnFullScreen":
        e.fullScreen();
        break;
    case "btnPrint":
        e.focus();
        edt.doCmd("Print");
        break;
    case "btnSpellCheck":
        e.hide();
        if (e.spellCheckMode == "ieSpell") modelessDialogShow(e.scriptPath + "spellcheck.htm", 500, 222);
        else if (e.spellCheckMode == "NetSpell") checkSpellingById("idContent" + edt.oName);
        break;
    case "btnCut":
        e.doCmd("Cut");
        break;
    case "btnCopy":
        e.doCmd("Copy");
        break;
    case "btnUndo":
        e.doUndo();
        break;
    case "btnRedo":
        e.doRedo();
        break;
    case "btnBold":
        e.applyBold();
        break;
    case "btnItalic":
        e.applyItalic();
        break;
    case "btnUnderline":
        e.applyLine("underline");
        break;
    case "btnStrikethrough":
        e.applyLine("line-through");
        break;
    case "btnSuperscript":
        e.doCmd("Superscript");
        break;
    case "btnSubscript":
        e.doCmd("Subscript");
        break;
    case "btnJustifyLeft":
        e.applyJustifyLeft();
        break;
    case "btnJustifyCenter":
        e.applyJustifyCenter();
        break;
    case "btnJustifyRight":
        e.applyJustifyRight();
        break;
    case "btnJustifyFull":
        e.applyJustifyFull();
        break;
    case "btnNumbering":
        e.applyNumbering();
        break;
    case "btnBullets":
        e.applyBullets();
        break;
    case "btnIndent":
        e.doCmd("Indent");
        break;
    case "btnOutdent":
        e.doCmd("Outdent");
        break;
    case "btnLTR":
        e.applyBlockDirLTR();
        break;
    case "btnRTL":
        e.applyBlockDirRTL();
        break;
    case "btnFontDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webfonts.htm", e.dialogSize["FontDialog"].w, e.dialogSize["FontDialog"].h);
        break;
    case "btnTextDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtext.htm", e.dialogSize["TextDialog"].w, e.dialogSize["TextDialog"].h);
        break;
    case "btnCompleteTextDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtextcomplete.htm", e.dialogSize["CompleteTextDialog"].w, e.dialogSize["CompleteTextDialog"].h);
        break;
    case "btnLinkDialog":
        e.hide();
        if (e.fileBrowser == "" && e.enableCssButtons == false) {
            modelessDialogShow(e.dialogPath + "weblink.htm", 300, 400)
        } else {
            modelessDialogShow(e.dialogPath + "weblink.htm", e.dialogSize["LinkDialog"].w, e.dialogSize["LinkDialog"].h)
        }
        break;
    case "btnImageDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webimage.htm", e.dialogSize["ImageDialog"].w, e.dialogSize["ImageDialog"].h);
        break;
    case "btnYoutubeDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webyoutube.htm", e.dialogSize["YoutubeDialog"].w, e.dialogSize["YoutubeDialog"].h);
        break;
    case "btnTableDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webtable.htm", e.dialogSize["TableDialog"].w, e.dialogSize["TableDialog"].h);
        break;
    case "btnFlashDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webflash.htm", e.dialogSize["FlashDialog"].w, e.dialogSize["FlashDialog"].h);
        break;
    case "btnCharsDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webchars.htm", e.dialogSize["CharsDialog"].w, e.dialogSize["CharsDialog"].h);
        break;
    case "btnSearchDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "websearch.htm", e.dialogSize["SearchDialog"].w, e.dialogSize["SearchDialog"].h);
        break;
    case "btnSourceDialog":
        e.hide();
        changeActiveEditor(e.oName);
        modelessDialogShow(e.dialogPath + "websource.htm", e.dialogSize["SourceDialog"].w, e.dialogSize["SourceDialog"].h);
        break;
    case "btnBookmarkDialog":
        e.hide();
        modelessDialogShow(e.dialogPath + "webbookmark.htm", e.dialogSize["BookmarkDialog"].w, e.dialogSize["BookmarkDialog"].h);
        break;
    case "btnPreview":
        e.hide();
        modelessDialogShow(e.dialogPath + "webpreview.htm", e.dialogSize["Preview"].w, e.dialogSize["Preview"].h);
        break;
    case "btnContentBlock":
        e.hide();
        eval(e.cmdContentBlock);
        break;
    case "btnInternalLink":
        e.hide();
        eval(e.cmdInternalLink);
        break;
    case "btnInternalImage":
        e.hide();
        eval(e.cmdInternalImage);
        break;
    case "btnCustomObject":
        e.hide();
        eval(e.cmdCustomObject);
        break;
    case "btnGuidelines":
        e.runtimeBorder(true);
        break;
    case "btnAbsolute":
        e.makeAbsolute();
        break;
    case "btnLine":
        e.doCmd("InsertHorizontalRule");
        break;
    case "btnRemoveFormat":
        e.doClean();
        break;
    case "btnClearAll":
        e.clearAll();
        break;
    case "btnStyles":
        e.hide();
        e.openStyleSelect();
        break;
    case "btnParagraph":
        e.hide();
        e.selectParagraph();
        break;
    case "btnFontName":
        e.hide();
        e.expandSelection();
        realtimeFontSelect(e.oName);
        break;
    case "btnFontSize":
        e.hide();
        e.expandSelection();
        realtimeSizeSelect(e.oName);
        break;
    case "btnCustomTag":
        e.hide();
        break;
    case "btnQuote":
        e.applyQuote();
        break;
    default:
        for (var i = 0; i < e.arrCustomButtons.length; i++) {
            if (e.arrCustomButtons[i][0] == btn) {
                eval(e.arrCustomButtons[i][1]);
                break
            }
        }
    }
};

function ddAction(tb, id, edt, sfx) {
    var oN = sfx;
    var e = edt;
    var btn = id.substring(0, id.lastIndexOf(oN));
    switch (btn) {
    case "btnPasteClip":
        e.doPaste();
        break;
    case "btnPasteWord":
        if (!e.customDialogShow("PasteWord")) modelessDialogShow(e.scriptPath + "paste_word.htm", 400, 280);
        break;
    case "btnPasteText":
        e.doPasteText();
        break
    }
    var idx = 0;
    if (btn.indexOf("btnParagraphFormatting") != -1) {} else if (btn.indexOf("btnParagraph") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyParagraph("<" + e.arrParagraph[parseInt(idx)][1] + ">")
    } else if (btn.indexOf("btnFontName") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyFontName(e.arrFontName[parseInt(idx)])
    } else if (btn.indexOf("btnFontSize") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.applyFontSize(e.arrFontSize[parseInt(idx)][1])
    } else if (btn.indexOf("btnCustomTag") != -1) {
        idx = btn.substr(btn.indexOf("_") + 1);
        e.insertCustomTag(parseInt(idx))
    }
};

function changeHeight(v) {
    var cH = String(this.height);
    var edtObj = document.getElementById("idArea" + this.oName);
    if (cH.indexOf("%") > -1) {
        cH = edtObj.childNodes[0].offsetHeight - edtObj.rows[0].cells[0].childNodes[0].offsetHeight - (this.useTagSelector ? 20 : 0)
    }
    if (!this.minHeight) this.minHeight = parseInt(cH, 10);
    var newHeight = parseInt(cH, 10) + v;
    this.height = newHeight + "px";
    edtObj.style.height = this.height
};

function _isWordContent(str) {
    return (/msonormal/i.test(str) || /<\\?\?xml[^>]*>/gi.test(str) || /<\/?o:p[^>]*>/gi.test(str) || /<\/?u1:p[^>]*>/gi.test(str) || /<\/?v:[^>]*>/gi.test(str) || /<\/?o:[^>]*>/gi.test(str) || /<\/?st1:[^>]*>/gi.test(str) || /<\/?w:wrap[^>]*>/gi.test(str) || /<\/?w:anchorlock[^>]*>/gi.test(str))
};

function fixWord() {
    var pasteFrame = document.getElementById("idContentWord" + this.oName);
    var idSource = pasteFrame.contentWindow;
    var isWord = _isWordContent(idSource.document.body.innerHTML);
    if (!isWord) {
        this.insertHTML(idSource.document.body.innerHTML);
        return
    }
    for (var i = 0; i < idSource.document.body.all.length; i++) {
        idSource.document.body.all[i].removeAttribute("class", "", 0);
        idSource.document.body.all[i].removeAttribute("className", "", 0);
        idSource.document.body.all[i].removeAttribute("style", "", 0)
    }
    var str = idSource.document.body.innerHTML;
    str = String(str).replace(/<\\?\?xml[^>]*>/gi, "");
    str = String(str).replace(/<\/?o:p[^>]*>/gi, "");
    str = String(str).replace(/<\/?u1:p[^>]*>/gi, "");
    str = String(str).replace(/<\/?v:[^>]*>/gi, "");
    str = String(str).replace(/<\/?o:[^>]*>/gi, "");
    str = String(str).replace(/<\/?st1:[^>]*>/gi, "");
    str = String(str).replace(/<\/?w:wrap[^>]*>/gi, "");
    str = String(str).replace(/<\/?w:anchorlock[^>]*>/gi, "");
    str = String(str).replace(/<\/?SPAN[^>]*>/gi, "");
    str = String(str).replace(/<\/?FONT[^>]*>/gi, "");
    str = String(str).replace(/<\/?STRONG[^>]*>/gi, "");
    str = String(str).replace(/<\/?H1[^>]*>/gi, "");
    str = String(str).replace(/<\/?H2[^>]*>/gi, "");
    str = String(str).replace(/<\/?H3[^>]*>/gi, "");
    str = String(str).replace(/<\/?H4[^>]*>/gi, "");
    str = String(str).replace(/<\/?H5[^>]*>/gi, "");
    str = String(str).replace(/<\/?H6[^>]*>/gi, "");

    str = String(str).replace(/<\/?P[^>]*><\/P>/gi, "");
    var reg = new RegExp(String.fromCharCode(8217), "gi");
    str = String(str).replace(reg, "'");
    str = String(str).replace(/  /gi, " ");
    str = String(str).replace(/\n\n/gi, "\n");
    this.insertHTML(str)
};

function getOuterHTML(node) {
    var sHTML = "";
    switch (node.nodeType) {
    case 1:
        sHTML = "<" + node.nodeName;
        var tagVal = "",
            embedTag = "";
        for (var atr = 0; atr < node.attributes.length; atr++) {
            if (node.attributes[atr].nodeName.substr(0, 4) == "_moz") continue;
            if (node.attributes[atr].nodeValue.substr(0, 4) == "_moz") continue;
            if (node.nodeName == 'TEXTAREA' && node.attributes[atr].nodeName.toLowerCase() == 'value') {
                tagVal = node.attributes[atr].nodeValue
            } else {
                if (node.attributes[atr].nodeName.toLowerCase() == "althtml") {
                    embedTag = node.attributes[atr].nodeValue.replace(/^\s+|\s+$/g, "");
                    embedTag = embedTag.replace(/>$/ig, " \/>").replace(/\/ \/>$/ig, "\/>").replace(/\s*\/>/gi, "></embed>")
                } else {
                    sHTML += ' ' + node.attributes[atr].nodeName + '="' + node.attributes[atr].nodeValue.replace(/"/gi, "'") + '"'
                }
            }
        }
        sHTML += '>';
        if (embedTag != "") sHTML += embedTag;
        if (node.nodeName == 'TEXTAREA') {
            sHTML += tagVal
        } else if (node.nodeName == 'OBJECT') {
            var ch;
            for (var i = 0; i < node.childNodes.length; i++) {
                ch = node.childNodes[i];
                if (ch.nodeType == 1) {
                    if (ch.tagName == "PARAM") {
                        sHTML += "<param name=\"" + ch.name + "\" value=\"" + ch.value.replace(/"/gi, "'") + "\"/>\n"
                    } else if (ch.tagName == "EMBED") {
                        sHTML += getOuterHTML(ch)
                    }
                }
            }
        } else {
            sHTML += node.innerHTML
        }
        sHTML += "</" + node.nodeName + ">";
        break;
    case 8:
        sHTML = "<!" + "--" + node.nodeValue + "--" + ">";
        break;
    case 3:
        sHTML = node.nodeValue;
        break
    }
    return sHTML
};

function customDialogShow(s) {
    for (var j = 0; j < this.customDialog.length; j++) {
        if (this.customDialog[j][0] == s) {
            eval(this.customDialog[j][1]);
            return true
        }
    }
    return false
};

function GetEmoticons() {
    var sHtml = '';
    var arrEmoticons = [
        ["smiley.png", "smiley"],
        ["smiley-lol.png", "lol"],
        ["smiley-confuse.png", "confuse"],
        ["smiley-cool.png", "cool"],
        ["smiley-cry.png", "cry"],
        ["smiley-wink.png", "wink"],
        ["smiley-surprise.png", "surprise"],
        ["smiley-sad.png", "sad"],
        ["smiley-red.png", "red"],
        ["smiley-neutral.png", "neutral"],
        ["smiley-kiss.png", "kiss"],
        ["smiley-mad.png", "mad"],
        ["smiley-money.png", "money"],
        ["smiley-sleep.png", "sleep"],
        ["smiley-yell.png", "yell"],
        ["smiley-roll.png", "roll"],
        ["smiley-grin.png", "grin"],
        ["smiley-razz.png", "razz"],
        ["smiley-sweat.png", "sweat"],
        ["smiley-eek.png", "eek"],
        ["smiley-zipper.png", "zipper"],
        ["heart.png", "love"],
        ["heart-break.png", "heart break"],
        ["light-bulb.png", "idea"]
    ];
    var icPath = this.scriptPath + this.iconPath;
    sHtml += "<div style='width:193px;height:72px;margin:2px;'>";
    for (var j = 0; j < arrEmoticons.length; j++) {
        sHtml += "<img unselectable='on' src='" + icPath + arrEmoticons[j][0] + "' onmouseover='this.style.border=\"#aaaaaa 1px solid\"' onmouseout='this.style.border=\"#ffffff 1px solid\"' style='float:left;margin:1px;cursor:pointer;border:#ffffff 1px solid;padding:2px;' onclick='" + this.oName + ".insertEmoticon(\"" + icPath + arrEmoticons[j][0] + "\",\"" + arrEmoticons[j][1] + "\");' alt='" + arrEmoticons[j][1] + "' />"
    }
    sHtml += "</div>";
    return sHtml
}

function insertEmoticon(img, alt) {
    this.insertHTML("<img src='" + img + "' alt='" + alt + "' />");
    var box = document.getElementById("ddEmoticons" + this.oName);
    box.style.display = "none";
    this.isActive = false
}

function applyQuote() {
    var oEditor = eval("idContent" + this.oName);
    var oSel = oEditor.document.selection.createRange();
    var oQuote = (oSel.parentElement != null ? GetElement(oSel.parentElement(), "BLOCKQUOTE") : GetElement(oSel.item(0), "BLOCKQUOTE"));
    if (oQuote) {
        var oQuote = (oSel.parentElement != null ? GetElement(oSel.parentElement(), "BLOCKQUOTE") : GetElement(oSel.item(0), "BLOCKQUOTE"));
        var oSelRange = oEditor.document.body.createTextRange();
        oSelRange.moveToElementText(oQuote);
        oSel.setEndPoint("StartToStart", oSelRange);
        oSel.setEndPoint("EndToEnd", oSelRange);
        oSel.select();
        this.saveForUndo();
        sHTML = oQuote.innerHTML;
        sHTML = fixPathEncode(sHTML);
        var oTmp = oQuote.parentElement;
        if (oTmp.innerHTML == oQuote.outerHTML) {
            oTmp.innerHTML = "<p>" + sHTML + "</p>";
            fixPathDecode(oEditor);
            var oSelRange = oEditor.document.body.createTextRange();
            oSelRange.moveToElementText(oTmp);
            oSel.setEndPoint("StartToStart", oSelRange);
            oSel.setEndPoint("EndToEnd", oSelRange);
            oSel.select();
            realTime(this.oName);
            this.selectElement(0);
            return
        } else {
            oQuote.outerHTML = "";
            oSel.pasteHTML("<p>" + sHTML + "</p>");
            fixPathDecode(oEditor);
            this.focus();
            realTime(this.oName)
        }
        this.runtimeBorder(false);
        this.runtimeStyles()
    } else {
        this.applyParagraph('<pre>');
        var oPre = (oSel.parentElement != null ? GetElement(oSel.parentElement(), "PRE") : GetElement(oSel.item(0), "PRE"));
        var oSelRange = oEditor.document.body.createTextRange();
        oSelRange.moveToElementText(oPre);
        oSel.setEndPoint("StartToStart", oSelRange);
        oSel.setEndPoint("EndToEnd", oSelRange);
        oSel.select();
        this.saveForUndo();
        sHTML = oPre.innerHTML;
        sHTML = fixPathEncode(sHTML);
        var oTmp = oPre.parentElement;
        if (oTmp.innerHTML == oPre.outerHTML) {
            oTmp.innerHTML = "<blockquote>" + sHTML + "</blockquote>";
            fixPathDecode(oEditor);
            var oSelRange = oEditor.document.body.createTextRange();
            oSelRange.moveToElementText(oTmp);
            oSel.setEndPoint("StartToStart", oSelRange);
            oSel.setEndPoint("EndToEnd", oSelRange);
            oSel.select();
            realTime(this.oName);
            this.selectElement(0);
            return
        } else {
            oPre.outerHTML = "";
            oSel.pasteHTML("<blockquote>" + sHTML + "</blockquote>");
            fixPathDecode(oEditor);
            this.focus();
            realTime(this.oName)
        }
        this.runtimeBorder(false);
        this.runtimeStyles()
    }
}