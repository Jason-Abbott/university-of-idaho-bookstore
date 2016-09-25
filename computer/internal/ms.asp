<!-----------------------------------------
--File: menu.asp
--Date: 10/22/96
--Author: Dave and Bryan Baker
--Beverly Hills Software
--http://www.bhs.com
--(C) 1996 Microsoft Corporation
------------------------------------------>
<HTML>
<HEAD>
<TITLE>Menu</TITLE>
</HEAD>
<BODY BGCOLOR="#FFCC00" Link="#000080" VLINK="#000080">
<CENTER>
<IMG SRC="http://iisa.microsoft.com/iis3/images/activ.gif" ALT="select an activity" WIDTH="120" HEIGHT="18" BORDER="0" VSPACE="0" HSPACE=0><BR>


<OBJECT ID="Button1" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="Learn about IIS">
</OBJECT><BR>
<IMG SRC="http://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR>
<OBJECT ID="Button2" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="Evaluating">
</OBJECT>
<BR>
<IMG SRC="http://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR>
<OBJECT ID="Button3" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="Web Showcase">
</OBJECT>
<BR>
<IMG SRC="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR>
<OBJECT ID="Button4" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="Free Software">
</OBJECT>
<BR>
<IMG SRC="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR>
<OBJECT ID="Button5" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="Support & Training">
</OBJECT>
<BR>
<IMG SRC="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR><OBJECT ID="Button6" WIDTH=130 HEIGHT=18 CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667" CODEBASE="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/ikcntrls.cab#version=1,0,0,5">
	<PARAM NAME=Caption VALUE="IIS News">
</OBJECT>
<BR>
<IMG SRC="http://iisa.microsoft.comhttp://iisa.microsoft.com/iis3/images/spacer.gif" ALT=" "WIDTH=130 HEIGHT=2 BORDER=0><BR>
<OBJECT ID="Menu1" WIDTH=100 HEIGHT=51 CLASSID="CLSID:F5131C24-E56D-11CF-B78A-444553540000"	CODEBASE="http://iisa.microsoft.com/iis3/ikcntrls.cab#Version=1,0,0,8"></OBJECT>
<SCRIPT LANGUAGE="VBScript">
<!--
Sub Menu1_OnClick(id)
		On Error Resume Next
			select case id
				case "LearnIISA01"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/iis3beta/"
				case "LearnIISA02"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/denali/default.asp"
				case "LearnIISA03"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/Netshow/"
				case "LearnIISA04"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/FrontPage/"
				case "LearnIISA05"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/IndexServer/"
				case "LearnIISA06"
				parent.frames("content").Location.Href="http://iisb.microsoft.com/advworks/default.asp"
				case "LearnIISB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/iis2/default.htm"
				case "LearnIISC"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/learnaboutiis/iis1/default.htm"
				case "EvaluatingA"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/evaluating/review/default.htm"
				case "EvaluatingB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/evaluating/compare/default.asp"
				case "ShowcaseA"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/webshowcase/Coolsites/default.asp"
				case "ShowcaseB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/webshowcase/iiscommunity/default.htm"
				case "ShowcaseC"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/webshowcase/casestudies/default.htm"
				case "ShowcaseD"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/webshowcase/testimonials/default.htm"
				case "ShowcaseE"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/webshowcase/ontheweb/default.asp"
				case "FreeSoftwareA"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/freesoftware/DownloadIIS3"
				case "FreeSoftwareB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/freesoftware/IIS2"
				case "FreeSoftwareC"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/freesoftware/DownloadIIS1"
				case "SupportA"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/support/online/"
				case "SupportB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/support/training/default.htm"
				case "IISNewsA"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/iisnews/hotnews/default.asp"
				case "IISNewsB"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/iisnews/press/"
				case "IISNewsC"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/iisnews/IndustryArticles/"
				case "IISNewsD"
				parent.frames("content").Location.Href="http://iisa.microsoft.com/iis3/iisnews/events/"
			end select

	end sub
	
Sub Button1_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("LearnIISA", "IIS 3.0 Beta", "", "LearnIIS")
	call Menu1.AddItem("LearnIISA01", "Overview", "", "LearnIISA")
	call Menu1.AddItem("LearnIISA02", "Active Server Pages", "", "LearnIISA")
	call Menu1.AddItem("LearnIISA05", "Index Server", "", "LearnIISA")
	call Menu1.AddItem("LearnIISA03", "NetShow", "", "LearnIISA")
	call Menu1.AddItem("LearnIISA04", "Front Page 97 Extensions", "", "LearnIISA")

	call Menu1.AddItem("LearnIISA06", "Run the Demo", "", "LearnIISA")
	call Menu1.AddItem("LearnIISB", "IIS 2.0", "", "LearnIIS")
	call Menu1.AddItem("LearnIISC", "IIS 1.0", "", "LearnIIS")
	call Menu1.Popup(Button1.Left + Button1.Width, Button1.Top)
End Sub
Sub Button2_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("EvaluatingA", "Reviewer's Guide", "", "Evaluating")
	call Menu1.AddItem("EvaluatingB", "How Does IIS Compare?", "", "Evaluating")
	call Menu1.Popup(Button2.Left + Button2.Width, Button2.Top)
End Sub
Sub Button3_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("ShowcaseA", "Cool IIS Sites", "", "Showcase")
'	call Menu1.AddItem("ShowcaseB", "IIS Community", "", "Showcase")
	call Menu1.AddItem("ShowcaseC", "Case Studies", "", "Showcase")
	call Menu1.AddItem("ShowcaseD", "Testimonials", "", "Showcase")
	call Menu1.AddItem("ShowcaseE", "Microsoft on the Web", "", "Showcase")
	call Menu1.Popup(Button3.Left + Button3.Width, Button3.Top)
End Sub
Sub Button4_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("FreeSoftwareA", "Download IIS 3.0 Beta", "", "FreeSoftware")
	call Menu1.AddItem("FreeSoftwareB", "IIS 2.0", "", "FreeSoftware")
	call Menu1.AddItem("FreeSoftwareC", "Download IIS 1.0", "", "FreeSoftware")
	call Menu1.Popup(Button4.Left + Button4.Width, Button4.Top)
End Sub
Sub Button5_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("SupportA", "On-line Support", "", "Support")
	call Menu1.AddItem("SupportB", "Training", "", "Support")
	call Menu1.Popup(Button5.Left + Button5.Width, Button5.Top)
End Sub
Sub Button6_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("IISNewsA", "Hot News", "", "IISNews")
	call Menu1.AddItem("IISNewsB", "Press Announcements", "", "IISNews")
	call Menu1.AddItem("IISNewsC", "Industry Articles", "", "IISNews")
	call Menu1.AddItem("IISNewsD", "Events", "", "IISNews")
	call Menu1.Popup(Button6.Left + Button6.Width, Button6.Top)
End Sub
-->
</SCRIPT>


</ceNTER>
</BODY>
</HTML>

