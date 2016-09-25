<html>
<head>
<title>Management Options</title>
<body bgcolor="#000000" background="satin.jpg">

<a href="/default.htm"><img src="title1.jpg" border=0></a>

<center>
<OBJECT ID="Button1" WIDTH=100 HEIGHT=18
   CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667"
   CODEBASE="ikbutton.ocx">
	<PARAM NAME=Caption VALUE="Order">
</OBJECT>
<img src="space_2h.gif">

<OBJECT ID="Button2" WIDTH=100 HEIGHT=18
   CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667"
   CODEBASE="ikbutton.ocx">
	<PARAM NAME=Caption VALUE="Update">
</OBJECT>
<img src="space_2h.gif">

<OBJECT ID="Button3" WIDTH=100 HEIGHT=18
   CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667"
   CODEBASE="ikbutton.ocx">
	<PARAM NAME=Caption VALUE="Some">
</OBJECT>
<img src="space_2h.gif">

<OBJECT ID="Button4" WIDTH=100 HEIGHT=18
   CLASSID="CLSID:B10CBD8D-F9B6-11CF-9B38-0080AD11B667"
   CODEBASE="ikbutton.ocx">
	<PARAM NAME=Caption VALUE="Schedule">
</OBJECT>
<img src="space_2h.gif">

<OBJECT ID="Menu1" WIDTH=100 HEIGHT=18
   CLASSID="CLSID:F5131C24-E56D-11CF-B78A-444553540000"
   CODEBASE="ikmenu.ocx">
</OBJECT>

<SCRIPT LANGUAGE="VBScript">
<!--
Sub Menu1_OnClick(id)
   On Error Resume Next
   select case id
      case "OrderN"
         parent.frames("body").Location.Href="order_new.asp"
      case "OrderV"
         parent.frames("body").Location.Href="order_view.asp"
   
      case "ManEA"
         parent.frames("body").Location.Href="./employees/add.asp"
      case "ManEE"
         parent.frames("body").Location.Href="./employees/edit.asp"
      case "ManER"
         parent.frames("body").Location.Href="./employees/remove.asp"

      case "ManCA"
         parent.frames("body").Location.Href="./customers/add.asp"
      case "ManCE"
         parent.frames("body").Location.Href="./customers/edit.asp"
      case "ManCR"
         parent.frames("body").Location.Href="./customers/remove.asp"

      case "ManDA"
         parent.frames("body").Location.Href="./departments/add.asp"
      case "ManDE"
         parent.frames("body").Location.Href="./departments/edit.asp"
      case "ManDR"
         parent.frames("body").Location.Href="./departments/remove.asp"

      case "ManVA"
         parent.frames("body").Location.Href="./vendors/add.asp"
      case "ManVE"
         parent.frames("body").Location.Href="./vendors/edit.asp"
      case "ManVR"
         parent.frames("body").Location.Href="./vendors/remove.asp"

      case "ManPA"
         parent.frames("body").Location.Href="./publishers/add.asp"
      case "ManPE"
         parent.frames("body").Location.Href="./publishers/edit.asp"
      case "ManPR"
         parent.frames("body").Location.Href="./publishers/remove.asp"       
   end select
end sub

Sub Button1_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()
	call Menu1.AddItem("OrderN", "New", "", "Order")
   call Menu1.AddItem("OrderR", "Receive", "", "Order")
   call Menu1.Popup(Button1.Left + Button1.Width, Button1.Top)
End Sub

Sub Button2_MouseDown(Button, Shift, X, Y)
	call Menu1.RemoveAllItems()

	call Menu1.AddItem("ManE", "Employees", "", "Man")
   call Menu1.AddItem("ManEA", "Add", "", "ManE")
	call Menu1.AddItem("ManEE", "Edit", "", "ManE")
   call Menu1.AddItem("ManER", "Remove", "", "ManE")

	call Menu1.AddItem("ManC", "Customers", "", "Man")
   call Menu1.AddItem("ManCA", "Add", "", "ManC")
	call Menu1.AddItem("ManCE", "Edit", "", "ManC")
   call Menu1.AddItem("ManCR", "Remove", "", "ManC")

	call Menu1.AddItem("ManD", "Departments", "", "Man")
   call Menu1.AddItem("ManDA", "Add", "", "ManD")
	call Menu1.AddItem("ManDE", "Edit", "", "ManD")
   call Menu1.AddItem("ManDR", "Remove", "", "ManD")

	call Menu1.AddItem("ManV", "Vendors", "", "Man")
   call Menu1.AddItem("ManVA", "Add", "", "ManV")
	call Menu1.AddItem("ManVE", "Edit", "", "ManV")
   call Menu1.AddItem("ManVR", "Remove", "", "ManV")

	call Menu1.AddItem("ManP", "Publishers", "", "Man")
   call Menu1.AddItem("ManPA", "Add", "", "ManP")
	call Menu1.AddItem("ManPE", "Edit", "", "ManP")
   call Menu1.AddItem("ManPR", "Remove", "", "ManP")

	call Menu1.AddItem("ManDB", "Products", "", "Man")
   call Menu1.AddItem("ManDBA", "Add", "", "ManDB")
	call Menu1.AddItem("ManDBE", "Edit", "", "ManDB")
   call Menu1.AddItem("ManDBR", "Remove", "", "ManDB")
   call Menu1.AddItem("ManDBI", "Import", "", "ManDB")
   call Menu1.AddItem("ManDBIN", "NACSCORP", "", "ManDBI")
   call Menu1.AddItem("ManDBII", "Ingram Micro", "", "ManDBI")
   call Menu1.AddItem("ManDBID", "Douglas Stewart", "", "ManDBI")

   call Menu1.Popup(Button2.Left + Button2.Width, Button2.Top)
End Sub

Sub Button4_MouseDown(Button, Shift, X, Y)
   parent.frames("body").Location.Href="calander.asp"
End Sub

-->

</SCRIPT>

<font face="arial narrow, arial" color="#ffffff">
<b>Computer<br>
Department<br>
Management<br>
System<br></b>
<font size=2>Preview Version</font><br>
<font size=1>Copyright 1996<br>Jason Abbott<br>
</font>
</font>
</center>
</BODY>
</HTML>

