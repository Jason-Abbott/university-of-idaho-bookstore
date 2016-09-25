<table cellspacing=3 cellpadding=0 bgcolor="#fcf4b2">
<tr>

<% dim contact(2), phone, email

Select Case Request.ServerVariables("PATH_INFO")
   Case "/computer/resale.asp" , "/computer/search.asp"
      contact(0) = "Erin Adams"
      contact(1) = "<a href='http://www.cs.uidaho.edu/~harr9446'>Mike Harrison</a>"
      contact(2) = "<a href='http://www.uidaho.edu/~skirch/'>Sam Kirchmeier</a>"
      phone = "5518"
      email = "uipcstore"
   Case "/computer/repair.asp"
      contact(0) = "<a href='http://www.uidaho.edu/~jabbott'>Jason Abbott</a>"
      contact(1) = "<a href='http://www.uidaho.edu/~benk/'>Ben Kirchmeier</a>"
      phone = "5518"
      email = "uipcstore"
   Case "/computer/license.asp"
      contact(0) = "Matt Dessert"
      contact(1) = "<a href='http://www.uidaho.edu/~skirch/'>Sam Kirchmeier</a>"
      phone = "5518"
      email = "license"
   Case "/books/packets.asp"
      contact(0) = "Patty Carscallen"
      phone = "2517"
      email = "pattyc"
   Case "/books/buyback.asp"
      contact(0) = "Larry Martin"
      phone = "7038"
      email = "larry"
   Case Else
      phone = "6469"
      email = "uibooks"
End Select

if contact(0) = "" then %>

   <td valign=top align=right valign=top><font size=2 face="Arial"><b>Location:</b></font></td>
   <td><font size=2 face="Arial">Deakin Street<br>Moscow, ID  83844</td>
<tr>
   <td valign=top align=right><font size=2 face="Arial"><b>Hours:</b></font></td>
   <td><font size=2 face="Arial">7:30-5:30 M-F,<br>9:00-4:00 Sat.</td>

<% else %>

   <td valign=top align=right valign=top><font size=2 face="Arial"><b>Contact(s):</b></font></td>
   <td><font size=2 face="Arial">

   <% for i = 0 to UBound(contact,1) %>
   <%= contact(i) %><br>
   <% Next %>
   </td>

<% end if %>

<tr>
   <td valign=top align=right><font size=2 face="Arial"><b>Phone:</b></font></td>
   <td><font size=2 face="Arial">(208)�885-<%= phone %></td>
<tr>
   <td valign=top align=right><font size=2 face="Arial"><b>FAX:</b></font></td>
   <td><font size=2 face="Arial">(208)�885-5953</td>
<tr>
   <td valign=top align=right><font size=2 face="Arial"><b>e-mail:</b></font></td>
   <td><font size=2 face="Arial"><a href="mailto:<%= email %>@uidaho.edu"><%= email %>@uidaho.edu</td>
</table>
