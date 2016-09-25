<!--#INCLUDE FILE="computer/header.asp"-->

      </td>
   </table>

<%

thing = Request.QueryString("pn")
query = "SELECT pn, vendor, publisher, pub_pn, description, details, updated, sell, mac, pc FROM items WHERE pn = '" + thing + "'"

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "store"
Set RS = Conn.Execute(query)

%>

<center>
<table cellspacing=1 cellpadding=5>
<tr>
   <td colspan=4 bgcolor="a6c5e5"><font face="arial"><font size=4><b><%=RS("description")%></font></b>
   <br><font size=2><%=RS("details")%></font></font></td>
<tr>
   <td colspan=4 align=right>from <%=RS("publisher")%></td>
<tr>
   <td colspan=2 align=center bgcolor="#eeeeee"><font face="arial">Part Numbers</font></td>
   <td colspan=2 align=center bgcolor="#eeeeee"><font face="arial">Platforms</font></td>
<tr>
   <td align=right>Manufacturer's:</td>
   <td><%=RS("pub_pn")%></td>
   <td align=right>PC:</td>

<tr>
   <td align=right>Distributor's:</td>
   <td><%=RS("pn")%></td>
   <td align=right>Macintosh:</td>

<tr>
   <td colspan=2><font size=1 face="arial">Last updated <%=RS("updated")%></font></td>
   <td colspan=2 align=right><font size=4 face="arial">$<%=RS("sell")%></font></td>

</table>
</center>

<%
RS.close
Conn.close
%>
