<!--#INCLUDE FILE="header.asp"-->

      </td>
   </table>

<%

if Not (Request.Form("description") = "") Then
   aryDesc = Split(Request.Form("description"))
   query = "(description LIKE '%" + aryDesc(0) + "%' "
   for i = 1 to UBound(aryDesc,1)
      query = query + Request.Form("comb1") + " description LIKE '%" + aryDesc(i) + "%'"
      Next
   query = query + ")"
End if

if Not (Request.Form("publisher") = "") Then
   aryPubl = Split(Request.Form("publisher"))
   if (Request.Form("description") = "") Then
      query = "(publisher LIKE '%" + aryPubl(0) + "%' "
   else
      query = query + " and (publisher LIKE '%" + aryPubl(0) + "%' "      
   end if
   for i = 1 to UBound(aryPubl,1)
      query = query + Request.Form("comb2") + " publisher LIKE '%" + aryPubl(i) + "%'"
      Next
   query = query + ")"
End if

query = "SELECT pn, publisher, description, sell, academic, software FROM items WHERE " + query

Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "store"
Set RS = Conn.Execute(query)

%>

Academically discounted* items are in bold.
<p>
<center>
<table cellspacing=1 cellpadding=3>
<tr>
   <td align=center bgcolor="a6c5e5"><font face="arial"><b>Description</b></font></td>
   <td align=center bgcolor="a6c5e5"><font face="arial"><b>Price</b></font></td>
   
<%while not RS.eof%>

<tr>
   <td align=right><font face="arial"><font size=1 color="#555588">
      <%=RS("publisher")%></font><br>
      <% if (RS("academic") = -1) Then %><b><% End if %>
      <font size=2><a href="detail.asp?pn=<%=RS("pn")%>&type=<%=RS("software")%>">
      <%=RS("description")%></a>
      </font></font>
   </td>
   <td bgcolor="#fcf4b2" align=right>
   <font size=2>$<%=RS("sell")%></font></td>

<%

RS.MoveNext
wend
RS.close
Conn.close

%>

</table>
</center>
<p>
<font size=2>
All prices are subject to change without notice.  Our product database comes from
multiple vendors so there's a potential for finding multiple copies of a title at
slightly different prices.  While I work to resolve this make sure that we order
the lower priced copy by giving us the distributor's part number (available by
clicking on the product description) for order.
<p>
*Only University of Idaho students, faculty and staff are eligible to purchase
academically discounted software (in bold).  All other titles are available to
the general public.

<!--#INCLUDE FILE="../footer.htm"-->
