<%option explicit

dim cnnSQLConnection
dim rsGetQuotePartsByQuoteId, cmdGetQuotePartsByQuoteId
dim intRowNumber

set cnnSQLConnection									= nothing
set rsGetQuotePartsByQuoteId					= nothing
set cmdGetQuotePartsByQuoteId					= nothing

set cnnSQLConnection									= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuotePartsByQuoteId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetQuotePartsByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuotePartsByQuoteId.CommandText = "{call spGetQuotePartsByQuoteId( ?,?,? ) }"

cmdGetQuotePartsByQuoteId.Parameters.Append cmdGetQuotePartsByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuotePartsByQuoteId.Parameters.Append cmdGetQuotePartsByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuotePartsByQuoteId.Parameters.Append cmdGetQuotePartsByQuoteId.CreateParameter( "@IsPrint", adTinyInt, adParamInput, 1 )

cmdGetQuotePartsByQuoteId.Parameters( "@QuoteId" ) = Request.QueryString( "intQuoteId" )
cmdGetQuotePartsByQuoteId.Parameters( "@IsPrint" ) = 1

on error resume next
set rsGetQuotePartsByQuoteId = cmdGetQuotePartsByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0

if not( rsGetQuotePartsByQuoteId is nothing ) then
	if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
		if( not rsGetQuotePartsByQuoteId.EOF ) then%>
<HTML>
<TITLE>Parts order for Work Order&nbsp;&nbsp;#<%=rsGetQuotePartsByQuoteId( "WorkOrderNumber" )%></TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="reports_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 100%; }
	.title { FONT-SIZE: 120%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder; }
  </STYLE>  
</HEAD>

<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width="640">
    <tr><td><table width="640">
      <tr><td class="title">PARTS ORDER</td><td class="title" align=right><%=rsGetQuotePartsByQuoteId( "WorkOrderNumber" )%></td></tr>
    </table></td></tr>
    
    <%set rsGetQuotePartsByQuoteId = rsGetQuotePartsByQuoteId.NextRecordset
			if not( rsGetQuotePartsByQuoteId is nothing ) then
				if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
					if( not rsGetQuotePartsByQuoteId.EOF ) then
						intRowNumber = 1%>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    
    <tr><td><table>
			<%while( not rsGetQuotePartsByQuoteId.EOF )%>
      <tr><td class="label">ITEM&nbsp;<%=intRowNumber%>:&nbsp;</td><td><%=rsGetQuotePartsByQuoteId( "PartNumber" )%></td>
      <tr><td class="label">SOURCE:&nbsp;</td><td><%=rsGetQuotePartsByQuoteId( "SupplierName" )%></td></tr>
					<td class="label">PRODUCT ID:&nbsp;</td><td><%=rsGetQuotePartsByQuoteId( "ProductId" )%></td></tr>
      <tr><td>&nbsp;</td></tr>
      <%rsGetQuotePartsByQuoteId.MoveNext
				intRowNumber = intRowNumber + 1
      wend%>
    </table></td></tr>
    <%	end if
			end if
		end if%>    
   <tr><td>&nbsp;</td></tr>
    
	<tr><td><table>
	  <tr><td class="label">DATE ORDERED:_______/_______/___________</td>
	    <td class="label">ORDERED BY:_______/_______/___________</td></tr>
	  <tr><td colspan=2>&nbsp;</td></tr>
	  <tr><td class="label">HOW SHIPPED:__________________________</td>
	    <td class="label">EXP. DELIVERY:_________________________</td></tr>
	</table></td></tr>
    
  </table>
</BODY>
</HTML>
<%	end if
	end if
end if

if not( rsGetQuotePartsByQuoteId is nothing ) then
	if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
		rsGetQuotePartsByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuotePartsByQuoteId					= nothing
set cmdGetQuotePartsByQuoteId					= nothing
set cnnSQLConnection									= nothing
%>