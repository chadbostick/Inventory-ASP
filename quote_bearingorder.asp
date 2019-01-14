<%option explicit

dim cnnSQLConnection
dim rsGetQuoteBearingsByQuoteId, cmdGetQuoteBearingsByQuoteId
dim intRowNumber

set cnnSQLConnection											= nothing
set rsGetQuoteBearingsByQuoteId						= nothing
set cmdGetQuoteBearingsByQuoteId					= nothing

set cnnSQLConnection											= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteBearingsByQuoteId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetQuoteBearingsByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteBearingsByQuoteId.CommandText = "{call spGetQuoteBearingsByQuoteId( ?,?,? ) }"

cmdGetQuoteBearingsByQuoteId.Parameters.Append cmdGetQuoteBearingsByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteBearingsByQuoteId.Parameters.Append cmdGetQuoteBearingsByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4)
cmdGetQuoteBearingsByQuoteId.Parameters.Append cmdGetQuoteBearingsByQuoteId.CreateParameter( "@IsPrint", adTinyInt, adParamInput, 1)

cmdGetQuoteBearingsByQuoteId.Parameters( "@QuoteId" )	= Request.QueryString( "intQuoteId" )
cmdGetQuoteBearingsByQuoteId.Parameters( "@IsPrint" )	= 1

on error resume next
set rsGetQuoteBearingsByQuoteId = cmdGetQuoteBearingsByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
		
on error goto 0

if not( rsGetQuoteBearingsByQuoteId is nothing ) then
		if( rsGetQuoteBearingsByQuoteId.State = adStateOpen ) then
			if( not rsGetQuoteBearingsByQuoteId.EOF ) then%>
<HTML>
<TITLE>Bearing Order for Work Order&nbsp;&nbsp;#<%=rsGetQuoteBearingsByQuoteId( "WorkOrderNumber" )%></TITLE>
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
      <tr><td class="title">BEARING ORDER</td><td class="title" align=right><%=rsGetQuoteBearingsByQuoteId( "WorkOrderNumber" )%></td></tr>
    </table></td></tr>
    
    <%set rsGetQuoteBearingsByQuoteId = rsGetQuoteBearingsByQuoteId.NextRecordset
			if not( rsGetQuoteBearingsByQuoteId is nothing ) then
				if( rsGetQuoteBearingsByQuoteId.State = adStateOpen ) then
					if( not rsGetQuoteBearingsByQuoteId.EOF ) then
						intRowNumber = 1%>
						
						<tr><td>&nbsp;</td></tr>
						<tr><td>&nbsp;</td></tr>
    
    <tr><td><table>
			<%while( not rsGetQuoteBearingsByQuoteId.EOF )%>
      <tr><td class="label">CEN BRG CODE:&nbsp;</td><td><%=rsGetQuoteBearingsByQuoteId( "CenBearingCode" )%></td>
        <td class="label">PRODUCT ID:&nbsp;</td><td><%=rsGetQuoteBearingsByQuoteId( "ProductId" )%></td></tr>
      <tr><td class="label">&nbsp;&nbsp;BEARING&nbsp;<%=intRowNumber%>:&nbsp;</td><td colspan=3><%=rsGetQuoteBearingsByQuoteId( "BearingDescription" )%></td></tr>
      <tr><td class="label">&nbsp;&nbsp;SOURCE:&nbsp;</td><td colspan=3><%=rsGetQuoteBearingsByQuoteId( "SupplierName" )%></td></tr>
      <tr><td class="label">&nbsp;&nbsp;CSS PRICE:&nbsp;</td><td><%=FormatCurrency( rsGetQuoteBearingsByQuoteId( "CSSPrice" ), 2 )%></td>
        <td class="label">PRICE:&nbsp;</td><td><%=FormatCurrency( rsGetQuoteBearingsByQuoteId( "BearingCost" ), 2 )%></td>
      <tr><td>&nbsp;</td></tr>
      <%rsGetQuoteBearingsByQuoteId.MoveNext
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

if not( rsGetQuoteBearingsByQuoteId is nothing ) then
	if( rsGetQuoteBearingsByQuoteId.State = adStateOpen ) then
		rsGetQuoteBearingsByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuoteBearingsByQuoteId				= nothing
set cmdGetQuoteBearingsByQuoteId			= nothing
set cnnSQLConnection									= nothing
%>