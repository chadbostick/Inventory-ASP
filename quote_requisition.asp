<%option explicit

dim cnnSQLConnection
dim rsGetQuoteSubWorkByQuoteId, cmdGetQuoteSubWorkByQuoteId
dim intRowNumber, strFreightLBS, intFreightChargeSub

set cnnSQLConnection										= nothing
set rsGetQuoteSubWorkByQuoteId					= nothing
set cmdGetQuoteSubWorkByQuoteId					= nothing

set cnnSQLConnection										= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteSubWorkByQuoteId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetQuoteSubWorkByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteSubWorkByQuoteId.CommandText = "{call spGetQuoteSubWorkByQuoteId( ?,?,? ) }"

cmdGetQuoteSubWorkByQuoteId.Parameters.Append cmdGetQuoteSubWorkByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteSubWorkByQuoteId.Parameters.Append cmdGetQuoteSubWorkByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuoteSubWorkByQuoteId.Parameters.Append cmdGetQuoteSubWorkByQuoteId.CreateParameter( "@IsPrint", adTinyInt, adParamInput, 1 )

cmdGetQuoteSubWorkByQuoteId.Parameters( "@QuoteId" ) = Request.QueryString( "intQuoteId" )
cmdGetQuoteSubWorkByQuoteId.Parameters( "@IsPrint" ) = 1

on error resume next
set rsGetQuoteSubWorkByQuoteId = cmdGetQuoteSubWorkByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0

if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
	if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
		if( not rsGetQuoteSubWorkByQuoteId.EOF ) then%>
<HTML>
<HEAD>
  <TITLE>Requisition Form for Work Order&nbsp;&nbsp;#<%=rsGetQuoteSubWorkByQuoteId( "WorkOrderNumber" )%></TITLE>
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
      <tr><td class="title">REQUISITION FORM</td><td class="title" align=right><%=rsGetQuoteSubWorkByQuoteId( "WorkOrderNumber" )%></td></tr>
    </table></td></tr>
    
    <tr><td><table width="640">
      <tr><td width="400"><%=FormatDateTime( Date(), vbLongDate )%></td><td width="240">DELIVERY:&nbsp;<%=rsGetQuoteSubWorkByQuoteId( "ExpDeliveryDate" )%></td></tr>
    </table></td></tr>
    
    <tr><td>&nbsp;</td></tr>
    
    <tr><td><%=rsGetQuoteSubWorkByQuoteId( "Customer" )%></td></tr>
    <tr><td><%=rsGetQuoteSubWorkByQuoteId( "SpindleType" )%></td></tr>
    
    <%strFreightLBS				= rsGetQuoteSubWorkByQuoteId( "FreightLBS" )
			intFreightChargeSub	= CInt( rsGetQuoteSubWorkByQuoteId( "FreightChargeSub" ) )
			
			set rsGetQuoteSubWorkByQuoteId = rsGetQuoteSubWorkByQuoteId.NextRecordset
			if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
				if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
					if( not rsGetQuoteSubWorkByQuoteId.EOF ) then
						intRowNumber = 1%>
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    
    <tr><td><table>
			<%while( not rsGetQuoteSubWorkByQuoteId.EOF )%>
      <tr><td class="label">SUBWORK&nbsp;<%=intRowNumber%>)</td>
        <td><%=rsGetQuoteSubWorkByQuoteId( "SubWorkDescription" )%></td></tr>
      <tr><td>SOURCE:</td>
        <td><%=rsGetQuoteSubWorkByQuoteId( "SupplierName" )%></td></tr>
      <tr><td>&nbsp;</td></tr>
      <%rsGetQuoteSubWorkByQuoteId.MoveNext
				intRowNumber = intRowNumber + 1
      wend%>
    </table></td></tr>
    <%	end if
			end if
		end if%>
		
		<tr><td>&nbsp;</td></tr>
    
    <tr><td class="label">SHIPPING INFO:&nbsp;<%=strFreightLBS%><%if( intFreightChargeSub > 0 ) then Response.Write( "&nbsp;&nbsp;>>>&nbsp;" + FormatCurrency( intFreightChargeSub, 2 ) ) end if%></td></tr>
    
    <tr><td>&nbsp;</td></tr>
    <tr><td>&nbsp;</td></tr>
    
    <tr><td align=right>INITIALS:______________________</td>
    
  </table>
</BODY>
</HTML>
<%	end if
	end if
end if

if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
	if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
		rsGetQuoteSubWorkByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuoteSubWorkByQuoteId				= nothing
set cmdGetQuoteSubWorkByQuoteId				= nothing
set cnnSQLConnection									= nothing
%>