<%option explicit

dim cnnSQLConnection
dim rsGetQuoteLaborByQuoteId, cmdGetQuoteLaborByQuoteId
dim intQuoteId

set cnnSQLConnection					= nothing
set rsGetQuoteLaborByQuoteId	= nothing
set cmdGetQuoteLaborByQuoteId	= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteLaborByQuoteId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intQuoteId = Request.QueryString( "intQuoteId" )

cmdGetQuoteLaborByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteLaborByQuoteId.CommandText = "{call spGetQuoteLaborByQuoteId( ?,? ) }"
cmdGetQuoteLaborByQuoteId.Parameters.Append cmdGetQuoteLaborByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteLaborByQuoteId.Parameters.Append cmdGetQuoteLaborByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuoteLaborByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetQuoteLaborByQuoteId = cmdGetQuoteLaborByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0

if not( rsGetQuoteLaborByQuoteId is nothing ) then
		if( rsGetQuoteLaborByQuoteId.State = adStateOpen ) then
			if( not rsGetQuoteLaborByQuoteId.EOF ) then%>		
<HTML>
<TITLE>Time Tracking for Work Order&nbsp;&nbsp;#<%=rsGetQuoteLaborByQuoteId( "WorkOrderNumber" )%></TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="reports_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; TEXT-DECORATION: underline }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder }
	.smalllabel { FONT-SIZE: 70%; FONT-WEIGHT: bolder; }
	.timetable { PADDING-BOTTOM: 10px }
  </STYLE>  
</HEAD>
<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width="640">
    <tr><td><table width=640>
    <tr><td align=left class="label">W.O. #<%=rsGetQuoteLaborByQuoteId( "WorkOrderNumber" )%></td>
      <td align=right class="label">Assembly Started: ____________</td></tr>
    <tr><td>&nbsp;</td>
      <td align=right class="label">Assembly Finished: ___________</td></tr>
    </table></td></tr>
      
    <tr><td class="title" align=middle>Centerline Time Tracking</td></tr>    
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Disassembly & Evaluation: <%=rsGetQuoteLaborByQuoteId( "DisassemblyEvaluation" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "HoursDisassembly" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Clean and Inspect: <%=rsGetQuoteLaborByQuoteId( "CleanAndInspect" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "HoursCleanAndInspect" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Spacer Preparation: <%=rsGetQuoteLaborByQuoteId( "InhouseGrinding" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "GrindingHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Balancing: <%=rsGetQuoteLaborByQuoteId( "Balancing" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "BalancingHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Electrical Work: <%=rsGetQuoteLaborByQuoteId( "ElectricalWork" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "ElectricalHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Bearing Preparation: <%=rsGetQuoteLaborByQuoteId( "GreaseBearings" )%></td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "GreaseHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Assembly & Test Run:</td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "AssemblyAndTestHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
    
    <tr><td><table width=640>
      <tr><td class="label" width="370">Miscellaneous Work Needed:</td>
        <td class="smalllabel" width="120" align="right">Est. Time: <u>&nbsp;<%=rsGetQuoteLaborByQuoteId( "MscWorkHours" )%>&nbsp;</u> hrs.</td>
        <td class="smalllabel" width="150" align="right">Actual Time: ______ hrs.</td></tr>
    </table></td></tr>
    <tr><td class="timetable" width="640"><table>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
      <tr><td width="160">Technician:______________</td>
        <td width="130">Start:____________</td>
        <td width="130" align="right">End:____________</td>
        <td width="220" align="right">Notes:___________________________</td></tr>
    </table></td></tr>
  </table>
</BODY>
</HTML>
<%	end if
	end if
end if

if not( rsGetQuoteLaborByQuoteId is nothing ) then
	if( rsGetQuoteLaborByQuoteId.State = adStateOpen ) then
		rsGetQuoteLaborByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuoteLaborByQuoteId	= nothing
set cmdGetQuoteLaborByQuoteId	= nothing
set cnnSQLConnection					= nothing
%>
