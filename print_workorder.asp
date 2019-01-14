<%option explicit

dim cnnSQLConnection
dim rsGetWorkOrderByWorkOrderId, cmdGetWorkOrderByWorkOrderId
dim intWorkOrderId

set cnnSQLConnection								= nothing
set rsGetWorkOrderByWorkOrderId			= nothing
set cmdGetWorkOrderByWorkOrderId		= nothing

set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetWorkOrderByWorkOrderId		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.QueryString( "intWorkOrderId" ) = "" ) then
	intWorkOrderId = 0
else
	intWorkOrderId = Request.QueryString( "intWorkOrderId" )
end if

cmdGetWorkOrderByWorkOrderId.ActiveConnection = cnnSQLConnection
cmdGetWorkOrderByWorkOrderId.CommandText = "{call spGetWorkOrderByWorkOrderId( ?,?,?,? ) }"
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@IsPrint", adTinyInt, adParamInput, 1 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@CustomerId", adInteger, adParamOutput, 4 )
cmdGetWorkOrderByWorkOrderId.Parameters( "@WorkOrderId" )	= intWorkOrderId
cmdGetWorkOrderByWorkOrderId.Parameters( "@IsPrint" )	= 1

on error resume next
set rsGetWorkOrderByWorkOrderId = cmdGetWorkOrderByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
					
on error goto 0
'**********************************************************************************
%>
<%if not( rsGetWorkOrderByWorkOrderId is nothing ) then
		if( rsGetWorkOrderByWorkOrderId.State = adStateOpen ) then
			if( not rsGetWorkOrderByWorkOrderId.EOF ) then%>
<HTML>
<TITLE>Work Sheet for Work Order&nbsp;&nbsp;#<%=rsGetWorkOrderByWorkOrderId( "WorkOrderNumber" )%></TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="reports_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; TEXT-DECORATION: underline }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder }
	.smalllabel { FONT-SIZE: 80%; FONT-WEIGHT: bolder; }
	.timetable { PADDING-BOTTOM: 10px }
	.emptyspace { FONT-SIZE: 40%; }
  </STYLE>  
</HEAD>
<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width="640" ID="Table1">
      
    <tr><td><table width="640" ID="Table2">
      <tr><td class="label">DATE IN:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "DateIn" )%></td><td class="label" width="190">WO#:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "WorkOrderNumber" )%></td></tr>
			<tr><td class="label" width="450">CUSTOMER:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "Customer" )%></td><td class="label" width="190">PRIORITY:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "WorkOrderPriority" )%></td></tr>
      <tr><td class="label" width="450">REP:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "SalesRepName" )%></td><td class="label" width="190">PO #: <%=rsGetWorkOrderByWorkOrderId( "PONumber" )%></td></tr>
      <tr><td class="label" colspan=2>SPINDLE TYPE:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "SpindleType" )%>&nbsp;&nbsp;&nbsp;SER #:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "SerialNumber" )%></td></tr>
      <tr><td class="label" colspan=2>PROJECT #:&nbsp;<%=rsGetWorkOrderByWorkOrderId( "ProjectId" )%></td></tr>
    </table></td></tr>

    <tr><td><table width="640" ID="Table4">
      <tr><td class="label">REPAIR SPECIFIC INFO: <%=rsGetWorkOrderByWorkOrderId( "AdditionalInfo" )%></td></tr>
    </table></td></tr>
    
    <!---
    <tr><td><table width="640" ID="Table3">
      <tr><td class="emptyspace">&nbsp;</td></tr>
      <tr><td class="label">SPECIFIC REPAIR INFO:_________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
    </table></td></tr>
    --->

    <tr><td><table width="640" ID="Table10">
      <tr><td class="emptyspace">&nbsp;</td></tr>
      <tr><td class="label">INCOMING VISUAL:______________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
    </table></td></tr>

    <tr><td><table width="640" ID="Table11">
      <tr><td class="emptyspace">&nbsp;</td></tr>
      <tr><td class="label" colspan=2>SEIZED:______________ TURNS ROUGH:___________ RUNABLE: _____________</td></tr>
      <tr><td class="emptyspace">&nbsp;</td></tr>
      <tr><td class="label" width="150">PULLEY: Y / N</td>
          <td class="label" width="490">TOOL HOLDER: Y / N</td></tr>
    </table></td></tr>
    <tr><td><table width="640" ID="Table5">
      <tr><td class="emptyspace">&nbsp;</td></tr>
      <tr><td class="label">TECHNICIAN COMMENTS: (ASSEMBLY)</td></tr>
    </table></td></tr>
   
    <tr><td><table width="640" ID="Table6">
      <tr><td class="label">NOTES:________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="label">BALANCE INFO:_________________________________________________________________</td></tr>
      <tr><td class="label">______________________________________________________________________________</td></tr>
      <tr><td class="emptyspace">&nbsp;</td></tr>
    </table></td></tr>

    <tr><td><table width="640" ID="Table4">
      <tr><td class="label">TECHNICIAN:_____________________________________DATE:_________________________</td></tr>
    </table></td></tr>

    <tr><td class="label"><table width="640" ID="Table7">
      <tr><td class="label">DISASSEMBLY & MEASURING</td><td class="label" align=right>Actual Time: _________ hrs.</td></tr>
    </table></td></tr>
    
    <tr><td class="label"><table width="640" ID="Table8">
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label">TECH/DATE:____________ START:_________ END:__________ NOTES:_______________</td></tr>
      <tr><td class="label" style="padding-top:2em;">AR Status : &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;OK&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;HOLD</td></tr>
    </table></td></tr>


  </table>
    
  </table>
</BODY>
</HTML>
<%	end if
	end if
end if

if not( rsGetWorkOrderByWorkOrderId is nothing ) then
	if( rsGetWorkOrderByWorkOrderId.State = adStateOpen ) then
		rsGetWorkOrderByWorkOrderId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetWorkOrderByWorkOrderId			= nothing
set cmdGetWorkOrderByWorkOrderId		= nothing
set cnnSQLConnection								= nothing
%>
