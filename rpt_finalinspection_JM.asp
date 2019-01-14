<%option explicit

dim cnnSQLConnection
dim rsReportFinalInspection, cmdReportFinalInspection
dim rsGetCompanyInformation, cmdGetCompanyInformation
dim rsGetQuotePartsByWorkOrderId, cmdGetQuotePartsByWorkOrderId
dim rsGetQuoteBearingsByWorkOrderId, cmdGetQuoteBearingsByWorkOrderId
dim rsGetQuoteSubWorkByWorkOrderId, cmdGetQuoteSubWorkByWorkOrderId
dim intRowNumber

set cnnSQLConnection					= nothing
set rsReportFinalInspection		= nothing
set cmdReportFinalInspection	= nothing
set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set rsGetQuotePartsByWorkOrderId				= nothing
set cmdGetQuotePartsByWorkOrderId				= nothing
set rsGetQuoteBearingsByWorkOrderId				= nothing
set cmdGetQuoteBearingsByWorkOrderId				= nothing
set cmdGetQuoteSubWorkByWorkOrderId				= nothing
set rsGetQuoteSubWorkByWorkOrderId				= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdReportFinalInspection	= Server.CreateObject( "ADODB.Command" )
set cmdGetCompanyInformation	= Server.CreateObject( "ADODB.Command" )
set cmdGetQuotePartsByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteBearingsByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteSubWorkByWorkOrderId				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportFinalInspection.ActiveConnection	= cnnSQLConnection
cmdReportFinalInspection.CommandText			= "{call spReportFinalInspection( ? ) }"

cmdReportFinalInspection.Parameters.Append cmdReportFinalInspection.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdReportFinalInspection.Parameters( "@WorkOrderId" ) = Request.QueryString( "intWorkOrderId" )

on error resume next
set rsReportFinalInspection = cmdReportFinalInspection.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'**********************************************************************************
cmdGetQuotePartsByWorkOrderId.ActiveConnection	= cnnSQLConnection
cmdGetQuotePartsByWorkOrderId.CommandText		= "{call spGetQuotePartsByWorkOrderId( ?,? ) }"

cmdGetQuotePartsByWorkOrderId.Parameters.Append cmdGetQuotePartsByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuotePartsByWorkOrderId.Parameters.Append cmdGetQuotePartsByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdGetQuotePartsByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= Request.QueryString( "intWorkOrderId" )

on error resume next
set rsGetQuotePartsByWorkOrderId = cmdGetQuotePartsByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetQuoteBearingsByWorkOrderId.ActiveConnection	= cnnSQLConnection
cmdGetQuoteBearingsByWorkOrderId.CommandText		= "{call spGetQuoteBearingsByWorkOrderId( ?,? ) }"

cmdGetQuoteBearingsByWorkOrderId.Parameters.Append cmdGetQuoteBearingsByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteBearingsByWorkOrderId.Parameters.Append cmdGetQuoteBearingsByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdGetQuoteBearingsByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= Request.QueryString( "intWorkOrderId" )

on error resume next
set rsGetQuoteBearingsByWorkOrderId = cmdGetQuoteBearingsByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetQuoteSubWorkByWorkOrderId.ActiveConnection	= cnnSQLConnection
cmdGetQuoteSubWorkByWorkOrderId.CommandText		= "{call spGetQuoteSubWorkByWorkOrderId( ?,? ) }"

cmdGetQuoteSubWorkByWorkOrderId.Parameters.Append cmdGetQuoteSubWorkByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteSubWorkByWorkOrderId.Parameters.Append cmdGetQuoteSubWorkByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdGetQuoteSubWorkByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= Request.QueryString( "intWorkOrderId" )

on error resume next
set rsGetQuoteSubWorkByWorkOrderId = cmdGetQuoteSubWorkByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'*******************************************************************************
cmdGetCompanyInformation.ActiveConnection = cnnSQLConnection
cmdGetCompanyInformation.CommandText = "{call spGetCompanyInformation }"

on error resume next
set rsGetCompanyInformation = cmdGetCompanyInformation.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'*******************************************************************************
if not( rsReportFinalInspection is nothing ) then
	if( rsReportFinalInspection.State = adStateOpen ) then
		if( not rsReportFinalInspection.EOF ) then%>
<HTML>
<TITLE>JM Final Inspection Report</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; BORDER-BOTTOM: black solid thin }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder; BORDER-TOP: black solid thin }
	.smalllabel { BORDER-TOP: black solid thin; FONT-SIZE: 80%; FONT-WEIGHT: bolder }
	.inspectionitem { }
	.inspectiontable { }
  </STYLE>
</HEAD>

<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width=640>
  
    <tr><td><table width=640 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo.png" width=282 height=138></td>
       <td align=right><table>   
        <tr><td class="contact_info" width=100%>
         <table>
         <tr><td class="po_number" align=right colspan=2><%=rsReportFinalInspection( "WorkOrderNumber" )%></td>
          <tr><td colspan=2><%=rsGetCompanyInformation( "CompanyName" )%></td></tr>
          <tr><td colspan=2><%=rsGetCompanyInformation( "Address" )%></td></tr>
          <tr><td><%=rsGetCompanyInformation( "City" )%>,&nbsp;<%=rsGetCompanyInformation( "StateOrProvince" )%>&nbsp;<%=rsGetCompanyInformation( "PostalCode" )%></td><td align=middle><%=rsGetCompanyInformation( "Country" )%></td></tr>
          <tr><td>Phone:&nbsp;<%=rsGetCompanyInformation( "PhoneNumber" )%>&nbsp;&nbsp;&nbsp;</td><td align=right>Fax:&nbsp;<%=rsGetCompanyInformation( "FaxNumber" )%></td></tr>
          <tr><td colspan=2><b>Date:&nbsp;<%=rsReportFinalInspection( "DateOut" )%></b></td></tr>
         </table></td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="title" align=middle>INSPECTION REPORT</td></tr>
    
    <tr><td class="inspectiontable"><table width=640>
      <tr><td>&nbsp;</td></tr>
      <tr><td class="inspectionitem" colspan=3><b>CUSTOMER:&nbsp;</b><%=rsReportFinalInspection( "Customer" )%></td></tr>
      <tr><td class="inspectionitem"><b>SPINDLE TYPE:&nbsp;</b><%=rsReportFinalInspection( "SpindleType" )%></td>
        <td class="inspectionitem"><b>SER. #:&nbsp;</b><%=rsReportFinalInspection( "SerialNumber" )%></td>
        <td class="inspectionitem"><b>P.O.#:&nbsp;</b><%=rsReportFinalInspection( "PONumber" )%></td></tr>
      <tr><td colspan=3>&nbsp;</td></tr>
      <tr><td colspan=3><b>INCOMING INSPECTION: </b><%=rsReportFinalInspection( "IncomingInspection" )%></td></tr>
    </table></td></tr>
    
    <tr><td class="inspectiontable"><table width=640>        
	  <tr><td class="smalllabel" align="middle">PARTS</td></tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td class="inspectionitem"><b>PARTS:&nbsp;</b>      
			<%if not( rsGetQuotePartsByWorkOrderId is nothing ) then
				if( rsGetQuotePartsByWorkOrderId.State = adStateOpen ) then
					intRowNumber = 1
					while( not rsGetQuotePartsByWorkOrderId.EOF )
			%>
						<br><b><%=rsGetQuotePartsByWorkOrderId( "Qty" )%></b>&nbsp;&nbsp;&nbsp;<%=rsGetQuotePartsByWorkOrderId( "ProductName" )%>
			<%
						intRowNumber = intRowNumber + 1
						rsGetQuotePartsByWorkOrderId.MoveNext
					wend
				end if
			end if%>      
      </td></tr>
      
      <tr><td class="inspectionitem"><b>BEARINGS:&nbsp;</b>      
			<%if not( rsGetQuoteBearingsByWorkOrderId is nothing ) then
				if( rsGetQuoteBearingsByWorkOrderId.State = adStateOpen ) then
					intRowNumber = 1
					while( not rsGetQuoteBearingsByWorkOrderId.EOF )
			%>
						<br><b><%=rsGetQuoteBearingsByWorkOrderId( "Qty" )%></b>&nbsp;&nbsp;&nbsp;<%=rsGetQuoteBearingsByWorkOrderId( "PartNumber" )%>
			<%
						intRowNumber = intRowNumber + 1
						rsGetQuoteBearingsByWorkOrderId.MoveNext
					wend
				end if
			end if%>      
      </td></tr>
      
      <tr><td class="inspectionitem"><b>REWORK:&nbsp;</b>      
			<%if not( rsGetQuoteSubWorkByWorkOrderId is nothing ) then
				if( rsGetQuoteSubWorkByWorkOrderId.State = adStateOpen ) then
					intRowNumber = 1
					while( not rsGetQuoteSubWorkByWorkOrderId.EOF )
			%>
						<br><%=rsGetQuoteSubWorkByWorkOrderId( "SubWorkDescription" )%>
			<%
						intRowNumber = intRowNumber + 1
						rsGetQuoteSubWorkByWorkOrderId.MoveNext
					wend
				end if
			end if%>      
      </td></tr>
      
    </table></td></tr>
      
<tr><td class="inspectiontable"><table width=640>
	  <tr><td class="smalllabel" colspan=3 align="middle">FINAL INSPECTION</td></tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td class="inspectionitem"><b>Oil/Grease Type:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "OilGreaseType_Final" )%></td>
        <td class="inspectionitem"><b>Interval/DPM:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "IntervalDPM_Final" )%></td></tr>
      <tr><td class="inspectionitem"><b>Main Pressure:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "MainPressure_Final" )%></td>
        <td class="inspectionitem"><b>Tube Pressure:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "TubePressure_Final" )%></td></tr>
      <tr><td class="inspectionitem"><b>BREAK-IN TIME:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "BreakInTime_Final" )%></td>
        <td class="inspectionitem"><b>DRAWFORCE:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "DrawbarForce_Final" )%></td></tr>
      <tr><td class="inspectionitem"><b>COOLANT FLOW GPM:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "CoolantFlow_Final" )%></td>
        <td class="inspectionitem"><b>ROOM TEMP:&nbsp;&nbsp;&nbsp;</b>20C</td></tr>
      <tr><td class="inspectionitem"><b>FRONT TEMP:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "FrontTemp_Final" )%></td>
        <td class="inspectionitem"><b>REAR TEMP:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "RearTemp_Final" )%></td></tr>
      <tr><td class="inspectionitem"><b>RUNOUT FRONT:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "RunoutFront_Final" )%></td>
        <td class="inspectionitem"><b>RUNOUT REAR:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "RunoutRear_Final" )%></td></tr>	
      <tr><td class="inspectionitem"><b>RUNOUT FRONT LOCATION:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "RunoutFrontLocation_Final" )%></td>
        <td class="inspectionitem"><b>RUNOUT REAR LOCATION:&nbsp;&nbsp;&nbsp;</b><%=rsReportFinalInspection( "RunoutRearLocation_Final" )%></td></tr>	
      <tr><td colspan=3>&nbsp;</td></tr>
      <tr><td colspan=3><b>REMARKS:&nbsp;</b><%=rsReportFinalInspection( "Remarks" )%></td></tr>
      
	  <tr><td colspan=3>Centerline, Inc. compiles and archives full data sets for all spindle repairs performed (incoming and outgoing values, vibration analysis and photographic documentation). This Final Inspection Report is intended as an overview of spindle repair. Please contact Customer Service if you require further information.</td></tr>
    </table></td></tr>
    
  </table>
</BODY>
</HTML>
<%end if
end if
end if

if not( rsReportFinalInspection is nothing ) then
	if( rsReportFinalInspection.State = adStateOpen ) then
		rsReportFinalInspection.Close
	end if
end if

if not( rsGetCompanyInformation is nothing ) then
	if( rsGetCompanyInformation.State = adStateOpen ) then
		rsGetCompanyInformation.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set rsReportFinalInspection		= nothing
set cmdReportFinalInspection	= nothing
set cnnSQLConnection					= nothing
%>
