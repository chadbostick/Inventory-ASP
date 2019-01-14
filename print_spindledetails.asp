<%option explicit

dim cnnSQLConnection
dim rsGetSpindleBySpindleId, cmdGetSpindleBySpindleId
dim intSpindleId

set cnnSQLConnection					= nothing
set rsGetSpindleBySpindleId				= nothing
set cmdGetSpindleBySpindleId			= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleBySpindleId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intSpindleId = Request.QueryString( "intSpindleId" )

cmdGetSpindleBySpindleId.ActiveConnection = cnnSQLConnection
cmdGetSpindleBySpindleId.CommandText = "{call spGetSpindleBySpindleId( ?,? ) }"
cmdGetSpindleBySpindleId.Parameters.Append cmdGetSpindleBySpindleId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleBySpindleId.Parameters.Append cmdGetSpindleBySpindleId.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
cmdGetSpindleBySpindleId.Parameters( "@SpindleId" )	= intSpindleId

on error resume next
set rsGetSpindleBySpindleId = cmdGetSpindleBySpindleId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Spindle Information</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; TEXT-DECORATION: underline }
	.label { PADDING-TOP: 15px; FONT-WEIGHT: bolder }
	.smalllabel { BORDER-TOP: black solid thin; FONT-SIZE: 80%; FONT-WEIGHT: bolder }
	.item { PADDING-TOP: 15px }
	.infotable { BORDER-BOTTOM: black solid thin }
	.box { BORDER: black solid 1px }
	.underscore { BORDER-BOTTOM: black solid 1px }
  </STYLE>
</HEAD>

<BODY onload="javascript: if (window.print) window.print();">
<%if not( rsGetSpindleBySpindleId is nothing ) then
		if( rsGetSpindleBySpindleId.State = adStateOpen ) then
			if( not rsGetSpindleBySpindleId.EOF ) then%>
  <table align=left width=640>
    
    <tr><td class="title" align=middle>SPINDLE DETAILS</td></tr>
    
    <tr><td><table class="infotable" width=640>
      
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table2">
      <tr><td class="label" width=100>Spindle Type:</td><td class="item"><%=rsGetSpindleBySpindleId( "SpindleType" )%></td>
        <td class="label" width=120>Category:</td><td class="item"><%=rsGetSpindleBySpindleId( "SpindleCategoryName" )%></td></tr>
      <tr><td class="label" width=100>RPM:</td><td class="item"><%=rsGetSpindleBySpindleId( "RPM" )%></td>
        <td class="label" width=120>Weight:</td><td class="item"><%=rsGetSpindleBySpindleId( "Weight" )%></td></tr>
      <tr><td class="label" width=100>Drawing:</td><td class="item"><%=rsGetSpindleBySpindleId( "DrawingNumber" )%></td>></tr>
      </table></td></tr>
     
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table1">
      <tr><td class="label" width=100>Volts:</td><td class="item"><%=rsGetSpindleBySpindleId( "Volts" )%></td>
        <td class="label" width=120>Hz:</td><td class="item"><%=rsGetSpindleBySpindleId( "Hz" )%></td></tr>
      <tr><td class="label" width=100>HP:</td><td class="item"><%=rsGetSpindleBySpindleId( "HP" )%></td>
        <td class="label" width=120>Amps:</td><td class="item"><%=rsGetSpindleBySpindleId( "Amps" )%></td></tr>
      <tr><td class="label" width=100>Phase:</td><td class="item"><%=rsGetSpindleBySpindleId( "Phase" )%></td>
        <td class="label" width=120>Poles:</td><td class="item"><%=rsGetSpindleBySpindleId( "Poles" )%></td></tr>
      <tr><td class="label" width=100>Cooling Method:</td><td class="item"><%=rsGetSpindleBySpindleId( "CoolingMethod" )%></td>
        <td class="label" width=120>Thermistor:</td><td class="item"><%=rsGetSpindleBySpindleId( "Thermistor" )%></td></tr>
      </table></td></tr>
     
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table4">
      <tr><td class="label" width=100>Tool Holder:</td><td class="item"><%=rsGetSpindleBySpindleId( "ToolHolder" )%></td>
        <td class="label" width=120>Pull Pin:</td><td class="item"><%=rsGetSpindleBySpindleId( "PullPin" )%></td></tr>
      <tr><td class="label" width=100>E.M. Dimension:</td><td class="item"><%=rsGetSpindleBySpindleId( "EMDimension" )%></td>
        <td class="label" width=120>Ejection Path:</td><td class="item"><%=rsGetSpindleBySpindleId( "EjectionPath" )%></td></tr>
      <tr><td class="label" width=100>Tool Out Pressure:</td><td class="item"><%=rsGetSpindleBySpindleId( "ToolOutPressure" )%></td>
        <td class="label" width=120>Return Pressure:</td><td class="item"><%=rsGetSpindleBySpindleId( "ReturnPressure" )%></td></tr>
      <tr><td class="label" width=100>Drawbar Force:</td><td class="item" colspan=3><%=rsGetSpindleBySpindleId( "DrawbarForce" )%></td></tr>
      <tr><td class="label" width=100>Lubrication:</td><td class="item" colspan=3><%=rsGetSpindleBySpindleId( "Lubrication" )%></td></tr>
      </table></td></tr>
        
     <tr><td colspan=2><table width=640 ID="Table3">
        <tr><td class="label">Grease</td><td class="item"><%=rsGetSpindleBySpindleId( "Grease" )%></td>
        <td class="label">Oil Mist</td><td class="item"><%=rsGetSpindleBySpindleId( "OilMist" )%></td>
        <td class="label">Oil Jet</td><td class="item"><%=rsGetSpindleBySpindleId( "OilJet" )%></td></tr>
     </table></td></tr>
     
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table5">
      <tr><td class="label" width=100>Oil/Grease Type:</td><td class="item"><%=rsGetSpindleBySpindleId( "OilGreaseType" )%></td>
        <td class="label" width=120>Interval/DPM:</td><td class="item"><%=rsGetSpindleBySpindleId( "IntervalDPM" )%></td></tr>
      <tr><td class="label" width=100>Main Pressure:</td><td class="item"><%=rsGetSpindleBySpindleId( "MainPressure" )%></td>
        <td class="label" width=120>Tube Pressure:</td><td class="item"><%=rsGetSpindleBySpindleId( "TubePressure" )%></td></tr>
      </table></td></tr>
     
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table6">
      <tr><td class="label">Lube Notes:</td></tr>
      <tr><td class="item"><%=rsGetSpindleBySpindleId( "LubeNotes" )%></td></tr>
      <tr><td class="label">Balancing Requirements:</td></tr>
      <tr><td class="item"><%=rsGetSpindleBySpindleId( "BalancingRequirements" )%></td></tr>
      <tr><td class="label">Bearing Information:</td></tr>
      <tr><td class="item"><%=rsGetSpindleBySpindleId( "BearingInformation" )%></td></tr>
      <tr><td class="label">General Spindle Notes:</td></tr>
      <tr><td class="item"><%=rsGetSpindleBySpindleId( "GeneralSpindleNotes" )%></td></tr>
      </table></td></tr>
  </table>
<%	end if
	end if
end if%>
</BODY>
</HTML>
<%
if not( rsGetSpindleBySpindleId is nothing ) then
	if( rsGetSpindleBySpindleId.State = adStateOpen ) then
		rsGetSpindleBySpindleId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleBySpindleId			= nothing
set cmdGetSpindleBySpindleId		= nothing
set cnnSQLConnection						= nothing
%>