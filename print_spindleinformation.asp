<%option explicit

dim cnnSQLConnection
dim rsGetSpindleInformationByWorkOrderId, cmdGetSpindleInformationByWorkOrderId
dim intWorkOrderId

set cnnSQLConnection					= nothing
set rsGetSpindleInformationByWorkOrderId				= nothing
set cmdGetSpindleInformationByWorkOrderId			= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleInformationByWorkOrderId			= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intWorkOrderId = Request.QueryString( "intWorkOrderId" )

cmdGetSpindleInformationByWorkOrderId.ActiveConnection = cnnSQLConnection
cmdGetSpindleInformationByWorkOrderId.CommandText = "{call spGetSpindleInformationByWorkOrderId( ?,? ) }"
cmdGetSpindleInformationByWorkOrderId.Parameters.Append cmdGetSpindleInformationByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleInformationByWorkOrderId.Parameters.Append cmdGetSpindleInformationByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
cmdGetSpindleInformationByWorkOrderId.Parameters( "@WorkOrderId" )	= intWorkOrderId

on error resume next
set rsGetSpindleInformationByWorkOrderId = cmdGetSpindleInformationByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Spindle Repair Information</TITLE>
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
<%if not( rsGetSpindleInformationByWorkOrderId is nothing ) then
		if( rsGetSpindleInformationByWorkOrderId.State = adStateOpen ) then
			if( not rsGetSpindleInformationByWorkOrderId.EOF ) then%>
  <table align=left width=640>
    
    <tr><td class="title" align=middle>SPINDLE INFORMATION</td></tr>
    
    <tr><td><table class="infotable" width=640>
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0>
      <tr><td class="label" width=100>Spindle Type:</td><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "SpindleType" )%></td>
        <td class="label" width=120>Serial Number:</td><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "SerialNumber" )%></td></tr>
      <tr><td class="label" width=100>Spindle Speed:</td><td class="underscore" width=210>&nbsp;</td>
        <td class="label" width=120>Operating Speed:</td><td class="underscore" width=210>&nbsp;</td></tr>
      </table></td></tr>
        
	  <tr><td class="smalllabel" colspan=2 align="middle">Lubrication</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="item" colspan=2><%=rsGetSpindleInformationByWorkOrderId( "Lubrication" )%></td></tr>
      <%if ( rsGetSpindleInformationByWorkOrderId( "Grease" ) <> "" )  then %>
        <tr><td class="label">Grease</td><td class="item"><%=rsGetSpindleInformationByWorkOrderId( "Grease" )%></td></tr>
      <%end if
        if ( rsGetSpindleInformationByWorkOrderId( "OilMist" ) <> "" ) then %>
        <tr><td class="label">Oil Mist</td><td class="item"><%=rsGetSpindleInformationByWorkOrderId( "OilMist" )%></td></tr>
      <%end if
        if ( rsGetSpindleInformationByWorkOrderId( "OilJet" ) <> "" ) then %>
        <tr><td class="label">Oil Jet</td><td class="item"><%=rsGetSpindleInformationByWorkOrderId( "OilJet" )%></td></tr>
      <%end if%>
        <tr><td class="item" colspan=2><%=rsGetSpindleInformationByWorkOrderId( "LubeNotes" )%></td></tr>
      </table></td></tr>
      
	  <tr><td class="smalllabel" colspan=2 align="middle">Tool Change</td></tr>
      <tr><td class="item" colspan=2 align="middle">Please check one and specify</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Manual</td>
          <td>Specify:______________________________________________________________________</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Automatic</td>
          <td>Specify:______________________________________________________________________</td></tr>
      </table></td></tr>
      
	  <tr><td class="smalllabel" colspan=2 align="middle">Motor Data</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="label" width=120>Category:</td><td class="item"><%=rsGetSpindleInformationByWorkOrderId( "SpindleCategoryName" )%></td></tr>
      </table></td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="label" width=30>V:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "Volts" )%></td>
          <td class="label" width=30>Hz:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "Hz" )%></td>
          <td class="label" width=50>HP/kW:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "HP" )%></td>
          <td class="label" width=30>A:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "Amps" )%></td>
          <td class="label" width=70>Poles:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "Poles" )%></td>
          <td class="label" width=70>Phases:</td><td class="item" width=60><%=rsGetSpindleInformationByWorkOrderId( "Phase" )%></td></tr>
      </table></td></tr>
    </table></td></tr>
    
    
      <tr><td><table width=100% cellpadding=0 cellspacing=0 border=0 ID="Table1">
      <tr><td class="label" width=100>Bal Velocity:</td><% if (rsGetSpindleInformationByWorkOrderId( "BalVelocity" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "BalVelocity" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "BalVelocityFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "BalVelocityFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>GSE:</td><% if (rsGetSpindleInformationByWorkOrderId( "GSE" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "GSE" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "GSEFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "GSEFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Break In:</td><% if (rsGetSpindleInformationByWorkOrderId( "BreakIn" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "BreakIn" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "BreakInFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "BreakInFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Drawforce:</td><% if (rsGetSpindleInformationByWorkOrderId( "ActualDrawforce" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "ActualDrawforce" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "ActualDrawforceFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "ActualDrawforceFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Cooling:</td><% if (rsGetSpindleInformationByWorkOrderId( "Cooling" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "Cooling" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "CoolingFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "CoolingFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>RoomTemp:</td><% if (rsGetSpindleInformationByWorkOrderId( "RoomTemp" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RoomTemp" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "RoomTempFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RoomTempFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Front Temp:</td><% if (rsGetSpindleInformationByWorkOrderId( "FrontTemp" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "FrontTemp" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "FrontTempFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "FrontTempFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>RearTemp:</td><% if (rsGetSpindleInformationByWorkOrderId( "RearTemp" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RearTemp" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "RearTempFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RearTempFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Runout:</td><% if (rsGetSpindleInformationByWorkOrderId( "RunoutFront" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RunoutFront" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "RunoutFrontFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RunoutFrontFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>
      <tr><td class="label" width=100>Rear:</td><% if (rsGetSpindleInformationByWorkOrderId( "Rear" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "Rear" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td>
        <td class="label" width=120>Final:</td><% if (rsGetSpindleInformationByWorkOrderId( "RearFinal" ) <> "") then%><td class="item" width=210><%=rsGetSpindleInformationByWorkOrderId( "RearFinal" )%><%else%><td class="underscore" width=210>&nbsp;<%end if%></td></tr>

      </table></td></tr>
    
    
  </table>
<%	end if
	end if
end if%>
</BODY>
</HTML>
<%
if not( rsGetSpindleInformationByWorkOrderId is nothing ) then
	if( rsGetSpindleInformationByWorkOrderId.State = adStateOpen ) then
		rsGetSpindleInformationByWorkOrderId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleInformationByWorkOrderId			= nothing
set cmdGetSpindleInformationByWorkOrderId		= nothing
set cnnSQLConnection						= nothing
%>