<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetSpindleCategoryList, cmdGetSpindleCategoryList
dim rsGetSpindleBySpindleId, cmdGetSpindleBySpindleId, cmdEditSpindle
dim rsGetProjectList
dim intRowNumber, intSpindleId

set cnnSQLConnection					= nothing
set rsGetSpindleCategoryList					= nothing
set cmdGetSpindleCategoryList					= nothing
set rsGetSpindleBySpindleId				= nothing
set cmdGetSpindleBySpindleId			= nothing
set cmdEditSpindle						= nothing
set rsGetProjectList					= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleCategoryList			= Server.CreateObject( "ADODB.Command" )
set cmdGetSpindleBySpindleId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
cmdGetSpindleCategoryList.ActiveConnection = cnnSQLConnection
cmdGetSpindleCategoryList.CommandText = "{call spGetSpindleCategoryList}"

on error resume next
set rsGetSpindleCategoryList = cmdGetSpindleCategoryList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0

intSpindleId = Request.QueryString( "intSpindleId" )

cmdGetSpindleBySpindleId.ActiveConnection = cnnSQLConnection
'cmdGetSpindleBySpindleId.CommandText = "{call spGetSpindleBySpindleId( ?,? ) }"
cmdGetSpindleBySpindleId.CommandText			= "spGetSpindleBySpindleId"
cmdGetSpindleBySpindleId.CommandType			= adCmdStoredProc
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
<HEAD>
<TITLE>Centerline - View / Edit Spindle</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
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
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	var objSpindleCategoryField;
	
	function SelectSpindleCategory( objFormElement )
	{
		if ( objFormElement.value == -1 )
		{
			OpenSpindleCategory( 0 );
			objSpindleCategoryField = objFormElement[1];
		}
	}
	
	function SpindleCategorySaved( intId, strName )
	{
		objSpindleCategoryField.text = strName;
		objSpindleCategoryField.value = intId;
	}
	
	function OpenSpindleCategory( strSpindleCategoryId )
	{
	  OpenWindow( 'Spindlecategory_details.asp?intSpindleCategoryId='+strSpindleCategoryId,'Centerline','width=640,height=120,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 160 )+',screenY='+GetCenterY( 480 )+',scrollbars=no,dependent=yes' );
	}
	
  function OpenSpindleParts( strSpindleId )
	{
	  OpenWindow( 'spindle_parts_details.asp?intSpindleId='+strSpindleId,'EditSpindleParts','width=722,height=500,left='+GetCenterX( 722 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 722 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}	
	
	function PrintSpindle( strSpindleId )
	{		
		if( strSpindleId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
			//OpenWindow( 'print_spindledetails.asp?intSpindleId='+strSpindleId,'PrintSpindle'+strSpindleId,'width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
			window.print();
	}
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 onload="javascript: if (window.print) window.print();">

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetSpindleBySpindleId is nothing ) then
		if( rsGetSpindleBySpindleId.State = adStateOpen ) then
			if( not rsGetSpindleBySpindleId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intSpindleId = 0 ) then%>
   <tr><td class="sectiontitle" width=20%>New Spindle</td>
   <%else%>
   <tr><td class="sectiontitle" width=20%>Spindle&nbsp;<%=intSpindleId%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Spindle Form -->
<tr><td align=center>
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Type</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "SpindleType" ))%></td>
         <td class="label_2">Category</td><td class="input_2"><%=rsGetSpindleBySpindleId( "SpindleCategoryName" )%></td></tr>
   	 <tr><td class="label_1">Drawing #</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "DrawingNumber" ))%></td></tr>
   	 <tr><td class="label_1">Speed (RPM)</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RPM" ))%></td>
   	   <td class="label_2">Weight</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Weight" ))%></td></tr>
   	 <tr><td class="label_1">Vibration</td><td>
   			<table width="100%" cellspacing=0><tr><td class="generic_label">Front</td><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Vibration" ))%></td>
   				<td class="generic_label">Rear</td><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "VibrationRear" ))%></td>
   			</tr></table></td>
   			<td class="label_2">Coolant Flow (GPM)</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantFlow" ))%></td></tr>
	   <tr><td class="label_1">GSE</td><td>
   			<table width="100%" cellspacing=0 ID="Table1"><tr><td class="generic_label">Front</td><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "GSE" ))%></td>
   				<td class="generic_label">Rear</td><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "GSERear" ))%></td>
   			</tr></table></td></tr>
   	 <tr><td class="label_1">Shaft Temp (&deg;C)</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ShaftTemp" ))%></td>
   			<td class="label_2">Coolant Temp Incoming Set</td><td>
   			<table width="100%" cellspacing=0 ID="Table2"><tr><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantTempIncomingSet" ))%></td>
   				<td class="generic_label">Actual (&deg;C)</td><td width="33%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantTempActual" ))%></td>
   			</tr></table></td></tr>
   	 <tr><td class="label_1">Front Temp (&deg;C)</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "FrontTemp" ))%></td></tr>
   	 <tr><td class="label_1">Rear Temp (&deg;C)</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RearTemp" ))%></td>
   	   <td class="label_2">Coolant Press. Incoming</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantPressureIncoming" ))%></td></tr>
   	 <tr><td class="label_1">Break-In Time</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "BreakInTime" ))%></td>
   	   <td class="label_2">Cooling Method</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolingMethod" ))%></td></tr>
 
  </table>
</td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Volts</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Volts" ))%></td>
       <td class="label_2">Power (HP)</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "HP" ))%></td></tr>
     <tr><td class="label_1">Amps</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Amps" ))%></td>
			<td class="label_2">Phase</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Phase" ))%></td></tr>
     <tr><td class="label_1">Hertz</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Hz" ))%></td>
       <td class="label_2">Thermistor</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Thermistor" ))%></td></tr>
     <tr><td class="label_1">Poles</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Poles" ))%></td>
       <td class="label_2">Amp Draw</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "AmpDraw" ))%></td></tr>
     <tr><td class="label_1">Connectors</td><td>
   			<table width="100%" cellspacing=0 ID="Table3"><tr><td class="generic_label">Ctrl</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ConnectorCtrl" ))%></td>
   				<td class="generic_label">Power</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ConnectorPower" ))%></td>
   			</tr></table></td>
   			<td class="label_2">Converter</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Converter" ))%></td></tr>
   </table>
</td></tr> 
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Tool Type</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolHolder" ))%></td>
       <td class="label_2">Pull Pin</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "PullPin" ))%></td></tr>
     <tr><td class="label_1">E.M. Dimension</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "EMDimension" ))%></td>
       <td class="label_2">Ejection Path</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "EjectionPath" ))%></td></tr>
     <tr><td class="label_1">Tool Out Pressure</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolOutPressure" ))%></td>
			 <td class="label_2">Return Pressure</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ReturnPressure" ))%></td></tr>
     <tr><td class="label_1">Draw Force</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "DrawbarForce" ))%></td></tr>
     <tr><td class="label_1">Tool Change Function</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolChangeFunction" ))%></td>
			 <td class="label_2">Proximity Switch Function</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ProximitySwitchFunction" ))%></td></tr>
 </table>
</td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Lubrication</td><td class="item" colspan=3><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Lubrication" ))%></td></tr>
     <tr><td colspan=4><table width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
       <tr><td class="generic_label" width=102>Grease</td><td width=98><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Grease" ))%></td>
	     <td class="generic_label" width=90 align=center>Oil Mist</td><td width=110><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilMist" ))%></td>
         <td class="generic_label" width=90 align=center>Oil Jet</td><td width=110><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilJet" ))%></td></tr>
     </table></td></tr>
     <tr><td class="label_1">Oil/Grease Type</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilGreaseType" ))%></td>
       <td class="label_2">Interval/DPM</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "IntervalDPM" ))%></td></tr>
     <tr><td class="label_1">Main Pressure</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "MainPressure" ))%></td>
       <td class="label_2">Tube Pressure</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "TubePressure" ))%></td></tr>
     <tr><td class="label_1" colspan=4>Lube Notes</td></tr>
     <tr><td class="item" colspan=4><%=rsGetSpindleBySpindleId( "LubeNotes" )%></td></tr>
 </table>
</td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td colspan=4><table width=600 border=0 cellpadding=0 cellspacing=0 align=middle >
       <tr><td class="generic_label" width=102>Preload</td><td width=98><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Preload" ))%></td>
	     <td class="generic_label" width=90 align=center>Radial Play</td><td width=110><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RadialPlay" ))%></td>
         <td class="generic_label" width=90 align=center>Axial Play</td><td width=110><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "AxialPlay" ))%></td></tr>
     </table></td></tr>
     <tr><td class="label_1">Runout Front</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront" ))%></td>
					<td class="generic_label">Location</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFrontLocation" ))%></td>
				</tr></table></td>
				<td class="label_2">Runout Rear</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear" ))%></td>
					<td class="generic_label">Location</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRearLocation" ))%></td>
				</tr></table></td></tr>
     <tr><td class="label_1">&nbsp;</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront2" ))%></td>
					<td class="generic_label">Location</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront2Location" ))%></td>
				</tr></table></td>
				<td class="label_2">&nbsp;</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear2" ))%></td>
					<td class="generic_label">Location</td><td width="35%"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear2Location" ))%></td>
				</tr></table></td></tr>
     <tr><td class="label_1">Tool Contact</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolContact" ))%></td>
			 <td class="label_2">Tool Contact</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolContactRear" ))%></td></tr>
     <tr><td class="label_1">Tool Gap</td><td class="item"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolGap" ))%></td>
			 <td class="label_2">Tool Gap</td><td class="input_2"><%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolGapRear" ))%></td></tr>
		 <tr><td class="label_1" colspan=4>Other</td></tr>
     <tr><td class="item" colspan=4><%=rsGetSpindleBySpindleId( "Other" )%></td></tr>
 </table>
</td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1" colspan=4>Balancing Requirements</td></tr>
     <tr><td class="item" colspan=4><%=rsGetSpindleBySpindleId( "BalancingRequirements" )%></td></tr>
     <tr><td class="label_1" colspan=4>Bearing Information</td></tr>
     <tr><td class="item" colspan=4><%=rsGetSpindleBySpindleId( "BearingInformation" )%></td></tr>
 </table>
</td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1" colspan=4>General Spindle Notes</td></tr>
     <tr><td class="item" colspan=4><%=rsGetSpindleBySpindleId( "GeneralSpindleNotes" )%></td></tr>
 </table>
</td></tr>
<!-- End Spindle Form -->
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

set rsGetSpindleCategoryList					= nothing
set cmdGetSpindleCategoryList					= nothing
set rsGetSpindleBySpindleId			= nothing
set cmdGetSpindleBySpindleId		= nothing
set cmdEditSpindle							= nothing
set cnnSQLConnection						= nothing
%>