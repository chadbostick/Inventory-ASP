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
'**********************************************************************************

if( Request.Form.Count > 0 ) then
	set cmdEditSpindle	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditSpindle.ActiveConnection	= cnnSQLConnection
	'cmdEditSpindle.CommandText			= "{? = call spEditSpindle( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) }"
	cmdEditSpindle.CommandText			= "spEditSpindle"
	cmdEditSpindle.CommandType			= adCmdStoredProc
	
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@SpindleType", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@SpindleCategoryId", adInteger, adParamInput, 4 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@DrawingNumber", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RPM", adVarChar, adParamInput, 10 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Weight", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Vibration", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@VibrationRear", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@CoolantFlow", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@GSE", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@GSERear", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ShaftTemp", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@CoolantTempIncomingSet", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@CoolantTempActual", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@FrontTemp", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RearTemp", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@CoolantPressureIncoming", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@BreakInTime", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@CoolingMethod", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Volts", adVarChar, adParamInput, 10 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@HP", adVarChar, adParamInput, 10 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Amps", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Phase", adVarChar, adParamInput, 1 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Hz", adVarChar, adParamInput, 10 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Thermistor", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Poles", adVarChar, adParamInput, 10 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@AmpDraw", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ConnectorCtrl", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ConnectorPower", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Converter", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolHolder", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@PullPin", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@EMDimension", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@EjectionPath", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolOutPressure", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ReturnPressure", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@DrawbarForce", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolChangeFunction", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ProximitySwitchFunction", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Lubrication", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Grease", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@OilMist", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@OilJet", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@OilGreaseType", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@IntervalDPM", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@MainPressure", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@TubePressure", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@LubeNotes", adLongVarChar, adParamInput, 1048576 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Preload", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RadialPlay", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@AxialPlay", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutFront", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutFrontLocation", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutFront2", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutFront2Location", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutRear", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutRearLocation", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutRear2", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@RunoutRear2Location", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolContact", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolContactRear", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolGap", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@ToolGapRear", adVarChar, adParamInput, 100 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@Other", adLongVarChar, adParamInput, 1048576 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@BalancingRequirements", adLongVarChar, adParamInput, 1048576 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@BearingInformation", adLongVarChar, adParamInput, 1048576 )
	cmdEditSpindle.Parameters.Append cmdEditSpindle.CreateParameter( "@GeneralSpindleNotes", adLongVarChar, adParamInput, 1048576 )
	
	
	if( Request.Form( "txtSpindleId" ) <> "" ) then cmdEditSpindle.Parameters( "@SpindleId" ) = Request.Form( "txtSpindleId" )
	if( Request.Form( "txtSpindleType" ) <> "" ) then cmdEditSpindle.Parameters( "@SpindleType" ) = Request.Form( "txtSpindleType" )
	if( Request.Form( "selSpindleCategory" ) <> "" ) then cmdEditSpindle.Parameters( "@SpindleCategoryId" ) = Request.Form( "selSpindleCategory" )
	if( Request.Form( "txtDrawingNumber" ) <> "" ) then cmdEditSpindle.Parameters( "@DrawingNumber" ) = Request.Form( "txtDrawingNumber" )
	if( Request.Form( "txtRPM" ) <> "" ) then cmdEditSpindle.Parameters( "@RPM" ) = Request.Form( "txtRPM" )
	if( Request.Form( "txtWeight" ) <> "" ) then cmdEditSpindle.Parameters( "@Weight" ) = Request.Form( "txtWeight" )
	if( Request.Form( "txtVibration" ) <> "" ) then cmdEditSpindle.Parameters( "@Vibration" ) = Request.Form( "txtVibration" )
	if( Request.Form( "txtVibrationRear" ) <> "" ) then cmdEditSpindle.Parameters( "@VibrationRear" ) = Request.Form( "txtVibrationRear" )
	if( Request.Form( "txtCoolantFlow" ) <> "" ) then cmdEditSpindle.Parameters( "@CoolantFlow" ) = Request.Form( "txtCoolantFlow" )
	if( Request.Form( "txtGSE" ) <> "" ) then cmdEditSpindle.Parameters( "@GSE" ) = Request.Form( "txtGSE" )
	if( Request.Form( "txtGSERear" ) <> "" ) then cmdEditSpindle.Parameters( "@GSERear" ) = Request.Form( "txtGSERear" )
	if( Request.Form( "txtShaftTemp" ) <> "" ) then cmdEditSpindle.Parameters( "@ShaftTemp" ) = Request.Form( "txtShaftTemp" )
	if( Request.Form( "txtCoolantTempIncomingSet" ) <> "" ) then cmdEditSpindle.Parameters( "@CoolantTempIncomingSet" ) = Request.Form( "txtCoolantTempIncomingSet" )
	if( Request.Form( "txtCoolantTempActual" ) <> "" ) then cmdEditSpindle.Parameters( "@CoolantTempActual" ) = Request.Form( "txtCoolantTempActual" )
	if( Request.Form( "txtFrontTemp" ) <> "" ) then cmdEditSpindle.Parameters( "@FrontTemp" ) = Request.Form( "txtFrontTemp" )
	if( Request.Form( "txtRearTemp" ) <> "" ) then cmdEditSpindle.Parameters( "@RearTemp" ) = Request.Form( "txtRearTemp" )
	if( Request.Form( "txtCoolantPressureIncoming" ) <> "" ) then cmdEditSpindle.Parameters( "@CoolantPressureIncoming" ) = Request.Form( "txtCoolantPressureIncoming" )
	if( Request.Form( "txtBreakInTime" ) <> "" ) then cmdEditSpindle.Parameters( "@BreakInTime" ) = Request.Form( "txtBreakInTime" )
	if( Request.Form( "txtCoolingMethod" ) <> "" ) then cmdEditSpindle.Parameters( "@CoolingMethod" ) = Request.Form( "txtCoolingMethod" )
	if( Request.Form( "txtVolts" ) <> "" ) then cmdEditSpindle.Parameters( "@Volts" ) = Request.Form( "txtVolts" )
	if( Request.Form( "txtHP" ) <> "" ) then cmdEditSpindle.Parameters( "@HP" ) = Request.Form( "txtHP" )
	if( Request.Form( "txtAmps" ) <> "" ) then cmdEditSpindle.Parameters( "@Amps" ) = Request.Form( "txtAmps" )
	if( Request.Form( "txtPhase" ) <> "" ) then cmdEditSpindle.Parameters( "@Phase" ) = Request.Form( "txtPhase" )
	if( Request.Form( "txtHz" ) <> "" ) then cmdEditSpindle.Parameters( "@Hz" ) = Request.Form( "txtHz" )
	if( Request.Form( "txtThermistor" ) <> "" ) then cmdEditSpindle.Parameters( "@Thermistor" ) = Request.Form( "txtThermistor" )
	if( Request.Form( "txtPoles" ) <> "" ) then cmdEditSpindle.Parameters( "@Poles" ) = Request.Form( "txtPoles" )
	if( Request.Form( "txtAmpDraw" ) <> "" ) then cmdEditSpindle.Parameters( "@AmpDraw" ) = Request.Form( "txtAmpDraw" )
	if( Request.Form( "txtConnectorCtrl" ) <> "" ) then cmdEditSpindle.Parameters( "@ConnectorCtrl" ) = Request.Form( "txtConnectorCtrl" )
	if( Request.Form( "txtConnectorPower" ) <> "" ) then cmdEditSpindle.Parameters( "@ConnectorPower" ) = Request.Form( "txtConnectorPower" )
	if( Request.Form( "txtConverter" ) <> "" ) then cmdEditSpindle.Parameters( "@Converter" ) = Request.Form( "txtConverter" )
	if( Request.Form( "txtToolHolder" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolHolder" ) = Request.Form( "txtToolHolder" )
	if( Request.Form( "txtPullPin" ) <> "" ) then cmdEditSpindle.Parameters( "@PullPin" ) = Request.Form( "txtPullPin" )
	if( Request.Form( "txtEMDimension" ) <> "" ) then cmdEditSpindle.Parameters( "@EMDimension" ) = Request.Form( "txtEMDimension" )
	if( Request.Form( "txtEjectionPath" ) <> "" ) then cmdEditSpindle.Parameters( "@EjectionPath" ) = Request.Form( "txtEjectionPath" )
	if( Request.Form( "txtToolOutPressure" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolOutPressure" ) = Request.Form( "txtToolOutPressure" )
	if( Request.Form( "txtReturnPressure" ) <> "" ) then cmdEditSpindle.Parameters( "@ReturnPressure" ) = Request.Form( "txtReturnPressure" )
	if( Request.Form( "txtDrawbarForce" ) <> "" ) then cmdEditSpindle.Parameters( "@DrawbarForce" ) = Request.Form( "txtDrawbarForce" )
	if( Request.Form( "txtToolChangeFunction" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolChangeFunction" ) = Request.Form( "txtToolChangeFunction" )
	if( Request.Form( "txtProximitySwitchFunction" ) <> "" ) then cmdEditSpindle.Parameters( "@ProximitySwitchFunction" ) = Request.Form( "txtProximitySwitchFunction" )
	if( Request.Form( "txtLubrication" ) <> "" ) then cmdEditSpindle.Parameters( "@Lubrication" ) = Request.Form( "txtLubrication" )
	if( Request.Form( "txtGrease" ) <> "" ) then cmdEditSpindle.Parameters( "@Grease" ) = Request.Form( "txtGrease" )
	if( Request.Form( "txtOilMist" ) <> "" ) then cmdEditSpindle.Parameters( "@OilMist" ) = Request.Form( "txtOilMist" )
	if( Request.Form( "txtOilJet" ) <> "" ) then cmdEditSpindle.Parameters( "@OilJet" ) = Request.Form( "txtOilJet" )
	if( Request.Form( "txtOilGreaseType" ) <> "" ) then cmdEditSpindle.Parameters( "@OilGreaseType" ) = Request.Form( "txtOilGreaseType" )
	if( Request.Form( "txtIntervalDPM" ) <> "" ) then cmdEditSpindle.Parameters( "@IntervalDPM" ) = Request.Form( "txtIntervalDPM" )
	if( Request.Form( "txtMainPressure" ) <> "" ) then cmdEditSpindle.Parameters( "@MainPressure" ) = Request.Form( "txtMainPressure" )
	if( Request.Form( "txtTubePressure" ) <> "" ) then cmdEditSpindle.Parameters( "@TubePressure" ) = Request.Form( "txtTubePressure" )
	if( Request.Form( "txtLubeNotes" ) <> "" ) then cmdEditSpindle.Parameters( "@LubeNotes" ) = Request.Form( "txtLubeNotes" )
	if( Request.Form( "txtPreload" ) <> "" ) then cmdEditSpindle.Parameters( "@Preload" ) = Request.Form( "txtPreload" )
	if( Request.Form( "txtRadialPlay" ) <> "" ) then cmdEditSpindle.Parameters( "@RadialPlay" ) = Request.Form( "txtRadialPlay" )
	if( Request.Form( "txtAxialPlay" ) <> "" ) then cmdEditSpindle.Parameters( "@AxialPlay" ) = Request.Form( "txtAxialPlay" )
	if( Request.Form( "txtRunoutFront" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutFront" ) = Request.Form( "txtRunoutFront" )
	if( Request.Form( "txtRunoutFrontLocation" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutFrontLocation" ) = Request.Form( "txtRunoutFrontLocation" )
	if( Request.Form( "txtRunoutFront2" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutFront2" ) = Request.Form( "txtRunoutFront2" )
	if( Request.Form( "txtRunoutFront2Location" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutFront2Location" ) = Request.Form( "txtRunoutFront2Location" )
	if( Request.Form( "txtRunoutRear" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutRear" ) = Request.Form( "txtRunoutRear" )
	if( Request.Form( "txtRunoutRearLocation" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutRearLocation" ) = Request.Form( "txtRunoutRearLocation" )
	if( Request.Form( "txtRunoutRear2" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutRear2" ) = Request.Form( "txtRunoutRear2" )
	if( Request.Form( "txtRunoutRear2Location" ) <> "" ) then cmdEditSpindle.Parameters( "@RunoutRear2Location" ) = Request.Form( "txtRunoutRear2Location" )
	if( Request.Form( "txtToolContact" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolContact" ) = Request.Form( "txtToolContact" )
	if( Request.Form( "txtToolContactRear" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolContactRear" ) = Request.Form( "txtToolContactRear" )
	if( Request.Form( "txtToolGap" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolGap" ) = Request.Form( "txtToolGap" )
	if( Request.Form( "txtToolGapRear" ) <> "" ) then cmdEditSpindle.Parameters( "@ToolGapRear" ) = Request.Form( "txtToolGapRear" )
	if( Request.Form( "txtOther" ) <> "" ) then cmdEditSpindle.Parameters( "@Other" ) = Request.Form( "txtOther" )
	if( Request.Form( "txtBalancingRequirements" ) <> "" ) then cmdEditSpindle.Parameters( "@BalancingRequirements" ) = Request.Form( "txtBalancingRequirements" )
	if( Request.Form( "txtBearingInformation" ) <> "" ) then cmdEditSpindle.Parameters( "@BearingInformation" ) = Request.Form( "txtBearingInformation" )
	if( Request.Form( "txtGeneralSpindleNotes" ) <> "" ) then cmdEditSpindle.Parameters( "@GeneralSpindleNotes" ) = Request.Form( "txtGeneralSpindleNotes" )
	
	
	on error resume next
	cmdEditSpindle.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0
	
	intSpindleId = cmdEditSpindle.Parameters( "@Return" ).Value
	
else
	intSpindleId = Request.QueryString( "intSpindleId" )
end if

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
			OpenPrintBox( 'print_spindle_details.asp?intSpindleId='+strSpindleId );
			//if (window.print) window.print();
	}
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

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
 <form id="frmSpindle" name="frmSpindle" action="" method="post">
 <input type="hidden" name="txtSpindleId" id="txtSpindleId" value="<%=intSpindleId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Type</td><td class="input_1"><input type="text" class="forminput" name="txtSpindleType" maxlength=100 isRequired="true" onchange="this.isDirty=true;" fieldName="Type" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "SpindleType" ))%>"></td>
         <td class="label_2">Category</td><td class="input_2"><select class="forminput" name="selSpindleCategory" size=1 isRequired="true" onchange="this.isDirty=true;" fieldName="Category" onChange="javascript:SelectSpindleCategory(this);">
   			<option value=""></option>
   			<option value="-1">- ADD NEW SPINDLE CATEGORY -</option>
   	    <%if not( rsGetSpindleCategoryList is nothing ) then
   					if( rsGetSpindleCategoryList.State = adStateOpen ) then
   						while( not rsGetSpindleCategoryList.EOF )%><option value="<%=rsGetSpindleCategoryList( "SpindleCategoryId" )%>"<%if( rsGetSpindleCategoryList( "SpindleCategoryId" ) = rsGetSpindleBySpindleId( "SpindleCategoryId" ) ) then%>SELECTED<%end if%>><%=rsGetSpindleCategoryList( "SpindleCategoryName" )%></option>
   							<%rsGetSpindleCategoryList.MoveNext
   						wend
   					end if
   				end if%>
   	  </select></td></tr>
   	 <tr><td class="label_1">Drawing #</td><td class="input_1"><input type="text" class="forminput" name="txtDrawingNumber" maxlength=100 onchange="this.isDirty=true;" fieldName="Drawing Number" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "DrawingNumber" ))%>"></td>
   	   <% if intSpindleId > 0 then %><td class="label_2">Parts</td><td class="generic_input" colspan=2><a href="javascript:OpenSpindleParts(<%=intSpindleId%>)">Edit Spindle Parts</a></td><% end if %></tr>
   	 <tr><td class="label_1">Speed (RPM)</td><td class="input_1"><input type="text" class="forminput" name="txtRPM" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="RPM" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RPM" ))%>"></td>
   	   <td class="label_2">Weight</td><td class="input_2"><input type="text" class="forminput" name="txtWeight" maxlength=100 onchange="this.isDirty=true;" fieldName="Weight" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Weight" ))%>"></td></tr>
   	 <tr><td class="label_1">Vibration</td><td>
   			<table width="100%" cellspacing=0><tr><td class="generic_label">Front</td><td width="33%"><input type="text" class="forminput" name="txtVibration" maxlength=100 onchange="this.isDirty=true;" fieldName="Vibration" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Vibration" ))%>"></td>
   				<td class="generic_label">Rear</td><td width="33%"><input type="text" class="forminput" name="txtVibrationRear" maxlength=100 onchange="this.isDirty=true;" fieldName="Vibration Rear" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "VibrationRear" ))%>"></td>
   			</tr></table></td>
   			<td class="label_2">Coolant Flow (GPM)</td><td class="input_2"><input type="text" class="forminput" name="txtCoolantFlow" maxlength=100 onchange="this.isDirty=true;" fieldName="Weight" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantFlow" ))%>"></td></tr>
	   <tr><td class="label_1">GSE</td><td>
   			<table width="100%" cellspacing=0 ID="Table1"><tr><td class="generic_label">Front</td><td width="33%"><input type="text" class="forminput" name="txtGSE" maxlength=100 onchange="this.isDirty=true;" fieldName="RPM" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "GSE" ))%>"></td>
   				<td class="generic_label">Rear</td><td width="33%"><input type="text" class="forminput" name="txtGSERear" maxlength=100 onchange="this.isDirty=true;" fieldName="RPM" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "GSERear" ))%>"></td>
   			</tr></table></td></tr>
   	 <tr><td class="label_1">Shaft Temp (&deg;C)</td><td class="input_1"><input type="text" class="forminput" name="txtShaftTemp" maxlength=100 onchange="this.isDirty=true;" fieldName="Shaft Temp" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ShaftTemp" ))%>"></td>
   			<td class="label_2">Coolant Temp Incoming Set</td><td>
   			<table width="100%" cellspacing=0 ID="Table2"><tr><td width="33%"><input type="text" class="forminput" name="txtCoolantTempIncomingSet" maxlength=100 onchange="this.isDirty=true;" fieldName="Coolant Temp Incoming Set" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantTempIncomingSet" ))%>"></td>
   				<td class="generic_label">Actual (&deg;C)</td><td width="33%"><input type="text" class="forminput" name="txtCoolantTempActual" maxlength=100 onchange="this.isDirty=true;" fieldName="Coolant Temp Actual" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantTempActual" ))%>"></td>
   			</tr></table></td></tr>
   	 <tr><td class="label_1">Front Temp (&deg;C)</td><td class="input_1"><input type="text" class="forminput" name="txtFrontTemp" maxlength=100 onchange="this.isDirty=true;" fieldName="Front Temp" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "FrontTemp" ))%>"></td></tr>
   	 <tr><td class="label_1">Rear Temp (&deg;C)</td><td class="input_1"><input type="text" class="forminput" name="txtRearTemp" maxlength=100 onchange="this.isDirty=true;" fieldName="Rear Temp" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RearTemp" ))%>"></td>
   	   <td class="label_2">Coolant Press. Incoming</td><td class="input_2"><input type="text" class="forminput" name="txtCoolantPressureIncoming" maxlength=100 onchange="this.isDirty=true;" fieldName="Coolant Pressure Incoming" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolantPressureIncoming" ))%>"></td></tr>
   	 <tr><td class="label_1">Break-In Time</td><td class="input_1"><input type="text" class="forminput" name="txtBreakInTime" maxlength=100 onchange="this.isDirty=true;" fieldName="BreakInTime" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "BreakInTime" ))%>"></td>
   	   <td class="label_2">Cooling Method</td><td class="input_2"><input type="text" class="forminput" name="txtCoolingMethod" maxlength=100 onchange="this.isDirty=true;" fieldName="Cooling Method" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "CoolingMethod" ))%>"></td></tr>
 
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
     <tr><td class="label_1">Volts</td><td class="input_1"><input type="text" class="forminput" name="txtVolts" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="Volts" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Volts" ))%>"></td>
       <td class="label_2">Power (HP)</td><td class="input_1"><input type="text" class="forminput" name="txtHP" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="Power" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "HP" ))%>"></td></tr>
     <tr><td class="label_1">Amps</td><td class="input_1"><input type="text" class="forminput" name="txtAmps" maxlength=100 onchange="this.isDirty=true;" fieldName="Amps" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Amps" ))%>"></td>
			<td class="label_2">Phase</td><td class="input_2"><input type="text" class="forminput" name="txtPhase" maxlength=1 isNumber="true" onchange="this.isDirty=true;" fieldName="Phase" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Phase" ))%>"></td></tr>
     <tr><td class="label_1">Hertz</td><td class="input_1"><input type="text" class="forminput" name="txtHz" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="Hz" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Hz" ))%>"></td>
       <td class="label_2">Thermistor</td><td class="input_2"><input type="text" class="forminput" name="txtThermistor" onchange="this.isDirty=true;" fieldName="Thermistor" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Thermistor" ))%>"></td></tr>
     <tr><td class="label_1">Poles</td><td class="input_1"><input type="text" class="forminput" name="txtPoles" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="Poles" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Poles" ))%>"></td>
       <td class="label_2">Amp Draw</td><td class="input_2"><input type="text" class="forminput" name="txtAmpDraw" onchange="this.isDirty=true;" fieldName="Amp Draw" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "AmpDraw" ))%>"></td></tr>
     <tr><td class="label_1">Connectors</td><td>
   			<table width="100%" cellspacing=0 ID="Table3"><tr><td class="generic_label">Ctrl</td><td width="35%"><input type="text" class="forminput" name="txtConnectorCtrl" maxlength=100 onchange="this.isDirty=true;" fieldName="Connector Ctrl" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ConnectorCtrl" ))%>"></td>
   				<td class="generic_label">Power</td><td width="35%"><input type="text" class="forminput" name="txtConnectorPower" maxlength=100 onchange="this.isDirty=true;" fieldName="ConnectorPower" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ConnectorPower" ))%>"></td>
   			</tr></table></td>
   			<td class="label_2">Converter</td><td class="input_2"><input type="text" class="forminput" name="txtConverter" maxlength=100 onchange="this.isDirty=true;" fieldName="Weight" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Converter" ))%>"></td></tr>
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
     <tr><td class="label_1">Tool Type</td><td class="input_1"><input type="text" class="forminput" name="txtToolHolder" maxlength=100 onchange="this.isDirty=true;" fieldName="Tool Type" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolHolder" ))%>"></td>
       <td class="label_2">Retention Knob</td><td class="input_2"><input type="text" class="forminput" name="txtPullPin" maxlength=100 onchange="this.isDirty=true;" fieldName="Pull Pin" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "PullPin" ))%>"></td></tr>
     <tr><td class="label_1">E.M. Dimension</td><td class="input_1"><input type="text" class="forminput" name="txtEMDimension" maxlength=100 onchange="this.isDirty=true;" fieldName="E.M. Dimension" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "EMDimension" ))%>"></td>
       <td class="label_2">Ejection Path</td><td class="input_2"><input type="text" class="forminput" name="txtEjectionPath" maxlength=100 onchange="this.isDirty=true;" fieldName="Ejection Path" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "EjectionPath" ))%>"></td></tr>
     <tr><td class="label_1">Tool Out Pressure</td><td class="input_1"><input type="text" class="forminput" name="txtToolOutPressure" onchange="this.isDirty=true;" fieldName="Tool Out Pressure" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolOutPressure" ))%>"></td>
			 <td class="label_2">Return Pressure</td><td class="input_2"><input type="text" class="forminput" name="txtReturnPressure" maxlength=100 onchange="this.isDirty=true;" fieldName="Return Pressure" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ReturnPressure" ))%>"></td></tr>
     <tr><td class="label_1">Draworce</td><td class="input_1"><input type="text" class="forminput" name="txtDrawbarForce" maxlength=100 onchange="this.isDirty=true;" fieldName="Drawbar Force" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "DrawbarForce" ))%>"></td></tr>
     <tr><td class="label_1">Tool Change Function</td><td class="input_1"><input type="text" class="forminput" name="txtToolChangeFunction" onchange="this.isDirty=true;" fieldName="Tool Change Function" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolChangeFunction" ))%>"></td>
			 <td class="label_2">Proximity Switch Function</td><td class="input_2"><input type="text" class="forminput" name="txtProximitySwitchFunction" maxlength=100 onchange="this.isDirty=true;" fieldName="Proximity Switch Function" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ProximitySwitchFunction" ))%>"></td></tr>
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
     <tr><td class="label_1">Lubrication</td><td class="input_1" colspan=3><input type="text" class="forminput" name="txtLubrication" maxlength=100 onchange="this.isDirty=true;" fieldName="Lubrication" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Lubrication" ))%>"></td></tr>
     <tr><td colspan=4><table width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
       <tr><td class="generic_label" width=102>Grease</td><td width=98><input type="text" class="forminput" name="txtGrease" maxlength=100 onchange="this.isDirty=true;" fieldName="Grease" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Grease" ))%>"></td>
	     <td class="generic_label" width=90 align=center>Oil Mist</td><td width=110><input type="text" class="forminput" name="txtOilMist" onchange="this.isDirty=true;" fieldName="Oil Mist" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilMist" ))%>"></td>
         <td class="generic_label" width=90 align=center>Oil Jet</td><td width=110><input type="text" class="forminput" name="txtOilJet" onchange="this.isDirty=true;" fieldName="Oil Jet" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilJet" ))%>"></td></tr>
     </table></td></tr>
     <tr><td class="label_1">Oil/Grease Type</td><td class="input_1"><input type="text" class="forminput" name="txtOilGreaseType" onchange="this.isDirty=true;" fieldName="Oil Grease Type" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "OilGreaseType" ))%>"></td>
       <td class="label_2">Interval/DPM</td><td class="input_2"><input type="text" class="forminput" name="txtIntervalDPM" maxlength=100 onchange="this.isDirty=true;" fieldName="Interval DPM" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "IntervalDPM" ))%>"></td></tr>
     <tr><td class="label_1">Main Pressure</td><td class="input_1"><input type="text" class="forminput" name="txtMainPressure" maxlength=100 onchange="this.isDirty=true;" fieldName="Main Pressure" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "MainPressure" ))%>"></td>
       <td class="label_2">Tube Pressure</td><td class="input_2"><input type="text" class="forminput" name="txtTubePressure" maxlength=100 onchange="this.isDirty=true;" fieldName="Tube Pressure" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "TubePressure" ))%>"></td></tr>
     <tr><td class="label_1" colspan=4>Lube Notes</td></tr>
     <tr><td class="input_1" colspan=4><textarea class="forminput" name="txtLubeNotes" onchange="this.isDirty=true;" fieldName="Lube Notes" rows=4 wrap=soft><%=rsGetSpindleBySpindleId( "LubeNotes" )%></textarea></td></tr>
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
       <tr><td class="generic_label" width=102>Preload</td><td width=98><input type="text" class="forminput" name="txtPreload" maxlength=100 onchange="this.isDirty=true;" fieldName="Preload" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "Preload" ))%>"></td>
	     <td class="generic_label" width=90 align=center>Radial Play</td><td width=110><input type="text" class="forminput" name="txtRadialPlay" onchange="this.isDirty=true;" fieldName="Radial Play" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RadialPlay" ))%>"></td>
         <td class="generic_label" width=90 align=center>Axial Play</td><td width=110><input type="text" class="forminput" name="txtAxialPlay" onchange="this.isDirty=true;" fieldName="Axial Play" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "AxialPlay" ))%>"></td></tr>
     </table></td></tr>
     <tr><td class="label_1">Runout Front</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><input type="text" class="forminput" name="txtRunoutFront" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Front" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront" ))%>"></td>
					<td class="generic_label">Location</td><td width="35%"><input type="text" class="forminput" name="txtRunoutFrontLocation" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Front Location" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFrontLocation" ))%>"></td>
				</tr></table></td>
				<td class="label_2">Runout Rear</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><input type="text" class="forminput" name="txtRunoutRear" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Rear" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear" ))%>"></td>
					<td class="generic_label">Location</td><td width="35%"><input type="text" class="forminput" name="txtRunoutRearLocation" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Rear Location" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRearLocation" ))%>"></td>
				</tr></table></td></tr>
     <tr><td class="label_1">&nbsp;</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><input type="text" class="forminput" name="txtRunoutFront2" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Front 2" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront2" ))%>"></td>
					<td class="generic_label">Location</td><td width="35%"><input type="text" class="forminput" name="txtRunoutFront2Location" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Front 2 Location" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutFront2Location" ))%>"></td>
				</tr></table></td>
				<td class="label_2">&nbsp;</td>
				<td><table width="100%" cellspacing=0><tr><td width="35%"><input type="text" class="forminput" name="txtRunoutRear2" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Rear2" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear2" ))%>"></td>
					<td class="generic_label">Location</td><td width="35%"><input type="text" class="forminput" name="txtRunoutRear2Location" maxlength=100 onchange="this.isDirty=true;" fieldName="Runout Rear 2 Location" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "RunoutRear2Location" ))%>"></td>
				</tr></table></td></tr>
     <tr><td class="label_1">Tool Contact</td><td class="input_1"><input type="text" class="forminput" name="txtToolContact" onchange="this.isDirty=true;" fieldName="Tool Contact" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolContact" ))%>"></td>
			 <td class="label_2">Tool Contact</td><td class="input_2"><input type="text" class="forminput" name="txtToolContactRear" maxlength=100 onchange="this.isDirty=true;" fieldName="ToolContactRear" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolContactRear" ))%>"></td></tr>
     <tr><td class="label_1">Tool Gap</td><td class="input_1"><input type="text" class="forminput" name="txtToolGap" onchange="this.isDirty=true;" fieldName="Tool Gap" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolGap" ))%>"></td>
			 <td class="label_2">Tool Gap</td><td class="input_2"><input type="text" class="forminput" name="txtToolGapRear" maxlength=100 onchange="this.isDirty=true;" fieldName="Tool Gap Rear" value="<%=EscapeDoubleQuotes(rsGetSpindleBySpindleId( "ToolGapRear" ))%>"></td></tr>
		 <tr><td class="label_1" colspan=4>Other</td></tr>
     <tr><td class="input_1" colspan=4><textarea class="forminput" name="txtOther" onchange="this.isDirty=true;" fieldName="Other" rows=2 wrap=soft><%=rsGetSpindleBySpindleId( "Other" )%></textarea></td></tr>
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
     <tr><td class="input_1" colspan=4><textarea class="forminput" name="txtBalancingRequirements" onchange="this.isDirty=true;" fieldName="Balancing Requirements" rows=3 wrap=soft><%=rsGetSpindleBySpindleId( "BalancingRequirements" )%></textarea></td></tr>
     <tr><td class="label_1" colspan=4>Bearing Information</td></tr>
     <tr><td class="input_1" colspan=4><textarea class="forminput" name="txtBearingInformation" rows=3 onchange="this.isDirty=true;" fieldName="Bearing Information" wrap=soft><%=rsGetSpindleBySpindleId( "BearingInformation" )%></textarea></td></tr>
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
     <tr><td class="input_1" colspan=4><textarea class="forminput" name="txtGeneralSpindleNotes" onchange="this.isDirty=true;" fieldName="General Spindle Notes" rows=24 wrap=soft><%=rsGetSpindleBySpindleId( "GeneralSpindleNotes" )%></textarea></td></tr>
 </table>
</td></tr>

<!-- Bottom Rule -->
<tr><td height=12>
 <TABLE class="" width=100% cellpadding=0 cellspacing=0>
   <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
 </TABLE>
</td></tr>
<!-- End Bottom Rule -->
 
<tr><td align=center>
 	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=256 align=right>
          <tr><td width=64 align=right><a href="javascript:validateForm(frmSpindle);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmSpindle.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:PrintSpindle( '<%=intSpindleId%>' );"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:GetDirty(frmSpindle);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
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