<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetQCInspection, cmdGetQCInspection
dim cmdEditQCInspection
dim intQCInspectionId, intWorkOrderId, intWorkOrderNumber, intRowNumber, strCurrentPage

dim rsGetSpindleBySpindleId, cmdGetSpindleBySpindleId, cmdEditSpindle

set cnnSQLConnection								= nothing
set rsGetQCInspection								= nothing
set cmdGetQCInspection							= nothing
set cmdEditQCInspection							= nothing


set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetQCInspection							= Server.CreateObject( "ADODB.Command" )

set rsGetSpindleBySpindleId				= nothing
set cmdGetSpindleBySpindleId			= nothing

set cmdGetSpindleBySpindleId	= Server.CreateObject( "ADODB.Command" )


cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then
	
		set cmdEditQCInspection	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQCInspection.ActiveConnection	= cnnSQLConnection
		'cmdEditQCInspection.CommandText			= "{? = call spEditQCInspection( ?,?,?,?,?,? ) }"
		cmdEditQCInspection.CommandText			= "spEditQCInspection"
		cmdEditQCInspection.CommandType			= adCmdStoredProc
		
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
		
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RPM_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Weight_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Vibration_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@VibrationRear_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantFlow_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@GSE_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@GSERear_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ShaftTemp_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantTempIncomingSet_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantTempActual_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@FrontTemp_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RearTemp_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantPressureIncoming_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@BreakInTime_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolingMethod_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Volts_In", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@HP_In", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Amps_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Phase_In", adInteger, adParamInput, 1 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Hz_In", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Thermistor_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Poles_In", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@AmpDraw_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ConnectorCtrl_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ConnectorPower_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Converter_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolHolder_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@PullPin_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@EMDimension_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@EjectionPath_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolOutPressure_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ReturnPressure_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@DrawbarForce_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolChangeFunction_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ProximitySwitchFunction_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Lubrication_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Grease_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilMist_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilJet_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilGreaseType_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@IntervalDPM_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@MainPressure_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@TubePressure_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Preload_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RadialPlay_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@AxialPlay_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFrontLocation_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront2_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront2Location_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRearLocation_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear2_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear2Location_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolContact_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolContactRear_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolGap_In", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolGapRear_In", adVarChar, adParamInput, 100 )
		
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RPM_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Weight_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Vibration_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@VibrationRear_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantFlow_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@GSE_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@GSERear_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ShaftTemp_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantTempIncomingSet_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantTempActual_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@FrontTemp_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RearTemp_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolantPressureIncoming_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@BreakInTime_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@CoolingMethod_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Volts_Final", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@HP_Final", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Amps_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Phase_Final", adInteger, adParamInput, 1 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Hz_Final", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Thermistor_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Poles_Final", adInteger, adParamInput, 4 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@AmpDraw_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ConnectorCtrl_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ConnectorPower_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Converter_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolHolder_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@PullPin_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@EMDimension_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@EjectionPath_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolOutPressure_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ReturnPressure_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@DrawbarForce_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolChangeFunction_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ProximitySwitchFunction_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Lubrication_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Grease_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilMist_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilJet_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@OilGreaseType_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@IntervalDPM_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@MainPressure_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@TubePressure_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@Preload_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RadialPlay_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@AxialPlay_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFrontLocation_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront2_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutFront2Location_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRearLocation_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear2_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@RunoutRear2Location_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolContact_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolContactRear_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolGap_Final", adVarChar, adParamInput, 100 )
		cmdEditQCInspection.Parameters.Append cmdEditQCInspection.CreateParameter( "@ToolGapRear_Final", adVarChar, adParamInput, 100 )
		
			
		if( Request.Form( "txtWorkOrderId" ) <> "" ) then cmdEditQCInspection.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderId" )
		
		if( Request.Form( "txtRPM_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RPM_In" ) = Request.Form( "txtRPM_In" )
		if( Request.Form( "txtWeight_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Weight_In" ) = Request.Form( "txtWeight_In" )
		if( Request.Form( "txtVibration_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Vibration_In" ) = Request.Form( "txtVibration_In" )
		if( Request.Form( "txtVibrationRear_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@VibrationRear_In" ) = Request.Form( "txtVibrationRear_In" )
		if( Request.Form( "txtCoolantFlow_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantFlow_In" ) = Request.Form( "txtCoolantFlow_In" )
		if( Request.Form( "txtGSE_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@GSE_In" ) = Request.Form( "txtGSE_In" )
		if( Request.Form( "txtGSERear_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@GSERear_In" ) = Request.Form( "txtGSERear_In" )
		if( Request.Form( "txtShaftTemp_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ShaftTemp_In" ) = Request.Form( "txtShaftTemp_In" )
		if( Request.Form( "txtCoolantTempIncomingSet_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantTempIncomingSet_In" ) = Request.Form( "txtCoolantTempIncomingSet_In" )
		if( Request.Form( "txtCoolantTempActual_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantTempActual_In" ) = Request.Form( "txtCoolantTempActual_In" )
		if( Request.Form( "txtFrontTemp_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@FrontTemp_In" ) = Request.Form( "txtFrontTemp_In" )
		if( Request.Form( "txtRearTemp_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RearTemp_In" ) = Request.Form( "txtRearTemp_In" )
		if( Request.Form( "txtCoolantPressureIncoming_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantPressureIncoming_In" ) = Request.Form( "txtCoolantPressureIncoming_In" )
		if( Request.Form( "txtBreakInTime_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@BreakInTime_In" ) = Request.Form( "txtBreakInTime_In" )
		if( Request.Form( "txtCoolingMethod_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolingMethod_In" ) = Request.Form( "txtCoolingMethod_In" )
		if( Request.Form( "txtVolts_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Volts_In" ) = Request.Form( "txtVolts_In" )
		if( Request.Form( "txtHP_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@HP_In" ) = Request.Form( "txtHP_In" )
		if( Request.Form( "txtAmps_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Amps_In" ) = Request.Form( "txtAmps_In" )
		if( Request.Form( "txtPhase_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Phase_In" ) = Request.Form( "txtPhase_In" )
		if( Request.Form( "txtHz_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Hz_In" ) = Request.Form( "txtHz_In" )
		if( Request.Form( "txtThermistor_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Thermistor_In" ) = Request.Form( "txtThermistor_In" )
		if( Request.Form( "txtPoles_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Poles_In" ) = Request.Form( "txtPoles_In" )
		if( Request.Form( "txtAmpDraw_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@AmpDraw_In" ) = Request.Form( "txtAmpDraw_In" )
		if( Request.Form( "txtConnectorCtrl_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ConnectorCtrl_In" ) = Request.Form( "txtConnectorCtrl_In" )
		if( Request.Form( "txtConnectorPower_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ConnectorPower_In" ) = Request.Form( "txtConnectorPower_In" )
		if( Request.Form( "txtConverter_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Converter_In" ) = Request.Form( "txtConverter_In" )
		if( Request.Form( "txtToolHolder_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolHolder_In" ) = Request.Form( "txtToolHolder_In" )
		if( Request.Form( "txtPullPin_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@PullPin_In" ) = Request.Form( "txtPullPin_In" )
		if( Request.Form( "txtEMDimension_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@EMDimension_In" ) = Request.Form( "txtEMDimension_In" )
		if( Request.Form( "txtEjectionPath_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@EjectionPath_In" ) = Request.Form( "txtEjectionPath_In" )
		if( Request.Form( "txtToolOutPressure_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolOutPressure_In" ) = Request.Form( "txtToolOutPressure_In" )
		if( Request.Form( "txtReturnPressure_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ReturnPressure_In" ) = Request.Form( "txtReturnPressure_In" )
		if( Request.Form( "txtDrawbarForce_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@DrawbarForce_In" ) = Request.Form( "txtDrawbarForce_In" )
		if( Request.Form( "txtToolChangeFunction_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolChangeFunction_In" ) = Request.Form( "txtToolChangeFunction_In" )
		if( Request.Form( "txtProximitySwitchFunction_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ProximitySwitchFunction_In" ) = Request.Form( "txtProximitySwitchFunction_In" )
		if( Request.Form( "txtLubrication_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Lubrication_In" ) = Request.Form( "txtLubrication_In" )
		if( Request.Form( "txtGrease_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Grease_In" ) = Request.Form( "txtGrease_In" )
		if( Request.Form( "txtOilMist_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilMist_In" ) = Request.Form( "txtOilMist_In" )
		if( Request.Form( "txtOilJet_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilJet_In" ) = Request.Form( "txtOilJet_In" )
		if( Request.Form( "txtOilGreaseType_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilGreaseType_In" ) = Request.Form( "txtOilGreaseType_In" )
		if( Request.Form( "txtIntervalDPM_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@IntervalDPM_In" ) = Request.Form( "txtIntervalDPM_In" )
		if( Request.Form( "txtMainPressure_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@MainPressure_In" ) = Request.Form( "txtMainPressure_In" )
		if( Request.Form( "txtTubePressure_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@TubePressure_In" ) = Request.Form( "txtTubePressure_In" )
		if( Request.Form( "txtPreload_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@Preload_In" ) = Request.Form( "txtPreload_In" )
		if( Request.Form( "txtRadialPlay_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RadialPlay_In" ) = Request.Form( "txtRadialPlay_In" )
		if( Request.Form( "txtAxialPlay_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@AxialPlay_In" ) = Request.Form( "txtAxialPlay_In" )
		if( Request.Form( "txtRunoutFront_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront_In" ) = Request.Form( "txtRunoutFront_In" )
		if( Request.Form( "txtRunoutFrontLocation_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFrontLocation_In" ) = Request.Form( "txtRunoutFrontLocation_In" )
		if( Request.Form( "txtRunoutFront2_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront2_In" ) = Request.Form( "txtRunoutFront2_In" )
		if( Request.Form( "txtRunoutFront2Location_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront2Location_In" ) = Request.Form( "txtRunoutFront2Location_In" )
		if( Request.Form( "txtRunoutRear_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear_In" ) = Request.Form( "txtRunoutRear_In" )
		if( Request.Form( "txtRunoutRearLocation_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRearLocation_In" ) = Request.Form( "txtRunoutRearLocation_In" )
		if( Request.Form( "txtRunoutRear2_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear2_In" ) = Request.Form( "txtRunoutRear2_In" )
		if( Request.Form( "txtRunoutRear2Location_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear2Location_In" ) = Request.Form( "txtRunoutRear2Location_In" )
		if( Request.Form( "txtToolContact_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolContact_In" ) = Request.Form( "txtToolContact_In" )
		if( Request.Form( "txtToolContactRear_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolContactRear_In" ) = Request.Form( "txtToolContactRear_In" )
		if( Request.Form( "txtToolGap_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolGap_In" ) = Request.Form( "txtToolGap_In" )
		if( Request.Form( "txtToolGapRear_In" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolGapRear_In" ) = Request.Form( "txtToolGapRear_In" )
		
		if( Request.Form( "txtRPM_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RPM_Final" ) = Request.Form( "txtRPM_Final" )
		if( Request.Form( "txtWeight_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Weight_Final" ) = Request.Form( "txtWeight_Final" )
		if( Request.Form( "txtVibration_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Vibration_Final" ) = Request.Form( "txtVibration_Final" )
		if( Request.Form( "txtVibrationRear_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@VibrationRear_Final" ) = Request.Form( "txtVibrationRear_Final" )
		if( Request.Form( "txtCoolantFlow_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantFlow_Final" ) = Request.Form( "txtCoolantFlow_Final" )
		if( Request.Form( "txtGSE_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@GSE_Final" ) = Request.Form( "txtGSE_Final" )
		if( Request.Form( "txtGSERear_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@GSERear_Final" ) = Request.Form( "txtGSERear_Final" )
		if( Request.Form( "txtShaftTemp_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ShaftTemp_Final" ) = Request.Form( "txtShaftTemp_Final" )
		if( Request.Form( "txtCoolantTempIncomingSet_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantTempIncomingSet_Final" ) = Request.Form( "txtCoolantTempIncomingSet_Final" )
		if( Request.Form( "txtCoolantTempActual_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantTempActual_Final" ) = Request.Form( "txtCoolantTempActual_Final" )
		if( Request.Form( "txtFrontTemp_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@FrontTemp_Final" ) = Request.Form( "txtFrontTemp_Final" )
		if( Request.Form( "txtRearTemp_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RearTemp_Final" ) = Request.Form( "txtRearTemp_Final" )
		if( Request.Form( "txtCoolantPressureIncoming_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolantPressureIncoming_Final" ) = Request.Form( "txtCoolantPressureIncoming_Final" )
		if( Request.Form( "txtBreakInTime_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@BreakInTime_Final" ) = Request.Form( "txtBreakInTime_Final" )
		if( Request.Form( "txtCoolingMethod_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@CoolingMethod_Final" ) = Request.Form( "txtCoolingMethod_Final" )
		if( Request.Form( "txtVolts_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Volts_Final" ) = Request.Form( "txtVolts_Final" )
		if( Request.Form( "txtHP_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@HP_Final" ) = Request.Form( "txtHP_Final" )
		if( Request.Form( "txtAmps_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Amps_Final" ) = Request.Form( "txtAmps_Final" )
		if( Request.Form( "txtPhase_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Phase_Final" ) = Request.Form( "txtPhase_Final" )
		if( Request.Form( "txtHz_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Hz_Final" ) = Request.Form( "txtHz_Final" )
		if( Request.Form( "txtThermistor_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Thermistor_Final" ) = Request.Form( "txtThermistor_Final" )
		if( Request.Form( "txtPoles_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Poles_Final" ) = Request.Form( "txtPoles_Final" )
		if( Request.Form( "txtAmpDraw_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@AmpDraw_Final" ) = Request.Form( "txtAmpDraw_Final" )
		if( Request.Form( "txtConnectorCtrl_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ConnectorCtrl_Final" ) = Request.Form( "txtConnectorCtrl_Final" )
		if( Request.Form( "txtConnectorPower_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ConnectorPower_Final" ) = Request.Form( "txtConnectorPower_Final" )
		if( Request.Form( "txtConverter_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Converter_Final" ) = Request.Form( "txtConverter_Final" )
		if( Request.Form( "txtToolHolder_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolHolder_Final" ) = Request.Form( "txtToolHolder_Final" )
		if( Request.Form( "txtPullPin_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@PullPin_Final" ) = Request.Form( "txtPullPin_Final" )
		if( Request.Form( "txtEMDimension_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@EMDimension_Final" ) = Request.Form( "txtEMDimension_Final" )
		if( Request.Form( "txtEjectionPath_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@EjectionPath_Final" ) = Request.Form( "txtEjectionPath_Final" )
		if( Request.Form( "txtToolOutPressure_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolOutPressure_Final" ) = Request.Form( "txtToolOutPressure_Final" )
		if( Request.Form( "txtReturnPressure_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ReturnPressure_Final" ) = Request.Form( "txtReturnPressure_Final" )
		if( Request.Form( "txtDrawbarForce_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@DrawbarForce_Final" ) = Request.Form( "txtDrawbarForce_Final" )
		if( Request.Form( "txtToolChangeFunction_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolChangeFunction_Final" ) = Request.Form( "txtToolChangeFunction_Final" )
		if( Request.Form( "txtProximitySwitchFunction_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ProximitySwitchFunction_Final" ) = Request.Form( "txtProximitySwitchFunction_Final" )
		if( Request.Form( "txtLubrication_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Lubrication_Final" ) = Request.Form( "txtLubrication_Final" )
		if( Request.Form( "txtGrease_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Grease_Final" ) = Request.Form( "txtGrease_Final" )
		if( Request.Form( "txtOilMist_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilMist_Final" ) = Request.Form( "txtOilMist_Final" )
		if( Request.Form( "txtOilJet_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilJet_Final" ) = Request.Form( "txtOilJet_Final" )
		if( Request.Form( "txtOilGreaseType_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@OilGreaseType_Final" ) = Request.Form( "txtOilGreaseType_Final" )
		if( Request.Form( "txtIntervalDPM_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@IntervalDPM_Final" ) = Request.Form( "txtIntervalDPM_Final" )
		if( Request.Form( "txtMainPressure_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@MainPressure_Final" ) = Request.Form( "txtMainPressure_Final" )
		if( Request.Form( "txtTubePressure_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@TubePressure_Final" ) = Request.Form( "txtTubePressure_Final" )
		if( Request.Form( "txtPreload_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@Preload_Final" ) = Request.Form( "txtPreload_Final" )
		if( Request.Form( "txtRadialPlay_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RadialPlay_Final" ) = Request.Form( "txtRadialPlay_Final" )
		if( Request.Form( "txtAxialPlay_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@AxialPlay_Final" ) = Request.Form( "txtAxialPlay_Final" )
		if( Request.Form( "txtRunoutFront_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront_Final" ) = Request.Form( "txtRunoutFront_Final" )
		if( Request.Form( "txtRunoutFrontLocation_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFrontLocation_Final" ) = Request.Form( "txtRunoutFrontLocation_Final" )
		if( Request.Form( "txtRunoutFront2_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront2_Final" ) = Request.Form( "txtRunoutFront2_Final" )
		if( Request.Form( "txtRunoutFront2Location_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutFront2Location_Final" ) = Request.Form( "txtRunoutFront2Location_Final" )
		if( Request.Form( "txtRunoutRear_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear_Final" ) = Request.Form( "txtRunoutRear_Final" )
		if( Request.Form( "txtRunoutRearLocation_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRearLocation_Final" ) = Request.Form( "txtRunoutRearLocation_Final" )
		if( Request.Form( "txtRunoutRear2_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear2_Final" ) = Request.Form( "txtRunoutRear2_Final" )
		if( Request.Form( "txtRunoutRear2Location_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@RunoutRear2Location_Final" ) = Request.Form( "txtRunoutRear2Location_Final" )
		if( Request.Form( "txtToolContact_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolContact_Final" ) = Request.Form( "txtToolContact_Final" )
		if( Request.Form( "txtToolContactRear_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolContactRear_Final" ) = Request.Form( "txtToolContactRear_Final" )
		if( Request.Form( "txtToolGap_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolGap_Final" ) = Request.Form( "txtToolGap_Final" )
		if( Request.Form( "txtToolGapRear_Final" ) <> "" ) then cmdEditQCInspection.Parameters( "@ToolGapRear_Final" ) = Request.Form( "txtToolGapRear_Final" )		
		
		on error resume next
		cmdEditQCInspection.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if

		on error goto 0
		
		intWorkOrderId = Request.Form( "txtWorkOrderId" )
else
	intWorkOrderId = Request.QueryString( "intWorkOrderId" )
end if

cmdGetQCInspection.ActiveConnection = cnnSQLConnection
cmdGetQCInspection.CommandText			= "spGetQCInspectionByWorkOrderId"
cmdGetQCInspection.CommandType			= adCmdStoredProc
cmdGetQCInspection.Parameters.Append cmdGetQCInspection.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQCInspection.Parameters.Append cmdGetQCInspection.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4)
cmdGetQCInspection.Parameters( "@WorkOrderId" )	= intWorkOrderId

on error resume next
set rsGetQCInspection = cmdGetQCInspection.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0

'**********************************************************************************
if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<HTML>
<HEAD>
<TITLE>Centerline - QC Inspection Details</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/print.css" media="print" TITLE="print_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	function OpenSpindle( strSpindleId )
	{
	  OpenWindow( 'spindle_details.asp?intSpindleId='+strSpindleId,'CenterlineSpindle'+strSpindleId,'width=660,height=540,left='+GetCenterX( 660 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
  </script>
</head>
<body topmargin=0 leftmargin=0 id="qcinspectionsheet">

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetQCInspection is nothing ) then
		if( rsGetQCInspection.State = adStateOpen ) then
			if( not rsGetQCInspection.EOF ) then%>
						
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>QC Inspection</td>
     <td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>W.O. #&nbsp;&nbsp;<a href="javascript:OpenWorkOrder('<%=intWorkOrderId%>');"><%=rsGetQCInspection( "WorkOrderNumber" )%></a>&nbsp;&nbsp;:&nbsp;&nbsp;Spindle&nbsp;&nbsp;<a href="javascript:OpenSpindle('<%=rsGetQCInspection("SpindleId")%>');"><%=rsGetQCInspection("SpindleType")%></a></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- Cost Form -->
<form id="frmQCInspection" name="frmQCInspection" action="" method="post">
<input type="hidden" name="txtWorkOrderId" value="<%=intWorkOrderId%>">

<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>	
		<!-- table heading -->
		<thead>
				<th class="itemlist" width="192">&nbsp;</th>
				<th class="itemlist aleft" width="98">Default Info</th>
				<th class="itemlist" width="110">Incoming</th>
				<th class="itemlist" width="110">Final</th>
		</thead>
		
		<tr><td class="generic_label" width=192>Speed (RPM)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RPM" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRPM_In" onchange="this.isDirty=true;" fieldName="RPM Incoming" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RPM_In" ))%>" ID="Text2"></td>
         <td width=110><input type="text" class="forminput" name="txtRPM_Final" onchange="this.isDirty=true;" fieldName="RPM Final" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RPM_Final" ))%>" ID="Text3"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Weight</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Weight" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtWeight_In" onchange="this.isDirty=true;" fieldName="Weight Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Weight_In" ))%>" ID="Text1"></td>
         <td width=110><input type="text" class="forminput" name="txtWeight_Final" onchange="this.isDirty=true;" fieldName="Weight Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Weight_Final" ))%>" ID="Text4"></td></tr>
		<tr><td class="generic_label" width=192>Vibration Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtVibration_In" onchange="this.isDirty=true;" fieldName="Vibration Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration_In" ))%>" ID="Text5"></td>
         <td width=110><input type="text" class="forminput" name="txtVibration_Final" onchange="this.isDirty=true;" fieldName="Vibration Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration_Final" ))%>" ID="Text6"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Vibration Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtVibrationRear_In" onchange="this.isDirty=true;" fieldName="VibrationRear Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear_In" ))%>" ID="Text7"></td>
         <td width=110><input type="text" class="forminput" name="txtVibrationRear_Final" onchange="this.isDirty=true;" fieldName="VibrationRear Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear_Final" ))%>" ID="Text8"></td></tr>
		<tr><td class="generic_label" width=192>GSE Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "GSE" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtGSE_In" onchange="this.isDirty=true;" fieldName="GSE Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "GSE_In" ))%>" ID="Text11"></td>
         <td width=110><input type="text" class="forminput" name="txtGSE_Final" onchange="this.isDirty=true;" fieldName="GSE Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "GSE_Final" ))%>" ID="Text12"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>GSE Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtGSERear_In" onchange="this.isDirty=true;" fieldName="GSERear Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear_In" ))%>" ID="Text13"></td>
         <td width=110><input type="text" class="forminput" name="txtGSERear_Final" onchange="this.isDirty=true;" fieldName="GSERear Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear_Final" ))%>" ID="Text14"></td></tr>
		<tr><td class="generic_label" width=192>Shaft Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtShaftTemp_In" onchange="this.isDirty=true;" fieldName="ShaftTemp Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp_In" ))%>" ID="Text15"></td>
         <td width=110><input type="text" class="forminput" name="txtShaftTemp_Final" onchange="this.isDirty=true;" fieldName="ShaftTemp Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp_Final" ))%>" ID="Text16"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Front Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtFrontTemp_In" onchange="this.isDirty=true;" fieldName="FrontTemp Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp_In" ))%>" ID="Text21"></td>
         <td width=110><input type="text" class="forminput" name="txtFrontTemp_Final" onchange="this.isDirty=true;" fieldName="FrontTemp Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp_Final" ))%>" ID="Text22"></td></tr>
		<tr><td class="generic_label" width=192>Rear Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRearTemp_In" onchange="this.isDirty=true;" fieldName="RearTemp Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp_In" ))%>" ID="Text23"></td>
         <td width=110><input type="text" class="forminput" name="txtRearTemp_Final" onchange="this.isDirty=true;" fieldName="RearTemp Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp_Final" ))%>" ID="Text24"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Coolant Flow (GPM)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtCoolantFlow_In" onchange="this.isDirty=true;" fieldName="CoolantFlow Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow_In" ))%>" ID="Text9"></td>
         <td width=110><input type="text" class="forminput" name="txtCoolantFlow_Final" onchange="this.isDirty=true;" fieldName="CoolantFlow Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow_Final" ))%>" ID="Text10"></td></tr>
		<tr><td class="generic_label" width=192>Coolant Temp Incoming Set</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtCoolantTempIncomingSet_In" onchange="this.isDirty=true;" fieldName="CoolantTempIncomingSet Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet_In" ))%>" ID="Text17"></td>
         <td width=110><input type="text" class="forminput" name="txtCoolantTempIncomingSet_Final" onchange="this.isDirty=true;" fieldName="CoolantTempIncoming Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet_Final" ))%>" ID="Text18"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Coolant Temp Actual</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtCoolantTempActual_In" onchange="this.isDirty=true;" fieldName="CoolantTempActual Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual_In" ))%>" ID="Text19"></td>
         <td width=110><input type="text" class="forminput" name="txtCoolantTempActual_Final" onchange="this.isDirty=true;" fieldName="CoolantTempActual Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual_Final" ))%>" ID="Text20"></td></tr>
		<tr><td class="generic_label" width=192>Coolant Pressure Incoming</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtCoolantPressureIncoming_In" onchange="this.isDirty=true;" fieldName="CoolantPressureIncoming Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming_In" ))%>" ID="Text25"></td>
         <td width=110><input type="text" class="forminput" name="txtCoolantPressureIncoming_Final" onchange="this.isDirty=true;" fieldName="CoolantPressureIncoming Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming_Final" ))%>" ID="Text26"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Cooling Method</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtCoolingMethod_In" onchange="this.isDirty=true;" fieldName="CoolingMethod Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod_In" ))%>" ID="Text29"></td>
         <td width=110><input type="text" class="forminput" name="txtCoolingMethod_Final" onchange="this.isDirty=true;" fieldName="CoolingMethod Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod_Final" ))%>" ID="Text30"></td></tr>
		
		<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>

		<tr class="divider"><td class="generic_label" width=192>Volts</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Volts" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtVolts_In" onchange="this.isDirty=true;" fieldName="Volts Incoming" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Volts_In" ))%>" ID="Text31"></td>
         <td width=110><input type="text" class="forminput" name="txtVolts_Final" onchange="this.isDirty=true;" fieldName="Volts Final" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Volts_Final" ))%>" ID="Text32"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Power (HP)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "HP" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtHP_In" onchange="this.isDirty=true;" fieldName="HP Incoming" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "HP_In" ))%>" ID="Text33"></td>
         <td width=110><input type="text" class="forminput" name="txtHP_Final" onchange="this.isDirty=true;" fieldName="HP Final" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "HP_Final" ))%>" ID="Text34"></td></tr>
		<tr><td class="generic_label" width=192>Amps</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Amps" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtAmps_In" onchange="this.isDirty=true;" fieldName="Amps Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Amps_In" ))%>" ID="Text35"></td>
         <td width=110><input type="text" class="forminput" name="txtAmps_Final" onchange="this.isDirty=true;" fieldName="Amps Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Amps_Final" ))%>" ID="Text36"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Phase</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Phase" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtPhase_In" onchange="this.isDirty=true;" fieldName="Phase Incoming" maxlength=1 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Phase_In" ))%>" ID="Text37"></td>
         <td width=110><input type="text" class="forminput" name="txtPhase_Final" onchange="this.isDirty=true;" fieldName="Phase Final" maxlength=1 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Phase_Final" ))%>" ID="Text38"></td></tr>
		<tr><td class="generic_label" width=192>Hz</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Hz" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtHz_In" onchange="this.isDirty=true;" fieldName="Hz Incoming" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Hz_In" ))%>" ID="Text39"></td>
         <td width=110><input type="text" class="forminput" name="txtHz_Final" onchange="this.isDirty=true;" fieldName="Hz Final" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Hz_Final" ))%>" ID="Text40"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Thermistor</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtThermistor_In" onchange="this.isDirty=true;" fieldName="Thermistor Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor_In" ))%>" ID="Text41"></td>
         <td width=110><input type="text" class="forminput" name="txtThermistor_Final" onchange="this.isDirty=true;" fieldName="Thermistor Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor_Final" ))%>" ID="Text42"></td></tr>
		<tr><td class="generic_label" width=192>Poles</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Poles" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtPoles_In" onchange="this.isDirty=true;" fieldName="Poles Incoming" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Poles_In" ))%>" ID="Text43"></td>
         <td width=110><input type="text" class="forminput" name="txtPoles_Final" onchange="this.isDirty=true;" fieldName="Poles Final" maxlength=10 isNumber="true" value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Poles_Final" ))%>" ID="Text44"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Amp Draw</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtAmpDraw_In" onchange="this.isDirty=true;" fieldName="Amp Draw Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw_In" ))%>" ID="Text45"></td>
         <td width=110><input type="text" class="forminput" name="txtAmpDraw_Final" onchange="this.isDirty=true;" fieldName="Amp Draw Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw_Final" ))%>" ID="Text46"></td></tr>
		<tr><td class="generic_label" width=192>Connector Ctrl</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtConnectorCtrl_In" onchange="this.isDirty=true;" fieldName="ConnectorCtrl Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl_In" ))%>" ID="Text47"></td>
         <td width=110><input type="text" class="forminput" name="txtConnectorCtrl_Final" onchange="this.isDirty=true;" fieldName="ConnectorCtrl Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl_Final" ))%>" ID="Text48"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Connector Power</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtConnectorPower_In" onchange="this.isDirty=true;" fieldName="ConnectorPower Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower_In" ))%>" ID="Text49"></td>
         <td width=110><input type="text" class="forminput" name="txtConnectorPower_Final" onchange="this.isDirty=true;" fieldName="ConnectorPower Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower_Final" ))%>" ID="Text50"></td></tr>
		<tr><td class="generic_label" width=192>Converter</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Converter" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtConverter_In" onchange="this.isDirty=true;" fieldName="Converter Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Converter_In" ))%>" ID="Text51"></td>
         <td width=110><input type="text" class="forminput" name="txtConverter_Final" onchange="this.isDirty=true;" fieldName="Converter Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Converter_Final" ))%>" ID="Text52"></td></tr>
		
		<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Tool Type</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolHolder_In" onchange="this.isDirty=true;" fieldName="ToolHolder Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder_In" ))%>" ID="Text53"></td>
         <td width=110><input type="text" class="forminput" name="txtToolHolder_Final" onchange="this.isDirty=true;" fieldName="ToolHolder Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder_Final" ))%>" ID="Text54"></td></tr>
		<tr><td class="generic_label" width=192>Pull Pin</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtPullPin_In" onchange="this.isDirty=true;" fieldName="PullPin Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin_In" ))%>" ID="Text55"></td>
         <td width=110><input type="text" class="forminput" name="txtPullPin_Final" onchange="this.isDirty=true;" fieldName="PullPin Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin_Final" ))%>" ID="Text56"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>EM Dimension</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtEMDimension_In" onchange="this.isDirty=true;" fieldName="EMDimension Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension_In" ))%>" ID="Text57"></td>
         <td width=110><input type="text" class="forminput" name="txtEMDimension_Final" onchange="this.isDirty=true;" fieldName="EMDimension Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension_Final" ))%>" ID="Text58"></td></tr>
		<tr><td class="generic_label" width=192>Ejection Path</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtEjectionPath_In" onchange="this.isDirty=true;" fieldName="EjectionPath Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath_In" ))%>" ID="Text59"></td>
         <td width=110><input type="text" class="forminput" name="txtEjectionPath_Final" onchange="this.isDirty=true;" fieldName="EjectionPath Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath_Final" ))%>" ID="Text60"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Tool Out Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolOutPressure_In" onchange="this.isDirty=true;" fieldName="ToolOutPressure Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure_In" ))%>" ID="Text115"></td>
         <td width=110><input type="text" class="forminput" name="txtToolOutPressure_Final" onchange="this.isDirty=true;" fieldName="ToolOutPressure Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure_Final" ))%>" ID="Text116"></td></tr>
		<tr><td class="generic_label" width=192>Return Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtReturnPressure_In" onchange="this.isDirty=true;" fieldName="ReturnPressure Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure_In" ))%>" ID="Text61"></td>
         <td width=110><input type="text" class="forminput" name="txtReturnPressure_Final" onchange="this.isDirty=true;" fieldName="ReturnPressure Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure_Final" ))%>" ID="Text62"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Drawforce</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtDrawbarForce_In" onchange="this.isDirty=true;" fieldName="DrawbarForce Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce_In" ))%>" ID="Text63"></td>
         <td width=110><input type="text" class="forminput" name="txtDrawbarForce_Final" onchange="this.isDirty=true;" fieldName="DrawbarForce Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce_Final" ))%>" ID="Text64"></td></tr>
		<tr><td class="generic_label" width=192>Tool Change Function</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolChangeFunction_In" onchange="this.isDirty=true;" fieldName="ToolChangeFunction Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction_In" ))%>" ID="Text65"></td>
         <td width=110><input type="text" class="forminput" name="txtToolChangeFunction_Final" onchange="this.isDirty=true;" fieldName="ToolChangeFunction Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction_Final" ))%>" ID="Text66"></td></tr>
		<tr class="evenrow pagebreak"><td class="generic_label" width=192>Proximity Switch Function</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtProximitySwitchFunction_In" onchange="this.isDirty=true;" fieldName="Proximity Switch Function Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction_In" ))%>" ID="Text67"></td>
         <td width=110><input type="text" class="forminput" name="txtProximitySwitchFunction_Final" onchange="this.isDirty=true;" fieldName="ProximitySwitchFunction Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction_Final" ))%>" ID="Text68"></td></tr>
		
		<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>

		<tr><td class="generic_label" width=192>Lubrication</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtLubrication_In" onchange="this.isDirty=true;" fieldName="Lubrication Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication_In" ))%>" ID="Text69"></td>
         <td width=110><input type="text" class="forminput" name="txtLubrication_Final" onchange="this.isDirty=true;" fieldName="Lubrication Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication_Final" ))%>" ID="Text70"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Grease</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Grease" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtGrease_In" onchange="this.isDirty=true;" fieldName="Grease Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Grease_In" ))%>" ID="Text71"></td>
         <td width=110><input type="text" class="forminput" name="txtGrease_Final" onchange="this.isDirty=true;" fieldName="Grease Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Grease_Final" ))%>" ID="Text72"></td></tr>
		<tr><td class="generic_label" width=192>Oil Mist</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtOilMist_In" onchange="this.isDirty=true;" fieldName="OilMist Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist_In" ))%>" ID="Text73"></td>
         <td width=110><input type="text" class="forminput" name="txtOilMist_Final" onchange="this.isDirty=true;" fieldName="OilMist Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist_Final" ))%>" ID="Text74"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Oil Jet</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtOilJet_In" onchange="this.isDirty=true;" fieldName="OilJet Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet_In" ))%>" ID="Text75"></td>
         <td width=110><input type="text" class="forminput" name="txtOilJet_Final" onchange="this.isDirty=true;" fieldName="OilJet Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet_Final" ))%>" ID="Text76"></td></tr>
		<tr><td class="generic_label" width=192>Oil Grease Type</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtOilGreaseType_In" onchange="this.isDirty=true;" fieldName="OilGreaseType Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType_In" ))%>" ID="Text77"></td>
         <td width=110><input type="text" class="forminput" name="txtOilGreaseType_Final" onchange="this.isDirty=true;" fieldName="OilGreaseType Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType_Final" ))%>" ID="Text78"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Interval DPM</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtIntervalDPM_In" onchange="this.isDirty=true;" fieldName="IntervalDPM Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM_In" ))%>" ID="Text79"></td>
         <td width=110><input type="text" class="forminput" name="txtIntervalDPM_Final" onchange="this.isDirty=true;" fieldName="IntervalDPM Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM_Final" ))%>" ID="Text80"></td></tr>
		<tr><td class="generic_label" width=192>Main Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtMainPressure_In" onchange="this.isDirty=true;" fieldName="MainPressure Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure_In" ))%>" ID="Text81"></td>
         <td width=110><input type="text" class="forminput" name="txtMainPressure_Final" onchange="this.isDirty=true;" fieldName="MainPressure Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure_Final" ))%>" ID="Text82"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Tube Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtTubePressure_In" onchange="this.isDirty=true;" fieldName="TubePressure Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure_In" ))%>" ID="Text83"></td>
         <td width=110><input type="text" class="forminput" name="txtTubePressure_Final" onchange="this.isDirty=true;" fieldName="TubePressure Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure_Final" ))%>" ID="Text84"></td></tr>
		
		<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>

		<tr><td class="generic_label" width=192>Preload</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Preload" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtPreload_In" onchange="this.isDirty=true;" fieldName="Preload Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Preload_In" ))%>" ID="Text85"></td>
         <td width=110><input type="text" class="forminput" name="txtPreload_Final" onchange="this.isDirty=true;" fieldName="Preload Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "Preload_Final" ))%>" ID="Text86"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Radial Play</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRadialPlay_In" onchange="this.isDirty=true;" fieldName="RadialPlay Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay_In" ))%>" ID="Text87"></td>
         <td width=110><input type="text" class="forminput" name="txtRadialPlay_Final" onchange="this.isDirty=true;" fieldName="RadialPlay Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay_Final" ))%>" ID="Text88"></td></tr>
		<tr><td class="generic_label" width=192>Axial Play</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtAxialPlay_In" onchange="this.isDirty=true;" fieldName="AxialPlay Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay_In" ))%>" ID="Text89"></td>
         <td width=110><input type="text" class="forminput" name="txtAxialPlay_Final" onchange="this.isDirty=true;" fieldName="AxialPlay Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay_Final" ))%>" ID="Text90"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Runout Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutFront_In" onchange="this.isDirty=true;" fieldName="RunoutFront Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront_In" ))%>" ID="Text91"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutFront_Final" onchange="this.isDirty=true;" fieldName="RunoutFront Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront_Final" ))%>" ID="Text92"></td></tr>
		<tr><td class="generic_label" width=192>Runout Front Location</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutFrontLocation_In" onchange="this.isDirty=true;" fieldName="RunoutFrontLocation Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation_In" ))%>" ID="Text93"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutFrontLocation_Final" onchange="this.isDirty=true;" fieldName="RunoutFrontLocation Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation_Final" ))%>" ID="Text94"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Runout Front 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutFront2_In" onchange="this.isDirty=true;" fieldName="RunoutFront2 Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2_In" ))%>" ID="Text95"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutFront2_Final" onchange="this.isDirty=true;" fieldName="RunoutFront2 Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2_Final" ))%>" ID="Text96"></td></tr>
		<tr><td class="generic_label" width=192>Runout Front Location 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutFront2Location_In" onchange="this.isDirty=true;" fieldName="RunoutFront2Location Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location_In" ))%>" ID="Text97"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutFront2Location_Final" onchange="this.isDirty=true;" fieldName="RunoutFront2Location Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location_Final" ))%>" ID="Text98"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Runout Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutRear_In" onchange="this.isDirty=true;" fieldName="RunoutRear Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear_In" ))%>" ID="Text99"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutRear_Final" onchange="this.isDirty=true;" fieldName="RunoutRear Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear_Final" ))%>" ID="Text100"></td></tr>
		<tr><td class="generic_label" width=192>Runout Rear Location</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutRearLocation_In" onchange="this.isDirty=true;" fieldName="RunoutRearLocation Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation_In" ))%>" ID="Text101"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutRearLocation_Final" onchange="this.isDirty=true;" fieldName="RunoutRearLocation Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation_Final" ))%>" ID="Text102"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Runout Rear 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutRear2_In" onchange="this.isDirty=true;" fieldName="RunoutRear2 Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2_In" ))%>" ID="Text103"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutRear2_Final" onchange="this.isDirty=true;" fieldName="RunoutRear2 Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2_Final" ))%>" ID="Text104"></td></tr>
		<tr><td class="generic_label" width=192>Runout Rear Location 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtRunoutRear2Location_In" onchange="this.isDirty=true;" fieldName="RunoutRear2Location Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location_In" ))%>" ID="Text105"></td>
         <td width=110><input type="text" class="forminput" name="txtRunoutRear2Location_Final" onchange="this.isDirty=true;" fieldName="RunoutRear2Location Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location_Final" ))%>" ID="Text106"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Tool Contact</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolContact_In" onchange="this.isDirty=true;" fieldName="ToolContact Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact_In" ))%>" ID="Text107"></td>
         <td width=110><input type="text" class="forminput" name="txtToolContact_Final" onchange="this.isDirty=true;" fieldName="ToolContact Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact_Final" ))%>" ID="Text108"></td></tr>
		<tr><td class="generic_label" width=192>Tool Contact Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolContactRear_In" onchange="this.isDirty=true;" fieldName="ToolContactRear Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear_In" ))%>" ID="Text109"></td>
         <td width=110><input type="text" class="forminput" name="txtToolContactRear_Final" onchange="this.isDirty=true;" fieldName="ToolContactRear Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear_Final" ))%>" ID="Text110"></td></tr>
		<tr class="evenrow"><td class="generic_label" width=192>Tool Gap</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolGap_In" onchange="this.isDirty=true;" fieldName="ToolGap Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap_In" ))%>" ID="Text111"></td>
         <td width=110><input type="text" class="forminput" name="txtToolGap_Final" onchange="this.isDirty=true;" fieldName="ToolGap Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap_Final" ))%>" ID="Text112"></td></tr>
		<tr><td class="generic_label" width=192>Tool Gap Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtToolGapRear_In" onchange="this.isDirty=true;" fieldName="ToolGapRear Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear_In" ))%>" ID="Text113"></td>
         <td width=110><input type="text" class="forminput" name="txtToolGapRear_Final" onchange="this.isDirty=true;" fieldName="ToolGapRear Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear_Final" ))%>" ID="Text114"></td></tr>
  		
		<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>
		       
    <tr class="evenrow"><td class="generic_label" width=192>Break In Time</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime" ))%></td>
	     <td width=110><input type="text" class="forminput" name="txtBreakInTime_In" onchange="this.isDirty=true;" fieldName="BreakInTime Incoming" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime_In" ))%>" ID="Text27"></td>
         <td width=110><input type="text" class="forminput" name="txtBreakInTime_Final" onchange="this.isDirty=true;" fieldName="BreakInTime Final" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime_Final" ))%>" ID="Text28"></td></tr>
         
    <tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>
	<tr class="print_only">
    	<td colspan="2" class="generic_label underline">Check Serial Number:</td>
    	<td colspan="2" class="generic_label underline">Baker:</td>
    </tr>
	<tr class="print_only">
		<td colspan="4" class="generic_label"><div class="failure">Incoming Inspection Notes / Failure Analysis:</div></td>
	</tr>
	
	<tr>
		<td class="generic_label" width=192>Balancing Requirements</td>
		<td colspan="3" class="generic_input"><%=rsGetSpindleBySpindleId( "BalancingRequirements" )%></td>
	</tr>
	<tr>
		<td class="generic_label" width=192>BearingInformation</td>
		<td colspan="3" class="generic_input"><%=rsGetSpindleBySpindleId( "BearingInformation" )%></td>
	</tr>
	<tr>
		<td class="generic_label" width=192>General Spindle Notes</td>
		<td colspan="3" class="generic_input"><%=rsGetSpindleBySpindleId( "GeneralSpindleNotes" )%></td>
	</tr>
	
	
	
	
	
	
	
	
	
	
	
	
	
		
	</table>
</td></tr>

<tr><td colspan="4" style="background-color:#006;font-size:.125em;">&nbsp;</td></tr>

<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle ID="Table2">
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=256 align=right ID="Table3">
	          <tr><td width=64 align=right><a href="javascript:window.navigate( '<%=Request.QueryString( "strPrevPage" )%>' );"><IMG src="images/back_button.gif" width=60 height=20 border=0></a></td>
							<td width=64 align=right><a href="javascript:validateForm(frmQCInspection);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmQCInspection.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmQCInspection);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>
<!-- End Form Buttons -->
</form>
<%	set rsGetQCInspection = rsGetQCInspection.NextRecordset
		end if
	end if
end if%>
<!-- End QC Temp Log Form -->

 </body>
</html>
<%
if not( rsGetQCInspection is nothing ) then
	if( rsGetQCInspection.State = adStateOpen ) then
		rsGetQCInspection.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQCInspection				= nothing
set cmdGetQCInspection			= nothing
set cmdEditQCInspection											= nothing
set cnnSQLConnection									= nothing
%>
