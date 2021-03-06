<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetQCInspection, cmdGetQCInspection
dim cmdEditQCInspection
dim intQCInspectionId, intWorkOrderId, intWorkOrderNumber, intRowNumber, strCurrentPage

set cnnSQLConnection								= nothing
set rsGetQCInspection								= nothing
set cmdGetQCInspection							= nothing
set cmdEditQCInspection							= nothing


set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetQCInspection							= Server.CreateObject( "ADODB.Command" )

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


<style type="text/css" media="all">
body, table, th, td {font : .95em verdana;}
th {font-weight : bold;padding-bottom:3px;}
td {padding : 5px 0;}
.evenrow {color:#000;background-color:#bfbfbf;}
td.section {color:#fff;background-color:#006;height:20px;font-size:1.1em;font-weight:bold;}
</style>


</head>
<body topmargin=10 leftmargin=10>

<!-- Page Contents -->
<table border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetQCInspection is nothing ) then
		if( rsGetQCInspection.State = adStateOpen ) then
			if( not rsGetQCInspection.EOF ) then%>
						
<!--
Top Title Bar 
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>QC Inspection</td>
     <td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>W.O. #&nbsp;&nbsp;<a href="javascript:OpenWorkOrder('<%=intWorkOrderId%>');"><%=rsGetQCInspection( "WorkOrderNumber" )%></a>&nbsp;&nbsp;:&nbsp;&nbsp;Spindle&nbsp;&nbsp;<a href="javascript:OpenSpindle('<%=rsGetQCInspection("SpindleId")%>');"><%=rsGetQCInspection("SpindleType")%></a></td></tr>
 </table>
</td></tr>
End Top Title Bar
//-->

<!-- Cost Form -->
<form id="frmQCInspection" name="frmQCInspection" action="" method="post">
<input type="hidden" name="txtWorkOrderId" value="<%=intWorkOrderId%>">

<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>	
		<tr>
				<td class="generic_label">Spindle Type</t>
				<td colspan=3><%=rsGetQCInspection("SpindleType")%></td>
		</tr>
		<tr>
				<td class="generic_label">Serial Number</t>
				<td colspan=3>&nbsp;</td>
		</tr>
		<tr>
				<td class="generic_label">Work Order Number</t>
				<td colspan=3><%=rsGetQCInspection( "WorkOrderNumber" )%></td>
		</tr>

		<!-- Bottom Rule -->
		<tr><td colspan="4"><hr /></td></tr>
		<!-- End Bottom Rule -->
		
		<!-- table heading -->
		<tr>
				<th align="left" width="192">&nbsp;</th>
				<th align="left" width="98">Default</th>
				<th align="left" width="110" style="color:red">Incoming</th>
				<th align="left" width="110" style="color:green">Final</th>
		</tr>

		<!-- Bottom Rule -->
		<tr><td colspan="4" class="section">General</td></tr>
		<!-- End Bottom Rule -->
	
		

		
		<tr><td class="generic_label" width=192>Speed (RPM)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RPM" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RPM_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RPM_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Weight</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Weight" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Weight_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Weight_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Vibration Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Vibration_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Vibration Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "VibrationRear_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>GSE Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "GSE" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "GSE_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "GSE_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>GSE Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "GSERear_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Shaft Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ShaftTemp_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Front Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "FrontTemp_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Rear Temp</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RearTemp_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Coolant Flow (GPM)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantFlow_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Coolant Temp Incoming Set</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempIncomingSet_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Coolant Temp Actual</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantTempActual_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Coolant Pressure Incoming</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolantPressureIncoming_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Cooling Method</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "CoolingMethod_Final" ))%></td></tr>

		
		<!-- Bottom Rule -->
		<tr><td colspan="4" class="section">Electrical</td></tr>
		<!-- End Bottom Rule -->

		<tr><td class="generic_label" width=192>Volts</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Volts" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Volts_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Volts_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Power (HP)</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "HP" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "HP_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "HP_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Amps</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Amps" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Amps_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Amps_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Phase</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Phase" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Phase_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Phase_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Hz</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Hz" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Hz_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Hz_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Thermistor</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Thermistor_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Poles</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Poles" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Poles_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Poles_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Amp Draw</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "AmpDraw_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Connector Ctrl</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorCtrl_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Connector Power</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ConnectorPower_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Converter</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Converter" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Converter_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Converter_Final" ))%></td></tr>

		
		<!-- Bottom Rule -->
		<tr><td colspan="4" class="section">Tool Holding</td></tr>
		<!-- End Bottom Rule -->


		<tr class="evenrow"><td class="generic_label" width=192>Tool Type</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolHolder_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Pull Pin</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "PullPin_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>EM Dimension</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "EMDimension_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Ejection Path</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "EjectionPath_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Tool Out Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolOutPressure_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Return Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ReturnPressure_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Drawforce</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "DrawbarForce_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Tool Change Function</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolChangeFunction_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Proximity Switch Function</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ProximitySwitchFunction_Final" ))%></td></tr>

		
		<!-- Bottom Rule -->
		<tr><td colspan="4" class="section">Lubrication</td></tr>
		<!-- End Bottom Rule -->


		<tr><td class="generic_label" width=192>Lubrication</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Lubrication_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Grease</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Grease" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Grease_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Grease_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Oil Mist</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilMist_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Oil Jet</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilJet_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Oil Grease Type</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "OilGreaseType_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Interval DPM</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "IntervalDPM_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Main Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "MainPressure_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Tube Pressure</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "TubePressure_Final" ))%></td></tr>

		
		<!-- Bottom Rule -->
		<tr><td colspan="4" class="section">Conforminess to Geometry and Stiffitude</td></tr>
		<!-- End Bottom Rule -->


		<tr><td class="generic_label" width=192>Preload</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "Preload" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Preload_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "Preload_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Radial Play</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RadialPlay_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Axial Play</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "AxialPlay_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Runout Front</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Runout Front Location</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFrontLocation_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Runout Front 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Runout Front Location 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutFront2Location_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Runout Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Runout Rear Location</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRearLocation_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Runout Rear 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Runout Rear Location 2</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "RunoutRear2Location_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Tool Contact</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContact_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Tool Contact Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolContactRear_Final" ))%></td></tr>

		<tr class="evenrow"><td class="generic_label" width=192>Tool Gap</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGap_Final" ))%></td></tr>

		<tr><td class="generic_label" width=192>Tool Gap Rear</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "ToolGapRear_Final" ))%></td></tr>

  		
		<!-- Bottom Rule -->
		<tr><td colspan="4"><hr /></td></tr>
		<!-- End Bottom Rule -->

		       
    <tr class="evenrow"><td class="generic_label" width=192>Break In Time</td><td width=98 class="generic_input"><%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime" ))%></td>
	     <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime_In" ))%></td>
         <td width=110><%=EscapeDoubleQuotes(rsGetQCInspection( "BreakInTime_Final" ))%></td></tr>

		
	</table>
</td></tr>





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
