<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetProjectsList, cmdGetProjectsList
dim rsGetWorkOrderByWorkOrderId, cmdGetWorkOrderByWorkOrderId
dim cmdEditWorkOrder 
dim rsGetQuotesByWorkOrderId, cmdGetQuotesByWorkOrderId
dim rsGetQuotePartsByWorkOrderId, cmdGetQuotePartsByWorkOrderId
dim rsGetQuoteBearingsByWorkOrderId, cmdGetQuoteBearingsByWorkOrderId
dim rsGetQuoteSubWorkByWorkOrderId, cmdGetQuoteSubWorkByWorkOrderId
dim intWorkOrderId, intWorkOrderNumber, intRowNumber, intCustomerId, intProjectId, strCurrentPage
dim rsGetWorkOrderPrioritiesList, cmdGetWorkOrderPrioritiesList
dim rsGetLocationsList, cmdGetLocationsList
dim rsGetShippingMethodsList, cmdGetShippingMethodsList
dim rsGetEmployeesList, cmdGetEmployeesList
dim rsGetCustomerContactsByCustomerId, cmdGetCustomerContactsByCustomerId
dim rsGetQCTempLogsByWorkOrderId, cmdGetQCTempLogsByWorkOrderId

set cnnSQLConnection								= nothing
set rsGetProjectsList								= nothing
set cmdGetProjectsList							= nothing
set rsGetWorkOrderByWorkOrderId			= nothing
set cmdGetWorkOrderByWorkOrderId		= nothing
set cmdEditWorkOrder								= nothing
set rsGetQuotesByWorkOrderId				= nothing
set cmdGetQuotesByWorkOrderId				= nothing
set rsGetQuotePartsByWorkOrderId				= nothing
set cmdGetQuotePartsByWorkOrderId				= nothing
set rsGetQuoteBearingsByWorkOrderId				= nothing
set cmdGetQuoteBearingsByWorkOrderId				= nothing
set cmdGetQuoteSubWorkByWorkOrderId				= nothing
set rsGetQuoteSubWorkByWorkOrderId				= nothing
set rsGetWorkOrderPrioritiesList		= nothing
set cmdGetWorkOrderPrioritiesList		= nothing
set rsGetLocationsList							= nothing
set cmdGetLocationsList							= nothing
set rsGetShippingMethodsList				= nothing
set cmdGetShippingMethodsList				= nothing
set rsGetEmployeesList							= nothing
set cmdGetEmployeesList							= nothing
set rsGetCustomerContactsByCustomerId			= nothing
set cmdGetCustomerContactsByCustomerId			= nothing
set rsGetQCTempLogsByWorkOrderId			= nothing
set cmdGetQCTempLogsByWorkOrderId			= nothing

set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetProjectsList							= Server.CreateObject( "ADODB.Command" )
set cmdGetWorkOrderByWorkOrderId		= Server.CreateObject( "ADODB.Command" )
set cmdGetQuotesByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuotePartsByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteBearingsByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteSubWorkByWorkOrderId				= Server.CreateObject( "ADODB.Command" )
set cmdGetWorkOrderPrioritiesList		= Server.CreateObject( "ADODB.Command" )
set cmdGetLocationsList							= Server.CreateObject( "ADODB.Command" )
set cmdGetShippingMethodsList				= Server.CreateObject( "ADODB.Command" )
set cmdGetEmployeesList							= Server.CreateObject( "ADODB.Command" )
set cmdGetCustomerContactsByCustomerId								= Server.CreateObject( "ADODB.Command" )
set cmdGetQCTempLogsByWorkOrderId								= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


'**********************************************************************************
cmdGetProjectsList.ActiveConnection = cnnSQLConnection
cmdGetProjectsList.CommandText = "{call spGetProjectList}"

on error resume next
set rsGetProjectsList = cmdGetProjectsList.Execute
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
	if( Request.Form( "txtWorkOrderId" ) <> "" ) then
	set cmdEditWorkOrder	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditWorkOrder.ActiveConnection	= cnnSQLConnection
	cmdEditWorkOrder.CommandText			= "{? = call spEditWorkOrder( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderNumber", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@PromiseDate", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@DateIn", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@DateOut", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@SerialNumber", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@NewSpindle", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@PONumber", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Labor", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Material", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Subcontract", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Cost", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Charge", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Date", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@DateExp", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@SalesRep", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ShippingMethodId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@TrackingWaybill", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@LocationId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderPriorityId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BoeingSpindleId", adInteger, adParamInput, 4 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BoeingWorkOrderNumber", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Parts", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Bearings", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Lube", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BalVelocity", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BalVelocityFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@GSE", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@GSEFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BreakIn", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@BreakInFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RoomTemp", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RoomTempFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@FrontTemp", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@FrontTempFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RearTemp", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RearTempFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Cooling", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@CoolingFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RunoutFront", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RunoutFrontFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Rear", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@RearFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Other", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@IncomingInspection", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Comments", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Remarks", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@DateRec", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ExpectedDelDate", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Commission", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@CommissionReceivedDate", adVarChar, adParamInput, 10 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@AdditionalInfo", adLongVarChar, adParamInput, 1048576 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ActualDrawforce", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ActualDrawforceFinal", adVarChar, adParamInput, 100 )
	cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderContactId", adInteger, adParamInput, 4 )
	
	if( Request.Form( "txtWorkOrderId" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderId" )
	if( Request.Form( "txtWorkOrderNumber" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderNumber" ) = Request.Form( "txtWorkOrderNumber" )
	if( Request.Form( "hidProjectId" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ProjectId" ) = Request.Form( "hidProjectId" )
	if( Request.Form( "txtPromiseDate" ) <> "" ) then cmdEditWorkOrder.Parameters( "@PromiseDate" ) = Request.Form( "txtPromiseDate" )
	if( Request.Form( "txtDateIn" ) <> "" ) then cmdEditWorkOrder.Parameters( "@DateIn" ) = Request.Form( "txtDateIn" )
	if( Request.Form( "txtDateOut" ) <> "" ) then cmdEditWorkOrder.Parameters( "@DateOut" ) = Request.Form( "txtDateOut" )
	if( Request.Form( "txtSerialNumber" ) <> "" ) then cmdEditWorkOrder.Parameters( "@SerialNumber" ) = Request.Form( "txtSerialNumber" )
	if( Request.Form( "selNewSpindle" ) <> "" ) then cmdEditWorkOrder.Parameters( "@NewSpindle" ) = Request.Form( "selNewSpindle" )
	if( Request.Form( "txtPONumber" ) <> "" ) then cmdEditWorkOrder.Parameters( "@PONumber" ) = Request.Form( "txtPONumber" )
	if( Request.Form( "txtLabor" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Labor" ) = Request.Form( "txtLabor" )
	if( Request.Form( "txtMaterial" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Material" ) = Request.Form( "txtMaterial" )
	if( Request.Form( "txtSubcontract" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Subcontract" ) = Request.Form( "txtSubcontract" )
	if( Request.Form( "txtCost" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Cost" ) = Request.Form( "txtCost" )
	if( Request.Form( "txtCharge" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Charge" ) = Request.Form( "txtCharge" )
	if( Request.Form( "txtDate" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Date" ) = Request.Form( "txtDate" )
	if( Request.Form( "txtDateExp" ) <> "" ) then cmdEditWorkOrder.Parameters( "@DateExp" ) = Request.Form( "txtDateExp" )
	if( Request.Form( "txtSalesRep" ) <> "" ) then cmdEditWorkOrder.Parameters( "@SalesRep" ) = Request.Form( "txtSalesRep" )
	if( Request.Form( "selShippingMethod" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ShippingMethodId" ) = Request.Form( "selShippingMethod" )
	if( Request.Form( "txtTrackingWaybill" ) <> "" ) then cmdEditWorkOrder.Parameters( "@TrackingWaybill" ) = Request.Form( "txtTrackingWaybill" )
	if( Request.Form( "selLocation" ) <> "" ) then cmdEditWorkOrder.Parameters( "@LocationId" ) = Request.Form( "selLocation" )
	if( Request.Form( "selPriority" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderPriorityId" ) = Request.Form( "selPriority" )
	if( Request.Form( "txtBoeingSpindleId" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BoeingSpindleId" ) = Request.Form( "txtBoeingSpindleId" )
	if( Request.Form( "txtBoeingWorkOrderNumber" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BoeingWorkOrderNumber" ) = Request.Form( "txtBoeingWorkOrderNumber" )
	if( Request.Form( "txtParts" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Parts" ) = Request.Form( "txtParts" )
	if( Request.Form( "txtBearings" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Bearings" ) = Request.Form( "txtBearings" )
	if( Request.Form( "txtLube" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Lube" ) = Request.Form( "txtLube" )
	if( Request.Form( "txtBalVelocity" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BalVelocity" ) = Request.Form( "txtBalVelocity" )
	if( Request.Form( "txtBalVelocityFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BalVelocityFinal" ) = Request.Form( "txtBalVelocityFinal" )
	if( Request.Form( "txtGSE" ) <> "" ) then cmdEditWorkOrder.Parameters( "@GSE" ) = Request.Form( "txtGSE" )
	if( Request.Form( "txtGSEFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@GSEFinal" ) = Request.Form( "txtGSEFinal" )
	if( Request.Form( "txtBreakIn" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BreakIn" ) = Request.Form( "txtBreakIn" )
	if( Request.Form( "txtBreakInFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@BreakInFinal" ) = Request.Form( "txtBreakInFinal" )
	if( Request.Form( "txtRoomTemp" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RoomTemp" ) = Request.Form( "txtRoomTemp" )
	if( Request.Form( "txtRoomTempFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RoomTempFinal" ) = Request.Form( "txtRoomTempFinal" )
	if( Request.Form( "txtFrontTemp" ) <> "" ) then cmdEditWorkOrder.Parameters( "@FrontTemp" ) = Request.Form( "txtFrontTemp" )
	if( Request.Form( "txtFrontTempFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@FrontTempFinal" ) = Request.Form( "txtFrontTempFinal" )
	if( Request.Form( "txtRearTemp" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RearTemp" ) = Request.Form( "txtRearTemp" )
	if( Request.Form( "txtRearTempFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RearTempFinal" ) = Request.Form( "txtRearTempFinal" )
	if( Request.Form( "txtCooling" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Cooling" ) = Request.Form( "txtCooling" )
	if( Request.Form( "txtCoolingFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@CoolingFinal" ) = Request.Form( "txtCoolingFinal" )
	if( Request.Form( "txtRunoutFront" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RunoutFront" ) = Request.Form( "txtRunoutFront" )
	if( Request.Form( "txtRunoutFrontFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RunoutFrontFinal" ) = Request.Form( "txtRunoutFrontFinal" )
	if( Request.Form( "txtRear" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Rear" ) = Request.Form( "txtRear" )
	if( Request.Form( "txtRearFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@RearFinal" ) = Request.Form( "txtRearFinal" )
	if( Request.Form( "txtOther" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Other" ) = Request.Form( "txtOther" )
	if( Request.Form( "txtIncomingInspection" ) <> "" ) then cmdEditWorkOrder.Parameters( "@IncomingInspection" ) = Request.Form( "txtIncomingInspection" )
	if( Request.Form( "txtComments" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Comments" ) = Request.Form( "txtComments" )
	if( Request.Form( "txtRemarks" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Remarks" ) = Request.Form( "txtRemarks" )
	if( Request.Form( "txtDateRec" ) <> "" ) then cmdEditWorkOrder.Parameters( "@DateRec" ) = Request.Form( "txtDateRec" )
	if( Request.Form( "txtExpectedDelDate" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ExpectedDelDate" ) = Request.Form( "txtExpectedDelDate" )
	if( Request.Form( "txtCommission" ) <> "" ) then cmdEditWorkOrder.Parameters( "@Commission" ) = Request.Form( "txtCommission" )
	if( Request.Form( "txtCommissionRecDate" ) <> "" ) then cmdEditWorkOrder.Parameters( "@CommissionReceivedDate" ) = Request.Form( "txtCommissionRecDate" )
	if( Request.Form( "txtAdditionalInfo" ) <> "" ) then cmdEditWorkOrder.Parameters( "@AdditionalInfo" ) = Request.Form( "txtAdditionalInfo" )
	if( Request.Form( "txtDrawforce" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ActualDrawforce" ) = Request.Form( "txtDrawforce" )
	if( Request.Form( "txtDrawforceFinal" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ActualDrawforceFinal" ) = Request.Form( "txtDrawforceFinal" )
	if( Request.Form( "selWorkOrderContact" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderContactId" ) = Request.Form( "selWorkOrderContact" )
		
	on error resume next
	cmdEditWorkOrder.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0
	
	intWorkOrderId = cmdEditWorkOrder.Parameters( "@Return" ).Value
	
	elseif( Request.Form( "txtWorkOrderIdDetail" ) <> "" ) then
	
		if ( Request.Form( "txtPurchaseOrderId" ) = "" ) then
	
			set cmdEditWorkOrder	= Server.CreateObject( "ADODB.Command" )
		
			cmdEditWorkOrder.ActiveConnection	= cnnSQLConnection
			cmdEditWorkOrder.CommandText				= "{? = call spEditCall( ?,?,?,?,?,?,? ) }"
			
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )	
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@CallId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@EmployeeId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@CallComments", adLongVarChar, adParamInput, 1048576 )
			
			if( Request.Form( "txtCustomerIdDetail" ) <> "" ) then cmdEditWorkOrder.Parameters( "@CustomerId" ) = Request.Form( "txtCustomerIdDetail" )
			if( Request.Form( "selEmployeeId" ) <> "" ) then cmdEditWorkOrder.Parameters( "@EmployeeId" ) = Request.Form( "selEmployeeId" )
			if( Request.Form( "txtProjectIdDetail" ) <> "" ) then cmdEditWorkOrder.Parameters( "@ProjectId" ) = Request.Form( "txtProjectIdDetail" )
			if( Request.Form( "txtWorkOrderIdDetail" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderIdDetail" )
			if( Request.Form( "txtCallComments" ) <> "" ) then cmdEditWorkOrder.Parameters( "@CallComments" ) = Request.Form( "txtCallComments" )
			
			on error resume next
			cmdEditWorkOrder.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if
		
		else	
			set cmdEditWorkOrder	= Server.CreateObject( "ADODB.Command" )
		
			cmdEditWorkOrder.ActiveConnection	= cnnSQLConnection
			cmdEditWorkOrder.CommandText				= "{? = call spEditWorkOrderPO( ?,? ) }"
			
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
			cmdEditWorkOrder.Parameters.Append cmdEditWorkOrder.CreateParameter( "@PurchaseOrderId", adInteger, adParamInput, 4 )
			
			if( Request.Form( "txtWorkOrderIdDetail" ) <> "" ) then cmdEditWorkOrder.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderIdDetail" )
			if( Request.Form( "txtPurchaseOrderId" ) <> "" ) then cmdEditWorkOrder.Parameters( "@PurchaseOrderId" ) = Request.Form( "txtPurchaseOrderId" )
			
			on error resume next
			cmdEditWorkOrder.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if		
		end if
									
		on error goto 0
	
		intProjectId = Request.Form( "txtProjectIdDetail" )
		intWorkOrderId = Request.Form( "txtWorkOrderIdDetail" )
		
	end if

else
	if( Request.QueryString( "intWorkOrderId" ) = "" ) then
		intWorkOrderId = 0
	else
		intWorkOrderId = Request.QueryString( "intWorkOrderId" )
	end if
end if

cmdGetWorkOrderByWorkOrderId.ActiveConnection = cnnSQLConnection
cmdGetWorkOrderByWorkOrderId.CommandText = "{call spGetWorkOrderByWorkOrderId( ?,?,?,? ) }"
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@IsPrint", adTinyInt, adParamInput, 1 )
cmdGetWorkOrderByWorkOrderId.Parameters.Append cmdGetWorkOrderByWorkOrderId.CreateParameter( "@CustomerId", adInteger, adParamOutput, 4 )
cmdGetWorkOrderByWorkOrderId.Parameters( "@WorkOrderId" )	= intWorkOrderId

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
cmdGetQuotesByWorkOrderId.ActiveConnection	= cnnSQLConnection
cmdGetQuotesByWorkOrderId.CommandText		= "{call spGetQuotesByWorkOrderId( ?,? ) }"

cmdGetQuotesByWorkOrderId.Parameters.Append cmdGetQuotesByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuotesByWorkOrderId.Parameters.Append cmdGetQuotesByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdGetQuotesByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= intWorkOrderId

on error resume next
set rsGetQuotesByWorkOrderId = cmdGetQuotesByWorkOrderId.Execute
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

cmdGetQuotePartsByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= intWorkOrderId

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

cmdGetQuoteBearingsByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= intWorkOrderId

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

cmdGetQuoteSubWorkByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= intWorkOrderId

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
'**********************************************************************************
cmdGetQCTempLogsByWorkOrderId.ActiveConnection	= cnnSQLConnection
cmdGetQCTempLogsByWorkOrderId.CommandText		= "{call spGetQCTempLogsByWorkOrderId( ?,? ) }"

cmdGetQCTempLogsByWorkOrderId.Parameters.Append cmdGetQCTempLogsByWorkOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQCTempLogsByWorkOrderId.Parameters.Append cmdGetQCTempLogsByWorkOrderId.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )

cmdGetQCTempLogsByWorkOrderId.Parameters( "@WorkOrderId" ).Value			= intWorkOrderId

on error resume next
set rsGetQCTempLogsByWorkOrderId = cmdGetQCTempLogsByWorkOrderId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetWorkOrderPrioritiesList.ActiveConnection = cnnSQLConnection
cmdGetWorkOrderPrioritiesList.CommandText = "{call spGetWorkOrderPriorityList}"

on error resume next
set rsGetWorkOrderPrioritiesList = cmdGetWorkOrderPrioritiesList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetLocationsList.ActiveConnection = cnnSQLConnection
cmdGetLocationsList.CommandText = "{call spGetLocationList}"

on error resume next
set rsGetLocationsList = cmdGetLocationsList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetShippingMethodsList.ActiveConnection = cnnSQLConnection
cmdGetShippingMethodsList.CommandText = "{call spGetShippingMethodList}"

on error resume next
set rsGetShippingMethodsList = cmdGetShippingMethodsList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0
'**********************************************************************************
cmdGetEmployeesList.ActiveConnection = cnnSQLConnection
cmdGetEmployeesList.CommandText = "{call spGetEmployeeList}"

on error resume next
set rsGetEmployeesList = cmdGetEmployeesList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetCustomerContactsByCustomerId.ActiveConnection = cnnSQLConnection
cmdGetCustomerContactsByCustomerId.CommandText = "{call spGetCustomerContactsByCustomerId ( ?,? ) }"
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
cmdGetCustomerContactsByCustomerId.Parameters( "@CustomerId" )	= cmdGetWorkOrderByWorkOrderId.Parameters( "@CustomerId" )

on error resume next
set rsGetCustomerContactsByCustomerId = cmdGetCustomerContactsByCustomerId.Execute
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
'**********************************************************************************
%>
<HTML>
<HEAD>
<TITLE>Centerline - View / Edit Work Order</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function PrintFinalInspectionReport( strWorkOrderId )
	{
		if( document.frmWorkOrder.txtDateOut.value == '' )
			alert( 'You must assign a Date Out to this work order, before the Final Inspection Report can be printed' );
		else
			OpenPrintBox( 'rpt_finalinspection.asp?intWorkOrderId='+strWorkOrderId );
	}
	function PrintSpindleInformation( strSpindleId )
	{
			OpenPrintBox( 'print_spindle_details.asp?intSpindleId='+strSpindleId );
			//if (window.print) window.print();
	}
	function PrintWorkOrderInformation( strWorkOrderId )
	{
			OpenPrintBox( 'print_spindleinformation.asp?intWorkOrderId='+strWorkOrderId );
			OpenPrintBox('print_workorder.asp?intWorkOrderId='+strWorkOrderId);
			if (window.print) window.print();
	}
	function PrintReport()
	{
		var nReportCount;

		for( nReportCount = 0; nReportCount < document.frmWorkOrder.rdoPrint.length; nReportCount++ )
		{
			if( document.frmWorkOrder.rdoPrint[nReportCount].checked == true )
			{
				eval( document.frmWorkOrder.rdoPrint[nReportCount].value  );
			}
		}
	}
	
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
  function OpenPurchaseOrder( strPurchaseOrderId )
	{
	  OpenWindow( 'po_details.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO'+strPurchaseOrderId,'width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenSpindle( strSpindleId )
	{
	  OpenWindow( 'spindle_details.asp?intSpindleId='+strSpindleId,'CenterlineSpindle'+strSpindleId,'width=660,height=540,left='+GetCenterX( 660 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenCustomer( strCustomerId )
	{
	  OpenWindow( 'customer_details.asp?intCustomerId='+strCustomerId,'CenterlineCustomer'+strCustomerId,'width=657,height=550,left='+GetCenterX( 657 )+',top='+GetCenterY( 550 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 550 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenProject( strProjectId )
	{
	  OpenWindow( 'project_details.asp?intProjectId='+strProjectId,'CenterlineProject'+strProjectId,'width=657,height=540,left='+GetCenterX( 657 )+',top='+GetCenterY( 540 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 540 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenProjectSearch(  )
	{
	  OpenWindow( 'project_search.asp','ProjectSearch','width=691,height=500,left='+GetCenterX( 691 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 691 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenPOSearch(  )
	{
	  OpenWindow( 'po_search.asp','POSearch','width=691,height=500,left='+GetCenterX( 691 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 691 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectProject( intProjectId, strProjectName )
	{
		document.forms.frmWorkOrder.hidProjectId.value = intProjectId;
		document.forms.frmWorkOrder.txtProjectName.value = strProjectName;
	}
	function OpenQuote( strQuoteId )
	{
	  OpenWindow( 'quote_details.asp?intQuoteId='+strQuoteId,'CenterlineQuote'+strQuoteId,'width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
  function OpenProduct( strProductId ) 
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'Centerline','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
  function OpenQCInspection( intWorkOrderId )
	{
	  OpenWindow( 'qcinspection_details.asp?intWorkOrderId='+intWorkOrderId,'Centerline','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function ConfirmDelete( intCallId, strCallComments )
	{
		if( confirm( 'Are you sure you want to remove item ' + strCallComments + '?' ) )
		{
			OpenWindow( 'call_delete.asp?intCallId='+intCallId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function ConfirmDeleteWorkOrderPO( intWorkOrderId, intPurchaseOrderId, intPONum )
	{
		if( confirm( 'Are you sure you want to remove item ' + intPONum + '?' ) )
		{
			OpenWindow( 'workorderpo_delete.asp?intWorkOrderId='+intWorkOrderId+'&intPurchaseOrderId='+intPurchaseOrderId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function ConfirmDeleteQCTempLog( intQCTempLogId, strQCDate )
	{
		if( confirm( 'Are you sure you want to remove item ' + strQCDate + '?' ) )
		{
			OpenWindow( 'qctemplog_delete.asp?intQCTempLogId='+intQCTempLogId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	function SelectPO( intPOId )
	{
		document.forms.frmWorkOrderPOs.txtPurchaseOrderId.value = intPOId;
	}
	</script>
</HEAD>
<BODY topmargin=0 leftmargin=0 onload="javascript: if (window.print) window.print();">
<%if not( rsGetWorkOrderByWorkOrderId is nothing ) then
		if( rsGetWorkOrderByWorkOrderId.State = adStateOpen ) then
			if( not rsGetWorkOrderByWorkOrderId.EOF ) then
				intCustomerId = rsGetWorkOrderByWorkOrderId( "CustomerId" )
				intProjectId = rsGetWorkOrderByWorkOrderId( "ProjectId" )
				if( intWorkOrderId <> 0 ) then
					intWorkOrderNumber = rsGetWorkOrderByWorkOrderId( "WorkOrderNumber" )
				end if%>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intWorkOrderId = 0 ) then%>
   <tr><td class="sectiontitle" width=25%>New Work Order</td>
   <%else%>
   <tr><td class="sectiontitle" width=25%>Work Order <%=rsGetWorkOrderByWorkOrderId( "WorkOrderNumber" )%></td>
   <%end if%>
     <td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=left><%=rsGetWorkOrderByWorkOrderId( "Customer" )%>&nbsp;&nbsp;:&nbsp;&nbsp;Project <%=rsGetWorkOrderByWorkOrderId( "ProjectName" )%></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Work Order Form -->
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
  		<tr><td class="label_1">Project</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "ProjectName" )%></td>
  	    <td class="label_2">Work Order #</td><td class="generic_input"><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "WorkOrderNumber" ))%></td></tr>
			<tr><td class="label_1">Spindle Type</td><td colspan="3" class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "SpindleType" )%></td></tr>
			<tr><td class="label_1">Contact</td><td class="generic_input">
   	    <%if not( rsGetCustomerContactsByCustomerId is nothing ) then
   					if( rsGetCustomerContactsByCustomerId.State = adStateOpen ) then
   						while( not rsGetCustomerContactsByCustomerId.EOF )
   							if( rsGetCustomerContactsByCustomerId( "CustomerContactId" ) = rsGetWorkOrderByWorkOrderId( "WorkOrderContactId" ) ) then%>
   								<%=rsGetCustomerContactsByCustomerId( "Contact" )%>
   							<%end if
   							rsGetCustomerContactsByCustomerId.MoveNext
   						wend
   					end if
   				end if%>
   			</td>
   	  <td class="label_2">Sales Rep</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "SalesRepName" )%></td></tr>
			<tr><td class="label_1">Log In Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "DateIn" )%></td>
				<td class="label_2">Log Out Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "DateOut" )%></td></tr>
			<tr><td class="label_1">Received Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "DateRec" )%></td>
       <td class="label_2">Promised RTS Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "PromiseDate" )%></td></tr>
     <tr><td class="label_1">EPA Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "Date" )%></td>
      <td class="label_2">Latest RTS Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "ExpectedDelDate" )%></td></tr>
     <tr><td class="label_1">Shipping</td><td class="generic_input">
   	    <%if not( rsGetShippingMethodsList is nothing ) then
   					if( rsGetShippingMethodsList.State = adStateOpen ) then
   						while( not rsGetShippingMethodsList.EOF )%>
   							<%if( rsGetShippingMethodsList( "ShippingMethodId" ) = rsGetWorkOrderByWorkOrderId( "ShippingMethodId" ) ) then%>
   								<%=rsGetShippingMethodsList( "ShippingMethod" )%>
   							<%end if
   							rsGetShippingMethodsList.MoveNext
   						wend
   					end if
   				end if
   			%>
				</td>
       <td class="label_2">Ship Date</td><td class="generic_input"><%=rsGetWorkOrderByWorkOrderId( "DateExp" )%></td></tr>
     <tr><td class="label_1">Tracking/Waybill Number</td><td class="generic_input" colspan=3><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "TrackingWaybill" ))%></td></tr>
     <tr><td colspan=2>&nbsp;</td></tr>
     <tr><td class="label_1" colspan=2>(EPA - Earliest Possible Assembly)</td>
				<td class="label_2" colspan=2>(RTS - Ready To Ship)</td></tr>
 </table>
</td></tr>
	  
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
 
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   	<tr><td class="label_1">Serial #</td><td class="generic_input"><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "SerialNumber" ))%></td>
       <td class="label_2">New Spindle</td><td class="generic_input">
				<%if( UCase( rsGetWorkOrderByWorkOrderId( "NewSpindle" ) ) = "YES" ) then%>YES<%end if%>
   	    <%if( UCase( rsGetWorkOrderByWorkOrderId( "NewSpindle" ) ) = "NO" ) then%>NO<%end if%>
   	  </select></td></tr>
	    
     <tr><td class="label_1">Boeing Spindle Id</td><td class="generic_input"><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "BoeingSpindleId" ))%></td>
       <td class="label_2">Boeing Work Order #</td><td class="generic_input"><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "BoeingWorkOrderNumber" ))%></td></tr>
     <tr><td class="label_1">PO #</td><td class="generic_input"><%=EscapeDoubleQuotes(rsGetWorkOrderByWorkOrderId( "PONumber" ))%></td>
       <td class="label_2">Priority</td><td class="generic_input">
   	    <%if not( rsGetWorkOrderPrioritiesList is nothing ) then
   					if( rsGetWorkOrderPrioritiesList.State = adStateOpen ) then
   						while( not rsGetWorkOrderPrioritiesList.EOF )%><%if( rsGetWorkOrderPrioritiesList( "WorkOrderPriorityId" ) = rsGetWorkOrderByWorkOrderId( "WorkOrderPriorityId" ) ) then%><%=rsGetWorkOrderPrioritiesList( "WorkOrderPriority" )%><%end if%>
   							<%rsGetWorkOrderPrioritiesList.MoveNext
   						wend
   					end if
   				end if%>
   	  </select></td>
 </table>
</td></tr>

<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
 
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		<tr><td class="label_1" colspan=4>Repair Specific Info</td>
		<tr><td class="generic_input" colspan=4><%=rsGetWorkOrderByWorkOrderId( "AdditionalInfo" )%></td></tr>
		<tr><td class="label_1" colspan=4>Incoming Inspection</td>
		<tr><td class="generic_input" colspan=4><%=rsGetWorkOrderByWorkOrderId( "IncomingInspection" )%></td></tr>
		<tr><td class="label_1" colspan=4>Centerline Comments</td>
		<tr><td class="generic_input" colspan=4><%=rsGetWorkOrderByWorkOrderId( "Comments" )%></td></tr>
		<tr><td class="label_1" colspan=4>Final Remarks</td>
		<tr><td class="generic_input" colspan=4><%=rsGetWorkOrderByWorkOrderId( "Remarks" )%></td></tr>
		<tr><td colspan=4>&nbsp;</td></tr>
	</table>
</td></tr>
	    <%	end if
	end if
end if%></BODY>
</HTML>
<%
if not( rsGetWorkOrderByWorkOrderId is nothing ) then
	if( rsGetWorkOrderByWorkOrderId.State = adStateOpen ) then
		rsGetWorkOrderByWorkOrderId.Close
	end if
end if

if not( rsGetProjectsList is nothing ) then
	if( rsGetProjectsList.State = adStateOpen ) then
		rsGetProjectsList.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuotesByWorkOrderId				= nothing
set cmdGetQuotesByWorkOrderId				= nothing
set rsGetWorkOrderPrioritiesList		= nothing
set cmdGetWorkOrderPrioritiesList		= nothing
set rsGetLocationsList							= nothing
set cmdGetLocationsList							= nothing
set rsGetShippingMethodsList				= nothing
set cmdGetShippingMethodsList				= nothing
set rsGetProjectsList								= nothing
set cmdGetProjectsList							= nothing
set rsGetWorkOrderByWorkOrderId			= nothing
set cmdGetWorkOrderByWorkOrderId		= nothing
set rsGetQCTempLogsByWorkOrderId			= nothing
set cmdGetQCTempLogsByWorkOrderId		= nothing
set cmdEditWorkOrder								= nothing
set cnnSQLConnection								= nothing
%>