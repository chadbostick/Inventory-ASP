<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetShippingMethodsList, cmdGetShippingMethodsList, cmdEditPurchaseOrder
dim rsGetEmployeeList, cmdGetEmployeeList
dim rsGetSupplierList, cmdGetSupplierList
dim rsGetProductListByName, cmdGetProductListByName
dim rsGetPurchaseOrderByPurchaseOrderId, cmdGetPurchaseOrderByPurchaseOrderId
dim intRowNumber, intPurchaseOrderId, strCurrentPage

set cnnSQLConnection											= nothing
set rsGetShippingMethodsList							= nothing
set cmdGetShippingMethodsList							= nothing
set rsGetEmployeeList											= nothing
set cmdGetEmployeeList										= nothing
set rsGetSupplierList											= nothing
set cmdGetSupplierList										= nothing
set rsGetProductListByName											= nothing
set cmdGetProductListByName											= nothing
set rsGetPurchaseOrderByPurchaseOrderId		= nothing
set cmdGetPurchaseOrderByPurchaseOrderId	= nothing
set cmdEditPurchaseOrder									= nothing

set cnnSQLConnection											= Server.CreateObject( "ADODB.Connection" )
set cmdGetShippingMethodsList							= Server.CreateObject( "ADODB.Command" )
set cmdGetEmployeeList										= Server.CreateObject( "ADODB.Command" )
set cmdGetSupplierList										= Server.CreateObject( "ADODB.Command" )
set cmdGetProductListByName											= Server.CreateObject( "ADODB.Command" )
set cmdGetPurchaseOrderByPurchaseOrderId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

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
cmdGetEmployeeList.ActiveConnection = cnnSQLConnection
cmdGetEmployeeList.CommandText = "{call spGetEmployeeList}"

on error resume next
set rsGetEmployeeList = cmdGetEmployeeList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0
'**********************************************************************************
cmdGetSupplierList.ActiveConnection = cnnSQLConnection
cmdGetSupplierList.CommandText = "{call spGetSupplierList}"

on error resume next
set rsGetSupplierList = cmdGetSupplierList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
			
on error goto 0
'**********************************************************************************
cmdGetProductListByName.ActiveConnection = cnnSQLConnection
cmdGetProductListByName.CommandText = "{call spGetProductListByName}"

on error resume next
set rsGetProductListByName = cmdGetProductListByName.Execute
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
	if( Request.Form( "txtPOIDMaster" ) <> "" ) then
		set cmdEditPurchaseOrder	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditPurchaseOrder.ActiveConnection	= cnnSQLConnection
		'cmdEditPurchaseOrder.CommandText			= "{? = call spEditPurchaseOrder( ?,?,?,?,?,?,?,?,?,?,?,?,?,? ) }"
		cmdEditPurchaseOrder.CommandText = "spEditPurchaseOrder"
		cmdEditPurchaseOrder.CommandType = adCmdStoredProc
	
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderID", adInteger, adParamInput, 4 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderNumber", adVarChar, adParamInput, 100 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderDescription", adLongVarChar, adParamInput, 1048576 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@SupplierID", adInteger, adParamInput, 4 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@EmployeeID", adInteger, adParamInput, 4 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@OrderDate", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@DateRequired", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@DatePromised", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@DateClosed", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ShipDate", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@DateShippedToSupplier", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ShippingMethodID", adInteger, adParamInput, 4 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TrackingNumber", adVarChar, adParamInput, 100 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@FreightCharge", adVarChar, adParamInput, 10 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@InternalNotes", adLongVarChar, adParamInput, 1048576 )
		cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PODescription", adLongVarChar, adParamInput, 1048576 )
			
		if( Request.Form( "txtPOIDMaster" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PurchaseOrderID" ) = Request.Form( "txtPOIDMaster" )
		if( Request.Form( "txtPONumber" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PurchaseOrderNumber" ) = Request.Form( "txtPONumber" )
		if( Request.Form( "txtDescription" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PurchaseOrderDescription" ) = Request.Form( "txtDescription" )
		if( Request.Form( "selSupplier" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@SupplierID" ) = Request.Form( "selSupplier" )
		if( Request.Form( "selBuyer" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@EmployeeID" ) = Request.Form( "selBuyer" )
		if( Request.Form( "txtOrderDate" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@OrderDate" ) = Request.Form( "txtOrderDate" )
		if( Request.Form( "txtDateRequired" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@DateRequired" ) = Request.Form( "txtDateRequired" )
		if( Request.Form( "txtDatePromised" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@DatePromised" ) = Request.Form( "txtDatePromised" )
		if( Request.Form( "txtDateClosed" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@DateClosed" ) = Request.Form( "txtDateClosed" )
		if( Request.Form( "txtShipDate" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@ShipDate" ) = Request.Form( "txtShipDate" )
		if( Request.Form( "txtDateShippedToSupplier" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@DateShippedToSupplier" ) = Request.Form( "txtDateShippedToSupplier" )
		if( Request.Form( "selShippingMethod" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@ShippingMethodID" ) = Request.Form( "selShippingMethod" )
		if( Request.Form( "txtFreightCharge" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@FreightCharge" ) = Request.Form( "txtFreightCharge" )
		if( Request.Form( "txtTrackingNumber" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@TrackingNumber" ) = Request.Form( "txtTrackingNumber" )
		if( Request.Form( "txtInternalNotes" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@InternalNotes" ) = Request.Form( "txtInternalNotes" )
		if( Request.Form( "txtPODescription" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PODescription" ) = Request.Form( "txtPODescription" )

		on error resume next
		cmdEditPurchaseOrder.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
					
		on error goto 0
	
		intPurchaseOrderId = cmdEditPurchaseOrder.Parameters( "@Return" ).Value
		
	elseif( Request.Form( "txtPOIDDetail" ) <> "" ) then
		
		if (Request.Form( "txtWorkOrderId" ) = "" ) then	
			set cmdEditPurchaseOrder	= Server.CreateObject( "ADODB.Command" )
		
			cmdEditPurchaseOrder.ActiveConnection	= cnnSQLConnection
			cmdEditPurchaseOrder.CommandText			= "{? = call spEditPurchaseOrderDetail( ?,?,?,?,?,?,?,? ) }"
			
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionDate", adVarChar, adParamInput, 10 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ProductID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionDescription", adLongVarChar, adParamInput, 1048576 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@UnitPrice", adCurrency, adParamInput, 10 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@UnitsOrdered", adVarChar, adParamInput, 10 )
			
			if( Request.Form( "txtTransactionDate" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@TransactionDate" ) = Request.Form( "txtTransactionDate" )
			if( Request.Form( "txtProductId" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@ProductID" ) = Request.Form( "txtProductId" )
			if( Request.Form( "txtPOIDDetail" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PurchaseOrderID" ) = Request.Form( "txtPOIDDetail" )
			if( Request.Form( "txtTransactionDescription" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@TransactionDescription" ) = Request.Form( "txtTransactionDescription" )
			if( Request.Form( "txtUnitPrice" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@UnitPrice" ) = Request.Form( "txtUnitPrice" )
			if( Request.Form( "txtOrderQty" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@UnitsOrdered" ) = Request.Form( "txtOrderQty" )
					
			on error resume next
			cmdEditPurchaseOrder.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if
		
		else	
			set cmdEditPurchaseOrder	= Server.CreateObject( "ADODB.Command" )
		
			cmdEditPurchaseOrder.ActiveConnection	= cnnSQLConnection
			cmdEditPurchaseOrder.CommandText				= "{? = call spEditWorkOrderPO( ?,? ) }"
			
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderId", adInteger, adParamInput, 4 )
			
			if( Request.Form( "txtWorkOrderId" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@WorkOrderId" ) = Request.Form( "txtWorkOrderId" )
			if( Request.Form( "txtPOIdDetail" ) <> "" ) then cmdEditPurchaseOrder.Parameters( "@PurchaseOrderId" ) = Request.Form( "txtPOIdDetail" )
			
			on error resume next
			cmdEditPurchaseOrder.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if	
		end if
					
		on error goto 0
		
		intPurchaseOrderId = Request.Form( "txtPOIdDetail" )
		
	end if
		
else
	intPurchaseOrderId = Request.QueryString( "intPurchaseOrderId" )
end if

cmdGetPurchaseOrderByPurchaseOrderId.ActiveConnection = cnnSQLConnection
cmdGetPurchaseOrderByPurchaseOrderId.CommandText = "{call spGetPurchaseOrderByPurchaseOrderId( ?,? ) }"
cmdGetPurchaseOrderByPurchaseOrderId.Parameters.Append cmdGetPurchaseOrderByPurchaseOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetPurchaseOrderByPurchaseOrderId.Parameters.Append cmdGetPurchaseOrderByPurchaseOrderId.CreateParameter( "@PurchaseOrderId", adInteger, adParamInput, 4 )
cmdGetPurchaseOrderByPurchaseOrderId.Parameters( "@PurchaseOrderId" )	= intPurchaseOrderId

on error resume next
set rsGetPurchaseOrderByPurchaseOrderId = cmdGetPurchaseOrderByPurchaseOrderId.Execute
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
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - View / Edit PO</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	function ConfirmDelete( intTransactionId, strProduct )
	{
		if( confirm( 'Are you sure you want to remove item ' + strProduct + '?' ) )
		{
			OpenWindow( 'potransactions_delete.asp?intTransactionId='+intTransactionId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}	
	function ProductSelected()
	{
		document.frmPODetail.txtProductId.value = document.frmPODetail.selProductName.options[document.frmPODetail.selProductName.selectedIndex].value;
	}	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}	
	function PrintCenterlinePO( strPurchaseOrderId )
	{
		if( strPurchaseOrderId == '0' )
			alert( 'You must save your purchase order before continuing.' );
		else
		{
			if( document.frmPrintPO.radLogo[0].checked == true )
				OpenWindow( 'print_pocenterline.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO','width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
			else
				OpenWindow( 'print_pocss.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO','width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
			
		}
	}	
	function PrintCenterlineSubWork( strPurchaseOrderId )
	{
		if( strPurchaseOrderId == '0' )
			alert( 'You must save your purchase order before continuing.' );
		else
		{
			if( document.frmPrintPO.radLogo[0].checked == true )
				OpenWindow( 'print_subworkcenterline.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO','width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
			else
				OpenWindow( 'print_subworkcss.asp?intPurchaseOrderId='+strPurchaseOrderId,'CenterlinePO','width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
			
		}
	}	
	function OpenEmployee( strEmployeeId )
	{
	  OpenWindow( 'employee_details.asp?intEmployeeId='+strEmployeeId,'Centerline','width=640,height=160,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 160 )+',scrollbars=no,dependent=yes' );
	}
	function SelectEmployee( objFormElement )
	{
		if ( objFormElement.value == -1 )
		{
			OpenEmployee( 0 );
			objEmployeeField = objFormElement[1];
		}
	}	
	function OpenSpindle( strSpindleId )
	{
	  OpenWindow( 'spindle_details.asp?intSpindleId='+strSpindleId,'CenterlineSpindle'+strSpindleId,'width=660,height=540,left='+GetCenterX( 660 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
	}
	function EmployeeSaved( intId, strName )
	{
		objEmployeeField.text = strName;
		objEmployeeField.value = intId;
	}		
  function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'CenterlineProduct'+strProductId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenProductSearch( bMultiple )
	{
		var strSearchURL = "product_search.asp?isMultiple=false";
		
		if ( bMultiple ) strSearchURL = "product_search.asp?isMultiple=true";
			
	  OpenWindow( strSearchURL,'ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectProduct( intProductId, strProductName, strPartNumber, strUnitPrice, isMultiple )
	{
		document.forms.frmPODetail.hidProductId.value = intProductId;
		document.forms.frmPODetail.txtProductId.value = intProductId;
		document.forms.frmPODetail.txtProductName.value = strProductName;
		document.forms.frmPODetail.txtOrderQty.value = 1;
		document.forms.frmPODetail.txtUnitPrice.value = strUnitPrice;
		if (isMultiple)
		{
		  document.forms.frmPODetail.submit();
		}
	}
  function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}	
	function OpenWorkOrderSearch(  )
	{
	  OpenWindow( 'workorder_search.asp','ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectWorkOrder( intWorkOrderId, intWorkOrderNum )
	{
		document.forms.frmWorkOrderPOs.txtWorkOrderId.value = intWorkOrderId;
	}
	function ConfirmDeleteWorkOrderPO( intWorkOrderId, intPurchaseOrderId, intPONum )
	{
		if( confirm( 'Are you sure you want to remove item ' + intPONum + '?' ) )
		{
			OpenWindow( 'workorderpo_delete.asp?intWorkOrderId='+intWorkOrderId+'&intPurchaseOrderId='+intPurchaseOrderId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	function OpenUsedPartsPopup( intPurchaseOrderId )
	{
	  OpenWindow( 'po_parts_prevused.asp?poID=' + intPurchaseOrderId, 'SelectPrevUsedParts','width=660,height=200,left='+GetCenterX( 660 )+',top='+GetCenterY( 200 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 200 )+',scrollbars=yes,dependent=yes' );
	}
  function OpenTransaction( strTransactionId )
	{
	  OpenWindow( 'transaction_details.asp?isRefreshParent=true&intTransactionId='+strTransactionId,'Transaction'+strTransactionId,'width=640,height=300,left='+GetCenterX( 640 )+',top='+GetCenterY( 300 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 300 )+',scrollbars=no,dependent=yes' );
	}
	// End hiding script from old browsers -->
  </script>
</head>
<body topmargin=0 leftmargin=0>
<script type="text/javascript" src="/js/tooltips.js"></script>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetPurchaseOrderByPurchaseOrderId is nothing ) then
		if( rsGetPurchaseOrderByPurchaseOrderId.State = adStateOpen ) then
			if( not rsGetPurchaseOrderByPurchaseOrderId.EOF ) then%>
 <!--Top Title Bar -->
 <tr><td class="pagetitle">  
  <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
    <%if( intPurchaseOrderId = 0 ) then%>
		<tr><td class="sectiontitle" width=20%>New PO</td>
		<%else%>
    <tr><td class="sectiontitle" width=20%>PO&nbsp;&nbsp;<%=intPurchaseOrderId%></td>
		<%end if%>
      <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
      <td class="sectioncaption" align=right><%=FormatDateTime( Now(), vbShortDate )%></td></tr>
  </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- PO Form -->
<tr><td align=center>
	<form id="frmPOMaster" name="frmPOMaster" action="" method="post">
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		<tr><td class="label_1">PO ID</td><td class="input_1"><input type="text" class="forminput" name="txtPOIDMaster" readonly value="<%=intPurchaseOrderId%>"></td>
				<td class="label_2">PO Number</td><td class="input_2"><input type="text" class="forminput" name="txtPONumber" maxlength=100 onchange="this.isDirty=true;" isRequired="true" fieldName="PO Number" value="<%=EscapeDoubleQuotes(rsGetPurchaseOrderByPurchaseOrderId( "PurchaseOrderNumber" ))%>"></td></tr>
		<%if( IsNull( rsGetPurchaseOrderByPurchaseOrderId( "OrderDate" ) ) or ( rsGetPurchaseOrderByPurchaseOrderId( "OrderDate" ) = "" ) ) then%>
		<tr><td class="label_1">Date Ordered</td><td class="input_1"><input type="text" class="forminput" name="txtOrderDate" maxlength=32 onchange="this.isDirty=true;" isDate="true" fieldName="Date Ordered" value="<%=FormatDateTime( Now(), vbShortDate )%>"></td>
		<%else%>
		<tr><td class="label_1">Date Ordered</td><td class="input_1"><input type="text" class="forminput" name="txtOrderDate" maxlength=32 onchange="this.isDirty=true;" isDate="true" fieldName="Date Ordered" value="<%=FormatDateTime( rsGetPurchaseOrderByPurchaseOrderId( "OrderDate" ), vbShortDate )%>"></td>		
		<%end if%>
			<td class="label_2">Supplier</td><td class="input_2">
				<select class="forminput" name="selSupplier" size=1 isRequired="true" onchange="this.isDirty=true;" fieldName="Supplier">
				<option value=""></option>
   	    <%if not( rsGetSupplierList is nothing ) then
   					if( rsGetSupplierList.State = adStateOpen ) then
   						while( not rsGetSupplierList.EOF )%>
   							<option value="<%=rsGetSupplierList( "SupplierId" )%>"<%if( rsGetSupplierList( "SupplierId" ) = rsGetPurchaseOrderByPurchaseOrderId( "SupplierId" ) ) then%>SELECTED<%end if%>><%=rsGetSupplierList( "SupplierName" )%></option>
   							<%rsGetSupplierList.MoveNext
   						wend
   					end if
   				end if
   			%>
				</select></td></tr>
		<tr><td class="label_1">Required<br />Ship Date</td><td class="input_1"><input type="text" class="forminput" name="txtDateRequired" onchange="this.isDirty=true;" maxlength=32 isDate="true" fieldName="Date Required" value="<%=rsGetPurchaseOrderByPurchaseOrderId( "DateRequired" )%>"></td>
				<td class="label_2">Shipping</td><td class="input_2"><select class="forminput" name="selShippingMethod" size=1 onchange="this.isDirty=true;" isRequired="true" fieldName="Shipping Method">
				<option value=""></option>
   	    <%if not( rsGetShippingMethodsList is nothing ) then
   					if( rsGetShippingMethodsList.State = adStateOpen ) then
   						while( not rsGetShippingMethodsList.EOF )%>
   							<option value="<%=rsGetShippingMethodsList( "ShippingMethodId" )%>"<%if( rsGetShippingMethodsList( "ShippingMethodId" ) = rsGetPurchaseOrderByPurchaseOrderId( "ShippingMethodId" ) ) then%>SELECTED<%end if%>><%=rsGetShippingMethodsList( "ShippingMethod" )%></option>
   							<%rsGetShippingMethodsList.MoveNext
   						wend
   					end if
   				end if
   			%>
				</select></td></tr>
		<tr><td class="label_1">Promised<br />Ship Date</td><td class="input_1"><input type="text" class="forminput" name="txtDatePromised" onchange="this.isDirty=true;" maxlength=32 isDate="true" fieldName="Date Promised" value="<%=rsGetPurchaseOrderByPurchaseOrderId( "DatePromised" )%>"></td>
		  <td class="label_2">Ship Date</td><td class="input_2"><input type="text" class="forminput" name="txtShipDate" maxlength=100 onchange="this.isDirty=true;" fieldName="Ship Date" value="<%=EscapeDoubleQuotes(rsGetPurchaseOrderByPurchaseOrderId( "ShipDate" ))%>" ID="Text3"></td></tr>
		<tr><td class="label_1">Date to Supplier</td><td class="input_1"><input type="text" class="forminput" name="txtDateShippedToSupplier" onchange="this.isDirty=true;" maxlength=32 isDate="true" fieldName="Date Shipped To Supplier" value="<%=rsGetPurchaseOrderByPurchaseOrderId( "DateShippedToSupplier" )%>" ID="Text2"></td>
		  <td class="label_2">Tracking No.</td><td class="input_2"><input type="text" class="forminput" name="txtTrackingNumber" maxlength=100 onchange="this.isDirty=true;" fieldName="Tracking Number" value="<%=EscapeDoubleQuotes(rsGetPurchaseOrderByPurchaseOrderId( "TrackingNumber" ))%>" ID="Text4"></td></tr>
		<tr><td class="label_1">Date Closed</td><td class="input_1"><input type="text" class="forminput" onchange="this.isDirty=true;" name="txtDateClosed" maxlength=32 isDate="true" fieldName="Date Closed" value="<%=rsGetPurchaseOrderByPurchaseOrderId( "DateClosed" )%>"></td>
				<td class="label_2">Buyer</td><td class="input_2">
				<select class="forminput" name="selBuyer" size=1 isRequired="true" onchange="this.isDirty=true;" fieldName="Buyer" ID="Select1">
				<option value=""></option>
   	    <%if not( rsGetEmployeeList is nothing ) then
   					if( rsGetEmployeeList.State = adStateOpen ) then
   						while( not rsGetEmployeeList.EOF )%>
   							<option value="<%=rsGetEmployeeList( "EmployeeId" )%>"<%if( rsGetEmployeeList( "EmployeeId" ) = rsGetPurchaseOrderByPurchaseOrderId( "EmployeeId" ) ) then%>SELECTED<%end if%>><%=rsGetEmployeeList( "EmployeeName" )%></option>
   							<%rsGetEmployeeList.MoveNext
   						wend
   					end if
   				end if
   			%>
				</select></td></tr>
		<tr><td class="label_1">PO Description</td><td class="input_1" colspan=3><textarea class="forminput" onchange="this.isDirty=true;" fieldName="Transaction Description" name="txtDescription" rows=6 wrap=soft ID="Textarea1"><%=rsGetPurchaseOrderByPurchaseOrderId( "PurchaseOrderDescription" )%></textarea></td></tr>
		<tr><td class="label_1">Transaction Description</td><td class="input_1" colspan=3><textarea class="forminput" onchange="this.isDirty=true;" fieldName="PO Description" name="txtPODescription" rows=6 wrap=soft ID="Textarea2"><%=rsGetPurchaseOrderByPurchaseOrderId( "PODescription" )%></textarea></td></tr>
		<tr><td class="label_1">Internal Notes</td><td class="input_1" colspan=3><textarea class="forminput" onchange="this.isDirty=true;" fieldName="Internal Notes" name="txtInternalNotes" rows=6 wrap=soft ID="Textarea3"><%=rsGetPurchaseOrderByPurchaseOrderId( "InternalNotes" )%></textarea></td></tr>
	</table>
</td></tr>

<!-- Form Bottom Rule -->
<tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>  <!-- End Form Bottom Rule -->

<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmPOMaster);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmPOMaster.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmPOMaster);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>
<!-- End Form Buttons -->

		
</form>
</td></tr>
<!-- End PO Form -->
<%	set rsGetPurchaseOrderByPurchaseOrderId = rsGetPurchaseOrderByPurchaseOrderId.NextRecordset
		end if
	end if
end if%>

<%if( intPurchaseOrderId <> 0 ) then%> 
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=4></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Products</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Transactions Table -->
 <form id="frmPODetail" name="frmPODetail" action="" method="post">
 <input type="hidden" name="txtPOIDDetail" id="txtPOIDDetail" value="<%=intPurchaseOrderId%>">
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Del</th>
    <th class="itemlist" width="15" align="center">Edit</th>
    <th class="itemlist" width="60">Date</th>
    <th class="itemlist" width="30">ID</th>
    <th class="itemlist" width="200">Product Name</th>
    <th class="itemlist">Transaction</th>
    <th class="itemlist" width=25>Qty.</th>
    <th class="itemlist" width=60>Cost</th>
    <th class="itemlist" width=70>Extended</th>
  </tr>  
 <!-- table data -->
  <%if not( rsGetPurchaseOrderByPurchaseOrderId is nothing ) then
			if( rsGetPurchaseOrderByPurchaseOrderId.State = adStateOpen ) then
				intRowNumber = 1
		
				while( not rsGetPurchaseOrderByPurchaseOrderId.EOF )
				if( ( intRowNumber Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
					<%else%>
					<tr class="oddrow">
					<%end if%>
						<td class="itemlist" width="15" align="center"><a href="javascript:ConfirmDelete('<%=rsGetPurchaseOrderByPurchaseOrderId( "TransactionId" )%>', '<%=EscapeQuotes(rsGetPurchaseOrderByPurchaseOrderId( "ProductName" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
						<td class="itemlist" width="15" align="center"><a href="javascript:OpenTransaction('<%=rsGetPurchaseOrderByPurchaseOrderId( "TransactionId" )%>');"><img src="images/form_edit.gif" width=15 height=15 border=0></a></td>
				
						<%if( not IsNull( rsGetPurchaseOrderByPurchaseOrderId( "TransactionDate" ) ) ) then%>
							<td class="itemlist" width=60 align="center"><%=FormatDateTime( rsGetPurchaseOrderByPurchaseOrderId( "TransactionDate" ), vbShortDate )%></td>
						<%else%>
							<td class="itemlist" width=60 align="center"><%=rsGetPurchaseOrderByPurchaseOrderId( "TransactionDate" )%></td>
						<%end if%>
					  <td class="itemlist" width=30 align=right><a href="javascript:OpenProduct('<%=rsGetPurchaseOrderByPurchaseOrderId( "ProductID" )%>');"><%=rsGetPurchaseOrderByPurchaseOrderId( "ProductID" )%></a></td>
					  <td class="itemlist" width=200><%=rsGetPurchaseOrderByPurchaseOrderId( "ProductName" )%></td>
					  <td class="itemlist"><%=rsGetPurchaseOrderByPurchaseOrderId( "TransactionDescription" )%></td>
					  <td class="itemlist" width=25 align=right><%=rsGetPurchaseOrderByPurchaseOrderId( "UnitsOrdered" )%></td>
					  <%if( IsNull( rsGetPurchaseOrderByPurchaseOrderId( "UnitPrice" ) ) ) then%>
					  <td class="itemlist" width=60 align=right><%=rsGetPurchaseOrderByPurchaseOrderId( "UnitPrice" )%></td>
					  <%else%>
					  <td class="itemlist" width=60 align=right><%=FormatCurrency( rsGetPurchaseOrderByPurchaseOrderId( "UnitPrice" ) )%></td>
					  <%end if
					  if( IsNull( rsGetPurchaseOrderByPurchaseOrderId( "SubTotal" ) ) ) then%>
					  <td class="itemlist" width=70 align=right><%=rsGetPurchaseOrderByPurchaseOrderId( "SubTotal" )%></td>
					  <%else%>
					  <td class="itemlist" width=70 align=right><%=FormatCurrency( rsGetPurchaseOrderByPurchaseOrderId( "SubTotal" ) )%></td>
					  <%end if%>
					</tr>
			<%
				intRowNumber = intRowNumber + 1
				rsGetPurchaseOrderByPurchaseOrderId.MoveNext
				wend
			end if
		end if%>		
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td width=15 class="itemlist">&nbsp;</td>
			<td width=15 class="itemlist">&nbsp;</td>
			<td width=60 class="itemlist" align=middle><%=FormatDateTime( Now(), vbShortDate )%></td>
			<td width=30><input type="text" class="forminput" name="txtProductId" readonly value=""></td>
			<td><table cellpadding=0 cellspacing=0 width="200"><tr><td><a href="javascript:OpenProductSearch(false);"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidProductId" isRequired="true" fieldName="Product" value=""><input class="forminput" name="txtProductName" readonly value=""></td></tr></table></td>
			<td><input type="text" class="forminput" name="txtTransactionDescription" value=""></td>
			<td width=25><input type="text" class="forminput" name="txtOrderQty" isRequired="true" isQty="true" fieldName="Qty" maxlength=32 value=""></td>
			<td width=60><input type="text" class="forminput" name="txtUnitPrice" isRequired="true" isCurrency="true" fieldName="Cost" maxlength=32 value=""></td>
			<td width=70>&nbsp;</td>
		</tr>
 </table></td>
 </tr>
 </form>
	<tr><td><table cellspacing=0 cellpadding=0 width=600 align=middle>
	  <tr><td><table cellspacing=0 cellpadding=0 width=408 align=left>
        <form id="frmPrintPO" name="frmPrintPO" action="" method="post">
	    <tr><td align=left class="generic_label" width=64><input type="radio" name="radLogo" value="centerline" checked>Centerline</td>
	      <td align=left class="generic_label" width=64><input type="radio" name="radLogo" value="css">CSS</td>
		  <td align=left width=64><a href="javascript:PrintCenterlinePO( '<%=intPurchaseOrderId%>' );"><img onmouseover="Tip('Prints PO with the line item\'s DETAILED DESCRIPION field')" src="images/print_po_button.gif" width=60 height=20 border=0></a></td>
		  <td align=left width=216><a href="javascript:PrintCenterlineSubWork( '<%=intPurchaseOrderId%>' );"><img onmouseover="Tip('Prints PO with the line item\'s TRANSACTION field')" src="images/print_sub_work_button.gif" width=60 height=20 border=0></a></td>
		</form>
	  </table></td>
	  <td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr><td align=right width=64><a href="javascript:validateForm(frmPODetail);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	  </table></td></tr>
	</table></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td><table cellspacing=0 cellpadding=0 width=300 align=middle ID="Table4">
	  <tr><td class="itemlist"><a href="javascript:OpenUsedPartsPopup( '<%=intPurchaseOrderId%>' );">Select Default Parts</a></td>
	    <td class="itemlist"><a href="javascript:OpenProductSearch( true );">Select Multiple Parts</a></td></tr>
	</table></td></tr>
 </table>
 
 </td></tr>
 <!-- End Middle Title Bar -->
 
 



<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0 ID="Table11">
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=18%>Work Orders</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

<!-- Purchase Order Table -->
<form name="frmWorkOrderPOs" action="" method="post" ID="Form1">
<input type="hidden" name="txtPOIdDetail" value="<%=intPurchaseOrderId%>" ID="Hidden1">
<tr><td align=center>
 <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle ID="Table1">
 <tr>
 <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle ID="Table2">
  
 <!-- table heading -->
 <tr>
		<th class="itemlist" width="20" align="center">Del</th>
		<th class="itemlist" width="40">Num</th>
		<th class="itemlist" width="60">Date In</th>
		<th class="itemlist" width="60">Date Out</th>
		<th class="itemlist">Spindle</th>
		<th class="itemlist">Serial #</th>
 </tr>
  
 <!-- table data -->
 <%	set rsGetPurchaseOrderByPurchaseOrderId = rsGetPurchaseOrderByPurchaseOrderId.NextRecordset
		if not( rsGetPurchaseOrderByPurchaseOrderId is nothing ) then
			if( rsGetPurchaseOrderByPurchaseOrderId.State = adStateOpen ) then
				intRowNumber = 1
		
				while( not rsGetPurchaseOrderByPurchaseOrderId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>					
						<tr class="oddrow">
					<%end if%>
							<td class="itemlist" width="20" align=middle><a href="javascript:ConfirmDeleteWorkOrderPO('<%=rsGetPurchaseOrderByPurchaseOrderId( "WorkOrderId" )%>', '<%=EscapeQuotes(rsGetPurchaseOrderByPurchaseOrderId( "PurchaseOrderId" ))%>', '<%=EscapeQuotes(rsGetPurchaseOrderByPurchaseOrderId( "WorkOrderNumber" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="40"><a href="javascript:OpenWorkOrder( '<%=rsGetPurchaseOrderByPurchaseOrderId( "WorkOrderId" )%>' )"><%=rsGetPurchaseOrderByPurchaseOrderId( "WorkOrderNumber" )%></a></td>
							<td class="itemlist"><%if( IsNull( rsGetPurchaseOrderByPurchaseOrderId( "DateIn" ) ) or ( rsGetPurchaseOrderByPurchaseOrderId( "DateIn" ) = "" ) ) then%>&nbsp;<%else%><%=FormatDateTime( rsGetPurchaseOrderByPurchaseOrderId( "DateIn" ), vbShortDate )%><%end if%></td>
							<td class="itemlist"><%if( IsNull( rsGetPurchaseOrderByPurchaseOrderId( "DateOut" ) ) or ( rsGetPurchaseOrderByPurchaseOrderId( "DateOut" ) = "" ) ) then%>&nbsp;<%else%><%=FormatDateTime( rsGetPurchaseOrderByPurchaseOrderId( "DateOut" ), vbShortDate )%><%end if%></td>
							<td class="itemlist"><a href="javascript:OpenSpindle('<%=rsGetPurchaseOrderByPurchaseOrderId( "SpindleId" )%>');"><%=rsGetPurchaseOrderByPurchaseOrderId( "SpindleType" )%></a></td>
							<td class="itemlist"><%=rsGetPurchaseOrderByPurchaseOrderId( "SerialNumber" )%></td>
						</tr>
					<%intRowNumber = intRowNumber + 1
					rsGetPurchaseOrderByPurchaseOrderId.MoveNext
				wend
			end if
		end if%>
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td class="itemlist" width=20 align=middle>&nbsp;</td>
			<td class="itemlist" colspan="6">
				<table cellpadding=0 cellspacing=0 width=100% ID="Table3"><tr><td><a href="javascript:OpenWorkOrderSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input class="forminput" name="txtWorkOrderId" readonly isRequired="true" onchange="this.isDirty=true;" fieldName="Work Order Id" value="" ID="Text1"></td></tr></table>
			</td>
		</tr>
</table></td></tr>
<!-- End Purchase Order Table -->

<tr><td align=center>
<table width=128 border=0 cellspacing=0 cellpadding=0 align=right ID="Table14">
  <tr><td width=64 align=right><a href="javascript:validateForm(frmWorkOrderPOs);"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
		<td width=64 align=right><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
</table></td></tr>
</form></table>

 <%end if%>
</body>
</html>
<%
if not( rsGetShippingMethodsList is nothing ) then
	if( rsGetShippingMethodsList.State = adStateOpen ) then
		rsGetShippingMethodsList.Close
	end if
end if

if not( rsGetEmployeeList is nothing ) then
	if( rsGetEmployeeList.State = adStateOpen ) then
		rsGetEmployeeList.Close
	end if
end if

if not( rsGetSupplierList is nothing ) then
	if( rsGetSupplierList.State = adStateOpen ) then
		rsGetSupplierList.Close
	end if
end if

if not( rsGetProductListByName is nothing ) then
	if( rsGetProductListByName.State = adStateOpen ) then
		rsGetProductListByName.Close
	end if
end if

if not( rsGetPurchaseOrderByPurchaseOrderId is nothing ) then
	if( rsGetPurchaseOrderByPurchaseOrderId.State = adStateOpen ) then
		rsGetPurchaseOrderByPurchaseOrderId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetShippingMethodsList							= nothing
set cmdGetShippingMethodsList							= nothing
set rsGetEmployeeList											= nothing
set cmdGetEmployeeList										= nothing
set rsGetSupplierList											= nothing
set cmdGetSupplierList										= nothing
set rsGetProductListByName											= nothing
set cmdGetProductListByName											= nothing
set rsGetPurchaseOrderByPurchaseOrderId		= nothing
set cmdGetPurchaseOrderByPurchaseOrderId	= nothing
set cmdEditPurchaseOrder									= nothing
set cnnSQLConnection											= nothing
%>
