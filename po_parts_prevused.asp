<%@ Language=VBScript %>
<%option explicit

dim cnnSQLConnection
dim rsGetSpindleProductsByPurchaseOrderId, cmdGetSpindleProductsByPurchaseOrderId
dim intPurchaseOrderId, intRowNumber, strCurrentPage
dim cmdEditPurchaseOrder
dim i
dim windowFreshener, strCategory

set cnnSQLConnection									= nothing
set rsGetSpindleProductsByPurchaseOrderId						= nothing
set cmdGetSpindleProductsByPurchaseOrderId					= nothing
set cmdEditPurchaseOrder										= nothing

set cnnSQLConnection									= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleProductsByPurchaseOrderId					= Server.CreateObject( "ADODB.Command" )
windowFreshener = false

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


					
'user hit save on the form, so we need to save all the items to the DB
if( Request.Form.Count > 0 ) then
		intPurchaseOrderId = Request.Form("intPurchaseOrderId")
		windowFreshener = true
		
			set cmdEditPurchaseOrder	= Server.CreateObject( "ADODB.Command" )
		
			cmdEditPurchaseOrder.ActiveConnection	= cnnSQLConnection
			cmdEditPurchaseOrder.CommandText = "spEditPurchaseOrderDetail"
			cmdEditPurchaseOrder.CommandType = adCmdStoredProc 
			
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionDate", adVarChar, adParamInput, 10 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@ProductID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@PurchaseOrderID", adInteger, adParamInput, 4 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@TransactionDescription", adLongVarChar, adParamInput, 1048576 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@UnitPrice", adCurrency, adParamInput, 10 )
			cmdEditPurchaseOrder.Parameters.Append cmdEditPurchaseOrder.CreateParameter( "@UnitsOrdered", adVarChar, adParamInput, 10 )
		
				
		for i = 1 to CInt(Request.form("checkboxCount")) 
			if Request.Form("usedPart"+CStr(i)) = "on" then
				cmdEditPurchaseOrder.Parameters( "@TransactionDate" ) = FormatDateTime( Now(), vbShortDate )
				cmdEditPurchaseOrder.Parameters( "@ProductId" ) = Request.Form("productId" + CStr(i))		
				cmdEditPurchaseOrder.Parameters( "@PurchaseOrderId" ) = intPurchaseOrderId
				cmdEditPurchaseOrder.Parameters( "@UnitPrice" ) = Request.Form("cost" + CStr(i))
				cmdEditPurchaseOrder.Parameters( "@UnitsOrdered" ) = Request.Form("qty" + CStr(i))
				
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
			end if
		next
		
		'close window, and refresh parent
		
else

	intPurchaseOrderId = Request.QueryString( "POId" )
	
	cmdGetSpindleProductsByPurchaseOrderId.ActiveConnection = cnnSQLConnection
	cmdGetSpindleProductsByPurchaseOrderId.CommandText = "spGetSpindleProductsByPurchaseOrderId"
	cmdGetSpindleProductsByPurchaseOrderId.CommandType = adCmdStoredProc 
	cmdGetSpindleProductsByPurchaseOrderId.Parameters.Append cmdGetSpindleProductsByPurchaseOrderId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetSpindleProductsByPurchaseOrderId.Parameters.Append cmdGetSpindleProductsByPurchaseOrderId.CreateParameter( "@PurchaseOrderId", adInteger, adParamInput, 4 )
	cmdGetSpindleProductsByPurchaseOrderId.Parameters( "@PurchaseOrderId" )	= intPurchaseOrderId
	
	'Response.Write( "no variables set")
	
	on error resume next
	set rsGetSpindleProductsByPurchaseOrderId = cmdGetSpindleProductsByPurchaseOrderId.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
end if
%>


<HTML>
<HEAD>
<TITLE>Centerline - View / Add Previous Used Parts</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function CloseAndRefreshParent()
	{
		window.opener.RefreshCurrentPage();
		window.close();
	}
	
	function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'CenterlineProduct'+strProductId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
  </script>
</HEAD>
<BODY <%if (windowFreshener = true) then %> onload="javascript:CloseAndRefreshParent();" <%end if%>>

<form id="prevUsedParts" name="prevUsedParts" action="" method="post">
<table width="600" ID="Table1">
<tr><td>
  <table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle ID="Table2">
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Add</th>
    <th class="itemlist" width="25" align="center">Id</th>
    <th class="itemlist" width="200">Part #</th>
    <th class="itemlist" width="200">Description</th>
    <th class="itemlist" width="70">Cost</th>
    <th class="itemlist" width="30">Qty</th>
    <th class="itemlist" width="30">OH</th>
    <th class="itemlist" width="30">OO</th>
  </tr>  
  
 <!-- table data -->
<%if not( rsGetSpindleProductsByPurchaseOrderId is nothing ) then
	if( rsGetSpindleProductsByPurchaseOrderId.State = adStateOpen ) then
		intRowNumber = 1
			
		while( not rsGetSpindleProductsByPurchaseOrderId.EOF )
			if( ( intRowNumber Mod 2 ) = 1 ) then%>
				<tr class="evenrow">
			<%else%>
				<tr class="oddrow">
			<%end if%>
 
			<input type=hidden name="spindleProduct<%=intRowNumber%>" value="<%=rsGetSpindleProductsByPurchaseOrderId("SpindleProductId")%>" ID="Hidden1">
			<td class="itemlist" width="15" align=middle><input type="checkbox" class="forminput" name="usedpart<%=intRowNumber%>" fieldName="usedpart" ID="Checkbox1"></td>
			
			<input type=hidden name="productID<%=intRowNumber%>" value="<%=rsGetSpindleProductsByPurchaseOrderId("ProductID")%>" ID="Hidden2">
			<td class="itemlist" width="25"><a href="javascript:OpenProduct( '<%=rsGetSpindleProductsByPurchaseOrderId( "ProductID" )%>' );"><%=rsGetSpindleProductsByPurchaseOrderId( "ProductID" )%></a></td>
			
			<td class="itemlist" width="200"><%=rsGetSpindleProductsByPurchaseOrderId( "PartNumber" )%></td>
			
			<td class="itemlist" width="200"><%=rsGetSpindleProductsByPurchaseOrderId( "ProductShortDescription" )%></td>
			
			<input type=hidden name="cost<%=intRowNumber%>" value="<%=rsGetSpindleProductsByPurchaseOrderId( "Cost" )%>" ID="Hidden4">
			<td class="itemlist" width="70" align="right"><%=FormatCurrency( rsGetSpindleProductsByPurchaseOrderId( "Cost" ), 2 )%></td>
			
			<input type=hidden name="qty<%=intRowNumber%>" value="<%=rsGetSpindleProductsByPurchaseOrderId("Quantity")%>" ID="Hidden6">
			<td class="itemlist" width="30" align="right"><%=rsGetSpindleProductsByPurchaseOrderId( "Quantity" )%></td>
			
			<td class="itemlist" width="30" align="right"><%=rsGetSpindleProductsByPurchaseOrderId( "OnHand" )%></td>
			<td class="itemlist" width="50" align="right"><%=rsGetSpindleProductsByPurchaseOrderId( "OnOrder" )%></td>
			
			</tr>
			<%intRowNumber = intRowNumber + 1
			rsGetSpindleProductsByPurchaseOrderId.MoveNext
		wend%>
		
		<input type="hidden" name="checkboxCount" value="<%=intRowNumber-1%>" ID="Hidden7">
		<input type="hidden" name="intPurchaseOrderId" value="<%=intPurchaseOrderId%>" ID="Hidden8">
		<input type="hidden" name="strCategory" value="<%=strCategory%>" ID="Hidden9">
	<%end if
end if%>
</table></td></tr>


<tr><td><table cellspacing=0 cellpadding=0 width=128 align=right border=0 ID="Table3">
	<tr><td width="64" align="right"><a href="javascript:validateForm(prevUsedParts);"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td>
	<td width="64" align="right"><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
</table></td></tr>


</form>
</BODY>
</HTML>

<%
if not( rsGetSpindleProductsByPurchaseOrderId is nothing ) then
	if( rsGetSpindleProductsByPurchaseOrderId.State = adStateOpen ) then
		rsGetSpindleProductsByPurchaseOrderId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleProductsByPurchaseOrderId				= nothing
set cmdGetSpindleProductsByPurchaseOrderId			= nothing
set cmdEditPurchaseOrder								= nothing
set cnnSQLConnection							= nothing
%>
