<%option explicit

dim cnnSQLConnection
dim rsPrintPurchaseOrderCenterline, cmdPrintPurchaseOrderCenterline
dim rsGetCompanyInformation, cmdGetCompanyInformation
dim intPurchaseOrderId, curOrderTotal

set rsPrintPurchaseOrderCenterline		= nothing
set cmdPrintPurchaseOrderCenterline		= nothing
set rsGetCompanyInformation						= nothing
set cmdGetCompanyInformation					= nothing

set cnnSQLConnection									= Server.CreateObject( "ADODB.Connection" )
set cmdPrintPurchaseOrderCenterline		= Server.CreateObject( "ADODB.Command" )
set cmdGetCompanyInformation					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intPurchaseOrderId = Request.QueryString( "intPurchaseOrderId" )

cmdPrintPurchaseOrderCenterline.ActiveConnection = cnnSQLConnection
cmdPrintPurchaseOrderCenterline.CommandText = "{call spPrintSubWorkCenterline( ? ) }"
cmdPrintPurchaseOrderCenterline.Parameters.Append cmdPrintPurchaseOrderCenterline.CreateParameter( "@PurchaseOrderId", adInteger, adParamInput, 4 )
cmdPrintPurchaseOrderCenterline.Parameters( "@PurchaseOrderId" )	= intPurchaseOrderId

on error resume next
set rsPrintPurchaseOrderCenterline = cmdPrintPurchaseOrderCenterline.Execute
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
if not( rsPrintPurchaseOrderCenterline is nothing ) then
		if( rsPrintPurchaseOrderCenterline.State = adStateOpen ) then
			if( not rsPrintPurchaseOrderCenterline.EOF ) then%>
<HTML>
<TITLE>Purchase Order - Centerline</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
</HEAD>

<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width=640>
  
    <tr><td><table width=640 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo.png" width=282 height=138></td>
       <td align=right><table>
        <tr><td class="report_title">Purchase Order</td></tr>    
        <tr><td class="contact_info" width=100%>
         <table>
          <tr><td colspan=2><%=rsGetCompanyInformation( "CompanyName" )%></td></tr>
          <tr><td colspan=2><%=rsGetCompanyInformation( "Address" )%></td></tr>
          <tr><td><%=rsGetCompanyInformation( "City" )%>,&nbsp;<%=rsGetCompanyInformation( "StateOrProvince" )%>&nbsp;<%=rsGetCompanyInformation( "PostalCode" )%></td><td align=middle><%=rsGetCompanyInformation( "Country" )%></td></tr>
          <tr><td>Phone:&nbsp;<%=rsGetCompanyInformation( "PhoneNumber" )%>&nbsp;&nbsp;&nbsp;</td><td align=right>Fax:&nbsp;<%=rsGetCompanyInformation( "FaxNumber" )%></td></tr>
         </table></td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="purchaser_info"><table width="100%">
      <tr><td class="purchaser_label">Ordered From:</td>
        <td colspan=3><%=rsPrintPurchaseOrderCenterline( "SupplierName" )%></td>
        <td align=right><table cellspacing=0 cellpadding=0><tr><td class="purchaser_label">PO Number:</td>
          <td class="po_number">&nbsp;&nbsp;<%=rsPrintPurchaseOrderCenterline( "PurchaseOrderNumber" )%></td></tr></table></td></tr>
      <tr><td>&nbsp;</td><td><%=rsPrintPurchaseOrderCenterline( "Address" )%></td></tr>
      <tr><td>&nbsp;</td><td colspan=2><%=rsPrintPurchaseOrderCenterline( "City" )%>,&nbsp;<%=rsPrintPurchaseOrderCenterline( "StateOrProvince" )%>&nbsp;<%=rsPrintPurchaseOrderCenterline( "PostalCode" )%></td><td><%=rsPrintPurchaseOrderCenterline( "Country" )%></td></tr>
      <tr><td>&nbsp;</td><td colspan=2><%=rsPrintPurchaseOrderCenterline( "ContactName" )%></td>
        <td align=right><table cellspacing=0 cellpadding=0><tr><td class="purchaser_label" align="right">Phone:&nbsp;&nbsp;</td><td><%=rsPrintPurchaseOrderCenterline( "PhoneNumber" )%></td></tr></table></td>
        <td align=right><table cellspacing=0 cellpadding=0><tr><td class="purchaser_label" align="right">Fax:&nbsp;&nbsp;</td><td><%=rsPrintPurchaseOrderCenterline( "FaxNumber" )%></td></tr></table></td></tr>
    </table></td></tr>
    
    <tr><td class="order_info"><table width="100%">
      <tr><td class="order_label">Order Date: </td><td><%=rsPrintPurchaseOrderCenterline( "OrderDate" )%></td>
        <td class="order_label">Date Required: </td><td><%=rsPrintPurchaseOrderCenterline( "DateRequired" )%></td>
        <td class="order_label">Buyer: </td><td><%=rsPrintPurchaseOrderCenterline( "BuyerName" )%></td></tr>
      <tr><td class="order_label">PO ID: </td><td><%=rsPrintPurchaseOrderCenterline( "PurchaseOrderId" )%></td>
        <td class="order_label">Date Promised: </td><td><%=rsPrintPurchaseOrderCenterline( "DatePromised" )%></td>
        <td class="order_label">Ship Via: </td><td><%=rsPrintPurchaseOrderCenterline( "ShippingMethod" )%></td></tr>
    </table></td></tr>
    
    <tr><td class="desc_info"><table align=left>
      <tr><td class="desc_label">Description:&nbsp;&nbsp;</td><td class="order_desc"><%=rsPrintPurchaseOrderCenterline( "PurchaseOrderDescription" )%></td></tr>
    </table></td></tr>
    
    <%set rsPrintPurchaseOrderCenterline = rsPrintPurchaseOrderCenterline.NextRecordset
			if not( rsPrintPurchaseOrderCenterline is nothing ) then
				if( rsPrintPurchaseOrderCenterline.State = adStateOpen ) then
					if( not rsPrintPurchaseOrderCenterline.EOF ) then
						curOrderTotal = 0%>    
						<tr><td class="item_info"><table cellspacing=0 cellpadding=4 width=640>
						  <tr><td class="item_label" width=5%>ID</td>
						    <td class="item_label" width=20%>Product Name</td>
						    <td class="item_label" width=35%>Description</td>
						    <td class="item_label" width=10% align=right>Units</td>
						    <td class="item_label" width=15% align=right>Unit Price</td>
						    <td class="item_label" width=15% align=right>Subtotal</td></tr>
						  <%while not rsPrintPurchaseOrderCenterline.EOF%>
						  <tr><td class="item_idnum"><%=rsPrintPurchaseOrderCenterline( "ProductId" )%></td>
						    <td class="item_item"><%=rsPrintPurchaseOrderCenterline( "ProductName" )%></td>
						    <td class="item_desc"><%=rsPrintPurchaseOrderCenterline( "TransactionDescription" )%></td>
						    <td class="item_units"><%=rsPrintPurchaseOrderCenterline( "UnitsOrdered" )%></td>
						    <td class="item_unit_price"><%=FormatCurrency( rsPrintPurchaseOrderCenterline( "UnitPrice" ), 2 )%></td>
						    <td class="item_subtotal"><%=FormatCurrency( rsPrintPurchaseOrderCenterline( "SubTotal" ), 2 )%></td></tr>
						  <% rsPrintPurchaseOrderCenterline.MoveNext
						  wend%>
        <%end if
				end if
			end if%> 
			<%set rsPrintPurchaseOrderCenterline = rsPrintPurchaseOrderCenterline.NextRecordset
			if not( rsPrintPurchaseOrderCenterline is nothing ) then
				if( rsPrintPurchaseOrderCenterline.State = adStateOpen ) then
					if( not rsPrintPurchaseOrderCenterline.EOF ) then
			%>
			
      <tr><td class="order_total_label" colspan=5>Order Total</td>
        <td class="order_total"><%=FormatCurrency( rsPrintPurchaseOrderCenterline( "OrderTotal" ), 2 )%></td></tr>
        <%end if
				end if
			end if%> 
        
    </table></td></tr>        
    
  </table>
</BODY>
</HTML>
<%	end if
	end if
end if

if not( rsPrintPurchaseOrderCenterline is nothing ) then
	if( rsPrintPurchaseOrderCenterline.State = adStateOpen ) then
		rsPrintPurchaseOrderCenterline.Close
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

set rsGetCompanyInformation					= nothing
set cmdGetCompanyInformation				= nothing
set rsPrintPurchaseOrderCenterline	= nothing
set cmdPrintPurchaseOrderCenterline	= nothing
set cnnSQLConnection								= nothing
%>
