<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetInventoryTransactionByInventoryTransactionId, cmdGetInventoryTransactionByInventoryTransactionId
dim cmdEditInventoryTransaction
dim intInventoryTransactionId, windowFreshener, isRefreshParent
windowFreshener = false

if ( Request.QueryString( "isRefreshParent" ) = "true" ) then
	isRefreshParent = true
else
	isRefreshParent = false
end if

set cnnSQLConnection																		= nothing
set rsGetInventoryTransactionByInventoryTransactionId		= nothing
set cmdGetInventoryTransactionByInventoryTransactionId	= nothing
set cmdEditInventoryTransaction													= nothing

set cnnSQLConnection																		= Server.CreateObject( "ADODB.Connection" )
set cmdGetInventoryTransactionByInventoryTransactionId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
if( Request.Form.Count > 0 ) then
	windowFreshener = true
	set cmdEditInventoryTransaction	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditInventoryTransaction.ActiveConnection	= cnnSQLConnection
	cmdEditInventoryTransaction.CommandText			= "{? = call spEditInventoryTransaction( ?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@TransactionID", adInteger, adParamInput, 4 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@TransactionDescription", adLongVarChar, adParamInput, 1048576 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@UnitPrice", adCurrency, adParamInput, 10 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@UnitsOrdered", adVarChar, adParamInput, 10 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@DateReceived", adVarChar, adParamInput, 10 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@UnitsReceived", adVarChar, adParamInput, 10 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@UnitsSold", adVarChar, adParamInput, 10 )
	cmdEditInventoryTransaction.Parameters.Append cmdEditInventoryTransaction.CreateParameter( "@UnitsShrinkage", adVarChar, adParamInput, 10 )
	
	if( Request.Form( "txtTransactionId" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@TransactionID" ) = Request.Form( "txtTransactionId" )
	if( Request.Form( "txtTransactionDescription" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@TransactionDescription" ) = Request.Form( "txtTransactionDescription" )
	if( Request.Form( "txtUnitPrice" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@UnitPrice" ) = Request.Form( "txtUnitPrice" )
	if( Request.Form( "txtUnitsOrdered" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@UnitsOrdered" ) = Request.Form( "txtUnitsOrdered" )
	if( Request.Form( "txtDateReceived" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@DateReceived" ) = Request.Form( "txtDateReceived" )
	if( Request.Form( "txtUnitsReceived" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@UnitsReceived" ) = Request.Form( "txtUnitsReceived" )
	if( Request.Form( "txtUnitsSold" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@UnitsSold" ) = Request.Form( "txtUnitsSold" )
	if( Request.Form( "txtUnitsShrinkage" ) <> "" ) then cmdEditInventoryTransaction.Parameters( "@UnitsShrinkage" ) = Request.Form( "txtUnitsShrinkage" )
	
	on error resume next
	cmdEditInventoryTransaction.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0
	
	intInventoryTransactionId = cmdEditInventoryTransaction.Parameters( "@Return" ).Value
else
	if( Request.QueryString( "intTransactionId" ) = "" ) then
		intInventoryTransactionId = 0
	else
		intInventoryTransactionId = Request.QueryString( "intTransactionId" )
	end if
end if

cmdGetInventoryTransactionByInventoryTransactionId.ActiveConnection = cnnSQLConnection
cmdGetInventoryTransactionByInventoryTransactionId.CommandText = "{call spGetInventoryTransactionByInventoryTransactionId( ?,? ) }"
cmdGetInventoryTransactionByInventoryTransactionId.Parameters.Append cmdGetInventoryTransactionByInventoryTransactionId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetInventoryTransactionByInventoryTransactionId.Parameters.Append cmdGetInventoryTransactionByInventoryTransactionId.CreateParameter( "@InventoryTransactionId", adInteger, adParamInput, 4 )
cmdGetInventoryTransactionByInventoryTransactionId.Parameters( "@InventoryTransactionId" ) = intInventoryTransactionId

on error resume next
set rsGetInventoryTransactionByInventoryTransactionId = cmdGetInventoryTransactionByInventoryTransactionId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0
'**********************************************************************************
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - View / Edit Transaction</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	function OpenPurchaseOrder( strPurchaseOrderId )
	{
	  OpenWindow( 'po_details.asp?intPurchaseOrderId='+strPurchaseOrderId,'PurchaseOrder','width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'Product','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	function CloseAndRefreshParent()
	{
		window.opener.RefreshCurrentPage();
		window.close();
	}
	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if ((windowFreshener = true) and (isRefreshParent = true)) then %> onload="javascript:CloseAndRefreshParent();" <%end if%>>

<%if not( rsGetInventoryTransactionByInventoryTransactionId is nothing ) then
		if( rsGetInventoryTransactionByInventoryTransactionId.State = adStateOpen ) then
			if( not rsGetInventoryTransactionByInventoryTransactionId.EOF ) then%>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=30%>Transaction #<%=intInventoryTransactionId%></td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Transaction Form -->
<form id="frmTransaction" name="frmTransaction" action="" method="post">
<input type="hidden" name="txtTransactionId" id="txtTransactionId" value="<%=intInventoryTransactionId%>">
<tr><td align=center>
    <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
  		<tr><td class="label_1">Date</td><td class="input_1"><input type="text" class="forminput" name="txtDate" maxlength=16 readonly value="<%=rsGetInventoryTransactionByInventoryTransactionId( "TransactionDate" )%>"></td>
       <td class="label_2">PO</td><td class="generic_input"><a href="javascript:OpenPurchaseOrder( '<%=rsGetInventoryTransactionByInventoryTransactionId( "PurchaseOrderId" )%>' );"><%=rsGetInventoryTransactionByInventoryTransactionId( "PurchaseOrderId" )%></a></td></tr>
     <tr><td class="label_1">Product</td><td class="generic_input"><a href="javascript:OpenProduct( '<%=rsGetInventoryTransactionByInventoryTransactionId( "ProductId" )%>' );"><%=rsGetInventoryTransactionByInventoryTransactionId( "ProductId" )%></a></td>
	      
     <tr><td class="label_1">Units Ordered</td><td class="input_1"><input type="text" class="forminput" name="txtUnitsOrdered" onchange="this.isDirty=true;" fieldName="Units Ordered"  isQty="true" maxlength=10 value="<%=rsGetInventoryTransactionByInventoryTransactionId( "UnitsOrdered" )%>"></td>
       <td class="label_2">Unit Price</td><td class="input_2"><input type="text" class="forminput" name="txtUnitPrice" onchange="this.isDirty=true;" isCurrency="true" fieldName="Unit Price" maxlength=10 value="<%=FormatCurrency( rsGetInventoryTransactionByInventoryTransactionId( "UnitPrice" ) )%>"></td></tr>
     <tr><td class="label_1">Units Received</td><td class="input_1"><input type="text" class="forminput" name="txtUnitsReceived" onchange="this.isDirty=true;" fieldName="Units Received"  isQty="true" maxlength=10 value="<%=rsGetInventoryTransactionByInventoryTransactionId( "UnitsReceived" )%>"></td>
       <td class="label_2">Date Rcvd</td><td class="input_2"><input type="text" class="forminput" name="txtDateReceived" onchange="this.isDirty=true;" isDate="true" fieldName="Date Recvd" maxlength=10 value="<%=rsGetInventoryTransactionByInventoryTransactionId( "DateReceived" )%>"></td></tr>
     <tr><td class="label_1">Units Sold</td><td class="input_1"><input type="text" class="forminput" name="txtUnitsSold" isQty="true" fieldName="Units Sold" onchange="this.isDirty=true;" maxlength=10 value="<%=rsGetInventoryTransactionByInventoryTransactionId( "UnitsSold" )%>"></td>
       <td class="label_2">Shrinkage</td><td class="input_2"><input type="text" class="forminput" name="txtUnitsShrinkage" isQty="true" fieldName="Shrinkage" onchange="this.isDirty=true;" maxlength=10 value="<%=rsGetInventoryTransactionByInventoryTransactionId( "UnitsShrinkage" )%>"></td></tr>
     <tr><td class="label_1" colspan=4>Description</td>
 	<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtTransactionDescription" onchange="this.isDirty=true;" fieldName="Description" rows=6 wrap=soft><%=rsGetInventoryTransactionByInventoryTransactionId( "TransactionDescription" )%></textarea></td></tr>
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
    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
          <tr><td width=64 align=right><a href="javascript:validateForm(frmTransaction);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmTransaction.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:GetDirty(frmTransaction);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
<!-- End Product Form -->
 <%	end if
	end if
end if%>
</BODY>
</HTML>
<%
if not( rsGetInventoryTransactionByInventoryTransactionId is nothing ) then
	if( rsGetInventoryTransactionByInventoryTransactionId.State = adStateOpen ) then
		rsGetInventoryTransactionByInventoryTransactionId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetInventoryTransactionByInventoryTransactionId		= nothing
set cmdGetInventoryTransactionByInventoryTransactionId	= nothing
set cmdEditInventoryTransaction													= nothing
set cnnSQLConnection																		= nothing
%>