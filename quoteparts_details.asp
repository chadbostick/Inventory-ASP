<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetQuotePartsByQuoteId, cmdGetQuotePartsByQuoteId
dim cmdEditQuote, cmdEditSpindleProduct, cmdCheckSpindleProduct
dim intQuoteId, intRowNumber, strCurrentPage
dim theSpindleId, theProductId, retVal

set cnnSQLConnection							= nothing
set rsGetQuotePartsByQuoteId					= nothing
set cmdGetQuotePartsByQuoteId					= nothing
set cmdEditQuote								= nothing
set cmdEditSpindleProduct						= nothing
set cmdCheckSpindleProduct						= nothing


set cnnSQLConnection							= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuotePartsByQuoteId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	if( Request.Form( "txtQuoteIdMaster" ) <> "" ) then
	
		set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQuote.ActiveConnection	= cnnSQLConnection
		cmdEditQuote.CommandText			= "{? = call spEditQuotePart( ?,?,?,? ) }"
		
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteID", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@PartsCommission", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@FreightChargeParts", adVarChar, adParamInput, 10 )
			
		if( Request.Form( "txtQuoteIdMaster" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteIdMaster" )
		
		'if( Left( Request.Form( "txtPartsCommission" ), 1 ) = "$" ) then
		'	if( Request.Form( "txtPartsCommission" ) <> "" ) then cmdEditQuote.Parameters( "@PartsCommission" ) = Mid( Request.Form( "txtPartsCommission" ), 2 )
		'else
		'	if( Request.Form( "txtPartsCommission" ) <> "" ) then cmdEditQuote.Parameters( "@PartsCommission" ) = Request.Form( "txtPartsCommission" )
		'end if
		
		if( Left( Request.Form( "txtFreightCharge" ), 1 ) = "$" ) then
			if( Request.Form( "txtFreightCharge" ) <> "" ) then cmdEditQuote.Parameters( "@FreightChargeParts" ) = Mid( Request.Form( "txtFreightCharge" ), 2 )
		else
			if( Request.Form( "txtFreightCharge" ) <> "" ) then cmdEditQuote.Parameters( "@FreightChargeParts" ) = Request.Form( "txtFreightCharge" )
		end if
		
		on error resume next
		cmdEditQuote.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if

		on error goto 0
	
		intQuoteId = cmdEditQuote.Parameters( "@Return" ).Value
		
	elseif( Request.Form( "txtQuoteIdDetail" ) <> "" ) then
	
		set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQuote.ActiveConnection	= cnnSQLConnection
		cmdEditQuote.CommandText			= "{? = call spEditQuotePartDetail( ?,?,?,?,?,?,?,? ) }"
		
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuotePartId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ProductId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@PartCost", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Markup", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Qty", adVarChar, adParamInput, 10 )
		
		if( Request.Form( "txtQuoteIdDetail" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteIdDetail" )
		if( Request.Form( "hidProductId" ) <> "" ) then cmdEditQuote.Parameters( "@ProductId" ) = Request.Form( "hidProductId" )
		
		'if( Request.Form( "txtCost" ) <> "" ) then cmdEditQuote.Parameters( "@PartCost" ) = Request.Form( "txtCost" )
		if( Left( Request.Form( "txtCost" ), 1 ) = "$" ) then
			cmdEditQuote.Parameters( "@PartCost" ) = Mid( Request.Form( "txtCost" ), 2 )
		else
			cmdEditQuote.Parameters( "@PartCost" ) = Request.Form( "txtCost" )
		end if
		
		if( Request.Form( "txtMarkup" ) <> "" ) then cmdEditQuote.Parameters( "@Markup" ) = Request.Form( "txtMarkup" )
		if( Request.Form( "hidSupplierId" ) <> "" ) then cmdEditQuote.Parameters( "@SupplierId" ) = Request.Form( "hidSupplierId" )
		if( Request.Form( "txtQty" ) <> "" ) then cmdEditQuote.Parameters( "@Qty" ) = Request.Form( "txtQty" )
		
		on error resume next
		cmdEditQuote.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if

		on error goto 0
		
		
		
		
		'get the spindle and product to add
		theSpindleId = Request.Form( "theSpindleId" )
		theProductId = Request.Form( "hidProductId" )
		
		'if the spindle and product are already associated in the DB then don't add them this time
		set cmdCheckSpindleProduct	= Server.CreateObject( "ADODB.Command" )
	
		cmdCheckSpindleProduct.ActiveConnection	= cnnSQLConnection
		cmdCheckSpindleProduct.CommandText			= "{? = call spCheckSpindleProduct( ?,?,? ) }"
		
		cmdCheckSpindleProduct.Parameters.Append cmdCheckSpindleProduct.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdCheckSpindleProduct.Parameters.Append cmdCheckSpindleProduct.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdCheckSpindleProduct.Parameters.Append cmdCheckSpindleProduct.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
		cmdCheckSpindleProduct.Parameters.Append cmdCheckSpindleProduct.CreateParameter( "@ProductId", adInteger, adParamInput, 4 )
		
		cmdCheckSpindleProduct.Parameters( "@SpindleId" ) = theSpindleId
		cmdCheckSpindleProduct.Parameters( "@ProductId" ) = theProductId
		
		on error resume next
		cmdCheckSpindleProduct.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
		on error goto 0
		
		'Find out if the part is already associated with our spindle
		retVal = cmdCheckSpindleProduct.Parameters( "@Return" ).Value
		
		'if the part isn't associated, then add it to the association table
		if (retVal = 0) then
			set cmdEditSpindleProduct	= Server.CreateObject( "ADODB.Command" )
	
			cmdEditSpindleProduct.ActiveConnection	= cnnSQLConnection
			cmdEditSpindleProduct.CommandText			= "{? = call spEditSpindleProduct( ?,?,?,?,?,?,?,? ) }"
		
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@SpindleProductId", adInteger, adParamInput, 4 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@ProductId", adInteger, adParamInput, 4 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@PartCost", adVarChar, adParamInput, 10 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@Markup", adVarChar, adParamInput, 10 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
			cmdEditSpindleProduct.Parameters.Append cmdEditSpindleProduct.CreateParameter( "@Qty", adVarChar, adParamInput, 10 )
		
		
		
			cmdEditSpindleProduct.Parameters( "@SpindleProductId" ) = 0
			cmdEditSpindleProduct.Parameters( "@SpindleId" ) = Request.Form( "theSpindleId" )
			if( Request.Form( "hidProductId" ) <> "" ) then cmdEditSpindleProduct.Parameters( "@ProductId" ) = Request.Form( "hidProductId" )
			
			'if( Request.Form( "txtCost" ) <> "" ) then cmdEditSpindleProduct.Parameters( "@PartCost" ) = Request.Form( "txtCost" )
			if( Left( Request.Form( "txtCost" ), 1 ) = "$" ) then
				cmdEditSpindleProduct.Parameters( "@PartCost" ) = Mid( Request.Form( "txtCost" ), 2 )
			else
				cmdEditSpindleProduct.Parameters( "@PartCost" ) = Request.Form( "txtCost" )
			end if
		
			if( Request.Form( "txtMarkup" ) <> "" ) then cmdEditSpindleProduct.Parameters( "@Markup" ) = Request.Form( "txtMarkup" )
			if( Request.Form( "hidSupplierId" ) <> "" ) then cmdEditSpindleProduct.Parameters( "@SupplierId" ) = Request.Form( "hidSupplierId" )
			if( Request.Form( "txtQty" ) <> "" ) then cmdEditSpindleProduct.Parameters( "@Qty" ) = Request.Form( "txtQty" )
		
		
			on error resume next
			cmdEditSpindleProduct.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if

			on error goto 0
		end if	
		
		
		intQuoteId = Request.Form( "txtQuoteIdDetail" )
		
	end if
else
	intQuoteId = Request.QueryString( "intQuoteId" )
end if

cmdGetQuotePartsByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuotePartsByQuoteId.CommandText = "{call spGetQuotePartsByQuoteId( ?,? ) }"
cmdGetQuotePartsByQuoteId.Parameters.Append cmdGetQuotePartsByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuotePartsByQuoteId.Parameters.Append cmdGetQuotePartsByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuotePartsByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetQuotePartsByQuoteId = cmdGetQuotePartsByQuoteId.Execute
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
<TITLE>Centerline - View / Edit Parts Cost</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function ConfirmDelete( intQuotePartId, strPart )
	{
		if( confirm( 'Are you sure you want to remove item ' + strPart + '?' ) )
		{
			OpenWindow( 'quoteparts_delete.asp?intQuotePartId='+intQuotePartId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'CenterlineProduct'+strProductId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenPartsPopup( )
	{
	  OpenWindow( 'quoteparts_details_popup.asp','SelectPart','width=660,height=200,left='+GetCenterX( 660 )+',top='+GetCenterY( 200 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 200 )+',scrollbars=no,dependent=yes' );
	}
	
	function OpenUsedPartsPopup(intSpindleID, intPartType)
	{
	  OpenWindow( 'quote_parts_prevused.asp?spindleID=' + intSpindleID + '&partType=' + intPartType + '&intQuoteId=<%=intQuoteId%>' + '&strCategory=Part', 'SelectPrevUsedParts','width=660,height=400,left='+GetCenterX( 660 )+',top='+GetCenterY( 400 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 400 )+',scrollbars=yes,dependent=yes' );
	}
	
	function EditQuotePart( intQuotePartId)
	{
		OpenWindow( 'quote_parts_details_edit.asp?intQuotePartId='+intQuotePartId + '&intQuoteId=<%=intQuoteId%>','EditRecord','width=700,height=160,left='+GetCenterX( 700 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 700 )+',screenY='+GetCenterY( 160 )+',scrollbars=yes,dependent=yes' );
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
		
	function OpenProductSearch(  )
	{
	  OpenWindow( 'product_search.asp','ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectProduct( intProductId, strProductName, strPartNumber, strUnitPrice )
	{
		document.forms.frmParts.hidProductId.value = intProductId;
		document.forms.frmParts.txtProductName.value = strPartNumber;
		document.forms.frmParts.txtCost.value = strUnitPrice;
		document.forms.frmParts.txtMarkup.value = 1.5;
		document.forms.frmParts.txtQty.value = 1;
	}
		
	function OpenSupplierSearch(  )
	{
	  OpenWindow( 'supplier_search.asp','SupplierSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectSupplier( intSupplierId, strSupplierName )
	{
		document.forms.frmParts.hidSupplierId.value = intSupplierId;
		document.forms.frmParts.txtSupplierName.value = strSupplierName;
	}
  </script>
</head>
<body topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=660 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetQuotePartsByQuoteId is nothing ) then
		if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
			if( not rsGetQuotePartsByQuoteId.EOF ) then
				theSpindleId = rsGetQuotePartsByQuoteId("SpindleId")		
%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>Parts Cost</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>Quote #&nbsp;&nbsp;<%=intQuoteId%></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- Cost Form -->
<form id="frmPartsCost" name="frmPartsCost" action="" method="post">
<input type="hidden" name="txtQuoteIdMaster" value="<%=intQuoteId%>">
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		<tr><td class="label_1">Spindle Type</td><td class="input_2" colspan=3><input type="text" class="forminput" name="txtSpindleType" maxlength=255 readonly value="<%=rsGetQuotePartsByQuoteId( "SpindleType" )%>"></td></tr>
		<tr><td class="label_1">Freight Charge</td><td class="input_1"><input type="text" class="forminput" name="txtFreightCharge" onchange="this.isDirty=true;" maxlength=10 isCurrency="true" fieldName="Freight Charge" value="<%=FormatCurrency( rsGetQuotePartsByQuoteId( "FreightChargeParts" ), 2 )%>"></td>
		<!--<td class="label_2">Commission</td><td class="input_2"><input type="text" class="forminput" name="txtPartsCommission" maxlength=10 value="<%=rsGetQuotePartsByQuoteId( "PartsCommission" )%>"></td></tr>-->
		<tr><td class="label_1">Total Charge</td><td class="input_1"><input type="text" class="forminput" name="txtTotalCharge" readonly value="<%=FormatCurrency( rsGetQuotePartsByQuoteId( "TotalPartCharge" ), 2 )%>"></td>
		<td class="label_2">Total Cost</td><td class="input_2"><input type="text" class="forminput" name="txtTotalCost" readonly value="<%=FormatCurrency( rsGetQuotePartsByQuoteId( "TotalPartCost" ), 2 )%>"></td></tr>
	</table>
</td></tr> 

<!-- Form Bottom Rule -->
<tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>
 <!-- End Form Bottom Rule -->

<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=320 align=right>
	          <tr><td width=64 align=right><a href="javascript:window.navigate( '<%=Request.QueryString( "strPrevPage" )%>' );"><IMG src="images/back_button.gif" width=60 height=20 border=0></a></td>
							<td width=64 align=right><a href="javascript:validateForm(frmPartsCost);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmPartsCost.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:OpenPrintBox('quote_partsorder.asp?intQuoteId=<%=intQuoteId%>');"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmPartsCost);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
	 <p style="font-size:10px;color:red;text-align:left;padding-left:34px;">28% base is added to "Total Charge"</p>
</td></tr>
<!-- End Form Buttons -->
</form>
<%	set rsGetQuotePartsByQuoteId = rsGetQuotePartsByQuoteId.NextRecordset
		end if
	end if
end if%>
<!-- End PO Form -->

<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Parts</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmParts" name="frmParts" action="" method="post">
 <input type="hidden" name="txtQuoteIdDetail" value="<%=intQuoteId%>">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="txtProductId" value="">
 <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=660 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=660 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr style="font-size:.8em">
    <th class="itemlist" width="15" align="center">Del</th>
    <th class="itemlist" width="15" align="center">Edit</th>
    <th class="itemlist" width="25" align="center">Id</th>
    <th class="itemlist" width="110">Part #</th>
    <th class="itemlist" width="110">Product Name</th>
    <th class="itemlist" width="120">Supplier</th>
    <th class="itemlist" width="70">Cost</th>
    <th class="itemlist" width="50">Markup</th>
    <th class="itemlist" width="30">Qty</th>
    <th class="itemlist" width="30">OH</th>
    <th class="itemlist" width="30" style="color:green;" title="Quoted as well as on Order">OQ/O</th>
  </tr>  
 <!-- table data -->
	<%if not( rsGetQuotePartsByQuoteId is nothing ) then
		if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
			intRowNumber = 1
			
			
			while( not rsGetQuotePartsByQuoteId.EOF )					
				if( ( intRowNumber Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
				<%else%>
					<tr class="oddrow">
				<%end if%>
				
				<td class="itemlist" width="15" align=middle><a href="javascript:ConfirmDelete('<%=rsGetQuotePartsByQuoteId( "QuotePartId" )%>', '<%=EscapeQuotes(rsGetQuotePartsByQuoteId( "PartNumber" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
				<td class="itemlist" width="60" align=middle><a href="javascript:EditQuotePart('<%=rsGetQuotePartsByQuoteId( "QuotePartId" )%>');"><img src="images/form_edit.gif" width=15 height=15 border=0></a></td>
				<td class="itemlist" width="25" align="center"><a href="javascript:OpenProduct( '<%=rsGetQuotePartsByQuoteId( "ProductId" )%>' );"><%=rsGetQuotePartsByQuoteId( "ProductId" )%></a></td>
				<td class="itemlist" width="110"><%=rsGetQuotePartsByQuoteId( "PartNumber" )%></td>
				<td class="itemlist" width="110"><%=rsGetQuotePartsByQuoteId( "ProductName" )%></td>
				<td class="itemlist" width="120"><%=EscapeDoubleQuotes(rsGetQuotePartsByQuoteId( "SupplierName" ))%></td>
				<td class="itemlist" width="70" align="right"><%=FormatCurrency( rsGetQuotePartsByQuoteId( "PartCost" ), 2 )%></td>
				<td class="itemlist" width="50" align="right"><%=rsGetQuotePartsByQuoteId( "Markup" )%></td>
				<td class="itemlist" width="30" align="right"><%=rsGetQuotePartsByQuoteId( "Qty" )%></td>
				<td class="itemlist" width="30" align="right"><%=rsGetQuotePartsByQuoteId( "OnHand" )%></td>
				<td class="itemlist" width="30" align="right"><%=rsGetQuotePartsByQuoteId( "OnOrder" )%></td>
				</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetQuotePartsByQuoteId.MoveNext
			wend
		end if
	end if%>
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td width="15">&nbsp;</td>
			<td width="60">&nbsp;</td>
			<td class="itemlist" width="25" align="center"><a name="txtProductLink" href=""><div id="txtProductIdDiv"></div></a></td>
			<td colspan="2"><table cellpadding=0 cellspacing=0 width="220"><tr><td><a href="javascript:OpenProductSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidProductId" value=""><input class="forminput" name="txtProductName" readonly isRequired="true" fieldName="Product" value=""></td></tr></table></td>
			<td><table cellpadding=0 cellspacing=0 width="120"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value=""><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value=""></td></tr></table></td>
			<td width="70"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtMarkup" isRequired="true" isNumber="true" fieldName="Markup" value=""></td>
			<td><input type="text" class="forminput" name="txtQty" isRequired="true" isQty="true" fieldName="Qty" value=""></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
 </table></td>
 </tr>
	<tr><table width="660"><tr>
	
			<%if not (IsNull(theSpindleId)) then%>
				<td class="itemlist" width="408"><a href="javascript:OpenUsedPartsPopup(<%=theSpindleId%>, -1);">Select default parts</a></td>
			<%end if%>	
	<td><table cellspacing=0 cellpadding=0 width=192 align=right>
		<tr>
		  <td align=right width=64><a href="javascript:validateForm(frmParts);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:OpenPrintBox('quote_partsorder.asp?intQuoteId=<%=intQuoteId%>');"><img src="images/print_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	</table></td></tr>
	
	
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
</body>
</html>
<%
if not( rsGetQuotePartsByQuoteId is nothing ) then
	if( rsGetQuotePartsByQuoteId.State = adStateOpen ) then
		rsGetQuotePartsByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuotePartsByQuoteId					= nothing
set cmdGetQuotePartsByQuoteId					= nothing
set cmdEditQuote											= nothing
set cnnSQLConnection									= nothing
%>
