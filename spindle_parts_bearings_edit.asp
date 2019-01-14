<%option explicit

dim cnnSQLConnection
dim rsGetSpindleProductDetailsBySpindleProductId, cmdGetSpindleProductDetailsBySpindleProductId
dim cmdEditSpindleProduct
dim intQuoteId, intRowNumber, strCurrentPage
dim theSpindleProductId, theProductId, retVal, theSubWorkDesc, windowFreshener, theSpindleId

set cnnSQLConnection											= nothing
set rsGetSpindleProductDetailsBySpindleProductId				= nothing
set cmdGetSpindleProductDetailsBySpindleProductId				= nothing
set cmdEditSpindleProduct										= nothing
windowFreshener													= false

set cnnSQLConnection											= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleProductDetailsBySpindleProductId				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	windowFreshener = true

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
		
		
		
	cmdEditSpindleProduct.Parameters( "@SpindleProductId" ) = Request.Form("theSpindleProductId")
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
		
	
else
	theSpindleProductId = Request.QueryString( "intSpindleProductId" )
	theSpindleId = Request.QueryString( "intSpindleId" )
end if



cmdGetSpindleProductDetailsBySpindleProductId.ActiveConnection = cnnSQLConnection
cmdGetSpindleProductDetailsBySpindleProductId.CommandText = "{call spGetSpindleProductDetailsBySpindleProductId( ?,? ) }"

cmdGetSpindleProductDetailsBySpindleProductId.Parameters.Append cmdGetSpindleProductDetailsBySpindleProductId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleProductDetailsBySpindleProductId.Parameters.Append cmdGetSpindleProductDetailsBySpindleProductId.CreateParameter( "@SpindleProductId", adInteger, adParamInput, 4 )

cmdGetSpindleProductDetailsBySpindleProductId.Parameters( "@SpindleProductId" )	= theSpindleProductId

on error resume next
set rsGetSpindleProductDetailsBySpindleProductId = cmdGetSpindleProductDetailsBySpindleProductId.Execute
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
	function CloseAndRefreshParent()
	{
		window.opener.RefreshCurrentPage();
		window.close();
	}
	
	function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'CenterlineProduct'+strProductId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenPartsPopup( )
	{
	  OpenWindow( 'quoteparts_details_popup.asp','SelectPart','width=660,height=200,left='+GetCenterX( 660 )+',top='+GetCenterY( 200 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 200 )+',scrollbars=no,dependent=yes' );
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
		
	function OpenProductSearch(  )
	{
	  OpenWindow( 'product_search.asp','ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function SelectProduct( intProductId, strProductName )
	{
		document.forms.frmParts.hidProductId.value = intProductId;
		document.forms.frmParts.txtProductName.value = strProductName;
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
<body topmargin=0 leftmargin=0 <%if (windowFreshener = true) then %> onload="javascript:CloseAndRefreshParent();" <%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>



<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Edit Product</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmParts" name="frmParts" action="" method="post">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="txtProductId" value="">
 <input type="hidden" name="theSpindleProductId" value="<%=theSpindleProductId%>">
 <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=585 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=585 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="25" align="center">Id</th>
    <th class="itemlist" width="250">Product Description</th>
    <th class="itemlist" width="150">Supplier</th>
    <th class="itemlist" width="70">Cost</th>
    <th class="itemlist" width="60">Markup</th>
    <th class="itemlist" width="30">Qty</th>
  </tr>  
 <!-- table data -->
	<%if not( rsGetSpindleProductDetailsBySpindleProductId is nothing ) then
		if( rsGetSpindleProductDetailsBySpindleProductId.State = adStateOpen ) then
			intRowNumber = 1
			
			
			while( not rsGetSpindleProductDetailsBySpindleProductId.EOF )					
				if( ( intRowNumber Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
				<%else%>
					<tr class="oddrow">
				<%end if%>
				
				<td class="itemlist" width="25" align="center"><a href="javascript:OpenProduct( '<%=rsGetSpindleProductDetailsBySpindleProductId( "ProductId" )%>' );"><%=rsGetSpindleProductDetailsBySpindleProductId( "ProductId" )%></a></td>
				<td><table cellpadding=0 cellspacing=0 width="250"><tr><td><a href="javascript:OpenProductSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidProductId" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "ProductId" )%>"><input class="forminput" name="txtProductName" readonly isRequired="true" fieldName="Product" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "ProductDescription" )%>"></td></tr></table></td>
				<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "SupplierId" )%>"><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "SupplierName" )%>"></td></tr></table></td>
				<td width="70"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value="<%=FormatCurrency( rsGetSpindleProductDetailsBySpindleProductId( "PartCost" ), 2 )%>"></td>
				<td width="60"><input type="text" class="forminput" name="txtMarkup" isRequired="true" isNumber="true" fieldName="Markup" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "Markup" )%>"></td>
				<td width="30"><input type="text" class="forminput" name="txtQty" isRequired="true" isQty="true" fieldName="Qty" value="<%=rsGetSpindleProductDetailsBySpindleProductId( "Qty" )%>"></td>
				</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetSpindleProductDetailsBySpindleProductId.MoveNext
			wend
		end if
	end if%>
		
			
 </table></td>
 </tr>
	<tr><table width="585"><tr>
	
			
	<td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr>
		  <td align=right width=64><a href="javascript:validateForm(frmParts);"><img src="images/save_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	</table></td></tr>
	
	
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
 
</body>
</html>
<%
if not( rsGetSpindleProductDetailsBySpindleProductId is nothing ) then
	if( rsGetSpindleProductDetailsBySpindleProductId.State = adStateOpen ) then
		rsGetSpindleProductDetailsBySpindleProductId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleProductDetailsBySpindleProductId		= nothing
set cmdGetSpindleProductDetailsBySpindleProductId		= nothing
set cmdEditSpindleProduct								= nothing
set cnnSQLConnection									= nothing

%>
