<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetProductByProductId, cmdGetProductByProductId
dim rsGetCategories, cmdGetCategories
dim rsGetProductOwners, cmdGetProductOwners
dim cmdEditProduct
dim intRowNumber, intProductId

set cnnSQLConnection						= nothing
set rsGetProductByProductId			= nothing
set cmdGetProductByProductId		= nothing
set rsGetCategories							= nothing
set cmdGetCategories						= nothing
set rsGetProductOwners					= nothing
set cmdGetProductOwners					= nothing
set cmdEditProduct							= nothing

set cnnSQLConnection						= Server.CreateObject( "ADODB.Connection" )
set cmdGetProductByProductId		= Server.CreateObject( "ADODB.Command" )
set cmdGetCategories						= Server.CreateObject( "ADODB.Command" )
set cmdGetProductOwners					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then
	set cmdEditProduct	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditProduct.ActiveConnection	= cnnSQLConnection
	cmdEditProduct.CommandText			= "{? = call spEditProduct( ?,?,?,?,?,?,?,?,?,?,?,? ) }"

	
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ProductID", adInteger, adParamInput, 4 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@PartNumber", adVarChar, adParamInput, 100 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ProductName", adLongVarChar, adParamInput, 1048576 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ProductDescription", adLongVarChar, adParamInput, 1048576 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@CategoryID", adInteger, adParamInput, 4 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@SerialNumber", adVarChar, adParamInput, 100 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@UnitPrice", adVarChar, adParamInput, 10 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ReorderLevel", adInteger, adParamInput, 4 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@LeadTime", adVarChar, adParamInput, 100 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@DrawingID", adVarChar, adParamInput, 100 )
	cmdEditProduct.Parameters.Append cmdEditProduct.CreateParameter( "@ProductOwnerID", adInteger, adParamInput, 4 )
	
	if( Request.Form( "txtProductId" ) <> "" ) then cmdEditProduct.Parameters( "@ProductID" ) = Request.Form( "txtProductId" )
	if( Request.Form( "txtPartNumber" ) <> "" ) then cmdEditProduct.Parameters( "@PartNumber" ) = Request.Form( "txtPartNumber" )
	if( Request.Form( "txtProductName" ) <> "" ) then cmdEditProduct.Parameters( "@ProductName" ) = Request.Form( "txtProductName" )
	if( Request.Form( "txtDescription" ) <> "" ) then cmdEditProduct.Parameters( "@ProductDescription" ) = Request.Form( "txtDescription" )
	if( Request.Form( "selCategory" ) <> "" ) then cmdEditProduct.Parameters( "@CategoryID" ) = Request.Form( "selCategory" )
	if( Request.Form( "txtSerialNumber" ) <> "" ) then cmdEditProduct.Parameters( "@SerialNumber" ) = Request.Form( "txtSerialNumber" )
	
	if( Left( Request.Form( "txtUnitPrice" ), 1 ) = "$" ) then
		if( Request.Form( "txtUnitPrice" ) <> "" ) then cmdEditProduct.Parameters( "@UnitPrice" ) = Mid( Request.Form( "txtUnitPrice" ), 2 )
	else
		if( Request.Form( "txtUnitPrice" ) <> "" ) then cmdEditProduct.Parameters( "@UnitPrice" ) = Request.Form( "txtUnitPrice" )
	end if
	
	if( Request.Form( "txtReorderLvl" ) <> "" ) then cmdEditProduct.Parameters( "@ReorderLevel" ) = Request.Form( "txtReorderLvl" )
	if( Request.Form( "txtLeadTime" ) <> "" ) then cmdEditProduct.Parameters( "@LeadTime" ) = Request.Form( "txtLeadTime" )
	if( Request.Form( "txtDrawingNumber" ) <> "" ) then cmdEditProduct.Parameters( "@DrawingID" ) = Request.Form( "txtDrawingNumber" )
	if( Request.Form( "selProductOwner" ) <> "" ) then cmdEditProduct.Parameters( "@ProductOwnerID" ) = Request.Form( "selProductOwner" )
		
	on error resume next
	cmdEditProduct.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
						
	on error goto 0
	
	intProductId = cmdEditProduct.Parameters( "@Return" ).Value
	
else
	if( Request.QueryString( "intProductId" ) = "" ) then
		intProductId = 0
	else
		intProductId = Request.QueryString( "intProductId" )
	end if
end if

cmdGetProductByProductId.ActiveConnection = cnnSQLConnection
cmdGetProductByProductId.CommandText = "{call spGetProductByProductId( ?,? ) }"
cmdGetProductByProductId.Parameters.Append cmdGetProductByProductId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProductByProductId.Parameters.Append cmdGetProductByProductId.CreateParameter( "@ProductId", adInteger, adParamInput, 4 )
cmdGetProductByProductId.Parameters( "@ProductId" )	= intProductId

on error resume next
set rsGetProductByProductId = cmdGetProductByProductId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
						
on error goto 0
'**********************************************************************************
cmdGetCategories.ActiveConnection = cnnSQLConnection
cmdGetCategories.CommandText = "{call spGetCategoryList}"

on error resume next
set rsGetCategories = cmdGetCategories.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
						
on error goto 0
'**********************************************************************************
cmdGetProductOwners.ActiveConnection = cnnSQLConnection
cmdGetProductOwners.CommandText = "{call spGetProductOwnerList}"

on error resume next
set rsGetProductOwners = cmdGetProductOwners.Execute
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
<TITLE>Centerline - View / Edit Product</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
		<!-- Hide script from old browsers
		function OpenPODetail( strPOId )
		{
		  OpenWindow( 'po_details.asp?intPurchaseOrderId='+strPOId,'CenterlinePO'+strPOId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
		}
		// End hiding script from old browsers -->
	</script>
</HEAD>
<BODY topmargin=0 leftmargin=0>
<script type="text/javascript" src="/js/tooltips.js"></script>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetProductByProductId is nothing ) then
		if( rsGetProductByProductId.State = adStateOpen ) then
			if( not rsGetProductByProductId.EOF ) then%>
 <!--Top Title Bar -->
 <tr><td class="pagetitle">  
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
		<%if( intProductId = 0 ) then%>
		<tr><td class="sectiontitle" width=20%>New Product</td>
		<%else%>
    <tr><td class="sectiontitle" width=20%>Product&nbsp;&nbsp;<%=rsGetProductByProductId( "ProductId" )%></td>
    <%end if%>
      <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
      <td align=right>&nbsp;</td></tr>
  </TABLE>
</td></tr>  <!-- End Top Title Bar -->

 <!-- PO Form -->
 <tr><td align=center>
  <form id="frmProduct" name="frmProduct" action="" method="post">
  <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
    <tr><td class="label_1">Product ID</td><td class="input_1"><input type="text" class="forminput" name="txtProductId" readonly maxlength=32 value="<%=intProductId%>"></td>
      <td class="label_2">Drawing #</td><td class="input_2"><input type="text" class="forminput" name="txtDrawingNumber" onchange="this.isDirty=true;" fieldName="Drawing #" maxlength=32 value="<%=rsGetProductByProductId( "DrawingId" )%>"></td><tr>
    <tr><td class="label_1">Part #</td><td class="input_1"><input type="text" class="forminput" name="txtPartNumber" maxlength=50 onchange="this.isDirty=true;" isRequired="true" fieldName="Part Number" value="<%=EscapeDoubleQuotes(rsGetProductByProductId( "PartNumber" ))%>"></td>
      <td class="label_2">Lead Time</td><td class="input_2"><input type="text" class="forminput" name="txtLeadTime" onchange="this.isDirty=true;" fieldName="Lead Time" maxlength=30 value="<%=rsGetProductByProductId( "LeadTime" )%>"></td></tr>
    <tr><td class="label_1"><span style="color:green" onmouseover="Tip('Manufacturer\'s List Price, Currency Denominator and Date of Pricing')">List Price</span></td><td class="input_1"><input type="text" class="forminput" name="txtSerialNumber" onchange="this.isDirty=true;" fieldName="Serial Number" maxlength=50 value="<%=rsGetProductByProductId( "SerialNumber" )%>"></td>
      <td class="label_2"><span style="color:green" onmouseover="Tip('Our cost at the latest time of purchase, update as often as necessary')">CEN Cost</span></td><td class="input_2"><input type="text" class="forminput" name="txtUnitPrice" maxlength=10 onchange="this.isDirty=true;" isCurrency="true" fieldName="Unit Price" value="<%=FormatCurrency( rsGetProductByProductId( "UnitPrice" ), 2 )%>"></td></tr>
    <tr><td colspan=2>&nbsp;</td>
   	  <td class="label_2">Reorder Lvl</td><td class="input_2"><input type="text" class="forminput" name="txtReorderLvl" maxlength=10 isNumber="true" onchange="this.isDirty=true;" fieldName="Reorder Lvl" value="<%=rsGetProductByProductId( "ReorderLevel" )%>"></td></tr>
   	<tr><td class="label_1">Category</td><td class="input_1">
   	  <select class="forminput" name="selCategory" size=1 onchange="this.isDirty=true;" fieldName="Category" ID="Select1">
   			<%if not( rsGetCategories is nothing ) then
   					if( rsGetCategories.State = adStateOpen ) then
   						while( not rsGetCategories.EOF )%>
   							<option value="<%=rsGetCategories( "CategoryId" )%>"<%if( rsGetCategories( "CategoryName" ) = rsGetProductByProductId( "CategoryName" ) ) then%>SELECTED<%end if%>><%=rsGetCategories( "CategoryName" )%></option>
   							<%rsGetCategories.MoveNext
   						wend
   					end if
   				end if
   			%>
   	  </select></td>
   	  <td class="label_2">On Hand</td><td class="input_2"><input type="text" class="forminput" name="txtOnHand" maxlength=32 readonly value="<%=rsGetProductByProductId( "OnHand" )%>"></td></tr>
   	<tr><td class="label_1">Owner</td><td class="input_1">
   	  <select class="forminput" name="selProductOwner" size=1 onchange="this.isDirty=true;" fieldName="Category">
   			<%if not( rsGetProductOwners is nothing ) then
   					if( rsGetProductOwners.State = adStateOpen ) then
   						while( not rsGetProductOwners.EOF )%>
   							<option value="<%=rsGetProductOwners( "ProductOwnerId" )%>"<%if( rsGetProductOwners( "ProductOwner" ) = rsGetProductByProductId( "ProductOwner" ) ) then%>SELECTED<%end if%>><%=rsGetProductOwners( "ProductOwner" )%></option>
   							<%rsGetProductOwners.MoveNext
   						wend
   					end if
   				end if
   			%>
   	  </select></td>
			<td class="label_2">On Order</td><td class="input_2"><input type="text" class="forminput" name="txtOnOrder" readonly maxlength=32 value="<%=rsGetProductByProductId( "OnOrder" )%>"></td></tr>
	 <tr><td class="label_1">Product Name</td><td class="input_1" colspan=3><textarea class="forminput" name="txtProductName" onchange="this.isDirty=true;" fieldName="Product Name" rows=6 wrap=soft><%=rsGetProductByProductId( "ProductName" )%></textarea></td></tr>
	 <tr><td class="label_1">Detailed Description</td><td class="input_1" colspan=3><textarea class="forminput" name="txtDescription" onchange="this.isDirty=true;" fieldName="Detailed Description" rows=6 wrap=soft><%=rsGetProductByProductId( "ProductDescription" )%></textarea></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmProduct);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmProduct.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmProduct);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
  </form>
 </td></tr>
 <!-- End PO Form -->
 <%		set rsGetProductByProductId = rsGetProductByProductId.NextRecordset
		end if
	end if
 end if%>
 
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=4></td></tr>
    <tr><td class="smallsectiontitle" width=15%>Transactions</td>
      <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
 <!-- Transactions Table -->
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist">Date</th>
    <th class="itemlist">PO ID</th>
    <th class="itemlist">Description</th>
    <th class="itemlist"># Ordered</th>
    <th class="itemlist">Date Received</th>
    <th class="itemlist"># Recvd</th>
    <th class="itemlist"># Sold</th>
  </tr>  
  
  <!-- table data -->
  <%if( rsGetProductByProductId.State = adStateOpen ) then
		intRowNumber = 1
		
  while( not rsGetProductByProductId.EOF )
		if( ( intRowNumber Mod 2 ) = 1 ) then%>
			<tr class="evenrow">
			<%else%>
			<tr class="oddrow">
			<%end if%>
		<%if( not IsNull( rsGetProductByProductId( "TransactionDate" ) ) ) then%>
			<td class="itemlist" align=middle><%=FormatDateTime( rsGetProductByProductId( "TransactionDate" ), vbShortDate )%></td>
		<%else%>
			<td class="itemlist" align=middle><%=rsGetProductByProductId( "TransactionDate" )%></td>
		<%end if%>
		<td class="itemlist" align="center"><a href="javascript:OpenPODetail( '<%=rsGetProductByProductId( "PurchaseOrderId" )%>' );"><%=rsGetProductByProductId( "PurchaseOrderId" )%></a></td>
    <td class="itemlist"><%=rsGetProductByProductId( "TransactionDescription" )%></td>
    <td class="itemlist" align=right><%=rsGetProductByProductId( "UnitsOrdered" )%></td>
    <%if( not IsNull( rsGetProductByProductId( "DateReceived" ) ) ) then%>
			<td class="itemlist" align=middle><%=FormatDateTime( rsGetProductByProductId( "DateReceived" ), vbShortDate )%></td>
		<%else%>
			<td class="itemlist" align=middle><%=rsGetProductByProductId( "DateReceived" )%></td>
		<%end if%>
    <td class="itemlist" align=right><%=rsGetProductByProductId( "UnitsReceived" )%></td>
    <td class="itemlist" align=right><%=rsGetProductByProductId( "UnitsSold" )%></td>
  </tr>
  <%
  intRowNumber = intRowNumber + 1
  rsGetProductByProductId.MoveNext
	wend
	end if%>  
 </table></td></tr>

  </table>
 
 </td></tr>  <!-- End Middle Title Bar -->

</BODY>
</HTML>
<%
if not( rsGetProductByProductId is nothing ) then
	if( rsGetProductByProductId.State = adStateOpen ) then
		rsGetProductByProductId.Close
	end if
end if

if not( rsGetCategories is nothing ) then
	if( rsGetCategories.State = adStateOpen ) then
		rsGetCategories.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProductByProductId		= nothing
set rsGetCategories						= nothing
set cmdGetProductByProductId	= nothing
set cmdGetCategories					= nothing
set cmdEditProduct						= nothing
set cnnSQLConnection					= nothing
%>
