<%option explicit

dim cnnSQLConnection
dim rsGetQuoteSubWorkDetailsByQuoteSubWorkId, cmdGetQuoteSubWorkDetailsByQuoteSubWorkId
dim cmdEditQuote
dim intQuoteId, intRowNumber, strCurrentPage
dim theQuoteSubWorkId, theProductId, retVal, theSubWorkDesc, windowFreshener, theQuoteId

set cnnSQLConnection											= nothing
set rsGetQuoteSubWorkDetailsByQuoteSubWorkId					= nothing
set cmdGetQuoteSubWorkDetailsByQuoteSubWorkId					= nothing
set cmdEditQuote												= nothing
windowFreshener													= false

set cnnSQLConnection											= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteSubWorkDetailsByQuoteSubWorkId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	windowFreshener = true

	set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditQuote.ActiveConnection	= cnnSQLConnection
	cmdEditQuote.CommandText			= "{? = call spEditQuoteSubWorkDetail( ?,?,?,?,?,? ) }"
		
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteSubWorkId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkCost", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkDescription", adLongVarChar, adParamInput, 1048576 )
		
	if( Request.Form( "intQuoteId" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "intQuoteId" )
	cmdEditQuote.Parameters( "@QuoteSubWorkID" ) = Request.Form( "intQuoteSubWorkId" )
	
	'if( Request.Form( "txtCost" ) <> "" ) then cmdEditQuote.Parameters( "@SubWorkCost" ) = Request.Form( "txtCost" )
	if( Left( Request.Form( "txtCost" ), 1 ) = "$" ) then
		cmdEditQuote.Parameters( "@SubWorkCost" ) = Mid( Request.Form( "txtCost" ), 2 )
	else
		cmdEditQuote.Parameters( "@SubWorkCost" ) = Request.Form( "txtCost" )
	end if
		
	if( Request.Form( "hidSupplierId" ) <> "" ) then cmdEditQuote.Parameters( "@SupplierId" ) = Request.Form( "hidSupplierId" )
	if( Request.Form( "txtSubWorkDescription" ) <> "" ) then cmdEditQuote.Parameters( "@SubWorkDescription" ) = Request.Form( "txtSubWorkDescription" )
	
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
		
	
else
	theQuoteSubWorkId = Request.QueryString( "intQuoteSubWorkId" )
	theQuoteId = Request.QueryString( "intQuoteId" )
end if



cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.ActiveConnection = cnnSQLConnection
cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.CommandText = "{call spGetQuoteSubWorkDetailsByQuoteSubWorkId( ?,? ) }"
cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.Parameters.Append cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.Parameters.Append cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.CreateParameter( "@QuoteSubWorkId", adInteger, adParamInput, 4 )
cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.Parameters( "@QuoteSubWorkId" )	= theQuoteSubWorkId

on error resume next
set rsGetQuoteSubWorkDetailsByQuoteSubWorkId = cmdGetQuoteSubWorkDetailsByQuoteSubWorkId.Execute
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
<TITLE>Centerline - View / Edit Rework Cost</TITLE>
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
   <tr><td class="smallsectiontitle" width=15%>Edit Rework</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmParts" name="frmParts" action="" method="post">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="txtProductId" value="">
 <input type="hidden" name="intQuoteSubWorkId" value="<%=theQuoteSubWorkId%>">
 <input type="hidden" name="intQuoteId" value="<%=theQuoteId%>">
 <tr><td align=center>
  <table width=585 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=585 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="385">Rework Description</th>
    <th class="itemlist" width="150">Source</th>
    <th class="itemlist" width="50">Cost</th>
  </tr>  
 <!-- table data -->  
		<%if not( rsGetQuoteSubWorkDetailsByQuoteSubWorkId is nothing ) then
			if( rsGetQuoteSubWorkDetailsByQuoteSubWorkId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetQuoteSubWorkDetailsByQuoteSubWorkId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td width="385"><input type="text" class="forminput" isRequired="true" fieldName="Sub Work" name="txtSubWorkDescription" value="<%=rsGetQuoteSubWorkDetailsByQuoteSubWorkId( "SubWorkDescription" )%>"></td>
							<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value="<%=rsGetQuoteSubWorkDetailsByQuoteSubWorkId( "SupplierId" )%>"><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value="<%=rsGetQuoteSubWorkDetailsByQuoteSubWorkId( "SupplierName" )%>"></td></tr></table></td>
							<td width="50"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value="<%=FormatCurrency( rsGetQuoteSubWorkDetailsByQuoteSubWorkId( "SubWorkCost" ), 2 )%>"></td>
						</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetQuoteSubWorkDetailsByQuoteSubWorkId.MoveNext
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
if not( rsGetQuoteSubWorkDetailsByQuoteSubWorkId is nothing ) then
	if( rsGetQuoteSubWorkDetailsByQuoteSubWorkId.State = adStateOpen ) then
		rsGetQuoteSubWorkDetailsByQuoteSubWorkId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuoteSubWorkDetailsByQuoteSubWorkId			= nothing
set cmdGetQuoteSubWorkDetailsByQuoteSubWorkId			= nothing
set cmdEditQuote										= nothing
set cnnSQLConnection									= nothing

%>
