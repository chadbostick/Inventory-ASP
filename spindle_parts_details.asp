<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetSpindleDetailsBySpindleId, cmdGetSpindleDetailsBySpindleId
dim cmdEditSpindleProduct, cmdCheckSpindleProduct, cmdCheckSpindleSubWorkDesc, cmdEditSpindleSubWork
dim intQuoteId, intRowNumber, strCurrentPage
dim theSpindleId, theProductId, retVal, theSubWorkDesc

set cnnSQLConnection							= nothing
set rsGetSpindleDetailsBySpindleId				= nothing
set cmdGetSpindleDetailsBySpindleId				= nothing
set cmdEditSpindleProduct						= nothing
set cmdCheckSpindleProduct						= nothing
set cmdCheckSpindleSubWorkDesc					= nothing
set cmdEditSpindleSubWork						= nothing


set cnnSQLConnection							= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleDetailsBySpindleId				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	if Request.Form("productType") <> "subwork" then	
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
		
		
	elseif Request.Form("productType") = "subwork" then
	
		'get the spindle and product to add
		theSpindleId = Request.Form( "theSpindleId" )
		theSubWorkDesc = Request.Form( "txtSubWorkDescription" )
		
		'if the spindle and product are already associated in the DB then don't add them this time
		set cmdCheckSpindleSubWorkDesc	= Server.CreateObject( "ADODB.Command" )
	
		cmdCheckSpindleSubWorkDesc.ActiveConnection	= cnnSQLConnection
		cmdCheckSpindleSubWorkDesc.CommandText			= "{? = call spCheckSpindleSubWorkDesc( ?,?,? ) }"
		
		cmdCheckSpindleSubWorkDesc.Parameters.Append cmdCheckSpindleSubWorkDesc.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdCheckSpindleSubWorkDesc.Parameters.Append cmdCheckSpindleSubWorkDesc.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdCheckSpindleSubWorkDesc.Parameters.Append cmdCheckSpindleSubWorkDesc.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
		cmdCheckSpindleSubWorkDesc.Parameters.Append cmdCheckSpindleSubWorkDesc.CreateParameter( "@SubWorkDesc", adVarChar, adParamInput, 255 )
		
		cmdCheckSpindleSubWorkDesc.Parameters( "@SpindleId" ) = theSpindleId
		cmdCheckSpindleSubWorkDesc.Parameters( "@SubWorkDesc" ) = theSubWorkDesc
		
		on error resume next
		cmdCheckSpindleSubWorkDesc.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
		on error goto 0
		
		'Find out if the part is already associated with our spindle
		retVal = cmdCheckSpindleSubWorkDesc.Parameters( "@Return" ).Value
		
		'if the part isn't associated, then add it to the association table
		if (retVal = 0) then
			set cmdEditSpindleSubWork	= Server.CreateObject( "ADODB.Command" )
	
			cmdEditSpindleSubWork.ActiveConnection	= cnnSQLConnection
			cmdEditSpindleSubWork.CommandText			= "{? = call spEditSpindleSubWork( ?,?,?,?,?,? ) }"
		
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SpindleSubWorkId", adInteger, adParamInput, 4 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SubWorkCost", adVarChar, adParamInput, 10 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SubWorkDesc", adVarChar, adParamInput, 255 )
		
		
		
			cmdEditSpindleSubWork.Parameters( "@SpindleSubWorkId" ) = 0
			cmdEditSpindleSubWork.Parameters( "@SpindleId" ) = Request.Form( "theSpindleId" )
			
			'cmdEditSpindleSubWork.Parameters( "@SubWorkCost" ) = Request.Form("txtCost")
			if( Left( Request.Form( "txtCost" ), 1 ) = "$" ) then
				cmdEditSpindleSubWork.Parameters( "@SubWorkCost" ) = Mid( Request.Form( "txtCost" ), 2 )
			else
				cmdEditSpindleSubWork.Parameters( "@SubWorkCost" ) = Request.Form( "txtCost" )
			end if
		
			cmdEditSpindleSubWork.Parameters( "@SupplierId" ) = Request.Form( "hidSupplierId" )
			cmdEditSpindleSubWork.Parameters( "@SubWorkDesc" ) = Request.Form("txtSubWorkDescription")
		
		
			on error resume next
			cmdEditSpindleSubWork.Execute
			if( Err.number <> 0 ) then
				Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
				Response.Write( Err.Description )
				Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
				Response.Write( "</font></body></html>" )
				Response.End
			end if

			on error goto 0
		end if	
	end if
		
	
else
	theSpindleId = Request.QueryString( "intSpindleId" )
end if



cmdGetSpindleDetailsBySpindleId.ActiveConnection = cnnSQLConnection
cmdGetSpindleDetailsBySpindleId.CommandText = "{call spGetSpindleDetailsBySpindleId( ?,? ) }"

cmdGetSpindleDetailsBySpindleId.Parameters.Append cmdGetSpindleDetailsBySpindleId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleDetailsBySpindleId.Parameters.Append cmdGetSpindleDetailsBySpindleId.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )

cmdGetSpindleDetailsBySpindleId.Parameters( "@SpindleId" )	= theSpindleId

on error resume next
set rsGetSpindleDetailsBySpindleId = cmdGetSpindleDetailsBySpindleId.Execute
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
  var productType;
  
	function ConfirmDelete( intQuotePartId, strPart )
	{
		if( confirm( 'Are you sure you want to remove item ' + strPart + '?' ) )
		{
			OpenWindow( 'quoteparts_delete.asp?intQuotePartId='+intQuotePartId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function ConfirmPartBearingDelete( intSpindleProductId, strPart )
	{
		if( confirm( 'Are you sure you want to remove item ' + strPart + '?' ) )
		{
			OpenWindow( 'spindle_parts_bearings_delete.asp?intSpindleProductId='+intSpindleProductId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function ConfirmSubworkDelete( intSpindleSubWorkId, strPart )
	{
		if( confirm( 'Are you sure you want to remove item ' + strPart + '?' ) )
		{
			OpenWindow( 'spindle_subwork_delete.asp?intSpindleSubWorkId='+intSpindleSubWorkId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function EditPartBearing( intSpindleProductId)
	{
		OpenWindow( 'spindle_parts_bearings_edit.asp?intSpindleProductId='+intSpindleProductId + '&intSpindleId=<%=theSpindleId%>','EditRecord','width=700,height=160,left='+GetCenterX( 700 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 700 )+',screenY='+GetCenterY( 160 )+',scrollbars=yes,dependent=yes' );
	}
	
	function EditSubWork( intSpindleSubWorkId)
	{
		OpenWindow( 'spindle_subwork_edit.asp?intSpindleSubWorkId='+intSpindleSubWorkId + '&intSpindleId=<%=theSpindleId%>','EditRecord','width=700,height=160,left='+GetCenterX( 700 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 700 )+',screenY='+GetCenterY( 160 )+',scrollbars=yes,dependent=yes' );
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
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
		
	function OpenProductSearch(  )
	{
		productType = "part";
	  OpenWindow( 'product_search.asp','ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenBearingProductSearch(  )
	{
		productType = "bearing";
	  OpenWindow( 'product_search.asp','ProductSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function SelectProduct( intProductId, strProductName )
	{
		if (productType == "part")
		{
			document.forms.frmParts.hidProductId.value = intProductId;
			document.forms.frmParts.txtProductName.value = strProductName;
		}
		
		if (productType == "bearing")
		{
			document.forms.frmBearings.hidProductId.value = intProductId;
			document.forms.frmBearings.txtProductName.value = strProductName;
		}
	}
		
	function OpenSupplierSearch(  )
	{
		productType = "part";
	  OpenWindow( 'supplier_search.asp','SupplierSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenBearingSupplierSearch(  )
	{
		productType = "bearing";
	  OpenWindow( 'supplier_search.asp','SupplierSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenSubWorkSupplierSearch(  )
	{
		productType = "subwork";
	  OpenWindow( 'supplier_search.asp','SupplierSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	
	function SelectSupplier( intSupplierId, strSupplierName )
	{
		if (productType == "part")
		{
			document.forms.frmParts.hidSupplierId.value = intSupplierId;
			document.forms.frmParts.txtSupplierName.value = strSupplierName;
		}
		
		if (productType == "bearing")
		{
			document.forms.frmBearings.hidSupplierId.value = intSupplierId;
			document.forms.frmBearings.txtSupplierName.value = strSupplierName;
		}
		
		if (productType == "subwork")
		{
			document.forms.frmSubWorkDetail.hidSupplierId.value = intSupplierId;
			document.forms.frmSubWorkDetail.txtSupplierName.value = strSupplierName;
		}
	}
  </script>
</head>
<body topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>


<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
		<%if( theSpindleId = 0 ) then%>
		<tr><td class="sectiontitle" width=20%>New Spindle</td>
		<%else%>
		<tr><td class="sectiontitle" width=20%>Spindle&nbsp;&nbsp;#<%=theSpindleId%></td>
		<%end if%>
     <td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
 </table>
</td></tr>
<!-- End Top Title Bar -->


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
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="txtProductId" value="">
 <input type="hidden" name="productType" value="part">
 <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=660 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=660 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Del</th>
    <th class="itemlist" width="60" align="center">Edit</th>
    <th class="itemlist" width="25" align="center">Id</th>
    <th class="itemlist" width="250">Product Description</th>
    <th class="itemlist" width="150">Supplier</th>
    <th class="itemlist" width="70">Cost</th>
    <th class="itemlist" width="60">Markup</th>
    <th class="itemlist" width="30">Qty</th>
  </tr>  
 <!-- table data -->
	<%if not( rsGetSpindleDetailsBySpindleId is nothing ) then
		if( rsGetSpindleDetailsBySpindleId.State = adStateOpen ) then
			intRowNumber = 1
			
			'rsGetSpindleDetailsBySpindleId.MoveFirst
			while( not rsGetSpindleDetailsBySpindleId.EOF )					
				if( ( intRowNumber Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
				<%else%>
					<tr class="oddrow">
				<%end if%>
				
				<td class="itemlist" width="15" align=middle><a href="javascript:ConfirmPartBearingDelete('<%=rsGetSpindleDetailsBySpindleId( "SpindleProductId" )%>', '<%=rsGetSpindleDetailsBySpindleId( "ProductDescription" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
				<td class="itemlist" width="60" align=middle><a href="javascript:EditPartBearing('<%=rsGetSpindleDetailsBySpindleId( "SpindleProductId" )%>');"><img src="images/edit_button.gif" width=60 height=20 border=0></a></td>
				<td class="itemlist" width="25" align="center"><a href="javascript:OpenProduct( '<%=rsGetSpindleDetailsBySpindleId( "ProductId" )%>' );"><%=rsGetSpindleDetailsBySpindleId( "ProductId" )%></a></td>
				<td class="itemlist" width="250"><%=rsGetSpindleDetailsBySpindleId( "ProductDescription" )%></td>
				<td class="itemlist" width="150"><%=rsGetSpindleDetailsBySpindleId( "SupplierName" )%></td>
				<td class="itemlist" width="70" align="right"><%=FormatCurrency( rsGetSpindleDetailsBySpindleId( "PartCost" ), 2 )%></td>
				<td class="itemlist" width="60" align="right"><%=rsGetSpindleDetailsBySpindleId( "Markup" )%></td>
				<td class="itemlist" width="30" align="right"><%=rsGetSpindleDetailsBySpindleId( "Qty" )%></td>
				</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetSpindleDetailsBySpindleId.MoveNext
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
			<td><table cellpadding=0 cellspacing=0 width="250"><tr><td><a href="javascript:OpenProductSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidProductId" value=""><input class="forminput" name="txtProductName" readonly isRequired="true" fieldName="Product" value=""></td></tr></table></td>
			<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value=""><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value=""></td></tr></table></td>
			<td width="70"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value=""></td>
			<td width="60"><input type="text" class="forminput" name="txtMarkup" isRequired="true" isNumber="true" fieldName="Markup" value=""></td>
			<td width="30"><input type="text" class="forminput" name="txtQty" isRequired="true" isQty="true" fieldName="Qty" value=""></td>
		</tr>
 </table></td>
 </tr>
	<tr><table width="660"><tr>
	
			
	<td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr>
		  <td align=right width=64><a href="javascript:validateForm(frmParts);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	</table></td></tr>
	
	
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
 
 
 
 
 
 <!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Bearings</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Bearings Table -->
 <form id="frmBearings" name="frmBearings" action="" method="post">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="txtProductId" value="">
 <input type="hidden" name="productType" value="bearing">
 <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=660 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=660 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Del</th>
    <th class="itemlist" width="60" align="center">Edit</th>
    <th class="itemlist" width="25" align="center">Id</th>
    <th class="itemlist" width="250">Bearing Description</th>
    <th class="itemlist" width="150">Source</th>
    <th class="itemlist" width="70">Cost</th>
    <th class="itemlist" width="50">Markup</th>
    <th class="itemlist" width="30">Qty</th>
  </tr>  
 <!-- table data -->
	<%set rsGetSpindleDetailsBySpindleId = rsGetSpindleDetailsBySpindleId.NextRecordset
		if not( rsGetSpindleDetailsBySpindleId is nothing ) then
			if( rsGetSpindleDetailsBySpindleId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetSpindleDetailsBySpindleId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td class="itemlist" width="15" align=middle><a href="javascript:ConfirmPartBearingDelete('<%=rsGetSpindleDetailsBySpindleId( "SpindleProductId" )%>', '<%=rsGetSpindleDetailsBySpindleId( "BearingDescription" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="60" align=middle><a href="javascript:EditPartBearing('<%=rsGetSpindleDetailsBySpindleId( "SpindleProductId" )%>');"><img src="images/edit_button.gif" width=60 height=20 border=0></a></td>
							<td class="itemlist" width="25" align="center"><a href="javascript:OpenProduct( '<%=rsGetSpindleDetailsBySpindleId( "ProductId" )%>' );"><%=rsGetSpindleDetailsBySpindleId( "ProductId" )%></a></td>
							<td class="itemlist" width="250"><%=rsGetSpindleDetailsBySpindleId( "BearingDescription" )%></td>
							<td class="itemlist" width="150"><%=rsGetSpindleDetailsBySpindleId( "SupplierName" )%></td>
							<td class="itemlist" width="70" align="right"><%=FormatCurrency( rsGetSpindleDetailsBySpindleId( "BearingCost" ), 2 )%></td>
							<td class="itemlist" width="50" align="right"><%=rsGetSpindleDetailsBySpindleId( "Markup" )%></td>
							<td class="itemlist" width="30" align="right"><%=rsGetSpindleDetailsBySpindleId( "Qty" )%></td>
						</tr>
						<%intRowNumber = intRowNumber + 1
						rsGetSpindleDetailsBySpindleId.MoveNext
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
			<td class="itemlist" width="25" align="center"><a name="txtBearingLink" href=""><div id="txtBearingIdDiv"></div></a></td>
			<td><table cellpadding=0 cellspacing=0 width="250"><tr><td><a href="javascript:OpenBearingProductSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidProductId" value=""><input class="forminput" name="txtProductName" readonly isRequired="true" fieldName="Product" value=""></td></tr></table></td>
			<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenBearingSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value=""><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value=""></td></tr></table></td>
			<td width="70"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value=""></td>
			<td width="50"><input type="text" class="forminput" name="txtMarkup" isRequired="true" isNumber="true" fieldName="Markup" value=""></td>
			<td width="30"><input type="text" class="forminput" name="txtQty" isRequired="true" isQty="true" fieldName="Qty" value=""></td>
		</tr>
 </table></td>
 </tr><tr><td><table width=660><tr>
 
	
	
	<tr><td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr><td align=right width=64><a href="javascript:validateForm(frmBearings);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	</table></td></tr>

 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
 
 
 
 
 
 <!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Sub Work</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmSubWorkDetail" name="frmSubWorkDetail" action="" method="post">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="productType" value="subwork">
  <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=660 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=660 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
	<th class="itemlist" width="15">Del</th>
	<th class="itemlist" width="60" align="center">Edit</th>
    <th class="itemlist" width="385">Sub Work Description</th>
    <th class="itemlist" width="150">Source</th>
    <th class="itemlist" width="50">Cost</th>
  </tr>  
 <!-- table data -->  
		<%set rsGetSpindleDetailsBySpindleId = rsGetSpindleDetailsBySpindleId.NextRecordset
		if not( rsGetSpindleDetailsBySpindleId is nothing ) then
			if( rsGetSpindleDetailsBySpindleId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetSpindleDetailsBySpindleId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td class="itemlist" width="15" align=center><a href="javascript:ConfirmSubworkDelete('<%=rsGetSpindleDetailsBySpindleId( "SpindleSubWorkId" )%>', '<%=rsGetSpindleDetailsBySpindleId( "SubWorkDescription" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="60" align=middle><a href="javascript:EditSubWork('<%=rsGetSpindleDetailsBySpindleId( "SpindleSubWorkId" )%>');"><img src="images/edit_button.gif" width=60 height=20 border=0></a></td>
							<td class="itemlist" width="385"><%=rsGetSpindleDetailsBySpindleId( "SubWorkDescription" )%></td>
							<td class="itemlist" width="150"><%=rsGetSpindleDetailsBySpindleId( "SupplierName" )%></td>
							<td class="itemlist" width="50"><%=FormatCurrency( rsGetSpindleDetailsBySpindleId( "SubWorkCost" ), 2 )%></td>
						</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetSpindleDetailsBySpindleId.MoveNext
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
			<td width="385"><input type="text" class="forminput" isRequired="true" fieldName="Sub Work" name="txtSubWorkDescription" value=""></td>
			<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSubWorkSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value=""><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value=""></td></tr></table></td>
			<td width="50"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value=""></td>
		</tr>
 </table></td>
 </tr>
	<tr><table width="660"><tr>
	
				
	<td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr>
			<td align=right width=64><a href="javascript:validateForm(frmSubWorkDetail);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	
	

 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>
 <!-- End Middle Title Bar -->
	
				
	<td><table cellspacing=0 cellpadding=0 width=64 align=right>
		<tr>
			<td align=right width=64><a href="javascript:window.print();"><img src="images/print_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
</body>
</html>
<%
if not( rsGetSpindleDetailsBySpindleId is nothing ) then
	if( rsGetSpindleDetailsBySpindleId.State = adStateOpen ) then
		rsGetSpindleDetailsBySpindleId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleDetailsBySpindleId				= nothing
set cmdGetSpindleDetailsBySpindleId				= nothing
set cmdEditSpindleProduct						= nothing
set cmdCheckSpindleProduct						= nothing
set cmdCheckSpindleSubWorkDesc					= nothing
set cmdEditSpindleSubWork						= nothing
set cnnSQLConnection							= nothing

%>
