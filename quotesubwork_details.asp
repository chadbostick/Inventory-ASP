<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetQuoteSubWorkByQuoteId, cmdGetQuoteSubWorkByQuoteId
dim cmdEditQuote, cmdEditSpindleSubWork, cmdCheckSpindleSubWorkDesc
dim intQuoteId, intRowNumber, strCurrentPage
dim theSpindleId, theProductId, retVal, theSupplierId, theSubWorkDesc

set cnnSQLConnection							= nothing
set rsGetQuoteSubWorkByQuoteId					= nothing
set cmdGetQuoteSubWorkByQuoteId					= nothing
set cmdEditQuote								= nothing
set cmdEditSpindleSubWork						= nothing
set cmdCheckSpindleSubWorkDesc					= nothing



set cnnSQLConnection							= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteSubWorkByQuoteId					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then
	if( Request.Form( "txtQuoteIdMaster" ) <> "" ) then
	
		set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQuote.ActiveConnection	= cnnSQLConnection
		cmdEditQuote.CommandText			= "{? = call spEditQuoteSubWork( ?,?,?,?,?,? ) }"
		
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteID", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@FreightLBS", adVarChar, adParamInput, 100 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@FreightChargeSub", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ExpDeliveryDate", adVarChar, adParamInput, 20 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkCommission", adVarChar, adParamInput, 10 )
		
		if( Request.Form( "txtQuoteIdMaster" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteIdMaster" )
		if( Request.Form( "txtFreight" ) <> "" ) then cmdEditQuote.Parameters( "@FreightLBS" ) = Left( Request.Form( "txtFreight" ), 100 )
		
		if( Left( Request.Form( "txtFreightChargeSub" ), 1 ) = "$" ) then
			if( Request.Form( "txtFreightChargeSub" ) <> "" ) then cmdEditQuote.Parameters( "@FreightChargeSub" ) = Mid( Request.Form( "txtFreightChargeSub" ), 2 )
		else
			if( Request.Form( "txtFreightChargeSub" ) <> "" ) then cmdEditQuote.Parameters( "@FreightChargeSub" ) = Request.Form( "txtFreightChargeSub" )
		end if
		
		if( Request.Form( "txtExpDeliveryDate" ) <> "" ) then cmdEditQuote.Parameters( "@ExpDeliveryDate" ) = Request.Form( "txtExpDeliveryDate" )
		
		'if( Left( Request.Form( "txtSubWorkCommission" ), 1 ) = "$" ) then
		'	if( Request.Form( "txtSubWorkCommission" ) <> "" ) then cmdEditQuote.Parameters( "@SubWorkCommission" ) = Mid( Request.Form( "txtSubWorkCommission" ), 2 )
		'else
		'	if( Request.Form( "txtSubWorkCommission" ) <> "" ) then cmdEditQuote.Parameters( "@SubWorkCommission" ) = Request.Form( "txtSubWorkCommission" )
		'end if
		
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
		cmdEditQuote.CommandText			= "{? = call spEditQuoteSubWorkDetail( ?,?,?,?,?,? ) }"
		
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteSubWorkId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkCost", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkDescription", adLongVarChar, adParamInput, 1048576 )
		
		if( Request.Form( "txtQuoteIdDetail" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteIdDetail" )
		
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
		cmdCheckSpindleSubWorkDesc.Parameters.Append cmdCheckSpindleSubWorkDesc.CreateParameter( "@SubWorkDesc", adLongVarChar, adParamInput, 1048576 )
		
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
			cmdEditSpindleSubWork.Parameters.Append cmdEditSpindleSubWork.CreateParameter( "@SubWorkDesc", adLongVarChar, adParamInput, 1048576 )
		
		
		
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
		
		intQuoteId = Request.Form( "txtQuoteIdDetail" )
		
	end if
else
	intQuoteId = Request.QueryString( "intQuoteId" )
end if

cmdGetQuoteSubWorkByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteSubWorkByQuoteId.CommandText = "{call spGetQuoteSubWorkByQuoteId( ?,? ) }"
cmdGetQuoteSubWorkByQuoteId.Parameters.Append cmdGetQuoteSubWorkByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteSubWorkByQuoteId.Parameters.Append cmdGetQuoteSubWorkByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuoteSubWorkByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetQuoteSubWorkByQuoteId = cmdGetQuoteSubWorkByQuoteId.Execute
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
	function ConfirmDelete( intQuoteSubWorkId, strSubWork )
	{
		if( confirm( 'Are you sure you want to remove item ' + strSubWork + '?' ) )
		{
			OpenWindow( 'quotesubwork_delete.asp?intQuoteSubWorkId='+intQuoteSubWorkId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function OpenSubWorkPopup( )
	{
	  OpenWindow( 'quotesubwork_details_popup.asp','SelectSubWork','width=660,height=200,left='+GetCenterX( 660 )+',top='+GetCenterY( 200 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 200 )+',scrollbars=no,dependent=yes' );
	}
	
	function OpenUsedSubWorkPopup(intSpindleID, intPartType)
	{
	  OpenWindow( 'quote_subwork_prevused.asp?spindleID=' + intSpindleID + '&partType=' + intPartType + '&intQuoteId=<%=intQuoteId%>' + '&strCategory=SubWork', 'SelectPrevUsedParts','width=660,height=200,left='+GetCenterX( 660 )+',top='+GetCenterY( 200 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 200 )+',scrollbars=yes,dependent=yes' );
	}
	
	function EditQuoteSubWork( intQuoteSubWorkId)
	{
		OpenWindow( 'quote_subwork_details_edit.asp?intQuoteSubWorkId='+intQuoteSubWorkId + '&intQuoteId=<%=intQuoteId%>','EditRecord','width=700,height=160,left='+GetCenterX( 700 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 700 )+',screenY='+GetCenterY( 160 )+',scrollbars=yes,dependent=yes' );
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
		
	function OpenSupplierSearch(  )
	{
	  OpenWindow( 'supplier_search.asp','SupplierSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectSupplier( intSupplierId, strSupplierName )
	{
		document.forms.frmSubWorkDetail.hidSupplierId.value = intSupplierId;
		document.forms.frmSubWorkDetail.txtSupplierName.value = strSupplierName;
	}
  </script>
</head>
<body topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
		if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
			if( not rsGetQuoteSubWorkByQuoteId.EOF ) then
			theSpindleId = rsGetQuoteSubWorkByQuoteId("SpindleId")%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>Rework Cost</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>Quote #&nbsp;&nbsp;<%=intQuoteId%></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- Cost Form -->
<form id="frmSubWorkMaster" name="frmSubWorkMaster" action="" method="post">
<input type="hidden" name="txtQuoteIdMaster" value="<%=intQuoteId%>">
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
			<tr><td class="long_left_label_1">Spindle Type</td><td class="input_2" colspan=3><input type="text" class="forminput" name="txtSpindleType" maxlength=255 readonly value="<%=rsGetQuoteSubWorkByQuoteId( "SpindleType" )%>"></td></tr>
	    <tr><td class="label_1" colspan=4>Freight LBS, Carrier, Priority</td>
		<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtFreight" onchange="this.isDirty=true;" fieldName="Freight LBS, Carrier, Priority" rows=3 wrap=soft><%=rsGetQuoteSubWorkByQuoteId( "FreightLBS" )%></textarea></td></tr>
	    <tr><td class="long_left_label_1">Freight Charge Sub</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtFreightChargeSub" maxlength=10 isCurrency="true" onchange="this.isDirty=true;" fieldName="Freight Charge Sub" value="<%=FormatCurrency( rsGetQuoteSubWorkByQuoteId( "FreightChargeSub" ), 2 )%>"></td>
		  <td class="long_left_label_1">Expedited Delivery Date</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtExpDeliveryDate" onchange="this.isDirty=true;" fieldName="Exp Delivery Date" isDate="true" maxlength=20 value="<%=rsGetQuoteSubWorkByQuoteId( "ExpDeliveryDate" )%>"></td></tr>
		<tr><td class="long_left_label_1">Freight Charge Sub Work</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtFreightChargeSubWork" maxlength=10 readonly value="<%=FormatCurrency( rsGetQuoteSubWorkByQuoteId( "FreightChargeSubWork" ), 2 )%>"></td>
			<td class="long_left_label_1">Total Cost</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtTotalCharges" maxlength=32 readonly value="<%=FormatCurrency( rsGetQuoteSubWorkByQuoteId( "TotalReworkCost" ), 2 )%>" ID="Text2"></td></tr>
		<!--<tr><td class="long_left_label_1">Commission</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtSubWorkCommission" maxlength=10 value="<%=rsGetQuoteSubWorkByQuoteId( "SubWorkCommission" )%>"></td><td colspan=2>&nbsp;</td></tr>-->
	</table>
	<p style="font-size:10px;color:red;text-align:left;padding-left:34px;">markup 1.8</p>
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
							<td width=64 align=right><a href="javascript:validateForm(frmSubWorkMaster);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmSubWorkMaster.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:OpenPrintBox('quote_requisition.asp?intQuoteId=<%=intQuoteId%>');"><IMG src="images/print_button.gif" width=60 height=20 border=0></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmSubWorkMaster);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>
<!-- End Form Buttons -->
</form>
<%	set rsGetQuoteSubWorkByQuoteId = rsGetQuoteSubWorkByQuoteId.NextRecordset
		end if
	end if
end if%>
<!-- End PO Form -->

<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Rework</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmSubWorkDetail" name="frmSubWorkDetail" action="" method="post">
 <input type="hidden" name="txtQuoteIdDetail" value="<%=intQuoteId%>">
 <input type="hidden" name="txtSupplierId" value="">
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
		<%if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
			if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetQuoteSubWorkByQuoteId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td class="itemlist" width="15" align=center><a href="javascript:ConfirmDelete('<%=rsGetQuoteSubWorkByQuoteId( "QuoteSubWorkId" )%>', '<%=EscapeQuotes(rsGetQuoteSubWorkByQuoteId( "SubWorkDescription" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="60" align=middle><a href="javascript:EditQuoteSubWork('<%=rsGetQuoteSubWorkByQuoteId( "QuoteSubWorkId" )%>');"><img src="images/edit_button.gif" width=60 height=20 border=0></a></td>
							<td class="itemlist" width="385"><%=rsGetQuoteSubWorkByQuoteId( "SubWorkDescription" )%></td>
							<td class="itemlist" width="150"><%=rsGetQuoteSubWorkByQuoteId( "SupplierName" )%></td>
							<td class="itemlist" width="50"><%=FormatCurrency( rsGetQuoteSubWorkByQuoteId( "SubWorkCost" ), 2 )%></td>
						</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetQuoteSubWorkByQuoteId.MoveNext
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
			<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value=""><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value=""></td></tr></table></td>
			<td width="50"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value=""></td>
		</tr>
 </table></td>
 </tr>
	<tr><table width="660"><tr>
	
			<%if not (IsNull(theSpindleId)) then%>
				<td class="itemlist" width="408"><a href="javascript:OpenUsedSubWorkPopup(<%=theSpindleId%>, -1);">Select default rework</a></td>
			<%end if%>	
	<td><table cellspacing=0 cellpadding=0 width=192 align=right>
		<tr>
			<td align=right width=64><a href="javascript:validateForm(frmSubWorkDetail);"><img src="images/add_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:OpenPrintBox('quote_requisition.asp?intQuoteId=<%=intQuoteId%>');"><img src="images/print_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
</body>
</html>
<%
if not( rsGetQuoteSubWorkByQuoteId is nothing ) then
	if( rsGetQuoteSubWorkByQuoteId.State = adStateOpen ) then
		rsGetQuoteSubWorkByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuoteSubWorkByQuoteId				= nothing
set cmdGetQuoteSubWorkByQuoteId				= nothing
set cmdEditQuote											= nothing
set cnnSQLConnection									= nothing
%>
