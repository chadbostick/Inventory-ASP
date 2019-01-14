<%option explicit

dim cnnSQLConnection
dim rsGetWorkOrderList, cmdGetWorkOrderList
dim rsGetQuoteByQuoteId, cmdGetQuoteByQuoteId
dim cmdEditQuote
dim intQuoteId, intDefaultQuoteId
dim rsGetCustomerContactsByCustomerId, cmdGetCustomerContactsByCustomerId
dim rsGetEmployeesList, cmdGetEmployeesList

set cnnSQLConnection							= nothing
set rsGetWorkOrderList							= nothing
set cmdGetWorkOrderList							= nothing
set rsGetQuoteByQuoteId							= nothing
set cmdGetQuoteByQuoteId						= nothing
set cmdEditQuote								= nothing
set rsGetCustomerContactsByCustomerId			= nothing
set cmdGetCustomerContactsByCustomerId			= nothing
set rsGetEmployeesList						= nothing
set cmdGetEmployeesList					= nothing

set cnnSQLConnection							= Server.CreateObject( "ADODB.Connection" )
set cmdGetWorkOrderList							= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteByQuoteId						= Server.CreateObject( "ADODB.Command" )
set cmdGetEmployeesList					= Server.CreateObject( "ADODB.Command" )
set cmdGetCustomerContactsByCustomerId								= Server.CreateObject( "ADODB.Command" )

intDefaultQuoteId = 0

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
cmdGetWorkOrderList.ActiveConnection	= cnnSQLConnection
cmdGetWorkOrderList.CommandText				= "{call spGetWorkOrderList}"

set rsGetWorkOrderList = cmdGetWorkOrderList.Execute
'**********************************************************************************			
if( Request.Form.Count > 0 ) then
	set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditQuote.ActiveConnection	= cnnSQLConnection
	cmdEditQuote.CommandText			= "{? = call spEditQuote( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteID", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@WorkOrderNumber", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SerialNumber", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@DateApproved", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@DateQuoted", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Notes", adLongVarChar, adParamInput, 1048576 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteSpecificComments", adLongVarChar, adParamInput, 1048576 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@DefaultQuoteId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteContactId", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@DeliveryInformation", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ExpeditedDeliveryInformation", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuotedById", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@HandlingCharge", adVarChar, adParamInput, 10 )
	
	if( Request.Form( "txtQuoteId" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteId" )
	if( Request.Form( "hidWorkOrderId" ) <> "" ) then cmdEditQuote.Parameters( "@WorkOrderId" ) = Request.Form( "hidWorkOrderId" )
	if( Request.Form( "txtWorkOrderNumber" )<> "" ) then cmdEditQuote.Parameters( "@WorkOrderNumber" ) = Request.Form( "txtWorkOrderNumber" )
	if( Request.Form( "txtSerialNumber" ) <> "" ) then cmdEditQuote.Parameters( "@SerialNumber" ) = Request.Form( "txtSerialNumber" )
	if( Request.Form( "txtDateApproved" ) <> "" ) then cmdEditQuote.Parameters( "@DateApproved" ) = Request.Form( "txtDateApproved" )
	if( Request.Form( "txtDateQuoted" ) <> "" ) then cmdEditQuote.Parameters( "@DateQuoted" ) = Request.Form( "txtDateQuoted" )
	if( Request.Form( "txtNotes" ) <> "" ) then cmdEditQuote.Parameters( "@Notes" ) = Request.Form( "txtNotes" )
	if( Request.Form( "txtQuoteSpecificComments" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteSpecificComments" ) = Request.Form( "txtQuoteSpecificComments" )
	if( Request.Form( "txtDefaultQuoteId" ) <> "" ) then cmdEditQuote.Parameters( "@DefaultQuoteId" ) = Request.Form( "txtDefaultQuoteId" )
	if( Request.Form( "selQuoteContactId" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteContactId" ) = Request.Form( "selQuoteContactId" )
	if( Request.Form( "txtDeliveryInformation" ) <> "" ) then cmdEditQuote.Parameters( "@DeliveryInformation" ) = Request.Form( "txtDeliveryInformation" )
	if( Request.Form( "txtExpeditedDeliveryInformation" ) <> "" ) then cmdEditQuote.Parameters( "@ExpeditedDeliveryInformation" ) = Request.Form( "txtExpeditedDeliveryInformation" )
	if( Request.Form( "selQuotedById" ) <> "" ) then cmdEditQuote.Parameters( "@QuotedById" ) = Request.Form( "selQuotedById" )
	if( Left( Request.Form( "txtHandlingCharge" ), 1 ) = "$" ) then
		cmdEditQuote.Parameters( "@HandlingCharge" ) = Mid( Request.Form( "txtHandlingCharge" ), 2 )
	else
		cmdEditQuote.Parameters( "@HandlingCharge" ) = Request.Form( "txtHandlingCharge" )
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
else
	if( Request.QueryString( "intQuoteId" ) = "" ) then
		intQuoteId = 0
	else
		intQuoteId = Request.QueryString( "intQuoteId" )
	end if
end if

cmdGetQuoteByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteByQuoteId.CommandText = "{call spGetQuoteByQuoteId( ?,?,? ) }"
cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@CustomerId", adInteger, adParamOutput, 4 )
cmdGetQuoteByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetQuoteByQuoteId = cmdGetQuoteByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
								
on error goto 0
'**********************************************************************************
cmdGetCustomerContactsByCustomerId.ActiveConnection = cnnSQLConnection
cmdGetCustomerContactsByCustomerId.CommandText = "{call spGetCustomerContactsByCustomerId ( ?,? ) }"
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
cmdGetCustomerContactsByCustomerId.Parameters( "@CustomerId" )	= cmdGetQuoteByQuoteId.Parameters( "@CustomerId" )

on error resume next
set rsGetCustomerContactsByCustomerId = cmdGetCustomerContactsByCustomerId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0

'**********************************************************************************
cmdGetEmployeesList.ActiveConnection = cnnSQLConnection
cmdGetEmployeesList.CommandText = "{call spGetEmployeeList}"

on error resume next
set rsGetEmployeesList = cmdGetEmployeesList.Execute
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
<TITLE>Centerline - View / Edit Quote</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script language="javascript">
	function OpenDefaultQuote( strQuoteId )
	{
	  OpenWindow( 'quote_details.asp?intQuoteId='+strQuoteId+'&strDefaultQuote=true','CenterlineQuote0','width=657,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenLaborDetails( strQuoteId )
	{
		<%if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" ) ) + "';" )
		else
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) ) + "';" )
		end if%>
		
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
			OpenWindow( 'quotelabor_details.asp?strPrevPage='+strPrevPage+'&intQuoteId='+strQuoteId,'CenterlineQuote<%=intQuoteId%>','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenPartsDetails( strQuoteId )
	{
		<%if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" ) ) + "';" )
		else
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) ) + "';" )
		end if%>
		
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
			OpenWindow( 'quoteparts_details.asp?strPrevPage='+strPrevPage+'&intQuoteId='+strQuoteId,'CenterlineQuote<%=intQuoteId%>','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes,status=yes' );
	}
	
	function OpenBearingsDetails( strQuoteId )
	{
		<%if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" ) ) + "';" )
		else
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) ) + "';" )
		end if%>
		
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
			OpenWindow( 'quotebearings_details.asp?strPrevPage='+strPrevPage+'&intQuoteId='+strQuoteId,'CenterlineQuote<%=intQuoteId%>','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenSubWorkDetails( strQuoteId )
	{
		<%if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" ) ) + "';" )
		else
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) ) + "';" )
		end if%>
		
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
			OpenWindow( 'quotesubwork_details.asp?strPrevPage='+strPrevPage+'&intQuoteId='+strQuoteId,'CenterlineQuote<%=intQuoteId%>','width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	
	function OpenSpindle( strSpindleId )
	{
	  OpenWindow( 'spindle_details.asp?intSpindleId='+strSpindleId,'CenterlineSpindle'+strSpindleId,'width=660,height=540,left='+GetCenterX( 660 )+',top='+GetCenterY( 600 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 600 )+',scrollbars=yes,dependent=yes' );
	}
	
	
	function PrintQuote( strQuoteId )
	{
		<%if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" ) ) + "';" )
		else
			Response.Write( "var strPrevPage = '" + Server.URLEncode( Request.ServerVariables( "URL" ) ) + "';" )
		end if%>
		
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
		{	
			var isFullQuote = document.forms.frmQuote.chkExpanded.checked;
			OpenWindow( 'print_quote.asp?isFullQuote='+isFullQuote+'&strPrevPage='+strPrevPage+'&intQuoteId='+strQuoteId,'PrintQuote'+strQuoteId,'width=690,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
		}
	}	
	
	function EmailQuote( strQuoteId )
	{
		if( strQuoteId == '0' )
			alert( 'You must save your quote before continuing.' );
		else
		{	
			var isFullQuote = document.forms.frmQuote.chkExpanded.checked;
			OpenWindow( 'email_quote.asp?isFullQuote='+isFullQuote+'&intQuoteId='+strQuoteId,'EmailQuote'+strQuoteId,'width=640,height=180,left='+GetCenterX( 640 )+',top='+GetCenterY( 180 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 180 )+',scrollbars=no,dependent=yes' );
		}
	}
	
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenCustomer( strCustomerId )
	{
	  OpenWindow( 'customer_details.asp?intCustomerId='+strCustomerId,'CenterlineCustomer'+strCustomerId,'width=657,height=550,left='+GetCenterX( 657 )+',top='+GetCenterY( 550 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 550 )+',scrollbars=yes,dependent=yes' );
	}
	function OpenProject( strProjectId )
	{
	  OpenWindow( 'project_details.asp?intProjectId='+strProjectId,'CenterlineProject'+strProjectId,'width=657,height=540,left='+GetCenterX( 657 )+',top='+GetCenterY( 540 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 540 )+',scrollbars=yes,dependent=yes' );
	}		
	function OpenWorkOrderSearch(  )
	{
	  OpenWindow( 'workorder_search.asp','WorkOrderSearch','width=689,height=500,left='+GetCenterX( 689 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 689 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectWorkOrder( intWorkOrderId, strWorkOrderNumber )
	{
		document.forms.frmQuote.hidWorkOrderId.value = intWorkOrderId;
		document.forms.frmQuote.txtWorkOrderNumber.value = strWorkOrderNumber;
	}
	</script>
</head>
<body topmargin=0 leftmargin=0>
<%if not( rsGetQuoteByQuoteId is nothing ) then
		if( rsGetQuoteByQuoteId.State = adStateOpen ) then
			if( not rsGetQuoteByQuoteId.EOF ) then%>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
		<%if( intQuoteId = 0 ) then%>
		<tr><td class="sectiontitle" width=20%>New Quote</td>
		<%else%>
		<tr><td class="sectiontitle" width=20%>Quote&nbsp;&nbsp;#<%=intQuoteId%></td>
		<%end if%>
     <td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=left width=325><a href="javascript:OpenCustomer(<%=rsGetQuoteByQuoteId( "CustomerId" )%>)"><%=rsGetQuoteByQuoteId( "Customer" )%></a>&nbsp;&nbsp;:&nbsp;&nbsp;Project <a href="javascript:OpenProject(<%=rsGetQuoteByQuoteId( "ProjectId" )%>)"><%=rsGetQuoteByQuoteId( "ProjectName" )%></a>&nbsp;&nbsp;:&nbsp;&nbsp;W.O. # <a href="javascript:OpenWorkOrder(<%=rsGetQuoteByQuoteId( "WorkOrderId" )%>)"><%=rsGetQuoteByQuoteId( "WorkOrderNumber" )%></a></td></tr>
 </table>
</td></tr>
<!-- End Top Title Bar -->

<!-- Quote Form -->
<form id="frmQuote" name="frmQuote" action="" method="post">
<tr><td align=center>
	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
			<input type="hidden" name="txtQuoteId" value="<%=intQuoteId%>">
			<input type="hidden" name="txtDefaultQuoteId" value="<%=intDefaultQuoteId%>">
		<tr><td class="label_1">Work Order #</td><td class="input_1"><table cellpadding=0 cellspacing=0 width=100%><tr><td><a href="javascript:OpenWorkOrderSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidWorkOrderId" value="<%=rsGetQuoteByQuoteId( "WorkOrderId" )%>"><input class="forminput" name="txtWorkOrderNumber" readonly isRequired="true" onchange="this.isDirty=true;" fieldName="Work Order #" value="<%=rsGetQuoteByQuoteId( "WorkOrderNumber" )%>"></td></tr></table></td>
			<td class="label_2">Quoted By</td><td class="input_2"><select class="forminput" name="selQuotedById" size=1 fieldName="Quoted By" onChange="javascript:this.isDirty=true;">
   			<option value=""></option>
   			<%if not( rsGetEmployeesList is nothing ) then
   					if( rsGetEmployeesList.State = adStateOpen ) then
   						while( not rsGetEmployeesList.EOF )%><option value="<%=rsGetEmployeesList( "EmployeeId" )%>"<%if( rsGetEmployeesList( "EmployeeId" ) = rsGetQuoteByQuoteId( "QuotedById" ) ) then%>SELECTED<%end if%>><%=rsGetEmployeesList( "EmployeeName" )%></option>
   							<%rsGetEmployeesList.MoveNext
   						wend
   					end if
   				end if%>
   			</select></td></tr>
		<tr><td class="label_1">Date Quoted</td><td class="input_1"><input type="text" class="forminput" name="txtDateQuoted" onchange="this.isDirty=true;" fieldName="Date Quoted" maxlength=20 value="<%if( IsDate( rsGetQuoteByQuoteId( "DateQuoted" ) ) ) then Response.Write( FormatDateTime( rsGetQuoteByQuoteId( "DateQuoted" ), vbShortDate ) ) end if%>"></td>
   				<td class="label_2">Date Appr.</td><td class="input_2"><input type="text" class="forminput" name="txtDateApproved" onchange="this.isDirty=true;" fieldName="Date Appr." maxlength=20 value="<%if( IsDate( rsGetQuoteByQuoteId( "DateApproved" ) ) ) then Response.Write( FormatDateTime( rsGetQuoteByQuoteId( "DateApproved" ), vbShortDate ) ) end if%>"></td></tr>
   	<tr><td class="label_1">Spindle Type</td><td class="generic_input"><a href="javascript:OpenSpindle(<%=rsGetQuoteByQuoteId( "SpindleId" )%>);"><%=rsGetQuoteByQuoteId( "SpindleType" )%></a></td><td class="label_2">Serial #</td><td class="generic_input"><%=rsGetQuoteByQuoteId( "SerialNumber" )%></td></tr>
   	<tr><td class="label_1">Work Order</td><td class="generic_input"><a href="javascript:OpenWorkOrder(<%=rsGetQuoteByQuoteId( "WorkOrderId" )%>);"><%=rsGetQuoteByQuoteId( "WorkOrderNumber" )%></a></td>
   		<td class="label_2">Contact</td><td class="input_2"><select class="forminput" name="selQuoteContactId" size=1 fieldName="Quote Contact" onChange="javascript:this.isDirty=true;">
   			<option value=""></option>
   	    <%if not( rsGetCustomerContactsByCustomerId is nothing ) then
   					if( rsGetCustomerContactsByCustomerId.State = adStateOpen ) then
   						while( not rsGetCustomerContactsByCustomerId.EOF )%><option value="<%=rsGetCustomerContactsByCustomerId( "CustomerContactId" )%>"<%if( rsGetCustomerContactsByCustomerId( "CustomerContactId" ) = rsGetQuoteByQuoteId( "QuoteContactId" ) ) then%>SELECTED<%end if%>><%=rsGetCustomerContactsByCustomerId( "Contact" )%></option>
   							<%rsGetCustomerContactsByCustomerId.MoveNext
   						wend
   					end if
   				end if%>
   			</select></td></tr>
		<tr><td class="label_1" colspan=4>Notes</td>
		<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtNotes" rows=3 wrap=soft><%=rsGetQuoteByQuoteId( "Notes" )%></textarea></td></tr>
		<tr><td class="label_1" colspan=4>Quote Specific Comments</td>
		<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtQuoteSpecificComments" rows=5 cols=80 wrap=soft><%=rsGetQuoteByQuoteId( "QuoteSpecificComments" )%></textarea></td></tr>
	</table>
 </td></tr>

 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
  <tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td class="label_1"><a href="javascript:OpenLaborDetails( '<%=intQuoteId%>' );">Labor</a></td><td class="input_1"><input type="text" class="forminput" name="txtLabor" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalLaborCost" ), 2 )%>"></td>
		  <td class="label_2"><a href="javascript:OpenBearingsDetails( '<%=intQuoteId%>' );">Bearings</a></td><td class="input_2"><input type="text" class="forminput" name="txtBearings" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalBearingCharge" ), 2 )%>"></td></tr>
		<tr><td class="label_1" colspan=4>&nbsp;</td></tr>
	    <tr><td class="label_1"><a href="javascript:OpenPartsDetails( '<%=intQuoteId%>' );">Parts</a></td><td class="input_1"><input type="text" class="forminput" name="txtParts" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalPartCharge" ), 2 )%>"></td>
		  <td class="label_2"><a href="javascript:OpenSubWorkDetails( '<%=intQuoteId%>' );">Rework</a></td><td class="input_2"><input type="text" class="forminput" name="txtRework" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalReworkCharge" ), 2 )%>"></td></tr>
		<tr><td class="label_1" colspan=4>&nbsp;</td></tr>
	    <tr><td class="label_1">Freight</td><td class="input_1"><input type="text" class="forminput" name="txtFreight" maxlength=10 readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalFreightCharge" ), 2 )%>"></td>
		  <td class="label_2">Handling</td><td class="input_2"><input type="text" isRequired="true" isCurrency="true" fieldName="Handling" class="forminput" name="txtHandlingCharge" maxlength=10 value="<%=FormatCurrency( rsGetQuoteByQuoteId( "HandlingCharge" ), 2 )%>"></td></tr>
	</table>
 </td></tr>
	  
 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>  <!-- End Middle Title Bar -->
 
  <tr><td align=center>
  	<!--<p style="font-size:10px;color:red;">price increase notes in Bearings, Parts, Labor and Rework screens.</p>//-->
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td class="label_1" colspan=4>&nbsp;</td>
	    <tr><td class="label_1">Total Repair Price</td><td class="input_1"><input type="text" class="forminput" name="txtTotalRepairPrice" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalRepairPrice" ), 2 )%>"></td>
		  <td class="label_2">Total Incl. Exp. Charge</td><td class="input_2"><input type="text" class="forminput" name="txtTotalInclExpCharge" readonly value="<%=FormatCurrency( rsGetQuoteByQuoteId( "TotalExpeditingCharge" ), 2 )%>"></td></tr>
	    <tr><td class="label_1" colspan=4>&nbsp;</td>
	    <tr><td class="label_1">Delivery Info</td><td class="input_1" colspan="3"><input type="text" class="forminput" maxlength=100 name="txtDeliveryInformation" value="<%=rsGetQuoteByQuoteId( "DeliveryInformation" )%>"></td></tr>
		  <tr><td class="label_1">Exp. Delivery Info</td><td class="input_1" colspan="3"><input type="text" class="forminput" maxlength=100 name="txtExpeditedDeliveryInformation" value="<%=rsGetQuoteByQuoteId( "ExpeditedDeliveryInformation" )%>"></td></tr>
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
	    <% if( intQuoteId = 0 ) then %>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=256 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmQuote);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmQuote.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:PrintQuote( '<%=intQuoteId%>' );"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmQuote);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	     <% else %>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=600 align=right>
	          <tr><td width=64 align=left><a href="javascript:frmQuote.txtQuoteId.value=0; frmQuote.txtDefaultQuoteId.value=<%=intQuoteId%>; frmQuote.submit();"><IMG src="images/copy_button.gif" width=60 height=20 border=0></a></td>
	            <td width=128 align="right" class="generic_label"><input type="checkbox" name="chkExpanded" id="chkExpanded" >Detailed</td>
	            <td width=64 align=right><a href="javascript:PrintQuote( '<%=intQuoteId%>' );"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:EmailQuote( '<%=intQuoteId%>' );"><IMG src="images/email_button.gif" width=60 height=20 border=0></a></td>
	            <td width=216 align=right><a href="javascript:validateForm(frmQuote);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmQuote.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmQuote);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	     <% end if %>
	 </table>
</td></tr>
<!-- End Form Buttons -->
</form>
<!-- End PO Form -->
</table>
<%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetQuoteByQuoteId is nothing ) then
	if( rsGetQuoteByQuoteId.State = adStateOpen ) then
		rsGetQuoteByQuoteId.Close
	end if
end if

if not( rsGetWorkOrderList is nothing ) then
	if( rsGetWorkOrderList.State = adStateOpen ) then
		rsGetWorkOrderList.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetWorkOrderList							= nothing
set cmdGetWorkOrderList							= nothing
set rsGetQuoteByQuoteId							= nothing
set cmdGetQuoteByQuoteId						= nothing
set cmdEditQuote										= nothing
set cnnSQLConnection								= nothing
%>
