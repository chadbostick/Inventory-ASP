<%@ Language=VBScript %>
<%option explicit

dim cnnSQLConnection
dim rsGetSpindleSubWorkBySpindleId, cmdGetSpindleSubWorkBySpindleId
dim intSpindleId, intPartType, intRowNumber, strCurrentPage, intQuoteId
dim cmdEditQuote
dim i
dim windowFreshener, strCategory

set cnnSQLConnection									= nothing
set rsGetSpindleSubWorkBySpindleId						= nothing
set cmdGetSpindleSubWorkBySpindleId						= nothing
set cmdEditQuote										= nothing

set cnnSQLConnection									= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleSubWorkBySpindleId						= Server.CreateObject( "ADODB.Command" )
windowFreshener = false

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient



'user hit save on the form, so we need to save all the items to the DB
if( Request.Form.Count > 0 ) then
		'Response.Write ("form variables set")
		intQuoteId = Request.Form("intQuoteId")
		strCategory = Request.Form("strCategory")
		windowFreshener = true
		
		set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditQuote.ActiveConnection	= cnnSQLConnection
		cmdEditQuote.CommandText			= "{? = call spEditQuoteSubWorkDetail( ?,?,?,?,?,? ) }"
		
		
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Quote" + strCategory + "Id", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkCost", adVarChar, adParamInput, 10 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SupplierId", adInteger, adParamInput, 4 )
		cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@SubWorkDescription", adLongVarChar, adParamInput, 1048576 )
		
		
		
		for i = 1 to CInt(Request.form("checkboxCount"))
			if Request.Form("usedPart"+CStr(i)) = "on" then
				cmdEditQuote.Parameters( "@QuoteID" ) = intQuoteId
				cmdEditQuote.Parameters( "@Quote" + strCategory + "Id" ) = 0
				cmdEditQuote.Parameters( "@SubWorkCost" ) = Request.Form("SubWorkCost" + CStr(i))
				cmdEditQuote.Parameters( "@SupplierId" ) = Request.Form("supplierId" + CStr(i))
				cmdEditQuote.Parameters( "@SubWorkDescription" ) = Request.Form("SubWorkDesc" + CStr(i))
				
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
			end if
		next
		
		'close window, and refresh parent
		
else

	intSpindleId = Request.QueryString( "spindleId" )
	intPartType = Request.QueryString( "partType" )
	intQuoteId = Request.QueryString( "intQuoteId" )
	strCategory = Request.QueryString("strCategory")
	'Response.Write(CStr(intQuoteId))
	
	cmdGetSpindleSubWorkBySpindleId.ActiveConnection = cnnSQLConnection
	cmdGetSpindleSubWorkBySpindleId.CommandText = "{call spGetSpindleSubWorkBySpindleId( ?,? ) }"
	
	cmdGetSpindleSubWorkBySpindleId.Parameters.Append cmdGetSpindleSubWorkBySpindleId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetSpindleSubWorkBySpindleId.Parameters.Append cmdGetSpindleSubWorkBySpindleId.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
	
	cmdGetSpindleSubWorkBySpindleId.Parameters( "@SpindleId" )	= intSpindleId
	
	'Response.Write( "no variables set")
	
	on error resume next
	set rsGetSpindleSubWorkBySpindleId = cmdGetSpindleSubWorkBySpindleId.Execute
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
<TITLE>Centerline - View / Add Previous Used Rework</TITLE>
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
<table width="600">
<tr><td>
  <table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="15" align="center">Add</th>
    <th class="itemlist" width="250">Rework Description</th>
    <th class="itemlist" width="150">Source</th>
    <th class="itemlist" width="70">Cost</th>
  </tr>  
  
 <!-- table data -->
<%if not( rsGetSpindleSubWorkBySpindleId is nothing ) then
	if( rsGetSpindleSubWorkBySpindleId.State = adStateOpen ) then
		intRowNumber = 1
			
		while( not rsGetSpindleSubWorkBySpindleId.EOF )
			if( ( intRowNumber Mod 2 ) = 1 ) then%>
				<tr class="evenrow">
			<%else%>
				<tr class="oddrow">
			<%end if%>
 
			<input type=hidden name="spindleSubWork<%=intRowNumber%>" value="<%=rsGetSpindleSubWorkBySpindleId("SpindleSubWorkId")%>">
			<td class="itemlist" width="15" align=middle><input type="checkbox" class="forminput" name="usedpart<%=intRowNumber%>" fieldName="usedpart"></td>
			
			<input type=hidden name="SubWorkDesc<%=intRowNumber%>" value="<%=rsGetSpindleSubWorkBySpindleId("SubWorkDescription")%>">
			<td class="itemlist" width="250"><%=rsGetSpindleSubWorkBySpindleId( "SubWorkDescription" )%></td>
			
			<input type=hidden name="supplierId<%=intRowNumber%>" value="<%=rsGetSpindleSubWorkBySpindleId("SupplierId")%>">
			<td class="itemlist" width="150"><%=rsGetSpindleSubWorkBySpindleId( "SupplierName" )%></td>
			
			<input type=hidden name="SubWorkCost<%=intRowNumber%>" value="<%=rsGetSpindleSubWorkBySpindleId( "SubWorkCost" )%>">
			<td class="itemlist" width="70" align="right"><%=FormatCurrency( rsGetSpindleSubWorkBySpindleId( "SubWorkCost" ), 2 )%></td>
			
			</tr>
			<%intRowNumber = intRowNumber + 1
			rsGetSpindleSubWorkBySpindleId.MoveNext
		wend%>
		
		<input type="hidden" name="checkboxCount" value="<%=intRowNumber-1%>">
		<input type="hidden" name="intQuoteId" value="<%=intQuoteId%>">
		<input type="hidden" name="strCategory" value="<%=strCategory%>">
	<%end if
end if%>
</table></td></tr>


<tr><td><table cellspacing=0 cellpadding=0 width=128 align=right border=0>
	<tr><td width="64" align="right"><a href="javascript:validateForm(prevUsedParts);"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td>
	<td width="64" align="right"><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
</table></td></tr>


</form>
</BODY>
</HTML>

<%
if not( rsGetSpindleSubWorkBySpindleId is nothing ) then
	if( rsGetSpindleSubWorkBySpindleId.State = adStateOpen ) then
		rsGetSpindleSubWorkBySpindleId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleSubWorkBySpindleId				= nothing
set cmdGetSpindleSubWorkBySpindleId				= nothing
set cmdEditQuote								= nothing
set cnnSQLConnection							= nothing
%>
