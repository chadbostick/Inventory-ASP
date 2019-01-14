<%option explicit

dim cnnSQLConnection
dim rsGetQuotes, cmdGetQuotes
dim lngTotalQuoteCount, intRowStart, intRowCount, intNextRecord, intPreviousRecord, intFirstRecord, intLastRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage

set cnnSQLConnection	= nothing
set rsGetQuotes				= nothing
set cmdGetQuotes			= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuotes			= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetQuotes.ActiveConnection	= cnnSQLConnection

cmdGetQuotes.CommandText		= "{call spGetQuotes( ?,?,?,?,?,? ) }"

cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@OrderBy", adVarChar, adParamInputOutput, 64 )
cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetQuotes.Parameters.Append cmdGetQuotes.CreateParameter( "@JumpTo", adVarChar, adParamInput, 11 )

if( Request.QueryString( "intRowStart" ) <> "" ) then
	intRowStart = CInt( Request.QueryString( "intRowStart" ) )
else
	intRowStart = 0
end if

if( Request.QueryString( "blnIsJumpTo" ) = "1" ) then
	if( ( Request.QueryString( "strOrderBy" ) = "QuoteId" ) or _
			( Request.QueryString( "strOrderBy" ) = "DateQuoted" ) ) then
		strJumpTo				= Request.QueryString( "strSearch" )
	else
		strJumpTo				= left( Request.QueryString( "strSearch" ), 1 )
	end if
	strSearchString	= ""
else
	strSearchString = Request.QueryString( "strSearch" )
	strJumpTo				= ""
end if

cmdGetQuotes.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetQuotes.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetQuotes.Parameters( "@OrderBy" ).Value				= Request.QueryString( "strOrderBy" )
cmdGetQuotes.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetQuotes.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetQuotes = cmdGetQuotes.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0

intRowCount = cmdGetQuotes.Parameters( "@RowCount" ).Value
strOrderBy	= cmdGetQuotes.Parameters( "@OrderBy" ).Value

if( rsGetQuotes.State = adStateOpen ) then
		lngTotalQuoteCount = rsGetQuotes( "TotalRowCount" )
		
		set rsGetQuotes = rsGetQuotes.NextRecordset
		
		if( rsGetQuotes.State = adStateOpen ) then
			if( not rsGetQuotes.EOF ) then
				intRowStart = rsGetQuotes( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalQuoteCount ) then
	intNextRecord = 0
else
	intNextRecord = intRowStart + intRowCount
end if

if( ( intRowStart - intRowCount ) < 0 ) then
	intPreviousRecord = lngTotalQuoteCount - intRowCount
else
	intPreviousRecord = intRowStart - intRowCount
end if

intFirstRecord = 1
intLastRecord = lngTotalQuoteCount - intRowCount + 1

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Quote</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/main_menu.js"></script>
  <script language="Javascript">
	function OpenQuote( strQuoteId )
	{
	  OpenWindow( 'quote_details.asp?intQuoteId='+strQuoteId,'CenterlineQuote'+strQuoteId,'width=680,height=440,left='+GetCenterX( 657 )+',top='+GetCenterY( 440 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 440 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "quote.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "quote.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "quote.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "quote.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	</script>
</head>
<body topmargin=2 leftmargin=2 marginwidth=2 marginheight=2 onload="strSearchTop.focus();">
<table border=0 cellpadding=0 cellspacing=0>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_top_left.gif" width=202 height=54 border=0></a></td>
    <td><IMG SRC="images/menu/menu_top_middle.gif" width=460 height=54 border=0></td>
    <td><IMG SRC="images/menu/menu_icon_Quote.gif" width=108 height=54 border=0></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_left.gif" width=202 height=25 border=0></a></td>
    <td><a href="customer.asp" onmouseOver="imgOn('customer')" onMouseOut="imgOff('customer')">
      <IMG name ="customer" SRC="images/menu/menu_customer_off.gif" width=94 height=25 border=0></a></td>
    <td><a href="project.asp" onmouseOver="imgOn('project')" onMouseOut="imgOff('project')">
      <IMG name ="project" SRC="images/menu/menu_project_off.gif" width=76 height=25 border=0></a></td>
    <td><a href="spindle.asp" onmouseOver="imgOn('spindle')" onMouseOut="imgOff('spindle')">
      <IMG name ="spindle" SRC="images/menu/menu_spindle_off.gif" width=80 height=25 border=0></a></td>
    <td><a href="workorder.asp" onmouseOver="imgOn('work_order')" onMouseOut="imgOff('work_order')">
      <IMG name ="work_order" SRC="images/menu/menu_work_order_off.gif" width=103 height=25 border=0></a></td>
    <td><IMG name ="quote" SRC="images/menu/menu_quote_on.gif" width=70 height=25 border=0></td>
    <td><a href="po.asp" onmouseOver="imgOn('po')" onMouseOut="imgOff('po')">
      <IMG name ="po" SRC="images/menu/menu_po_off.gif" width=64 height=25 border=0></a></td>
    <td><a href="reports.asp" onmouseOver="imgOn('reports')" onMouseOut="imgOff('reports')">
      <IMG name ="reports" SRC="images/menu/menu_reports_off.gif" width=81 height=25 border=0></td></a></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
 <tr><td><a href="index.html"><IMG SRC="images/menu/menu_bottom_left.gif" width=202 height=39 border=0></a></td>
   <td><IMG SRC="images/menu/menu_bottom_middle.gif" width=194 height=39 border=0></td>
   <td><table border=0 cellpadding=0 cellspacing=0 width=374>
     <tr><td colspan=2><IMG SRC="images/menu/menu_bottom_search_top.gif" width=374 height=8 border=0></td></tr>
     <tr><td rowspan=2><IMG SRC="images/menu/menu_bottom_search_left.gif" width=8 height=31 border=0></td>
       <td width=302><table border=0 cellpadding=0 cellspacing=0 width=366>
         <tr><td valign=bottom width=30><a href="quote.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="quote.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="quote.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=30><a href="quote.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
           <td valign=bottom width=126><table cellpadding=0 cellspacing=0>
             <tr><td rowspan=3><IMG src="images/search_bar/input_field/input_field_left.gif" width=4 height=20 border=0></td>
               <td><IMG src="images/search_bar/input_field/input_field_top_bottom.gif" width=115 height=3 border=0></td>
               <td rowspan=3><IMG src="images/search_bar/input_field/input_field_right.gif" width=4 height=20 border=0></td></tr>
             <tr><td><input type=text class="graphical_top_search_input_field" id=strSearchTop name=strSearchTop maxlength=32 onkeypress="CheckEnter(window.event)"></td></tr>
             <tr><td><IMG src="images/search_bar/input_field/input_field_top_bottom.gif" width=115 height=3 border=0></td></tr>
      </table></td>
    <td valign=bottom width=63><a href="javascript:JumpToSubmit();"><IMG src="images/search_bar/jump_to_button.gif" width=60 height=20 border=0></a></td>
    <td valign=bottom width=63><a href="javascript:SearchSubmit();"><IMG src="images/search_bar/search_button.gif" width=60 height=20 border=0></a></td></tr>
   </table></td></tr>
     <tr><td valign=bottom><IMG SRC="images/menu/menu_bottom_search_bottom.gif" width=366 height=3 border=0></td></tr>
 </table></td></tr>
</table>

<!-- a little spacer betwen header and content -->
&nbsp;

<!-- Page Contents -->
<table width=770 border=0 cellpadding=0 cellspacing=0>
<tr><td align=right>

<!--Quote List -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- table heading -->
  <tr>
		<th class="itemlist" width="40"><a href="quote.asp?strOrderBy=QuoteId">Quote Id</a></th>
    <th class="itemlist" width="40"><a href="quote.asp?strOrderBy=WorkOrderNumber">Work Order</a></th>
    <th class="itemlist" width="202"><a href="quote.asp?strOrderBy=Customer">Customer</a></th>
    <th class="itemlist" width="170"><a href="quote.asp?strOrderBy=SpindleType">Spindle Type</a></th>
    <th class="itemlist" width="60"><a href="quote.asp?strOrderBy=DateQuoted">Date Quoted</a></th>
    <th class="itemlist" width="60">Date Approved</th>
    <th class="itemlist" width="100"><a href="quote.asp?strOrderBy=SerialNumber">Serial Number</a></th>
  </tr>
  <!-- table data -->
		<%if( rsGetQuotes.State = adStateOpen ) then
				while( not rsGetQuotes.EOF )
					if( ( CInt( rsGetQuotes( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
							<td class="itemlist" width="40" align="center"<%if( strOrderBy = "QuoteId" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenQuote( '<%=rsGetQuotes( "QuoteId" )%>' );"><%=rsGetQuotes( "QuoteId" )%></a></td>
							<td class="itemlist" width="40" align="center"<%if( strOrderBy = "WorkOrderNumber" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenWorkOrder( '<%=rsGetQuotes( "WorkOrderId" )%>' );"><%=rsGetQuotes( "WorkOrderNumber" )%></a></td>
							<td class="itemlist" width="202"<%if( strOrderBy = "Customer" ) then%> id="selected_evenrow"<%end if%>><%=rsGetQuotes( "Customer" )%></td>
							<td class="itemlist" width="170"<%if( strOrderBy = "SpindleType" ) then%> id="selected_evenrow"<%end if%>><%=rsGetQuotes( "SpindleType" )%></td>
							<td class="itemlist" width="60" align=right<%if( strOrderBy = "DateQuoted" ) then%> id="selected_evenrow"<%end if%>><%=rsGetQuotes( "DateQuoted" )%></td>
							<td class="itemlist" width="60" align=right><%=rsGetQuotes( "DateApproved" )%></td>
							<td class="itemlist" width="100"<%if( strOrderBy = "SerialNumber" ) then%> id="selected_evenrow"<%end if%>><%=rsGetQuotes( "SerialNumber" )%></td>
					<%else%>
						<tr class="oddrow">
							<td class="itemlist" width="40" align="center"<%if( strOrderBy = "QuoteId" ) then%> id="selected_oddrow"<%end if%>><a href="javascript:OpenQuote( '<%=rsGetQuotes( "QuoteId" )%>' );"><%=rsGetQuotes( "QuoteId" )%></a></td>
							<td class="itemlist" width="40" align="center"<%if( strOrderBy = "WorkOrderNumber" ) then%> id="selected_oddrow"<%end if%>><a href="javascript:OpenWorkOrder( '<%=rsGetQuotes( "WorkOrderId" )%>' );"><%=rsGetQuotes( "WorkOrderNumber" )%></a></td>
							<td class="itemlist" width="202"<%if( strOrderBy = "Customer" ) then%> id="selected_oddrow"<%end if%>><%=rsGetQuotes( "Customer" )%></td>
							<td class="itemlist" width="170"<%if( strOrderBy = "SpindleType" ) then%> id="selected_oddrow"<%end if%>><%=rsGetQuotes( "SpindleType" )%></td>
							<td class="itemlist" width="60" align=right<%if( strOrderBy = "DateQuoted" ) then%> id="selected_oddrow"<%end if%>><%=rsGetQuotes( "DateQuoted" )%></td>
							<td class="itemlist" width="60" align=right><%=rsGetQuotes( "DateApproved" )%></td>
							<td class="itemlist" width="100"<%if( strOrderBy = "SerialNumber" ) then%> id="selected_oddrow"<%end if%>><%=rsGetQuotes( "SerialNumber" )%></td>
					<%end if%>
						</tr>
					<%rsGetQuotes.MoveNext
				wend
		end if%>		
</table><!-- End of Quotes --></td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=10><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=30><a href="quote.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="quote.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="quote.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=30><a href="quote.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
    <td valign=bottom width=124><table cellpadding=0 cellspacing=0>
        <tr><td rowspan=3><IMG src="images/search_bar/input_field/input_field_left.gif" width=4 height=20 border=0></td>
          <td><IMG src="images/search_bar/input_field/input_field_top_bottom.gif" width=113 height=3 border=0></td>
          <td rowspan=3><IMG src="images/search_bar/input_field/input_field_right.gif" width=4 height=20 border=0></td></tr>
        <tr><td><input type=text class="graphical_input_field" id=strSearchBottom name=strSearchBottom maxlength=32 onkeypress="CheckEnter(window.event)"></td></tr>
        <tr><td><IMG src="images/search_bar/input_field/input_field_top_bottom.gif" width=113 height=3 border=0></td></tr>
      </table></td>
    <td valign=bottom width=63><a href="javascript:JumpToSubmit();"><IMG src="images/search_bar/jump_to_button.gif" width=60 height=20 border=0></a></td>
    <td valign=bottom width=63><a href="javascript:SearchSubmit();"><IMG src="images/search_bar/search_button.gif" width=60 height=20 border=0></a></td>
    <td align=right width=9><IMG src="images/search_bar/search_bar_right.gif" width=3 height=23 border=0></td>
    <td align=right valign=bottom width=236><a href="javascript:OpenQuote( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=7 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=364 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>  <!-- End of Search and Nav bar-->

</td></tr></table>  <!-- End of Page Contents -->
</body>
</html>
<%
if not( rsGetQuotes is nothing ) then
	if( rsGetQuotes.State = adStateOpen ) then
		rsGetQuotes.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetQuotes			= nothing
set cmdGetQuotes		= nothing
set cnnSQLConnection	= nothing
%>
