<%option explicit

dim cnnSQLConnection
dim rsGetCustomers, cmdGetCustomers
dim lngTotalCustomerCount, intRowStart, intRowCount, intNextRecord, intPreviousRecord, intFirstRecord, intLastRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage

set cnnSQLConnection	= nothing
set rsGetCustomers		= nothing
set cmdGetCustomers		= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetCustomers		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetCustomers.ActiveConnection	= cnnSQLConnection

cmdGetCustomers.CommandText		= "{call spGetCustomers( ?,?,?,?,?,? ) }"

cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@OrderBy", adVarChar, adParamInputOutput, 64 )
cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetCustomers.Parameters.Append cmdGetCustomers.CreateParameter( "@JumpTo", adVarChar, adParamInput, 9 )

if( Request.QueryString( "intRowStart" ) <> "" ) then
	intRowStart = CInt( Request.QueryString( "intRowStart" ) )
else
	intRowStart = 0
end if

if( Request.QueryString( "blnIsJumpTo" ) = "1" ) then
	strJumpTo				= left( Request.QueryString( "strSearch" ), 1 )
	strSearchString	= ""
else
	strSearchString = Request.QueryString( "strSearch" )
	strJumpTo				= ""
end if

cmdGetCustomers.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetCustomers.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetCustomers.Parameters( "@OrderBy" ).Value			= Request.QueryString( "strOrderBy" )
cmdGetCustomers.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetCustomers.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetCustomers = cmdGetCustomers.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
		
on error goto 0

intRowCount = cmdGetCustomers.Parameters( "@RowCount" ).Value
strOrderBy	= cmdGetCustomers.Parameters( "@OrderBy" ).Value

if( rsGetCustomers.State = adStateOpen ) then
		lngTotalCustomerCount = rsGetCustomers( "TotalRowCount" )
		
		set rsGetCustomers = rsGetCustomers.NextRecordset
		
		if( rsGetCustomers.State = adStateOpen ) then
			if( not rsGetCustomers.EOF ) then
				intRowStart = rsGetCustomers( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalCustomerCount ) then
	intNextRecord = 0
else
	intNextRecord = intRowStart + intRowCount
end if

if( ( intRowStart - intRowCount ) < 0 ) then
	intPreviousRecord = lngTotalCustomerCount - intRowCount
else
	intPreviousRecord = intRowStart - intRowCount
end if

intFirstRecord = 1
intLastRecord = lngTotalCustomerCount - intRowCount + 1

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Customer</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/main_menu.js"></script>
  <script language="Javascript">
  function OpenCustomer( strCustomerId )
	{
	  OpenWindow( 'customer_details.asp?intCustomerId='+strCustomerId,'CenterlineCustomer'+strCustomerId,'width=657,height=550,left='+GetCenterX( 657 )+',top='+GetCenterY( 550 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 550 )+',scrollbars=yes,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "customer.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "customer.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "customer.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "customer.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
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
    <td><IMG SRC="images/menu/menu_icon_customer.gif" width=108 height=54 border=0></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_left.gif" width=202 height=25 border=0></a></td>
    <td><IMG name ="customer" SRC="images/menu/menu_customer_on.gif" width=94 height=25 border=0></td>
    <td><a href="project.asp" onmouseOver="imgOn('project')" onMouseOut="imgOff('project')">
      <IMG name ="project" SRC="images/menu/menu_project_off.gif" width=76 height=25 border=0></a></td>
    <td><a href="spindle.asp" onmouseOver="imgOn('spindle')" onMouseOut="imgOff('spindle')">
      <IMG name ="spindle" SRC="images/menu/menu_spindle_off.gif" width=80 height=25 border=0></a></td>
    <td><a href="workorder.asp" onmouseOver="imgOn('work_order')" onMouseOut="imgOff('work_order')">
      <IMG name ="work_order" SRC="images/menu/menu_work_order_off.gif" width=103 height=25 border=0></a></td>
    <td><a href="quote.asp" onmouseOver="imgOn('quote')" onMouseOut="imgOff('quote')">
      <IMG name ="quote" SRC="images/menu/menu_quote_off.gif" width=70 height=25 border=0></a></td>
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
       <td width=366><table border=0 cellpadding=0 cellspacing=0 width=366>
         <tr><td valign=bottom width=30><a href="customer.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="customer.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="customer.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=30><a href="customer.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
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

<!--Spindle List -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="222"><a href="customer.asp?strOrderBy=CustomerName">Customer</a></th>
    <th class="itemlist" width="130"><a href="customer.asp?strOrderBy=Contact">Contact</a></th>
    <th class="itemlist" width="100"><a href="customer.asp?strOrderBy=SalesRep">Sales Rep</a></th>
    <th class="itemlist" width="100"><a href="customer.asp?strOrderBy=City">City</a></th>
    <th class="itemlist" width="40"><a href="customer.asp?strOrderBy=State">State</a></th>
    <th class="itemlist" width="80">Phone</th>
  </tr>
  <!-- table data -->
	<%
	if( rsGetCustomers.State = adStateOpen ) then
		while( not rsGetCustomers.EOF )
			if( ( CInt( rsGetCustomers( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
			<tr class="evenrow">
				<td class="itemlist" width="222"<%if( strOrderBy = "CustomerName" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenCustomer( '<%=rsGetCustomers( "CustomerId" )%>' );"><%=rsGetCustomers( "Customer" )%></a></td>
				<td class="itemlist" width="130"<%if( strOrderBy = "Contact" ) then%> id="selected_evenrow"<%end if%>><%=rsGetCustomers( "Contact" )%></td>
				<td class="itemlist" width="100"<%if( strOrderBy = "SalesRep" ) then%> id="selected_evenrow"<%end if%>><%=rsGetCustomers( "SalesRep" )%></td>
				<td class="itemlist" width="100"<%if( strOrderBy = "City" ) then%> id="selected_evenrow"<%end if%>><%=rsGetCustomers( "City" )%></td>
				<td class="itemlist" width="40"<%if( strOrderBy = "State" ) then%> id="selected_evenrow"<%end if%>><%=rsGetCustomers( "State" )%></td>
				<td class="itemlist" width="80"><%=rsGetCustomers( "TelephoneNumber" )%></td>
			<%else%>
			<tr class="oddrow">
				<td class="itemlist" width="222"<%if( strOrderBy = "CustomerName" ) then%> id="selected_oddrow"<%end if%>><a href="javascript:OpenCustomer( '<%=rsGetCustomers( "CustomerId" )%>' );"><%=rsGetCustomers( "Customer" )%></a></td>
				<td class="itemlist" width="130"<%if( strOrderBy = "Contact" ) then%> id="selected_oddrow"<%end if%>><%=rsGetCustomers( "Contact" )%></td>
				<td class="itemlist" width="100"<%if( strOrderBy = "SalesRep" ) then%> id="selected_oddrow"<%end if%>><%=rsGetCustomers( "SalesRep" )%></td>
				<td class="itemlist" width="100"<%if( strOrderBy = "City" ) then%> id="selected_oddrow"<%end if%>><%=rsGetCustomers( "City" )%></td>
				<td class="itemlist" width="40"<%if( strOrderBy = "State" ) then%> id="selected_oddrow"<%end if%>><%=rsGetCustomers( "State" )%></td>
				<td class="itemlist" width="80"><%=rsGetCustomers( "TelephoneNumber" )%></td>
			<%end if%>
			</tr>
			<%rsGetCustomers.MoveNext
		wend
	end if
	%>
</table><!-- End of Open Products --></td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=10><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=30><a href="customer.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="customer.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="customer.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=30><a href="customer.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
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
    <td align=right valign=bottom width=236><a href="javascript:OpenCustomer( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=7 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=364 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>  <!-- End of Search and Nav bar-->

</td></tr></table>  <!-- End of Page Contents -->
</body>
</html>
<%
if not( rsGetCustomers is nothing ) then
	if( rsGetCustomers.State = adStateOpen ) then
		rsGetCustomers.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCustomers			= nothing
set cmdGetCustomers		= nothing
set cnnSQLConnection	= nothing
%>