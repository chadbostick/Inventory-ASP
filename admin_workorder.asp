<%option explicit

dim cnnSQLConnection
dim rsGetWorkOrders, cmdGetWorkOrders
dim lngTotalWorkOrderCount, intRowStart, intRowCount, intNextRecord, intPreviousRecord, intFirstRecord, intLastRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage

set cnnSQLConnection	= nothing
set rsGetWorkOrders		= nothing
set cmdGetWorkOrders	= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetWorkOrders	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetWorkOrders.ActiveConnection	= cnnSQLConnection

cmdGetWorkOrders.CommandText		= "{call spGetWorkOrders( ?,?,?,?,?,? ) }"

cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@OrderBy", adVarChar, adParamInputOutput, 64 )
cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetWorkOrders.Parameters.Append cmdGetWorkOrders.CreateParameter( "@JumpTo", adVarChar, adParamInput, 11 )

if( Request.QueryString( "intRowStart" ) <> "" ) then
	intRowStart = CInt( Request.QueryString( "intRowStart" ) )
else
	intRowStart = 0
end if

if( Request.QueryString( "blnIsJumpTo" ) = "1" ) then
	if( ( Request.QueryString( "strOrderBy" ) = "WorkOrder" ) or _
			( Request.QueryString( "strOrderBy" ) = "DateIn" ) or _
			( Request.QueryString( "strOrderBy" ) = "DateOut" ) ) then
		strJumpTo				= Request.QueryString( "strSearch" )
	else
		strJumpTo				= left( Request.QueryString( "strSearch" ), 1 )
	end if
	strSearchString	= ""
else
	strSearchString = Request.QueryString( "strSearch" )
	strJumpTo				= ""
end if

cmdGetWorkOrders.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetWorkOrders.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetWorkOrders.Parameters( "@OrderBy" ).Value				= Request.QueryString( "strOrderBy" )
cmdGetWorkOrders.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetWorkOrders.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetWorkOrders = cmdGetWorkOrders.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
		
on error goto 0

intRowCount = cmdGetWorkOrders.Parameters( "@RowCount" ).Value
strOrderBy	= cmdGetWorkOrders.Parameters( "@OrderBy" ).Value

if( rsGetWorkOrders.State = adStateOpen ) then
		lngTotalWorkOrderCount = rsGetWorkOrders( "TotalRowCount" )
		
		set rsGetWorkOrders = rsGetWorkOrders.NextRecordset
		
		if( rsGetWorkOrders.State = adStateOpen ) then
			if( not rsGetWorkOrders.EOF ) then
				intRowStart = rsGetWorkOrders( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalWorkOrderCount ) then
	intNextRecord = 0
else
	intNextRecord = intRowStart + intRowCount
end if

if( ( intRowStart - intRowCount ) < 0 ) then
	intPreviousRecord = 0
else
	intPreviousRecord = intRowStart - intRowCount
end if

intFirstRecord = 1
intLastRecord = lngTotalWorkOrderCount - intRowCount + 1

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Work Order</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/admin_menu.js"></script>
  <script language="Javascript">
  function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'Centerline','width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "admin_workorder.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "admin_workorder.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "admin_workorder.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "admin_workorder.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function ConfirmDelete( intWorkOrderId, strWorkOrder )
	{
		if( confirm( 'Are you sure you want to remove ' + strWorkOrder + '?' ) )
		{
			OpenWindow( 'workorder_delete.asp?intWorkOrderId='+intWorkOrderId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
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
    <td><IMG SRC="images/menu/menu_icon_work_order.gif" width=108 height=54 border=0></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_left.gif" width=202 height=25 border=0></a></td>
    <td><IMG name SRC="images/menu/menu_delete_middle.gif" width=384 height=25 border=0></td>
    <td><a href="admin.htm" onmouseOver="imgOn('admin_menu')" onMouseOut="imgOff('admin_menu')">
      <IMG name ="admin_menu" SRC="images/menu/menu_admin_menu_off.gif" width=102 height=25 border=0></td></a>
    <td><a href="index.html" onmouseOver="imgOn('main_menu')" onMouseOut="imgOff('main_menu')">
      <IMG name ="main_menu" SRC="images/menu/menu_main_menu_off.gif" width=82 height=25 border=0></a></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
 <tr><td><a href="index.html"><IMG SRC="images/menu/menu_bottom_left.gif" width=202 height=39 border=0></a></td>
   <td><IMG SRC="images/menu/menu_bottom_middle.gif" width=194 height=39 border=0></td>
   <td><table border=0 cellpadding=0 cellspacing=0 width=374>
     <tr><td colspan=2><IMG SRC="images/menu/menu_bottom_search_top.gif" width=374 height=8 border=0></td></tr>
     <tr><td rowspan=2><IMG SRC="images/menu/menu_bottom_search_left.gif" width=8 height=31 border=0></td>
       <td width=302><table border=0 cellpadding=0 cellspacing=0 width=366>
         <tr><td valign=bottom width=30><a href="admin_workorder.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
				   <td valign=bottom width=27><a href="admin_workorder.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="admin_workorder.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=30><a href="admin_workorder.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
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
<!--Work Order -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" name="tItemList" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- caption -->
  <!--<caption valign=bottom align=left id="attention">* Attention Required</caption>-->
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width=20>Del</th>
	<th class="itemlist" width=30><a href="admin_workorder.asp?strOrderBy=WorkOrder">Num</a></th>
    <th class="itemlist" width=60><a href="admin_workorder.asp?strOrderBy=PONumber">PO #</a></th>
    <th class="itemlist" width=120><a href="admin_workorder.asp?strOrderBy=Customer">Customer</a></th>
    <th class="itemlist"><a href="admin_workorder.asp?strOrderBy=SpindleType">Spindle</a></th>
    <th class="itemlist" width=25><a href="admin_workorder.asp?strOrderBy=NewSpindle">New</a></th>
    <th class="itemlist" width=60><a href="admin_workorder.asp?strOrderBy=DateIn">Date In</a></th>
    <th class="itemlist" width=60><a href="admin_workorder.asp?strOrderBy=DateOut">Date Out</a></th>
    <th class="itemlist" width=50><a href="admin_workorder.asp?strOrderBy=Priority">Priority</a></th>
  </tr>
  <!-- table data -->
		<%if( rsGetWorkOrders.State = adStateOpen ) then
				while( not rsGetWorkOrders.EOF )
					if( ( CInt( rsGetWorkOrders( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
							<td class="itemlist" width=20 align=middle><a href="javascript:ConfirmDelete('<%=rsGetWorkOrders( "WorkOrderId" )%>', '<%=rsGetWorkOrders( "WorkOrderNumber" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width=30 align="center" <%if( strOrderBy = "WorkOrder" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenWorkOrder( '<%=rsGetWorkOrders( "WorkOrderId" )%>' );"><%=rsGetWorkOrders( "WorkOrderNumber" )%></a></td>
							<td class="itemlist" width=60 <%if( strOrderBy = "PONumber" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "PONumber" )%></td>
							<td class="itemlist" width=120 <%if( strOrderBy = "Customer" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "Customer" )%></td>
							<td class="itemlist" <%if( strOrderBy = "SpindleType" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "SpindleType" )%></td>
							<td class="itemlist" width=25 align="center" <%if( strOrderBy = "NewSpindle" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "NewSpindle" )%></td>
							<td class="itemlist" width=60 align="right" <%if( strOrderBy = "DateIn" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "DateIn" )%></td>
							<td class="itemlist" width=60 align="right" <%if( strOrderBy = "DateOut" ) then%> id="selected_evenrow"<%end if%>><%=rsGetWorkOrders( "DateOut" )%></td>
							<td class="itemlist" width=50 align="center" <%if( strOrderBy = "Priority" ) then%> id="selected_evenrow"<%end if%>><b><%=UCase( rsGetWorkOrders( "Priority" ) )%></b></td>
					<%else%>
						<tr class="oddrow">
							<td class="itemlist" width=20 align=middle><a href="javascript:ConfirmDelete('<%=rsGetWorkOrders( "WorkOrderId" )%>', '<%=rsGetWorkOrders( "WorkOrderNumber" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width=30 align="center" <%if( strOrderBy = "WorkOrder" ) then%> id="selected_oddrow"<%end if%>><a href="javascript:OpenWorkOrder( '<%=rsGetWorkOrders( "WorkOrderId" )%>' );"><%=rsGetWorkOrders( "WorkOrderNumber" )%></a></td>
							<td class="itemlist" width=60 <%if( strOrderBy = "PONumber" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "PONumber" )%></td>
							<td class="itemlist" width=120 <%if( strOrderBy = "Customer" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "Customer" )%></td>
							<td class="itemlist" <%if( strOrderBy = "SpindleType" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "SpindleType" )%></td>
							<td class="itemlist" width=25 align="center" <%if( strOrderBy = "NewSpindle" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "NewSpindle" )%></td>
							<td class="itemlist" width=60 align="right" <%if( strOrderBy = "DateIn" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "DateIn" )%></td>
							<td class="itemlist" width=60 align="right" <%if( strOrderBy = "DateOut" ) then%> id="selected_oddrow"<%end if%>><%=rsGetWorkOrders( "DateOut" )%></td>
							<td class="itemlist" width=50 align="center" <%if( strOrderBy = "Priority" ) then%> id="selected_oddrow"<%end if%>><b><%=UCase( rsGetWorkOrders( "Priority" ) )%></b></td>
					<%end if%>
						</tr>
					<%rsGetWorkOrders.MoveNext
				wend
		end if%>		
</table><!-- End of Work Order --></td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=10><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=30><a href="admin_workorder.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=27 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="admin_workorder.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="admin_workorder.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=30><a href="admin_workorder.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=27 height=20 border=0></a></td>
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
    <td align=right valign=bottom width=236><a href="javascript:OpenWorkOrder( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=7 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=364 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>  <!-- End of Search and Nav bar-->

</td></tr></table>  <!-- End of Page Contents -->
</body>
</HTML>
<%
if not( rsGetWorkOrders is nothing ) then
	if( rsGetWorkOrders.State = adStateOpen ) then
		rsGetWorkOrders.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetWorkOrders				= nothing
set cmdGetWorkOrders			= nothing
set cnnSQLConnection	= nothing
%>