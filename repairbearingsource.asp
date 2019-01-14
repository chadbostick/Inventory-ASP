<%option explicit

dim cnnSQLConnection
dim rsGetRepairBearingSources, cmdGetRepairBearingSources
dim lngTotalRepairBearingSourceCount, intRowStart, intRowCount, intNextRecord, intPreviousRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage

set cnnSQLConnection	= nothing
set rsGetRepairBearingSources		= nothing
set cmdGetRepairBearingSources		= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetRepairBearingSources		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetRepairBearingSources.ActiveConnection	= cnnSQLConnection

cmdGetRepairBearingSources.CommandText		= "{call spGetRepairBearingSources( ?,?,?,?,?,? ) }"

cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@OrderBy", adVarChar, adParamInput, 64 )
cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetRepairBearingSources.Parameters.Append cmdGetRepairBearingSources.CreateParameter( "@JumpTo", adVarChar, adParamInput, 9 )

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

cmdGetRepairBearingSources.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetRepairBearingSources.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetRepairBearingSources.Parameters( "@OrderBy" ).Value			= Request.QueryString( "strOrderBy" )
cmdGetRepairBearingSources.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetRepairBearingSources.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetRepairBearingSources = cmdGetRepairBearingSources.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0

intRowCount = cmdGetRepairBearingSources.Parameters( "@RowCount" ).Value

if( rsGetRepairBearingSources.State = adStateOpen ) then
		lngTotalRepairBearingSourceCount = rsGetRepairBearingSources( "TotalRowCount" )
		
		set rsGetRepairBearingSources = rsGetRepairBearingSources.NextRecordset
		
		if( rsGetRepairBearingSources.State = adStateOpen ) then
			if( not rsGetRepairBearingSources.EOF ) then
				intRowStart = rsGetRepairBearingSources( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalRepairBearingSourceCount ) then
	intNextRecord = 0
else
	intNextRecord = intRowStart + intRowCount
end if

if( ( intRowStart - intRowCount ) < 0 ) then
	intPreviousRecord = 0
else
	intPreviousRecord = intRowStart - intRowCount
end if

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Bearing Sources</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script language="Javascript" src="js/admin_menu.js"></script>
  <script language="Javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenRepairBearingSource( strBearingSourceId )
	{
	  OpenWindow( 'repairbearingsource_details.asp?intRepairBearingSourceId='+strBearingSourceId,'Centerline','width=640,height=120,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 160 )+',screenY='+GetCenterY( 480 )+',scrollbars=no,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "repairbearingsource.asp?blnIsJumpTo=1&strOrderBy=<%=Request.QueryString( "strOrderBy" )%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "repairbearingsource.asp?blnIsJumpTo=1&strOrderBy=<%=Request.QueryString( "strOrderBy" )%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "repairbearingsource.asp?blnIsJumpTo=0&strOrderBy=<%=Request.QueryString( "strOrderBy" )%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "repairbearingsource.asp?blnIsJumpTo=0&strOrderBy=<%=Request.QueryString( "strOrderBy" )%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function ConfirmDelete( intBearingSourceId, strBearingSource )
	{
		if( confirm( 'Are you sure you want to remove ' + strBearingSource + '?' ) )
		{
			OpenWindow( 'repairbearingsource_delete.asp?intRepairBearingSourceId='+intBearingSourceId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
  </script>
</head>
<body topmargin=2 leftmargin=2 marginwidth=2 marginheight=2>
<table border=0 cellpadding=0 cellspacing=0>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_top_left.gif" width=202 height=54 border=0></a></td>
    <td><IMG SRC="images/menu/menu_top_middle.gif" width=460 height=54 border=0></td>
    <td><IMG SRC="images/menu/menu_icon_admin.gif" width=108 height=54 border=0></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
  <tr><td><a href="index.html"><IMG SRC="images/menu/menu_left.gif" width=202 height=25 border=0></a></td>
    <td><IMG name SRC="images/menu/menu_bearings_source_middle.gif" width=384 height=25 border=0></td>
    <td><a href="admin.htm" onmouseOver="imgOn('admin_menu')" onMouseOut="imgOff('admin_menu')">
      <IMG name ="admin_menu" SRC="images/menu/menu_admin_menu_off.gif" width=102 height=25 border=0></td></a>
    <td><a href="index.html" onmouseOver="imgOn('main_menu')" onMouseOut="imgOff('main_menu')">
      <IMG name ="main_menu" SRC="images/menu/menu_main_menu_off.gif" width=82 height=25 border=0></a></td></tr>
 </table></td></tr>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=770>
 <tr><td><a href="index.html"><IMG SRC="images/menu/menu_bottom_left.gif" width=202 height=39 border=0></a></td>
   <td><IMG SRC="images/menu/menu_bottom_middle.gif" width=258 height=39 border=0></td>
   <td><table border=0 cellpadding=0 cellspacing=0 width=310>
     <tr><td colspan=2><IMG SRC="images/menu/menu_bottom_search_top.gif" width=310 height=8 border=0></td></tr>
     <tr><td rowspan=2><IMG SRC="images/menu/menu_bottom_search_left.gif" width=8 height=31 border=0></td>
       <td width=302><table border=0 cellpadding=0 cellspacing=0 width=302>
         <tr><td valign=bottom width=27><a href="repairbearingsource.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="repairbearingsource.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
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
     <tr><td valign=bottom><IMG SRC="images/menu/menu_bottom_search_bottom.gif" width=302 height=3 border=0></td></tr>
 </table></td></tr>
</table>

<!-- a little spacer betwen header and content -->
&nbsp;

<!-- Page Contents -->
<table width=770 border=0 cellpadding=2 cellspacing=2>
<tr><td>

<!--Bearing Sources -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- table heading -->
  <tr>
		<th class="itemlist" width=20>Del</th>
    <th class="itemlist"><a href="repairbearingsource.asp?strOrderBy=BearingSource">Bearing Source</a></th>
  </tr>  
  <!-- table data -->
	<%if( rsGetRepairBearingSources.State = adStateOpen ) then
			while( not rsGetRepairBearingSources.EOF )
				if( ( CInt( rsGetRepairBearingSources( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
				<%else%>
					<tr class="oddrow">
				<%end if%>
						<td class="itemlist" width=20 align=middle><a href="javascript:ConfirmDelete('<%=rsGetRepairBearingSources( "BearingSourceId" )%>', '<%=rsGetRepairBearingSources( "BearingSource" )%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
						<td class="itemlist"><a href="javascript:OpenRepairBearingSource('<%=rsGetRepairBearingSources( "BearingSourceId" )%>');"><%=rsGetRepairBearingSources( "BearingSource" )%></a></td>
				</tr>
			<%rsGetRepairBearingSources.MoveNext
		wend
	end if%>
</table>
<!-- End of Bearing Sources --></td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=8><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=27><a href="repairbearingsource.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="repairbearingsource.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=Request.QueryString( "strOrderBy" )%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
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
    <td align=right valign=bottom width=296><a href="javascript:OpenRepairBearingSource( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=5 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=304 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>
 <!-- End of Search and Nav bar-->

</td></tr></table><!-- End of Page Contents -->
</body>
</HTML>
<%
if not( rsGetRepairBearingSources is nothing ) then
	if( rsGetRepairBearingSources.State = adStateOpen ) then
		rsGetRepairBearingSources.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetRepairBearingSources	= nothing
set cmdGetRepairBearingSources	= nothing
set cnnSQLConnection			= nothing
%>