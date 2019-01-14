<%option explicit

dim cnnSQLConnection
dim rsGetProjects, cmdGetProjects
dim lngTotalProjectCount, intRowStart, intRowCount, intFirstRecord, intLastRecord, intNextRecord, intPreviousRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage

set cnnSQLConnection	= nothing
set rsGetProjects			= nothing
set cmdGetProjects		= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetProjects		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetProjects.ActiveConnection	= cnnSQLConnection

cmdGetProjects.CommandText		= "{call spGetProjects( ?,?,?,?,?,? ) }"

cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@OrderBy", adVarChar, adParamInputOutput, 64 )
cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@JumpTo", adVarChar, adParamInput, 11 )

if( Request.QueryString( "intRowStart" ) <> "" ) then
	intRowStart = CInt( Request.QueryString( "intRowStart" ) )
else
	intRowStart = 0
end if

if( Request.QueryString( "blnIsJumpTo" ) = "1" ) then
	if( ( Request.QueryString( "strOrderBy" ) = "StartDate" ) or _
			( Request.QueryString( "strOrderBy" ) = "EstCompletionDate" ) or _
			( Request.QueryString( "strOrderBy" ) = "CompletionDate" ) ) then
		strJumpTo				= Request.QueryString( "strSearch" )
	else
		strJumpTo				= left( Request.QueryString( "strSearch" ), 1 )
	end if
	strSearchString	= ""
else
	strSearchString = Request.QueryString( "strSearch" )
	strJumpTo				= ""
end if

cmdGetProjects.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetProjects.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetProjects.Parameters( "@OrderBy" ).Value				= Request.QueryString( "strOrderBy" )
cmdGetProjects.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetProjects.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetProjects = cmdGetProjects.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0

intRowCount = cmdGetProjects.Parameters( "@RowCount" ).Value
strOrderBy	= cmdGetProjects.Parameters( "@OrderBy" ).Value

if( rsGetProjects.State = adStateOpen ) then
		lngTotalProjectCount = rsGetProjects( "TotalRowCount" )
		
		set rsGetProjects = rsGetProjects.NextRecordset
		
		if( rsGetProjects.State = adStateOpen ) then
			if( not rsGetProjects.EOF ) then
				intRowStart = rsGetProjects( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalProjectCount ) then
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
intLastRecord = lngTotalProjectCount - intRowCount + 1

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Projects</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenProject( strProjectId )
	{
	  OpenWindow( 'project_details.asp?intProjectId='+strProjectId,'CenterlineProject'+strProjectId,'width=657,height=540,left='+GetCenterX( 657 )+',top='+GetCenterY( 540 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 540 )+',scrollbars=yes,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "project_search.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "project_search.asp?blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "project_search.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "project_search.asp?blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	function SelectProject( intProjectId, strProject )
	{
		window.opener.SelectProject( intProjectId, strProject);
		window.close();
	}
	</script>
</head>
<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 onload="strSearchTop.focus();">
<table border=0 cellpadding=0 cellspacing=0>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=672>
 <tr><td><IMG SRC="images/menu/menu_bottom_left_blank.gif" width=104 height=39 border=0></td>
   <td><IMG SRC="images/menu/menu_bottom_middle.gif" width=198 height=39 border=0></td>
   <td><table border=0 cellpadding=0 cellspacing=0 width=374>
     <tr><td colspan=2><IMG SRC="images/menu/menu_bottom_search_top.gif" width=374 height=8 border=0></td></tr>
     <tr><td rowspan=2><IMG SRC="images/menu/menu_bottom_search_left.gif" width=8 height=31 border=0></td>
       <td width=302><table border=0 cellpadding=0 cellspacing=0 width=366>
         <tr><td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=24 height=20 border=0></a></td>
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
     <tr><td valign=bottom><IMG SRC="images/menu/menu_bottom_search_bottom.gif" width=364 height=3 border=0></td></tr>
 </table></td></tr>
</table>

<!-- a little spacer betwen header and content -->
&nbsp;

<!-- Page Contents -->
<table width=672 border=0 cellpadding=0 cellspacing=2>
<tr><td>

<!--Spindle List -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- table heading -->
  <tr>
    <th class="itemlist" width=20>Sel</th>
    <th class="itemlist" width="90"><a href="project_search.asp?strOrderBy=ProjectName">Name</a></th>
    <th class="itemlist" width="90"><a href="project_search.asp?strOrderBy=ProjectType">Project Type</a></th>
    <th class="itemlist" width="182"><a href="project_search.asp?strOrderBy=Customer">Customer</a></th>
    <th class="itemlist" width="60"><a href="project_search.asp?strOrderBy=StartDate">Start Date</a></th>
    <th class="itemlist" width="60"><a href="project_search.asp?strOrderBy=CompletionDate">Complete</a></th>
    <th class="itemlist"><a href="project_search.asp?strOrderBy=ProjectDescription">Description</a></th>
  </tr>  
  <!-- table data -->
  <%if( rsGetProjects.State = adStateOpen ) then
				while( not rsGetProjects.EOF )
					if( ( CInt( rsGetProjects( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						  <td class="itemlist" width=20 align=middle><a href="javascript:SelectProject('<%=rsGetProjects( "ProjectId" )%>', '<%=rsGetProjects( "ProjectName" )%>');"><img src="images/form_select.gif" width=15 height=15 border=0></a></td>
						  <td class="itemlist" width="90" <%if( strOrderBy = "ProjectName" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenProject( '<%=rsGetProjects( "ProjectId" )%>' );"><%=rsGetProjects( "ProjectName" )%></a></td>
						  <td class="itemlist" width="90" <%if( strOrderBy = "ProjectType" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProjects( "ProjectType" )%></td>
						  <td class="itemlist" width="182" <%if( strOrderBy = "Customer" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProjects( "Customer" )%></td>
						  <td class="itemlist" width="60" align=right <%if( strOrderBy = "StartDate" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProjects( "StartDate" )%></td>
						  <td class="itemlist" width="60" align=right <%if( strOrderBy = "CompletionDate" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProjects( "CompletionDate" )%></td>
						  <td class="itemlist" width="170" <%if( strOrderBy = "ProjectDescription" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProjects( "ProjectDescription" )%></td>
					<%else%>
						<tr class="oddrow">
							<td class="itemlist" width=20 align=middle><a href="javascript:SelectProject('<%=rsGetProjects( "ProjectId" )%>', '<%=rsGetProjects( "ProjectName" )%>');"><img src="images/form_select.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="90" <%if( strOrderBy = "ProjectName" ) then%> id="selected_oddrow"<%end if%>><a href="javascript:OpenProject( '<%=rsGetProjects( "ProjectId" )%>' );"><%=rsGetProjects( "ProjectName" )%></a></td>
							<td class="itemlist" width="90" <%if( strOrderBy = "ProjectType" ) then%> id="selected_oddrow"<%end if%>><%=rsGetProjects( "ProjectType" )%></td>
							<td class="itemlist" width="182" <%if( strOrderBy = "Customer" ) then%> id="selected_oddrow"<%end if%>><%=rsGetProjects( "Customer" )%></td>
							<td class="itemlist" width="60" align=right <%if( strOrderBy = "StartDate" ) then%> id="selected_oddrow"<%end if%>><%=rsGetProjects( "StartDate" )%></td>
							<td class="itemlist" width="60" align=right <%if( strOrderBy = "CompletionDate" ) then%> id="selected_oddrow"<%end if%>><%=rsGetProjects( "CompletionDate" )%></td>
							<td class="itemlist" width="170" <%if( strOrderBy = "ProjectDescription" ) then%> id="selected_oddrow"<%end if%>><%=rsGetProjects( "ProjectDescription" )%></td>
					<%end if%>
						</tr>
				<%rsGetProjects.MoveNext
				wend
		end if%>
</table><!-- End of Open Projects --></td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=10><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="project_search.asp?intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=24 height=20 border=0></a></td>
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
    <td align=right valign=bottom width=260><a href="javascript:OpenProject( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=7 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=358 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>  <!-- End of Search and Nav bar-->

</td></tr></table>  <!-- End of Page Contents -->

</body>
</HTML>
<%
if not( rsGetProjects is nothing ) then
	if( rsGetProjects.State = adStateOpen ) then
		rsGetProjects.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProjects				= nothing
set cmdGetProjects			= nothing
set cnnSQLConnection			= nothing
%>