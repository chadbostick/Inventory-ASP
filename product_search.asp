<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetProducts, cmdGetProducts
dim lngTotalProductCount, intRowStart, intRowCount, intFirstRecord, intLastRecord, intNextRecord, intPreviousRecord, strOrderBy, strSearchString, strJumpTo
dim strCurrentPage
dim isMultiple

if ( Request.QueryString( "isMultiple" ) = "true" ) then
	isMultiple = "true"
else
	isMultiple = "false"
end if

set cnnSQLConnection	= nothing
set rsGetProducts			= nothing
set cmdGetProducts		= nothing

set cnnSQLConnection	= Server.CreateObject( "ADODB.Connection" )
set cmdGetProducts		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdGetProducts.ActiveConnection	= cnnSQLConnection

cmdGetProducts.CommandText		= "{call spGetProducts( ?,?,?,?,?,? ) }"

cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@RowStart", adInteger, adParamInputOutput, 4 )
cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@RowCount", adInteger, adParamInputOutput, 4 )
cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@OrderBy", adVarChar, adParamInputOutput, 64 )
cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@SearchString", adVarChar, adParamInput, 255 )
cmdGetProducts.Parameters.Append cmdGetProducts.CreateParameter( "@JumpTo", adVarChar, adParamInput, 11 )

if( Request.QueryString( "intRowStart" ) <> "" ) then
	intRowStart = CInt( Request.QueryString( "intRowStart" ) )
else
	intRowStart = 0
end if

if( Request.QueryString( "blnIsJumpTo" ) = "1" ) then
	if( ( Request.QueryString( "strOrderBy" ) = "ProductId" ) ) then
		strJumpTo				= Request.QueryString( "strSearch" )
	else
		strJumpTo				= left( Request.QueryString( "strSearch" ), 1 )
	end if
	strSearchString	= ""
else
	strSearchString = Request.QueryString( "strSearch" )
	strJumpTo				= ""
end if

cmdGetProducts.Parameters( "@RowStart" ).Value			= intRowStart
cmdGetProducts.Parameters( "@RowCount" ).Value			= Request.QueryString( "intRowCount" )
cmdGetProducts.Parameters( "@OrderBy" ).Value				= Request.QueryString( "strOrderBy" )
cmdGetProducts.Parameters( "@SearchString" ).Value	= strSearchString
cmdGetProducts.Parameters( "@JumpTo" ).Value				= strJumpTo

on error resume next
set rsGetProducts = cmdGetProducts.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
					
on error goto 0

intRowCount = cmdGetProducts.Parameters( "@RowCount" ).Value
strOrderBy	= cmdGetProducts.Parameters( "@OrderBy" ).Value

if( rsGetProducts.State = adStateOpen ) then
		lngTotalProductCount = rsGetProducts( "TotalRowCount" )
		
		set rsGetProducts = rsGetProducts.NextRecordset
		
		if( rsGetProducts.State = adStateOpen ) then
			if( not rsGetProducts.EOF ) then
				intRowStart = rsGetProducts( "RowNumber" )
			end if
		end if
end if

if( ( intRowStart + intRowCount ) > lngTotalProductCount ) then
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
intLastRecord = lngTotalProductCount - intRowCount + 1

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<html>
<head>
<title>Centerline - Products</title>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <script language="Javascript" src="js/admin_menu.js"></script>
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenProduct( strProductId )
	{
	  OpenWindow( 'product_details.asp?intProductId='+strProductId,'CenterlineProduct'+strProductId,'width=660,height=480,left='+GetCenterX( 660 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 660 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function JumpToSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "product_search.asp?isMultiple=<%=isMultiple%>&blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if ( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "product_search.asp?isMultiple=<%=isMultiple%>&blnIsJumpTo=1&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function SearchSubmit()
	{
		if( document.all('strSearchTop').value != "" )
		{
			document.location.href = "product_search.asp?isMultiple=<%=isMultiple%>&blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchTop').value;
		}
		else if( document.all('strSearchBottom').value != "" )
		{
			document.location.href = "product_search.asp?isMultiple=<%=isMultiple%>&blnIsJumpTo=0&strOrderBy=<%=strOrderBy%>&strSearch=" + document.all('strSearchBottom').value;
		}
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	function SelectProduct( intProductId, strProductName, strPartNumber, strUnitPrice, boolKeepOpen )
	{
		window.opener.SelectProduct( intProductId, strProductName, strPartNumber, strUnitPrice, boolKeepOpen);
		if (!boolKeepOpen)
		{
			window.close();
		}
	}
	</script>
</head>
<body topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 onload="strSearchTop.focus();">
<table border=0 cellpadding=0 cellspacing=0>
<tr><td><table border=0 cellpadding=0 cellspacing=0 width=672>
 <tr><td><IMG SRC="images/menu/menu_bottom_left_blank.gif" width=104 height=39 border=0></td>
   <td><IMG SRC="images/menu/menu_bottom_middle.gif" width=198 height=39 border=0></td>
   <td><table border=0 cellpadding=0 cellspacing=0 width=374>
     <tr><td colspan=2><IMG SRC="images/menu/menu_bottom_search_top.gif" width=370 height=8 border=0></td></tr>
     <tr><td rowspan=2><IMG SRC="images/menu/menu_bottom_search_left.gif" width=8 height=31 border=0></td>
       <td width=302><table border=0 cellpadding=0 cellspacing=0 width=366>
         <tr><td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
           <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=24 height=20 border=0></a></td>
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
<table width=672 border=0 cellpadding=0 cellspacing=0>
<tr><td>

<!--Product List -->
<table width=672 border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
  <td colspan=2><table class="itemlist" width=672 border=1 cellpadding=2 cellspacing=0 rules="cols" align=right>
  <!-- table heading -->
  <tr>
    <th class="itemlist" width=20>Sel</th>
    <th class="itemlist" width=25><a href="product_search.asp?isMultiple=<%=isMultiple%>&strOrderBy=ProductId">Id</a></th>
    <th class="itemlist" width=85><a href="product_search.asp?isMultiple=<%=isMultiple%>&strOrderBy=PartNumber">Part #</a></th>
    <th class="itemlist" width=65><a href="product_search.asp?isMultiple=<%=isMultiple%>&strOrderBy=CategoryName">Category</a></th>
    <th class="itemlist"><a href="product_search.asp?isMultiple=<%=isMultiple%>&strOrderBy=ProductShortDescription"> Product Description</a></th>
    <th class="itemlist" width=50>Reorder Lvl</th>
    <th class="itemlist" width=50>Units on Hand</th>
    <th class="itemlist width=50">Units on Order</th>
  </tr>  
  <!-- table data -->
		<%if( rsGetProducts.State = adStateOpen ) then
				while( not rsGetProducts.EOF )
					if( ( CInt( rsGetProducts( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>
						<tr class="oddrow">
					<%end if%>
							<td class="itemlist" width=20 align=middle><a href="javascript:SelectProduct('<%=rsGetProducts( "ProductId" )%>', '<%=EscapeQuotes(rsGetProducts( "ProductShortDescription" ))%>', '<%=EscapeQuotes(rsGetProducts( "PartNumber" ))%>', '<%=rsGetProducts( "UnitPrice" )%>', <%=isMultiple%> );"><img src="images/form_select.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width=25 <%if( strOrderBy = "ProductId" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProducts( "ProductId" )%></td>
							<td class="itemlist" width=85 <%if( strOrderBy = "PartNumber" ) then%> id="selected_evenrow"<%end if%>><a href="javascript:OpenProduct( '<%=rsGetProducts( "ProductId" )%>' );"><%=rsGetProducts( "PartNumber" )%></a></td>
							<td class="itemlist" width=65 <%if( strOrderBy = "CategoryName" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProducts( "CategoryName" )%></td>
							<td class="itemlist" <%if( strOrderBy = "ProductShortDescription" ) then%> id="selected_evenrow"<%end if%>><%=rsGetProducts( "ProductShortDescription" )%></td>
							<td class="itemlist" width=50 align=right><%=rsGetProducts( "ReorderLevel" )%></td>
							<td class="itemlist" width=50 align=right><%=rsGetProducts( "OnHand" )%></td>
							<td class="itemlist" width=50 align=right><%=rsGetProducts( "OnOrder" )%></td>
						</tr>
					<%rsGetProducts.MoveNext
				wend
		end if%>
</table>
<!-- End of Employees -->
</td></tr>

<!-- Search and Nav bar -->
<tr><td><table cellspacing=0 cellpadding=0 width=672 align=right>
  <tr><td colspan=10><IMG src="images/search_bar/search_bar_top.gif" width=672 height=3 border=0></td></tr>
  <tr><td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intFirstRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/first_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intPreviousRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/previous_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intNextRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/next_button.gif" width=24 height=20 border=0></a></td>
    <td valign=bottom width=27><a href="product_search.asp?isMultiple=<%=isMultiple%>&intRowStart=<%=intLastRecord%>&strOrderBy=<%=strOrderBy%><%if( strSearchString <> "" ) then%>&strSearch=<%Response.Write( strSearchString )%><%end if%>"><IMG src="images/search_bar/last_button.gif" width=24 height=20 border=0></a></td>
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
    <td align=right valign=bottom width=260><a href="javascript:OpenProduct( 0 );"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
    <td align=right valign=bottom width=63><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  <tr><td colspan=7 valign=bottom><IMG src="images/search_bar/search_bar_bottom.gif" width=358 height=3 border=0></td>
    <td><IMG src="images/search_bar/search_bar_bottom_corner.gif" width=9 height=7 border=0></td></tr>
 </table></td></tr>
 <!-- End of Search and Nav bar-->

</td></tr></table>  <!-- End of Page Contents -->
</body>
</html>
<%
if not( rsGetProducts is nothing ) then
	if( rsGetProducts.State = adStateOpen ) then
		rsGetProducts.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProducts			= nothing
set cmdGetProducts		= nothing
set cnnSQLConnection	= nothing
%>