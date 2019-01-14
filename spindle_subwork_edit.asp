<%option explicit

dim cnnSQLConnection
dim rsGetSpindleSubWorkDetailsBySpindleSubWorkId, cmdGetSpindleSubWorkDetailsBySpindleSubWorkId
dim cmdEditSpindleSubWork
dim intQuoteId, intRowNumber, strCurrentPage
dim theSpindleSubWorkId, theProductId, retVal, theSubWorkDesc, windowFreshener, theSpindleId

set cnnSQLConnection											= nothing
set rsGetSpindleSubWorkDetailsBySpindleSubWorkId				= nothing
set cmdGetSpindleSubWorkDetailsBySpindleSubWorkId				= nothing
set cmdEditSpindleSubWork										= nothing
windowFreshener													= false

set cnnSQLConnection											= Server.CreateObject( "ADODB.Connection" )
set cmdGetSpindleSubWorkDetailsBySpindleSubWorkId				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	windowFreshener = true

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
		
		
		
	cmdEditSpindleSubWork.Parameters( "@SpindleSubWorkId" ) = Request.Form("theSpindleSubWorkId")
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
		
	
else
	theSpindleSubWorkId = Request.QueryString( "intSpindleSubWorkId" )
	theSpindleId = Request.QueryString( "intSpindleId" )
end if



cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.ActiveConnection = cnnSQLConnection
cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.CommandText = "{call spGetSpindleSubWorkDetailsBySpindleSubWorkId( ?,? ) }"

cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.Parameters.Append cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.Parameters.Append cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.CreateParameter( "@SpindleSubWorkId", adInteger, adParamInput, 4 )

cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.Parameters( "@SpindleSubWorkId" )	= theSpindleSubWorkId

on error resume next
set rsGetSpindleSubWorkDetailsBySpindleSubWorkId = cmdGetSpindleSubWorkDetailsBySpindleSubWorkId.Execute
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
	function CloseAndRefreshParent()
	{
		window.opener.RefreshCurrentPage();
		window.close();
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
<body topmargin=0 leftmargin=0 <%if (windowFreshener = true) then %> onload="javascript:CloseAndRefreshParent();" <%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>



<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Edit Sub Work</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 
 <!-- Parts Table -->
 <form id="frmSubWorkDetail" name="frmSubWorkDetail" action="" method="post">
 <input type="hidden" name="txtSupplierId" value="">
 <input type="hidden" name="theSpindleSubWorkId" value="<%=theSpindleSubWorkId%>">
 <input type="hidden" name="theSpindleId" value="<%=theSpindleId%>">
 <tr><td align=center>
  <table width=585 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=585 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="385">Sub Work Description</th>
    <th class="itemlist" width="150">Source</th>
    <th class="itemlist" width="50">Cost</th>
  </tr>  
 <!-- table data -->  
		<%if not( rsGetSpindleSubWorkDetailsBySpindleSubWorkId is nothing ) then
			if( rsGetSpindleSubWorkDetailsBySpindleSubWorkId.State = adStateOpen ) then
				intRowNumber = 1
				
				while( not rsGetSpindleSubWorkDetailsBySpindleSubWorkId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
						<%else%>
						<tr class="oddrow">
						<%end if%>
							<td width="385"><input type="text" class="forminput" isRequired="true" fieldName="Sub Work" name="txtSubWorkDescription" value="<%=rsGetSpindleSubWorkDetailsBySpindleSubWorkId( "SubWorkDescription" )%>"></td>
							<td><table cellpadding=0 cellspacing=0 width="150"><tr><td><a href="javascript:OpenSupplierSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSupplierId" value="<%=rsGetSpindleSubWorkDetailsBySpindleSubWorkId( "SupplierId" )%>"><input class="forminput" name="txtSupplierName" readonly isRequired="true" fieldName="Supplier" value="<%=rsGetSpindleSubWorkDetailsBySpindleSubWorkId( "SupplierName" )%>"></td></tr></table></td>
							<td width="50"><input type="text" class="forminput" name="txtCost" isRequired="true" isCurrency="true" fieldName="Cost" value="<%=FormatCurrency( rsGetSpindleSubWorkDetailsBySpindleSubWorkId( "SubWorkCost" ), 2 )%>"></td>
						</tr>
				<%intRowNumber = intRowNumber + 1
				rsGetSpindleSubWorkDetailsBySpindleSubWorkId.MoveNext
				wend
			end if
		end if%>
		
 </table></td>
 </tr>
	<tr><table width="585"><tr>
	
				
	<td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr>
			<td align=right width=64><a href="javascript:validateForm(frmSubWorkDetail);"><img src="images/save_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
</body>
</html>
<%
if not( rsGetSpindleSubWorkDetailsBySpindleSubWorkId is nothing ) then
	if( rsGetSpindleSubWorkDetailsBySpindleSubWorkId.State = adStateOpen ) then
		rsGetSpindleSubWorkDetailsBySpindleSubWorkId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetSpindleSubWorkDetailsBySpindleSubWorkId		= nothing
set cmdGetSpindleSubWorkDetailsBySpindleSubWorkId		= nothing
set cmdEditSpindleSubWork								= nothing
set cnnSQLConnection									= nothing

%>