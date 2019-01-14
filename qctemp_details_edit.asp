<%option explicit

dim cnnSQLConnection
dim rsEditQCTempLog, cmdEditQCTempLog
dim intQCTempLogId, intRowNumber, strCurrentPage
dim intQCTempId, windowFreshener

set cnnSQLConnection				= nothing
set rsEditQCTempLog					= nothing
set cmdEditQCTempLog				= nothing
windowFreshener							= false

set cnnSQLConnection				= Server.CreateObject( "ADODB.Connection" )
set cmdEditQCTempLog				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient


if( Request.Form.Count > 0 ) then
	windowFreshener = true

	set cmdEditQCTempLog	= Server.CreateObject( "ADODB.Command" )

	cmdEditQCTempLog.ActiveConnection	= cnnSQLConnection
	cmdEditQCTempLog.CommandText			= "{? = call spEditQCTemp( ?,?,?,?,?,?,?,?,?,? ) }"

	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempId", adInteger, adParamInput, 4 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempLogId", adInteger, adParamInput, 4 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTime", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCSpeed", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCFront", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCRear", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCShaft", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempLocation", adVarChar, adParamInput, 100 )
	cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCNotes", adLongVarChar, adParamInput, 1048576 )


	if( Request.Form( "txtQCTempId" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempId" ) = Request.Form( "txtQCTempId" )
	if( Request.Form( "txtQCTempLogId" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempLogId" ) = Request.Form( "txtQCTempLogId" )
	if( Request.Form( "txtQCTime" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTime" ) = Request.Form( "txtQCTime" )
	if( Request.Form( "txtQCSpeed" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCSpeed" ) = Request.Form( "txtQCSpeed" )
	if( Request.Form( "txtQCFront" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCFront" ) = Request.Form( "txtQCFront" )
	if( Request.Form( "txtQCRear" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCRear" ) = Request.Form( "txtQCRear" )
	if( Request.Form( "txtQCShaft" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCShaft" ) = Request.Form( "txtQCShaft" )
	if( Request.Form( "txtQCTempLocation" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCTempLocation" ) = Request.Form( "txtQCTempLocation" )
	if( Request.Form( "txtQCNotes" ) <> "" ) then cmdEditQCTempLog.Parameters( "@QCNotes" ) = Request.Form( "txtQCNotes" )

	on error resume next
	cmdEditQCTempLog.Execute
	if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
	end if

	on error goto 0	

	intQCTempId = cmdEditQCTempLog.Parameters( "@QCTempId" )
	intQCTempLogId = cmdEditQCTempLog.Parameters( "@QCTempLogId" )
	
else
	intQCTempId = Request.QueryString( "intQCTempId" )
	intQCTempLogId = Request.QueryString( "intQCTempLogId" )
end if

cmdEditQCTempLog.ActiveConnection = cnnSQLConnection
cmdEditQCTempLog.CommandText = "{call spGetQCTempByQCTempId( ?,? ) }"
cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdEditQCTempLog.Parameters.Append cmdEditQCTempLog.CreateParameter( "@QCTempId", adInteger, adParamInput, 4 )
cmdEditQCTempLog.Parameters( "@QCTempId" )	= intQCTempId

on error resume next
set rsEditQCTempLog = cmdEditQCTempLog.Execute

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
  </script>
</head>
<body topmargin=0 leftmargin=0 <%if (windowFreshener = true) then %> onload="javascript:CloseAndRefreshParent();" <%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>



<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Edit QC Temp</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->
  
 <!-- Parts Table -->
 <form id="frmQCTemp" name="frmQCTemp" action="" method="post">
 <input type="hidden" name="txtQCTempId" value="<%=intQCTempId%>">
 <input type="hidden" name="txtQCTempLogId" value="<%=intQCTempLogId%>">
 <tr><td align=center>
  <table width=585 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=585 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="60">Time</th>
    <th class="itemlist" width="60">Speed</th>
    <th class="itemlist" width="60">Front</th>
    <th class="itemlist" width="60">Rear</th>
    <th class="itemlist" width="60">Shaft</th>
    <th class="itemlist" width="100">Temp Location</th>
    <th class="itemlist">Notes</th>
  </tr>  
 <!-- table data -->
	<%if not( rsEditQCTempLog is nothing ) then
		if( rsEditQCTempLog.State = adStateOpen ) then
			intRowNumber = 1
			
			
			while( not rsEditQCTempLog.EOF )					
				if( ( intRowNumber Mod 2 ) = 1 ) then%>
					<tr class="evenrow">
				<%else%>
					<tr class="oddrow">
				<%end if%>
				<td width="60"><input type="text" class="forminput" name="txtQCTime" isRequired="true" fieldName="Time" value="<%=rsEditQCTempLog( "QCTime" )%>"></td>
				<td width="60"><input type="text" class="forminput" name="txtQCSpeed" fieldName="Speed" value="<%=rsEditQCTempLog( "QCSpeed" )%>"></td>
				<td width="60"><input type="text" class="forminput" name="txtQCFront" fieldName="Front" value="<%=rsEditQCTempLog( "QCFront" )%>"></td>
				<td width="60"><input type="text" class="forminput" name="txtQCRear" fieldName="Rear" value="<%=rsEditQCTempLog( "QCRear" )%>"></td>
				<td width="60"><input type="text" class="forminput" name="txtQCShaft" fieldName="Shaft" value="<%=rsEditQCTempLog( "QCShaft" )%>"></td>
				<td width="100"><input type="text" class="forminput" name="txtQCTempLocation" fieldName="Temp Location" value="<%=rsEditQCTempLog( "QCTempLocation" )%>"></td>
				<td><input type="text" class="forminput" name="txtQCNotes" fieldName="Notes" value="<%=rsEditQCTempLog( "QCNotes" )%>"</td>
				</tr>
				<%intRowNumber = intRowNumber + 1
				rsEditQCTempLog.MoveNext
			wend
		end if
	end if%>
		
			
 </table></td>
 </tr>
	<tr><table width="585"><tr>
	
			
	<td><table cellspacing=0 cellpadding=0 width=128 align=right>
		<tr>
		  <td align=right width=64><a href="javascript:validateForm(frmQCTemp);"><img src="images/save_button.gif" width=60 height=20 border=0></a></td>
		  <td align=right width=64><a href="javascript:RefreshCurrentPage();"><img src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
	</table></td></tr>
	</table></td></tr>
	
	
	
	
 </table>
 </form>
 
 </td></tr>  <!-- End Middle Title Bar -->
 
</body>
</html>
<%
if not( rsEditQCTempLog is nothing ) then
	if( rsEditQCTempLog.State = adStateOpen ) then
		rsEditQCTempLog.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsEditQCTempLog			= nothing
set cmdEditQCTempLog			= nothing
set cnnSQLConnection									= nothing

%>
