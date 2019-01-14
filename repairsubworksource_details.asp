<%option explicit

dim cnnSQLConnection
dim rsGetRepairSubWorkSourceByRepairSubWorkSourceId, cmdGetRepairSubWorkSourceByRepairSubWorkSourceId, cmdEditRepairSubWorkSource
dim intRowNumber, intRepairSubWorkSourceId

set cnnSQLConnection																	= nothing
set rsGetRepairSubWorkSourceByRepairSubWorkSourceId		= nothing
set cmdGetRepairSubWorkSourceByRepairSubWorkSourceId	= nothing
set cmdEditRepairSubWorkSource												= nothing

set cnnSQLConnection																	= Server.CreateObject( "ADODB.Connection" )
set cmdGetRepairSubWorkSourceByRepairSubWorkSourceId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then

	set cmdEditRepairSubWorkSource	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditRepairSubWorkSource.ActiveConnection		= cnnSQLConnection
	cmdEditRepairSubWorkSource.CommandText				= "{? = call spEditRepairSubWorkSource( ?,?,? ) }"
	
	cmdEditRepairSubWorkSource.Parameters.Append cmdEditRepairSubWorkSource.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditRepairSubWorkSource.Parameters.Append cmdEditRepairSubWorkSource.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditRepairSubWorkSource.Parameters.Append cmdEditRepairSubWorkSource.CreateParameter( "@RepairSubWorkSourceId", adInteger, adParamInput, 4 )
	cmdEditRepairSubWorkSource.Parameters.Append cmdEditRepairSubWorkSource.CreateParameter( "@RepairSubWorkSource", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtSubWorkSourceId" ) <> "" ) then cmdEditRepairSubWorkSource.Parameters( "@RepairSubWorkSourceId" ) = Request.Form( "txtSubWorkSourceId" )
	if( Request.Form( "txtSubWorkSource" ) <> "" ) then cmdEditRepairSubWorkSource.Parameters( "@RepairSubWorkSource" ) = Request.Form( "txtSubWorkSource" )
	
	on error resume next
	cmdEditRepairSubWorkSource.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0

	intRepairSubWorkSourceId = cmdEditRepairSubWorkSource.Parameters( "@Return" ).Value
else
	intRepairSubWorkSourceId = Request.QueryString( "intRepairSubWorkSourceId" )
end if

cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.ActiveConnection = cnnSQLConnection
cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.CommandText = "{call spGetRepairSubWorkSourceByRepairSubWorkSourceId( ?,? ) }"
cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.Parameters.Append cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.Parameters.Append cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.CreateParameter( "@RepairSubWorkSourceId", adInteger, adParamInput, 4 )
cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.Parameters( "@RepairSubWorkSourceId" )	= intRepairSubWorkSourceId

on error resume next
set rsGetRepairSubWorkSourceByRepairSubWorkSourceId = cmdGetRepairSubWorkSourceByRepairSubWorkSourceId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Add Sub Work Source</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetRepairSubWorkSourceByRepairSubWorkSourceId is nothing ) then
		if( rsGetRepairSubWorkSourceByRepairSubWorkSourceId.State = adStateOpen ) then
			if( not rsGetRepairSubWorkSourceByRepairSubWorkSourceId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intRepairSubWorkSourceId = 0 ) then %>
   <tr><td class="sectiontitle" width=35%>New Sub Work Source</td>
   <%else%>
   <tr><td class="sectiontitle" width=35%>Sub Work Source #<%=intRepairSubWorkSourceId%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Parts Source Form -->
<form id="frmPartsSource" name="frmPartsSource" action="" method="post">
<input type="hidden" name="txtSubWorkSourceId" value="<%=intRepairSubWorkSourceId%>">
<tr><td align=center>	
   <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Sub Work Source</td><td class="input_1"><input type="text" class="forminput" name="txtSubWorkSource" maxlength=50 isRequired="true" fieldName="Sub Work Source" value="<%=rsGetRepairSubWorkSourceByRepairSubWorkSourceId( "SubWorkSource" )%>"></td></tr>
   </table>
</td></tr>
 
<!-- Bottom Rule -->
<tr><td height=12>
<TABLE class="" width=100% cellpadding=0 cellspacing=0>
  <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=4></td></tr>
</TABLE>
</td></tr>
<!-- End Bottom Rule -->
 
<tr><td align=center>
 	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
          <tr><td width=64 align=right><a href="javascript:validateForm(frmPartsSource);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmPartsSource.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
<!-- End Parts Source Form -->
<%	end if
	end if
end if%>
 
</BODY>
</HTML>

<%
if not( rsGetRepairSubWorkSourceByRepairSubWorkSourceId is nothing ) then
	if( rsGetRepairSubWorkSourceByRepairSubWorkSourceId.State = adStateOpen ) then
		rsGetRepairSubWorkSourceByRepairSubWorkSourceId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetRepairSubWorkSourceByRepairSubWorkSourceId		= nothing
set cmdGetRepairSubWorkSourceByRepairSubWorkSourceId	= nothing
set cmdEditRepairSubWorkSource												= nothing
set cnnSQLConnection																	= nothing
%>