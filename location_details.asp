<%option explicit

dim cnnSQLConnection
dim rsGetLocationByLocationId, cmdGetLocationByLocationId, cmdEditLocation
dim intRowNumber, intLocationId

set cnnSQLConnection												= nothing
set rsGetLocationByLocationId		= nothing
set cmdGetLocationByLocationId	= nothing
set cmdEditLocation										= nothing

set cnnSQLConnection												= Server.CreateObject( "ADODB.Connection" )
set cmdGetLocationByLocationId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then

	set cmdEditLocation	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditLocation.ActiveConnection	= cnnSQLConnection
	cmdEditLocation.CommandText				= "{? = call spEditLocation( ?,?,? ) }"
	
	cmdEditLocation.Parameters.Append cmdEditLocation.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditLocation.Parameters.Append cmdEditLocation.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditLocation.Parameters.Append cmdEditLocation.CreateParameter( "@LocationId", adInteger, adParamInput, 4 )
	cmdEditLocation.Parameters.Append cmdEditLocation.CreateParameter( "@Location", adVarChar, adParamInput, 100 )
		
	if( Request.Form( "txtLocationId" ) <> "" ) then cmdEditLocation.Parameters( "@LocationId" ) = Request.Form( "txtLocationId" )
	if( Request.Form( "txtLocation" ) <> "" ) then cmdEditLocation.Parameters( "@Location" ) = Request.Form( "txtLocation" )
	
	on error resume next
	cmdEditLocation.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0

	intLocationId = cmdEditLocation.Parameters( "@Return" ).Value
else
	intLocationId = Request.QueryString( "intLocationId" )
end if

cmdGetLocationByLocationId.ActiveConnection = cnnSQLConnection
cmdGetLocationByLocationId.CommandText = "{call spGetLocationByLocationId( ?,? ) }"
cmdGetLocationByLocationId.Parameters.Append cmdGetLocationByLocationId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetLocationByLocationId.Parameters.Append cmdGetLocationByLocationId.CreateParameter( "@LocationId", adInteger, adParamInput, 4 )
cmdGetLocationByLocationId.Parameters( "@LocationId" )	= intLocationId

on error resume next
set rsGetLocationByLocationId = cmdGetLocationByLocationId.Execute
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
<TITLE>Centerline - Add Location</TITLE>
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

<%
if not( rsGetLocationByLocationId is nothing ) then
	if( rsGetLocationByLocationId.State = adStateOpen ) then
		if( not rsGetLocationByLocationId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intLocationId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Location</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetLocationByLocationId( "Location" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Location Form -->
<tr><td align=center>
 <form id="frmLocation" name="frmLocation" action="" method="post">
 <input type="hidden" name="txtLocationId" value="<%=intLocationId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Location</td><td class="input_1"><input type="text" class="forminput" name="txtLocation" onchange="this.isDirty=true;" maxlength=100 isRequired="true" fieldName="Location" value="<%=rsGetLocationByLocationId( "Location" )%>"></td></tr>
   </table>
 
  <!-- Bottom Rule -->
 <tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>  <!-- End Bottom Rule -->
 
  <tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmLocation);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmLocation.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmLocation);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>   <!-- End Location Form -->
 <%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetLocationByLocationId is nothing ) then
	if( rsGetLocationByLocationId.State = adStateOpen ) then
		rsGetLocationByLocationId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetLocationByLocationId		= nothing
set cmdGetLocationByLocationId	= nothing
set cmdEditLocation							= nothing
set cnnSQLConnection						= nothing
%>