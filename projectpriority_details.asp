<%option explicit

dim cnnSQLConnection
dim rsGetProjectPriorityByProjectPriorityId, cmdGetProjectPriorityByProjectPriorityId, cmdEditProjectPriority
dim intRowNumber, intProjectPriorityId
dim bSaved

set cnnSQLConnection													= nothing
set rsGetProjectPriorityByProjectPriorityId		= nothing
set cmdGetProjectPriorityByProjectPriorityId	= nothing
set cmdEditProjectPriority										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetProjectPriorityByProjectPriorityId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved = false

if( Request.Form.Count > 0 ) then

	set cmdEditProjectPriority	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditProjectPriority.ActiveConnection	= cnnSQLConnection
	cmdEditProjectPriority.CommandText			= "{? = call spEditProjectPriority( ?,?,? ) }"
	
	cmdEditProjectPriority.Parameters.Append cmdEditProjectPriority.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditProjectPriority.Parameters.Append cmdEditProjectPriority.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditProjectPriority.Parameters.Append cmdEditProjectPriority.CreateParameter( "@ProjectPriorityId", adInteger, adParamInput, 4 )
	cmdEditProjectPriority.Parameters.Append cmdEditProjectPriority.CreateParameter( "@ProjectPriority", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtProjectPriorityId" ) <> "" ) then cmdEditProjectPriority.Parameters( "@ProjectPriorityId" ) = Request.Form( "txtProjectPriorityId" )
	if( Request.Form( "txtProjectPriority" ) <> "" ) then cmdEditProjectPriority.Parameters( "@ProjectPriority" ) = Request.Form( "txtProjectPriority" )
	
	on error resume next
	cmdEditProjectPriority.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
	
	on error goto 0

	intProjectPriorityId = cmdEditProjectPriority.Parameters( "@Return" ).Value
	bSaved = true
else
	intProjectPriorityId = Request.QueryString( "intProjectPriorityId" )
end if

cmdGetProjectPriorityByProjectPriorityId.ActiveConnection = cnnSQLConnection
cmdGetProjectPriorityByProjectPriorityId.CommandText = "{call spGetProjectPriorityByProjectPriorityId( ?,? ) }"
cmdGetProjectPriorityByProjectPriorityId.Parameters.Append cmdGetProjectPriorityByProjectPriorityId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProjectPriorityByProjectPriorityId.Parameters.Append cmdGetProjectPriorityByProjectPriorityId.CreateParameter( "@ProjectPriorityId", adInteger, adParamInput, 4 )
cmdGetProjectPriorityByProjectPriorityId.Parameters( "@ProjectPriorityId" )	= intProjectPriorityId

on error resume next
set rsGetProjectPriorityByProjectPriorityId = cmdGetProjectPriorityByProjectPriorityId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
	
on error goto 0
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Add Project Priority</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers	
	function UpdateOpener()
	{
		window.opener.ProjectPrioritySaved( <%=intProjectPriorityId%>, "<%=rsGetProjectPriorityByProjectPriorityId( "ProjectPriority" )%>" );
	}	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetProjectPriorityByProjectPriorityId is nothing ) then
	if( rsGetProjectPriorityByProjectPriorityId.State = adStateOpen ) then
		if( not rsGetProjectPriorityByProjectPriorityId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intProjectPriorityId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Project Priority</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetProjectPriorityByProjectPriorityId( "ProjectPriority" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmProjectPriority" name="frmProjectPriority" action="" method="post">
 <input type="hidden" name="txtProjectPriorityId" value="<%=intProjectPriorityId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Project Priority</td><td class="input_1"><input type="text" class="forminput" name="txtProjectPriority" maxlength=50 isRequired="true" fieldName="Project Priority" value="<%=rsGetProjectPriorityByProjectPriorityId( "ProjectPriority" )%>"></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmProjectPriority);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmProjectPriority.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:window.close();"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>   <!-- End Shipping Method Form -->
 <%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetProjectPriorityByProjectPriorityId is nothing ) then
	if( rsGetProjectPriorityByProjectPriorityId.State = adStateOpen ) then
		rsGetProjectPriorityByProjectPriorityId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProjectPriorityByProjectPriorityId		= nothing
set cmdGetProjectPriorityByProjectPriorityId	= nothing
set cmdEditProjectPriority										= nothing
set cnnSQLConnection													= nothing
%>