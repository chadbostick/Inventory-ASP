<%option explicit

dim cnnSQLConnection
dim rsGetEmployeeByEmployeeId, cmdGetEmployeeByEmployeeId, cmdEditEmployee
dim intRowNumber, intEmployeeId
dim bSaved

set cnnSQLConnection						= nothing
set rsGetEmployeeByEmployeeId		= nothing
set cmdGetEmployeeByEmployeeId	= nothing
set cmdEditEmployee							= nothing

set cnnSQLConnection						= Server.CreateObject( "ADODB.Connection" )
set cmdGetEmployeeByEmployeeId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved = false

if( Request.Form.Count > 0 ) then

	set cmdEditEmployee	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditEmployee.ActiveConnection	= cnnSQLConnection
	cmdEditEmployee.CommandText				= "{? = call spEditEmployee( ?,?,?,?,?,?,?,? ) }"
	
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@EmployeeId", adInteger, adParamInput, 4 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@FirstName", adVarChar, adParamInput, 100 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@LastName", adVarChar, adParamInput, 100 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@Title", adVarChar, adParamInput, 100 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@Extension", adVarChar, adParamInput, 100 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@WorkPhone", adVarChar, adParamInput, 100 )
	cmdEditEmployee.Parameters.Append cmdEditEmployee.CreateParameter( "@EmailAddress", adVarChar, adParamInput, 100 )
	
	if( Request.Form( "txtEmployeeId" ) <> "" ) then cmdEditEmployee.Parameters( "@EmployeeId" ) = Request.Form( "txtEmployeeId" )
	if( Request.Form( "txtFirstName" ) <> "" ) then cmdEditEmployee.Parameters( "@FirstName" ) = Request.Form( "txtFirstName" )
	if( Request.Form( "txtLastName" ) <> "" ) then cmdEditEmployee.Parameters( "@LastName" ) = Request.Form( "txtLastName" )
	if( Request.Form( "txtTitle" ) <> "" ) then cmdEditEmployee.Parameters( "@Title" ) = Request.Form( "txtTitle" )
	if( Request.Form( "txtExt" ) <> "" ) then cmdEditEmployee.Parameters( "@Extension" ) = Request.Form( "txtExt" )
	if( Request.Form( "txtPhone" ) <> "" ) then cmdEditEmployee.Parameters( "@WorkPhone" ) = Request.Form( "txtPhone" )
	if( Request.Form( "txtEmailAddress" ) <> "" ) then cmdEditEmployee.Parameters( "@EmailAddress" ) = Request.Form( "txtEmailAddress" )
	
	on error resume next
	cmdEditEmployee.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
			
	on error goto 0

	intEmployeeId = cmdEditEmployee.Parameters( "@Return" ).Value
	bSaved = true
else
	intEmployeeId = Request.QueryString( "intEmployeeId" )
end if

cmdGetEmployeeByEmployeeId.ActiveConnection = cnnSQLConnection
cmdGetEmployeeByEmployeeId.CommandText = "{call spGetEmployeeByEmployeeId( ?,? ) }"
cmdGetEmployeeByEmployeeId.Parameters.Append cmdGetEmployeeByEmployeeId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetEmployeeByEmployeeId.Parameters.Append cmdGetEmployeeByEmployeeId.CreateParameter( "@EmployeeId", adInteger, adParamInput, 4 )
cmdGetEmployeeByEmployeeId.Parameters( "@EmployeeId" )	= intEmployeeId

on error resume next
set rsGetEmployeeByEmployeeId = cmdGetEmployeeByEmployeeId.Execute
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
<TITLE>Centerline - Add / Edit Employee</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	
	function OpenCall () {
	  OpenWindow('view_edit_call.htm','NewCall','width=660,height=200,left=0,top=0,scrollbars=yes');
	}
	
	function UpdateOpener()
	{
		window.opener.EmployeeSaved( <%=intEmployeeId%>, "<%=rsGetEmployeeByEmployeeId( "LastName" )%>, <%=rsGetEmployeeByEmployeeId( "FirstName" )%>" );
	}
	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetEmployeeByEmployeeId is nothing ) then
	if( rsGetEmployeeByEmployeeId.State = adStateOpen ) then
		if( not rsGetEmployeeByEmployeeId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intEmployeeId = 0 ) then%>
   <tr><td class="sectiontitle" width=25%>&nbsp;&nbsp;New Employee</td>
   <%else%>
   <tr><td class="sectiontitle" width=25%>&nbsp;&nbsp;<%=rsGetEmployeeByEmployeeId( "LastName" )%>,&nbsp;<%=rsGetEmployeeByEmployeeId( "FirstName" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Employee Form -->
<tr><td align=center>
  <form id="frmEmployee" name="frmEmployee" action="" method="post">
		<input type="hidden" name="txtEmployeeId" value="<%=intEmployeeId%>">
    <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
			<tr><td class="label_1">First Name</td><td class="input_1"><input type="text" class="forminput" name="txtFirstName" onchange="this.isDirty=true;" maxlength=100 isRequired="true" fieldName="First Name" value="<%=rsGetEmployeeByEmployeeId( "FirstName" )%>"></td>
			  <td class="label_2">Last Name</td><td class="input_2"><input type="text" class="forminput" name="txtLastName" onchange="this.isDirty=true;" maxlength=100 isRequired="true" fieldName="Last Name" value="<%=rsGetEmployeeByEmployeeId( "LastName" )%>"></td></tr>
			<tr><td class="label_1">Title</td><td class="input_1"><input type="text" class="forminput" name="txtTitle" onchange="this.isDirty=true;" fieldName="Title" maxlength=100 value="<%=rsGetEmployeeByEmployeeId( "Title" )%>"></td>
			  <td class="label_2">Work Phone</td><td class="input_2"><input type="text" class="forminput" name="txtPhone" onchange="this.isDirty=true;" fieldName="Phone" maxlength=100 value="<%=rsGetEmployeeByEmployeeId( "WorkPhone" )%>"></td></tr>
			<tr><td class="label_1">Extension</td><td class="input_1"><input type="text" class="forminput" name="txtExt" onchange="this.isDirty=true;" fieldName="Extension" maxlength=100 value="<%=rsGetEmployeeByEmployeeId( "Extension" )%>"></td>
			  <td class="label_2">E-mail</td><td class="input_2"><input type="text" class="forminput" name="txtEmailAddress" onchange="this.isDirty=true;" fieldName="E-mail" maxlength=100 value="<%=rsGetEmployeeByEmployeeId( "EmailAddress" )%>"></td></tr>
	</table>
 </td></tr> 
 
 
  <!-- Bottom Rule -->
 <tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>  <!-- End Bottom Rule -->
 
  <tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmEmployee);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmEmployee.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmEmployee);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>  <!-- End Employee Form -->
 </td></tr></table>
<%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetEmployeeByEmployeeId is nothing ) then
	if( rsGetEmployeeByEmployeeId.State = adStateOpen ) then
		rsGetEmployeeByEmployeeId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetEmployeeByEmployeeId		= nothing
set cmdGetEmployeeByEmployeeId	= nothing
set cmdEditEmployee							= nothing
set cnnSQLConnection						= nothing
%>