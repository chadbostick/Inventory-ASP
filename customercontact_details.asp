<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetCustomerContactByCustomerContactId, cmdGetCustomerContactByCustomerContactId, cmdEditCustomerContact
dim intRowNumber, intCustomerContactId, intCustomerId
dim bSaved

set cnnSQLConnection													= nothing
set rsGetCustomerContactByCustomerContactId		= nothing
set cmdGetCustomerContactByCustomerContactId	= nothing
set cmdEditCustomerContact										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetCustomerContactByCustomerContactId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

bSaved = false

intCustomerContactId = Request.Form( "txtCustomerContactId" )
intCustomerId = Request.QueryString( "intCustomerId" )

if( Request.Form.Count > 0 ) then

	set cmdEditCustomerContact	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditCustomerContact.ActiveConnection	= cnnSQLConnection
	cmdEditCustomerContact.CommandText				= "{? = call spEditCustomerContact( ?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@CustomerContactId", adInteger, adParamInput, 4 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@Contact", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@ContactTitle", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@Department", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@TelephoneNumber", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@Extension", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@MobileNumber", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@FaxNumber", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@EmailAddress", adVarChar, adParamInput, 100 )
	cmdEditCustomerContact.Parameters.Append cmdEditCustomerContact.CreateParameter( "@Notes", adLongVarChar, adParamInput, 1048576 )
	
	if( Request.Form( "txtCustomerContactId" ) <> "" ) then cmdEditCustomerContact.Parameters( "@CustomerContactId" ) = Request.Form( "txtCustomerContactId" )
	if( Request.Form( "txtCustomerId" ) <> "" ) then cmdEditCustomerContact.Parameters( "@CustomerId" ) = Request.Form( "txtCustomerId" )
	if( Request.Form( "txtContact" ) <> "" ) then cmdEditCustomerContact.Parameters( "@Contact" ) = Request.Form( "txtContact" )
	if( Request.Form( "txtContactTitle" ) <> "" ) then cmdEditCustomerContact.Parameters( "@ContactTitle" ) = Request.Form( "txtContactTitle" )
	if( Request.Form( "txtDepartment" ) <> "" ) then cmdEditCustomerContact.Parameters( "@Department" ) = Request.Form( "txtDepartment" )
	if( Request.Form( "txtPhoneNumber" ) <> "" ) then cmdEditCustomerContact.Parameters( "@TelephoneNumber" ) = Request.Form( "txtPhoneNumber" )
	if( Request.Form( "txtExtension" ) <> "" ) then cmdEditCustomerContact.Parameters( "@Extension" ) = Request.Form( "txtExtension" )
	if( Request.Form( "txtMobileNumber" ) <> "" ) then cmdEditCustomerContact.Parameters( "@MobileNumber" ) = Request.Form( "txtMobileNumber" )
	if( Request.Form( "txtFaxNumber" ) <> "" ) then cmdEditCustomerContact.Parameters( "@FaxNumber" ) = Request.Form( "txtFaxNumber" )
	if( Request.Form( "txtEmailAddress" ) <> "" ) then cmdEditCustomerContact.Parameters( "@EmailAddress" ) = Request.Form( "txtEmailAddress" )
	if( Request.Form( "txtNotes" ) <> "" ) then cmdEditCustomerContact.Parameters( "@Notes" ) = Request.Form( "txtNotes" )
	
	on error resume next
	cmdEditCustomerContact.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
			
	on error goto 0

	intCustomerContactId = cmdEditCustomerContact.Parameters( "@Return" ).Value
	bSaved = true
else
	intCustomerContactId = Request.QueryString( "intCustomerContactId" )
end if

cmdGetCustomerContactByCustomerContactId.ActiveConnection = cnnSQLConnection
cmdGetCustomerContactByCustomerContactId.CommandText = "{call spGetCustomerContactByCustomerContactId( ?,? ) }"
cmdGetCustomerContactByCustomerContactId.Parameters.Append cmdGetCustomerContactByCustomerContactId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomerContactByCustomerContactId.Parameters.Append cmdGetCustomerContactByCustomerContactId.CreateParameter( "@CustomerContactId", adInteger, adParamInput, 4 )
cmdGetCustomerContactByCustomerContactId.Parameters( "@CustomerContactId" )	= intCustomerContactId

on error resume next
set rsGetCustomerContactByCustomerContactId = cmdGetCustomerContactByCustomerContactId.Execute
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
<TITLE>Centerline - Add / Edit CustomerContact</TITLE>
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
		window.opener.RefreshCurrentPage();
	}
	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if (bSaved) then%>onload="UpdateOpener();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetCustomerContactByCustomerContactId is nothing ) then
	if( rsGetCustomerContactByCustomerContactId.State = adStateOpen ) then
		if( not rsGetCustomerContactByCustomerContactId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intCustomerContactId = 0 ) then%>
   <tr><td class="sectiontitle" width=25%>&nbsp;&nbsp;New Contact</td>
   <%else%>
   <tr><td class="sectiontitle" width=25%>&nbsp;&nbsp;<%=rsGetCustomerContactByCustomerContactId( "Contact" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- CustomerContact Form -->
<tr><td align=center>
  <form id="frmCustomerContact" name="frmCustomerContact" action="" method="post">
		<input type="hidden" name="txtCustomerContactId" value="<%=intCustomerContactId%>">
		<input type="hidden" name="txtCustomerId" value="<%=intCustomerId%>">
    <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
			 <tr><td class="label_1">Contact</td><td class="input_1"><input type="text" class="forminput" name="txtContact" isRequired="true" fieldName="Contact" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "Contact" ))%>"></td>
		  <td class="label_2">Title</td><td class="input_2"><input type="text" class="forminput" name="txtContactTitle" fieldName="Title" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "ContactTitle" ))%>"></td></tr>
		<tr><td class="label_1">Dept.</td><td class="input_1"><input type="text" class="forminput" name="txtDepartment" fieldName="Department" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "Department" ))%>"></td>
		  <td class="label_2">E-mail</td><td class="input_2"><input type="text" class="forminput" name="txtEmailAddress" fieldName="E-mail" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "EmailAddress" ))%>"></td></tr>
		<tr><td class="label_1">Phone #</td><td class="input_1"><input type="text" class="forminput" name="txtPhoneNumber" fieldName="Phone" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "TelephoneNumber" ))%>"></td>
		  <td class="label_2">Extension</td><td class="input_2"><input type="text" class="forminput" name="txtExtension" fieldName="Extension" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "Extension" ))%>"></td></tr>
		<tr><td class="label_1">Mobile #</td><td class="input_1"><input type="text" class="forminput" name="txtMobileNumber" fieldName="Mobile" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "MobileNumber" ))%>"></td>
			<td class="label_2">Fax #</td><td class="input_2"><input type="text" class="forminput" name="txtFaxNumber" fieldName="Fax" onchange="this.isDirty=true;" maxlength=100 value="<%=EscapeDoubleQuotes(rsGetCustomerContactByCustomerContactId( "FaxNumber" ))%>"></td></tr>
		<tr><td class="label_1">Comments</td></tr>
		<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtNotes" fieldName="Notes" onchange="this.isDirty=true;" rows=4 wrap=soft ID="Textarea1"><%=rsGetCustomerContactByCustomerContactId( "Notes" )%></textarea></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmCustomerContact);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmCustomerContact.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmCustomerContact);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>  <!-- End CustomerContact Form -->
 </td></tr></table>
<%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetCustomerContactByCustomerContactId is nothing ) then
	if( rsGetCustomerContactByCustomerContactId.State = adStateOpen ) then
		rsGetCustomerContactByCustomerContactId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCustomerContactByCustomerContactId		= nothing
set cmdGetCustomerContactByCustomerContactId	= nothing
set cmdEditCustomerContact							= nothing
set cnnSQLConnection						= nothing
%>