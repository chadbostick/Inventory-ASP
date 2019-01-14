<%option explicit

dim cnnSQLConnection
dim rsGetCompanyInformation, cmdGetCompanyInformation, cmdEditCompanyInformation
dim intRowNumber, intCompanyInformationId

set cnnSQLConnection					= nothing
set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set cmdEditCompanyInformation	= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetCompanyInformation	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then

	set cmdEditCompanyInformation	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditCompanyInformation.ActiveConnection	= cnnSQLConnection
	cmdEditCompanyInformation.CommandText				= "{? = call spEditCompanyInformation( ?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@CompanyName", adVarChar, adParamInput, 50 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@Address", adVarChar, adParamInput, 255 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@City", adVarChar, adParamInput, 50 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@StateOrProvince", adVarChar, adParamInput, 20 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@PostalCode", adVarChar, adParamInput, 20 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@Country", adVarChar, adParamInput, 50 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@PhoneNumber", adVarChar, adParamInput, 30 )
	cmdEditCompanyInformation.Parameters.Append cmdEditCompanyInformation.CreateParameter( "@FaxNumber", adVarChar, adParamInput, 30 )
	
	if( Request.Form( "txtCompanyName" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@CompanyName" ) = Request.Form( "txtCompanyName" )
	if( Request.Form( "txtAddress" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@Address" ) = Request.Form( "txtAddress" )
	if( Request.Form( "txtCity" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@City" ) = Request.Form( "txtCity" )
	if( Request.Form( "txtStateProvince" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@StateOrProvince" ) = Request.Form( "txtStateProvince" )
	if( Request.Form( "txtPostalCode" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@PostalCode" ) = Request.Form( "txtPostalCode" )
	if( Request.Form( "txtCountry" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@Country" ) = Request.Form( "txtCountry" )
	if( Request.Form( "txtPhone" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@PhoneNumber" ) = Request.Form( "txtPhone" )
	if( Request.Form( "txtFax" ) <> "" ) then cmdEditCompanyInformation.Parameters( "@FaxNumber" ) = Request.Form( "txtFax" )
	
	on error resume next
	cmdEditCompanyInformation.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0
end if

cmdGetCompanyInformation.ActiveConnection = cnnSQLConnection
cmdGetCompanyInformation.CommandText = "{call spGetCompanyInformation }"

on error resume next
set rsGetCompanyInformation = cmdGetCompanyInformation.Execute
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
<TITLE>Centerline - View / Edit Company Info</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetCompanyInformation is nothing ) then
		if( rsGetCompanyInformation.State = adStateOpen ) then
			if( not rsGetCompanyInformation.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=35%>Company Information</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Centerline Form -->
<form id="frmCompanyInfo" name="frmCompanyInfo" action="" method="post">
<tr><td align=center>
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Name</td><td class="input_1"><input type="text" class="forminput" name="txtCompanyName" onchange="this.isDirty=true;" fieldName="Name" maxlength=50 value="<%=rsGetCompanyInformation( "CompanyName" )%>"></td>
       <td class="label_2">Postal Code</td><td class="input_2"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="Postal Code" name="txtPostalCode" maxlength=20 value="<%=rsGetCompanyInformation( "PostalCode" )%>"></td></tr>
     <tr><td class="label_1">Address</td><td class="input_1"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="Address" name="txtAddress" maxlength=255 value="<%=rsGetCompanyInformation( "Address" )%>"></td>
 	  <td class="label_2">Country</td><td class="input_2"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="Country" name="txtCountry" maxlength=50 value="<%=rsGetCompanyInformation( "Country" )%>"></td></tr>
     <tr><td class="label_1">City</td><td class="input_1"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="City" name="txtCity" maxlength=50 value="<%=rsGetCompanyInformation( "City" )%>"></td>
       <td class="label_2">Phone</td><td class="input_2"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="Phone" name="txtPhone" maxlength=30 value="<%=rsGetCompanyInformation( "PhoneNumber" )%>"></td></tr>
     <tr><td class="label_1">State</td><td class="input_1"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="State" name="txtStateProvince" maxlength=20 value="<%=rsGetCompanyInformation( "StateOrProvince" )%>"></td>
       <td class="label_2">Fax</td><td class="input_2"><input type="text" class="forminput" onchange="this.isDirty=true;" fieldName="Fax" name="txtFax" maxlength=30 value="<%=rsGetCompanyInformation( "FaxNumber" )%>"></td></tr>
  </table>
</td></tr>
<!-- End Customer Form -->
	    
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
          <tr><td width=64 align=right><a href="javascript:frmCompanyInfo.submit();"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmCompanyInfo.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:GetDirty(frmCompanyInfo);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
<%	end if
	end if
end if%>
</BODY>
</HTML>
<%
if not( rsGetCompanyInformation is nothing ) then
	if( rsGetCompanyInformation.State = adStateOpen ) then
		rsGetCompanyInformation.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set cmdEditCompanyInformation	= nothing
set cnnSQLConnection					= nothing
%>