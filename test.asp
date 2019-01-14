<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetCustomerByCustomerId, cmdGetCustomerByCustomerId, cmdEditCustomer
dim rsGetEmployeesList, rsGetEmployeesList2, cmdGetEmployeesList
dim intRowNumber, intCustomerId, strCurrentPage
dim rsGetProjects, cmdGetProjects
dim rsGetCustomerContacts, cmdGetCustomerContacts

set cnnSQLConnection						= nothing
set rsGetCustomerByCustomerId		= nothing
set cmdGetCustomerByCustomerId	= nothing
set cmdEditCustomer							= nothing
set rsGetEmployeesList						= nothing
set rsGetEmployeesList2						= nothing
set cmdGetEmployeesList					= nothing
set rsGetProjects			= nothing
set cmdGetProjects		= nothing
set rsGetCustomerContacts			= nothing
set cmdGetCustomerContacts		= nothing

set cmdGetProjects		= Server.CreateObject( "ADODB.Command" )
set cmdGetCustomerContacts		= Server.CreateObject( "ADODB.Command" )

set cnnSQLConnection						= Server.CreateObject( "ADODB.Connection" )
set cmdGetCustomerByCustomerId	= Server.CreateObject( "ADODB.Command" )
set cmdGetEmployeesList					= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
cmdGetEmployeesList.ActiveConnection = cnnSQLConnection
cmdGetEmployeesList.CommandText = "{call spGetEmployeeList}"

on error resume next
set rsGetEmployeesList = cmdGetEmployeesList.Execute
set rsGetEmployeesList2 = cmdGetEmployeesList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************

if( Request.Form.Count > 0 ) then
	if( Request.Form( "txtCustomerId" ) <> "" ) then
		set cmdEditCustomer	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditCustomer.ActiveConnection	= cnnSQLConnection
		cmdEditCustomer.CommandText				= "{? = call spEditCustomer( ?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Customer", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Address", adVarChar, adParamInput, 255 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@City", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@State", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Country", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Zip", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@SalesRepId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@DateEstablished", adVarChar, adParamInput, 100 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@MainContactId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Notes", adLongVarChar, adParamInput, 1048576 )
	
		if( Request.Form( "txtCustomerId" ) <> "" ) then cmdEditCustomer.Parameters( "@CustomerId" ) = Request.Form( "txtCustomerId" )
		if( Request.Form( "txtCustomer" ) <> "" ) then cmdEditCustomer.Parameters( "@Customer" ) = Request.Form( "txtCustomer" )
		if( Request.Form( "txtAddress" ) <> "" ) then cmdEditCustomer.Parameters( "@Address" ) = Request.Form( "txtAddress" )
		if( Request.Form( "txtCity" ) <> "" ) then cmdEditCustomer.Parameters( "@City" ) = Request.Form( "txtCity" )
		if( Request.Form( "txtState" ) <> "" ) then cmdEditCustomer.Parameters( "@State" ) = Request.Form( "txtState" )
		if( Request.Form( "txtCountry" ) <> "" ) then cmdEditCustomer.Parameters( "@Country" ) = Request.Form( "txtCountry" )
		if( Request.Form( "txtPostalCode" ) <> "" ) then cmdEditCustomer.Parameters( "@Zip" ) = Request.Form( "txtPostalCode" )
		if( Request.Form( "selSalesRep" ) <> "" ) then cmdEditCustomer.Parameters( "@SalesRepId" ) = Request.Form( "selSalesRep" )
		if( Request.Form( "txtDateEstablished" ) <> "" ) then cmdEditCustomer.Parameters( "@DateEstablished" ) = Request.Form( "txtDateEstablished" )
		if( Request.Form( "selMainContact" ) <> "" ) then cmdEditCustomer.Parameters( "@MainContactId" ) = Request.Form( "selMainContact" )
		if( Request.Form( "txtNotes" ) <> "" ) then cmdEditCustomer.Parameters( "@Notes" ) = Request.Form( "txtNotes" )
	
		on error resume next
		cmdEditCustomer.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
		
		on error goto 0
	
		intCustomerId = cmdEditCustomer.Parameters( "@Return" ).Value
		
	elseif( Request.Form( "txtCustomerIdDetail" ) <> "" ) then
	
		set cmdEditCustomer	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditCustomer.ActiveConnection	= cnnSQLConnection
		cmdEditCustomer.CommandText				= "{? = call spEditCall( ?,?,?,?,?,?,? ) }"
		
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )	
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@CallId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@EmployeeId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
		cmdEditCustomer.Parameters.Append cmdEditCustomer.CreateParameter( "@CallComments", adLongVarChar, adParamInput, 1048576 )
		
		if( Request.Form( "txtCustomerIdDetail" ) <> "" ) then cmdEditCustomer.Parameters( "@CustomerId" ) = Request.Form( "txtCustomerIdDetail" )
		if( Request.Form( "selEmployeeId" ) <> "" ) then cmdEditCustomer.Parameters( "@EmployeeId" ) = Request.Form( "selEmployeeId" )
		if( Request.Form( "selProjectId" ) <> "" ) then cmdEditCustomer.Parameters( "@ProjectId" ) = Request.Form( "selProjectId" )
		if( Request.Form( "txtCallComments" ) <> "" ) then cmdEditCustomer.Parameters( "@CallComments" ) = Request.Form( "txtCallComments" )
		
		on error resume next
		cmdEditCustomer.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
		
		on error goto 0
	
		intCustomerId = Request.Form( "txtCustomerIdDetail" )
		
	end if
	
else
	intCustomerId = Request.QueryString( "intCustomerId" )
end if


if intCustomerId > 0 then
	cmdGetProjects.ActiveConnection	= cnnSQLConnection

	cmdGetProjects.CommandText		= "{call spGetProjectsByCustomerId( ?,? ) }"

	cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetProjects.Parameters.Append cmdGetProjects.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
	
	cmdGetProjects.Parameters( "@CustomerId" ) = intCustomerId
	
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
	
	'********************************************************************************************
	
	cmdGetCustomerContacts.ActiveConnection	= cnnSQLConnection

	cmdGetCustomerContacts.CommandText		= "{call spGetCustomerContactsByCustomerId( ?,? ) }"

	cmdGetCustomerContacts.Parameters.Append cmdGetCustomerContacts.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetCustomerContacts.Parameters.Append cmdGetCustomerContacts.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
	
	cmdGetCustomerContacts.Parameters( "@CustomerId" ) = intCustomerId
	
	on error resume next
	set rsGetCustomerContacts = cmdGetCustomerContacts.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
		
	on error goto 0
end if



cmdGetCustomerByCustomerId.ActiveConnection = cnnSQLConnection
cmdGetCustomerByCustomerId.CommandText = "{call spGetCustomerByCustomerId( ?,? ) }"
cmdGetCustomerByCustomerId.Parameters.Append cmdGetCustomerByCustomerId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomerByCustomerId.Parameters.Append cmdGetCustomerByCustomerId.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
cmdGetCustomerByCustomerId.Parameters( "@CustomerId" )	= intCustomerId

on error resume next
set rsGetCustomerByCustomerId = cmdGetCustomerByCustomerId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
		
on error goto 0

if( Request.ServerVariables( "QUERY_STRING" ) <> "" ) then
	strCurrentPage = Request.ServerVariables( "URL" ) + "?" + Request.ServerVariables( "QUERY_STRING" )
else
	strCurrentPage = Request.ServerVariables( "URL" )
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Edit / Add Customer</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	var objEmployeeField;
	
	function ConfirmDelete( intCallId, strCallComments )
	{
		if( confirm( 'Are you sure you want to remove item ' + strCallComments + '?' ) )
		{
			OpenWindow( 'call_delete.asp?intCallId='+intCallId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function ConfirmContactDelete( intCustomerContactId, strContactName )
	{
		if( confirm( 'Are you sure you want to remove contact ' + strContactName + '?' ) )
		{
			OpenWindow( 'customercontact_delete.asp?intCustomerContactId='+intCustomerContactId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	function OpenEmployee( strEmployeeId )
	{
	  OpenWindow( 'Employee_details.asp?intEmployeeId='+strEmployeeId,'CentEmpl'+intEmployeeId,'width=640,height=160,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 160 )+',scrollbars=no,dependent=yes' );
	}
	
	
	function SelectEmployee( objFormElement )
	{
		if ( objFormElement.value == -1 )
		{
			OpenEmployee( 0 );
			objEmployeeField = objFormElement[1];
		}
	}
	
	function EmployeeSaved( intId, strName )
	{
		objEmployeeField.text = strName;
		objEmployeeField.value = intId;
	}
	
	function OpenCustomerContact( strCustomerContactId )
	{
	  OpenWindow( 'customercontact_details.asp?intCustomerContactId='+strCustomerContactId+'&intCustomerId=<%= intCustomerId %>','CentCust'+strCustomerContactId,'width=640,height=280,left='+GetCenterX( 640 )+',top='+GetCenterY( 280 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 280 )+',scrollbars=no,dependent=yes' );
	}
	
	function OpenProject( strProjectId )
	{
	  OpenWindow( 'project_details.asp?intProjectId='+strProjectId,'CentProj'+strProjectId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetCustomerByCustomerId is nothing ) then
	if( rsGetCustomerByCustomerId.State = adStateOpen ) then
		if( not rsGetCustomerByCustomerId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intCustomerId = 0 ) then%>
   <tr><td class="sectiontitle">New Customer</td>
   <%else%>
   <tr><td class="sectiontitle"><%=rsGetCustomerByCustomerId( "Customer" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Customer Form -->
<form id="frmCustomer" name="frmCustomer" action="" method="POST">
<input type="hidden" name="txtCustomerId" value="<%=intCustomerId%>">
<tr><td align=center>
		<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		<tr><td class="label_1">Company</td><td class="input_1"><input type="text" class="forminput" name="txtCustomer" maxlength=100 isRequired="true" fieldName="Company" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "Customer" ))%>"></td>
		  <td class="label_2">Address</td><td class="input_2"><input type="text" class="forminput" name="txtAddress" maxlength=255 fieldName="Address" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "Address" ))%>"></td></tr>
		<tr><td class="label_1">City</td><td class="input_1"><input type="text" class="forminput" name="txtCity" maxlength=100 fieldName="City" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "City" ))%>"></td>
		  <td class="label_2">State</td><td class="input_2"><input type="text" class="forminput" name="txtState" maxlength=100 fieldName="State" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "State" ))%>"></td></tr>
		<tr><td class="label_1">Country</td><td class="input_1"><input type="text" class="forminput" name="txtCountry" maxlength=100 fieldName="Country" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "Country" ))%>"></td>
		  <td class="label_2">Postal Code</td><td class="input_2"><input type="text" class="forminput" name="txtPostalCode" maxlength=100 fieldName="Postal Code" onchange="this.isDirty=true;" fieldName="Postal Code" value="<%=EscapeDoubleQuotes(rsGetCustomerByCustomerId( "Zip" ))%>"></td></tr>
		<tr><td class="label_1">Sales Rep</td><td class="input_1"><select class="forminput" name="selSalesRep" size=1 isRequired="true" fieldName="Sales Rep" onchange="this.isDirty=true;" ID="Select1">
				<option value=""></option>
   				<%if not( rsGetEmployeesList is nothing ) then
   					if( rsGetEmployeesList.State = adStateOpen ) then
   						while( not rsGetEmployeesList.EOF )%><option value="<%=rsGetEmployeesList( "EmployeeId" )%>"<%if( rsGetEmployeesList( "EmployeeId" ) = rsGetCustomerByCustomerId( "SalesRepId" ) ) then%>SELECTED<%end if%>><%=rsGetEmployeesList( "EmployeeName" )%></option>
   							<%rsGetEmployeesList.MoveNext
   						wend
   					end if
   				end if%>
   	    </select></td>
   	    <td class="label_2">Date Est.</td><td class="input_2"><input type="text" class="forminput" name="txtDateEstablished" maxlength=100 isDate="true" fieldName="Date Est." onchange="this.isDirty=true;" value="<%=rsGetCustomerByCustomerId( "DateEstablished" )%>"></td></tr>
   	<tr><td class="label_1">Main Contact</td><td class="input_1"><select class="forminput" name="selMainContact" size=1 fieldName="Main Contact" onchange="this.isDirty=true;">
				<option value=""></option>
   				<%if not( rsGetCustomerContacts is nothing ) then
   					if( rsGetCustomerContacts.State = adStateOpen ) then
   						while( not rsGetCustomerContacts.EOF )%><option value="<%=rsGetCustomerContacts( "CustomerContactId" )%>"<%if( rsGetCustomerContacts( "CustomerContactId" ) = rsGetCustomerByCustomerId( "MainContactId" ) ) then%>SELECTED<%end if%>><%=rsGetCustomerContacts( "Contact" )%></option>
   							<%rsGetCustomerContacts.MoveNext
   						wend
   						if not ( rsGetCustomerContacts.EOF and rsGetCustomerContacts.BOF) then
								rsGetCustomerContacts.MoveFirst
							end if
   					end if
   				end if%>
   	    </select></td>
   	    <td>&nbsp;</td>
   	<tr><td class="label_1">Comments</td></tr>
		<tr><td class="input_1" colspan=4><textarea class="forminput" name="txtNotes" fieldName="Notes" onchange="this.isDirty=true;" rows="10" wrap=soft><%=rsGetCustomerByCustomerId( "Notes" )%></textarea></td></tr>
	</table>
</td></tr>

<!-- Bottom Rule -->
<tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
</td></tr>
<!-- End Bottom Rule -->
 
<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmCustomer);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmCustomer.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmCustomer)"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>  <!-- End Form Buttons -->
</form>  <!-- End Customer Form -->
<%	set rsGetCustomerByCustomerId = rsGetCustomerByCustomerId.NextRecordset
		end if
	end if
end if%>

<%if( intCustomerId <> 0 ) then%> 
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Contacts</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

 <!-- Contacts Table -->
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="20" align="center">Del</th>
    <th class="itemlist">Contact</th>
    <th class="itemlist">Title</th>
    <th class="itemlist">Dept</th>
    <th class="itemlist">Phone</th>
    <th class="itemlist">Ext</th>
  </tr>
  
  <!-- table data -->
  <% if ( not ( rsGetCustomerContacts is nothing ) ) then %>
  <%if( rsGetCustomerContacts.State = adStateOpen ) then
				while( not rsGetCustomerContacts.EOF )
					if( ( CInt( rsGetCustomerContacts( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>
						<tr class="oddrow">
					<%end if%>
						  <td class="itemlist" width="20" align=middle><a href="javascript:ConfirmContactDelete('<%=rsGetCustomerContacts( "CustomerContactId" )%>', '<%=EscapeQuotes(rsGetCustomerContacts( "Contact" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist"><a href="javascript:OpenCustomerContact( '<%=rsGetCustomerContacts( "CustomerContactId" )%>' );"><%=rsGetCustomerContacts( "Contact" )%></a></td>
						  <td class="itemlist"><%=rsGetCustomerContacts( "ContactTitle" )%></td>
						  <td class="itemlist"><%=rsGetCustomerContacts( "Department" )%></td>
						  <td class="itemlist"><%=rsGetCustomerContacts( "TelephoneNumber" )%></td>
						  <td class="itemlist"><%=rsGetCustomerContacts( "Extension" )%></td>
						</tr>
				<%rsGetCustomerContacts.MoveNext
				wend				
   				if not ( rsGetCustomerContacts.EOF and rsGetCustomerContacts.BOF) then
						rsGetCustomerContacts.MoveFirst
					end if
		end if
		end if%>
  
 </table></td></tr>
 <!-- End Projects Table -->
 
 <tr><td align=center>
  <table width=128 border=0 cellspacing=0 cellpadding=0 align=right>
    <tr><td width=64 align=right><a href="javascript:OpenCustomerContact(0);"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
      <td width=64 align=right><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  </table></td></tr>
 </form>
 </table></td></tr>

 
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0 ID="Table2">
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Call Log</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

 <!-- Call Log Table -->
 <form id="frmCallLog" name="frmCallLog" action="" method="post">
 <input type="hidden" name="txtCustomerIdDetail" id="txtCustomerIdDetail" value="<%=intCustomerId%>">
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="20" align="center">Del</th>
    <th class="itemlist" width="115">Date</th>
    <th class="itemlist" width="70">Project</th>
    <th class="itemlist" width="35">W.O.</th>
    <th class="itemlist">Comments</th>
    <th class="itemlist" width="100">Employee</th>
  </tr>
  
  <!-- table data -->
  <%if not( rsGetCustomerByCustomerId is nothing ) then
			if( rsGetCustomerByCustomerId.State = adStateOpen ) then
				intRowNumber = 1
		
				while( not rsGetCustomerByCustomerId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>
						<tr class="oddrow">
					<%end if%>
							<td class="itemlist" width="20" align=middle><a href="javascript:ConfirmDelete('<%=rsGetCustomerByCustomerId( "CallId" )%>', '<%=EscapeQuotes(rsGetCustomerByCustomerId( "CallComments" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" align=middle width="115"><%=rsGetCustomerByCustomerId( "CallDate" )%></td>
							<td class="itemlist" width="70"><a href="javascript:OpenProject('<%=rsGetCustomerByCustomerId( "ProjectId" )%>')"><%=rsGetCustomerByCustomerId( "ProjectName" )%></a></td>
							<td class="itemlist" width="35"><a href="javascript:OpenWorkOrder('<%=rsGetCustomerByCustomerId( "WorkOrderId" )%>')"><%=rsGetCustomerByCustomerId( "WorkOrderNumber" )%></a></td>
							<td class="itemlist"><%=rsGetCustomerByCustomerId( "CallComments" )%></td>
							<td class="itemlist" width="100"><%=rsGetCustomerByCustomerId( "EmployeeName" )%></td>
						</tr>
				<%intRowNumber = intRowNumber + 1
					rsGetCustomerByCustomerId.MoveNext
				wend
			end if
		end if%>
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td class="itemlist" width=20 align=middle>&nbsp;</td>
			<td class="itemlist" width=120 align=middle><%=FormatDateTime( Now(), vbGeneralDate )%></td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td><input type="text" class="forminput" name="txtCallComments" isRequired="true" fieldName="Comments" value=""></td>
			<td><select class="forminput" name="selEmployeeId" size=1 isRequired="true" fieldName="Employee">
					<option value=""></option>
					<%if not( rsGetEmployeesList2 is nothing ) then
   						if( rsGetEmployeesList2.State = adStateOpen ) then
   							while( not rsGetEmployeesList2.EOF )
   							%><option value="<%=rsGetEmployeesList2( "EmployeeId" )%>"><%=rsGetEmployeesList2( "EmployeeName" )%></option>
   								<%rsGetEmployeesList2.MoveNext
   							wend
   						end if
   					end if
   				%>
					</select></td>
		</tr>
 </table></td></tr>
 <!-- End Call Log Table -->
 
 <tr><td align=center>
  <table width=128 border=0 cellspacing=0 cellpadding=0 align=right>
    <tr><td width=64 align=right><a href="javascript:validateForm(frmCallLog);"><IMG src="images/add_button.gif" width=60 height=20 border=0></a></td>
      <td width=64 align=right><a href="javascript:RefreshCurrentPage();"><IMG src="images/refresh_button.gif" width=60 height=20 border=0></a></td></tr>
  </table></td></tr>
 </form>
 </table></td></tr>

<%end if%>
 
 
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Projects</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

 <!-- Projects Table -->
 <tr><td align=center>
  <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
  <tr>
  <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
  <!-- table heading -->
  <tr>
    <th class="itemlist" width="90">Name</th>
    <th class="itemlist" width="60">Type</th>
    <th class="itemlist">Description</th>
    <th class="itemlist" width="60">Start Date</th>
    <th class="itemlist" width="60">Finish Date</th>
  </tr>
  
  <!-- table data -->
  <% if ( not ( rsGetProjects is nothing ) ) then %>
  <%if( rsGetProjects.State = adStateOpen ) then
				while( not rsGetProjects.EOF )
					if( ( CInt( rsGetProjects( "RowNumber" ) ) Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>
						<tr class="oddrow">
					<%end if%>
						  <td class="itemlist" width="90"><a href="javascript:OpenProject( '<%=rsGetProjects( "ProjectId" )%>' );"><%=rsGetProjects( "ProjectName" )%></a></td>
						  <td class="itemlist" width="60"><%=rsGetProjects( "ProjectType" )%></td>
						  <td class="itemlist"><%=rsGetProjects( "ProjectDescription" )%></td>
						  <td class="itemlist" width="60" align=right><%=rsGetProjects( "StartDate" )%></td>
						  <td class="itemlist" width="60" align=right><%=rsGetProjects( "CompletionDate" )%></td>
						</tr>
				<%rsGetProjects.MoveNext
				wend
		end if
		end if%>
		
   </table></td></tr>
 </table></td></tr>
 <!-- End Projects Table -->

</body>
</html>
<%
if not( rsGetCustomerByCustomerId is nothing ) then
	if( rsGetCustomerByCustomerId.State = adStateOpen ) then
		rsGetCustomerByCustomerId.Close
	end if
end if

if not( rsGetEmployeesList is nothing ) then
	if( rsGetEmployeesList.State = adStateOpen ) then
		rsGetEmployeesList.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCustomerByCustomerId		= nothing
set cmdGetCustomerByCustomerId	= nothing
set rsGetEmployeesList						= nothing
set cmdGetEmployeesList					= nothing
set cmdEditCustomer							= nothing
set cnnSQLConnection						= nothing
%>
