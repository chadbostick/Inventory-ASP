<%option explicit
%><!--#include file="global_functions.asp"--><%

dim cnnSQLConnection
dim rsGetCustomersList, cmdGetCustomersList
dim rsGetEmployeesList, cmdGetEmployeesList
dim rsGetProjectTypesList, cmdGetProjectTypesList
dim rsGetProjectPrioritiesList, cmdGetProjectPrioritiesList
dim rsGetProjectByProjectId, cmdGetProjectByProjectId
dim rsGetCustomerContactsByCustomerId, cmdGetCustomerContactsByCustomerId
dim cmdEditProject
dim intProjectId, intRowNumber, strCurrentPage, intCustomerId, intSpindleId

set cnnSQLConnection								= nothing
set rsGetCustomersList							= nothing
set cmdGetCustomersList							= nothing
set rsGetEmployeesList							= nothing
set cmdGetEmployeesList							= nothing
set rsGetProjectTypesList						= nothing
set cmdGetProjectTypesList					= nothing
set rsGetProjectPrioritiesList			= nothing
set cmdGetProjectPrioritiesList			= nothing
set rsGetCustomerContactsByCustomerId			= nothing
set cmdGetCustomerContactsByCustomerId			= nothing
set rsGetProjectByProjectId					= nothing
set cmdGetProjectByProjectId				= nothing
set cmdEditProject									= nothing

set cnnSQLConnection								= Server.CreateObject( "ADODB.Connection" )
set cmdGetCustomersList							= Server.CreateObject( "ADODB.Command" )
set cmdGetEmployeesList							= Server.CreateObject( "ADODB.Command" )
set cmdGetCustomerContactsByCustomerId								= Server.CreateObject( "ADODB.Command" )
set cmdGetProjectByProjectId				= Server.CreateObject( "ADODB.Command" )
set cmdGetProjectTypesList					= Server.CreateObject( "ADODB.Command" )
set cmdGetProjectPrioritiesList			= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

'**********************************************************************************
cmdGetCustomersList.ActiveConnection = cnnSQLConnection
cmdGetCustomersList.CommandText = "{call spGetCustomerList}"

on error resume next
set rsGetCustomersList = cmdGetCustomersList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetEmployeesList.ActiveConnection = cnnSQLConnection
cmdGetEmployeesList.CommandText = "{call spGetEmployeeList}"

on error resume next
set rsGetEmployeesList = cmdGetEmployeesList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetProjectTypesList.ActiveConnection = cnnSQLConnection
cmdGetProjectTypesList.CommandText = "{call spGetProjectTypeList}"

on error resume next
set rsGetProjectTypesList = cmdGetProjectTypesList.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
							
on error goto 0
'**********************************************************************************
cmdGetProjectPrioritiesList.ActiveConnection = cnnSQLConnection
cmdGetProjectPrioritiesList.CommandText = "{call spGetProjectPriorityList}"

on error resume next
set rsGetProjectPrioritiesList = cmdGetProjectPrioritiesList.Execute
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
	if( Request.Form( "txtProjectId" ) <> "" ) then
		set cmdEditProject	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditProject.ActiveConnection	= cnnSQLConnection
		cmdEditProject.CommandText			= "{? = call spEditProject( ?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectName", adVarChar, adParamInput, 100 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectTypeId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CustomerPO", adVarChar, adParamInput, 100 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectDescription", adLongVarChar, adParamInput, 1048576 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@StartDate", adVarChar, adParamInput, 20 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@EstCompletionDate", adVarChar, adParamInput, 20 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CompletionDate", adVarChar, adParamInput, 20 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@SpindleId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectContactId", adInteger, adParamInput, 4 )
	
		if( Request.Form( "txtProjectId" ) <> "" ) then cmdEditProject.Parameters( "@ProjectId" ) = Request.Form( "txtProjectId" )
		if( Request.Form( "txtProjectName" ) <> "" ) then cmdEditProject.Parameters( "@ProjectName" ) = Request.Form( "txtProjectName" )
		if( Request.Form( "selProjectType" ) <> "" ) then cmdEditProject.Parameters( "@ProjectTypeId" ) = Request.Form( "selProjectType" )
		if( Request.Form( "hidCustomerId" ) <> "" ) then cmdEditProject.Parameters( "@CustomerId" ) = Request.Form( "hidCustomerId" )
		if( Request.Form( "txtCustomerPO" ) <> "" ) then cmdEditProject.Parameters( "@CustomerPO" ) = Request.Form( "txtCustomerPO" )
		if( Request.Form( "txtDescription" ) <> "" ) then cmdEditProject.Parameters( "@ProjectDescription" ) = Request.Form( "txtDescription" )
		if( Request.Form( "txtStartDate" ) <> "" ) then cmdEditProject.Parameters( "@StartDate" ) = Request.Form( "txtStartDate" )
		if( Request.Form( "txtEstComplete" ) <> "" ) then cmdEditProject.Parameters( "@EstCompletionDate" ) = Request.Form( "txtEstComplete" )
		if( Request.Form( "txtCompletionDate" ) <> "" ) then cmdEditProject.Parameters( "@CompletionDate" ) = Request.Form( "txtCompletionDate" )
		if( Request.Form( "hidSpindleId" ) <> "" ) then cmdEditProject.Parameters( "@SpindleId" ) = Request.Form( "hidSpindleId" )
		if( Request.Form( "selProjectContact" ) <> "" ) then cmdEditProject.Parameters( "@ProjectContactId" ) = Request.Form( "selProjectContact" )
		
		on error resume next
		cmdEditProject.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
									
		on error goto 0
	
		intProjectId = cmdEditProject.Parameters( "@Return" ).Value
		
	elseif( Request.Form( "txtProjectIdDetail" ) <> "" ) then
	
		set cmdEditProject	= Server.CreateObject( "ADODB.Command" )
	
		cmdEditProject.ActiveConnection	= cnnSQLConnection
		cmdEditProject.CommandText				= "{? = call spEditCall( ?,?,?,?,?,?,? ) }"
		
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )	
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CallId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@EmployeeId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@WorkOrderId", adInteger, adParamInput, 4 )
		cmdEditProject.Parameters.Append cmdEditProject.CreateParameter( "@CallComments", adLongVarChar, adParamInput, 1048576 )
		
		if( Request.Form( "txtCustomerIdDetail" ) <> "" ) then cmdEditProject.Parameters( "@CustomerId" ) = Request.Form( "txtCustomerIdDetail" )
		if( Request.Form( "selEmployeeId" ) <> "" ) then cmdEditProject.Parameters( "@EmployeeId" ) = Request.Form( "selEmployeeId" )
		if( Request.Form( "txtProjectIdDetail" ) <> "" ) then cmdEditProject.Parameters( "@ProjectId" ) = Request.Form( "txtProjectIdDetail" )
		if( Request.Form( "txtCallComments" ) <> "" ) then cmdEditProject.Parameters( "@CallComments" ) = Request.Form( "txtCallComments" )
		
		on error resume next
		cmdEditProject.Execute
		if( Err.number <> 0 ) then
			Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
			Response.Write( Err.Description )
			Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
			Response.Write( "</font></body></html>" )
			Response.End
		end if
									
		on error goto 0
	
		intProjectId = Request.Form( "txtProjectIdDetail" )
		
	end if
		
else
	if( Request.QueryString( "intProjectId" ) = "" ) then
		intProjectId = 0
	else
		intProjectId = Request.QueryString( "intProjectId" )
	end if
end if

cmdGetProjectByProjectId.ActiveConnection = cnnSQLConnection
cmdGetProjectByProjectId.CommandText = "{call spGetProjectByProjectId( ?,?,?,? ) }"
cmdGetProjectByProjectId.Parameters.Append cmdGetProjectByProjectId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProjectByProjectId.Parameters.Append cmdGetProjectByProjectId.CreateParameter( "@ProjectId", adInteger, adParamInput, 4 )
cmdGetProjectByProjectId.Parameters.Append cmdGetProjectByProjectId.CreateParameter( "@CustomerId", adInteger, adParamOutput, 4 )
cmdGetProjectByProjectId.Parameters.Append cmdGetProjectByProjectId.CreateParameter( "@SpindleId", adInteger, adParamOutput, 4 )
cmdGetProjectByProjectId.Parameters( "@ProjectId" )	= intProjectId

on error resume next
set rsGetProjectByProjectId = cmdGetProjectByProjectId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
									
on error goto 0

intSpindleId = cmdGetProjectByProjectId.Parameters( "@SpindleId" )


'**********************************************************************************
cmdGetCustomerContactsByCustomerId.ActiveConnection = cnnSQLConnection
cmdGetCustomerContactsByCustomerId.CommandText = "{call spGetCustomerContactsByCustomerId ( ?,? ) }"
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetCustomerContactsByCustomerId.Parameters.Append cmdGetCustomerContactsByCustomerId.CreateParameter( "@CustomerId", adInteger, adParamInput, 4 )
cmdGetCustomerContactsByCustomerId.Parameters( "@CustomerId" )	= cmdGetProjectByProjectId.Parameters( "@CustomerId" )

on error resume next
set rsGetCustomerContactsByCustomerId = cmdGetCustomerContactsByCustomerId.Execute
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
'**********************************************************************************
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - View / Edit Project</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	var objProjectTypeField;
	var objProjectPriorityField;
	
	function ConfirmDelete( intCallId, strCallComments )
	{
		if( confirm( 'Are you sure you want to remove item ' + strCallComments + '?' ) )
		{
			OpenWindow( 'call_delete.asp?intCallId='+intCallId,'DeleteRecord','width=657,height=120,left='+GetCenterX( 657 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 120 )+',scrollbars=yes,dependent=yes' );
		}
	}
	
	function OpenWorkOrder( strWorkOrderId )
	{
	  OpenWindow( 'workorder_details.asp?intWorkOrderId='+strWorkOrderId,'CenterlineWorkOrder'+strWorkOrderId,'width=657,height=480,left='+GetCenterX( 657 )+',top='+GetCenterY( 480 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 480 )+',scrollbars=yes,dependent=yes' );
	}
	
	function OpenCustomer( strCustomerId )
	{
	  OpenWindow( 'customer_details.asp?intCustomerId='+strCustomerId,'CenterlineCustomer'+strCustomerId,'width=657,height=550,left='+GetCenterX( 657 )+',top='+GetCenterY( 550 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 550 )+',scrollbars=yes,dependent=yes' );
	}
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	
	function SelectProjectType( objFormElement )
	{
		if ( objFormElement.value == -1 )
		{
			OpenProjectType( 0 );
			objProjectTypeField = objFormElement[1];
		}
		objFormElement.isDirty=true;
	}	
	function ProjectTypeSaved( intId, strName )
	{
		objProjectTypeField.text = strName;
		objProjectTypeField.value = intId;
	}	
	function OpenProjectType( strProjectTypeId )
	{
	  OpenWindow( 'projecttype_details.asp?intProjectTypeId='+strProjectTypeId,'Centerline','width=640,height=120,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 160 )+',screenY='+GetCenterY( 480 )+',scrollbars=no,dependent=yes' );
	}
	
	
	function SelectProjectPriority( objFormElement )
	{
		if ( objFormElement.value == -1 )
		{
			OpenProjectPriority( 0 );
			objProjectPriorityField = objFormElement[1];
		}
	}	
	function ProjectPrioritySaved( intId, strName )
	{
		objProjectPriorityField.text = strName;
		objProjectPriorityField.value = intId;
	}	
	function OpenProjectPriority( strProjectPriorityId )
	{
	  OpenWindow( 'projectpriority_details.asp?intProjectPriorityId='+strProjectPriorityId,'Centerline','width=640,height=120,left='+GetCenterX( 640 )+',top='+GetCenterY( 120 )+',screenX='+GetCenterX( 160 )+',screenY='+GetCenterY( 120 )+',scrollbars=no,dependent=yes' );
	}
		
	function OpenCustomerSearch(  )
	{
	  OpenWindow( 'customer_search.asp','CustomerSearch','width=691,height=500,left='+GetCenterX( 691 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 691 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectCustomer( intCustomerId, strCustomer )
	{
		document.forms.frmProject.hidCustomerId.value = intCustomerId;
		document.forms.frmProject.txtCustomer.value = strCustomer;
	}
		
	function OpenSpindleSearch(  )
	{
	  OpenWindow( 'spindle_search.asp','SpindleSearch','width=691,height=500,left='+GetCenterX( 691 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 691 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}
	function SelectSpindle( intSpindleId, strSpindleType )
	{
		document.forms.frmProject.hidSpindleId.value = intSpindleId;
		document.forms.frmProject.txtSpindleType.value = strSpindleType;
	}
	
  function OpenSpindleParts( strSpindleId )
	{
	  OpenWindow( 'spindle_parts_details.asp?intSpindleId='+strSpindleId,'EditSpindleParts','width=722,height=500,left='+GetCenterX( 722 )+',top='+GetCenterY( 500 )+',screenX='+GetCenterX( 722 )+',screenY='+GetCenterY( 500 )+',scrollbars=yes,dependent=yes' );
	}	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0>
<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%if not( rsGetProjectByProjectId is nothing ) then
		if( rsGetProjectByProjectId.State = adStateOpen ) then
			if( not rsGetProjectByProjectId.EOF ) then
				intCustomerId = rsGetProjectByProjectId( "CustomerId" )%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=4><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intProjectId = 0 ) then%>
   <tr><td class="sectiontitle" width=20%>New Project</td>
   <%else%>
   <tr><td class="sectiontitle" width=20%>Project&nbsp;<%=intProjectId%></td>
   <%end if%>
		<td align=left width=23><IMG SRC="images/title_right.gif" width=23 height=25></td>
		<td class="sectioncaption" align=left width=325><a href="javascript:OpenCustomer(<%=rsGetProjectByProjectId( "CustomerId" )%>)"><%=rsGetProjectByProjectId( "Customer" )%></a></td>
		<td class="sectioncaption" align=right><%=FormatDateTime( Now(), vbShortDate )%></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Project Form -->
<form id="frmProject" name="frmProject" action="" method="post">
<input type="hidden" name="txtProjectId" value="<%=intProjectId%>">
<tr><td align=center>
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1">Project Name</td><td class="input_1" colspan=3><input type="text" class="forminput" name="txtProjectName" maxlength=50 isRequired="true" onchange="this.isDirty=true;" fieldName="Project Name" value="<%=EscapeDoubleQuotes(rsGetProjectByProjectId( "ProjectName" ))%>"></td></tr>
   	<tr><td class="label_1">Project Type</td><td class="input_1"><select class="forminput" name="selProjectType" size=1 isRequired="true" fieldName="Project Type" onChange="javascript:SelectProjectType(this);">
   			<option value=""></option>
   			<option value="-1">- ADD NEW PROJECT TYPE -</option>
   	    <%if not( rsGetProjectTypesList is nothing ) then
   					if( rsGetProjectTypesList.State = adStateOpen ) then
   						while( not rsGetProjectTypesList.EOF )%><option value="<%=rsGetProjectTypesList( "ProjectTypeId" )%>"<%if( rsGetProjectTypesList( "ProjectTypeId" ) = rsGetProjectByProjectId( "ProjectTypeId" ) ) then%>SELECTED<%end if%>><%=rsGetProjectTypesList( "ProjectType" )%></option>
   							<%rsGetProjectTypesList.MoveNext
   						wend
   					end if
   				end if%>
   	  </select></td>
       <td class="label_2">Sales Rep</td><td class="input_2"><input type="text" class="forminput" name="txtSalesRep" onchange="this.isDirty=true;" fieldName="Sales Rep" readonly value="<%=EscapeDoubleQuotes(rsGetProjectByProjectId( "EmployeeName" ))%>"></td></tr>
   	<tr><td class="label_1">Customer</td><td class="input_1"><table cellpadding=0 cellspacing=0 width=100%><tr><td><a href="javascript:OpenCustomerSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidCustomerId" value="<%=rsGetProjectByProjectId( "CustomerId" )%>"><input class="forminput" name="txtCustomer" readonly isRequired="true" fieldName="Customer" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetProjectByProjectId( "Customer" ))%>"></td></tr></table></td>
   	  <td class="label_2">Contact</td><td class="input_2"><select class="forminput" name="selProjectContact" size=1 fieldName="Project Contact" onChange="javascript:this.isDirty=true;">
   			<option value=""></option>
   	    <%if not( rsGetCustomerContactsByCustomerId is nothing ) then
   					if( rsGetCustomerContactsByCustomerId.State = adStateOpen ) then
   						while( not rsGetCustomerContactsByCustomerId.EOF )%><option value="<%=rsGetCustomerContactsByCustomerId( "CustomerContactId" )%>"<%if( rsGetCustomerContactsByCustomerId( "CustomerContactId" ) = rsGetProjectByProjectId( "ProjectContactId" ) ) then%>SELECTED<%end if%>><%=rsGetCustomerContactsByCustomerId( "Contact" )%></option>
   							<%rsGetCustomerContactsByCustomerId.MoveNext
   						wend
   					end if
   				end if%>
   	  </select></td></tr>
   	<tr><td colspan=2>&nbsp;</td>
   	  <td class="label_2">Project PO</td><td class="input_2"><input type="text" class="forminput" name="txtCustomerPO" onchange="this.isDirty=true;" fieldName="Project PO" maxlength=25 value="<%=rsGetProjectByProjectId( "CustomerPO" )%>"></td></tr>
   	<tr><td class="label_1">Spindle Type</td><td class="input_1" colspan=3><table cellpadding=0 cellspacing=0 width=100%><tr><td><a href="javascript:OpenSpindleSearch();"><IMG src="images/select_button.gif" width=60 height=20 border=0></a></td><td width=100%><input type="hidden" name="hidSpindleId" value="<%=rsGetProjectByProjectId( "SpindleId" )%>"><input class="forminput" name="txtSpindleType" readonly isRequired="true" fieldName="Spindle Type" onchange="this.isDirty=true;" value="<%=EscapeDoubleQuotes(rsGetProjectByProjectId( "SpindleType" ))%>"></td></tr></table></td></tr>
   	<% if intProjectId > 0 then %><tr><td class="label_1">Parts</td><td class="generic_input" colspan=3><a href="javascript:OpenSpindleParts(<%=intSpindleId%>)">Edit Spindle Parts</a></td></tr><% end if %>
 </table>
</td></tr>

 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>
 <!-- End Middle Title Bar -->
 
<tr><td align=center>
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
		 <%if( intProjectId = 0 ) then%>
   	 <tr><td class="label_1">Start Date</td><td class="input_1"><input type="text" class="forminput" name="txtStartDate" onchange="this.isDirty=true;" maxlength=10 isDate="true" fieldName="Start Date" value="<%=FormatDateTime( Now(), vbShortDate )%>"></td>
   	 <%else%>
   	 <tr><td class="label_1">Start Date</td><td class="input_1"><input type="text" class="forminput" name="txtStartDate" onchange="this.isDirty=true;" maxlength=10 isDate="true" fieldName="Start Date" value="<%=rsGetProjectByProjectId( "StartDate" )%>"></td>
   	 <%end if%>
       <td class="label_2">Est. Complete</td><td class="input_2"><input type="text" class="forminput" name="txtEstComplete" onchange="this.isDirty=true;" maxlength=10 isDate="true" fieldName="Est. Complete" value="<%=rsGetProjectByProjectId( "EstCompletionDate" )%>"></td></tr>
     <tr>
     <td class="label_1" colspan=2>&nbsp;</td>
   	  <td class="label_2">Finish Date</td><td class="input_2"><input type="text" class="forminput" name="txtCompletionDate" maxlength=10 isDate="true" onchange="this.isDirty=true;" fieldName="Finish Date" value="<%=rsGetProjectByProjectId( "CompletionDate" )%>"></td></tr>
 </table>
</td></tr>

 <!-- Middle Title Bar -->
 <tr><td class="pagesmalltitle">
  <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=2></td></tr>
  </TABLE>
 </td></tr>
 <!-- End Middle Title Bar -->
 
<tr><td align=center>
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
     <tr><td class="label_1" colspan=4>Project Description</td></tr>
     <tr><td class="input_1" colspan=4><textarea class="forminput custext" name="txtDescription" fieldName="Project Description" onchange="this.isDirty=true;" rows=15 wrap=soft><%=rsGetProjectByProjectId( "ProjectDescription" )%></textarea></td></tr>
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
          <tr><td width=64 align=right><a href="javascript:validateForm(frmProject);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmProject.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:GetDirty(frmProject);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
<!-- End Project Form -->
<%	set rsGetProjectByProjectId = rsGetProjectByProjectId.NextRecordset
		end if
	end if
end if%>

<%if( intProjectId <> 0 ) then%>
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=15%>Call Log</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

<!-- Call Log Table -->
<form id="frmCallLog" name="frmCallLog" action="" method="post">
<input type="hidden" name="txtCustomerIdDetail" id="txtCustomerIdDetail" value="<%=intCustomerId%>">
<input type="hidden" name="txtProjectIdDetail" id="txtProjectIdDetail" value="<%=intProjectId%>">
<tr><td align=center>
 <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
 <tr>
 <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
 <!-- table heading -->
 <tr>
		<th class="itemlist" width="20" align="center">Del</th>
		<th class="itemlist" width="115">Date</th>
    <th class="itemlist" width="35">W.O.</th>
		<th class="itemlist">Comments</th>
		<th class="itemlist" width="100">Employee</th>
 </tr>
  
 <!-- table data -->
 <%if not( rsGetProjectByProjectId is nothing ) then
			if( rsGetProjectByProjectId.State = adStateOpen ) then
				intRowNumber = 1
				
				'set rsGetEmployeesList = cmdGetEmployeesList.Execute
				
				rsGetEmployeesList.MoveFirst
		
				while( not rsGetProjectByProjectId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>					
						<tr class="oddrow">
					<%end if%>
							<td class="itemlist" width="20" align=middle><a href="javascript:ConfirmDelete('<%=rsGetProjectByProjectId( "CallId" )%>', '<%=EscapeQuotes(rsGetProjectByProjectId( "CallComments" ))%>');"><img src="images/form_delete.gif" width=15 height=15 border=0></a></td>
							<td class="itemlist" width="115"><%=rsGetProjectByProjectId( "CallDate" )%></td>
							<td class="itemlist" width="35"><a href="javascript:OpenWorkOrder('<%=rsGetProjectByProjectId( "WorkOrderId" )%>')"><%=rsGetProjectByProjectId( "WorkOrderNumber" )%></a></td>
							<td class="itemlist"><%=rsGetProjectByProjectId( "CallComments" )%></td>
							<td class="itemlist" width="100"><%=rsGetProjectByProjectId( "EmployeeName" )%></td>
						</tr>
					<%intRowNumber = intRowNumber + 1
					rsGetProjectByProjectId.MoveNext
				wend
			end if
		end if%>
		<%if( ( intRowNumber Mod 2 ) = 1 ) then%>
		<tr class="evenrow">
		<%else%>
		<tr class="oddrow">
		<%end if%>
			<td class="itemlist" width=20 align=middle>&nbsp;</td>
			<td class="itemlist" width=120><%=FormatDateTime( Now(), vbGeneralDate )%></td>
			<td>&nbsp;</td>
			<td><input type="text" class="forminput" name="txtCallComments" isRequired="true" fieldName="Comments" value=""></td>
			<td><select class="forminput" name="selEmployeeId" size=1 isRequired="true" fieldName="Employee">
					<option value=""></option>
					<%if not( rsGetEmployeesList is nothing ) then
   						if( rsGetEmployeesList.State = adStateOpen ) then
   							while( not rsGetEmployeesList.EOF )%><option value="<%=rsGetEmployeesList( "EmployeeId" )%>"><%=rsGetEmployeesList( "EmployeeName" )%></option>
   								<%rsGetEmployeesList.MoveNext
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
</form></table>
<%	set rsGetProjectByProjectId = rsGetProjectByProjectId.NextRecordset
end if%>

<%if( intProjectId <> 0 ) then%>
<!-- Middle Title Bar -->
<tr><td class="pagesmalltitle">
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/small_title_top.gif" width=100% height=5></td></tr>
   <tr><td class="smallsectiontitle" width=20%>Work Orders</td>
     <td align=left width=448><IMG SRC="images/small_title_right.gif" width=15 height=17></td>
 </TABLE>
</td></tr>
<!-- End Middle Title Bar -->

<!-- Work Orders Table -->
<tr><td align=center>
 <table width=600 border=0 cellpadding=0 cellspacing=3 align=middle>
 <tr>
 <td><table class="itemlist" width=600 border=1 cellpadding=2 cellspacing=0 rules="cols" align=middle>
  
 <!-- table heading -->
 <tr>
		<th class="itemlist" width="40">Num</th>
		<th class="itemlist" width="50">Promised</th>
		<th class="itemlist" width="50">In</th>
		<th class="itemlist" width="50">Out</th>
		<th class="itemlist" width="410">Serial #</th>
 </tr>
  
 <!-- table data -->
 <%if not( rsGetProjectByProjectId is nothing ) then
			if( rsGetProjectByProjectId.State = adStateOpen ) then
				intRowNumber = 1
		
				while( not rsGetProjectByProjectId.EOF )
					if( ( intRowNumber Mod 2 ) = 1 ) then%>
						<tr class="evenrow">
					<%else%>					
						<tr class="oddrow">
					<%end if%>
							<td class="itemlist" width="40"><a href="javascript:OpenWorkOrder( '<%=rsGetProjectByProjectId( "WorkOrderId" )%>' )"><%=rsGetProjectByProjectId( "WorkOrderNumber" )%></a></td>
							<td class="itemlist" width="50"><%=rsGetProjectByProjectId( "PromiseDate" )%></td>
							<td class="itemlist" width="50"><%=rsGetProjectByProjectId( "DateIn" )%></td>
							<td class="itemlist" width="50"><%=rsGetProjectByProjectId( "DateOut" )%></td>
							<td class="itemlist" width="410"><%=rsGetProjectByProjectId( "SerialNumber" )%></td>
						</tr>
					<%intRowNumber = intRowNumber + 1
					rsGetProjectByProjectId.MoveNext
				wend
			end if
		end if%>
</table></td></tr>
<%end if%>
<!-- End Work Orders Table -->

</table>
</BODY>
</HTML>
<%
if not( rsGetProjectByProjectId is nothing ) then
	if( rsGetProjectByProjectId.State = adStateOpen ) then
		rsGetProjectByProjectId.Close
	end if
end if

if not( rsGetCustomersList is nothing ) then
	if( rsGetCustomersList.State = adStateOpen ) then
		rsGetCustomersList.Close
	end if
end if

if not( rsGetEmployeesList is nothing ) then
	if( rsGetEmployeesList.State = adStateOpen ) then
		rsGetEmployeesList.Close
	end if
end if

if not( rsGetProjectTypesList is nothing ) then
	if( rsGetProjectTypesList.State = adStateOpen ) then
		rsGetProjectTypesList.Close
	end if
end if

if not( rsGetProjectPrioritiesList is nothing ) then
	if( rsGetProjectPrioritiesList.State = adStateOpen ) then
		rsGetProjectPrioritiesList.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProjectPrioritiesList	= nothing
set cmdGetProjectPrioritiesList	= nothing
set rsGetProjectTypesList				= nothing
set cmdGetProjectTypesList			= nothing
set rsGetEmployeesList					= nothing
set cmdGetEmployeesList					= nothing
set rsGetCustomersList					= nothing
set cmdGetCustomersList					= nothing
set rsGetProjectByProjectId			= nothing
set cmdGetProjectByProjectId		= nothing
set cmdEditProject							= nothing
set cnnSQLConnection						= nothing
%>