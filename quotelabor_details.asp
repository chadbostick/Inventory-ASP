<%option explicit

dim cnnSQLConnection
dim rsGetQuoteLaborByQuoteId, cmdGetQuoteLaborByQuoteId
dim cmdEditQuote
dim intQuoteId

set cnnSQLConnection					= nothing
set rsGetQuoteLaborByQuoteId	= nothing
set cmdGetQuoteLaborByQuoteId	= nothing
set cmdEditQuote							= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteLaborByQuoteId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then
	set cmdEditQuote	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditQuote.ActiveConnection	= cnnSQLConnection
	cmdEditQuote.CommandText			= "{? = call spEditQuoteLabor( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,? ) }"
	
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@QuoteID", adInteger, adParamInput, 4 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@DisassemblyEvaluation", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@HoursDisassembly", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@CleanAndInspect", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@HoursCleanAndInspect", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@InhouseGrinding", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@GrindingHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@Balancing", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@BalancingHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ElectricalWork", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@ElectricalHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@GreaseBearings", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@GreaseHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@AssemblyAndTest", adVarChar, adParamInput, 100 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@AssemblyAndTestHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@MscWorkNeeded", adLongVarChar, adParamInput, 1048576 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@MscWorkHours", adVarChar, adParamInput, 10 )
	cmdEditQuote.Parameters.Append cmdEditQuote.CreateParameter( "@LaborCommission", adVarChar, adParamInput, 10 )
	
	if( Request.Form( "txtQuoteId" ) <> "" ) then cmdEditQuote.Parameters( "@QuoteID" ) = Request.Form( "txtQuoteId" )
	if( Request.Form( "txtDisassemblyEvaluation" ) <> "" ) then cmdEditQuote.Parameters( "@DisassemblyEvaluation" ) = Request.Form( "txtDisassemblyEvaluation" )
	if( Request.Form( "txtDisassemblyEvaluationHours" ) <> "" ) then cmdEditQuote.Parameters( "@HoursDisassembly" ) = Request.Form( "txtDisassemblyEvaluationHours" )
	if( Request.Form( "txtCleanInspect" ) <> "" ) then cmdEditQuote.Parameters( "@CleanAndInspect" ) = Request.Form( "txtCleanInspect" )
	if( Request.Form( "txtCleanInspectHours" ) <> "" ) then cmdEditQuote.Parameters( "@HoursCleanAndInspect" ) = Request.Form( "txtCleanInspectHours" )
	if( Request.Form( "txtInhouseGrinding" ) <> "" ) then cmdEditQuote.Parameters( "@InhouseGrinding" ) = Request.Form( "txtInhouseGrinding" )
	if( Request.Form( "txtInhouseGrindingHours" ) <> "" ) then cmdEditQuote.Parameters( "@GrindingHours" ) = Request.Form( "txtInhouseGrindingHours" )
	if( Request.Form( "txtBalancing" ) <> "" ) then cmdEditQuote.Parameters( "@Balancing" ) = Request.Form( "txtBalancing" )
	if( Request.Form( "txtBalancingHours" ) <> "" ) then cmdEditQuote.Parameters( "@BalancingHours" ) = Request.Form( "txtBalancingHours" )
	if( Request.Form( "txtElectricalWork" ) <> "" ) then cmdEditQuote.Parameters( "@ElectricalWork" ) = Request.Form( "txtElectricalWork" )
	if( Request.Form( "txtElectricalWorkHours" ) <> "" ) then cmdEditQuote.Parameters( "@ElectricalHours" ) = Request.Form( "txtElectricalWorkHours" )
	if( Request.Form( "txtGreaseBearings" ) <> "" ) then cmdEditQuote.Parameters( "@GreaseBearings" ) = Request.Form( "txtGreaseBearings" )
	if( Request.Form( "txtGreaseBearingsHours" ) <> "" ) then cmdEditQuote.Parameters( "@GreaseHours" ) = Request.Form( "txtGreaseBearingsHours" )
	if( Request.Form( "txtAssemblyTest" ) <> "" ) then cmdEditQuote.Parameters( "@AssemblyAndTest" ) = Request.Form( "txtAssemblyTest" )
	if( Request.Form( "txtAssemblyTestHours" ) <> "" ) then cmdEditQuote.Parameters( "@AssemblyAndTestHours" ) = Request.Form( "txtAssemblyTestHours" )
	if( Request.Form( "txtMiscWork" ) <> "" ) then cmdEditQuote.Parameters( "@MscWorkNeeded" ) = Request.Form( "txtMiscWork" )
	if( Request.Form( "txtMiscWorkHours" ) <> "" ) then cmdEditQuote.Parameters( "@MscWorkHours" ) = Request.Form( "txtMiscWorkHours" )
		
	
	'if( Left( Request.Form( "txtLaborCommission" ), 1 ) = "$" ) then
	'	if( Request.Form( "txtLaborCommission" ) <> "" ) then cmdEditQuote.Parameters( "@LaborCommission" ) = Mid( Request.Form( "txtLaborCommission" ), 2 )
	'else
	'	if( Request.Form( "txtLaborCommission" ) <> "" ) then cmdEditQuote.Parameters( "@LaborCommission" ) = Request.Form( "txtLaborCommission" )
	'end if

	
	on error resume next
	cmdEditQuote.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if

	on error goto 0
	
	intQuoteId = cmdEditQuote.Parameters( "@Return" ).Value
else
	intQuoteId = Request.QueryString( "intQuoteId" )
end if

cmdGetQuoteLaborByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetQuoteLaborByQuoteId.CommandText = "{call spGetQuoteLaborByQuoteId( ?,? ) }"
cmdGetQuoteLaborByQuoteId.Parameters.Append cmdGetQuoteLaborByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetQuoteLaborByQuoteId.Parameters.Append cmdGetQuoteLaborByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetQuoteLaborByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetQuoteLaborByQuoteId = cmdGetQuoteLaborByQuoteId.Execute
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
<TITLE>Centerline - View / Edit Labor Cost</TITLE>
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

<%if not( rsGetQuoteLaborByQuoteId is nothing ) then
		if( rsGetQuoteLaborByQuoteId.State = adStateOpen ) then
			if( not rsGetQuoteLaborByQuoteId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=20%>Labor Cost</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td>
     <td class="sectioncaption" align=right>Quote #&nbsp;&nbsp;<%=intQuoteId%></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Labor Form -->
<form id="frmLabor" name="frmLabor" action="" method="post">
<tr><td align=center>
  	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
  	<input type="hidden" name="txtQuoteId" value="<%=intQuoteId%>">
  	 <tr><td class="label_1">Spindle Type</td><td class="input_2" colspan=3><input type="text" class="forminput" name="txtSpindleType" onchange="this.isDirty=true;" fieldName="Spindle Type" readonly value="<%=rsGetQuoteLaborByQuoteId( "SpindleType" )%>"></td></tr>
     <tr><td class="long_left_label_1">Disassembly & Evaluation</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtDisassemblyEvaluation" onchange="this.isDirty=true;" fieldName=" Disassembly Evaluation" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "DisassemblyEvaluation" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtDisassemblyEvaluationHours" isNumber="true" isonchange="this.isDirty=true;" fieldName="Disassembly Evaluation Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "HoursDisassembly" )%>"></td></tr>
     <tr><td class="long_left_label_1">Clean & Inspect</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtCleanInspect" onchange="this.isDirty=true;" fieldName="Clean and Inspect" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "CleanAndInspect" )%>"></td>
 	  <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtCleanInspectHours" isNumber="true" maxlength=10 onchange="this.isDirty=true;" fieldName="Clean and Inspect Hours" value="<%=rsGetQuoteLaborByQuoteId( "HoursCleanAndInspect" )%>"></td></tr>
     <tr><td class="long_left_label_1">Inhouse Grinding</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtInhouseGrinding" onchange="this.isDirty=true;" fieldName="Inhouse Grinding" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "InhouseGrinding" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtInhouseGrindingHours" isNumber="true" onchange="this.isDirty=true;" fieldName="Inhouse Grinding Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "GrindingHours" )%>"></td></tr>
     <tr><td class="long_left_label_1">Balancing</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtBalancing" onchange="this.isDirty=true;" fieldName="Balancing" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "Balancing" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtBalancingHours" isNumber="true" maxlength=10 onchange="this.isDirty=true;" fieldName="Balancing Hours" value="<%=rsGetQuoteLaborByQuoteId( "BalancingHours" )%>"></td></tr>
     <tr><td class="long_left_label_1">Electrical Work</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtElectricalWork" onchange="this.isDirty=true;" fieldName="Electrical Work" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "ElectricalWork" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtElectricalWorkHours" isNumber="true" onchange="this.isDirty=true;" fieldName="Electrical Work Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "ElectricalHours" )%>"></td></tr>
     <tr><td class="long_left_label_1">Bearing Preparation</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtGreaseBearings" onchange="this.isDirty=true;" fieldName="Grease Bearings" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "GreaseBearings" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtGreaseBearingsHours" isNumber="true" onchange="this.isDirty=true;" fieldName="Grease Bearings Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "GreaseHours" )%>"></td></tr>
     <tr><td class="long_left_label_1">Assembly & Test</td><td class="long_left_input_1"><input type="text" class="forminput" name="txtAssemblyTest" onchange="this.isDirty=true;" fieldName="Assembly Test" maxlength=100 value="<%=rsGetQuoteLaborByQuoteId( "AssemblyAndTest" )%>"></td>
       <td class="long_left_label_2">Hours</td><td class="long_left_input_2"><input type="text" class="forminput" name="txtAssemblyTestHours" isNumber="true" onchange="this.isDirty=true;" fieldName="Assembly Test Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "AssemblyAndTestHours" )%>"></td></tr>
     <tr><td class="long_left_label_1" colspan=4>Misc Work Needed</td>
 	<tr><td class="long_left_input_1" colspan=4><textarea class="forminput" name="txtMiscWork" onchange="this.isDirty=true;" fieldName="Misc Work" rows=3 wrap=soft><%=rsGetQuoteLaborByQuoteId( "MscWorkNeeded" )%></textarea></td></tr>
     <tr><td class="long_left_input_1" colspan=2>&nbsp;</td><td class="long_left_label_2">Hours</td><td class="input_2"><input type="text" class="forminput" name="txtMiscWorkHours" isNumber="true" onchange="this.isDirty=true;" fieldName="Misc Work Hours" maxlength=10 value="<%=rsGetQuoteLaborByQuoteId( "MscWorkHours" )%>"></td></tr>
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
 	<p style="font-size:10px;color:red;text-align:left;padding-left:34px;">rate $145.00 / hour</p>
    <tr><!--<td class="generic_label" style="width:90px">Commission</td><td style="width:50px"><input type="text" class="forminput" name="txtLaborCommission" value="<%=rsGetQuoteLaborByQuoteId( "LaborCommission" )%>"></td>-->
      <td class="generic_label" style="width:120px; padding-left:30px">Total Labor Cost</td><td style="width:90px"><input type="text" class="forminput" name="txtTotalLaborCost" readonly value="<%=FormatCurrency( rsGetQuoteLaborByQuoteId( "TotalLaborCost" ), 2 )%>"></td>
        <td class="generic_label" style="width:100px; padding-left:30px">Total Labor Hours</td><td style="width:60px"><input type="text" class="forminput" name="txtTotalLaborHours" readonly value="<%=rsGetQuoteLaborByQuoteId( "TotalLaborHours" )%>"></td></tr>
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
    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=320 align=right>
          <tr><td width=64 align=right><a href="javascript:window.navigate( '<%=Request.QueryString( "strPrevPage" )%>' );"><IMG src="images/back_button.gif" width=60 height=20 border=0></a></td>
						<td width=64 align=right><a href="javascript:frmLabor.submit();"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:frmLabor.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:OpenPrintBox('quote_timetracking.asp?intQuoteId=<%=intQuoteId%>');"><IMG src="images/print_button.gif" width=60 height=20 border=0></a></td>
            <td width=64 align=right><a href="javascript:GetDirty(frmLabor);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
     </table></td></tr>
 </table>
</td></tr>
</form>
<%	end if
	end if
end if%>
<!-- End Labor Form -->

</BODY>
</HTML>
<%
if not( rsGetQuoteLaborByQuoteId is nothing ) then
	if( rsGetQuoteLaborByQuoteId.State = adStateOpen ) then
		rsGetQuoteLaborByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set cnnSQLConnection					= nothing
set rsGetQuoteLaborByQuoteId	= nothing
set cmdGetQuoteLaborByQuoteId	= nothing
set cmdEditQuote							= nothing
%>
