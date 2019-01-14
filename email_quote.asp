<!--METADATA TYPE="typelib" UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" NAME="CDO for Windows 2000 Library" --> 
<!--METADATA TYPE="typelib" UUID="00000205-0000-0010-8000-00AA006D2EA4" NAME="ADODB Type Library" -->

<%option explicit

dim cnnSQLConnection
dim rsGetEmailByQuoteId, cmdGetEmailByQuoteId
dim rsGetQuoteByQuoteId, cmdGetQuoteByQuoteId
dim rsGetCompanyInformation, cmdGetCompanyInformation
dim rsGetQuotePartsList, cmdGetQuotePartsList
dim rsGetQuoteBearingsList, cmdGetQuoteBearingsList
dim rsGetQuoteSubWorkList, cmdGetQuoteSubWorkList
dim intQuoteId, isFullQuote
dim strFromEmail, strToEmail, strCCEmail, strSubject
dim bSent
dim StrBody
dim intRowNumber

set rsGetEmailByQuoteId				= nothing
set cmdGetEmailByQuoteId			= nothing
set rsGetQuoteByQuoteId				= nothing
set cmdGetQuoteByQuoteId			= nothing
set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set rsGetQuotePartsList				= nothing
set cmdGetQuotePartsList			= nothing
set rsGetQuoteBearingsList		= nothing
set cmdGetQuoteBearingsList		= nothing
set cmdGetQuoteSubWorkList		= nothing
set rsGetQuoteSubWorkList			= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetEmailByQuoteId			= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteByQuoteId			= Server.CreateObject( "ADODB.Command" )
set cmdGetCompanyInformation	= Server.CreateObject( "ADODB.Command" )
set cmdGetQuotePartsList			= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteBearingsList		= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteSubWorkList		= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intQuoteId			= Request.QueryString( "intQuoteId" )
isFullQuote			= Request.QueryString( "isFullQuote" )
strFromEmail		= Request.Form( "txtFromEmail" )
strToEmail			= Request.Form( "txtToEmail" )
strCCEmail			= Request.Form( "txtCCEmail" )
strSubject			= Request.Form( "txtSubject" )

bSent = false

if ( strToEmail <> "") then

	strBody = ""

	cmdGetQuoteByQuoteId.ActiveConnection = cnnSQLConnection
	cmdGetQuoteByQuoteId.CommandText = "{call spGetQuoteByQuoteId( ?,?,? ) }"
	cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
	cmdGetQuoteByQuoteId.Parameters.Append cmdGetQuoteByQuoteId.CreateParameter( "@CustomerId", adInteger, adParamOutput, 4 )
	cmdGetQuoteByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

	on error resume next
	set rsGetQuoteByQuoteId = cmdGetQuoteByQuoteId.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
						
	on error goto 0
	'**********************************************************************************
	cmdGetQuotePartsList.ActiveConnection	= cnnSQLConnection
	cmdGetQuotePartsList.CommandText		= "{call spGetQuotePartsList( ?,? ) }"

	cmdGetQuotePartsList.Parameters.Append cmdGetQuotePartsList.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetQuotePartsList.Parameters.Append cmdGetQuotePartsList.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )

	cmdGetQuotePartsList.Parameters( "@QuoteId" ).Value			= intQuoteId

	on error resume next
	set rsGetQuotePartsList = cmdGetQuotePartsList.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
								
	on error goto 0
	'**********************************************************************************
	cmdGetQuoteBearingsList.ActiveConnection	= cnnSQLConnection
	cmdGetQuoteBearingsList.CommandText		= "{call spGetQuoteBearingsList( ?,? ) }"

	cmdGetQuoteBearingsList.Parameters.Append cmdGetQuoteBearingsList.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetQuoteBearingsList.Parameters.Append cmdGetQuoteBearingsList.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )

	cmdGetQuoteBearingsList.Parameters( "@QuoteId" ).Value			= intQuoteId

	on error resume next
	set rsGetQuoteBearingsList = cmdGetQuoteBearingsList.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
								
	on error goto 0
	'**********************************************************************************
	cmdGetQuoteSubWorkList.ActiveConnection	= cnnSQLConnection
	cmdGetQuoteSubWorkList.CommandText		= "{call spGetQuoteSubWorkList( ?,? ) }"

	cmdGetQuoteSubWorkList.Parameters.Append cmdGetQuoteSubWorkList.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdGetQuoteSubWorkList.Parameters.Append cmdGetQuoteSubWorkList.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )

	cmdGetQuoteSubWorkList.Parameters( "@QuoteId" ).Value			= intQuoteId

	on error resume next
	set rsGetQuoteSubWorkList = cmdGetQuoteSubWorkList.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
								
	on error goto 0
	'*******************************************************************************
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

	on error goto 0
	'*******************************************************************************
	if not( rsGetQuoteByQuoteId is nothing ) then
			if( rsGetQuoteByQuoteId.State = adStateOpen ) then
				if( not rsGetQuoteByQuoteId.EOF ) then


				strBody = strBody & "<head><style><!--"
				strBody = strBody & "body { FONT-FAMILY: Tahoma, Verdana, Arial, 'Microsoft Sans Serif' }"
				strBody = strBody & ".contact_info { FONT-FAMILY: Verdana, Arial; FONT-SIZE: 10pt; PADDING-BOTTOM: 0.2in; PADDING-RIGHT: 20px; TEXT-ALIGN: center }"
				strBody = strBody & ".quotetitle{FONT-FAMILY: Verdana, Arial;FONT-SIZE: 12pt;FONT-WEIGHT: 900;PADDING-BOTTOM: 15px;TEXT-DECORATION: underline}"
				strBody = strBody & ".quotetext{FONT-FAMILY: Verdana, Arial;FONT-SIZE: 10pt}"
				strBody = strBody & ".quoteinfo{FONT-FAMILY: Verdana, Arial;FONT-SIZE: 10pt;FONT-WEIGHT: bolder}"
				strBody = strBody & ".quoteitem{BORDER-BOTTOM: silver thin solid;FONT-FAMILY: Verdana, Arial;FONT-SIZE: 10pt;FONT-STYLE: italic;MARGIN-LEFT: 20px;PADDING-LEFT: 20px}"
				strBody = strBody & ".quoteitemtext{FONT-FAMILY: Verdana, Arial;FONT-SIZE: 8pt;PADDING-LEFT: 30px;}"
				strBody = strBody & ".quotetable{PADDING-BOTTOM: 15px}"
				strBody = strBody & ".quotetotal{BORDER-BOTTOM: black thin solid}"
				strBody = strBody & ".quotenote{FONT-FAMILY: Verdana, Arial;FONT-SIZE: 80%;FONT-STYLE: italic}"
				strBody = strBody & ".quoteauthorizationtable{BORDER-TOP: black double;PADDING-TOP: 8px}"
				strBody = strBody & ".authorizationline{BORDER-BOTTOM: black 2px solid}"
				strBody = strBody & ".holidays {font-weight:bold;color:red;text-align:center;}"
				strBody = strBody & ".holigreen {color:darkgreen}"
				strBody = strBody & "--></style></head>"
				
				
				strBody = strBody & "  <table align=left width=100% "					
				strBody = strBody & "    <tr><td><table width=100% >"
				strBody = strBody & "      <tr><td valign=top width=282><img src=""http://centerline-inc.com/quotes/centerline_logo.jpg"" width=282 height=138></td>"
				strBody = strBody & "       <td align=right><table>"
				strBody = strBody & "        <tr><td class=contact_info width=100% >"
				strBody = strBody & "         <table>"
				strBody = strBody & "          <tr><td colspan=2 align=right class=quotetext>" & rsGetCompanyInformation( "CompanyName" ) & "</td></tr>"
				strBody = strBody & "          <tr><td colspan=2 align=right class=quotetext>" & rsGetCompanyInformation( "Address" ) & "</td></tr>"
				strBody = strBody & "          <tr><td colspan=2 align=right class=quotetext>" & rsGetCompanyInformation( "City" ) & ",&nbsp;" & rsGetCompanyInformation( "StateOrProvince" ) & "&nbsp;" & rsGetCompanyInformation( "PostalCode" ) & "</td></tr>"
				strBody = strBody & "          <tr><td colspan=2 align=right class=quotetext>Phone:&nbsp;" & rsGetCompanyInformation( "PhoneNumber" ) & "</td></tr>"
				strBody = strBody & "          <tr><td colspan=2 align=right class=quotetext>Fax:&nbsp;" & rsGetCompanyInformation( "FaxNumber" ) & "</td></tr>"
				strBody = strBody & "         </table></td></tr></table></td></tr></table></td></tr>"

				strBody = strBody & "    <tr><td class=quotetitle align=middle>REPAIR QUOTE</td></tr>"

				strBody = strBody & "    <tr><td class=quotetable><table>"
				strBody = strBody & "      <tr><td class=quotetext width=180>Prepared Expressly For:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "Customer" ) & "</td></tr>"
				if (rsGetQuoteByQuoteId( "PONumber" ) <> "") then 
					strBody = strBody & "<tr><td class=quotetext>Purchase Order Number:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "PONumber" ) & "</td></tr>"
				end if
				strBody = strBody & "      <tr><td class=quotetext width=180>Date Quoted:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "DateQuoted" ) & "</td></tr>"
				strBody = strBody & "      <tr><td class=quotetext width=180>Work Order Number:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "WorkOrderNumber" )& "</td></tr>"
				strBody = strBody & "      <tr><td class=quotetext width=180>Spindle Type:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "SpindleType" ) & "</td></tr>"
				strBody = strBody & "      <tr><td class=quotetext width=180>Serial Number:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "SerialNumber" ) & "</td></tr>"
				strBody = strBody & "    </table></td></tr>"
					  
				strBody = strBody & "    <tr><td class=quotetable><table width=500>"
				strBody = strBody & "      <tr><td class=quotetext>We are pleased to offer you the following quote:</td></tr>"
				strBody = strBody & "      <tr><td class=quoteitem>Bearings</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalBearingCharge" ), 2 ) & "</td></tr>"

				if (isFullQuote = "true") then
					strBody = strBody & "			<tr><td class=quoteitemtext>"
					if not( rsGetQuoteBearingsList is nothing ) then
						if( rsGetQuoteBearingsList.State = adStateOpen ) then
							intRowNumber = 1
							while( not rsGetQuoteBearingsList.EOF )
								strBody = strBody & "			<br><b>" & rsGetQuoteBearingsList( "Qty" ) & "</b>&nbsp;&nbsp;&nbsp;" & rsGetQuoteBearingsList( "PartNumber" )
								intRowNumber = intRowNumber + 1
								rsGetQuoteBearingsList.MoveNext
							wend
							strBody = strBody & "<br>&nbsp;"
						end if
					end if
				strBody = strBody & "</tr>"
				end if
				strBody = strBody & "<tr><td class=quoteitem>Parts</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalPartCharge" ), 2 ) & "</td></tr>"
				if (isFullQuote = "true") then
					strBody = strBody & "<tr><td class=quoteitemtext>"
						if not( rsGetQuotePartsList is nothing ) then
							if( rsGetQuotePartsList.State = adStateOpen ) then
								intRowNumber = 1
								while( not rsGetQuotePartsList.EOF )
									strBody = strBody & "<br><b>" & rsGetQuotePartsList( "Qty" ) & "</b>&nbsp;&nbsp;&nbsp;" & rsGetQuotePartsList( "ProductName" )
									intRowNumber = intRowNumber + 1
									rsGetQuotePartsList.MoveNext
								wend
								strBody = strBody & "<br>&nbsp;"
							end if
						end if
					strBody = strBody & "</tr>"
				end if
				strBody = strBody & "<tr><td class=quoteitem>Labor</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalLaborAndHandling" ), 2 ) & "</td></tr>"
				strBody = strBody & "<tr><td class=quoteitem>Rework</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalReworkCharge" ), 2) & "</td></tr>"
				if (isFullQuote = "true") then
				strBody = strBody & "<tr><td class=quoteitemtext>"
					if not( rsGetQuoteSubWorkList is nothing ) then
						if( rsGetQuoteSubWorkList.State = adStateOpen ) then
							intRowNumber = 1
							while( not rsGetQuoteSubWorkList.EOF )
								strBody = strBody & "<br>" & rsGetQuoteSubWorkList( "SubWorkDescription" )
								intRowNumber = intRowNumber + 1
								rsGetQuoteSubWorkList.MoveNext
							wend
							strBody = strBody & "<br>&nbsp;"
						end if
					end if
					strBody = strBody & "</tr>"
					end if
				strBody = strBody & "<tr><td>&nbsp;</td><td class=quotetotal>&nbsp;</td></tr>"
				strBody = strBody & "<tr><td class=quoteitem>Total Repair ~ Standard</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalRepairPrice" ), 2 ) & "</td></tr>"
				strBody = strBody & "<tr><td class=quoteitem>Total Repair ~ Expedited</td><td class=quoteinfo align=right>" & FormatCurrency( rsGetQuoteByQuoteId( "TotalExpeditingCharge" ), 2 ) & "</td></tr>"
			strBody = strBody & "</table></td></tr>"
	    
			strBody = strBody & "<tr><td class=quotetable><table width=100% >"
			strBody = strBody & "<tr><td class=quotetext width=180>Comments:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "QuoteSpecificComments" ) & "</td></tr>"
			strBody = strBody & "</table></td></tr>"
	    
			strBody = strBody & "<tr><td class=quotetable><table width=100% >"
				strBody = strBody & "<tr><td class=quotetext width=200>Delivery Information:</td><td class=quoteitem width=440>" & rsGetQuoteByQuoteId( "DeliveryInformation" ) & "</td></tr>"
				strBody = strBody & "<tr><td class=quotetext width=200>Exp. Delivery Information:</td><td class=quoteitem width=440>" & rsGetQuoteByQuoteId( "ExpeditedDeliveryInformation" ) & "</td></tr>"
			strBody = strBody & "</table></td></tr>"
	    
			strBody = strBody & "<tr><td class=quotenote><p>Please indicate your approval and provide P.O. information in the appropriate space below. All repair approvals should be either faxed to (580) 762 4722 or emailed to repair.quote@centerline-inc.com. <strong>Your acknowledgement indicates authority to proceed with the repair and acceptance of Centerline's Terms.</strong></p><p>All amounts are quoted in US Dollars. Payment terms are Net 30 days from Invoice Date with approved credit. Invoices Unpaid Past Terms Are Subject To 5% Service Charge/Month. Freight Is F.O.B. Ponca City, OK (unless contrary agreement is reached before quoting).</p><p>Quotes are valid for one week, parts pricing and delivery subject to prior sale. Materials left at Centerline without repair approval for a period of six months will be scrapped.</p></td></tr>"
	    
			strBody = strBody & "<tr><td class=quotetable><table width=100% >"
				strBody = strBody & "<tr><td width=300 class=quotetext>&nbsp;</td><td class=quotetext>Sincerely yours,</td></tr>"
				strBody = strBody & "<tr><td width=300 class=quotetext>&nbsp;</td><td class=quotetext>" & rsGetQuoteByQuoteId( "QuotedBy" ) & ", Sales Representative</td></tr>"
				strBody = strBody & "<tr><td width=300 class=quotetext>&nbsp;</td><td class=quotetext>Jason Mabry, Project Manager</td></tr>"
				strBody = strBody & "<tr><td width=300>&nbsp;</td><td class=quotetext>for Centerline, Inc.</td></tr>"
			strBody = strBody & "</table></td></tr>"
	    
			strBody = strBody & "<tr><td class=quoteauthorizationtable align=middle><table width=500 align=center>"
				strBody = strBody & "<tr><td class=quotetitle align=middle colspan=4>CUSTOMER AUTHORIZATION " & rsGetQuoteByQuoteId( "WorkOrderNumber" )& "</td></tr>"
				strBody = strBody & "<tr><td class=quoteinfo>PHONE:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "TelephoneNumber" ) & "</td>"
					strBody = strBody & "<td class=quoteinfo>FAX:</td><td class=quoteinfo>" & rsGetQuoteByQuoteId( "FaxNumber" ) & "</td></tr>"
			strBody = strBody & "</table></td></tr>"
			strBody = strBody & "<tr><td align=middle><table width=500 align=center>"
				strBody = strBody & "<tr><td class=quoteinfo>DATE APPROVED:</td><td class=authorizationline width=110>&nbsp;</td>"
					strBody = strBody & "<td class=quoteinfo>P.O.#:</td><td class=authorizationline width=190>&nbsp;</td></tr>"
			strBody = strBody & "</table></td></tr>"
			strBody = strBody & "<tr><td align=middle><table width=500 align=center>"
				strBody = strBody & "<tr><td class=quoteinfo>APPROVED BY:</td><td class=authorizationline width=350>&nbsp;</td></tr>"
			strBody = strBody & "</table></td></tr>"
			strBody = strBody & "<tr><td align=middle><table width=500 align=center>"
				strBody = strBody & "<tr><td class=quoteinfo>EXPECTED DELIVERY:</td><td class=authorizationline width=310>&nbsp;</td></tr>"



				





				
				
		strBody = strBody & "</table></td></tr>"
			strBody = strBody & "  </table>"
			end if
		end if
	end if

	if not( rsGetQuoteByQuoteId is nothing ) then
		if( rsGetQuoteByQuoteId.State = adStateOpen ) then
			rsGetQuoteByQuoteId.Close
		end if
	end if

	if not( rsGetCompanyInformation is nothing ) then
		if( rsGetCompanyInformation.State = adStateOpen ) then
			rsGetCompanyInformation.Close
		end if
	end if


	'Const cdoSendUsingMethod        = _
	'	"http://schemas.microsoft.com/cdo/configuration/sendusing"
	'Const cdoSendUsingPort          = 2
	'Const cdoSMTPServer             = _
	'	"http://schemas.microsoft.com/cdo/configuration/smtpserver"
	'Const cdoSMTPServerPort         = _
	'	"http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	'Const cdoSMTPConnectionTimeout  = _
	'	"http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	'Const cdoSMTPAuthenticate       = _
	'	"http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	'Const cdoBasic                  = 1
	'Const cdoSendUserName           = _
	'	"http://schemas.microsoft.com/cdo/configuration/sendusername"
	'Const cdoSendPassword           = _
	'	"http://schemas.microsoft.com/cdo/configuration/sendpassword"

	Dim objConfig  ' As CDO.Configuration
	Dim objMessage ' As CDO.Message
	Dim Fields     ' As ADODB.Fields

	' Get a handle on the config object and it's fields
	Set objConfig = Server.CreateObject("CDO.Configuration")
	Set Fields = objConfig.Fields

	' Set config fields we care about
	With Fields
		.Item(cdoSendUsingMethod)       = cdoSendUsingPort
		.Item(cdoSMTPServer)            = "smtp.centerline-inc.com"
		.Item(cdoSMTPServerPort)        = 465
		.Item(cdoSMTPConnectionTimeout) = 10
		.Item(cdoSMTPAuthenticate)      = cdoBasic
'.Item(cdoSendUserName)          = "engster@swbell.net"
'.Item(cdoSendPassword)          = "abcd1234"
		.Item(cdoSendUserName)          = "inventory@centerline-inc.com"
		.Item(cdoSendPassword)          = "S9p9zZtL"

		.Update
	End With

	Set objMessage = Server.CreateObject("CDO.Message")

	Set objMessage.Configuration = objConfig

	With objMessage
		.To       = strToEmail
		.CC	  = strFromEmail & "," & strCCEmail
		.From     = strFromEmail
		.Subject  = strSubject
		.HTMLBody = strBody
		.Send
	End With

	Set Fields = Nothing
	Set objMessage = Nothing
	Set objConfig = Nothing
	
	bSent = true
end if

cmdGetEmailByQuoteId.ActiveConnection = cnnSQLConnection
cmdGetEmailByQuoteId.CommandText = "{call spGetEmailByQuoteId( ?,? ) }"
cmdGetEmailByQuoteId.Parameters.Append cmdGetEmailByQuoteId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetEmailByQuoteId.Parameters.Append cmdGetEmailByQuoteId.CreateParameter( "@QuoteId", adInteger, adParamInput, 4 )
cmdGetEmailByQuoteId.Parameters( "@QuoteId" )	= intQuoteId

on error resume next
set rsGetEmailByQuoteId = cmdGetEmailByQuoteId.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if
					
on error goto 0
'*******************************************************************************
if not( rsGetEmailByQuoteId is nothing ) then
		if( rsGetEmailByQuoteId.State = adStateOpen ) then
			if( not rsGetEmailByQuoteId.EOF ) then%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE>Centerline - Send Quote Email</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script type="text/javascript" src="js/forms.js"></script>
  <script LANGUAGE=JavaScript>
	<!-- Hide script from old browsers
	
	// End hiding script from old browsers -->
  </script>
</HEAD>
<BODY topmargin=0 leftmargin=0 <%if (bSent) then%>onload="window.close();"<%end if%>>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetEmailByQuoteId is nothing ) then
	if( rsGetEmailByQuoteId.State = adStateOpen ) then
		if( not rsGetEmailByQuoteId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle" width=30%>&nbsp;&nbsp;E-mail Quote <%=rsGetEmailByQuoteId( "QuoteId" )%></td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Email Form -->
<tr><td align=center>
  <form id="frmEmailQuote" name="frmEmailQuote" action="" method="post">
		<input type="hidden" name="txtQuoteId" value="<%=intQuoteId%>">
    <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
			<tr><td class="label_1">From:</td><td class="input_1" colspan="3"><input type="text" class="forminput" name="txtFromEmail" maxlength=200 isRequired="true" fieldName="From" value="<%=rsGetEmailByQuoteId( "FromEmail" )%>"></td>
			<tr><td class="label_1">To:</td><td class="input_1" colspan="3"><input type="text" class="forminput" name="txtToEmail" maxlength=200 isRequired="true" fieldName="To" value="<%=rsGetEmailByQuoteId( "ToEmail" )%>"></td>
			<tr><td class="label_1">CC:</td><td class="input_1" colspan="3"><input type="text" class="forminput" name="txtCCEmail" maxlength=200></td></tr>
			<tr><td class="label_1">Subject:</td><td class="input_1" colspan="3"><input type="text" class="forminput" name="txtSubject" fieldName="Subject" maxlength=200 value="Centerline Inc. Repair Quote for PO#: <%=rsGetEmailByQuoteId( "PONum" )%>"></td></tr>
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
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmEmailQuote);"><IMG src="images/send_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmEmployee.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmEmailQuote);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
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
<%	end if
	end if
end if

if not( rsGetEmailByQuoteId is nothing ) then
	if( rsGetEmailByQuoteId.State = adStateOpen ) then
		rsGetEmailByQuoteId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetEmailByQuoteId				= nothing
set cmdGetEmailByQuoteId			= nothing
set cnnSQLConnection					= nothing
%>
