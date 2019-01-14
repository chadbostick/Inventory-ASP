<%option explicit

dim cnnSQLConnection
dim rsGetQuoteByQuoteId, cmdGetQuoteByQuoteId
dim rsGetCompanyInformation, cmdGetCompanyInformation
dim rsGetQuotePartsList, cmdGetQuotePartsList
dim rsGetQuoteBearingsList, cmdGetQuoteBearingsList
dim rsGetQuoteSubWorkList, cmdGetQuoteSubWorkList
dim intQuoteId, isFullQuote
dim intRowNumber

set rsGetQuoteByQuoteId				= nothing
set cmdGetQuoteByQuoteId			= nothing
set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set rsGetQuotePartsList				= nothing
set cmdGetQuotePartsList				= nothing
set rsGetQuoteBearingsList				= nothing
set cmdGetQuoteBearingsList				= nothing
set cmdGetQuoteSubWorkList				= nothing
set rsGetQuoteSubWorkList				= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetQuoteByQuoteId			= Server.CreateObject( "ADODB.Command" )
set cmdGetCompanyInformation	= Server.CreateObject( "ADODB.Command" )
set cmdGetQuotePartsList				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteBearingsList				= Server.CreateObject( "ADODB.Command" )
set cmdGetQuoteSubWorkList				= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

intQuoteId = Request.QueryString( "intQuoteId" )
isFullQuote = Request.QueryString( "isFullQuote" )

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
			if( not rsGetQuoteByQuoteId.EOF ) then%>
<HTML>
<TITLE>Repair Quote&nbsp;&nbsp;#<%=rsGetQuoteByQuoteId( "QuoteId" )%></TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
</HEAD>

<BODY onload="javascript: if (window.print) window.print();">
  <table align=left width="100%">
  
    <tr><td><table width="100%" class="form_header">
      <tr><td valign=top><img src="images/centerline_logo.png" width=282 height=138></td>
       <td align=right><table>   
        <tr>
		<td class="contact_info">
         <table width="100%">
          <tr><td><%=rsGetCompanyInformation( "CompanyName" )%></td></tr>
          <tr><td><%=rsGetCompanyInformation( "Address" )%></td></tr>
          <tr><td><%=rsGetCompanyInformation( "City" )%>,&nbsp;<%=rsGetCompanyInformation( "StateOrProvince" )%>&nbsp;<%=rsGetCompanyInformation( "PostalCode" )%></td></tr>
          <tr><td>Phone:&nbsp;<%=rsGetCompanyInformation( "PhoneNumber" )%></td></tr>
          <tr><td>Fax:&nbsp;<%=rsGetCompanyInformation( "FaxNumber" )%></td></tr>
         </table>
		</td>
		</tr>
        </table></td></tr>
    </table></td></tr>

    <tr><td class="quotetitle" align=middle>REPAIR QUOTE</td></tr>

    <tr><td class="quotetable"><table>
      <tr><td class="quotetext" width="180">Prepared Expressly For:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "Customer" )%></td></tr>
      <%if (rsGetQuoteByQuoteId( "PONumber" ) <> "") then %><tr><td class="quotetext">Purchase Order Number:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "PONumber" )%></td></tr><%end if%>
      <tr><td class="quotetext" width="180">Date Quoted:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "DateQuoted" )%></td></tr>
      <tr><td class="quotetext" width="180">Work Order Number:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "WorkOrderNumber" )%></td></tr>
      <tr><td class="quotetext" width="180">Spindle Type:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "SpindleType" )%></td></tr>
      <tr><td class="quotetext" width="180">Serial Number:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "SerialNumber" )%></td></tr>
    </table></td></tr>
    
    <tr><td class="quotetable"><table width="500">
      <tr><td class="quotetext">We are pleased to offer you the following quote:</td></tr>
      <tr><td class="quoteitem">Bearings</td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalBearingCharge" ), 2 )%></td></tr>
      <%if (isFullQuote = "true") then%>
			<tr><td class="quoteitemtext">
				<%if not( rsGetQuoteBearingsList is nothing ) then
					if( rsGetQuoteBearingsList.State = adStateOpen ) then
						intRowNumber = 1
						while( not rsGetQuoteBearingsList.EOF )
				%>
							<br><b><%=rsGetQuoteBearingsList( "Qty" )%></b>&nbsp;&nbsp;&nbsp;<%=rsGetQuoteBearingsList( "PartNumber" )%>
				<%
							intRowNumber = intRowNumber + 1
							rsGetQuoteBearingsList.MoveNext
						wend%>
						<br>&nbsp;					
					<%end if
				end if%>  
			</tr>
      <%end if%>
      <tr><td class="quoteitem">Parts</td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalPartCharge" ), 2 )%></td></tr>
      <%if (isFullQuote = "true") then%>
			<tr><td class="quoteitemtext">
				<%if not( rsGetQuotePartsList is nothing ) then
					if( rsGetQuotePartsList.State = adStateOpen ) then
						intRowNumber = 1
						while( not rsGetQuotePartsList.EOF )
				%>
							<br><b><%=rsGetQuotePartsList( "Qty" )%></b>&nbsp;&nbsp;&nbsp;<%=rsGetQuotePartsList( "ProductName" )%>
				<%
							intRowNumber = intRowNumber + 1
							rsGetQuotePartsList.MoveNext
						wend%>
						<br>&nbsp;					
					<%end if
				end if%>      
			</tr>
      <%end if%>
      <tr><td class="quoteitem">Labor</td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalLaborAndHandling" ), 2)%></td></tr>
      <tr><td class="quoteitem">Rework</td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalReworkCharge" ), 2)%></td></tr>
      <%if (isFullQuote = "true") then%>
			<tr><td class="quoteitemtext">
				<%if not( rsGetQuoteSubWorkList is nothing ) then
					if( rsGetQuoteSubWorkList.State = adStateOpen ) then
						intRowNumber = 1
						while( not rsGetQuoteSubWorkList.EOF )
				%>
							<br><%=rsGetQuoteSubWorkList( "SubWorkDescription" )%>
				<%
							intRowNumber = intRowNumber + 1
							rsGetQuoteSubWorkList.MoveNext
						wend%>
						<br>&nbsp;					
					<%end if
				end if%>      
				</tr>
				<%end if%>
      <tr><td>&nbsp;</td><td class="quotetotal">&nbsp;</td></tr>
      <tr><td class="quoteitem">Total Repair ~ <span class="stnd">Standard</span></td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalRepairPrice" ), 2 )%></td></tr>
      <tr><td class="quoteitem">Total Repair ~ <span class="expd">Expedited</span></td><td class="quoteinfo" align="right"><%=FormatCurrency( rsGetQuoteByQuoteId( "TotalExpeditingCharge" ), 2 )%></td></tr>
    </table></td></tr>
    
    <tr><td class="quotetable"><table width="100%">
     <tr><td class="quotetext" width="180">Comments:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "QuoteSpecificComments" )%></td></tr>
    </table></td></tr>
    
    <tr><td class="quotetable"><table width="100%">
      <tr><td class="quotetext" width="200"><span class="stnd">Standard</span> Delivery:</td><td class="quoteitem" width="440"><%=rsGetQuoteByQuoteId( "DeliveryInformation" )%></td></tr>
      <tr><td class="quotetext" width="200"><span class="expd">Expedited</span> Delivery:</td><td class="quoteitem" width="440"><%=rsGetQuoteByQuoteId( "ExpeditedDeliveryInformation" )%></td></tr>
    </table></td></tr>
    
    <tr>
		<td>
			<p>Please indicate your approval and provide P.O. information in the appropriate space below. All repair approvals should be either faxed to (580) 762 4722 or emailed to repair.quote@centerline-inc.com. <strong>Your acknowledgement indicates authority to proceed with the repair and acceptance of Centerline's Terms.</strong></p>
			<p>All amounts are quoted in US Dollars. Payment terms are Net 30 days from Invoice Date with approved credit. Invoices Unpaid Past Terms Are Subject To 5% Service Charge/Month. Freight Is F.O.B. Ponca City, OK (unless contrary agreement is reached before quoting).</p>	
			<p>Quotes are valid for one week, parts pricing and delivery subject to prior sale. Materials left at Centerline without repair approval for a period of six months will be scrapped.</p>		
		</td>
	</tr>
    
    <tr><td class="quotetable"><table width="100%">
      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext">Sincerely yours,</td></tr>
      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext"><%=rsGetQuoteByQuoteId( "QuotedBy" )%>, Sales Representative</td></tr>
      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext">Jason Mabry, Repair Projects Manager</td></tr>
      <tr><td width="300">&nbsp;</td><td class="quotetext">for Centerline, Inc.</td></tr>
    </table></td></tr>
    
    <tr><td class="quoteauthorizationtable" align=middle><table width="500" align=center>
      <tr><td class="quotetitle" align="middle" colspan=4>CUSTOMER AUTHORIZATION FOR WO <%=rsGetQuoteByQuoteId( "WorkOrderNumber" )%></td></tr>
      <tr><td class="quoteinfo">PHONE:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "TelephoneNumber" )%></td>
        <td class="quoteinfo">FAX:</td><td class="quoteinfo"><%=rsGetQuoteByQuoteId( "FaxNumber" )%></td></tr>
    </table></td></tr>
    <tr><td align=middle><table width="500" align=center>
      <tr><td class="quoteinfo">DATE APPROVED:</td><td class="authorizationline" width=110>&nbsp;</td>
        <td class="quoteinfo">P.O.#:</td><td class="authorizationline" width=190>&nbsp;</td></tr>
    </table></td></tr>
    <tr><td align=middle><table width="500" align=center>
      <tr><td class="quoteinfo">APPROVED BY:</td><td class="authorizationline" width=350>&nbsp;</td></tr>
    </table></td></tr>

  </table>
</BODY>
</HTML>
<%	end if
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

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing
set rsGetQuoteByQuoteId				= nothing
set cmdGetQuoteByQuoteId			= nothing
set cnnSQLConnection					= nothing
%>
