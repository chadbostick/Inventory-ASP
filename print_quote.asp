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
        <tr><td class="contact_info" width=100%>
         <table>
          <tr><td colspan=2 align="right"><%=rsGetCompanyInformation( "CompanyName" )%></td></tr>
          <tr><td colspan=2 align="right"><%=rsGetCompanyInformation( "Address" )%></td></tr>
          <tr><td colspan=2 align="right"><%=rsGetCompanyInformation( "City" )%>,&nbsp;<%=rsGetCompanyInformation( "StateOrProvince" )%>&nbsp;<%=rsGetCompanyInformation( "PostalCode" )%></td></tr>
          <tr><td colspan=2 align="right">Phone:&nbsp;<%=rsGetCompanyInformation( "PhoneNumber" )%></td></tr>
          <tr><td colspan=2 align="right">Fax:&nbsp;<%=rsGetCompanyInformation( "FaxNumber" )%></td></tr>
         </table></td></tr>
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
			<!--<p>All amounts are quoted in US Dollars. Payment terms are Net 30 days from Invoice Date with approved credit. Invoices Unpaid Past Terms Are Subject To 5% Service Charge/Month. Freight Is ExWorks (EXW) Ponca City, OK (unless contrary agreement is reached before quoting).</p>	
			<p>Quotes are valid for one week, parts pricing and delivery subject to prior sale. Materials left at Centerline without repair approval for a period of six months will be scrapped.</p>//-->
			<p>The attached Centerline, Inc.'s Terms and Conditions of Sale and Repair are a binding part of any agreement between Centerline and Customer.  All sales and repairs are subject to those provisions.  They should be examined by the Customer and may not be amended or altered by Customer without the express written consent of Centerline.  The Terms and Conditions of Sale and Repair are also available at http://pubs.centerline-inc.com/terms/</p>		
		</td>
	</tr>
    
    <tr><td class="quotetable"><table width="100%">
      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext">Sincerely yours,</td></tr>
      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext"><%=rsGetQuoteByQuoteId( "QuotedBy" )%><!--, Sales Representative//--></td></tr>
<!--      <tr><td width="300" class="quotetext">&nbsp;</td><td class="quotetext">Jason Mabry, Repair Projects Manager</td></tr>//-->
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
	
	<tr><td><table id="terms">
      <tr><td>
	  
<p class="page-break" align="center">
<p align="center">
    <strong><u>CENTERLINE, INC.</u></strong>
</p>
<p align="center">
    <strong>Terms and conditions of sale and</strong>
    <strong> REPAIR</strong>
</p>
<p>
    1. <u>Definitions:</u>
</p>
<p>
    A. "Company" means Centerline, Inc.
</p>
<p>
    B. "Customer" means the purchaser of goods and/or services from the
    Company.
</p>
<p>
    C. "Goods" means articles, goods and services to which this document
    applies, including repairs and replacement of Goods.
</p>
<p>
    2. <u>Contract:</u>
</p>
<p>
    These Terms and Conditions of Sale and repair will apply to all sale,
    repair and service contracts and transactions between Company and Customer
    ("Contract(s)"). A Contract shall be formed upon the Company's receipt of
    Customer's purchase order or other written or oral approval which such
    purchase order must be in compliance with the provisions and specifics of
    Company's quotation. No variation, waiver, amendment or addition to the
    Company's quotation or these Terms and Conditions shall be valid unless
    agreed in writing by an authorized representative of the Company. If the
    Customer transmits its own terms and conditions to Company as a part of a
    contract, and if any of the Customer's terms and conditions vary or
    conflict with the Company's quotation, these Terms and Conditions or other
    written agreement between the parties, the Company's terms and provisions,
    including these Terms and Conditions, shall be paramount and prevailing to
    resolve any such conflict or inconsistency.
</p>
<p>
    3. <u>Estimates:</u>
</p>
<p>
    Initial estimates of repairs, costs and charges from Company are not
    contractual. Company cannot provide a quote on Goods to be repaired until
    it has received the Goods and had an opportunity to fully inspect them. If
    Company receives Goods from Customer based on an estimate, and after
    inspection by Company and issuance of a quote, the Customer shall
    thereafter reject the quote, the Customer shall nevertheless pay Company at
    its normal shop rates for all work performed and parts supplied in the
    inspection and analysis of Customer's Goods, unless such payment is
    otherwise waived by Company. After payment is received from Customer,
    Company will return the Goods to Customer with transportation costs paid by
    Customer.
</p>
<p>
    4. <u>Quotations:</u>
</p>
<p>
    All quotations (quotes) are valid for seven (7) days from the date of the
    quotation, unless stated otherwise in the quotation or earlier revoked by
    Company. No quotation by the Company nor the publication by Company of any
    other document shall place Company under any duty or liability to the
    Customer until, the Customer's purchase order or other valid acceptance is
    received, and further subject to mutual approval of written drawings,
    specifications and submittals as described in paragraph 7. Provided
    further, notwithstanding the issuance of Customer's purchase order or other
    agreement of sale or repair between the parties, Company may nevertheless
    thereafter extend the final delivery date subject to availability of any
    parts required from vendors and/or increase the price on the quotation if
    the cost of any required parts are increased by Company's vendors after the
    date of the quote. If it is necessary for Company to extend the delivery
    date or increase the price on its quotation it shall promptly notify
    Customer of such event, at which time, Customer shall have 5 business days
    to elect to cancel the order. Provided that, if Customer timely elects to
    cancel the order, it shall remain obligated to pay Company for all
    completed Goods delivered to Customer and for all work performed and parts
    purchased by Company for the Customer's job. Company will be paid its shop
    rates for all work performed and parts supplied.
</p>
<p>
    5. <u>Price:</u>
</p>
<p>
    Unless stated otherwise on Company's quotation:
</p>
<p>
    A. All prices quoted are exclusive of installation, freight, packaging and
    taxes.
</p>
<p>
    B. All transportation of Goods is EXWORKS at Customer's cost.
</p>
<p>
    C. The Customer shall be obligated for and shall pay directly, either to
    the taxing authority or by way of reimbursement to Company, for all excise
    taxes, use taxes, sales taxes, business privilege taxes or other
    assessments which may be due and owing to any federal, state or local
    taxing authorities by reason of the sale of any Goods to Customer, and
    Customer agrees to indemnify and hold Company harmless therefrom. All of
    the foregoing shall be in addition to the quoted sales price.
</p>
<p>
    D. The contract price shall be paid by the Customer within thirty (30) days
    of the invoice date (unless stated otherwise by Company). Any goods
    delivered in separate installments may be billed by Company in separate
    installments and payment shall be due accordingly.
    <strong>
        Upon the event that any payment due by Customer is not paid when due,
        the amount outstanding shall bear interest at the rate of 1.5% per
        month until paid.
    </strong>
</p>
<p>
    6. <u>Cancellation and Modification of Orders:</u>
</p>
<p>
    Except as otherwise provided in paragraph 3 above, no order may be
    cancelled or modified by the Customer unless requested in writing and
    accepted by Company in writing, including terms of repricing and delivery.
    Cancellation charges may apply and will be advised by the Company.
</p>
<p>
    7. <u>Delivery:</u>
</p>
<p>
    A. The delivery dates stated are only approximate, and are not conditions
    of the Contract and are contingent upon the receipt of all payments due and
    information required to proceed with the order without delay. Any delay by
    Customer, its contractors or its representatives in supplying Company with
    plans, specification, directions or other information required to complete
    and deliver the Goods will extend the delivery date by an amount of time
    equal to such delay on the part of Customer. The final delivery date will
    be established by agreement between Company and Customer after a compliant
    purchase order or other confirmation is received from Customer, Company has
    had an opportunity to fully inspect any goods to be repaired and both
    parties have approved all drawing, depictions and other submittals which
    describes the work to be performed, and which have been approved by both
    parties.
</p>
<p>
    B. No order may be delayed or rescheduled by the Customer unless agreed to
    by the Company in writing.
</p>
<p>
    C. In the event of such agreed delay or rescheduling, the price of the
    Goods shall be subject to increase, and the Customer may pay any storage
    charges or other charges required to safeguard the Goods. The Customer
    agrees to pay the Company's invoices for the Goods (or, if not completed,
    that portion of the Goods which is ready for delivery) as though delivery
    were made on the original estimated date of delivery.
</p>
<p>
    8. <u>Warranties and Disclaimers:</u>
</p>
<p>
    IN REGARD TO ANY GOODS WHICH CONSTITUTES NEW PARTS OR FINISHED PRODUCTS
    WHICH COMPANY HAS ACQUIRED FROM A THIRD PARTY, SUCH AS A DISTRIBUTOR, AND
    TO THE EXTENT ASSIGNABLE, COMPANY AGREES TO ASSIGN TO CUSTOMER ALL OF THE
    ORIGINAL MANUFACTURER'S WARRANTIES ON SUCH GOODS. COMPANY PROVIDES NO
    WARRANTY OF ITS OWN ON ANY SUCH ITEMS. COMPANY WARRANTS ITS WORK, ASSEMBLY,
    REPAIRS AND OTHER SERVICES PROVIDED BY COMPANY AGAINST DEFECTS IN
    WORKMANSHIP FOR A PERIOD OF ONE YEAR FROM COMPLETION OF COMPANY'S SERVICES
    (BASED ON THE SHIPPING DATE TO CUSTOMER ) OR FOR 2,000 HOURS OF USE OF THE
    GOODS BY CUSTOMER, WHICHEVER FIRST OCCURS. ALL ASSIGNABLE WARRANTIES,
    INCLUDING LIMITATIONS AND DISCLAIMERS, FROM THE ORIGINAL MANUFACTURERS OR
    OTHER OEM SHALL APPLY IN REGARD TO THIRD PARTY PARTS INCORPORATED INTO
    REPAIR WORK.
    <strong>
        Customer's remedies under ANY WARRANTY HEREIN MADE BY COMPANY are
        limited to repair or replacement of the goods by company, at the
        company's premises or the place of installation of the goods, at
        Company's option, of any defective workmanship of the company. Any
        repair or replacement of the goods by any party other than company
        shall void this warranty. If it is determined that the warranty has
        been properly invoked by customer and that company is liable
        thereunder, company will pay for any freight and shipping costs of
        delivering the goods to and between customer's and company's place of
        business. Other than the foregoing, company makes no warranty of any
        kind whatsoever, expressed or implied, in regard to the services or
        goods. All implied warranties of merchantability and fitness for
        particular purpose are hereby disclaimed by company and are excluded
        from this agreement. Company specifically disclaims liability for
        consequential and incidental damages, direct or indirect, loss of use
        and loss of profits. The ability of company to disclaim warranties may
        be limited in some jurisdictions. To the extent that any limitations
        should apply, the remainder of all other restrictions and limitations
        set forth herein shall nevertheless apply. No prior or subsequent oral
        statement, marketing LITERATURE or representation made by company or
        any of its agents, employees or representatives shall be binding upon
        company and this agreement constitutes the exclusive and entire
        agreement between company and customer regarding warranty.
    </strong>
</p>
<p>
    <strong>
        The warranty is specifically conditioned upon the following:
    </strong>
</p>
<p>
    <strong>
        A. Receipt of written notice from the Customer of claimed defects
        within the warranty period.
    </strong>
</p>
<p>
    <strong> </strong>
    <strong>
        B. Reasonable opportunity of Company to inspect and repair such
        defects.
    </strong>
</p>
<p>
    <strong> C. Payment of the entire contract price. </strong>
</p>
<p>
    <strong>
        D. The Goods having been installed, erected or operated in conformity
        with any instructions provided to Customer by Company, and/or any OEM
        instructions, if applicable, and the Goods having been used in normal
        use and service for the purpose for which they are designed, having
        received normal and periodic maintenance, having not been subjected to
        misuse, negligence or accident and having not been altered or repaired
        by anyone other than Company's representatives in any respect with
        affects its condition or operation.
    </strong>
</p>
<p>
    <strong>
        E. The Company's obligation pursuant to this warranty against defects
        in workmanship shall terminate if Customer undertakes repair or
        replacement of alleged defective parts without the prior written
        consent of the Company.
    </strong>
</p>
<p>
    <strong>
        F. The Company shall not be held responsible for errors in drawings,
        specifications or samples after they have been submitted and approved
        by the Customer or its representative.
    </strong>
</p>
<p>
    9. <u>Title, Security and Risk of Loss Interest:</u>
</p>
<p>
    All new Goods sold by Company to Customer, which are not otherwise
    incorporated into Goods already owned by Customer, shall remain the sole
    and absolute property of the Company as legal and equitable owner until
    such time as Customer has paid Company the full purchase price due for the
    Goods, together with the full price of any other Goods subject to any other
    contract between Company and Customer. Notwithstanding the foregoing,
    Customer hereby grants to Company a purchase money security interest in the
    Goods. The security interest shall include all additions, accessions and
    substitutions thereto and therefore and all accessories, parts and
    equipment now or hereafter affixed thereto or used in connection therewith.
    The security interest shall also include the proceeds and products of the
    Goods, including insurance proceeds, and all money and property owned by
    Customer which hereafter may be possessed or controlled by Company. This
    security interest is given to secure all of the obligations of Customer to
    Company including the complete payment and satisfaction of all purchase and
    sale orders between the parties, plus all interest due thereon and all
    expenditures by Company involving the performance of or enforcement of any
    agreement, covenant or warranty given by Customer in favor of Company, all
    costs, attorney fees and other expenditures of Company in the collection
    and enforcement of any obligation or liability of Customer to Company and
    in the collection and enforcement of and realization upon Company's
    security interest in the Goods. Customer agrees to execute one or more
    financing statements or other instruments of encumbrance and perfection in
    any form satisfactory to Company, in order to perfect or continue
    perfection of the security interest of Company which arises hereunder, and
    for all jurisdictions in which Company determines filing or other
    perfection of such interest is necessary or desirable. Notwithstanding the
    timing of the transfer of ownership of the Goods, and/or the existence of a
    security interest in the Goods, Customer shall solely bear the risk of loss
    to the Goods, including all damages, diminishment in value, transit claims
    and any and all other injuries or damages which may result to or by the
    Goods from the time that the Goods leave Company's place of business by any
    freight or delivery service.
</p>
<p>
    10. <u>Default:</u>
</p>
<p>
    Customer shall be in default under this agreement should it fail to timely
    make any payment owing to Company under this or any other agreement between
    them, whether or not related to the Goods which are the subject of this
    specific order; or upon the breach, omission or failure of performance by
    Customer of any agreement, covenant or provision contained herein or under
    any other contract or agreement between these parties. The dissolution,
    termination of existence or insolvency of Customer, the appointment of a
    receiver over any part of Customer's property, an assignment for the
    benefit of creditors by Customer or the commencement of any proceeding
    under any bankruptcy or insolvency law by or against Customer shall also
    constitute an event of default by Customer.
</p>
<p>
    11. <u>Company's Remedies:</u>
</p>
<p>
    Upon an event of any breach, omission, failure or default on the part of
    Customer, or upon Company's reasonable belief that circumstances exist or
    have occurred by which its ability to collect any monies due or to become
    due from Customer is impaired, Company shall have all the remedies allowed
    by this contract and by applicable law, including but not limited a claim
    for all damages caused and the right to cease any further production,
    services , repairs or delivery of Goods under this or any other contract or
    agreement between Customer and Company, and all rights to retrieve and sell
    collateral as provided by the Uniform Commercial Code of any state in which
    this agreement is enforceable.
</p>
<p>
    12. <u>Non-waiver:</u>
</p>
<p>
    No act, delay or omission, including Company's waiver of any right or
    remedy because of any known or unknown default hereunder, shall constitute
    a waiver of any of Company's rights and remedies under the law or this
    agreement or under any other agreement between the parties. All rights and
    remedies of Company are cumulative and may be exercised singularly or
    concurrently and the exercise of any one or more remedy will not be a
    waiver of any other, at the same time, or at a later time. No waiver,
    change, modification or discharge of any of Company's rights or if
    Customer's duties will be effective unless in writing and signed by a duly
    authorized officer of Company and any such waiver will not be a bar to the
    exercise of any right or remedy on any subsequent occasion.
</p>
<p>
    13. <u>Governing Law and Venue:</u>
</p>
<p>
    All transactions between Company and Customer shall be interpreted and
    enforced in accordance with the substantive laws of the State of Oklahoma
    including, without limitation, the Oklahoma Uniform Commercial as amended
    from time to time. With respect to any suit, action or proceeding related
    to a sale of Goods under this agreement, or any other agreement between the
    parties, each party irrevocably submits to the exclusive jurisdiction of
    the District Court of Kay County, Oklahoma and the United States District
    Court located in the Western District of Oklahoma, and waives any objection
    which it might have at any time to venue of any proceedings brought into
    any such court, or that such court does not have jurisdiction over such
    party.
</p>
<p>
    14. <u>Severability:</u>
</p>
<p>
    If any provision of this agreement shall be for any reason held to be
    invalid or unenforceable, such invalidity or enforceability shall not
    affect any other provision hereof and this agreement shall be construed as
    if such invalid or unenforceable provision had never been contained herein.
</p>
<p>
    15. <u>Costs and Fees:</u>
</p>
<p>
    The Customer shall be liable for and agrees to pay Company for all costs,
    attorney fees and other expenses and disbursements by Company in the
    enforcement or collection of any invoice, obligation or liability of
    Customer to Company, or the realization upon or the enforcement of
    collection of any collateral in which Company has a security interest.
    Customer agrees promptly to reimburse Company for all such expenditures,
    and until such reimbursement, the amounts of such expenditures shall be
    considered a liability of Customer to Company which is secured by this
    agreement and shall bear interest accordingly.
</p>
<p>
    16. <u>Integration:</u>
</p>
<p>
    This instrument, along with the Customer's purchase order and the Company's
    quotation and sales order, constitute the entire agreement between these
    parties, except as to any written amendments executed by an authorized
    representative of Company. In the event of any conflict between the terms
    of the Customer's purchase order, these Terms and Conditions of Sale, the
    Company's quotation and the Company's sales order, the conflict shall be
    resolved in the following order of preference and priority: (1) these Terms
    and Conditions, (2) Company's quotation and sales order, and (3) Customer's
    purchase order and Customer's terms and conditions. Neither party shall be
    bound by any other prior terms, conditions, statements, representations or
    writings not otherwise herein contained. All previous negotiations,
    statements and preliminary instruments by the parties or their
    representatives are merged in this instrument and the Company's sales
    order.
</p>
<p>
    17. <u>Customer's Indemnification:</u>
</p>
<p>
    The Customer shall defend, indemnify and hold the Company harmless from any
    loss, damage, expense, claim, costs or attorney fees related to or arising
    out of claims for antitrust violations, unfair trade practices or
    infringement of patents, trademarks, copyrights or other intellectual
    property rights resulting from Company's compliance with Customer's design
    specifications, installation instructions, drawings or other directions
    supplied by or on behalf of Customer to Company in relation to the Goods.
</p>
<p>
    18. <u>Force Majeure:</u>
</p>
<p>
    A force majeure delay shall mean any delay caused by, but not limited to,
    an act of God; government action or failure of the government to act; war
    or acts of the public enemy; strike or other labor trouble; fire; floods;
    severe weather; riots or other causes beyond Company's reasonable control,
    provided that any such delay is not caused, in whole or in part by the acts
    or omissions of Company and further that Company is unable to make up for
    such delay with reasonable diligence. If any force majeure delays Company's
    performance, the delivery date or time for compliance will be extended by a
    period of time reasonably necessary to overcome the effect of such delay,
    without penalty to Company.
</p>
<p>
    19. <u>No Third Party Beneficiaries</u>
</p>
<p>
    Nothing in these terms and conditions, or any other agreement between
    Customer and Company, expressed or implied, is intended to confer upon any
    person or entity other than the parties hereto and their respective
    permitted successors and assigns any rights, benefits or obligations
    hereunder.
</p>	  
	  </td></tr>
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
