<%option explicit

dim cnnSQLConnection
dim rsGetCompanyInformation, cmdGetCompanyInformation

set cnnSQLConnection					= nothing
set rsGetCompanyInformation		= nothing
set cmdGetCompanyInformation	= nothing

set cnnSQLConnection					= Server.CreateObject( "ADODB.Connection" )
set cmdGetCompanyInformation	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

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

if not( rsGetCompanyInformation is nothing ) then
	if( rsGetCompanyInformation.State = adStateOpen ) then
		if( not rsGetCompanyInformation.EOF ) then%>
<HTML>
<TITLE>Receiving Inspection Sheet</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; TEXT-DECORATION: underline }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder }
	.smalllabel { BORDER-TOP: black solid thin; FONT-SIZE: 80%; FONT-WEIGHT: bolder }
	.inspectionitem { PADDING-TOP: 15px }
	.inspectiontable { BORDER-BOTTOM: black solid thin }
  </STYLE>
</HEAD>

<BODY>
  <table align=left width=640>
  
    <tr><td><table width=640 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo.jpg" width=282 height=138></td>
       <td align=right><table>   
        <tr><td class="contact_info" width=100%>
         <table>
          <tr><td colspan=2><%=rsGetCompanyInformation( "CompanyName" )%></td></tr>
          <tr><td colspan=2><%=rsGetCompanyInformation( "Address" )%></td></tr>
          <tr><td><%=rsGetCompanyInformation( "City" )%>,&nbsp;<%=rsGetCompanyInformation( "StateOrProvince" )%>&nbsp;<%=rsGetCompanyInformation( "PostalCode" )%></td><td align=middle><%=rsGetCompanyInformation( "Country" )%></td></tr>
          <tr><td>Phone:&nbsp;<%=rsGetCompanyInformation( "PhoneNumber" )%>&nbsp;&nbsp;&nbsp;</td><td align=right>Fax:&nbsp;<%=rsGetCompanyInformation( "FaxNumber" )%></td></tr>
          <tr><td>&nbsp;</tr>
          <tr><td colspan=2>Date: _______ / _______ / __________</td></tr>
         </table></td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="title" align=middle>RECEIVING INSPECTION SHEET</td></tr>
    
    <tr><td class="inspectiontable"><table width=640>
	  <tr><td class="smalllabel" colspan=2 align="middle">Spindle Information</td></tr>
      <tr><td class="inspectionitem">Spindle Type:________________________________</td>
        <td class="inspectionitem">Shipping Weight:_______________________________</td></tr>
      <tr><td class="inspectionitem">Serial Number:_______________________________</td>
        <td class="inspectionitem">Project Number:________________________________</td></tr>
        
	  <tr><td class="smalllabel" colspan=2 align="middle">Received From</td></tr>
      <tr><td class="inspectionitem">Company:___________________________________</td>
        <td class="inspectionitem">Contact Name:_________________________________</td></tr>
      <tr><td class="inspectionitem">Address:____________________________________</td>
        <td class="inspectionitem">City:__________________________________________</td></tr>
      <tr><td class="inspectionitem">State:_____________ Zip:______________________</td>
        <td class="inspectionitem">Phone:________________________________________</td></tr>
      <tr><td class="inspectionitem" colspan=2>Received Via:_______________________________________________________________________________</td></tr>
      
	  <tr><td class="smalllabel" colspan=2 align="middle">Shipped To (Same as received if left blank)</td></tr>
      <tr><td class="inspectionitem">Company:___________________________________</td>
        <td class="inspectionitem">Contact Name:_________________________________</td></tr>
      <tr><td class="inspectionitem">Address:____________________________________</td>
        <td class="inspectionitem">City:__________________________________________</td></tr>
      <tr><td class="inspectionitem">State:_____________ Zip:______________________</td>
        <td class="inspectionitem">Phone:________________________________________</td></tr>
      <tr><td class="inspectionitem" colspan=2>Received Via:_______________________________________________________________________________</td></tr>
    </table></td></tr>
    
    <tr><td class="inspectionitem">Box Condition:______________________________________________________________________________</td></tr>
    <tr><td class="inspectionitem">Comments (Damaged Freight, Missing Items, Included Items, etc.):___________________________________</td></tr>
    <tr><td class="inspectionitem">___________________________________________________________________________________________</td></tr>
    <tr><td class="inspectionitem">___________________________________________________________________________________________</td></tr>
    <tr><td class="inspectionitem">___________________________________________________________________________________________</td></tr>
    <tr><td class="inspectionitem">___________________________________________________________________________________________</td></tr>
    <tr><td class="inspectionitem">___________________________________________________________________________________________</td></tr>
  </table>
</BODY>
</HTML>
<%end if
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
set cnnSQLConnection					= nothing
%>