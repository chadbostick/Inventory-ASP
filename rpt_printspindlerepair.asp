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
<TITLE>Spindle Repair Information</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <STYLE TYPE="text/css">
    TD { FONT-SIZE: 80%; }
	.title { FONT-SIZE: 110%; FONT-WEIGHT: bolder; PADDING-BOTTOM: 15px; TEXT-DECORATION: underline }
	.label { FONT-SIZE: 90%; FONT-WEIGHT: bolder }
	.smalllabel { BORDER-TOP: black solid thin; FONT-SIZE: 80%; FONT-WEIGHT: bolder }
	.item { PADDING-TOP: 15px }
	.infotable { BORDER-BOTTOM: black solid thin }
	.box { BORDER: black solid 1px }
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
          <tr><td>&nbsp;</tr>
          <tr><td colspan=2>Project#: ________________________</td></tr>
         </table></td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="title" align=middle>SPINDLE REPAIR INFORMATION</td></tr>
    
    <tr><td><table class="infotable" width=640>
      <tr><td class="item">Make of Machine:___________________________</td>
        <td class="item">Make of Spindle:________________________________</td></tr>
      <tr><td class="item">Spindle Type:______________________________</td>
        <td class="item">Serial Number:_________________________________</td></tr>
      <tr><td class="item">Spindle Speed:_____________________________</td>
        <td class="item">Operating Speed:_______________________________</td></tr>
        
	  <tr><td class="smalllabel" colspan=2 align="middle">Lubrication</td></tr>
      <tr><td class="item" colspan=2 align="middle">Please check one</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="box" width="20">&nbsp;</td>
          <td colspan=2>Grease</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Oil Mist</td>
          <td>Settings:_______________________________________________________________________</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Oil Air</td>
          <td>Settings:_______________________________________________________________________</td></tr>
      </table></td></tr>
      
	  <tr><td class="smalllabel" colspan=2 align="middle">Tool Change</td></tr>
      <tr><td class="item" colspan=2 align="middle">Please check one and specify</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Manual</td>
          <td>Specify:______________________________________________________________________</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Automatic</td>
          <td>Specify:______________________________________________________________________</td></tr>
      </table></td></tr>
      
	  <tr><td class="smalllabel" colspan=2 align="middle">Motor Data</td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Belt Driven</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>Motorized</td></tr>
        <tr><td class="box" width="20">&nbsp;</td>
          <td>High Frequency</td></tr>
      </table></td></tr>
      <tr><td colspan=2><table width=640>
        <tr><td>V:__________</td><td>Hz:__________</td><td>HP/kW:__________</td><td>A:__________</td><td>Poles:__________</td><td>Phases:_________</td></tr>
      </table></td></tr>
    </table></td></tr>
    
    
    <tr><td><table width=640>
      <tr><td>Application:___________________________________________</td>
        <td class="box" width="20">&nbsp;</td>
        <td>Horizontal Mount</td>
        <td class="box" width="20">&nbsp;</td>
        <td>Vertical Mount</td></tr>
      <tr><td class="item">Date of Last Rebuild:___________________________________</td></tr>
    </table></td></tr>
    
    <tr><td class="item">Specify Problem:_____________________________________________________________________________</td></tr>
    <tr><td class="item">___________________________________________________________________________________________</td></tr>
    <tr><td class="item">___________________________________________________________________________________________</td></tr>
    <tr><td class="item">Additional Information:_________________________________________________________________________</td></tr>
    
    
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