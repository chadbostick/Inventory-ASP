<%option explicit

dim cnnSQLConnection
dim rsReportOpenProjects, cmdReportOpenProjects
dim strPrint

set cnnSQLConnection			= nothing
set rsReportOpenProjects	= nothing
set cmdReportOpenProjects	= nothing
set strPrint = nothing

strPrint = Request.QueryString( "Print" )

set cnnSQLConnection			= Server.CreateObject( "ADODB.Connection" )
set cmdReportOpenProjects	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

cmdReportOpenProjects.ActiveConnection	= cnnSQLConnection
cmdReportOpenProjects.CommandText				= "{call spReportOpenProjects }"

on error resume next
set rsReportOpenProjects = cmdReportOpenProjects.Execute
if( Err.number <> 0 ) then
	Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
	Response.Write( Err.Description )
	Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
	Response.Write( "</font></body></html>" )
	Response.End
end if

on error goto 0%>
<HTML>
<TITLE>Open Projects</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/reports.css" TITLE="tables_css">
  <script type="text/javascript" src="js/common.js"></script>
  <script language="Javascript">
  function OpenProject( strProjectId )
	{
	  OpenWindow( 'project_details.asp?intProjectId='+strProjectId,'CenterlineProject'+strProjectId,'width=657,height=540,left='+GetCenterX( 657 )+',top='+GetCenterY( 540 )+',screenX='+GetCenterX( 657 )+',screenY='+GetCenterY( 540 )+',scrollbars=yes,dependent=yes' );
	}
	</script>
</HEAD>

<BODY<%if strPrint = "true" then%> onload="javascript: if (window.print) window.print();"<%end if%>>
  <table align=left width=800>
  
    <tr><td><table width=800 class="form_header">
      <tr><td valign=top><img src="images/centerline_logo_s.jpg" width=149 height=73></td>
       <td align=right><table>
        <tr><td class="internal_report_title">Open Projects</td></tr>
        </table></td></tr>
    </table></td></tr>
    
    <tr><td class="internal_report_info"><table cellspacing=0 cellpadding=4 width=800>
      
      <tr><td class="internal_report_label" align="center" width=50>Id</td>
        <td class="internal_report_label" width=180>Project Name</td>
        <td class="internal_report_label" width=310>Project Description</td>
        <td class="internal_report_label" width=70>Start Date</td>
        <td class="internal_report_label" width=150>Customer</td>
        <td class="internal_report_label" width=40>Priority</td>
      </tr>
      
      <%if strPrint <> "true" then%>
      <tr><td colspan=6>
				<div><!--  style="height:420px; overflow: auto;"  //-->
				<table width="100%" border=0 cellspacing = 0 cellpadding=4 ID="Table1">
			<%end if%>
              
        <%if( not rsReportOpenProjects is nothing ) then
						if( rsReportOpenProjects.State = adStateOpen ) then
							while( not rsReportOpenProjects.EOF )%>
								<tr><td class="internal_report_group_name" align="center" valign=baseline width=50><a href="javascript:OpenProject( '<%=rsReportOpenProjects( "ProjectId" )%>' );"><%=rsReportOpenProjects( "ProjectId" )%></a></td>
								  <td class="internal_report_group_name" valign=baseline width=180><%=rsReportOpenProjects( "ProjectName" )%></td>
								  <td class="internal_report_group_name" valign=baseline width=310><%=rsReportOpenProjects( "ProjectDescription" )%></td>
								  <%if( IsNull( rsReportOpenProjects( "StartDate" ) ) ) then%>
								  <td class="internal_report_group_name" valign=baseline width=70>&nbsp;</td>
								  <%else%>
								  <td class="internal_report_group_name" valign=baseline width=70><%=rsReportOpenProjects( "StartDate" )%></td>
								  <%end if%>
								  <td class="internal_report_group_name" valign=baseline width=150><%=rsReportOpenProjects( "Customer" )%></td>
								  <td class="internal_report_group_name" valign=baseline width=40><b><%=rsReportOpenProjects( "ProjectPriority" )%></b></td>
								<%rsReportOpenProjects.MoveNext
							wend
					end if
				end if%> 
			<%if strPrint <> "true" then%>
				</table></div>
			</td></tr>
			<%end if%> 
    </table></td></tr>
  </table>
</BODY>
</HTML>
<%
if not( rsReportOpenProjects is nothing ) then
	if( rsReportOpenProjects.State = adStateOpen ) then
		rsReportOpenProjects.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsReportOpenProjects		= nothing
set cmdReportOpenProjects		= nothing
set cnnSQLConnection				= nothing
%>
