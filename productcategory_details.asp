<%option explicit

dim cnnSQLConnection
dim rsGetProductCategoryByProductCategoryId, cmdGetProductCategoryByProductCategoryId, cmdEditProductCategory
dim intRowNumber, intProductCategoryId

set cnnSQLConnection													= nothing
set rsGetProductCategoryByProductCategoryId		= nothing
set cmdGetProductCategoryByProductCategoryId	= nothing
set cmdEditProductCategory										= nothing

set cnnSQLConnection													= Server.CreateObject( "ADODB.Connection" )
set cmdGetProductCategoryByProductCategoryId	= Server.CreateObject( "ADODB.Command" )

cnnSQLConnection.Open Application( "ODBC_DSN" )
cnnSQLConnection.CursorLocation = adUseClient

if( Request.Form.Count > 0 ) then

	set cmdEditProductCategory	= Server.CreateObject( "ADODB.Command" )
	
	cmdEditProductCategory.ActiveConnection	= cnnSQLConnection
	cmdEditProductCategory.CommandText				= "{? = call spEditProductCategory( ?,?,? ) }"
	
	cmdEditProductCategory.Parameters.Append cmdEditProductCategory.CreateParameter( "@Return", adInteger, adParamReturnValue, 4 )
	cmdEditProductCategory.Parameters.Append cmdEditProductCategory.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
	cmdEditProductCategory.Parameters.Append cmdEditProductCategory.CreateParameter( "@ProductCategoryId", adInteger, adParamInput, 4 )
	cmdEditProductCategory.Parameters.Append cmdEditProductCategory.CreateParameter( "@ProductCategory", adVarChar, adParamInput, 50 )
		
	if( Request.Form( "txtProductCategoryId" ) <> "" ) then cmdEditProductCategory.Parameters( "@ProductCategoryId" ) = Request.Form( "txtProductCategoryId" )
	if( Request.Form( "txtProductCategory" ) <> "" ) then cmdEditProductCategory.Parameters( "@ProductCategory" ) = Request.Form( "txtProductCategory" )
		
	on error resume next
	cmdEditProductCategory.Execute
	if( Err.number <> 0 ) then
		Response.Write( "<html><body><font face=Verdana size=2>A server error has occurred on this page, below is the technical information regarding this error:<br><br>" )
		Response.Write( Err.Description )
		Response.Write( "<br><br><center>Please click the button below and re-check your form submittal</center><br><br><center><a href=""javascript:window.history.back();""><img src=images/back_button.gif width=60 height=20 border=0></a></center>" )
		Response.Write( "</font></body></html>" )
		Response.End
	end if
							
	on error goto 0

	intProductCategoryId = cmdEditProductCategory.Parameters( "@Return" ).Value
else
	intProductCategoryId = Request.QueryString( "intProductCategoryId" )
end if

cmdGetProductCategoryByProductCategoryId.ActiveConnection = cnnSQLConnection
cmdGetProductCategoryByProductCategoryId.CommandText = "{call spGetProductCategoryByProductCategoryId( ?,? ) }"
cmdGetProductCategoryByProductCategoryId.Parameters.Append cmdGetProductCategoryByProductCategoryId.CreateParameter( "@ErrorMessage", adVarChar, adParamOutput, 255 )
cmdGetProductCategoryByProductCategoryId.Parameters.Append cmdGetProductCategoryByProductCategoryId.CreateParameter( "@ProductCategoryId", adInteger, adParamInput, 4 )
cmdGetProductCategoryByProductCategoryId.Parameters( "@ProductCategoryId" )	= intProductCategoryId

on error resume next
set rsGetProductCategoryByProductCategoryId = cmdGetProductCategoryByProductCategoryId.Execute
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
<TITLE>Centerline - Add Product Category</TITLE>
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/default.css" TITLE="default_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/titles.css" TITLE="titles_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/forms.css" TITLE="forms_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/tables.css" TITLE="tables_css">
  <LINK REL=STYLESHEET TYPE="text/css" HREF="css/pagelayout.css" TITLE="layout_css">
  <script type="text/javascript" src="js/forms.js"></script>
</HEAD>
<BODY topmargin=0 leftmargin=0>

<!-- Page Contents -->
<table width=640 border=0 cellpadding=0 cellspacing=0>

<%
if not( rsGetProductCategoryByProductCategoryId is nothing ) then
	if( rsGetProductCategoryByProductCategoryId.State = adStateOpen ) then
		if( not rsGetProductCategoryByProductCategoryId.EOF ) then%>
<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <table class="sectiontitle" width=100% cellpadding=0 cellspacing=0>
   <tr><td colspan=3><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <%if( intProductCategoryId = 0 ) then%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;New Product Category</td>
   <%else%>
   <tr><td class="sectiontitle">&nbsp;&nbsp;<%=rsGetProductCategoryByProductCategoryId( "ProductCategory" )%></td>
   <%end if%>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </table>
</td></tr>  <!-- End Top Title Bar -->

<!-- Shipping Method Form -->
<tr><td align=center>
 <form id="frmProductCategory" name="frmProductCategory" action="" method="post">
 <input type="hidden" name="txtProductCategoryId" value="<%=intProductCategoryId%>">
 <table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
   <tr><td class="label_1">Product Category</td><td class="input_1"><input type="text" class="forminput" name="txtProductCategory" onchange="this.isDirty=true;" maxlength=50 isRequired="true" fieldName="Product Category" value="<%=rsGetProductCategoryByProductCategoryId( "ProductCategory" )%>"></td></tr>
   </table>
 
  <!-- Bottom Rule -->
 <tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0>
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
 </td></tr>  <!-- End Bottom Rule -->
 
  <tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle>
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right>
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmProductCategory);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmProductCategory.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmProductCategory);"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
 </td></tr>
 
 </form>   <!-- End Shipping Method Form -->
 <%	end if
	end if
end if%>
</body>
</html>
<%
if not( rsGetProductCategoryByProductCategoryId is nothing ) then
	if( rsGetProductCategoryByProductCategoryId.State = adStateOpen ) then
		rsGetProductCategoryByProductCategoryId.Close
	end if
end if

if not( cnnSQLConnection is nothing ) then
	cnnSQLConnection.Close
end if

set rsGetProductCategoryByProductCategoryId		= nothing
set cmdGetProductCategoryByProductCategoryId	= nothing
set cmdEditProductCategory										= nothing
set cnnSQLConnection													= nothing
%>