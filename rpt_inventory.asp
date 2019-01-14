<HTML>
<HEAD>
<TITLE>Centerline - Select Inventory Report</TITLE>
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
	
	function RefreshCurrentPage()
	{
		document.location.href = '<%=strCurrentPage%>';
	}
	
	function OpenEmployee( strEmployeeId )
	{
	  OpenWindow( 'employee_details.asp?intEmployeeId='+strEmployeeId,'CentEmpl'+intEmployeeId,'width=640,height=160,left='+GetCenterX( 640 )+',top='+GetCenterY( 160 )+',screenX='+GetCenterX( 640 )+',screenY='+GetCenterY( 160 )+',scrollbars=no,dependent=yes' );
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
<table width=640 border=0 cellpadding=0 cellspacing=0 ID="Table1">

<!--Top Title Bar -->
<tr><td class="pagetitle">  
 <TABLE class="sectiontitle" width=100% cellpadding=0 cellspacing=0 ID="Table2">
   <tr><td colspan=2><IMG SRC="images/title_top.gif" width=100% height=6></td></tr>
   <tr><td class="sectiontitle">Inventory Report</td>
     <td align=left width=348><IMG SRC="images/title_right.gif" width=23 height=25></td></tr>
 </TABLE>
</td></tr>
<!-- End Top Title Bar -->

<!-- Customer Form -->
<form id="frmInventoryReport" name="frmInventoryReport" action="" method="POST">
<tr><td align=center>
		<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle ID="Table3">
		<tr><td class="label_1">Product Name</td><td class="input_1"><input type="text" class="forminput" name="txtProductName" maxlength=100 isRequired="true" fieldName="Product Name" onchange="this.isDirty=true;"></td>
		  <td class="label_2">Customer</td><td class="input_2"><input type="text" class="forminput" name="txtCustomer" maxlength=255 fieldName="Customer" onchange="this.isDirty=true;"></td></tr>
		<tr><td class="label_1">Date In</td><td class="input_1"><input type="text" class="forminput" name="txtDateIn" maxlength=100 fieldName="DateIn" onchange="this.isDirty=true;"></td>
		  <td class="label_2">Date Out</td><td class="input_2"><input type="text" class="forminput" name="txtDateOut" maxlength=100 fieldName="DateOut" onchange="this.isDirty=true;"></td></tr>
		</table>
</td></tr>

<!-- Bottom Rule -->
<tr><td height=12>
  <TABLE class="" width=100% cellpadding=0 cellspacing=0 ID="Table6">
    <tr><td><IMG SRC="images/small_title_top.gif" width=100% height=3></td></tr>
  </TABLE>
</td></tr>
<!-- End Bottom Rule -->
 
<!-- Form Buttons -->
<tr><td align=center>
   	<table class="form" width=600 border=0 cellpadding=0 cellspacing=0 align=middle ID="Table7">
	    <tr><td colspan=4><table cellspacing=0 cellpadding=0 width=192 align=right ID="Table8">
	          <tr><td width=64 align=right><a href="javascript:validateForm(frmInventoryReport);"><IMG src="images/save_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:frmInventoryReport.reset();"><IMG src="images/clear_button.gif" width=60 height=20 border=0></a></td>
	            <td width=64 align=right><a href="javascript:GetDirty(frmInventoryReport)"><IMG src="images/close_button.gif" width=60 height=20 border=0></a></td></tr>
	     </table></td></tr>
	 </table>
</td></tr>  <!-- End Form Buttons -->
</form>  <!-- End Customer Form -->
</table>
</body>
</html>