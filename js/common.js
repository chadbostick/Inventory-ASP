function OpenWindow (URL, WinName, Features) {
  window.open(URL, WinName, Features);
}

function GetCenterX( nFormWidth )
{
	return ( screen.availWidth / 2 ) - ( nFormWidth / 2 );
}

function GetCenterY( nFormHeight )
{
	return ( screen.availHeight / 2 ) - ( nFormHeight / 2 );
}


function OpenPrintBox( sURL, sWin )
{
  OpenWindow( sURL, 'print','width=650,height=480,left=0,top=0,screenX=0,screenY=0,scrollbars=no,dependent=yes' );
}
	
	
function CheckEnter(e)
{
	if (e.keyCode == 13)
	{
		SearchSubmit();
	}
}