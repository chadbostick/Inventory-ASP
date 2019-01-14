function replace(string,text,by) {
// Replaces text with by in string
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return string;

    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return string;
    if (i == -1) return string;

    var newstr = string.substring(0,i) + by;

    if (i+txtLength < strLength)
        newstr += replace(string.substring(i+txtLength,strLength),text,by);

    return newstr;
}



function checkdate(objName) {
	var datefield = objName;
	if (chkdate(objName) == false) {
		datefield.select();
		alert("That date is invalid.  Please try again.");
		datefield.focus();
		return false;
	}
	else {
		return true;
	}
}



function chkdate(objName) {
	var strDatestyle = "US"; //United States date style
	//var strDatestyle = "EU";  //European date style
	var strDate;
	var strDateArray;
	var strDay;
	var strMonth;
	var strYear;
	var intday;
	var intMonth;
	var intYear;
	var booFound = false;
	var datefield = objName;
	var strSeparatorArray = new Array("-"," ","/",".");
	var intElementNr;
	var err = 0;
	var strMonthArray = new Array(12);
	
	strMonthArray[0] = "Jan";
	strMonthArray[1] = "Feb";
	strMonthArray[2] = "Mar";
	strMonthArray[3] = "Apr";
	strMonthArray[4] = "May";
	strMonthArray[5] = "Jun";
	strMonthArray[6] = "Jul";
	strMonthArray[7] = "Aug";
	strMonthArray[8] = "Sep";
	strMonthArray[9] = "Oct";
	strMonthArray[10] = "Nov";
	strMonthArray[11] = "Dec";
	strDate = datefield.value;
	
	if (strDate.length < 1) {
		return true;
	}
	
	for (intElementNr = 0; intElementNr < strSeparatorArray.length; intElementNr++) {
		if (strDate.indexOf(strSeparatorArray[intElementNr]) != -1) {
			strDateArray = strDate.split(strSeparatorArray[intElementNr]);
			if (strDateArray.length != 3) {
				err = 1;
				return false;
			}
			else {
				strDay = strDateArray[0];
				strMonth = strDateArray[1];
				strYear = strDateArray[2];
			}
			booFound = true;
		}
	}
	
	if (booFound == false) {
		if (strDate.length>5) {
			strDay = strDate.substr(0, 2);
			strMonth = strDate.substr(2, 2);
			strYear = strDate.substr(4);
		}
		else {			
			err = 1;
			return false;
		}
	}
	
	if (strYear.length == 2) {
		strYear = '20' + strYear;
	}
	
	// US style
	if (strDatestyle == "US") {
		strTemp = strDay;
		strDay = strMonth;
		strMonth = strTemp;
	}
	
	intday = parseInt(strDay, 10);
	if (isNaN(intday)) {
		err = 2;
		return false;
	}
	
	intMonth = parseInt(strMonth, 10);
	if (isNaN(intMonth)) {
		for (i = 0;i<12;i++) {
			if (strMonth.toUpperCase() == strMonthArray[i].toUpperCase()) {
				intMonth = i+1;
				strMonth = strMonthArray[i];
				i = 12;
			}
		}
		if (isNaN(intMonth)) {
			err = 3;
			return false;
		}
	}

	intYear = parseInt(strYear, 10);
	if (isNaN(intYear)) {
		err = 4;
		return false;
	}
	if (intMonth>12 || intMonth<1) {
		err = 5;
		return false;
	}

	if ((intMonth == 1 || intMonth == 3 || intMonth == 5 || intMonth == 7 || intMonth == 8 || intMonth == 10 || intMonth == 12) && (intday > 31 || intday < 1)) {
		err = 6;
		return false;
	}
	if ((intMonth == 4 || intMonth == 6 || intMonth == 9 || intMonth == 11) && (intday > 30 || intday < 1)) {
		err = 7;
		return false;
	}
	if (intMonth == 2) {
		if (intday < 1) {
			err = 8;
			return false;
		}
		if (LeapYear(intYear) == true) {
			if (intday > 29) {
				err = 9;
				return false;
			}
		}
		else {
			if (intday > 28) {
				err = 10;
				return false;
			}
		}
	}

	if (strDatestyle == "US") {
		//datefield.value = strMonthArray[intMonth-1] + " " + intday+" " + strYear;
		datefield.value= intMonth + "/" + intday + "/" + strYear;
	}
	else {
		//datefield.value = intday + " " + strMonthArray[intMonth-1] + " " + strYear;
	}


	return true;
}




function LeapYear(intYear) {
	if (intYear % 100 == 0) {
		if (intYear % 400 == 0) { return true; }
	}
	else {
		if ((intYear % 4) == 0) { return true; }
	}

	return false;
}




function doDateCheck(from, to) {
	if (Date.parse(from.value) <= Date.parse(to.value)) {
		alert("The dates are valid.");
	}
	else {
		if (from.value == "" || to.value == "") 
			alert("Both dates must be entered.");
		else 
			alert("To date must occur after the from date.");
	}
}




function chknumber(objName) {
	var strNumber = objName.value;
	
	if ( (strNumber == null) || (strNumber == "") || isblank(strNumber) ) {
		return true;
	}
	
	var n = parseFloat(strNumber);
	if ( isNaN(n) ) {
		return false;
	}
	else {
		objName.value = n;
		return true;
	}
}




function chkqty(objName) {
	var strQty = objName.value;
	
	if ( (strQty == null) || (strQty == "") || isblank(strQty) ) {
		return true;
	}
	
	var n = parseFloat(strQty);
	if ( isNaN(n) ) {
		return false;
	}
	else {
		var intDecimal = strQty.indexOf(".");
		if ( intDecimal >= 0 )
		{
				var strDecimal = strQty.substring(intDecimal, strQty.length);
				if ( (parseFloat(strDecimal) == 0.0) || (parseFloat(strDecimal) == 0.5) )
				{
					return true;
				}
				else
				{
					return false;
				}
		}
		else
		{
			return true;
		}
	}
}



function chkcurrency(objName) {
	var strCurrency = objName.value;
	
	if ( (strCurrency == null) || (strCurrency == "") || isblank(strCurrency) ) {
		return true;
	}
	
	var strNumber  = replace(strCurrency, " ", "");
	strNumber = replace(strNumber, ",", "");
	
	if (strCurrency.charAt(0) == "$") {
		strNumber = strNumber.substring(1, strNumber.length);
	}	
	
	var n = parseFloat(strNumber);
	if ( isNaN(n) ) {
		return false;
	}
	else {
		objName.value = "$" + n;
		return true;
	}
}



function isblank(s) {
	for (var i = 0; i < s.length; i++) {
		var c = s.charAt(i);
		if ((c != ' ') && (c != '\n') && (c != '\t')) return false;
	}
	return true;
}




function GetDirty(thisForm) {
	var errorString = "";
	var strDirtyErrors = "";
	
	for (var i = 0; i < thisForm.length; i++) {
		var thisElement = thisForm.elements[i];
		
		if (thisElement.isDirty) {
				strDirtyErrors += '       ' + thisElement.fieldName + '\n';
		}
	}
	
	if ( strDirtyErrors ) {
		errorString += "\n- The following field(s) were changed:\n";
		errorString += strDirtyErrors + "\n";
	}	
	
	if ( errorString ) {
		var msg = "____________________________________________________\n\n";
		msg += "One or more fields were changed.\n";
		msg += "Would you like to save these fields before closing?\n";
		msg += "____________________________________________________\n";
		msg += errorString;
		//alert( msg );
		
		if (confirm(msg))
			validateForm(thisForm);
		else
			window.close();
	}
	else
		window.close();
}




function validateForm(thisForm) {
	var errorString = "";
	var strQtyErrors = "";
	var strDateErrors = "";
	var strRequiredErrors = "";
	var strNumberErrors = "";
	var strCurrencyErrors = "";
	var strDirtyErrors = "";
	
	for (var i = 0; i < thisForm.length; i++) {
		var thisElement = thisForm.elements[i];
		
		if (thisElement.isQty) {
			if ( chkqty(thisElement) == false ) {
				strQtyErrors += '       ' + thisElement.fieldName + '\n';
			}
		}		
		
		if (thisElement.isDate) {
			if ( chkdate(thisElement) == false ) {
				strDateErrors += '       ' + thisElement.fieldName + '\n';
			}
		}
		
		if (thisElement.isRequired) {
			if ( (thisElement.value == null) || (thisElement.value == "") || isblank(thisElement.value) ) {
				strRequiredErrors += '       ' + thisElement.fieldName + '\n';
			}
		}
		
		if (thisElement.isNumber) {
			if ( chknumber(thisElement) == false ) {
				strNumberErrors += '       ' + thisElement.fieldName + '\n';
			}
		}
		
		if (thisElement.isCurrency) {
			if ( chkcurrency(thisElement) == false ) {
				strCurrencyErrors += '       ' + thisElement.fieldName + '\n';
			}
		}
		
		if (thisElement.isDirty) {
				strDirtyErrors += '       ' + thisElement.fieldName + '\n';
		}
	}
	
	if ( strRequiredErrors ) {
		errorString += "\n- The following field(s) are required:\n";
		errorString += strRequiredErrors + "\n";
	}
	
	if ( strNumberErrors ) {
		errorString += "\n- The following field(s) are invalid numbers:\n";
		errorString += strNumberErrors + "\n";
	}
	
	if ( strCurrencyErrors ) {
		errorString += "\n- The following field(s) are invalid currencies:\n";
		errorString += strCurrencyErrors + "\n";
	}
	
	if ( strDateErrors ) {
		errorString += "\n- The following field(s) are invalid dates:\n";
		errorString += strDateErrors + "\n";
	}
	
	if ( strQtyErrors ) {
		errorString += "\n- Quantity fields must be in whole (x.0) or half (x.5):\n";
		errorString += "\n- The following field(s) are invalid quantities:\n";
		errorString += strQtyErrors + "\n";
	}
	
	if ( strDirtyErrors ) {
		//errorString += "\n- The following field(s) were changed:\n";
		//errorString += strDirtyErrors + "\n";
	}
	
	
	if ( errorString ) {
		var msg = "____________________________________________________\n\n";
		msg += "The form was not submitted because of the following error(s).\n";
		msg += "Please correct the error(s) and re-submit.\n";
		msg += "____________________________________________________\n";
		msg += errorString;
		alert( msg );
	}
	else {
		thisForm.submit();
	}
}