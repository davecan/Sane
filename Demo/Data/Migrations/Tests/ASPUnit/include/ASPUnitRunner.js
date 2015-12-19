function ComboBoxUpdate(strSelectorFrameSrc, strSelectorFrameName) 
{	
	document.frmSelector.action = strSelectorFrameSrc;
	document.frmSelector.target = strSelectorFrameName;
	document.frmSelector.submit();
}		
