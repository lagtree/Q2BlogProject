{
	var searchPhrase = document.getElementById('searchPhrase').value;
	var Worksheet = 'C:\Users\0690182\OneDrive - Osceola County School District\CYBERS\Example.xlsx';
	var Excel = new ActiveXObject('Excel.Application');

	Excel.Visible = false;
	var Excel_file = Excel.Workbooks.Open(Worksheet, null, true, null,
        "abc", null, true, null, null, false, false, null, null, null);

	var range = Excel_file.ActiveSheet.Range('A:A');
	var jsRangeArray = new VBArray(range.Value).toArray();

	var found = false;
	for(cells in jsRangeArray)
	{
		if(jsRangeArray[cells] == searchPhrase)
		{
		   document.getElementById("results").innerHTML = "Found";
		   found = true;
		}
	}

	if(found == false)
	{
	    document.getElementById("results").innerHTML = "Not Found";
	}

	Excel.ActiveWorkbook.Close(true);
	Excel.Application.Quit();
	Excel = null;
}
