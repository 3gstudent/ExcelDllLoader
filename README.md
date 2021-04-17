# ExcelDllLoader
Execute DLL via the Excel.Application object's RegisterXLL() method

**Need install Microsoft Office first**

Learn from Ryan Hanson‏ @ryHanson

Link：

https://gist.github.com/ryhanson/227229866af52e2d963cf941af135a52

License: BSD 3-Clause

ExcelDllLoader.js:

- Check if Microsoft Office has been installed
- Download the dll from Github
- Save the dll to %appdata%\Microsoft\Windows\Recent
- Load it via the Excel.Application object's RegisterXLL() method

ExcelDllLoader(Base64decode).js：

- Download the Base64 encoded text from Github
- Base64 decoded and get the calc.dll
- Save the dll to c:\test\calc.dll
- Load it via the Excel.Application object's RegisterXLL() method

**Note:**

After the DLL is loaded, the DLL is automatically deleted.

Like this:

![Alt text](https://raw.githubusercontent.com/3gstudent/ExcelDllLoader/master/1.gif)

But if you change the path that DLL saves(eg: c:\test),the dll will not be automatically deleted.

Like this:

![Alt text](https://raw.githubusercontent.com/3gstudent/ExcelDllLoader/master/2.gif)

:)

Maybe explorer.exe cheats me.

![Alt text](https://raw.githubusercontent.com/3gstudent/ExcelDllLoader/master/3.png)

More details:

https://3gstudent.github.io/Use-Excel.Application-object's-RegisterXLL()-method-to-load-dll

