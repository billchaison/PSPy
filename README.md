# PSPy
Leveraging Python DLL bundled in with Windows applications to execute Python scripts from Powershell.

**This example shows how invoke Python from Powershell to launch calc.exe**

```powershell
Add-Type -TypeDefinition @"
   using System;
   using System.Diagnostics;
   using System.Runtime.InteropServices;
      public static class Py1
      {
         [DllImport("C:\\Program Files\\Rapid7\\Insight Agent\\components\\insight_agent\\3.3.1.20\\python311.dll", SetLastError=true, CharSet = CharSet.Ansi)]
         public static extern void Py_Initialize();
      }
"@

Add-Type -TypeDefinition @"
   using System;
   using System.Diagnostics;
   using System.Runtime.InteropServices;
      public static class Py2
      {
         [DllImport("C:\\Program Files\\Rapid7\\Insight Agent\\components\\insight_agent\\3.3.1.20\\python311.dll", SetLastError=true, CharSet = CharSet.Ansi)]
         public static extern IntPtr PyRun_SimpleString(
         [MarshalAs(UnmanagedType.LPStr)]string lpCommand);
      }
"@

Add-Type -TypeDefinition @"
   using System;
   using System.Diagnostics;
   using System.Runtime.InteropServices;
      public static class Py3
      {
         [DllImport("C:\\Program Files\\Rapid7\\Insight Agent\\components\\insight_agent\\3.3.1.20\\python311.dll", SetLastError=true, CharSet = CharSet.Ansi)]
         public static extern void Py_Finalize();
      }
"@

[Py1]::Py_Initialize()
[Py2]::PyRun_SimpleString("import os")
[Py2]::PyRun_SimpleString('os.system("c:\\Windows\\System32\\calc.exe")')
[Py3]::Py_Finalize()
```

![alt text](https://github.com/billchaison/PSPy/raw/main/pspy.png)
