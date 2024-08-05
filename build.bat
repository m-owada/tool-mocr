if exist app.ico (
  C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:winexe mocr.cs /win32icon:app.ico ^
  /r:C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll ^
  "/r:ref\Windows.Foundation.FoundationContract.winmd" "/r:ref\Windows.Foundation.UniversalApiContract.winmd" %*
) else (
  C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /t:winexe mocr.cs ^
  /r:C:\Windows\Microsoft.NET\assembly\GAC_MSIL\System.Runtime\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Runtime.dll ^
  "/r:ref\Windows.Foundation.FoundationContract.winmd" "/r:ref\Windows.Foundation.UniversalApiContract.winmd" %*
)
if not %errorlevel% == 0 (
  pause
)
