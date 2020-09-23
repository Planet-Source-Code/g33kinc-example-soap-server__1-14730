Instructions for Installing Example SoapServer

Requirements
1. Web Server that can process ASP Pages (IIS, PWS)
2. Microsoft XML 2.0

Instructions

1. You will find three folders (SoapServer FrontEnd, SoapServer Backend, and ASP)

2. You will need to compile both the SoapServer and SoapDest projects (Both are ActiveX DLL'S) in the SoapServer Backend Folder.

3. Once you have compiled them place the dll files in the windows\system or winnt\system32 directories.

4. Run Regsvr32 against both dll files.  (i.e. Regsvr32 SoapServer.dll)

5. Place soaprequest.asp from the ASP directory into a directory recognized by your webserver (i.e. wwwroot)

6. Run the Client Test from SoapServer FrontEnd Directory)



