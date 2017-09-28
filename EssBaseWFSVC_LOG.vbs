Sub doSleep( vTime  )
  If Not IsNumeric(vTime ) Then _
       Exit Sub

set WScriptShell = CreateObject("WScript.Shell")
  call WScriptShell.Run ("%COMSPEC% /c ping -n 1 -w " & vTime * 1000 & " 127.255.255.254 > nul", WshHide, WAIT_ON_RETURN)  
end Sub

Sub getSleepy
	 doSleep 0.01 
End Sub 

sub myEcho  (LogMessage )
Dim objWScript
    Set objWScript = CreateObject("WScript.Shell")
  objWScript.Echo   LogMessage
end sub

sub writeConole  (LogMessage ) 
  myEcho Chr(13) & getRightTime() & ";" & LogMessage
end sub

sub writeXmlToConsole   (vCurrXMLFile ) 
for each x in vCurrXMLFile.documentElement.childNodes
   myEcho Chr(13) & (x.nodename) & ": " & x.text
   myEcho Chr(13) 
next
end Sub 


function getCookieFile 
    Dim objFolders 
        Set objFolders = CreateObject("WScript.Shell").SpecialFolders
        getCookieFile = objFolders("mydocuments") & "\essribon.csc"
      set objFolders = Nothing    
end function 

function getlogFile 
    Dim objFolders 
        Set objFolders = CreateObject("WScript.Shell").SpecialFolders
        getlogFile = objFolders("mydocuments") & "\essribon.log"
      set objFolders = Nothing    
end function 

function getLog
        Dim objFSO,objFile,vStrFileName
        Set objFSO=CreateObject("Scripting.FileSystemObject")
        vStrFileName = getlogFile
		 if objFSO.FileExists(vStrFileName) then 
		   set objFile =  objFSO.OpenTextFile(vStrFileName, 1)
			vStrFileName = ""   
			Do Until objFile.AtEndOfStream				
					vStrFileName = vStrFileName & objFile.ReadLine 				
			Loop
		   set objFile = Nothing
		 else 
		   vStrFileName = ""      
		 end if 
		    set objFSO = Nothing 
    getLog =  vStrFileName        
end function


function getRightTime 
 Dim vStrTime 
		 vStrTime  =   Year (Now)

		  if ( Month (Now) < 10  ) then 
		   vStrTime = vStrTime &"0" &  Month (Now)
		   else 
		   vStrTime = vStrTime &  Month (Now)
		  end if 

		  if ( Day (Now) < 10  ) then 
		   vStrTime = vStrTime &"0" &  Day (Now)
		   else 
		   vStrTime = vStrTime &  Day (Now)
		  end if 
           vStrTime = vStrTime & " " & Time 
   getRightTime = vStrTime

end function 


sub WriteLog (LogMessage )
        vStrFileName = getlogFile
		call WriteFileLog (vStrFileName,LogMessage)		 
End sub

sub WriteFileLog (vCurrFileName,vCurrLogMessage )
        Dim objFSO,objFile,vStrFileName
        Set objFSO=CreateObject("Scripting.FileSystemObject")        
		 if objFSO.FileExists(vCurrFileName) then 
		    set objFile =  objFSO.OpenTextFile(vCurrFileName, 8, True)', TristateTrue ) 
		 else 
		    set objFile =  objFSO.CreateTextFile(vCurrFileName,true) 	 
		 end if 

		 LogMessage  =  LogMessage & ";" &  getRightTime () 

          'alert LogMessage   
           objFile.WriteLine LogMessage  
         
            objFile.Close
            set objFSO = Nothing 
            set objFile = Nothing 
End sub


sub WriteCookie

        Dim objFSO,objFile,vStrFileName
        Set objFSO=CreateObject("Scripting.FileSystemObject")
        vStrFileName = getCookieFile
		 if objFSO.FileExists(vStrFileName) then 
		     objFSO.deletefile vStrFileName
		 end if 
		 
         set objFile =  objFSO.CreateTextFile(vStrFileName,true) 

        For i = 0 To UBound(vArrSceduleRules)
		 if (len (vArrSceduleRules(i)) > 3 ) then 
           objFile.WriteLine vArrSceduleRules(i)
		  end if 
        Next
 
            objFile.Close
            set objFSO = Nothing 
            set objFile = Nothing 
End sub

sub ReadCookie
    Dim i 
    Dim vStrFileName  
		 vStrFileName =  getCookieFile
		   Dim objFSO,objFile
        Set objFSO=CreateObject("Scripting.FileSystemObject")
            
		 if objFSO.FileExists(vStrFileName) then
			set objFile = objFSO.OpenTextFile(vStrFileName,1)
			    For i = 0 To UBound(vArrSceduleRules)
        			vArrSceduleRules(i) = "" 
    			Next   
                i = 0			 
			Do Until objFile.AtEndOfStream
				
				vStrFileName = objFile.ReadLine 
				if len (vStrFileName) > 3 then 
				  vArrSceduleRules(i) = vStrFileName  				    
				  i=i+1        
				end if 
				
			Loop
				objFile.Close
                vArrSceduleRulesID = i

				set objFSO = Nothing 
				set objFile = Nothing 
			call drawScheduleForm	

        end if

End sub