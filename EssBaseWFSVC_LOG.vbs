Sub DoNothing
	 doevents
End Sub

Sub doSleep( vTime  )
  If Not IsNumeric(vTime ) Then _
       Exit Sub
 Dim dteStart,dteEnd
    dteStart = Time()
	dteEnd = DateAdd("s", vTime, dteStart)
 
	While dteEnd > Time()
		DoNothing
	Wend

end Sub




Sub getSleepy
	 doSleep 0.01 
End Sub 

sub myEcho  (LogMessage )
Dim objWScript   
  objWScriptShell.Echo   LogMessage
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
        getCookieFile = objWScriptShell.SpecialFolders("mydocuments") & "\essribon.csc"      
end function 

function getSSOFile     
        getSSOFile = objWScriptShell.SpecialFolders("mydocuments") & "\essribon.sso"      
end function 

function getlogFile      
        getlogFile = objWScriptShell.SpecialFolders ("mydocuments") & "\essribon.log"   
end function 

function getLog
        Dim objFile,vStrFileName
    
        vStrFileName = getlogFile
		 if objFileSystemObject.FileExists(vStrFileName) then 
		   set objFile =  objFileSystemObject.OpenTextFile(vStrFileName, 1)
			vStrFileName = ""   
			Do Until objFile.AtEndOfStream				
					vStrFileName = vStrFileName & objFile.ReadLine 				
			Loop
		   set objFile = Nothing
		 else 
		   vStrFileName = ""      
		 end if 

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
       ' Set objFSO= objFileSystemObject ' CreateObject("Scripting.FileSystemObject")        
		 if objFileSystemObject.FileExists(vCurrFileName) then 
		    set objFile =  objFileSystemObject.OpenTextFile(vCurrFileName, 8, True)', TristateTrue ) 
		 else 
		    set objFile =  objFileSystemObject.CreateTextFile(vCurrFileName,true) 	 
		 end if 

		 LogMessage  =  LogMessage & ";" &  getRightTime () 

          'alert LogMessage   
           objFile.WriteLine LogMessage  
         
            objFile.Close
          
            set objFile = Nothing 
End sub

sub WriteFileRaw ( vStrFileName,vArrRaw)

        Dim objFile
 
		 if objFileSystemObject.FileExists(vStrFileName) then 
		     objFileSystemObject.deletefile vStrFileName
		 end if 
		 
         set objFile =  objFileSystemObject.CreateTextFile(vStrFileName,true) 

        For i = 0 To UBound(vArrRaw)
		 if (len (vArrRaw(i)) > 3 ) then 
           objFile.WriteLine vArrRaw(i)
		  end if 
        Next
 
         objFile.Close      
        set objFile = Nothing 
End sub


sub WriteCookie

      call WriteFileRaw (getCookieFile,vArrSceduleRules)
   
End sub

sub  ReadFileRaw (vStrFileName,vCurrArr,i )      
	Dim objFile,vCurrStr               
	i=0 
		 if objFileSystemObject.FileExists(vStrFileName) then
			set objFile = objFileSystemObject.OpenTextFile(vStrFileName,1)
		 For j = 0 To UBound(vCurrArr)
		    vCurrArr(j)=""				 
		 next 	
		   
			Do Until objFile.AtEndOfStream				
				vCurrStr = objFile.ReadLine 
				if len (vCurrStr) > 3 then 
				  vCurrArr(i) = vCurrStr  				    
				  i=i+1        
				end if 				
			Loop
				objFile.Close          			 
				set objFile = Nothing 	 
        end if  
End sub


sub  CheckFileDateAndDelete (vStrFileName )      
	Dim objFile,vCurrStr               
		 if objFileSystemObject.FileExists(vStrFileName) then
			set objFile = objFileSystemObject.GetFile(vStrFileName) 'objFileSystemObject.OpenTextFile(vStrFileName,1)
	         If DateDiff("n", objFile.DateLastModified, Now) > 60 Then
			    set objFile = Nothing 
                objFileSystemObject.deletefile vStrFileName
             End If 			 
			 set objFile = Nothing 	 
        end if  
End sub


sub ReadCookie
	 call ReadFileRaw (getCookieFile,vArrSceduleRules,vArrSceduleRulesID)    
	 call drawScheduleForm	         

End sub


Public Function f_XORDecryption(DataIn  )  
    Dim lonDataPtr 
    Dim strDataOut  
    Dim intXOrValue1  
    Dim intXOrValue2  
  
    For lonDataPtr = 1 To (Len(DataIn) / 2)        
        intXOrValue1 = CLng("&H" & (Mid(DataIn, (2 * lonDataPtr) - 1, 2)))        
        intXOrValue2 = Asc(Mid(SecretWord, ((lonDataPtr Mod Len(SecretWord)) + 1), 1))        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next  
   f_XORDecryption = strDataOut   
End Function

Dim arrConnectionsTMP (100,11)

sub ReadSSO
   Dim arrCurrConnections (100) 
   Dim jARR , vSplitRow 

     call CheckFileDateAndDelete (getSSOFile)
	 call ReadFileRaw (getSSOFile,arrCurrConnections,jARR) ' Dim arrConnections (100,11)  
     if jARR > 0 then 
	  for i = 0 to ubound (arrConnections)
	       arrConnectionsTMP(i,0) = 0 
		   for j=1 to 10 
		     arrConnectionsTMP(i,j) = "" 
		   next 
	   next  

	   for i = 0 to jARR 
	      arrCurrConnections(i)=f_XORDecryption(arrCurrConnections(i))
	      vSplitRow=split ((arrCurrConnections(i)),"`")		  
		  if ubound(vSplitRow) > 9 then 
		   for j=0 to ubound(vSplitRow)-1 
		     arrConnectionsTMP(i,j) = vSplitRow(j)  
		   next 
		  end if 
	   next         
    end if 
End sub
 
Public Function f_XOREncryption(DataIn  )  
    Dim lonDataPtr  
    Dim strDataOut  
    Dim temp 
    Dim tempstring  
    Dim intXOrValue1  
    Dim intXOrValue2  
   
       For lonDataPtr = 1 To Len(DataIn)
   
        intXOrValue1 = Asc(Mid(DataIn, lonDataPtr, 1))   
         
        intXOrValue2 = Asc(Mid(SecretWord, ((lonDataPtr Mod Len(SecretWord)) + 1), 1))
        
        temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next  
   f_XOREncryption = strDataOut 

End Function
 

sub WriteSSO
   Dim arrCurrConnections (100) 
   Dim jARR , vSplitRow 
        jARR = 0 
   	   for i = 0 to ubound(arrConnections)
	      if arrConnections(i,0) > -1 then 
		   arrCurrConnections(i) = "" 
		   jARR = jARR + 1 		 
		   for j=0 to 11
		     arrCurrConnections(i) = arrCurrConnections(i)  & arrConnections(i,j) & "`"
		   next 
		     arrCurrConnections(i) = f_XOREncryption (arrCurrConnections(i))
		  end if 
	   next  
	 call WriteFileRaw (getSSOFile,arrCurrConnections)          
End sub
