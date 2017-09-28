window.ReSizeTo  1200,900


'HypExecuteCalcScriptEx2 
'HypExecuteCalcScriptString 
'HypGetCalcScript

Dim vSVsID,vSVSSO

Dim vConnAps,vConnEsb 
Dim vConnDb,vConnApp
Dim vAppSID,vCubeSID 
Dim vConnUser,vConnPass 



Sub setCopyRight 
	On Error  Goto 0 
	Dim strHTML                       
	strHTML = OutCopyRight.InnerHTML 
	strHTML = strHTML &  "<a  href=""https://www.linkedin.com/in/essbaseru""> &copy;  </a> "
	
	OutCopyRight.InnerHTML = strHTML 
	
End Sub 

Dim arrVariblesName,arrVariblesValues

 
function setSubsVariable ( vURL,vCurrSID,vCurrAppName,vCurrCubeName ,vCurrVaribleName,vCurrVaribleValue)
	On Error  Goto 0 
	Dim vRequest
	vRequest = "<req_SetSubVar> " _
			& "   <sID>" & vCurrSID & "</sID> " _
			& "   <app>" & vCurrAppName & "</app> " _	
			& "   <cube>" & vCurrCubeName & "</cube> " _
			& "   <VarName>" & vCurrVaribleName & "</VarName>" _    
			& "   <VarVal>" & vCurrVaribleValue & "</VarVal>" _    
		& "</req_SetSubVar>"			  
	Dim objDOMDocument
	Set objDOMDocument = getSVXMLAnswer(vURL,vRequest) 
        setSubsVariable =  getXMLValue(objDOMDocument,"/") 
	Set objDOMDocument = Nothing
	if 0 = len(setSubsVariable)  Then
		pErrorHandler "setSubsVariable error" ,1 
	end if 	
end function

sub setSubsVariables   
	On Error  Goto 0 
    ' Set objDivID =  document.getElementById("btnRun" & vCalcCurrScriptName)
    ' getElementsByTagName("input").item(1).value
    OutChangeVariables.innerHTML="" 
  dim strHTML   
	 For j = 0 To UBound(arrVariblesName)
     if ((instr("1234567890",Left(arrVariblesName (j),1)) = 0 ) and (instr(arrVariblesName (j),"CUBE_NAME") = 0 ) ) then              
           Set objDivID = document.getElementById("txt" & arrVariblesName (j))           
            if (  objDivID.value <>  arrVariblesValues(j) ) then 
               strHTML=strHTML & setSubsVariable (vConnAps,vAppSID,vConnApp,vConnDb,arrVariblesName (j),objDivID.value) 			  
               arrVariblesValues(j)=objDivID.value
            end if         
     end if            
    Next     
 Set objDivID = Nothing     
	call reloadVariables
end sub

function getVariablseForm 
Dim strHTML
 strHTML= "<div > <table> " '<div style=""height:200px;border:1px ;overflow:auto;""> 
     vCubeScriptXML = getCubeVariablesList(vConnAps,vAppSID,vConnApp,vConnDb) 
      if ( len(vCubeScriptXML)>5 ) then 
          arrRules3 = Split (vCubeScriptXML,"#") 
          arrVariblesValues = Split (arrRules3(1),"|")   
          arrVariblesName = Split (arrRules3(0),"|")  
          set arrRules3 = nothing
         

          For j = 0 To UBound(arrVariblesName)                     
             if ((instr("1234567890",Left(arrVariblesName (j),1)) = 0 ) and (instr(arrVariblesName (j),"CUBE_NAME") = 0 ) ) then
                arrVariblesValues (j) = replace(arrVariblesValues (j),"""","")             
                strHTML = strHTML & "<tr> <td align=""left"" width=""50"" class=""ui label""> &."  
                strHTML = strHTML &   arrVariblesName (j)   
                strHTML = strHTML & "</td> <td  width=""50"" >"
                strHTML = strHTML & " <input id=""txt" & arrVariblesName (j) & """ type=""text""  value="""& arrVariblesValues (j)&  """> "             
                strHTML = strHTML & "</td>  </tr> "                          
            end if             
          Next  
      end if 
    getVariablseForm = strHTML & "</table> </div> "  
end function

sub reloadVariables   
	On Error  Goto 0   
     OUTVaribleFORM.innerHTML=getVariablseForm 
     fVariables.btnSetVariables.disabled = false 
     fVariables.btnHideVariables.disabled = false      
end sub

sub hideVariables   
	On Error  Goto 0   
     OUTVaribleFORM.innerHTML="" 
     fVariables.btnSetVariables.disabled = true 
     fVariables.btnHideVariables.disabled = true 
end sub

sub setCSCScripsForm
    OUTScriptsFORM.innerHTML=getScriptsForm
    runbutton.disabled = false
   ' runExpButton.disabled = false
    loadCsCbutton.disabled = true
    hideCsCbutton.disabled = false     
end sub 

sub runExpClose                  
          outTabRuleBody.InnerHTML =  "" 'jsScriptBody
END sub


sub hideCSCScrips
    OUTScriptsFORM.innerHTML=""
    outTabRuleBody.InnerHTML =  ""
    runbutton.disabled = true
  '  runExpButton.disabled = true  
    loadCsCbutton.disabled = false 
    hideCsCbutton.disabled = true     
end sub 

function getFirstSegmnet 
	Dim strHTML
 	Dim arrEssbSrv,vCurrEsbServer 
   
    vCurrEsbServer=""
	strHTML=  "<div class=""ui segment"">"
	 strHTML= strHTML & "<b class=""ui basic""> Database list: </b> <br> <br> " 
    For i = 0 to (Ubound(arrConnections) ) 

	if (len(arrConnections (i,1)) > 5 ) then 
		if ( len (vCurrEsbServer) = 0) then 
			vCurrEsbServer = arrConnections (i,4) 
			arrEssbSrv=split (vCurrEsbServer,".")
			strHTML= strHTML & "<b class=""ui basic"">@"& arrEssbSrv(0)&"</b> <br>" 
			' arrEssbSrv = Nothing
		end if 

		if ( instr(arrConnections (i,4),vCurrEsbServer) = 0 ) then 
			vCurrEsbServer = arrConnections (i,4) 
			arrEssbSrv=split (vCurrEsbServer,".")
			strHTML= strHTML & "<b class=""ui basic"">@"& arrEssbSrv(0)&"</b> <br>" 
			' arrEssbSrv = Nothing
		end if 
		
		if ( instr(arrConnections (i,5),vConnApp) = 0 and instr(arrConnections (i,5),vConnDb) = 0  ) then 
			strHTML= strHTML & "<a class=""ui basic"" id=""" & i & """ onclick=""vbscript:Call changeCurrentCube(window.event.srcelement.id,1)"" href=""#"" >" &_ 
						"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"& lcase( arrConnections (i,5) & "." & arrConnections (i,6)) &"</a> <br> " 
		    strHTML = strHTML & "<div style=""height: 5px;""> </div>"
		end if 

	 end if  	  
	next  
	 strHTML= strHTML & "</div>"	
	getFirstSegmnet =  strHTML
end function


 
function getScriptsForm 
Dim strHTML 
Dim arrRules1,arrRules2 
   strHTML=""
    strHTML = getCubeScripts(vConnAps,vAppSID,vConnApp,vConnDb ) 'vConnApp	'
	'alert strHTML
	arrRules1 = Split (strHTML,"rtp=""0"">") 
	if Ubound (arrRules1) < 1 then 
	 alert strHTML
	end if 
    strHTML = "<div  width=""width:600px;border:1px ;overflow:auto;"" ><table> <tr  > <td>" 
	For J = 1 To UBound(arrRules1)
		arrRules2 = Split(arrRules1(j),"</rule>")
		if j = 12 or j = 24 then 
		 strHTML = strHTML & "</td><td>&nbsp;&nbsp;  </td><td>"
		end if 
		if (instr(ucase(arrRules2(0)),"Z") <> 1 ) and   (instr(ucase(arrRules2(0)),"T") <> 1 )then 
			strHTML = strHTML & "<div class=""field"">"
			strHTML = strHTML & "<div class=""ui radio checkbox"">"
			strHTML = strHTML & "<input type=""radio"" name=""CalcOption"" value=""" & vConnApp & "." &  vConnDb & "." &  arrRules2(0) & """> " & _
			 " >  &nbsp  &nbsp <a  id=""" & vConnApp &"."& vConnDb &"."& arrRules2(0)  & """  onclick=""vbscript:Call setScheduleForm(window.event.srcelement.id)"" href=""#"" > " & arrRules2(0) & "</a> <div id=""btnRun" & arrRules2(0)  & """ ></div>  "
			strHTML = strHTML & "</div>"      
			strHTML = strHTML & "</div>"                     
		end if                    
	Next 
   getScriptsForm= "</td> </tr> </table> " & strHTML  & " </div>"
    'alert strHTML
end function


Dim vArrSceduleRules (100) 
Dim vArrSceduleRulesID

sub setScheduleForm  (vCurrScriptName)
Dim strHTML ,i 
strHTML = 0 
 For i = 0 To UBound(vArrSceduleRules)
	if (instr(vArrSceduleRules(i),vCurrScriptName) > 0 ) then 
	 strHTML = 1
	end if 
 next 

  if (strHTML <1 ) then 
'   alert vArrSceduleRules (0) 
  vArrSceduleRulesID =  vArrSceduleRulesID  + 1   

   i = vArrSceduleRulesID
   vArrSceduleRules (i) = vCurrScriptName

   strHTML=OUTScheduleForm.innerHTML
      strHTML = strHTML & "<div class=""field"">"
			strHTML = strHTML & "<div class=""ui radio checkbox"">"
			strHTML = strHTML & "<input type=""radio"" name=""ScheduleOption"" value=""" & vCurrScriptName & """> " & _
			 " >  &nbsp  &nbsp <a  id=""" & vCurrScriptName  & "#" & i &  """  onclick=""vbscript:Call delScriptFromForm(window.event.srcelement.id)"" href=""#"" > " & vCurrScriptName& "</a>  "
			strHTML = strHTML & "</div>"      
			strHTML = strHTML & "</div>"                     		
   OUTScheduleForm.innerHTML = strHTML 
   else 
    alert vCurrScriptName & " is already added"
  end if   
end sub

Dim vIsDeletedRule

sub delScriptFromForm (vCurrScriptID)
  Dim Arr
   arr=split(vCurrScriptID,"#")
      vArrSceduleRules (arr(1)) = ""
	  vIsDeletedRule = true
   drawScheduleForm 
end sub

sub drawScheduleForm 
dim vArrClearRules(100)
if vIsDeletedRule then 
j=0 
for i = 0 To UBound(vArrSceduleRules)
	  if ( len(vArrSceduleRules(i)) > 3 ) then 
	    vArrClearRules(j)=vArrSceduleRules(i)
	    j=j+1
      end if  
	Next 
	for i = 0 To UBound(vArrSceduleRules)
	    vArrSceduleRules(i)=""
	    vArrSceduleRules(i)=vArrClearRules(i)	      
	Next 
end if 
  Dim strHTML ,i  
    strHTML = ""
	For i = 0 To UBound(vArrSceduleRules)
	  if ( len(vArrSceduleRules(i)) > 3 ) then 
	         strHTML = strHTML & "<div class=""field"">"
			strHTML = strHTML & "<div class=""ui radio checkbox"">"
			strHTML = strHTML & "<input type=""radio"" name=""ScheduleOption"" value=""" & vArrSceduleRules(i) & """> " & _
			 " >  &nbsp  &nbsp <a  id=""" & vArrSceduleRules(i)  & "#" & i &  """  onclick=""vbscript:Call delScriptFromForm(window.event.srcelement.id)"" href=""#"" > " & vArrSceduleRules(i) & "</a> <div id=""btnRun" & vArrSceduleRules(i) & """ ></div> "
			strHTML = strHTML & "</div>"      
			strHTML = strHTML & "</div>" 
      end if  
	Next
	 OUTScheduleForm.innerHTML = strHTML 
end sub

Dim objSelectedCSC
sub upCSCScrip
   On Error  Goto 0 
	Dim objButton
	Dim Arr,vSTr,vIsmoved
	'alert ScheduleOption
	vIsmoved=0 
	For Each objButton in fScheduleFormName.ScheduleOption
		If ( objButton.Checked  ) Then
		     For i = 1 To UBound(vArrSceduleRules)
			  if vIsmoved=0 then 
				if (instr(vArrSceduleRules(i),objButton.value) > 0 and len(vArrSceduleRules(i))>1 ) then 
				  vSTr=vArrSceduleRules (i-1)  
			      vArrSceduleRules (i-1)=vArrSceduleRules (i)
			      vArrSceduleRules (i) = vSTr
				  vIsmoved=1
				end if 
			  end if 	
			next 
		End If
	Next  
     drawScheduleForm  
END sub 

sub downCSCScrip
   On Error  Goto 0 
	Dim objButton
	Dim Arr,vSTr,vIsmoved
	'alert ScheduleOption
	vIsmoved=0 
	For Each objButton in fScheduleFormName.ScheduleOption
		If ( objButton.Checked  ) Then
		     For i = 0 To UBound(vArrSceduleRules)-1
			 if vIsmoved=0 then 
				if (instr(vArrSceduleRules(i),objButton.value) > 0  and len(vArrSceduleRules(i))>1 ) then 
				  vSTr=vArrSceduleRules (i+1)  
			      vArrSceduleRules (i+1)=vArrSceduleRules (i)
			      vArrSceduleRules (i) = vSTr
				  vIsmoved=1
				end if 
			 end if 	
			next 
		End If
	Next  
     drawScheduleForm   
END sub 

sub setButtonOnRun 
	runbutton.disabled = true
	hideCsCbutton.disabled = true 
	loadCsCbutton.disabled = true 
	fScheduleFormName.cscRUN.disabled = true  
end sub 

sub runCurrentRuleByID (i)
dim strAttr 
dim vIsFound
 if ( len(vArrSceduleRules(i))>1 ) then 
 strAttr = split ( vArrSceduleRules(i)  ,".")
 vIsFound = 0 
 	For j = 0 To UBound(arrConnections )
	  if ( vIsFound = 0 ) then 
	  'alert (arrConnections(j,0))&arrConnections(j,5)&arrConnections(j,6)
		If (instr (strAttr(0),arrConnections(j,5))  > 0 and  instr (strAttr(1),arrConnections(j,6)) > 0 ) then 
			 call getLoginSID (j) 
			 vIsFound = 1 
		end if 	 
	   end if 	
	next 
	 
	 if (vIsFound = 1 ) then 
	  
	        call runExpClose          
            call setButtonOnRun
			'alert vArrSceduleRules(i)
			Set objDivID =  document.getElementById("btnRun" & vArrSceduleRules(i)  )                    
			objDivID.innerHTML = objDivID.innerHTML  & "  <div class=""description"">  &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp  Started : &nbsp &nbsp"  & Time &  "</div>  "  ' <div class=""item"">  <div class=""content"">
			runStatus.InnerHTML = "<div class=""description"">  " & vArrSceduleRules(i) &" have started at " & Time &  "</div>  "  ' <div class=""item"">  <div class=""content"">    
			WriteLog vArrSceduleRules(i) & ";" & " start"
			getSleepy  

		if ( instr(ucase( strAttr(2)),"DEFAULT") = 0 ) then 
				strHTML = "<req_LaunchBusinessRule> " _
						& "   <sID>" & vSVsID & "</sID> " _	                         
						& "   <cube>" & vConnDb & "</cube> " _					
						& "   <rule>" & strAttr(2) & "</rule> " _					
				& "</req_LaunchBusinessRule>"			  
				
				Call getSVXMLAnswerAsyncCalc(vConnAps,strHTML,strAttr(2),vArrSceduleRules(i))                  
			end if 
	 else 
	  alert " Can not find connection for " & vArrSceduleRules(i)
	 end if 		  	 
				   
  end if
end sub



sub runCalcLaunch
   On Error  Goto 0 
	Dim objButton
	Dim strHTML,objDivID 
	Dim strAttr ,vIsFound 
	dim vIsRunning 
	
	vIsRunning = 0 
	For Each objButton in CalcOption
		If ( objButton.Checked and instr(ucase( objButton.Value),"DEFAULT") = 0 ) Then
          'alert objButton.Value
		 strAttr = split ( objButton.Value  ,".")
            vIsRunning = 1 
			call getLoginSID (vGlobalCubeID) 			 

            call runExpClose          
            call setButtonOnRun
			Set objDivID =  document.getElementById("btnRun" & strAttr(2)  )                    
			objDivID.innerHTML = objDivID.innerHTML  & "  <div class=""description"">  &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp  Started : &nbsp &nbsp"  & Time &  "</div>  "  ' <div class=""item"">  <div class=""content"">
			runStatus.InnerHTML = "<div class=""description"">  " & strAttr(2)&" have started at " & Time &  "</div>  "  ' <div class=""item"">  <div class=""content"">    
			WriteLog  vConnApp   & "." & vConnDb  & "." & strAttr(2)  & ";" & " start"
			getSleepy                     
		End If
	Next  
	
	if (vIsRunning = 1) then 
	For Each objButton in CalcOption
		If objButton.Checked Then           
			if ( instr(ucase(strAttr(2)),"DEFAULT") = 0 ) then 
				strHTML = "<req_LaunchBusinessRule> " _
						& "   <sID>" & vSVsID & "</sID> " _	                         
						& "   <cube>" & vConnDb & "</cube> " _					
						& "   <rule>" & strAttr(2) & "</rule> " _					
				& "</req_LaunchBusinessRule>"			  
			'	alert vConnAps 
			'	alert strHTML 
				Call getSVXMLAnswerAsyncCalc(vConnAps,strHTML,strAttr(2),"")                  
			end if                  
		End If
	Next  
	else 
	  alert "Please select the rule before run "
	end if      
END sub  

function getSecondtSegmnet 
Dim strHTML
    strHTML = "<table > <tr> <td width=""652"" >"
        strHTML = strHTML &  " <div class=""ui list""> <div class=""item""> <i class=""folder icon""></i> <div class=""content""> <div class=""header"">" &  lcase(vConnApp & "." & vConnDb) & "</div> "    
            strHTML = strHTML &  " <div class=""list""> <div class=""item""> <i class=""folder icon""></i>  <div class=""content""> <div class=""header""> Variables</div> "                                       
                 strHTML = strHTML & "<form class=""span6"" id=fVariables > "    
                    strHTML = strHTML & "<div id=""OUTVaribleFORM""  style=""height:150px;border:1px ;overflow:auto;""  >" &  "</div>"   'getVariablseForm & 
					strHTML = strHTML & "<br>"  
                    strHTML = strHTML & "<input  style=""width: 75px;"" class=""mini ui button"" id=btnSetVariables    type=""button"" value=""set"" onClick=""vbscript:setSubsVariables"">"
                    strHTML = strHTML & "<input  style=""width: 75px;"" class=""mini ui button"" id=btnReloadVariables type=""button"" value=""load"" onClick=""vbscript:reloadVariables"">"
                    strHTML = strHTML & "<input  style=""width: 75px;"" class=""mini ui button"" id=btnHideVariables   type=""button"" value=""hide"" onClick=""vbscript:hideVariables"">"
                 strHTML = strHTML & "  </form> "
                 strHTML = strHTML & "<div id=""OutChangeVariables"" align=""left""> </div></div></div></div> "
            
            strHTML = strHTML & " <div class=""list""> <div class=""item""> <i class=""folder icon""></i>  <div class=""content""> <div class=""header""> Calculations</div> "                                 
                strHTML = strHTML & "<div class=""ui form"">"
                    strHTML = strHTML & "<div class=""grouped fields"">"'                                      
                        strHTML = strHTML & "<div id=""OUTScriptsFORM""  style=""height:300px;width:500px;border:1px ;overflow:auto;""  align=""left""></div>"
 						strHTML = strHTML & "<br>"
						strHTML = strHTML & "<div style=""height: 5px;""> </div>"
                        strHTML = strHTML & "<input style=""width: 75px;"" id=loadCsCbutton class=""mini ui button"" type=""button"" value=""load"" name=""set_calc""  onClick=""vbscript:setCSCScripsForm""> "                         
                        strHTML = strHTML & "<input style=""width: 75px;"" id=hideCsCbutton class=""mini ui button"" type=""button"" value=""hide"" name=""hide_calc""  onClick=""vbscript:hideCSCScrips""> " 
						'strHTML = strHTML & "<input style=""width: 75px;"" id=runExpButton class=""mini ui button""  type=""button"" value=""view"" name=""exp_calc""  onClick=""vbscript:runExportScript""> "
						'strHTML = strHTML & "<br>"
						'strHTML = strHTML & "<div style=""height: 15px;""> </div>"
						strHTML = strHTML & "<input style=""width: 75px;"" id=runbutton class=""ui button"" type=""button"" value="" run"" name=""run_Calc""  onClick=""vbscript:runCalcLaunch""> " 
					    strHTML = strHTML & "<input style=""width: 75px;"" id=showLog class=""ui button""   type=""button"" value="" log"" name=""show_Log""  onClick=""vbscript:showLog""> " 

						
                    strHTML =  strHTML & " </div> "      
                strHTML =  strHTML & " </div> "     	
            strHTML =  strHTML & " </div></div></div></div> "      
        strHTML =  strHTML & " </div></div></div></div> "                                 
	
	strHTML = strHTML & "</td > </tr> </table >"
	getSecondtSegmnet =  strHTML
end function

function getThirdSegmnet 
Dim strHTML
    strHTML = ""
            strHTML = strHTML & "<form class=""span6"" id=""fScheduleFormID"" name=""fScheduleFormName"" > "   
			strHTML = strHTML & "<div id=""OUTScheduleForm"" style=""height:485px;border:1px ;overflow:auto;"" align=""left""></div>"
			strHTML = strHTML & "<div style=""height: 35px;""> </div>"
			strHTML = strHTML & "<input style=""width: 75px;"" id=cscUP  class=""mini ui button"" type=""button"" value="" up"" name=""up_calc""  onClick=""vbscript:upCSCScrip""> " 
            strHTML = strHTML & "<input style=""width: 75px;"" id=cscDown class=""mini ui button"" type=""button"" value=""down"" name=""down_calc""  onClick=""vbscript:downCSCScrip""> "
			strHTML = strHTML & "<br>"
			strHTML = strHTML & "<div style=""height: 15px;""> </div>"
			strHTML = strHTML & "<input style=""width: 75px;"" id=cscRUN class=""ui button"" type=""button"" value=""run"" name=""run_seq""  onClick=""vbscript:runCSCScrip""> "		
			strHTML = strHTML & "<input style=""width: 75px;"" id=saveCsCbutton class=""ui button"" type=""button"" value=""save"" name=""save_calc""  onClick=""vbscript:saveCSCScrips""> " 
			strHTML = strHTML & "<input style=""width: 75px;"" id=clearCsCbutton class=""ui button"" type=""button"" value=""clear"" name=""clear_calc""  onClick=""vbscript:clearCSCScrips""> " 
			strHTML = strHTML & "<br>"
			strHTML = strHTML & "</form>"	
			
	getThirdSegmnet =  strHTML
end function

Dim vGlobalRunnedCSCID 
sub runCSCScrip
   call WriteCookie
   vGlobalRunnedCSCID=0
   call runCurrentRuleByID (0)
end sub 


sub saveCSCScrips
   call WriteCookie
end sub 
  
sub clearCSCScrips
	For i = 0 To UBound(vArrSceduleRules)
		vArrSceduleRules(i) = "" 
	Next  
   call WriteCookie
   call ReadCookie
end sub 

Sub getForm  
	Dim strHTML
 
	strHTML = "<table> <tr> <td width=""150"" >"
		strHTML = strHTML & " <div class=""ui segment"" style=""height:760px;border:1px ;overflow:auto;"" >"   
		' first segment 
			strHTML = strHTML & getFirstSegmnet
		strHTML = strHTML & "</div> </td> <td width=""1"" bgcolor=""#CFEDFF""> </td> <td width=""652""> "
		strHTML = strHTML & " <div  class=""ui segment"" style=""height:760px;width=652px;border:1px ;overflow:auto;"" >" 
			' second segment 
			strHTML = strHTML & getSecondtSegmnet

		strHTML = strHTML & "</div> </td>  <td width=""1"" bgcolor=""#CFEDFF""> </td> <td  width=""396"" align=""top"" > "
		strHTML = strHTML & " <div class=""ui segment""  style=""height:760px;border:1px ;overflow:auto;"" >"  		
		' third segment 
			strHTML = strHTML & "<table valign=""top""> <tr valign=""top"" > <td> <div style=""height: 50;""> <b> Calcualtion sequence </b> </div>  </td> </tr> <tr> <td> "
			strHTML = strHTML & getThirdSegmnet
			strHTML = strHTML & " </td> </tr> </table>"
		strHTML = strHTML & "</div> </td> </tr> <tr> <td bgcolor=""#CFEDFF""> </td><td bgcolor=""#CFEDFF""> </td><td bgcolor=""#CFEDFF""> </td><td bgcolor=""#CFEDFF""> </td> <td bgcolor=""#CFEDFF""> </td>  </tr>"
		strHTML = strHTML & "</table>"
		strHTML = strHTML & " <div id=""runStatus""></div>  " 

	outFormCalcLaunch.InnerHTML = strHTML 

	OUTScriptsFORM.innerHTML= ""
    outTabRuleBody.InnerHTML= ""
	OUTVaribleFORM.innerHTML= "" 

    fVariables.btnSetVariables.disabled = true 
    fVariables.btnHideVariables.disabled = true 
    
    runbutton.disabled = true
  '  runExpButton.disabled = true
    loadCsCbutton.disabled = false
    hideCsCbutton.disabled = true  
	
End Sub     

function  getScriptList ()  
	
	Dim strHTML,vCubeScriptXML
	Dim xmlNodes,xmlObject 
	Dim arrRules1,arrRules2,arrRules3,i 
	
	vCubeScriptXML = getCubeScripts(vConnAps,vAppSID,vConnApp,vConnDb )
	
	arrRules1 = Split (vCubeScriptXML,"rtp=""0"">")   
	i=0         
	For J = 1 To UBound(arrRules1)
		arrRules2 = Split(arrRules1(j),"</rule>")
		if (instr(ucase(arrRules2(0)),"Z") <> 1 ) and   (instr(ucase(arrRules2(0)),"T") <> 1 )then 
			arrRules3  = arrRules3 &";"&  arrRules2(0)   
			i=i+1                    
		end if                    
	Next   
	getScriptList = arrRules3        
End function 

sub runExportScript
  dim strHTML 
	For Each objButton in CalcOption
		If ( objButton.Checked and instr(ucase( objButton.Value),"DEFAULT") = 0 ) Then            
		'	runExpButton.disabled = true 
 
			strHTML =  getScript(vConnAps,vSVsID,objButton.Value) 
            
            strHTML = "<div class=""ui form""> " & _
                         " <div class=""field""> " & _                            
                             "<input class=""ui button"" type=""button"" value=""" & objButton.Value & """ name=""close""  onClick=""runExpClose""> "& _                             
                             "<textarea rows=""30"" cols=""60"" > " & strHTML & " </textarea> " & _
                         "</div>" & _
                     "</div>"
                     
          outTabRuleBody.InnerHTML =  strHTML 'jsScriptBody 
		End If
	Next 
'	runExpButton.disabled = false 
END sub

sub showLog
  dim strHTML  
		strHTML =  getLog
		
		strHTML = "<div class=""ui form""> " & _
						" <div class=""field""> " & _                            
							"<input class=""ui button"" type=""button"" value=""Close"" name=""close""  onClick=""runExpClose""> "& _                             
							"<textarea rows=""30"" cols=""60"" > " & strHTML & " </textarea> " & _
						"</div>" & _
					"</div>"
					
        outTabRuleBody.InnerHTML =  strHTML 'jsScriptBody 	
END sub



Dim arrConnections (100,11)
Sub getLoginArray 
' maget login parameters from commnad line 
	Dim strHTML1,strHTML2
	Dim arrLine,arrCurrConnect 

   For i = 0 to (Ubound(arrConnections) ) 	 
	  arrConnections (i,0) = -100        
   next 
	 ' alert  objEssCSCLNCH.commandLine
	arrCommands = Split(objEssCSCLNCH.commandLine, "conn=") 
	strHTML=replace(arrCommands(1),"""","")
	arrCommands=Split(strHTML, " csc=")	
	strHTML1=replace(arrCommands(0),"""","")
	strHTML2=""
	if ubound(arrCommands)>0 then 
	  strHTML2=replace(arrCommands(1),"""","")
	end if   

    arrCommands=Split(strHTML1, "|")
     'rulyubimir'rulyubimir'http://wedcb786.frmon.danet:13080/aps/SmartView'wedcb785.frmon.danet:1424'Kz1TTCST'Kz1TTCST
		For i = 0 to (Ubound(arrCommands) ) 
		arrCurrConnect = split (arrCommands(i),"'") 
		if ( ubound (arrCurrConnect) >4 ) then  
			arrConnections (i,0) = 0 
			'alert arrCurrConnect(0)
			arrConnections (i,1) = arrCurrConnect(0) 'login
			arrConnections (i,2) = arrCurrConnect(1) 'pass
			arrConnections (i,3) = arrCurrConnect(2) 'aps
			arrConnections (i,4) = arrCurrConnect(3) 'esb
			arrConnections (i,5) = UCASE(arrCurrConnect(4)) 'app
			arrConnections (i,6) = UCASE(arrCurrConnect(5)) 'db	  
		end if 
		next 

	    arrCommands=Split(strHTML2, "|")
     'BY3TTCST.BY3TTCST.FRC_CALC_TTCST|BY3TTCST.BY3TTCST.ACT_CALC_TTCST

        if ( ubound (arrCommands) >0 ) then  
			For i = 0 To UBound(vArrSceduleRules)
				vArrSceduleRules(i) = "" 
			Next  

			For i = 0 to (Ubound(arrCommands) ) 		
				arrConnections (i,0) = 0 
				'alert arrCurrConnect(0)
				vArrSceduleRules(i)  = arrCommands(i)   
			next 
			WriteCookie
        end if 
	' alert arrConnections (0,1)
end Sub  

sub getLoginSID (i)

' initialisate corrent connection 
 if (0=arrConnections (i,0)) then 
    
	arrConnections (i,7) = getSID(arrConnections (i,3),arrConnections (i,1),arrConnections (i,2))
	
	arrConnections (i,8) = getSSO(arrConnections (i,3),arrConnections (i,7))

	arrConnections (i,9) = getOpennedApplication(arrConnections (i,3),arrConnections (i,7),arrConnections (i,8) ,arrConnections (i,4),arrConnections (i,5))
	
	arrConnections (i,10) = getOpennedCube(arrConnections (i,3),arrConnections (i,7),arrConnections (i,8) ,arrConnections (i,4),arrConnections (i,5),arrConnections (i,6)) 
	
	arrConnections (i,0) = 1
  
 end if 
	
		vConnUser = arrConnections (i,1)
		vConnPass = arrConnections (i,2)
		vConnAps =  arrConnections (i,3)
		vConnEsb =  arrConnections (i,4)
		vConnApp  = arrConnections (i,5)
		vConnDb =   arrConnections (i,6)
		vSVsID =    arrConnections (i,7) 	
		vSVSSO =    arrConnections (i,8)	 
		vAppSID =   arrConnections (i,9)		 
		vCubeSID =  arrConnections (i,10) 
end Sub

Dim vGlobalCubeID
 
sub changeCurrentCube(vNewCubID,vIsSecondRun)
Dim isCanRun
vGlobalCubeID = vNewCubID 

 isCanRun=1 
   if (1 = vIsSecondRun ) then 
		if (runbutton.disabled  and loadCsCbutton.disabled   ) then 
			alert "You can not change the cube during calculation process"
			isCanRun=0		
		end if	
	end if	

	if (isCanRun = 1) then 
		call getLoginSID (vNewCubID)
		call getForm  
		call setCSCScripsForm		
	end if 

    if (vIsSecondRun=0 ) then
     call ReadCookie
	else  
	 call WriteCookie
	 call ReadCookie
	end if  

end sub

Sub pErrorHandler (vErrStr, vTerminate  ) 
	dim vErrorHandler 
	vErrorHandler = OutErrorHandler.InnerHTML  
	vErrorHandler= vErrorHandler & "<BR> -------------"  
	vErrorHandler= vErrorHandler & "<BR > " & vErrStr
	vErrorHandler= vErrorHandler & "<BR > " & Err.Number
	vErrorHandler= vErrorHandler & "<BR > " & Hex(Err.Number)
	vErrorHandler= vErrorHandler & "<BR > " & Err.source  
	vErrorHandler= vErrorHandler & "<BR > " & Err.Description
	vErrorHandler= vErrorHandler & "<BR > " & "-------------"
	
	
	OutErrorHandler.innerHTML =   vErrorHandler 
	Err.Clear  
 
End sub

Sub pCheckError  (vErrStr ) 
	if 0 <> Err.Number then 
		pErrorHandler pErrorHandler,1 
	end if 	
End sub


Sub window_onLoad 
    vIsDeletedRule = false
	getLoginArray	
	changeCurrentCube 0,0 
	setCopyRight
	ReadCookie
End Sub


