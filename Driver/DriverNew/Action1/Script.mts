Services.StartTransaction "trancation1"

'Action1
mrowcount=datatable.GetSheet("Action1").GetRowCount
msgbox mrowcount

For i=1 to mrowcount step 1
datatable.SetCurrentRow(i)
Modexe=datatable("ModuleExe","Action1")
'msgbox Modexe

If Modexe="Y" Then
    Modid=datatable("ModuleID","Action1")
    msgbox Modid

'Action2
trowcount=datatable.GetSheet("Action2").GetRowCount
msgbox trowcount
For j=1 To trowcount Step 1
    Datatable.SetCurrentRow(j)
    If Modid=Datatable("ModuleID","Action2") and Datatable("TestCaseExe","Action2")="Y" then
    testcaseid=Datatable("TestCaseID","Action2")
    msgbox testcaseid
    
    'Action3
     tsrowcount=Datatable.GetSheet("Action3").GetRowCount
     msgbox tsrowcount
     For k = 1 to tsrowcount Step 1
        datatable.SetCurrentRow(k)
        If testcaseid=Datatable("TestCaseID","Action3") Then
        keyword=Datatable("Keyword","Action3")
        msgbox keyword
     
    Select case (keyword)

    Case "ln"
        Call Login("john","hp")

        Case "ca"
        Call Closeapp()

        Case "oo"
        Call OpenOrder(orno)
        Case "uo"
        Call UpdateOrder()
        
        Case "lnd"
        drowcount=datatable.GetSheet("Action4").GetRowCount
        
       For l = 1 To drowcount Step 1
       	
       		datatable.SetCurrentRow(l)
       		
       		Call login(datatable("username","Action4"),datatable("password","Action4"))
       		Call Closeapp()
       		
       Next
       
       Case "ood"
       
       orderrowcount=datatable.GetSheet("Action4").GetRowCount
       For m = 1 To orderrowcount Step 1
      	datatable.SetCurrentRow(m)
      	
      	Call OpenOrder(datatable("Orderno","Action4"))
 @@ hightlight id_;_9372948_;_script infofile_;_ZIP::ssf12.xml_;_
 @@ hightlight id_;_2140804656_;_script infofile_;_ZIP::ssf13.xml_;_
       Next

        End  Select

      End If
      Next

   End If
    Next

End If
Next
Services.EndTransaction "trancation1" @@ hightlight id_;_11797934_;_script infofile_;_ZIP::ssf11.xml_;_
