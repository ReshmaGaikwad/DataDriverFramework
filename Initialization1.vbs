Dim objuft
Set objuft=CreateObject("QuickTest.Application")
objuft.visible=True
objuft.Launch
objuft.open("C:\Users\sfjbs\Desktop\UFT\DataDriverFramework\Driver\Driver")
objuft.Test.Run
objuft.Test.Close
objuft.quit
Set objuft=nothing
