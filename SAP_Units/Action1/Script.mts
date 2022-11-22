'Set the Context for SAP GUI Window
AIUtil.SetContext Window("regexpwndtitle:=SAP Easy Access", "regexpwndclass:=SAP_FRONTEND_SESSION", "is owned window:=False", "is child window:=False")

'Enter the TCode
AIUtil("combobox", "E5").SetText "VA01"
'AIUtil("combobox", "v A").Type "VA01"
Dim mySendKey
Set mySendKey = CreateObject("WScript.shell")
'Send Enter key
mySendKey.SendKeys("~") 

'Start the Sales Order
 If AIUtil("text_box", "Order Type").Exist Then
 	AIUtil("text_box", "Order Type").SetText "OR"
 Else wait 2
 End If
AIUtil("text_box", "Order Type").SetText "OR"
AIUtil("text_box", "Sales Organization").SetText "1710"
AIUtil("text_box", micAnyText, micWithAnchorOnLeft, AIUtil.FindText("Distribution Channel")).SetText "10"
AIUtil("text_box", "Division").SetText "00"
AIUtil("button", "Continue").Click

'Complete the Sales Order
'orN = DataTable.Value("OrderNumber", "Global")
AIUtil("text_box", "Standard Order").SetText Parameter("OrderInput")
AIUtil("text_box", "", micFromTop, 2).SetText "EWM17-CU02"
AIUtil("text_box", "Ship-To Party:").SetText "EWM17-CU02"
AIUtil("text_box", "Cust. Reference").SetText "450000019998"
AIUtil("text_box", "Cust. Ref. Date").SetText "11/30/2022"
AIUtil("plus").Click
AIUtil("button", "Save").Click
AIUtil.FindTextBlock("Exit").Click




