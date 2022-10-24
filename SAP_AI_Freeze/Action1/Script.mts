'Set the Context for SAP GUI Window
AIUtil.SetContext SAPGuiSession("micclass:=SAPGuiSession")

'Enter the TCode
AIUtil("combobox", "E5").SetText "VA01"
Dim mySendKey
Set mySendKey = CreateObject("WScript.shell")
'Send Enter key
mySendKey.SendKeys("~") 

'Start the Sales Order
AIUtil.Context.Freeze
AIUtil("text_box", "Order Type").SetText "OR"
AIUtil("text_box", "Sales Organization").SetText"1710"
AIUtil("text_box", "Distribution Channel.").SetText "10"
AIUtil("text_box", "Division").SetText "00"
AIUtil("button", "Continue").Click
AIUtil.Context.UnFreeze

'Complete the Sales Order
AIUtil.Context.Freeze
AIUtil("text_box", "Sold-To Party:").SetText "EWM17-CU02"
AIUtil("text_box", "Ship-To Party:").SetText "EWM17-CU02"
AIUtil("text_box", "Cust. Reference").SetText "450000019998"
AIUtil("text_box", "Cust. Ref. Date").SetText "10/24/2022"
AIUtil("plus").Click
AIUtil("button", "Save").Click
AIUtil.FindTextBlock("Exit").Click
AIUtil.Context.UnFreeze
