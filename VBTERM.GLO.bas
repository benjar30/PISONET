Attribute VB_Name = "Module1"
' Public variables
Public Echo As Boolean        ' Echo On/Off flag.
Public CancelSend As Integer  ' Flag to stop sending a text file.

Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)




