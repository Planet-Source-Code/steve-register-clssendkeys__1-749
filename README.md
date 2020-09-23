<div align="center">

## clsSendKeys


</div>

### Description

Allows users to be able to send keystrokes to dos programs running in a windows95 dos box
 
### More Info
 
This class has one property, Destination, which needs to be the handle returned from the shell function of the dos program or any program

started with the shell function.

It also has one method called, SendKeys, this is the string to be sent to the destination.

Nothing except how to use a class module in their code

None that I am aware of


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steve Register](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steve-register.md)
**Level**          |Unknown
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steve-register-clssendkeys__1-749/archive/master.zip)

### API Declarations

none, everything is done in the class module


### Source Code

```
Option Explicit
'local variable(s) to hold property value(s)
Private mvarDestination As Long 'local copy
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SHIFT = &H10
Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Integer) As Long
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Sub SendAKey(ByVal keys As String)
  Dim vk%
  Dim shiftscan%
  Dim scan%
  Dim oemchar$
  Dim dl&
  Dim shiftkey%
  ' Get the virtual key code for this character
  vk% = VkKeyScan(Asc(keys)) And &HFF
  ' See if shift key needs to be pressed
  shiftkey% = VkKeyScan(Asc(keys)) And 256
  oemchar$ = " " ' 2 character buffer
  ' Get the OEM character - preinitialize the buffer
  CharToOem Left$(keys, 1), oemchar$
  ' Get the scan code for this key
  scan% = OemKeyScan(Asc(oemchar$)) And &HFF
  ' Send the key down
  If shiftkey% = 256 Then
  'if shift key needs to be pressed
    shiftscan% = MapVirtualKey(VK_SHIFT, 0)
    'press down the shift key
    keybd_event VK_SHIFT, shiftscan%, 0, 0
  End If
  'press key to be sent
  keybd_event vk%, scan%, 0, 0
  ' Send the key up
  If shiftkey% = 256 Then
  'keyup for shift key
    keybd_event VK_SHIFT, shiftscan%, KEYEVENTF_KEYUP, 0
  End If
  'keyup for key sent
  keybd_event vk%, scan%, KEYEVENTF_KEYUP, 0
End Sub
Public Sub SendKeys(ByVal keys As String)
  Dim x&, t As Integer
  'loop thru string to send one key at a time
  For x& = 1 To Len(keys)
      'activate target application
      AppActivate (mvarDestination)
      'send one key to target
      SendAKey Mid$(keys, x&, 1)
  Next x&
End Sub
Public Property Let Destination(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Destination = 5
  mvarDestination = vData
End Property
Public Property Get Destination() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Destination
  Destination = mvarDestination
End Property
```

