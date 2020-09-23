<div align="center">

## Monitor mouse and keyborad


</div>

### Description

Monitor mouse movement and key presses globel wide.

This code will check for mouse movement or

keyboard presses. Works like a screen saver.

It is globel wide. Not window dependent.

You could use it to monitor input or to detect

whats keys have been pressed.

You could use it as a independent screensaver.

You could use it to shutdown your computer

after certain amount of time has passed without

any key or mouse movement.

I can think of lots of things it could be used

for.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[GaryLake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/garylake.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/garylake-monitor-mouse-and-keyborad__1-7691/archive/master.zip)





### Source Code

```
'##### Setup ##########
'Start a standard project with one form.
'Make the form Height 2200 twips
'and Width 4400 twips.
'Put a Label on the form and
'make it cover the top
'2/3 of the form.
'Put a command button on the
'bottom of the form.
'Add a Timer to the form.
'Paste the code into the code window.
'Have your Immediate window
'showing to see what its doing.
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type
Private Sub Command1_Click()
Timer1.Interval = 100
Timer1.Enabled = True
Me.Visible = False
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
Label1.Caption = "Press the button and this form will disappear. " _
        & "You can work all you want and the form will stay hidden " _
        & "as long as the computer is not sitting idel. " _
        & "After a number of seconds have passed without " _
        & "keyboard or mouse movement it will reappear."
End Sub
Private Sub Timer1_Timer()
Dim MouseMoved As Boolean
Dim KeyPressed As Boolean
Dim KeyCounter As Integer
Dim CurrentCursorPosition As POINTAPI
Static LastCursorPosition As POINTAPI
Static TimePassed As Date
'Loop through every key on keyboard
For KeyCounter = 1 To 256
'Check with API for keypress
  If GetAsyncKeyState(KeyCounter) <> 0 Then
  Debug.Print "Key Pressed"
  Debug.Print Chr$(KeyCounter)
    KeyPressed = True
    Exit For
  End If
Next
'Get the cursor position from API call
GetCursorPos CurrentCursorPosition
'Check the new cursor position with
'the last cursor position
If CurrentCursorPosition.x <> LastCursorPosition.x Or _
  CurrentCursorPosition.y <> LastCursorPosition.y Then
  Debug.Print "Mouse Moved"
  Debug.Print "x= " & CurrentCursorPosition.x
  Debug.Print "y= " & CurrentCursorPosition.y
  MouseMoved = True
End If
'Save the present cursor position to
'check against new position on next pass
  LastCursorPosition = CurrentCursorPosition
  Debug.Print DateDiff("s", TimePassed, Now)
'if movement then reset TimePassed
'back to 0
  If KeyPressed Or MouseMoved = True Then
    TimePassed = Now
  End If
'if no movement then
  If KeyPressed Or MouseMoved = False Then
  'check how much time has passed
  'against the time present time
  'in seconds and if more than 5
  'then make the form visiable
  'and shut the time off.
  'The more than 100000 is
  'required for the first pass.
    If DateDiff("s", TimePassed, Now) > 5 And _
      DateDiff("s", TimePassed, Now) < 100000 Then
      Me.Visible = True
      Timer1.Enabled = False
      Exit Sub
    End If
  End If
 KeyPressed = False
 MouseMoved = False
End Sub
```

