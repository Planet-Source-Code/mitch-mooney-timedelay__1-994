<div align="center">

## TimeDelay


</div>

### Description

TimeDelay fuction is good for when you want to time out a loop, in milliseconds.

Does'nt use a timer control, Uses simple api declare.
 
### More Info
 
Delay as Long, milliseconds

Create a module with the following api declare and function

Usage can be

Do

FuncThatRetunsTrue

MoreCode

EctEct

Loop until ( FuncThatRetunsTrue=True) or (TimeDelay(60000)=True)

TimeDelay as boolean, turns turn if time reached,else false


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mitch Mooney](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mitch-mooney.md)
**Level**          |Unknown
**User Rating**    |4.2 (159 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mitch-mooney-timedelay__1-994/archive/master.zip)

### API Declarations

Declare Function GetTickCount& Lib "kernel32" ()


### Source Code

```

Public Function TimeDelay(ByVal Delay As Long) As Boolean
Static Start As Long
Dim Elapsed As Long
If Start = 0 Then                            'if start is 0 then set a
  Start = GetTickCount                       'Static value to compare
End If
Elapsed = GetTickCount
If (Elapsed - Start) >= Delay Then
  TimeDelay = True
  Start = 0                            'Remember to reset start
Else: TimeDelay = False                 'once true so subsquent
End If                                'calls wont "spoof" on you!
End Function
```

