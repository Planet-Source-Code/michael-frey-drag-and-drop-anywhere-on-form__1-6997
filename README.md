<div align="center">

## Drag and Drop anywhere on form


</div>

### Description

Use this code to drag and drop any control anywhere on a form or another control.
 
### More Info
 
Lets supose you want to drop a picture or a button on the form or on a big label. Add a Commandbutton (Command1) a big Label (Label1) and a picture (Picture1).

IMPORTANT

The picture and commandbutton property DRAGMODE have to be set to Automatic


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Frey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-frey.md)
**Level**          |Beginner
**User Rating**    |4.2 (21 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-frey-drag-and-drop-anywhere-on-form__1-6997/archive/master.zip)

### API Declarations

Public X1, Y1


### Source Code

```
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Source.Top = Y - Y1
Source.Left = X - X1
End Sub
Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
Source.Top = Label1.Top + Y - Y1
Source.Left = Label1.Left + X - X1
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
X1 = X
Y1 = Y
End Sub
```

