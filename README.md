<div align="center">

## \_ Replace text in All text boxes contained in frames, form, and/or pictureboxes


</div>

### Description

I noticed an article recently posted about this. I am posting this because it is a much more efficient way to do it. Simple code so I am not expecting votes, but if you like it please feel free :)
 
### More Info
 
'Example usage:

'TextBoxMod Me, 0, "Kryo"

'Would make ALL textboxe' text say "Kryo"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[KRYO\_11](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kryo-11.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kryo-11-replace-text-in-all-text-boxes-contained-in-frames-form-and-or-pictureboxes__1-50323/archive/master.zip)

### API Declarations

```
'Enum declared to make using following sub easier
Public Enum What2Clear
  [Clear All Textbox's] = 0
  [Clear Textbox's Contained In Frames] = 1
  [Clear Textbox's Contained In Picturebox's] = 2
  [Clear Textbox's Contained In Form] = 3
End Enum
```


### Source Code

```
Public Sub TextBoxMod(WhichForm As Form, CommandLine As What2Clear, Optional ReplaceWith As String = Empty)
  For Each Control In WhichForm 'Search's through given form
    If CommandLine = [Clear All Textbox's] Then
      If TypeOf Control Is TextBox Then Control.Text = ReplaceWith
      'Look for ALL textboxes
    ElseIf CommandLine = [Clear Textbox's Contained In Form] Then
      'Look for textboxes in Form ONLY
      If TypeOf Control Is TextBox And TypeOf Control.Container Is Form Then Control.Text = ReplaceWith
    ElseIf CommandLine = [Clear Textbox's Contained In Frames] Then
      'Look for textboxes in Frmaes ONLY
      If TypeOf Control Is TextBox And TypeOf Control.Container Is Frame Then Control.Text = ReplaceWith
    ElseIf CommandLine = [Clear Textbox's Contained In Picturebox's] Then
      'Look for textboxes in Pictureboxes ONLY
      If TypeOf Control Is TextBox And TypeOf Control.Container Is PictureBox Then Control.Text = ReplaceWith
    End If
  Next
End Sub
```

