Attribute VB_Name = "modResize"
'*********************************************************************
'Originally Made by Shaun Gurley in May 2002.  My First submission   *
'Planet source code!!  Resize BAS file. I made this originally as an *
'.ocx, but to keep the size of the file low I decided instead to     *
'keep it in a module, This module resizes pretty much any control    *
'out there.  Fonts are also resized.  Feel free to alter this code   *
'or use it as you see fit but please give me recognition if you do so*
'To use: add ModResize to a project and put: modResize.ReSize(form)in*
'the resize event of form to be used. Or  modResize.resize me        *
' *Fixed No more commenting out code for RichTextBox and Now works   *
'with multiple forms!!                                               *
'*********************************************************************

Option Explicit

Private Type ControlSize 'to store control dimensions
    RatioHeight As Single
    RatioWidth As Single
    RatioTop As Single
    RatioLeft As Single
End Type

Private Type SizeOfFont 'store font sizes
    TheFont As Single
End Type

Private Type LineSize 'store, you guessed it, line sizes
    LLeft As Single
    LWidth As Single
    LTop As Single
    LHeight As Single
End Type

Dim OriginalHeight As Single
Dim OriginalWidth As Single

Dim NewWidth As Single
Dim NewHeight As Single

Dim HeightChange As Single
Dim WidthChange As Single

Dim LineCount As Integer
Dim aForm As Form
Dim LastFormUsed As String

'dim arrays for controls and control dimensions as defined types
Private LineSizes() As LineSize
Private ctlSizeArray() As ControlSize
Private FontArray() As SizeOfFont

'store the initial dimensions of the controls in an array
Private Sub SizeData(aForm As Form)
Dim ctl As Control
Dim i As Integer
Dim j As Integer
 
On Error Resume Next
'lines are counted seperately
Linecounter aForm

'dimension array to hold controls
ReDim FontArray(0 To aForm.Controls.Count - 1)
ReDim ctlSizeArray(0 To aForm.Controls.Count - 1)
ReDim LineSizes(0 To LineCount - 1)

    For Each ctl In aForm.Controls
    
    'deal with lines seperately
        If TypeOf ctl Is Line Then
    
          'save original form size
            OriginalWidth = aForm.ScaleWidth
            OriginalHeight = aForm.Height
            
            'lines are arrayed thus easy to deal with
            With LineSizes(j)
                    .LLeft = ctl.X1  'save line positions
                    .LWidth = ctl.X2
                    .LTop = ctl.Y1
                    .LHeight = ctl.Y2
            End With
            
            j = j + 1
                
            Else
            
            'store the ratio of the size of the control to the form size
            'This is pretty much the key to properly resizing controls
            With ctlSizeArray(i)
                .RatioHeight = ctl.Height / aForm.ScaleHeight
                .RatioWidth = ctl.Width / aForm.ScaleWidth
                .RatioTop = ctl.Top / aForm.ScaleHeight
                .RatioLeft = ctl.Left / aForm.ScaleWidth
            End With
            
        End If
      
      'fontsize property is a bit different on Richtxt box so deal with it separately
      'If form has no RichTextBox this code has to be commented out
            If TypeName(ctl) = "RichTextBox" Then
                FontArray(i).TheFont = aForm.Controls(i).Font.Size / aForm.ScaleWidth
            Else
                FontArray(i).TheFont = aForm.Controls(i).FontSize / aForm.ScaleWidth
            End If
      
    i = i + 1
    
    Next
    
End Sub

Public Sub Resize(aForm As Form)
Dim ctl As Control
Dim i As Integer
Dim j As Integer
Static CountSwitch As Integer
On Error Resume Next

    'this if block resets the arrays if a new form is used
    If aForm.Name <> LastFormUsed Then
        ReDim LineSizes(0)
        ReDim ctlSizeArray(0)
        ReDim FontArray(0)
        LastFormUsed = aForm.Name
        CountSwitch = 0
    End If

If CountSwitch = 0 Then
    SizeData aForm
    CountSwitch = 1
End If

For Each ctl In aForm
'Thecontrols are resized according to scale
    If TypeOf ctl Is Line Then 'deal with lines
     
        LineSizer j, ctl, aForm
        j = j + 1
            
    Else
         'BoxSize sub deals with "Box" controls (listbox, comboBox)
        If TypeOf ctl Is DriveListBox Or TypeOf ctl Is ComboBox Then
            BoxSizer ctl, i, aForm
                    
        Else
             'if control is image setting "stretch" control to true
             'allows resizing of the picture.  ImageSize resizes Image ctl
            If TypeOf ctl Is Image Then
                    
                ctl.Stretch = True
                ImageSizer ctl, i, aForm
                            
            Else
                 'all the other controls are resized here
                With ctlSizeArray(i)
                    aForm.Controls(i).Move .RatioLeft * aForm.ScaleWidth, .RatioTop * aForm.ScaleHeight, _
                    .RatioWidth * aForm.ScaleWidth, .RatioHeight * aForm.ScaleHeight
                End With
                    
            End If
                
        End If
        
    End If
    
    'Can't forget the fonts.  *note:  .TTF fonts dont pixelate when large
    'also if form contains no RichTextBox this has to be commented out
    If TypeName(ctl) = "RichTextBox" Then
        aForm.Controls(i).Font.Size = FontArray(i).TheFont * aForm.ScaleWidth
    Else
        aForm.Controls(i).FontSize = FontArray(i).TheFont * aForm.ScaleWidth
    End If
    
i = i + 1
Next
    
End Sub

Private Sub Linecounter(aForm As Form) 'count lines, if any
Dim ctl As Control
    For Each ctl In aForm
        If TypeOf ctl Is Line Then
            LineCount = LineCount + 1
        End If
    Next
End Sub

'Into the resizing subs I send the control 2b resized, index of control,
'& the form the control resides on
Private Sub BoxSizer(ctl As Control, i As Integer, aForm As Form)
'resize boxes (listbox, combobox), as needed

    With ctlSizeArray(i)
        aForm.Controls(i).Move .RatioLeft * aForm.ScaleWidth
        ctl.Width = .RatioWidth * aForm.ScaleWidth
        ctl.Top = .RatioTop * aForm.ScaleHeight
    End With

End Sub

Private Sub LineSizer(j As Integer, ctl As Control, aForm As Form)
'Get the new width and Height
        NewWidth = aForm.ScaleWidth - OriginalWidth
        NewHeight = aForm.ScaleHeight - OriginalHeight
            
        'figure ratio of change
        HeightChange = aForm.ScaleHeight / OriginalHeight
        WidthChange = aForm.ScaleWidth / OriginalWidth
            
        'resize the line
        With LineSizes(j)
            ctl.Y1 = .LTop * HeightChange
            ctl.Y2 = .LHeight * HeightChange
            ctl.X1 = .LLeft * WidthChange
            ctl.X2 = .LWidth * WidthChange
        End With
            
End Sub

Private Sub ImageSizer(ctl As Control, i As Integer, aForm As Form)
'resize image controls

    With ctlSizeArray(i)
       aForm.Controls(i).Move .RatioLeft * aForm.ScaleWidth, .RatioTop * aForm.ScaleHeight, _
       .RatioWidth * aForm.ScaleWidth, .RatioHeight * aForm.ScaleHeight
    End With

End Sub
