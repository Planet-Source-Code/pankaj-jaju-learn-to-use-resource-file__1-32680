VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Learn How To Use Resource File"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   675
      Width           =   5220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About Me"
      Height          =   375
      Left            =   5505
      TabIndex        =   1
      Top             =   675
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   5220
   End
   Begin VB.CommandButton Command1 
      Caption         =   "My Country"
      Height          =   375
      Left            =   5505
      TabIndex        =   0
      Top             =   150
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   5505
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   150
      TabIndex        =   4
      Top             =   1200
      Width           =   5220
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' (((((((( (((((((( (((    ((( (((  (((( (((((((( (((((((((
' )))  ))) )))  ))) ))))   ))) ))) ))))  )))  )))     )))
' (((((((( (((((((( ((( (( ((( (((((((   ((((((((     (((
' )))      )))))))) )))  ))))) ))) ))))  ))))))))     )))
' (((      (((  ((( (((   (((( (((   ((( (((  ((( ((  (((
' )))      )))  ))) )))    ))) (((   ((( (((  ((( )))))))
'
'«··´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸..·´¯`·.
'INDIA IS GREAT                                          .·´
'Pankaj Jaju                                             ´·.
'E-mail:- pankaj_jaju@rediffmail.com                      .·´
'PLEASE SEND YOUR VALUABLE SUGGESTIONS                  ´·.
'«·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·.¸_¸.·´¯`·..·´
'
'THIS IS MY THIRD UPLOAD (previously Imgviewer and LoveThermometer)
'RESOLUTION:- 800x600(24-bit color)
'******************************************************************************************
'******************************************************************************************
'
'Advantages of Resource Files:-
'-------------------------------------------------
'Firstly, Performance is increased because strings, bitmaps, icons,
'and data can be loaded on demand from the resource file, instead
'of all being loaded at once when a form is loaded.
'Secondly, You can develope an application which will be suitable for
'different regional settings i.e for different languages or locations
'
'How to create a Resource file:-
'-------------------------
'Step1- Select Add New Resource File from the Project menu.
'       This command is only available when the Resource Editor Add-In is loaded.
'       To load the Resource Editor Add-In, select Add-In Manager from the Add-Ins menu.
'       In the Add-In Manager dialog box, select VB6 Resource Editor and check
'       the Loaded/Unloaded box.
'Step2- Select the Save button on the Resource Editor toolbar to save the resource file.
'       The file will be added to the Project Explorer
'
'Note:- You can have only one resource file in your project.
'*****  .RES files are bit-specific i.e Can't use 16-bit .RES file for 32-bit project
'
'******************************************************************************************
'******************************************************************************************

Option Explicit

Private Sub Command1_Click()
    Text1.Text = LoadResString(101) 'loads string from .RES file
End Sub

Private Sub Command1_GotFocus()
    Label1.Caption = "Tutorial Code Created By" & vbCrLf _
                   & "PANKAJ JAJU (pankaj_jaju@rediffmail.com)" & vbCrLf & vbCrLf _
                   & "Make your Regional Setting ''English(United States)''" & vbCrLf _
                   & "Now Click both the Buttons simultaneously"
    Text1.Text = ""
    Text2.Text = ""
    
    'loads bitmap from .RES file
    Image1.Picture = LoadResPicture(101, vbResBitmap) 'OR vbResIcon OR vbResCursor
End Sub

Private Sub Command2_Click()
    Text2.Text = LoadResString(102) 'loads string from .RES file
    
    If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Mid$(Text1.Text, 1, 5) = "INDIA" Then
        Label1.Caption = "Now Change Regional Setting Of Your Computer From" _
                         & "''English(United States)'' To ''English(United Kingdom)''" _
                         & "And Click Again Both Command Buttons" & vbCrLf & vbCrLf _
                         & "To Change Regional Setting :- Open Control Panel And " _
                         & "Then Click Regional Settings"
    ElseIf Len(Text1.Text) > 0 And Len(Text2.Text) > 0 And Mid$(Text1.Text, 1, 6) = "BHARAT" Then
        Label1.Caption = vbCrLf & "Now you can create an application that uses RES file" _
                         & vbCrLf & vbCrLf & "So go on and show to your friends that you are" _
                         & vbCrLf & "G E N I U S"
    Else
        Label1.Caption = vbCrLf & "Dont' forget to rate this code" & vbCrLf & vbCrLf _
                        & "Also mail your suggestions or queries at pankaj_jaju@rediffmail.com"
    End If
    
    'loads bitmap from .RES file
    Image1.Picture = LoadResPicture(102, vbResBitmap) 'OR vbResIcon OR vbResCursor
End Sub
