VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.UserControl UserControl1 
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   Picture         =   "UserControl1.ctx":0000
   PropertyPages   =   "UserControl1.ctx":0282
   ScaleHeight     =   5460
   ScaleWidth      =   7365
   ToolboxBitmap   =   "UserControl1.ctx":02A6
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "UserControl1.ctx":05B8
      Left            =   5760
      List            =   "UserControl1.ctx":05C2
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtalaram 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdstop 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Stop Alaram"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      AutoEnable      =   0   'False
      DeviceType      =   ""
      FileName        =   "C:\WINNT\Media\Mozart's Symphony No. 40.RMI"
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6000
      Top             =   1800
   End
   Begin VB.CommandButton cmdunsetalaram 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "Unset Alaram"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdsetalaram 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Set Alaram"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1200
      Top             =   1800
   End
   Begin VB.Label lblalaram 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Set Alaram Time in Hour/minutes/seconds AM/PM format:"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   3255
      Left            =   0
      Shape           =   2  'Oval
      Top             =   960
      Width           =   7215
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'hello people to use this timer please a Mozart's Symphony No. 40.RMI or any othe music file which u want to have as ur alaram music in the same folder as your application
'and then change the code in ala function in the line no 5 to MMControl1.FileName = App.Path & "\your music file name"
Dim curtime As Date
Dim time1 As Date

Private Sub cmdsetalaram_Click() 'to set the alaram
txtalaram.Visible = True
MMControl1.Command = "Prev"
lblalaram.Visible = True
txtalaram.Text = ""
cmdstop.Enabled = False
Combo1.Visible = True
'Combo1.Text = "AM" uncomment this line and comment next line function call as well as the function timeit at the bottom of this page if u get a error saying combo1.text is a read only property
timeit 'function call
cmdsetalaram.Enabled = False
cmdunsetalaram.Enabled = True
If Timer2.Enabled = False Then Timer2.Enabled = True
End Sub

Private Sub cmdunsetalaram_Click() ' to unset the alaram
Timer2.Enabled = False
cmdsetalaram.Enabled = True
cmdunsetalaram.Enabled = False
lblalaram.Visible = False
txtalaram.Visible = False
Combo1.Visible = False
End Sub

Private Sub cmdstop_Click() 'to stop the alaram
MMControl1.Command = "Stop"
MMControl1.Command = "Next"
Label1.BackColor = RGB(0, 253, 0)
lblalaram.Visible = False
txtalaram.Visible = False
cmdsetalaram.Enabled = True
cmdstop.Enabled = False
cmdunsetalaram.Enabled = False
txtalaram.Locked = False
Combo1.Visible = False
End Sub

Private Sub Timer1_Timer()
If Label1.Caption <> Time Then
Label1.Caption = Time
End If
End Sub

Private Sub Timer2_Timer()
If time1 = Time Then
ala
End If
End Sub

Private Sub txtalaram_Validate(Cancel As Boolean)
Dim str1 As String

str1 = txtalaram.Text
If str1 <> "" And str1 <> " : :  PM" And str1 <> "  :  :  AM" And checktime(str1) = True Then
time1 = CDate(str1)
Cancel = 0
If time1 = Time Then
ala
End If
Else
MsgBox "plz enter proper alaram time", vbSystemModal + vbExclamation, "Alaram Notification"
txtalaram.SelStart = 0
txtalaram.SelLength = Len(txtalaram.Text)
Cancel = 1
txtalaram.SetFocus
End If
End Sub

Private Sub UserControl_Initialize()
lblalaram.Visible = False
txtalaram.Visible = False
Timer1.Enabled = True
Combo1.Visible = False
Timer2.Enabled = True
cmdsetalaram.Enabled = True
cmdstop.Enabled = False
cmdunsetalaram.Enabled = False
End Sub


Public Function checktime(ByRef stringtime As String) As Boolean 'to check the time format
int1 = InStr(1, stringtime, ":")
If int1 > 0 Then
str3 = Mid(stringtime, 1, int1 - 1)
int2 = InStr(int1 + 1, stringtime, ":")
If int2 > 0 Then
str4 = Mid(stringtime, int1 + 1, int2 - 1 - int1)
str5 = Mid(stringtime, int2 + 1, Len(stringtime) - int2)
If IsNumeric(str3) And IsNumeric(str4) And IsNumeric(str5) And (str3 <= 24 And str3 >= 1) And (str4 <= 60 And str4 >= 0) And (str5 <= 60 And str5 >= 0) Then
checktime = True
Else
checktime = False
End If
Else
checktime = False
End If
Else
checktime = False
End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Shape1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Shape1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Shape1,Shape1,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    BorderColor = Shape1.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    Shape1.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=MMControl1,MMControl1,-1,FileName
Public Property Get AlarmMusicfile() As String
Attribute AlarmMusicfile.VB_Description = "Specifies the file to be opened by an Open command or saved by a Save command."
Attribute AlarmMusicfile.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    AlarmMusicfile = MMControl1.FileName
End Property

Public Property Let AlarmMusicfile(ByVal New_AlarmMusicfile As String)
    MMControl1.FileName() = New_AlarmMusicfile
    PropertyChanged "AlarmMusicfile"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Shape1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Shape1.BorderColor = PropBag.ReadProperty("BorderColor", -2147483643)
    MMControl1.FileName = PropBag.ReadProperty("AlarmMusicfile", "C:\WINNT\Media\Mozart's Symphony No. 40.RMI")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Shape1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderColor", Shape1.BorderColor, -2147483643)
    Call PropBag.WriteProperty("AlarmMusicfile", MMControl1.FileName, "C:\WINNT\Media\Mozart's Symphony No. 40.RMI")
End Sub


Public Sub ala() 'to start the alaram
cmdstop.Enabled = True
cmdsetalaram.Enabled = False
cmdunsetalaram.Enabled = False
Label1.BackColor = RGB(256, 0, 0)
MMControl1.FileName = App.Path & "\Mozart's Symphony No. 40.RMI"
MMControl1.Command = "Open"
MMControl1.Command = "Play"
txtalaram.Locked = True
cmdstop.SetFocus
End Sub

Public Sub timeit()
int1 = InStr(1, Time, ":")
str3 = Mid(Time, 1, int1 - 1)
int2 = InStr(int1 + 1, Time, ":")
str4 = Mid(Time, int1 + 1, int2 - 1 - int1)
str5 = Mid(Time, int2 + 1, Len(Time) - int2)
Combo1.Text = Trim(Mid(str5, 3, Len(str5) - 2))
End Sub
