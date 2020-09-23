VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SZAform 
   Caption         =   "Solar Zenith Angle Calculator - RJL 2003"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   Icon            =   "SZA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9255
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   6120
   End
   Begin VB.Frame Frame18 
      Caption         =   "Current Situation"
      Height          =   1695
      Left            =   120
      TabIndex        =   44
      Top             =   7560
      Width           =   10695
      Begin VB.CommandButton Command6 
         Caption         =   "Turn Current situation screen on/off (Turn off to minimise processer usage)"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1320
         Width           =   10335
      End
      Begin VB.Frame Frame24 
         Caption         =   "SZA"
         Height          =   975
         Left            =   4680
         TabIndex        =   54
         Top             =   240
         Width           =   1695
         Begin VB.Label Label8 
            Caption         =   "Label8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Local Time (PC)"
         Height          =   975
         Left            =   2520
         TabIndex        =   51
         Top             =   240
         Width           =   2055
         Begin VB.Label Label7 
            Caption         =   "Label7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame dframe 
         Caption         =   "Date"
         Height          =   975
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   2295
         Begin VB.Label Label6 
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Based on "
         Height          =   975
         Left            =   6480
         TabIndex        =   45
         Top             =   240
         Width           =   3975
         Begin VB.Frame Frame26 
            Caption         =   "Timezone GMT+/-"
            Height          =   615
            Left            =   2280
            TabIndex        =   56
            Top             =   240
            Width           =   1575
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "SZA.frx":0442
               Left            =   120
               List            =   "SZA.frx":048E
               TabIndex        =   57
               Text            =   "0"
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame23 
            Caption         =   "Longitude"
            Height          =   615
            Left            =   1200
            TabIndex        =   47
            Top             =   240
            Width           =   975
            Begin VB.TextBox curlong 
               Height          =   285
               Left            =   120
               TabIndex        =   49
               Text            =   "0"
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Latitude"
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   975
            Begin VB.TextBox curlat 
               Height          =   285
               Left            =   120
               TabIndex        =   48
               Text            =   "0"
               Top             =   240
               Width           =   735
            End
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9960
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame7 
      Caption         =   "Create Record for Specified Location"
      Height          =   5655
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame13 
         Caption         =   "Convert to local time?"
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   4935
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "SZA.frx":04F6
            Left            =   2520
            List            =   "SZA.frx":0542
            TabIndex        =   22
            Text            =   "1"
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Minus"
            Height          =   195
            Left            =   1680
            TabIndex        =   21
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Plus"
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   480
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CheckBox loctimeconvcheck 
            Caption         =   "Convert?"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "hour(s)"
            Height          =   255
            Left            =   3360
            TabIndex        =   23
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Local Time is GMT"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.CheckBox savecheck 
         Caption         =   "Save as text file?"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Frame Frame12 
         Caption         =   "Longitude of Location (E +ve)"
         Height          =   615
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   2295
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "16"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Latitude of Location (N +ve)"
         Height          =   615
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   2295
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Text            =   "69"
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Location Name"
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   4935
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Text            =   "Andoya"
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "start"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Frame Frame9 
         Caption         =   "To"
         Height          =   615
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1455
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "4/9/2"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "From"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "27/7/2"
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calculate Dusk"
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Calculate Dawn"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate Solar Noon"
      Height          =   615
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox msgbox 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   5880
      Width           =   10695
   End
   Begin VB.Frame Frame15 
      Caption         =   "Obtain Single Calculation"
      Height          =   5655
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate SZA"
         Height          =   495
         Left            =   120
         TabIndex        =   58
         Top             =   3960
         Width           =   5055
      End
      Begin VB.Frame Frame17 
         Caption         =   "Ouput"
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   4560
         Width           =   5175
         Begin VB.Frame Frame3 
            Caption         =   "Julian Day"
            Height          =   615
            Left            =   2520
            TabIndex        =   42
            Top             =   240
            Width           =   2055
            Begin VB.TextBox Text4 
               Height          =   285
               Left            =   120
               TabIndex        =   43
               Text            =   "0"
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "SZA"
            Height          =   615
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2295
            Begin VB.Label Label1 
               Caption         =   "SZA"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   480
               TabIndex        =   41
               Top             =   120
               Width           =   1455
            End
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Inputs"
         Height          =   2895
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   5175
         Begin VB.Frame Frame5 
            Caption         =   "Time (Default GMT) - For SZA Calc"
            Height          =   615
            Left            =   1200
            TabIndex        =   59
            Top             =   960
            Width           =   2775
            Begin VB.TextBox Text6 
               Height          =   285
               Left            =   840
               TabIndex        =   60
               Text            =   "20:00:00"
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Is the Time above local time or GMT? If local then complete this:"
            Height          =   1095
            Left            =   120
            TabIndex        =   32
            Top             =   1560
            Width           =   4935
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "SZA.frx":05AA
               Left            =   2520
               List            =   "SZA.frx":05F6
               TabIndex        =   36
               Text            =   "1"
               Top             =   600
               Width           =   735
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Minus"
               Height          =   195
               Left            =   1680
               TabIndex        =   35
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Plus"
               Height          =   255
               Left            =   1680
               TabIndex        =   34
               Top             =   480
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.CheckBox localtimecheck 
               Caption         =   "Time is local time (click this check box if it is)"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   240
               Width           =   4695
            End
            Begin VB.Label Label4 
               Caption         =   "hour(s)"
               Height          =   255
               Left            =   3360
               TabIndex        =   38
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "Local Time is GMT"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   600
               Width           =   1455
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Lat of Site ( N +ve)"
            Height          =   615
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1575
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   120
               TabIndex        =   31
               Text            =   "69"
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Date"
            Height          =   615
            Left            =   3600
            TabIndex        =   28
            Top             =   240
            Width           =   1455
            Begin VB.TextBox Text5 
               Height          =   285
               Left            =   120
               TabIndex        =   29
               Text            =   "16/2/3"
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Long of Site (E +ve)"
            Height          =   615
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   1695
            Begin VB.TextBox Text3 
               Height          =   285
               Left            =   120
               TabIndex        =   27
               Text            =   "16"
               Top             =   240
               Width           =   975
            End
         End
      End
   End
End
Attribute VB_Name = "SZAform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
' This application was written by Roland Leigh (RL40@le.ac.uk)
' Please use it freely, but leave in this comment as an acknowledgement
'****************************************************************************
' This small application demonstrates the use of the SZAfunctions module.
' It provides various facilities for calculating the elevation of the sun.
' Many more permutations of the code are possible, and should be quite easy to achieve
' with some basic code and use of the SZA function.
' I might write a version 2 with some additional features, such as the option to output
' SZA every x minutes for a given day. But I'll see how version 1 goes down first.
' If you have any comments, please email me.


Public Pi As Single


Private Sub Command1_Click()
'One-off SZA calculation
'Get Julian Day for the Text Box
dn = GetJulianDay(Text5.Text)
If localtimecheck.Value = 1 Then
'The line below has been adapted to include the "24-" as a
'Work-around for the fact that the TimeSerial function doesn't cope well with negative numbers
If Option3.Value = True Then timedif = Fix(Combo2.Text) Else timedif = 24 - Fix(Combo2.Text)


szatime = TimeSerial(Hour(Text6.Text) + timedif, Minute(Text6.Text), Second(Text6.Text))
Else: szatime = Text6.Text
End If
bSZA = SZA(Text2.Text, Text3.Text, szatime, Text5.Text)
msgbox.Text = "SZA is " & bSZA & " Degrees"

Text4.Text = dn
Label1.Caption = Round(bSZA, 4) & " Degrees"

End Sub

Private Sub Command2_Click()
'One-off Noon calculation
sn = SolarNoon(Text2.Text, Text3.Text, Text5.Text)
msgbox.Text = "Solar Noon is at " & sn & " GMT"
End Sub

Private Sub Command3_Click()
'Dawn
d = Dawn(Text2.Text, Text3.Text, Text5.Text)
msgbox.Text = "Dawn is at " & d & " GMT"
End Sub

Private Sub Command4_Click()
'Dusk
d = Dusk(Text2.Text, Text3.Text, Text5.Text)
msgbox.Text = "Dusk is at " & d & " GMT"

End Sub

Private Sub Command5_Click()
On Error GoTo eh
'To calculate Dawn, Dusk and Noon for a certain period and place.
startdate = Text1.Text
enddate = Text7.Text
'retains information on the years required
myyear = Year(startdate)
myyear2 = Year(enddate)
sday = GetJulianDay(startdate)
'allows for enddate to be in different year
eday = GetJulianDay(enddate) + ((myyear2 - myyear) * 365)
Location = Text8.Text
'The line below has been adapted to include the "24-" as a
'Work-around for the fact that the TimeSerial function doesn't cope well with negative numbers
If Option1.Value = True Then camptd = Fix(Combo1.Text) Else camptd = 24 - Fix(Combo1.Text)
'This line selects the text for the output header, either with local time definition or GMT.
If loctimeconvcheck.Value = 1 Then gorloc = "Local Time (GMT plus " & camptd & " hours)" Else gorloc = "GMT"
'Asks for savefile information
If savecheck.Value = 1 Then
CommonDialog1.ShowSave
FileName = CommonDialog1.FileName
fn = FreeFile
Open FileName For Output As fn
'Uses this header
Print #fn, "\ SZA Data for " & Location & " at " & Text9.Text & "N and " & Text10.Text & "E"
Print #fn, "\ From " & startdate & " to " & enddate & " (Julian Days " & sday & " to " & eday & ") All Times are " & gorloc
Print #fn, "\ Date, Julian Day, Dawn Time, Noon Time, Dusk Time"
End If
'Also printed to screen
msgbox.Text = "\ SZA Data for " & Location & " at " & Text9.Text & "N and " & Text10.Text & "E"
msgbox.Text = msgbox.Text & vbCrLf & "\ From " & startdate & " to " & enddate & " (Julian Days " & sday & " to " & eday & ") All Times are " & gorloc
msgbox.Text = msgbox.Text & vbCrLf & "\ Date, Julian Day, Dawn Time, Noon Time, Dusk Time"
For a = sday To eday
mydate = DateSerial(myyear, 1, a + 1)
sn = SolarNoon(Text9.Text, Text10.Text, mydate)
sunrise = Dawn(Text9.Text, Text10.Text, mydate)
sunset = Dusk(Text9.Text, Text10.Text, mydate)
If loctimeconvcheck.Value = 1 Then
'Changes for local time if required
sn = TimeSerial(Hour(sn) + camptd, Minute(sn), Second(sn))
sunrise = TimeSerial(Hour(sunrise) + camptd, Minute(sunrise), Second(sunrise))
sunset = TimeSerial(Hour(sunset) + camptd, Minute(sunset), Second(sunset))
End If
'To screen
msgbox.Text = msgbox.Text & vbCrLf & mydate & " " & a & " " & "Dawn: " & sunrise & " " & "Noon " & sn & " " & "Dusk " & sunset
'To file if required
If savecheck.Value = 1 Then Print #fn, mydate, a, sunrise, sn, sunset

SZAform.Refresh

Next a
If savecheck.Value = 1 Then Close fn

Exit Sub
'Quick and simple error handler
eh:
'Turn off timer, otherwise error will keep happening
Timer1.Enabled = False

suc = EHandler()
End Sub

Private Sub Command6_Click()
'Switches current situation display on and off
If Timer1.Enabled = False Then Timer1.Enabled = True Else Timer1.Enabled = False

End Sub

Private Sub Form_Load()
'Initial settings for current situation here
' and all other form settings for that matter..
curlat.Text = 69
curlong.Text = 16
Label7.Caption = Time
Label6.Caption = Date
Text5.Text = Date
Timer1.Enabled = True


End Sub


Private Sub Timer1_Timer()
'This timer updates the current situation section at the bottom.
On Error GoTo eh
clat = curlat.Text
clong = curlong.Text
tdiff = Combo3.Text
'The line below has been adapted to include the "24-" as a
'Work-around for the fact that the TimeSerial function doesn't cope well with negative numbers
If tdiff < 0 Then tdiff = 24 - tdiff
szatime = TimeSerial(Hour(Time) - Fix(tdiff), Minute(Time), Second(Time))
Label7.Caption = Time
Label6.Caption = Date
Label8.Caption = Round(SZA(clat, clong, szatime, Date), 4)
Exit Sub
'Quick and simpler error handler
eh:
suc = EHandler()
End Sub


