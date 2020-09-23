VERSION 5.00
Begin VB.Form NumForm2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Numbers"
   ClientHeight    =   2535
   ClientLeft      =   6105
   ClientTop       =   4590
   ClientWidth     =   3225
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "Numform2.frx":0000
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "NumForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 7.0"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long
Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

'This Shaped form was generated by VB Shaped Form Creator.  This copy has
'NOT been registered for commercial use.  It may only be used for non-
'profit making programs.  If you intend to sell your program, I think
'it's only fair you pay for mine.  Commercial registration costs $30,
'and can be performed online.  See "Registration" item on the help menu
'for details.

'Latest versions of VB Shaped Form Creator can be found at my website at
'http://www.byalexv.com/VBSFC.html or you can visit my main site
'with many other free programs and utilities at http://www.byalexv.com

'Lines starting with '! are required for reading the form shape using the
'Import Form command in VB Shaped Form Creator, but are not necessary for
'Visual Basic to display the form correctly.

'!Shaped Form Region Definition
'!1,1,0,0,0,0,0,1
'!:92,0,119,0,120,1,124,1,125,2,128,2,130,4,131,4,133,6,134,6,136,8,136,9,138,11,141,10,157,10,158,11,161,11,162,12,164,12,165,13,170,15,173,16,175,18,176,18,176,19,180,23,181,23,185,27,185,28,187,30,188,37,188,42,190,44,194,44,195,45,198,46,200,48,201,48,202,49,202,50,205,53,207,58,208,61,209,64,209,77,207,79,207,84,205,86,205,87,203,89,202,89,200,91,200,92,202,94,203,97,203,107,202,108,202,110,201,111,200,114,198,116,197,119,192,124,192,125,191,126,189,126,187,128,186,128,184,130,183,130,181,132,174,133,166,133,165,136,165,140,164,141,164,142,159,147,158,147,156,149,155,149,152,152,147,154,140,155,133,156,123,156,122,155,114,155,113,154,111,154,109,152,106,151,103,151,102,152,101,152,99,154,94,155,75,155,74,154,70,154,69,153,61,151,59,151,58,150,57,150,55,148,54,148,52,146,51,146,47,142,47,141,45,139,45,135,43,133,33,133,32,132,29,132,28,131,26,131,25,130,18,127,16,125,15,125,14,124,14,123,EndL
'!:10,119,8,114,7,109,7,96,8,95,8,94,10,92,10,91,8,89,7,89,4,86,4,85,2,83,1,80,0,73,1,64,2,61,2,60,3,59,1,57,1,56,3,56,4,55,3,54,3,53,5,53,8,50,8,49,9,48,10,48,12,46,17,44,20,44,21,43,21,42,22,42,23,41,22,40,22,33,23,32,23,30,24,29,24,28,25,28,26,27,26,25,27,25,31,21,32,21,34,19,34,18,35,18,37,16,40,15,47,12,58,9,67,9,68,10,70,10,71,11,72,11,74,9,74,8,76,6,77,6,79,4,82,3,85,2,90,1,End
    ReDim PolyPoints(0 To 193)
    For Counter = 0 To 193
        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP0Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 194, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 92
    Case 1
        GP0X = 119
    Case 2
        GP0X = 120
    Case 3
        GP0X = 124
    Case 4
        GP0X = 125
    Case 5
        GP0X = 128
    Case 6
        GP0X = 130
    Case 7
        GP0X = 131
    Case 8
        GP0X = 133
    Case 9
        GP0X = 134
    Case 10
        GP0X = 136
    Case 11
        GP0X = 136
    Case 12
        GP0X = 138
    Case 13
        GP0X = 141
    Case 14
        GP0X = 157
    Case 15
        GP0X = 158
    Case 16
        GP0X = 161
    Case 17
        GP0X = 162
    Case 18
        GP0X = 164
    Case 19
        GP0X = 165
    Case 20
        GP0X = 170
    Case 21
        GP0X = 173
    Case 22
        GP0X = 175
    Case 23
        GP0X = 176
    Case 24
        GP0X = 176
    Case 25
        GP0X = 180
    Case 26
        GP0X = 181
    Case 27
        GP0X = 185
    Case 28
        GP0X = 185
    Case 29
        GP0X = 187
    Case 30
        GP0X = 188
    Case 31
        GP0X = 188
    Case 32
        GP0X = 190
    Case 33
        GP0X = 194
    Case 34
        GP0X = 195
    Case 35
        GP0X = 198
    Case 36
        GP0X = 200
    Case 37
        GP0X = 201
    Case 38
        GP0X = 202
    Case 39
        GP0X = 202
    Case 40
        GP0X = 205
    Case 41
        GP0X = 207
    Case 42
        GP0X = 208
    Case 43
        GP0X = 209
    Case 44
        GP0X = 209
    Case 45
        GP0X = 207
    Case 46
        GP0X = 207
    Case 47
        GP0X = 205
    Case 48
        GP0X = 205
    Case 49
        GP0X = 203
    Case 50
        GP0X = 202
    Case 51
        GP0X = 200
    Case 52
        GP0X = 200
    Case 53
        GP0X = 202
    Case 54
        GP0X = 203
    Case 55
        GP0X = 203
    Case 56
        GP0X = 202
    Case 57
        GP0X = 202
    Case 58
        GP0X = 201
    Case 59
        GP0X = 200
    Case 60
        GP0X = 198
    Case 61
        GP0X = 197
    Case 62
        GP0X = 192
    Case 63
        GP0X = 192
    Case 64
        GP0X = 191
    Case 65
        GP0X = 189
    Case 66
        GP0X = 187
    Case 67
        GP0X = 186
    Case 68
        GP0X = 184
    Case 69
        GP0X = 183
    Case 70
        GP0X = 181
    Case 71
        GP0X = 174
    Case 72
        GP0X = 166
    Case 73
        GP0X = 165
    Case 74
        GP0X = 165
    Case 75
        GP0X = 164
    Case 76
        GP0X = 164
    Case 77
        GP0X = 159
    Case 78
        GP0X = 158
    Case 79
        GP0X = 156
    Case 80
        GP0X = 155
    Case 81
        GP0X = 152
    Case 82
        GP0X = 147
    Case 83
        GP0X = 140
    Case 84
        GP0X = 133
    Case 85
        GP0X = 123
    Case 86
        GP0X = 122
    Case 87
        GP0X = 114
    Case 88
        GP0X = 113
    Case 89
        GP0X = 111
    Case 90
        GP0X = 109
    Case 91
        GP0X = 106
    Case 92
        GP0X = 103
    Case 93
        GP0X = 102
    Case 94
        GP0X = 101
    Case 95
        GP0X = 99
    Case 96
        GP0X = 94
    Case 97
        GP0X = 75
    Case 98
        GP0X = 74
    Case 99
        GP0X = 70
    Case 100
        GP0X = 69
    Case 101
        GP0X = 61
    Case 102
        GP0X = 59
    Case 103
        GP0X = 58
    Case 104
        GP0X = 57
    Case 105
        GP0X = 55
    Case 106
        GP0X = 54
    Case 107
        GP0X = 52
    Case 108
        GP0X = 51
    Case 109
        GP0X = 47
    Case 110
        GP0X = 47
    Case 111
        GP0X = 45
    Case 112
        GP0X = 45
    Case 113
        GP0X = 43
    Case 114
        GP0X = 33
    Case 115
        GP0X = 32
    Case 116
        GP0X = 29
    Case 117
        GP0X = 28
    Case 118
        GP0X = 26
    Case 119
        GP0X = 25
    Case 120
        GP0X = 18
    Case 121
        GP0X = 16
    Case 122
        GP0X = 15
    Case 123
        GP0X = 14
    Case 124
        GP0X = 14
    Case 125
        GP0X = 10
    Case 126
        GP0X = 8
    Case 127
        GP0X = 7
    Case 128
        GP0X = 7
    Case 129
        GP0X = 8
    Case 130
        GP0X = 8
    Case 131
        GP0X = 10
    Case 132
        GP0X = 10
    Case 133
        GP0X = 8
    Case 134
        GP0X = 7
    Case 135
        GP0X = 4
    Case 136
        GP0X = 4
    Case 137
        GP0X = 2
    Case 138
        GP0X = 1
    Case 139
        GP0X = 0
    Case 140
        GP0X = 1
    Case 141
        GP0X = 2
    Case 142
        GP0X = 2
    Case 143
        GP0X = 3
    Case 144
        GP0X = 1
    Case 145
        GP0X = 1
    Case 146
        GP0X = 3
    Case 147
        GP0X = 4
    Case 148
        GP0X = 3
    Case 149
        GP0X = 3
    Case 150
        GP0X = 5
    Case 151
        GP0X = 8
    Case 152
        GP0X = 8
    Case 153
        GP0X = 9
    Case 154
        GP0X = 10
    Case 155
        GP0X = 12
    Case 156
        GP0X = 17
    Case 157
        GP0X = 20
    Case 158
        GP0X = 21
    Case 159
        GP0X = 21
    Case 160
        GP0X = 22
    Case 161
        GP0X = 23
    Case 162
        GP0X = 22
    Case 163
        GP0X = 22
    Case 164
        GP0X = 23
    Case 165
        GP0X = 23
    Case 166
        GP0X = 24
    Case 167
        GP0X = 24
    Case 168
        GP0X = 25
    Case 169
        GP0X = 26
    Case 170
        GP0X = 26
    Case 171
        GP0X = 27
    Case 172
        GP0X = 31
    Case 173
        GP0X = 32
    Case 174
        GP0X = 34
    Case 175
        GP0X = 34
    Case 176
        GP0X = 35
    Case 177
        GP0X = 37
    Case 178
        GP0X = 40
    Case 179
        GP0X = 47
    Case 180
        GP0X = 58
    Case 181
        GP0X = 67
    Case 182
        GP0X = 68
    Case 183
        GP0X = 70
    Case 184
        GP0X = 71
    Case 185
        GP0X = 72
    Case 186
        GP0X = 74
    Case 187
        GP0X = 74
    Case 188
        GP0X = 76
    Case 189
        GP0X = 77
    Case 190
        GP0X = 79
    Case 191
        GP0X = 82
    Case 192
        GP0X = 85
    Case 193
        GP0X = 90
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 0
    Case 1
        GP0Y = 0
    Case 2
        GP0Y = 1
    Case 3
        GP0Y = 1
    Case 4
        GP0Y = 2
    Case 5
        GP0Y = 2
    Case 6
        GP0Y = 4
    Case 7
        GP0Y = 4
    Case 8
        GP0Y = 6
    Case 9
        GP0Y = 6
    Case 10
        GP0Y = 8
    Case 11
        GP0Y = 9
    Case 12
        GP0Y = 11
    Case 13
        GP0Y = 10
    Case 14
        GP0Y = 10
    Case 15
        GP0Y = 11
    Case 16
        GP0Y = 11
    Case 17
        GP0Y = 12
    Case 18
        GP0Y = 12
    Case 19
        GP0Y = 13
    Case 20
        GP0Y = 15
    Case 21
        GP0Y = 16
    Case 22
        GP0Y = 18
    Case 23
        GP0Y = 18
    Case 24
        GP0Y = 19
    Case 25
        GP0Y = 23
    Case 26
        GP0Y = 23
    Case 27
        GP0Y = 27
    Case 28
        GP0Y = 28
    Case 29
        GP0Y = 30
    Case 30
        GP0Y = 37
    Case 31
        GP0Y = 42
    Case 32
        GP0Y = 44
    Case 33
        GP0Y = 44
    Case 34
        GP0Y = 45
    Case 35
        GP0Y = 46
    Case 36
        GP0Y = 48
    Case 37
        GP0Y = 48
    Case 38
        GP0Y = 49
    Case 39
        GP0Y = 50
    Case 40
        GP0Y = 53
    Case 41
        GP0Y = 58
    Case 42
        GP0Y = 61
    Case 43
        GP0Y = 64
    Case 44
        GP0Y = 77
    Case 45
        GP0Y = 79
    Case 46
        GP0Y = 84
    Case 47
        GP0Y = 86
    Case 48
        GP0Y = 87
    Case 49
        GP0Y = 89
    Case 50
        GP0Y = 89
    Case 51
        GP0Y = 91
    Case 52
        GP0Y = 92
    Case 53
        GP0Y = 94
    Case 54
        GP0Y = 97
    Case 55
        GP0Y = 107
    Case 56
        GP0Y = 108
    Case 57
        GP0Y = 110
    Case 58
        GP0Y = 111
    Case 59
        GP0Y = 114
    Case 60
        GP0Y = 116
    Case 61
        GP0Y = 119
    Case 62
        GP0Y = 124
    Case 63
        GP0Y = 125
    Case 64
        GP0Y = 126
    Case 65
        GP0Y = 126
    Case 66
        GP0Y = 128
    Case 67
        GP0Y = 128
    Case 68
        GP0Y = 130
    Case 69
        GP0Y = 130
    Case 70
        GP0Y = 132
    Case 71
        GP0Y = 133
    Case 72
        GP0Y = 133
    Case 73
        GP0Y = 136
    Case 74
        GP0Y = 140
    Case 75
        GP0Y = 141
    Case 76
        GP0Y = 142
    Case 77
        GP0Y = 147
    Case 78
        GP0Y = 147
    Case 79
        GP0Y = 149
    Case 80
        GP0Y = 149
    Case 81
        GP0Y = 152
    Case 82
        GP0Y = 154
    Case 83
        GP0Y = 155
    Case 84
        GP0Y = 156
    Case 85
        GP0Y = 156
    Case 86
        GP0Y = 155
    Case 87
        GP0Y = 155
    Case 88
        GP0Y = 154
    Case 89
        GP0Y = 154
    Case 90
        GP0Y = 152
    Case 91
        GP0Y = 151
    Case 92
        GP0Y = 151
    Case 93
        GP0Y = 152
    Case 94
        GP0Y = 152
    Case 95
        GP0Y = 154
    Case 96
        GP0Y = 155
    Case 97
        GP0Y = 155
    Case 98
        GP0Y = 154
    Case 99
        GP0Y = 154
    Case 100
        GP0Y = 153
    Case 101
        GP0Y = 151
    Case 102
        GP0Y = 151
    Case 103
        GP0Y = 150
    Case 104
        GP0Y = 150
    Case 105
        GP0Y = 148
    Case 106
        GP0Y = 148
    Case 107
        GP0Y = 146
    Case 108
        GP0Y = 146
    Case 109
        GP0Y = 142
    Case 110
        GP0Y = 141
    Case 111
        GP0Y = 139
    Case 112
        GP0Y = 135
    Case 113
        GP0Y = 133
    Case 114
        GP0Y = 133
    Case 115
        GP0Y = 132
    Case 116
        GP0Y = 132
    Case 117
        GP0Y = 131
    Case 118
        GP0Y = 131
    Case 119
        GP0Y = 130
    Case 120
        GP0Y = 127
    Case 121
        GP0Y = 125
    Case 122
        GP0Y = 125
    Case 123
        GP0Y = 124
    Case 124
        GP0Y = 123
    Case 125
        GP0Y = 119
    Case 126
        GP0Y = 114
    Case 127
        GP0Y = 109
    Case 128
        GP0Y = 96
    Case 129
        GP0Y = 95
    Case 130
        GP0Y = 94
    Case 131
        GP0Y = 92
    Case 132
        GP0Y = 91
    Case 133
        GP0Y = 89
    Case 134
        GP0Y = 89
    Case 135
        GP0Y = 86
    Case 136
        GP0Y = 85
    Case 137
        GP0Y = 83
    Case 138
        GP0Y = 80
    Case 139
        GP0Y = 73
    Case 140
        GP0Y = 64
    Case 141
        GP0Y = 61
    Case 142
        GP0Y = 60
    Case 143
        GP0Y = 59
    Case 144
        GP0Y = 57
    Case 145
        GP0Y = 56
    Case 146
        GP0Y = 56
    Case 147
        GP0Y = 55
    Case 148
        GP0Y = 54
    Case 149
        GP0Y = 53
    Case 150
        GP0Y = 53
    Case 151
        GP0Y = 50
    Case 152
        GP0Y = 49
    Case 153
        GP0Y = 48
    Case 154
        GP0Y = 48
    Case 155
        GP0Y = 46
    Case 156
        GP0Y = 44
    Case 157
        GP0Y = 44
    Case 158
        GP0Y = 43
    Case 159
        GP0Y = 42
    Case 160
        GP0Y = 42
    Case 161
        GP0Y = 41
    Case 162
        GP0Y = 40
    Case 163
        GP0Y = 33
    Case 164
        GP0Y = 32
    Case 165
        GP0Y = 30
    Case 166
        GP0Y = 29
    Case 167
        GP0Y = 28
    Case 168
        GP0Y = 28
    Case 169
        GP0Y = 27
    Case 170
        GP0Y = 25
    Case 171
        GP0Y = 25
    Case 172
        GP0Y = 21
    Case 173
        GP0Y = 21
    Case 174
        GP0Y = 19
    Case 175
        GP0Y = 18
    Case 176
        GP0Y = 18
    Case 177
        GP0Y = 16
    Case 178
        GP0Y = 15
    Case 179
        GP0Y = 12
    Case 180
        GP0Y = 9
    Case 181
        GP0Y = 9
    Case 182
        GP0Y = 10
    Case 183
        GP0Y = 10
    Case 184
        GP0Y = 11
    Case 185
        GP0Y = 11
    Case 186
        GP0Y = 9
    Case 187
        GP0Y = 8
    Case 188
        GP0Y = 6
    Case 189
        GP0Y = 6
    Case 190
        GP0Y = 4
    Case 191
        GP0Y = 3
    Case 192
        GP0Y = 2
    Case 193
        GP0Y = 1
    End Select
End Function

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
'    ReleaseCapture
'    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub
Private Sub label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim P As Integer

 FormMain.Label1.Caption = ""
 With Label1(Index)
   If Button = vbLeftButton Then
      FormMain.Command1(CurrentIndex).Font = "Igloo"
      FormMain.Command1(CurrentIndex).FontSize = 26
      FormMain.Command1(CurrentIndex).Caption = Label1(Index)
   Else
     .Caption = ""
   End If
 End With
  
 NumForm2.Visible = False
 Call FormMain.Finished
 If FormMain.Finishok Then Call FormMain.Over

 End Sub

Private Sub Timer1_Timer()
    Dim A, B, C, D, I As Integer
    ' FormName, ControlName, Control.Left In Pixels, Control.Right In Pixels, Control.Top In Pixels, Control.Bottom In Pixels
    
   For I = 0 To 8
     A = Label1(I).Left
     B = Label1(I).Left + Label1(I).Width
     C = Label1(I).Top
     D = Label1(I).Top + Label1(I).Height
     MouseMove NumForm2, Label1(I), Val(A), Val(B), Val(C), Val(D)
   Next I
End Sub

