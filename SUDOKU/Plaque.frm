VERSION 5.00
Begin VB.Form Startpic 
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   9405
   ClientLeft      =   3840
   ClientTop       =   1050
   ClientWidth     =   7680
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "Plaque.frx":0000
   ScaleHeight     =   9405
   ScaleWidth      =   7680
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TODD"
      BeginProperty Font 
         Name            =   "RockFont"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1920
      TabIndex        =   3
      Top             =   6960
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "KEITH"
      BeginProperty Font 
         Name            =   "RockFont"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BY"
      BeginProperty Font 
         Name            =   "RockFont"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   2040
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUDOKU"
      BeginProperty Font 
         Name            =   "RockFont"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
   End
End
Attribute VB_Name = "Startpic"
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

'This procedure was generated by VB Shaped Form Creator.  This copy has
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
'!:360,3,404,3,405,4,406,4,408,6,411,7,414,10,417,17,418,20,420,25,420,31,421,32,422,35,424,40,426,42,428,47,429,56,429,60,428,61,429,62,429,68,428,69,428,71,427,72,426,75,424,80,422,85,420,90,420,92,419,93,416,93,416,95,415,96,414,99,413,100,414,101,415,104,416,113,417,118,418,123,419,128,420,133,421,140,422,145,423,152,424,159,425,168,425,178,426,179,427,194,428,195,430,209,431,212,432,225,433,230,434,241,435,246,436,257,437,262,437,275,438,276,440,296,441,311,444,330,445,337,446,344,445,347,446,352,447,363,448,370,449,381,450,388,451,399,459,434,460,435,463,442,465,444,467,449,469,451,473,460,475,465,477,467,496,506,500,515,502,520,504,522,505,525,505,527,506,528,509,535,511,537,512,540,512,549,511,550,500,552,499,553,494,554,486,554,485,555,480,556,477,557,472,557,465,564,465,565,463,567,463,568,461,570,461,571,459,573,459,574,457,576,457,577,455,579,455,580,453,582,453,583,448,588,448,590,445,593,445,594,443,596,442,599,439,602,439,603,437,605,EndL
'!:436,608,428,616,425,617,424,616,421,615,420,614,419,615,418,620,418,623,419,624,419,626,418,627,391,627,390,626,379,625,376,624,365,623,362,622,351,621,348,620,343,620,342,619,323,616,313,616,312,615,292,613,275,612,143,612,142,613,50,615,19,615,17,613,16,608,17,607,17,603,18,602,19,599,20,592,21,589,21,587,22,587,26,583,26,577,28,577,29,576,27,541,27,523,26,522,27,521,29,501,29,500,27,498,26,493,26,356,25,355,26,354,26,210,25,209,25,198,24,197,25,196,26,147,26,105,25,104,26,103,26,97,25,96,24,96,22,94,21,94,20,93,20,91,18,89,17,89,14,86,14,85,10,81,10,80,8,78,7,75,5,73,4,68,1,65,0,58,0,49,1,48,1,45,3,43,4,40,4,38,8,34,8,31,9,30,10,27,10,26,16,20,16,19,19,16,19,15,21,13,21,12,22,11,83,12,109,12,110,11,122,11,122,12,123,13,227,13,229,15,236,14,239,15,241,15,243,13,263,13,264,12,272,12,273,11,288,10,291,9,306,8,311,7,326,6,337,6,338,5,EndL
'!:357,4,End
    ReDim PolyPoints(0 To 250)
    For Counter = 0 To 250
        PolyPoints(Counter).X = GP0X(Counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(Counter).Y = GP0Y(Counter) * ScaleY * 15 / STPPY + OffsetY
    Next Counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 251, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 360
    Case 1
        GP0X = 404
    Case 2
        GP0X = 405
    Case 3
        GP0X = 406
    Case 4
        GP0X = 408
    Case 5
        GP0X = 411
    Case 6
        GP0X = 414
    Case 7
        GP0X = 417
    Case 8
        GP0X = 418
    Case 9
        GP0X = 420
    Case 10
        GP0X = 420
    Case 11
        GP0X = 421
    Case 12
        GP0X = 422
    Case 13
        GP0X = 424
    Case 14
        GP0X = 426
    Case 15
        GP0X = 428
    Case 16
        GP0X = 429
    Case 17
        GP0X = 429
    Case 18
        GP0X = 428
    Case 19
        GP0X = 429
    Case 20
        GP0X = 429
    Case 21
        GP0X = 428
    Case 22
        GP0X = 428
    Case 23
        GP0X = 427
    Case 24
        GP0X = 426
    Case 25
        GP0X = 424
    Case 26
        GP0X = 422
    Case 27
        GP0X = 420
    Case 28
        GP0X = 420
    Case 29
        GP0X = 419
    Case 30
        GP0X = 416
    Case 31
        GP0X = 416
    Case 32
        GP0X = 415
    Case 33
        GP0X = 414
    Case 34
        GP0X = 413
    Case 35
        GP0X = 414
    Case 36
        GP0X = 415
    Case 37
        GP0X = 416
    Case 38
        GP0X = 417
    Case 39
        GP0X = 418
    Case 40
        GP0X = 419
    Case 41
        GP0X = 420
    Case 42
        GP0X = 421
    Case 43
        GP0X = 422
    Case 44
        GP0X = 423
    Case 45
        GP0X = 424
    Case 46
        GP0X = 425
    Case 47
        GP0X = 425
    Case 48
        GP0X = 426
    Case 49
        GP0X = 427
    Case 50
        GP0X = 428
    Case 51
        GP0X = 430
    Case 52
        GP0X = 431
    Case 53
        GP0X = 432
    Case 54
        GP0X = 433
    Case 55
        GP0X = 434
    Case 56
        GP0X = 435
    Case 57
        GP0X = 436
    Case 58
        GP0X = 437
    Case 59
        GP0X = 437
    Case 60
        GP0X = 438
    Case 61
        GP0X = 440
    Case 62
        GP0X = 441
    Case 63
        GP0X = 444
    Case 64
        GP0X = 445
    Case 65
        GP0X = 446
    Case 66
        GP0X = 445
    Case 67
        GP0X = 446
    Case 68
        GP0X = 447
    Case 69
        GP0X = 448
    Case 70
        GP0X = 449
    Case 71
        GP0X = 450
    Case 72
        GP0X = 451
    Case 73
        GP0X = 459
    Case 74
        GP0X = 460
    Case 75
        GP0X = 463
    Case 76
        GP0X = 465
    Case 77
        GP0X = 467
    Case 78
        GP0X = 469
    Case 79
        GP0X = 473
    Case 80
        GP0X = 475
    Case 81
        GP0X = 477
    Case 82
        GP0X = 496
    Case 83
        GP0X = 500
    Case 84
        GP0X = 502
    Case 85
        GP0X = 504
    Case 86
        GP0X = 505
    Case 87
        GP0X = 505
    Case 88
        GP0X = 506
    Case 89
        GP0X = 509
    Case 90
        GP0X = 511
    Case 91
        GP0X = 512
    Case 92
        GP0X = 512
    Case 93
        GP0X = 511
    Case 94
        GP0X = 500
    Case 95
        GP0X = 499
    Case 96
        GP0X = 494
    Case 97
        GP0X = 486
    Case 98
        GP0X = 485
    Case 99
        GP0X = 480
    Case 100
        GP0X = 477
    Case 101
        GP0X = 472
    Case 102
        GP0X = 465
    Case 103
        GP0X = 465
    Case 104
        GP0X = 463
    Case 105
        GP0X = 463
    Case 106
        GP0X = 461
    Case 107
        GP0X = 461
    Case 108
        GP0X = 459
    Case 109
        GP0X = 459
    Case 110
        GP0X = 457
    Case 111
        GP0X = 457
    Case 112
        GP0X = 455
    Case 113
        GP0X = 455
    Case 114
        GP0X = 453
    Case 115
        GP0X = 453
    Case 116
        GP0X = 448
    Case 117
        GP0X = 448
    Case 118
        GP0X = 445
    Case 119
        GP0X = 445
    Case 120
        GP0X = 443
    Case 121
        GP0X = 442
    Case 122
        GP0X = 439
    Case 123
        GP0X = 439
    Case 124
        GP0X = 437
    Case 125
        GP0X = 436
    Case 126
        GP0X = 428
    Case 127
        GP0X = 425
    Case 128
        GP0X = 424
    Case 129
        GP0X = 421
    Case 130
        GP0X = 420
    Case 131
        GP0X = 419
    Case 132
        GP0X = 418
    Case 133
        GP0X = 418
    Case 134
        GP0X = 419
    Case 135
        GP0X = 419
    Case 136
        GP0X = 418
    Case 137
        GP0X = 391
    Case 138
        GP0X = 390
    Case 139
        GP0X = 379
    Case 140
        GP0X = 376
    Case 141
        GP0X = 365
    Case 142
        GP0X = 362
    Case 143
        GP0X = 351
    Case 144
        GP0X = 348
    Case 145
        GP0X = 343
    Case 146
        GP0X = 342
    Case 147
        GP0X = 323
    Case 148
        GP0X = 313
    Case 149
        GP0X = 312
    Case 150
        GP0X = 292
    Case 151
        GP0X = 275
    Case 152
        GP0X = 143
    Case 153
        GP0X = 142
    Case 154
        GP0X = 50
    Case 155
        GP0X = 19
    Case 156
        GP0X = 17
    Case 157
        GP0X = 16
    Case 158
        GP0X = 17
    Case 159
        GP0X = 17
    Case 160
        GP0X = 18
    Case 161
        GP0X = 19
    Case 162
        GP0X = 20
    Case 163
        GP0X = 21
    Case 164
        GP0X = 21
    Case 165
        GP0X = 22
    Case 166
        GP0X = 26
    Case 167
        GP0X = 26
    Case 168
        GP0X = 28
    Case 169
        GP0X = 29
    Case 170
        GP0X = 27
    Case 171
        GP0X = 27
    Case 172
        GP0X = 26
    Case 173
        GP0X = 27
    Case 174
        GP0X = 29
    Case 175
        GP0X = 29
    Case 176
        GP0X = 27
    Case 177
        GP0X = 26
    Case 178
        GP0X = 26
    Case 179
        GP0X = 25
    Case 180
        GP0X = 26
    Case 181
        GP0X = 26
    Case 182
        GP0X = 25
    Case 183
        GP0X = 25
    Case 184
        GP0X = 24
    Case 185
        GP0X = 25
    Case 186
        GP0X = 26
    Case 187
        GP0X = 26
    Case 188
        GP0X = 25
    Case 189
        GP0X = 26
    Case 190
        GP0X = 26
    Case 191
        GP0X = 25
    Case 192
        GP0X = 24
    Case 193
        GP0X = 22
    Case 194
        GP0X = 21
    Case 195
        GP0X = 20
    Case 196
        GP0X = 20
    Case 197
        GP0X = 18
    Case 198
        GP0X = 17
    Case 199
        GP0X = 14
    Case 200
        GP0X = 14
    Case 201
        GP0X = 10
    Case 202
        GP0X = 10
    Case 203
        GP0X = 8
    Case 204
        GP0X = 7
    Case 205
        GP0X = 5
    Case 206
        GP0X = 4
    Case 207
        GP0X = 1
    Case 208
        GP0X = 0
    Case 209
        GP0X = 0
    Case 210
        GP0X = 1
    Case 211
        GP0X = 1
    Case 212
        GP0X = 3
    Case 213
        GP0X = 4
    Case 214
        GP0X = 4
    Case 215
        GP0X = 8
    Case 216
        GP0X = 8
    Case 217
        GP0X = 9
    Case 218
        GP0X = 10
    Case 219
        GP0X = 10
    Case 220
        GP0X = 16
    Case 221
        GP0X = 16
    Case 222
        GP0X = 19
    Case 223
        GP0X = 19
    Case 224
        GP0X = 21
    Case 225
        GP0X = 21
    Case 226
        GP0X = 22
    Case 227
        GP0X = 83
    Case 228
        GP0X = 109
    Case 229
        GP0X = 110
    Case 230
        GP0X = 122
    Case 231
        GP0X = 122
    Case 232
        GP0X = 123
    Case 233
        GP0X = 227
    Case 234
        GP0X = 229
    Case 235
        GP0X = 236
    Case 236
        GP0X = 239
    Case 237
        GP0X = 241
    Case 238
        GP0X = 243
    Case 239
        GP0X = 263
    Case 240
        GP0X = 264
    Case 241
        GP0X = 272
    Case 242
        GP0X = 273
    Case 243
        GP0X = 288
    Case 244
        GP0X = 291
    Case 245
        GP0X = 306
    Case 246
        GP0X = 311
    Case 247
        GP0X = 326
    Case 248
        GP0X = 337
    Case 249
        GP0X = 338
    Case 250
        GP0X = 357
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 3
    Case 1
        GP0Y = 3
    Case 2
        GP0Y = 4
    Case 3
        GP0Y = 4
    Case 4
        GP0Y = 6
    Case 5
        GP0Y = 7
    Case 6
        GP0Y = 10
    Case 7
        GP0Y = 17
    Case 8
        GP0Y = 20
    Case 9
        GP0Y = 25
    Case 10
        GP0Y = 31
    Case 11
        GP0Y = 32
    Case 12
        GP0Y = 35
    Case 13
        GP0Y = 40
    Case 14
        GP0Y = 42
    Case 15
        GP0Y = 47
    Case 16
        GP0Y = 56
    Case 17
        GP0Y = 60
    Case 18
        GP0Y = 61
    Case 19
        GP0Y = 62
    Case 20
        GP0Y = 68
    Case 21
        GP0Y = 69
    Case 22
        GP0Y = 71
    Case 23
        GP0Y = 72
    Case 24
        GP0Y = 75
    Case 25
        GP0Y = 80
    Case 26
        GP0Y = 85
    Case 27
        GP0Y = 90
    Case 28
        GP0Y = 92
    Case 29
        GP0Y = 93
    Case 30
        GP0Y = 93
    Case 31
        GP0Y = 95
    Case 32
        GP0Y = 96
    Case 33
        GP0Y = 99
    Case 34
        GP0Y = 100
    Case 35
        GP0Y = 101
    Case 36
        GP0Y = 104
    Case 37
        GP0Y = 113
    Case 38
        GP0Y = 118
    Case 39
        GP0Y = 123
    Case 40
        GP0Y = 128
    Case 41
        GP0Y = 133
    Case 42
        GP0Y = 140
    Case 43
        GP0Y = 145
    Case 44
        GP0Y = 152
    Case 45
        GP0Y = 159
    Case 46
        GP0Y = 168
    Case 47
        GP0Y = 178
    Case 48
        GP0Y = 179
    Case 49
        GP0Y = 194
    Case 50
        GP0Y = 195
    Case 51
        GP0Y = 209
    Case 52
        GP0Y = 212
    Case 53
        GP0Y = 225
    Case 54
        GP0Y = 230
    Case 55
        GP0Y = 241
    Case 56
        GP0Y = 246
    Case 57
        GP0Y = 257
    Case 58
        GP0Y = 262
    Case 59
        GP0Y = 275
    Case 60
        GP0Y = 276
    Case 61
        GP0Y = 296
    Case 62
        GP0Y = 311
    Case 63
        GP0Y = 330
    Case 64
        GP0Y = 337
    Case 65
        GP0Y = 344
    Case 66
        GP0Y = 347
    Case 67
        GP0Y = 352
    Case 68
        GP0Y = 363
    Case 69
        GP0Y = 370
    Case 70
        GP0Y = 381
    Case 71
        GP0Y = 388
    Case 72
        GP0Y = 399
    Case 73
        GP0Y = 434
    Case 74
        GP0Y = 435
    Case 75
        GP0Y = 442
    Case 76
        GP0Y = 444
    Case 77
        GP0Y = 449
    Case 78
        GP0Y = 451
    Case 79
        GP0Y = 460
    Case 80
        GP0Y = 465
    Case 81
        GP0Y = 467
    Case 82
        GP0Y = 506
    Case 83
        GP0Y = 515
    Case 84
        GP0Y = 520
    Case 85
        GP0Y = 522
    Case 86
        GP0Y = 525
    Case 87
        GP0Y = 527
    Case 88
        GP0Y = 528
    Case 89
        GP0Y = 535
    Case 90
        GP0Y = 537
    Case 91
        GP0Y = 540
    Case 92
        GP0Y = 549
    Case 93
        GP0Y = 550
    Case 94
        GP0Y = 552
    Case 95
        GP0Y = 553
    Case 96
        GP0Y = 554
    Case 97
        GP0Y = 554
    Case 98
        GP0Y = 555
    Case 99
        GP0Y = 556
    Case 100
        GP0Y = 557
    Case 101
        GP0Y = 557
    Case 102
        GP0Y = 564
    Case 103
        GP0Y = 565
    Case 104
        GP0Y = 567
    Case 105
        GP0Y = 568
    Case 106
        GP0Y = 570
    Case 107
        GP0Y = 571
    Case 108
        GP0Y = 573
    Case 109
        GP0Y = 574
    Case 110
        GP0Y = 576
    Case 111
        GP0Y = 577
    Case 112
        GP0Y = 579
    Case 113
        GP0Y = 580
    Case 114
        GP0Y = 582
    Case 115
        GP0Y = 583
    Case 116
        GP0Y = 588
    Case 117
        GP0Y = 590
    Case 118
        GP0Y = 593
    Case 119
        GP0Y = 594
    Case 120
        GP0Y = 596
    Case 121
        GP0Y = 599
    Case 122
        GP0Y = 602
    Case 123
        GP0Y = 603
    Case 124
        GP0Y = 605
    Case 125
        GP0Y = 608
    Case 126
        GP0Y = 616
    Case 127
        GP0Y = 617
    Case 128
        GP0Y = 616
    Case 129
        GP0Y = 615
    Case 130
        GP0Y = 614
    Case 131
        GP0Y = 615
    Case 132
        GP0Y = 620
    Case 133
        GP0Y = 623
    Case 134
        GP0Y = 624
    Case 135
        GP0Y = 626
    Case 136
        GP0Y = 627
    Case 137
        GP0Y = 627
    Case 138
        GP0Y = 626
    Case 139
        GP0Y = 625
    Case 140
        GP0Y = 624
    Case 141
        GP0Y = 623
    Case 142
        GP0Y = 622
    Case 143
        GP0Y = 621
    Case 144
        GP0Y = 620
    Case 145
        GP0Y = 620
    Case 146
        GP0Y = 619
    Case 147
        GP0Y = 616
    Case 148
        GP0Y = 616
    Case 149
        GP0Y = 615
    Case 150
        GP0Y = 613
    Case 151
        GP0Y = 612
    Case 152
        GP0Y = 612
    Case 153
        GP0Y = 613
    Case 154
        GP0Y = 615
    Case 155
        GP0Y = 615
    Case 156
        GP0Y = 613
    Case 157
        GP0Y = 608
    Case 158
        GP0Y = 607
    Case 159
        GP0Y = 603
    Case 160
        GP0Y = 602
    Case 161
        GP0Y = 599
    Case 162
        GP0Y = 592
    Case 163
        GP0Y = 589
    Case 164
        GP0Y = 587
    Case 165
        GP0Y = 587
    Case 166
        GP0Y = 583
    Case 167
        GP0Y = 577
    Case 168
        GP0Y = 577
    Case 169
        GP0Y = 576
    Case 170
        GP0Y = 541
    Case 171
        GP0Y = 523
    Case 172
        GP0Y = 522
    Case 173
        GP0Y = 521
    Case 174
        GP0Y = 501
    Case 175
        GP0Y = 500
    Case 176
        GP0Y = 498
    Case 177
        GP0Y = 493
    Case 178
        GP0Y = 356
    Case 179
        GP0Y = 355
    Case 180
        GP0Y = 354
    Case 181
        GP0Y = 210
    Case 182
        GP0Y = 209
    Case 183
        GP0Y = 198
    Case 184
        GP0Y = 197
    Case 185
        GP0Y = 196
    Case 186
        GP0Y = 147
    Case 187
        GP0Y = 105
    Case 188
        GP0Y = 104
    Case 189
        GP0Y = 103
    Case 190
        GP0Y = 97
    Case 191
        GP0Y = 96
    Case 192
        GP0Y = 96
    Case 193
        GP0Y = 94
    Case 194
        GP0Y = 94
    Case 195
        GP0Y = 93
    Case 196
        GP0Y = 91
    Case 197
        GP0Y = 89
    Case 198
        GP0Y = 89
    Case 199
        GP0Y = 86
    Case 200
        GP0Y = 85
    Case 201
        GP0Y = 81
    Case 202
        GP0Y = 80
    Case 203
        GP0Y = 78
    Case 204
        GP0Y = 75
    Case 205
        GP0Y = 73
    Case 206
        GP0Y = 68
    Case 207
        GP0Y = 65
    Case 208
        GP0Y = 58
    Case 209
        GP0Y = 49
    Case 210
        GP0Y = 48
    Case 211
        GP0Y = 45
    Case 212
        GP0Y = 43
    Case 213
        GP0Y = 40
    Case 214
        GP0Y = 38
    Case 215
        GP0Y = 34
    Case 216
        GP0Y = 31
    Case 217
        GP0Y = 30
    Case 218
        GP0Y = 27
    Case 219
        GP0Y = 26
    Case 220
        GP0Y = 20
    Case 221
        GP0Y = 19
    Case 222
        GP0Y = 16
    Case 223
        GP0Y = 15
    Case 224
        GP0Y = 13
    Case 225
        GP0Y = 12
    Case 226
        GP0Y = 11
    Case 227
        GP0Y = 12
    Case 228
        GP0Y = 12
    Case 229
        GP0Y = 11
    Case 230
        GP0Y = 11
    Case 231
        GP0Y = 12
    Case 232
        GP0Y = 13
    Case 233
        GP0Y = 13
    Case 234
        GP0Y = 15
    Case 235
        GP0Y = 14
    Case 236
        GP0Y = 15
    Case 237
        GP0Y = 15
    Case 238
        GP0Y = 13
    Case 239
        GP0Y = 13
    Case 240
        GP0Y = 12
    Case 241
        GP0Y = 12
    Case 242
        GP0Y = 11
    Case 243
        GP0Y = 10
    Case 244
        GP0Y = 9
    Case 245
        GP0Y = 8
    Case 246
        GP0Y = 7
    Case 247
        GP0Y = 6
    Case 248
        GP0Y = 6
    Case 249
        GP0Y = 5
    Case 250
        GP0Y = 4
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
