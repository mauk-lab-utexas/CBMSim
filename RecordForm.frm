VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RecordForm 
   BackColor       =   &H00000000&
   Caption         =   "Cells to record"
   ClientHeight    =   13425
   ClientLeft      =   30
   ClientTop       =   10485
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   ScaleHeight     =   13425
   ScaleWidth      =   6180
   Visible         =   0   'False
   Begin VB.TextBox ScaleText 
      Height          =   375
      Left            =   7080
      TabIndex        =   437
      Text            =   "1200"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton ScaleButton 
      Caption         =   "Done"
      Height          =   375
      Left            =   7080
      TabIndex        =   436
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CSDurationButton 
      Caption         =   "Done"
      Height          =   375
      Left            =   5640
      TabIndex        =   435
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox CSDurationText 
      Height          =   375
      Left            =   5640
      TabIndex        =   434
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.ras|*.ras"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   12975
      Left            =   360
      TabIndex        =   397
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   35
         Left            =   120
         TabIndex        =   433
         Top             =   12720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   34
         Left            =   120
         TabIndex        =   432
         Top             =   12360
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   33
         Left            =   120
         TabIndex        =   431
         Top             =   12000
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   430
         Top             =   11640
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   31
         Left            =   120
         TabIndex        =   429
         Top             =   11280
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   428
         Top             =   10920
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   29
         Left            =   120
         TabIndex        =   427
         Top             =   10560
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   28
         Left            =   120
         TabIndex        =   426
         Top             =   10200
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   27
         Left            =   120
         TabIndex        =   425
         Top             =   9840
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   26
         Left            =   120
         TabIndex        =   424
         Top             =   9480
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   25
         Left            =   120
         TabIndex        =   423
         Top             =   9120
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   422
         Top             =   8760
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   23
         Left            =   120
         TabIndex        =   421
         Top             =   8400
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   420
         Top             =   8040
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   419
         Top             =   7680
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   418
         Top             =   7320
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   417
         Top             =   6960
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   18
         Left            =   120
         TabIndex        =   416
         Top             =   6600
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   415
         Top             =   6240
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   414
         Top             =   5880
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   413
         Top             =   5520
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   412
         Top             =   5160
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   411
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   410
         Top             =   4440
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   409
         Top             =   4080
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   408
         Top             =   3720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   407
         Top             =   3360
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   406
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   405
         Top             =   2640
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   404
         Top             =   2280
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   403
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   402
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   401
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   400
         Top             =   840
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   399
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Option2"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   398
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   11
      Left            =   360
      TabIndex        =   388
      Top             =   4320
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   95
         Left            =   3360
         TabIndex        =   396
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   94
         Left            =   2880
         TabIndex        =   395
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   93
         Left            =   2400
         TabIndex        =   394
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   92
         Left            =   1920
         TabIndex        =   393
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   91
         Left            =   1440
         TabIndex        =   392
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   90
         Left            =   960
         TabIndex        =   391
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   89
         Left            =   480
         TabIndex        =   390
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   88
         Left            =   0
         TabIndex        =   389
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   379
      Top             =   2160
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   47
         Left            =   3360
         TabIndex        =   387
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   46
         Left            =   2880
         TabIndex        =   386
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   45
         Left            =   2400
         TabIndex        =   385
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   44
         Left            =   1920
         TabIndex        =   384
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   43
         Left            =   1440
         TabIndex        =   383
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   42
         Left            =   960
         TabIndex        =   382
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   41
         Left            =   480
         TabIndex        =   381
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   40
         Left            =   0
         TabIndex        =   380
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   370
      Top             =   2520
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   55
         Left            =   3360
         TabIndex        =   378
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   54
         Left            =   2880
         TabIndex        =   377
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   53
         Left            =   2400
         TabIndex        =   376
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   52
         Left            =   1920
         TabIndex        =   375
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   51
         Left            =   1440
         TabIndex        =   374
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   50
         Left            =   960
         TabIndex        =   373
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   49
         Left            =   480
         TabIndex        =   372
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   48
         Left            =   0
         TabIndex        =   371
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   22
      Left            =   360
      TabIndex        =   361
      Top             =   8280
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   176
         Left            =   0
         TabIndex        =   369
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   177
         Left            =   480
         TabIndex        =   368
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   178
         Left            =   960
         TabIndex        =   367
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   179
         Left            =   1440
         TabIndex        =   366
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   180
         Left            =   1920
         TabIndex        =   365
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   181
         Left            =   2400
         TabIndex        =   364
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   182
         Left            =   2880
         TabIndex        =   363
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   183
         Left            =   3360
         TabIndex        =   362
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   352
      Top             =   1800
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   39
         Left            =   3360
         TabIndex        =   360
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   38
         Left            =   2880
         TabIndex        =   359
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   37
         Left            =   2400
         TabIndex        =   358
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   36
         Left            =   1920
         TabIndex        =   357
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   35
         Left            =   1440
         TabIndex        =   356
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   34
         Left            =   960
         TabIndex        =   355
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   33
         Left            =   480
         TabIndex        =   354
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   32
         Left            =   0
         TabIndex        =   353
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   35
      Left            =   360
      TabIndex        =   343
      Top             =   12960
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   287
         Left            =   3360
         TabIndex        =   351
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   286
         Left            =   2880
         TabIndex        =   350
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   285
         Left            =   2400
         TabIndex        =   349
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   284
         Left            =   1920
         TabIndex        =   348
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   283
         Left            =   1440
         TabIndex        =   347
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   282
         Left            =   960
         TabIndex        =   346
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   281
         Left            =   480
         TabIndex        =   345
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   280
         Left            =   0
         TabIndex        =   344
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   34
      Left            =   360
      TabIndex        =   334
      Top             =   12600
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   279
         Left            =   3360
         TabIndex        =   342
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   278
         Left            =   2880
         TabIndex        =   341
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   277
         Left            =   2400
         TabIndex        =   340
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   276
         Left            =   1920
         TabIndex        =   339
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   275
         Left            =   1440
         TabIndex        =   338
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   274
         Left            =   960
         TabIndex        =   337
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   273
         Left            =   480
         TabIndex        =   336
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   272
         Left            =   0
         TabIndex        =   335
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   33
      Left            =   360
      TabIndex        =   325
      Top             =   12240
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   271
         Left            =   3360
         TabIndex        =   333
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   270
         Left            =   2880
         TabIndex        =   332
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   269
         Left            =   2400
         TabIndex        =   331
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   268
         Left            =   1920
         TabIndex        =   330
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   267
         Left            =   1440
         TabIndex        =   329
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   266
         Left            =   960
         TabIndex        =   328
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   265
         Left            =   480
         TabIndex        =   327
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   264
         Left            =   0
         TabIndex        =   326
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   32
      Left            =   360
      TabIndex        =   316
      Top             =   11880
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   263
         Left            =   3360
         TabIndex        =   324
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   262
         Left            =   2880
         TabIndex        =   323
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   261
         Left            =   2400
         TabIndex        =   322
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   260
         Left            =   1920
         TabIndex        =   321
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   259
         Left            =   1440
         TabIndex        =   320
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   258
         Left            =   960
         TabIndex        =   319
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   257
         Left            =   480
         TabIndex        =   318
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   256
         Left            =   0
         TabIndex        =   317
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   31
      Left            =   360
      TabIndex        =   307
      Top             =   11520
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   255
         Left            =   3360
         TabIndex        =   315
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   254
         Left            =   2880
         TabIndex        =   314
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   253
         Left            =   2400
         TabIndex        =   313
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   252
         Left            =   1920
         TabIndex        =   312
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   251
         Left            =   1440
         TabIndex        =   311
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   250
         Left            =   960
         TabIndex        =   310
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   249
         Left            =   480
         TabIndex        =   309
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   248
         Left            =   0
         TabIndex        =   308
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   30
      Left            =   360
      TabIndex        =   298
      Top             =   11160
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   247
         Left            =   3360
         TabIndex        =   306
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   246
         Left            =   2880
         TabIndex        =   305
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   245
         Left            =   2400
         TabIndex        =   304
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   244
         Left            =   1920
         TabIndex        =   303
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   243
         Left            =   1440
         TabIndex        =   302
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   242
         Left            =   960
         TabIndex        =   301
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   241
         Left            =   480
         TabIndex        =   300
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   240
         Left            =   0
         TabIndex        =   299
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   29
      Left            =   360
      TabIndex        =   289
      Top             =   10800
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   239
         Left            =   3360
         TabIndex        =   297
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   238
         Left            =   2880
         TabIndex        =   296
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   237
         Left            =   2400
         TabIndex        =   295
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   236
         Left            =   1920
         TabIndex        =   294
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   235
         Left            =   1440
         TabIndex        =   293
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   234
         Left            =   960
         TabIndex        =   292
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   233
         Left            =   480
         TabIndex        =   291
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   232
         Left            =   0
         TabIndex        =   290
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   28
      Left            =   360
      TabIndex        =   280
      Top             =   10440
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   231
         Left            =   3360
         TabIndex        =   288
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   230
         Left            =   2880
         TabIndex        =   287
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   229
         Left            =   2400
         TabIndex        =   286
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   228
         Left            =   1920
         TabIndex        =   285
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   227
         Left            =   1440
         TabIndex        =   284
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   226
         Left            =   960
         TabIndex        =   283
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   225
         Left            =   480
         TabIndex        =   282
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   224
         Left            =   0
         TabIndex        =   281
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   27
      Left            =   360
      TabIndex        =   271
      Top             =   10080
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   223
         Left            =   3360
         TabIndex        =   279
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   222
         Left            =   2880
         TabIndex        =   278
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   221
         Left            =   2400
         TabIndex        =   277
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   220
         Left            =   1920
         TabIndex        =   276
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   219
         Left            =   1440
         TabIndex        =   275
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   218
         Left            =   960
         TabIndex        =   274
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   217
         Left            =   480
         TabIndex        =   273
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   216
         Left            =   0
         TabIndex        =   272
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   26
      Left            =   360
      TabIndex        =   262
      Top             =   9720
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   215
         Left            =   3360
         TabIndex        =   270
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   214
         Left            =   2880
         TabIndex        =   269
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   213
         Left            =   2400
         TabIndex        =   268
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   212
         Left            =   1920
         TabIndex        =   267
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   211
         Left            =   1440
         TabIndex        =   266
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   210
         Left            =   960
         TabIndex        =   265
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   209
         Left            =   480
         TabIndex        =   264
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   208
         Left            =   0
         TabIndex        =   263
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   25
      Left            =   360
      TabIndex        =   253
      Top             =   9360
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   207
         Left            =   3360
         TabIndex        =   261
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   206
         Left            =   2880
         TabIndex        =   260
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   205
         Left            =   2400
         TabIndex        =   259
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   204
         Left            =   1920
         TabIndex        =   258
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   203
         Left            =   1440
         TabIndex        =   257
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   202
         Left            =   960
         TabIndex        =   256
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   201
         Left            =   480
         TabIndex        =   255
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   200
         Left            =   0
         TabIndex        =   254
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   24
      Left            =   360
      TabIndex        =   244
      Top             =   9000
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   199
         Left            =   3360
         TabIndex        =   252
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   198
         Left            =   2880
         TabIndex        =   251
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   197
         Left            =   2400
         TabIndex        =   250
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   196
         Left            =   1920
         TabIndex        =   249
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   195
         Left            =   1440
         TabIndex        =   248
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   194
         Left            =   960
         TabIndex        =   247
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   193
         Left            =   480
         TabIndex        =   246
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   192
         Left            =   0
         TabIndex        =   245
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   23
      Left            =   360
      TabIndex        =   235
      Top             =   8640
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   191
         Left            =   3360
         TabIndex        =   243
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   190
         Left            =   2880
         TabIndex        =   242
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   189
         Left            =   2400
         TabIndex        =   241
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   188
         Left            =   1920
         TabIndex        =   240
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   187
         Left            =   1440
         TabIndex        =   239
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   186
         Left            =   960
         TabIndex        =   238
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   185
         Left            =   480
         TabIndex        =   237
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   184
         Left            =   0
         TabIndex        =   236
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   21
      Left            =   360
      TabIndex        =   226
      Top             =   7920
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   175
         Left            =   3360
         TabIndex        =   234
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   174
         Left            =   2880
         TabIndex        =   233
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   173
         Left            =   2400
         TabIndex        =   232
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   172
         Left            =   1920
         TabIndex        =   231
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   171
         Left            =   1440
         TabIndex        =   230
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   170
         Left            =   960
         TabIndex        =   229
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   169
         Left            =   480
         TabIndex        =   228
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   168
         Left            =   0
         TabIndex        =   227
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   20
      Left            =   360
      TabIndex        =   217
      Top             =   7560
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   167
         Left            =   3360
         TabIndex        =   225
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   166
         Left            =   2880
         TabIndex        =   224
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   165
         Left            =   2400
         TabIndex        =   223
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   164
         Left            =   1920
         TabIndex        =   222
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   163
         Left            =   1440
         TabIndex        =   221
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   162
         Left            =   960
         TabIndex        =   220
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   161
         Left            =   480
         TabIndex        =   219
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   160
         Left            =   0
         TabIndex        =   218
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   19
      Left            =   360
      TabIndex        =   208
      Top             =   7200
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   159
         Left            =   3360
         TabIndex        =   216
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   158
         Left            =   2880
         TabIndex        =   215
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   157
         Left            =   2400
         TabIndex        =   214
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   156
         Left            =   1920
         TabIndex        =   213
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   155
         Left            =   1440
         TabIndex        =   212
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   154
         Left            =   960
         TabIndex        =   211
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   153
         Left            =   480
         TabIndex        =   210
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   152
         Left            =   0
         TabIndex        =   209
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   18
      Left            =   360
      TabIndex        =   199
      Top             =   6840
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   151
         Left            =   3360
         TabIndex        =   207
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   150
         Left            =   2880
         TabIndex        =   206
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   149
         Left            =   2400
         TabIndex        =   205
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   148
         Left            =   1920
         TabIndex        =   204
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   147
         Left            =   1440
         TabIndex        =   203
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   146
         Left            =   960
         TabIndex        =   202
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   145
         Left            =   480
         TabIndex        =   201
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   144
         Left            =   0
         TabIndex        =   200
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   17
      Left            =   360
      TabIndex        =   190
      Top             =   6480
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   143
         Left            =   3360
         TabIndex        =   198
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   142
         Left            =   2880
         TabIndex        =   197
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   141
         Left            =   2400
         TabIndex        =   196
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   140
         Left            =   1920
         TabIndex        =   195
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   139
         Left            =   1440
         TabIndex        =   194
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   138
         Left            =   960
         TabIndex        =   193
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   137
         Left            =   480
         TabIndex        =   192
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   136
         Left            =   0
         TabIndex        =   191
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   16
      Left            =   360
      TabIndex        =   181
      Top             =   6120
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   135
         Left            =   3360
         TabIndex        =   189
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   134
         Left            =   2880
         TabIndex        =   188
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   133
         Left            =   2400
         TabIndex        =   187
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   132
         Left            =   1920
         TabIndex        =   186
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   131
         Left            =   1440
         TabIndex        =   185
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   130
         Left            =   960
         TabIndex        =   184
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   129
         Left            =   480
         TabIndex        =   183
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   128
         Left            =   0
         TabIndex        =   182
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   15
      Left            =   360
      TabIndex        =   172
      Top             =   5760
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   127
         Left            =   3360
         TabIndex        =   180
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   126
         Left            =   2880
         TabIndex        =   179
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   125
         Left            =   2400
         TabIndex        =   178
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   124
         Left            =   1920
         TabIndex        =   177
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   123
         Left            =   1440
         TabIndex        =   176
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   122
         Left            =   960
         TabIndex        =   175
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   121
         Left            =   480
         TabIndex        =   174
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   120
         Left            =   0
         TabIndex        =   173
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   14
      Left            =   360
      TabIndex        =   163
      Top             =   5400
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   119
         Left            =   3360
         TabIndex        =   171
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   118
         Left            =   2880
         TabIndex        =   170
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   117
         Left            =   2400
         TabIndex        =   169
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   116
         Left            =   1920
         TabIndex        =   168
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   115
         Left            =   1440
         TabIndex        =   167
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   114
         Left            =   960
         TabIndex        =   166
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   113
         Left            =   480
         TabIndex        =   165
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   112
         Left            =   0
         TabIndex        =   164
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   13
      Left            =   360
      TabIndex        =   154
      Top             =   5040
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   111
         Left            =   3360
         TabIndex        =   162
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   110
         Left            =   2880
         TabIndex        =   161
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   109
         Left            =   2400
         TabIndex        =   160
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   108
         Left            =   1920
         TabIndex        =   159
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   107
         Left            =   1440
         TabIndex        =   158
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   106
         Left            =   960
         TabIndex        =   157
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   105
         Left            =   480
         TabIndex        =   156
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   104
         Left            =   0
         TabIndex        =   155
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   12
      Left            =   360
      TabIndex        =   145
      Top             =   4680
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   103
         Left            =   3360
         TabIndex        =   153
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   102
         Left            =   2880
         TabIndex        =   152
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   101
         Left            =   2400
         TabIndex        =   151
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   100
         Left            =   1920
         TabIndex        =   150
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   99
         Left            =   1440
         TabIndex        =   149
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   98
         Left            =   960
         TabIndex        =   148
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   97
         Left            =   480
         TabIndex        =   147
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   96
         Left            =   0
         TabIndex        =   146
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   10
      Left            =   360
      TabIndex        =   136
      Top             =   3960
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   87
         Left            =   3360
         TabIndex        =   144
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   86
         Left            =   2880
         TabIndex        =   143
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   85
         Left            =   2400
         TabIndex        =   142
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   84
         Left            =   1920
         TabIndex        =   141
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   83
         Left            =   1440
         TabIndex        =   140
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   82
         Left            =   960
         TabIndex        =   139
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   81
         Left            =   480
         TabIndex        =   138
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   80
         Left            =   0
         TabIndex        =   137
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   127
      Top             =   3600
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   79
         Left            =   3360
         TabIndex        =   135
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   78
         Left            =   2880
         TabIndex        =   134
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   77
         Left            =   2400
         TabIndex        =   133
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   76
         Left            =   1920
         TabIndex        =   132
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   75
         Left            =   1440
         TabIndex        =   131
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   74
         Left            =   960
         TabIndex        =   130
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   73
         Left            =   480
         TabIndex        =   129
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   72
         Left            =   0
         TabIndex        =   128
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   118
      Top             =   3240
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   71
         Left            =   3360
         TabIndex        =   126
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   70
         Left            =   2880
         TabIndex        =   125
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   69
         Left            =   2400
         TabIndex        =   124
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   68
         Left            =   1920
         TabIndex        =   123
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   67
         Left            =   1440
         TabIndex        =   122
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   66
         Left            =   960
         TabIndex        =   121
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   65
         Left            =   480
         TabIndex        =   120
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   64
         Left            =   0
         TabIndex        =   119
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   109
      Top             =   2880
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   63
         Left            =   3360
         TabIndex        =   117
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   62
         Left            =   2880
         TabIndex        =   116
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   61
         Left            =   2400
         TabIndex        =   115
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   60
         Left            =   1920
         TabIndex        =   114
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   59
         Left            =   1440
         TabIndex        =   113
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   58
         Left            =   960
         TabIndex        =   112
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   57
         Left            =   480
         TabIndex        =   111
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   56
         Left            =   0
         TabIndex        =   110
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   100
      Top             =   1440
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   31
         Left            =   3360
         TabIndex        =   108
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   30
         Left            =   2880
         TabIndex        =   107
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   29
         Left            =   2400
         TabIndex        =   106
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   28
         Left            =   1920
         TabIndex        =   105
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   27
         Left            =   1440
         TabIndex        =   104
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   26
         Left            =   960
         TabIndex        =   103
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   25
         Left            =   480
         TabIndex        =   102
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   24
         Left            =   0
         TabIndex        =   101
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   91
      Top             =   1080
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   23
         Left            =   3360
         TabIndex        =   99
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   22
         Left            =   2880
         TabIndex        =   98
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   21
         Left            =   2400
         TabIndex        =   97
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   20
         Left            =   1920
         TabIndex        =   96
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   19
         Left            =   1440
         TabIndex        =   95
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   18
         Left            =   960
         TabIndex        =   94
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   17
         Left            =   480
         TabIndex        =   93
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   92
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   82
      Top             =   720
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   15
         Left            =   3360
         TabIndex        =   90
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   14
         Left            =   2880
         TabIndex        =   89
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   13
         Left            =   2400
         TabIndex        =   88
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   12
         Left            =   1920
         TabIndex        =   87
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   11
         Left            =   1440
         TabIndex        =   86
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   85
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   84
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   83
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   35
      Left            =   4080
      TabIndex        =   81
      Top             =   12960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   34
      Left            =   4080
      TabIndex        =   80
      Top             =   12600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   33
      Left            =   4080
      TabIndex        =   79
      Top             =   12240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   32
      Left            =   4080
      TabIndex        =   78
      Top             =   11880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   31
      Left            =   4080
      TabIndex        =   77
      Top             =   11520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   30
      Left            =   4080
      TabIndex        =   76
      Top             =   11160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   29
      Left            =   4080
      TabIndex        =   75
      Top             =   10800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   28
      Left            =   4080
      TabIndex        =   74
      Top             =   10440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   27
      Left            =   4080
      TabIndex        =   73
      Top             =   10080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   26
      Left            =   4080
      TabIndex        =   72
      Top             =   9720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   25
      Left            =   4080
      TabIndex        =   71
      Top             =   9360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   24
      Left            =   4080
      TabIndex        =   70
      Top             =   9000
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   23
      Left            =   4080
      TabIndex        =   69
      Top             =   8640
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   22
      Left            =   4080
      TabIndex        =   68
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   21
      Left            =   4080
      TabIndex        =   67
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   20
      Left            =   4080
      TabIndex        =   66
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   19
      Left            =   4080
      TabIndex        =   65
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   18
      Left            =   4080
      TabIndex        =   64
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   17
      Left            =   4080
      TabIndex        =   63
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   16
      Left            =   4080
      TabIndex        =   62
      Top             =   6120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   15
      Left            =   4080
      TabIndex        =   61
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   14
      Left            =   4080
      TabIndex        =   60
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   13
      Left            =   4080
      TabIndex        =   59
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   12
      Left            =   4080
      TabIndex        =   58
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   11
      Left            =   4080
      TabIndex        =   57
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   10
      Left            =   4080
      TabIndex        =   56
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   9
      Left            =   4080
      TabIndex        =   55
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   8
      Left            =   4080
      TabIndex        =   54
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   7
      Left            =   4080
      TabIndex        =   53
      Top             =   2880
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   6
      Left            =   4080
      TabIndex        =   52
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   5
      Left            =   4080
      TabIndex        =   51
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   4
      Left            =   4080
      TabIndex        =   50
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   3
      Left            =   4080
      TabIndex        =   49
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   2
      Left            =   4080
      TabIndex        =   48
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   1
      Left            =   4080
      TabIndex        =   47
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Index           =   0
      Left            =   4080
      TabIndex        =   46
      Top             =   360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   7
         Left            =   3360
         TabIndex        =   8
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   6
         Left            =   2880
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   5
         Left            =   2400
         TabIndex        =   6
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   5
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   35
      Left            =   0
      TabIndex        =   45
      Top             =   13080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   34
      Left            =   0
      TabIndex        =   44
      Top             =   12720
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   33
      Left            =   0
      TabIndex        =   43
      Top             =   12360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   32
      Left            =   0
      TabIndex        =   42
      Top             =   12000
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   31
      Left            =   0
      TabIndex        =   41
      Top             =   11640
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   30
      Left            =   0
      TabIndex        =   40
      Top             =   11280
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   29
      Left            =   0
      TabIndex        =   39
      Top             =   10920
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   28
      Left            =   0
      TabIndex        =   38
      Top             =   10560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   27
      Left            =   0
      TabIndex        =   37
      Top             =   10200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   26
      Left            =   0
      TabIndex        =   36
      Top             =   9840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   25
      Left            =   0
      TabIndex        =   35
      Top             =   9480
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   24
      Left            =   0
      TabIndex        =   34
      Top             =   9120
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   23
      Left            =   0
      TabIndex        =   33
      Top             =   8760
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   22
      Left            =   0
      TabIndex        =   32
      Top             =   8400
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   21
      Left            =   0
      TabIndex        =   31
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   30
      Top             =   7680
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   19
      Left            =   0
      TabIndex        =   29
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   18
      Left            =   0
      TabIndex        =   28
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   17
      Left            =   0
      TabIndex        =   27
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   26
      Top             =   6240
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   15
      Left            =   0
      TabIndex        =   25
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   14
      Left            =   0
      TabIndex        =   24
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   13
      Left            =   0
      TabIndex        =   23
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   12
      Left            =   0
      TabIndex        =   22
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   21
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   20
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  MF      PC     Nuc     CF       gr      Gol       St     Bask     cell #"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu RecordFileMenu 
      Caption         =   "&File"
      Begin VB.Menu RasterMenu 
         Caption         =   "&Open Rasters"
         Index           =   1
      End
      Begin VB.Menu RasterMenu 
         Caption         =   "&Save Rasters"
         Index           =   2
      End
      Begin VB.Menu RasterMenu 
         Caption         =   "&Close Record Window"
         Index           =   3
      End
   End
   Begin VB.Menu RecordSelectAllMenu 
      Caption         =   "&Select All"
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Mossy Fibers"
         Index           =   1
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Purkinje Cells"
         Index           =   2
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Nucleus Cells"
         Index           =   3
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Climbing Fibers"
         Index           =   4
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All granule cells"
         Index           =   5
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Golgi Cells"
         Index           =   6
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Stellate Cells"
         Index           =   7
      End
      Begin VB.Menu SelectAllMenu 
         Caption         =   "All Basket Cells"
         Index           =   8
      End
   End
   Begin VB.Menu ClearAllMenu 
      Caption         =   "&Clear all Selections"
   End
   Begin VB.Menu RecordHistoScaleMenu 
      Caption         =   "&Histo Scale"
      Begin VB.Menu ScaleHistoMenu 
         Caption         =   "&Up"
         Index           =   1
      End
      Begin VB.Menu ScaleHistoMenu 
         Caption         =   "&Down"
         Index           =   2
      End
   End
   Begin VB.Menu UpdateMenu 
      Caption         =   "&Update"
   End
   Begin VB.Menu RecordExportDataMenu 
      Caption         =   "&Export Data"
      Visible         =   0   'False
      Begin VB.Menu ExportMenu 
         Caption         =   "Export Histo to Excel"
         Index           =   1
      End
      Begin VB.Menu ExportMenu 
         Caption         =   "Export Something else"
         Index           =   2
      End
   End
   Begin VB.Menu CSDurationMenu 
      Caption         =   "CS Duration"
      Visible         =   0   'False
   End
   Begin VB.Menu RecordFormColorsMenu 
      Caption         =   "Colors"
      Visible         =   0   'False
      Begin VB.Menu ColorsMenu 
         Caption         =   "Background color"
         Index           =   1
      End
      Begin VB.Menu ColorsMenu 
         Caption         =   "Raster Dot Color"
         Index           =   2
      End
      Begin VB.Menu ColorsMenu 
         Caption         =   "Histogram Color"
         Index           =   3
      End
      Begin VB.Menu ColorsMenu 
         Caption         =   "CS color"
         Index           =   4
      End
   End
   Begin VB.Menu RecordFormDotSizeMenu 
      Caption         =   "Dot Size"
      Visible         =   0   'False
      Begin VB.Menu DotSizeMenu 
         Caption         =   "1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu DotSizeMenu 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu DotSizeMenu 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu DotSizeMenu 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu DotSizeMenu 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu DotSizeMenu 
         Caption         =   "6"
         Index           =   6
      End
   End
   Begin VB.Menu RecordFormScaleMenu 
      Caption         =   "Scale"
      Visible         =   0   'False
      Begin VB.Menu ScaleMenu 
         Caption         =   "Scale Rows"
         Index           =   1
      End
   End
End
Attribute VB_Name = "RecordForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ClearAllMenu_Click()
Dim i As Integer

    For i = 0 To 287
        Option1(i).Value = False
    Next i
    For i = 0 To 35
        Text1(i).Text = ""
        RRCellType(i + 1) = 0
        RRCellNum(i + 1) = 0
    Next i
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

    
End Sub

Private Sub Command3_Click()
    RecordForm.Visible = False
End Sub

Private Sub Command4_Click()
   
End Sub

Private Sub Command5_Click(Index As Integer)

End Sub

Private Sub ColorsMenu_Click(Index As Integer)
    CommonDialog1.ShowColor
    Select Case Index
        Case 1 'background
            RecordFormBackColor = CommonDialog1.Color
        Case 2 'Dots
            RecordFormDotColor = CommonDialog1.Color
        Case 3  'Histo
            RecordFormHistoColor = CommonDialog1.Color
        Case 4  'CS indicator
            RecordFormCSColor = CommonDialog1.Color
    End Select
End Sub

Private Sub CSDurationButton_Click()
    RecordFormCSDurationFromMenu = Val(CSDurationText.Text)
    CSDurationText.Visible = False
    CSDurationButton.Visible = False
End Sub

Private Sub CSDurationMenu_Click()
    CSDurationButton.Visible = True
    CSDurationText.Visible = True
End Sub

Private Sub DotSizeMenu_Click(Index As Integer)
Dim i
    For i = 1 To 6
        DotSizeMenu(i).Checked = False
    Next i
    DotSizeMenu(Index).Checked = True
    RecordFormDotSize = Index
End Sub

Private Sub ExportMenu_Click(Index As Integer)
Dim i
    Select Case Index
        Case 1  'export histo to excel
            CSDurationText.LinkTopic = "excel|Sheet1"
            For i = 1 To 1000
                CSDurationText.Text = RecordFormHisto(i)
                CSDurationText.LinkItem = "R" & i & "C1"
                CSDurationText.LinkMode = vbLinkManual
                CSDurationText.LinkPoke
            Next i
    End Select
End Sub

Private Sub Form_Load()
Dim i As Integer

    For i = 0 To 35
        Label2(i).Caption = Str(i + 1)
    Next i
    
    If Do_Big_Rasters = 1 Then
        For i = 0 To 35
            Text1(i).Visible = False
            Frame1(i).Visible = False
            'Label2(i).Visible = False
        Next i
        Label1.Visible = False
        RecordForm.RecordExportDataMenu.Visible = True
        RecordForm.Width = cbm_main.SysInfo1.WorkAreaWidth
        RecordForm.CSDurationMenu.Visible = True
        RecordForm.RecordFormColorsMenu.Visible = True
        RecordForm.RecordFormDotSizeMenu.Visible = True
        RecordForm.RecordFormScaleMenu.Visible = True
        RecordForm.Frame2.Visible = True
        DrawRasters 1
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Debug.Print Button
End Sub

Private Sub Form_Resize()
    RecordForm.ScaleLeft = -50
    RecordForm.ScaleWidth = 1050
    RecordForm.ScaleTop = 1200
    RecordForm.ScaleHeight = -1200
    RecordForm.DrawWidth = 2
End Sub

Private Sub OpenCommand_Click()

End Sub

Private Sub Option2_Click(Index As Integer)
    DrawRasters Index + 1
End Sub
Private Sub DrawRasters(cell As Integer)
Dim x As Integer
Dim y As Integer
    Erase RecordFormHisto
    RecordForm.ScaleTop = RecordFormScaleRows
    RecordForm.ScaleHeight = -1 * RecordFormScaleRows
    RecordForm.DrawWidth = RecordFormDotSize
    If RecordFormCSDurationFromMenu = 0 Then
        RecordFormCSDuration = cs_duration(1)
    Else
        RecordFormCSDuration = RecordFormCSDurationFromMenu
    End If
     Select Case RRCellType(cell)
        Case 1  'MF
            RecordForm.Caption = "Mossy Fiber: "
        Case 2  'PC
            RecordForm.Caption = "Purkinje cell: "
        Case 3  'Nuc
            RecordForm.Caption = "Nucleus cell: "
        Case 4  'cf
            RecordForm.Caption = "Climbing fiber: "
        Case 5  'gr
            RecordForm.Caption = "granule cell: "
        Case 6  'Go
            RecordForm.Caption = "Golgi cell: "
        Case 7  'Stellate
            RecordForm.Caption = "Stellate cell: "
        Case 8  'Basket
            RecordForm.Caption = "Basket cell: "
    End Select
    RecordForm.Caption = RecordForm.Caption + Str(RRCellNum(cell))
    RecordForm.BackColor = RecordFormBackColor
    RecordForm.Cls
    RecordForm.Line (201, 0)-(200 + RecordFormCSDuration / 5, 1200), RecordFormCSColor, BF
    For y = 1 To 1000
        For x = 1 To 1000
            If Rasters(cell, x, y) = True Then
                RecordForm.PSet (x, y), RecordFormDotColor
                RecordFormHisto(x) = RecordFormHisto(x) + 1
            End If
        Next x
        DoEvents
    Next y
    DrawHistogram
End Sub
Private Sub DrawHistogram()
Dim x As Integer
    RecordForm.ScaleTop = 1200
    RecordForm.ScaleHeight = -1200
    For x = 1 To 1000
        RecordForm.Line (x - 1, 1001)-(x, 1001 + (RecordFormHisto(x) * RRScale)), RecordFormHistoColor, BF
    Next x
End Sub

Private Sub RasterMenu_Click(Index As Integer)
Dim filename As String
Dim i As Integer
    Select Case Index
        Case 1
            CommonDialog1.filename = ""
            CommonDialog1.ShowOpen
            If CommonDialog1.filename <> "" Then
                filename = CommonDialog1.filename
                Open filename For Binary As #2
                Get #2, , Rasters
                Close #2
                l = Len(filename)
                filename = Mid(filename, 1, l - 4) + "Type"
                Close #2
                Open filename For Binary As #2
                Get #2, , RRCellType
                Close #2
                filename = CommonDialog1.filename
                l = Len(filename)
                filename = Mid(filename, 1, l - 4) + "Num"
                Close #2
                Open filename For Binary As #2
                Get #2, , RRCellNum
                Close #2
                
                
                RecordForm.Width = cbm_main.SysInfo1.WorkAreaWidth
                RecordForm.RecordExportDataMenu.Visible = True
                RecordForm.CSDurationMenu.Visible = True
                RecordForm.RecordFormColorsMenu.Visible = True
                RecordForm.RecordFormDotSizeMenu.Visible = True
                RecordForm.RecordFormScaleMenu.Visible = True
                For i = 0 To 35
                    Text1(i).Visible = False
                    Frame1(i).Visible = False
                    Label2(i).Caption = i + 1
                Next i
                
                
                Label1.Visible = False
                Frame2.Visible = True
                RecordForm.ScaleLeft = -50
                RecordForm.ScaleWidth = 1050
                RecordForm.ScaleTop = 1200
                RecordForm.ScaleHeight = -1200
                RecordForm.DrawWidth = 2
                RecordForm.BackColor = vbBlack
                
            End If
        Case 2
            CommonDialog1.filename = ""
            CommonDialog1.ShowSave
            If CommonDialog1.filename <> "" Then
        
                filename = CommonDialog1.filename
                Debug.Print filename
        
                Close #2
                Open filename For Binary As #2
                Put #2, , Rasters
                Close #2
                l = Len(filename)
                filename = Mid(filename, 1, l - 4) + "Type"
                Close #2
                Open filename For Binary As #2
                Put #2, , RRCellType
                Close #2
                filename = CommonDialog1.filename
                l = Len(filename)
                filename = Mid(filename, 1, l - 4) + "Num"
                Close #2
                Open filename For Binary As #2
                Put #2, , RRCellNum
                Close #2
            End If
        
        Case 3
             RecordForm.Visible = False
        
    End Select
End Sub

Private Sub SaveCommand_Click()

End Sub

Private Sub ScaleButton_Click()
    RecordFormScaleRows = Val(ScaleText.Text)
    ScaleText.Visible = False
    ScaleButton.Visible = False
End Sub

Private Sub ScaleHistoMenu_Click(Index As Integer)
Dim i As Integer
Dim c As Integer
    Index = Index - 1
    For i = 0 To 35
        If Option2(i).Value = True Then c = i
    Next i
    If Index = 0 Then
        RRScale = RRScale * 1.1
    Else
        RRScale = RRScale * 0.9
    End If
    DrawRasters c
End Sub

Private Sub ScaleMenu_Click(Index As Integer)
    
    Select Case Index
        Case 1  ' scale rows
            ScaleText.Visible = True
            ScaleButton.Visible = True
    End Select
End Sub

Private Sub SelectAllMenu_Click(Index As Integer)
Dim i As Integer
Dim c As Integer
    Index = Index - 1
    Select Case Index
        Case 0  'MF
            For i = 0 To 280 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 1  'PC
            For i = 1 To 185 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 2  'Nuc
            For i = 2 To 58 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 3  'cf
            For i = 3 To 27 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 4  'gr
            For i = 4 To 287 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
         Case 5  'Go
            For i = 5 To 287 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 6  'Stellate
            For i = 6 To 287 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
        Case 7  'Basket
            For i = 7 To 287 Step 8
                Option1(i).Value = True
                Text1(c).Text = c + 1
                c = c + 1
            Next i
    End Select
End Sub

Private Sub UpdateMenu_Click()
Dim i, j As Integer
Dim temp As Integer
Dim c As Integer

    For i = 0 To 35
        If Text1(i).Text <> "" Then
            RRCellNum(i + 1) = Val(Text1(i))
            temp = i * 8
            c = 0
            For j = temp To temp + 7
                c = c + 1
                If Option1(j) Then
                    RRCellType(i + 1) = c
                End If
            Next j
        End If
        
    Next i
    RecordForm.Width = cbm_main.SysInfo1.WorkAreaWidth
    RecordForm.RecordExportDataMenu.Visible = True
    RecordForm.CSDurationMenu.Visible = True
    RecordForm.RecordFormColorsMenu.Visible = True
    RecordForm.RecordFormDotSizeMenu.Visible = True
    RecordForm.RecordFormScaleMenu.Visible = True
    RecordForm.RecordFormDotSizeMenu.Visible = True
    RecordForm.RecordFormScaleMenu.Visible = True
    For i = 0 To 35
        Text1(i).Visible = False
        Frame1(i).Visible = False
        Label2(i).Caption = i + 1
    Next i
   
'    Command1.Visible = False
'    Command2.Visible = False
'    Command3.Visible = False
    Label1.Visible = False
    Frame2.Visible = True
    RecordForm.ScaleLeft = -50
    RecordForm.ScaleWidth = 1050
    
    
    RecordForm.DrawWidth = 2
    RecordForm.BackColor = vbBlack
'    Command4.Visible = True
End Sub
