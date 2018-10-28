VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   7275
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8220
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   6852
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "colon.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Partition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String
Dim out_file As String
Dim atts As String
Dim sample As String
Dim class As String
Dim x As Integer
Dim y As Integer


Private Sub Partition_click()
    List1.Clear
    'check whether the file name is empty
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        'check whether the data file exists
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            'variale define
            Dim each_row_output() As String
            Dim data(63, 2001) As String
            Dim sum_1(2000) As Double 'sum of normal
            Dim avg_1(2000) As Double 'average of normal
            Dim sum_2(2000) As Double 'sum of tumor
            Dim avg_2(2000) As Double 'average of tumor
            Dim std_1(2000) As Double 'variance of normal
            Dim std_2(2000) As Double 'variance of tumor
            Dim t_value(2000) As Double 't value of 2000 gene
            Dim Index As String
            
            Open in_file For Input As #1
            
            'first two rows
            Line Input #1, sample
            Line Input #1, class
            
            y = 0 'set index of row from 0
            
            'use split function to seperate ',' in class and save to data(x,0)=column0
            For x = 0 To 62
                each_row_output = Split(class, ",")
                data(x, y) = each_row_output(x)
            Next

            y = 1 'index start from 1
            
            'read each line until end of file
            Do While Not EOF(1)
                Line Input #1, atts
                each_row_output = Split(atts, ",")
                For x = 0 To 62 'data rotate 90 degree, save from col2
                    data(x, y) = each_row_output(x)
                Next
                y = y + 1 'col2 -> col3
            Loop
            
            'calculate sum of class1(normal), class2(tumor)
            For y = 1 To 2000
                For x = 1 To 62
                    If CInt(data(x, 0)) = 1 Then 'CInt:transforms str to int
                        sum_1(y) = sum_1(y) + CDbl(data(x, y)) 'CDbl transforms double to str
                    ElseIf CInt(data(x, 0)) = 2 Then
                        sum_2(y) = sum_2(y) + CDbl(data(x, y))
                    End If
                Next
            Next
            
            'calculate mean
            For x = 1 To 2000
                avg_1(x) = sum_1(x) / 22 '22 is #normal gene per cell
                avg_2(x) = sum_2(x) / 40 '40 is #tumor gene per cell
            Next
            
            'variance
            For y = 1 To 2000
                std_1(y) = 0
                std_2(y) = 0
                For x = 1 To 62
                    If CInt(data(x, 0)) = 1 Then
                        std_1(y) = std_1(y) + (CDbl(data(x, y)) - avg_1(y)) ^ 2
                    ElseIf CInt(data(x, 0)) = 2 Then
                        std_2(y) = std_2(y) + (CDbl(data(x, y)) - avg_2(y)) ^ 2
                    End If
                Next
            Next
            
            For x = 1 To 2000
                std_1(x) = std_1(x) / (22 - 1)
                std_2(x) = std_2(x) / (40 - 1)
            Next
            
            
            't value
            For x = 1 To 2000
                t_value(x) = (avg_1(x) - avg_2(x)) / (std_1(x) / 22 + std_2(x) / 40) ^ 0.5
            Next
                        
            'bubble sort
            For x = 1 To 2000
                For y = x To 2000
                    If t_value(x) > t_value(y) Then
                        'change t value
                        Tag = t_value(x)
                        t_value(x) = t_value(y)
                        t_value(y) = Tag
                        'change index of t value
                        Index = data(0, x)
                        data(0, x) = data(0, y)
                        data(0, y) = Index
                    End If
                Next y
            Next x
            
            'output to list
            For x = 1 To 2000
                List1.AddItem "Gene seq" & data(0, x) & Chr(9) & "t = " & t_value(x)
            Next x
               
            Close #1
        End If
    End If
End Sub


