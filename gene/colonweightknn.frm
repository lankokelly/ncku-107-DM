VERSION 5.00
Begin VB.Form Partition 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   12870
   Begin VB.CommandButton Command2 
      Caption         =   "t-value"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "score"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton tbtn 
      Caption         =   "t-value"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9960
      TabIndex        =   6
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton scorebtn 
      Caption         =   "score"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5760
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   10815
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Text            =   "colon.txt"
      Top             =   360
      Width           =   5895
   End
   Begin VB.CommandButton Partition 
      Appearance      =   0  '平面
      BackColor       =   &H00400000&
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8520
      MaskColor       =   &H00400000&
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "sort from large to small"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label small 
      BackStyle       =   0  '透明
      Caption         =   "sort from small to large"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000001&
      BackStyle       =   0  '透明
      Caption         =   "leave-one-out cross validation"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   2535
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
Dim k As Integer
Dim data(63, 2001) As String
Dim data2(63, 2001) As String
Dim score(2000, 2) As Integer 'col=0存score col=1存index

Private Sub Partition_click()
    List1.Clear
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            'variale define
            Dim each_row_output() As String
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
                data2(x, y) = each_row_output(x)
            Next
            y = 1 'index start from 1
            
            'read each line until end of file
            Do While Not EOF(1)
                Line Input #1, atts
                each_row_output = Split(atts, ",")
                For x = 0 To 62 'data rotate 90 degree, save from col2
                    data(x, y) = each_row_output(x)
                    data2(x, y) = each_row_output(x)
                Next
                y = y + 1 'col2 -> col3
            Loop
            
'tvalue-------------------------------------------------------------------------------------------------
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
                'List1.AddItem "Gene seq" & data(0, x) & Chr(9) & "t = " & t_value(x)
            Next x
            
'score-------------------------------------------------------------------------------------------------
            'count score
            For x = 1 To 2000
                score(x - 1, 1) = x 'index
                For y = 1 To 62
                    If CDbl(data2(y, 0)) = 1 Then
                        For k = 1 To 62
                            If CDbl(data2(k, 0)) = 2 Then
                                If CDbl(data2(y, x)) > CDbl(data2(k, x)) Then
                                    score(x - 1, 0) = score(x - 1, 0) + 1
                                End If
                            End If
                        Next
                    End If
                Next
            Next
            
            'bubble sort
            For x = 0 To 1999
                For y = 0 To 1999
                    If score(x, 0) > score(y, 0) Then
                        'change score
                        Tag = score(x, 0)
                        score(x, 0) = score(y, 0)
                        score(y, 0) = Tag
                        'change index of score
                        Index = score(x, 1)
                        score(x, 1) = score(y, 1)
                        score(y, 1) = Index
                    End If
                Next y
            Next x
            List1.AddItem "successfully read"
            Close #1
        End If
    End If
End Sub


Private Sub scorebtn_Click()
List1.Clear
'--score排序-----------------------------------------------------------------------
            Dim disdis As Double
            Dim sumdis(63, 63) As Double
            Dim sumdis2(63, 63) As Double
            Dim sumdis3(63, 63) As Double
            Dim sumdis4(63, 63) As Double
            Dim a As Integer
            Dim b As Integer
            Dim min_1(63) As Double
            Dim min_2(63) As Double
            Dim min_3(63) As Double
            Dim min_4(63) As Double
            Dim min_5(63) As Double
            Dim min_6(63) As Double
            Dim class_1(63) As String
            Dim class_2(63) As String
            Dim class_3(63) As String
            Dim class_4(63) As String
            Dim class_5(63) As String
            Dim class_6(63) As String
            Dim classofcell(63, 63) As String
            
            '存距離平方(不開根號)
            For x = 1 To 62
                For y = 1 To 62 '想像成另一張表的x
                    For k = 1 To 50 'scoreindex' '選前50小啦
                        a = score(k - 1, 1) '每個前50小的真正在data2的index
                        sumdis(x, y) = sumdis(x, y) + (CDbl(data2(x, a)) - CDbl(data2(y, a))) ^ 2
                    Next
                    For k = 1 To 100 'scoreindex' '選前100小啦
                        a = score(k - 1, 1)
                        sumdis2(x, y) = sumdis2(x, y) + (CDbl(data2(x, a)) - CDbl(data2(y, a))) ^ 2
                    Next
                    For k = 1 To 150 'scoreindex' '選前150小啦
                        a = score(k - 1, 1)
                        sumdis3(x, y) = sumdis3(x, y) + (CDbl(data2(x, a)) - CDbl(data2(y, a))) ^ 2
                    Next
                    For k = 1 To 200 'scoreindex' '選前200小啦
                        a = score(k - 1, 1)
                        sumdis4(x, y) = sumdis4(x, y) + (CDbl(data2(x, a)) - CDbl(data2(y, a))) ^ 2
                    Next
                    classofcell(x, y) = data2(y, 0)
                Next
            Next
'------------------------------------------------------------------------
            '比前3小距離
            For x = 1 To 62
                class_1(x) = "10000"
                class_2(x) = "10000"
                class_3(x) = "10000"
                class_4(x) = "10000"
                class_5(x) = "10000"
                class_6(x) = "10000"
                min_1(x) = 10000
                min_2(x) = 10000
                min_3(x) = 10000
                min_4(x) = 10000
                min_5(x) = 10000
                min_6(x) = 10000
                For y = 1 To 62
                    If x <> y Then
                        If sumdis(x, y) <= min_1(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = min_1(x)
                            min_1(x) = sumdis(x, y)
                            class_3(x) = class_2(x)
                            class_2(x) = class_1(x)
                            class_1(x) = classofcell(x, y)
                        ElseIf sumdis(x, y) >= min_1(x) And sumdis(x, y) <= min_2(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = sumdis(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = class_2(x)
                            class_2(x) = classofcell(x, y)
                        ElseIf sumdis(x, y) >= min_2(x) And sumdis(x, y) <= min_3(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = sumdis(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = classofcell(x, y)
                        ElseIf sumdis(x, y) >= min_3(x) And sumdis(x, y) <= min_4(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = sumdis(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = classofcell(x, y)
                        ElseIf sumdis(x, y) >= min_4(x) And sumdis(x, y) <= min_5(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = sumdis(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = classofcell(x, y)
                        ElseIf sumdis(x, y) >= min_5(x) And sumdis(x, y) <= min_6(x) Then
                            min_6(x) = sumdis(x, y)
                            class_6(x) = classofcell(x, y)
                        End If
                    End If
                Next
            Next
            
            
            '比前3小每個class的weight
            Dim sum1weight(63) As Double
            Dim sum2weight(63) As Double
            Dim clssspred3(63) As String
            For x = 1 To 62
                If class_1(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_1(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_1(x)
                End If
                If class_2(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_2(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_2(x)
                End If
                If class_3(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_3(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_3(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred3(x) = 1
                Else: clssspred3(x) = 2
                End If
            Next
            
            '算accuracy
            Dim acc As Double
            For x = 1 To 62
                If classofcell(x, x) = clssspred3(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "chosen number = 50"
            List1.AddItem "when k = 3," & "accuracy = " & acc / 62
            
            '比前4小每個class的weight
            Dim clssspred4(63) As String
            For x = 1 To 62
                If class_4(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_4(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_4(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred4(x) = 1
                Else: clssspred4(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred4(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 4," & "accuracy = " & acc / 62
            
            '比前5小每個class的weight
            Dim clssspred5(63) As String
            For x = 1 To 62
                If class_5(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_5(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_5(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred5(x) = 1
                Else: clssspred5(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred5(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 5," & "accuracy = " & acc / 62
            
             '比前6小每個class的weight
            Dim clssspred6(63) As String
            For x = 1 To 62
                If class_6(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_6(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_6(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred6(x) = 1
                Else: clssspred6(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred6(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 6," & "accuracy = " & acc / 62
'------------------------------------------------------------------------
            '100'比前3小距離
            For x = 1 To 62
                class_1(x) = "10000"
                class_2(x) = "10000"
                class_3(x) = "10000"
                class_4(x) = "10000"
                class_5(x) = "10000"
                class_6(x) = "10000"
                min_1(x) = 10000
                min_2(x) = 10000
                min_3(x) = 10000
                min_4(x) = 10000
                min_5(x) = 10000
                min_6(x) = 10000
                For y = 1 To 62
                    If x <> y Then
                        If sumdis2(x, y) <= min_1(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = min_1(x)
                            min_1(x) = sumdis2(x, y)
                            class_3(x) = class_2(x)
                            class_2(x) = class_1(x)
                            class_1(x) = classofcell(x, y)
                        ElseIf sumdis2(x, y) >= min_1(x) And sumdis2(x, y) <= min_2(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = sumdis2(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = class_2(x)
                            class_2(x) = classofcell(x, y)
                        ElseIf sumdis2(x, y) >= min_2(x) And sumdis2(x, y) <= min_3(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = sumdis2(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = classofcell(x, y)
                        ElseIf sumdis2(x, y) >= min_3(x) And sumdis2(x, y) <= min_4(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = sumdis2(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = classofcell(x, y)
                        ElseIf sumdis2(x, y) >= min_4(x) And sumdis2(x, y) <= min_5(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = sumdis2(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = classofcell(x, y)
                        ElseIf sumdis2(x, y) >= min_5(x) And sumdis2(x, y) <= min_6(x) Then
                            min_6(x) = sumdis2(x, y)
                            class_6(x) = classofcell(x, y)
                        End If
                    End If
                Next
            Next
            
            '比前3小每個class的weight
            For x = 1 To 62
                sum1weight(x) = 0
                sum2weight(x) = 0
                If class_1(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_1(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_1(x)
                End If
                If class_2(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_2(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_2(x)
                End If
                If class_3(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_3(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_3(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred3(x) = 1
                Else: clssspred3(x) = 2
                End If
            Next
            
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred3(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "chosen number = 100"
            List1.AddItem "when k = 3," & "accuracy = " & acc / 62
            
            '比前4小每個class的weight
            For x = 1 To 62
                If class_4(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_4(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_4(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred4(x) = 1
                Else: clssspred4(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred4(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 4," & "accuracy = " & acc / 62
            
            '比前5小每個class的weight
            For x = 1 To 62
                If class_5(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_5(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_5(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred5(x) = 1
                Else: clssspred5(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred5(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 5," & "accuracy = " & acc / 62
            
             '比前6小每個class的weight
            For x = 1 To 62
                If class_6(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_6(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_6(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred6(x) = 1
                Else: clssspred6(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred6(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 6," & "accuracy = " & acc / 62
'------------------------------------------------------------------------
        '150'比前3小距離
            For x = 1 To 62
                class_1(x) = "10000"
                class_2(x) = "10000"
                class_3(x) = "10000"
                class_4(x) = "10000"
                class_5(x) = "10000"
                class_6(x) = "10000"
                min_1(x) = 10000
                min_2(x) = 10000
                min_3(x) = 10000
                min_4(x) = 10000
                min_5(x) = 10000
                min_6(x) = 10000
                For y = 1 To 62
                    If x <> y Then
                        If sumdis3(x, y) <= min_1(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = min_1(x)
                            min_1(x) = sumdis3(x, y)
                            class_3(x) = class_2(x)
                            class_2(x) = class_1(x)
                            class_1(x) = classofcell(x, y)
                        ElseIf sumdis3(x, y) >= min_1(x) And sumdis3(x, y) <= min_2(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = sumdis3(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = class_2(x)
                            class_2(x) = classofcell(x, y)
                        ElseIf sumdis3(x, y) >= min_2(x) And sumdis3(x, y) <= min_3(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = sumdis3(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = classofcell(x, y)
                        ElseIf sumdis3(x, y) >= min_3(x) And sumdis3(x, y) <= min_4(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = sumdis3(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = classofcell(x, y)
                        ElseIf sumdis3(x, y) >= min_4(x) And sumdis3(x, y) <= min_5(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = sumdis3(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = classofcell(x, y)
                        ElseIf sumdis3(x, y) >= min_5(x) And sumdis3(x, y) <= min_6(x) Then
                            min_6(x) = sumdis3(x, y)
                            class_6(x) = classofcell(x, y)
                        End If
                    End If
                Next
            Next
            
            '比前3小每個class的weight
            For x = 1 To 62
                sum1weight(x) = 0
                sum2weight(x) = 0
                If class_1(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_1(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_1(x)
                End If
                If class_2(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_2(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_2(x)
                End If
                If class_3(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_3(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_3(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred3(x) = 1
                Else: clssspred3(x) = 2
                End If
            Next
            
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred3(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "chosen number = 150"
            List1.AddItem "when k = 3," & "accuracy = " & acc / 62
            
            '比前4小每個class的weight
            For x = 1 To 62
                If class_4(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_4(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_4(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred4(x) = 1
                Else: clssspred4(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred4(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 4," & "accuracy = " & acc / 62
            
            '比前5小每個class的weight
            For x = 1 To 62
                If class_5(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_5(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_5(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred5(x) = 1
                Else: clssspred5(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred5(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 5," & "accuracy = " & acc / 62
            
             '比前6小每個class的weight
            For x = 1 To 62
                If class_6(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_6(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_6(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred6(x) = 1
                Else: clssspred6(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred6(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 6," & "accuracy = " & acc / 62
'------------------------------------------------------------------------
'100'比前3小距離
            For x = 1 To 62
                class_1(x) = "10000"
                class_2(x) = "10000"
                class_3(x) = "10000"
                class_4(x) = "10000"
                class_5(x) = "10000"
                class_6(x) = "10000"
                min_1(x) = 10000
                min_2(x) = 10000
                min_3(x) = 10000
                min_4(x) = 10000
                min_5(x) = 10000
                min_6(x) = 10000
                For y = 1 To 62
                    If x <> y Then
                        If sumdis4(x, y) <= min_1(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = min_1(x)
                            min_1(x) = sumdis4(x, y)
                            class_3(x) = class_2(x)
                            class_2(x) = class_1(x)
                            class_1(x) = classofcell(x, y)
                        ElseIf sumdis4(x, y) >= min_1(x) And sumdis4(x, y) <= min_2(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = min_2(x)
                            min_2(x) = sumdis4(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = class_2(x)
                            class_2(x) = classofcell(x, y)
                        ElseIf sumdis4(x, y) >= min_2(x) And sumdis4(x, y) <= min_3(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = min_3(x)
                            min_3(x) = sumdis4(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = class_3(x)
                            class_3(x) = classofcell(x, y)
                        ElseIf sumdis4(x, y) >= min_3(x) And sumdis4(x, y) <= min_4(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = min_4(x)
                            min_4(x) = sumdis4(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = class_4(x)
                            class_4(x) = classofcell(x, y)
                        ElseIf sumdis4(x, y) >= min_4(x) And sumdis4(x, y) <= min_5(x) Then
                            min_6(x) = min_5(x)
                            min_5(x) = sumdis4(x, y)
                            class_6(x) = class_5(x)
                            class_5(x) = classofcell(x, y)
                        ElseIf sumdis4(x, y) >= min_5(x) And sumdis4(x, y) <= min_6(x) Then
                            min_6(x) = sumdis4(x, y)
                            class_6(x) = classofcell(x, y)
                        End If
                    End If
                Next
            Next
            
            '比前3小每個class的weight
            For x = 1 To 62
                sum1weight(x) = 0
                sum2weight(x) = 0
                If class_1(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_1(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_1(x)
                End If
                If class_2(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_2(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_2(x)
                End If
                If class_3(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_3(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_3(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred3(x) = 1
                Else: clssspred3(x) = 2
                End If
            Next
            
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred3(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "chosen number = 200"
            List1.AddItem "when k = 3," & "accuracy = " & acc / 62
            
            '比前4小每個class的weight
            For x = 1 To 62
                If class_4(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_4(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_4(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred4(x) = 1
                Else: clssspred4(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred4(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 4," & "accuracy = " & acc / 62
            
            '比前5小每個class的weight
            For x = 1 To 62
                If class_5(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_5(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_5(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred5(x) = 1
                Else: clssspred5(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred5(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 5," & "accuracy = " & acc / 62
            
             '比前6小每個class的weight
            For x = 1 To 62
                If class_6(x) = "1" Then
                    sum1weight(x) = sum1weight(x) + 1 / min_6(x)
                Else: sum2weight(x) = sum2weight(x) + 1 / min_6(x)
                End If
                If sum1weight(x) > sum2weight(x) Then
                    clssspred6(x) = 1
                Else: clssspred6(x) = 2
                End If
            Next
            '算accuracy
            acc = 0
            For x = 1 To 62
                If classofcell(x, x) = clssspred6(x) Then
                acc = acc + 1
                End If
            Next
            List1.AddItem "when k = 6," & "accuracy = " & acc / 62
End Sub

