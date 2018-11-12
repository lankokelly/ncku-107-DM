VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13455
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   13455
   Begin VB.CommandButton Command5 
      Caption         =   "Randomize data and cut into 5-fold"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   6360
      TabIndex        =   13
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "k = 6"
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "k = 5"
      Height          =   615
      Left            =   2880
      TabIndex        =   11
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "k = 4"
      Height          =   615
      Left            =   1800
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "k = 3"
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton backward 
      Caption         =   "Backward"
      Height          =   615
      Left            =   7920
      TabIndex        =   6
      Top             =   2520
      Width           =   1680
   End
   Begin VB.CommandButton forward 
      Caption         =   "Forward"
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   2520
      Width           =   1680
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
      Height          =   4905
      Left            =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   12255
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
      Height          =   405
      Left            =   2400
      TabIndex        =   1
      Text            =   "yeast.txt"
      Top             =   650
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
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
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "# of k-nearest neighbors : "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   1920
      Y2              =   3360
   End
   Begin VB.Label Label3 
      Caption         =   "Feature Selection : "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Output"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
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
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   1815
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
Dim X As Integer
Dim Y As Integer
Dim col As Integer
Dim RanNum As Integer
Dim num_in_fold(6) As Integer
Dim data(1484, 10) As String  'data1 used in equal-width,changing value after discretization
Dim distance(296, 1188) As Double 'fold1的(0-1187筆data ,0-1187存跟別人的距離)
Dim class_predict_candidate(296, 1188) As String '對照distance()，存根某training data的class
Dim distance_2(297, 1187) As Double 'fold2的(0-1187筆data ,0-1187存跟別人的距離)
Dim class_predict_candidate_2(297, 1187) As String '對照distance()，存根某training data的class
Dim distance_3(297, 1187) As Double 'fold3的(0-1187筆data ,0-1187存跟別人的距離)
Dim class_predict_candidate_3(297, 1187) As String '對照distance()，存根某training data的class
Dim distance_4(297, 1187) As Double 'fold4的(0-1187筆data ,0-1187存跟別人的距離)
Dim class_predict_candidate_4(297, 1187) As String '對照distance()，存根某training data的class
Dim distance_5(297, 1187) As Double 'fold5的(0-1187筆data ,0-1187存跟別人的距離)
Dim class_predict_candidate_5(297, 1187) As String '對照distance()，存根某training data的class
Dim fold_1(296, 10) As String
Dim fold_2(297, 10) As String
Dim fold_3(297, 10) As String
Dim fold_4(297, 10) As String
Dim fold_5(297, 10) As String

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label4_Click()

End Sub

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
 
            Open in_file For Input As #1
            Open App.Path & "\test.txt" For Output As #2
            'Open App.Path & "\test-freq.txt" For Output As #3
            
            Y = 0 'from row 0
            
            'read each line until end of file
            Do While Not EOF(1)
                Line Input #1, atts
                each_row_output = Split(atts, " ")
                col = 0
                For X = 0 To UBound(each_row_output)
                    If each_row_output(X) <> "" Then 'if space,ignore
                        data(Y, col) = each_row_output(X)
                    col = col + 1
                    End If
                Next
                Y = Y + 1 'goto next row
            Loop

            Close #1
        End If
    End If
End Sub
Private Sub Command5_click()
    'cut 5-fold
    Dim RanNum(1485) As Integer
    X = 1
    Do While X <= 1484
        RanNum(X) = Int(Rnd() * 1484) + 1 '隨機1-1484之間的整數
        Dim G As Boolean '判斷有無重複
        G = True '預設沒有重複

        For Y = 0 To X - 1  '開始判斷有沒有重複
            If RanNum(X) = RanNum(Y) Then
                G = False   '有重覆設定為重新選取
                Y = X   '跳出迴圈
            End If
        Next

        If G = True Then
            X = X + 1   '沒有重複則繼續取下一個數
        End If
    Loop
    
    '把原始資料丟到5個fold矩陣
    Dim temp_1 As Integer 'fold_1()的row，預設為0
    Dim temp_2 As Integer
    Dim temp_3 As Integer
    Dim temp_4 As Integer
    Dim temp_5 As Integer
    
    For X = 1 To 1484
        If RanNum(X) > 0 And RanNum(X) <= 296 Then '如果隨機的數字對應到的原data在此區間
            For col = 0 To 9
                fold_1(temp_1, col) = data(X - 1, col) 'data
            Next
            temp_1 = temp_1 + 1 'row++
        ElseIf RanNum(X) > 296 And RanNum(X) <= 593 Then
            For col = 0 To 9
                fold_2(temp_2, col) = data(X - 1, col)
            Next
            temp_2 = temp_2 + 1 '到下一個row
        ElseIf RanNum(X) > 593 And RanNum(X) <= 890 Then
            For col = 0 To 9
                fold_3(temp_3, col) = data(X - 1, col)
            Next
            temp_3 = temp_3 + 1 '到下一個row
        ElseIf RanNum(X) > 890 And RanNum(X) <= 1187 Then
            For col = 0 To 9
                fold_4(temp_4, col) = data(X - 1, col)
            Next
            temp_4 = temp_4 + 1 '到下一個row
        ElseIf RanNum(X) > 1187 And RanNum(X) <= 1484 Then
            For col = 0 To 9
                fold_5(temp_5, col) = data(X - 1, col)
            Next
            temp_5 = temp_5 + 1 '到下一個row
        End If
    Next
    
    
    'iteration1:------------------------------------------------------------------------------------
    'fold-1當testing，fold-2 到 fold-5當training
    '先將fold2到fold5的資料合併7
    Dim iteration_1_training(1188, 10) As String
    For X = 0 To 296
        For Y = 0 To 9
            iteration_1_training(X, Y) = fold_2(X, Y)
            iteration_1_training(X + 297, Y) = fold_3(X, Y)
            iteration_1_training(X + 594, Y) = fold_4(X, Y)
            iteration_1_training(X + 891, Y) = fold_5(X, Y)
        Next
    Next
    
    '算fold-1與iteration_1_training之歐式距離
    Dim euclidean As Double
    Dim tempclass As String
    
    For X = 0 To 295
        For Y = 0 To 1187
            For col = 1 To 8
                euclidean = euclidean + (CDbl(fold_1(X, col)) - CDbl(iteration_1_training(Y, col))) ^ 2
            Next col
            euclidean = euclidean ^ 0.5
            distance(X, Y) = euclidean
            tempclass = iteration_1_training(Y, 9)
            class_predict_candidate(X, Y) = tempclass
            euclidean = 0
        Next
    Next
    '--------------------------------------------------------------------------------------
    
    'iteration2:------------------------------------------------------------------------------------
    'fold-2當testing，其餘當training
    '先將fold資料合併
    Dim iteration_2_training(1187, 10) As String
    For X = 0 To 295
        For Y = 0 To 9
            iteration_2_training(X, Y) = fold_1(X, Y)
        Next
    Next
    For X = 0 To 296
        For Y = 0 To 9
            iteration_2_training(X + 296, Y) = fold_3(X, Y)
            iteration_2_training(X + 593, Y) = fold_4(X, Y)
            iteration_2_training(X + 890, Y) = fold_5(X, Y)
        Next
    Next
    
    '算fold-2與iteration_2_training之歐式距離
    euclidean = 0
    tempclass = ""
    
    For X = 0 To 296
        For Y = 0 To 1186
            For col = 1 To 8
                euclidean = euclidean + (CDbl(fold_2(X, col)) - CDbl(iteration_2_training(Y, col))) ^ 2
            Next col
            euclidean = euclidean ^ 0.5
            distance_2(X, Y) = euclidean
            tempclass = iteration_2_training(Y, 9)
            class_predict_candidate_2(X, Y) = tempclass
            euclidean = 0
        Next
    Next
    
    'iteration3:------------------------------------------------------------------------------------
    'fold-3當testing，其餘當training
    '先將fold資料合併
    Dim iteration_3_training(1187, 10) As String
    For X = 0 To 295
        For Y = 0 To 9
            iteration_3_training(X, Y) = fold_1(X, Y)
        Next
    Next
    For X = 0 To 296
        For Y = 0 To 9
            iteration_3_training(X + 296, Y) = fold_2(X, Y)
            iteration_3_training(X + 593, Y) = fold_4(X, Y)
            iteration_3_training(X + 890, Y) = fold_5(X, Y)
        Next
    Next
    
    '算fold-3與iteration_3_training之歐式距離
    euclidean = 0
    tempclass = ""
    
    For X = 0 To 296
        For Y = 0 To 1186
            For col = 1 To 8
                euclidean = euclidean + (CDbl(fold_3(X, col)) - CDbl(iteration_3_training(Y, col))) ^ 2
            Next col
            euclidean = euclidean ^ 0.5
            distance_3(X, Y) = euclidean
            tempclass = iteration_3_training(Y, 9)
            class_predict_candidate_3(X, Y) = tempclass
            euclidean = 0
        Next
    Next
   
    
    'iteration4:------------------------------------------------------------------------------------
    'fold-4當testing，其餘當training
    '先將fold資料合併
    Dim iteration_4_training(1187, 10) As String
    For X = 0 To 295
        For Y = 0 To 9
            iteration_4_training(X, Y) = fold_1(X, Y)
        Next
    Next
    For X = 0 To 296
        For Y = 0 To 9
            iteration_4_training(X + 296, Y) = fold_2(X, Y)
            iteration_4_training(X + 593, Y) = fold_3(X, Y)
            iteration_4_training(X + 890, Y) = fold_5(X, Y)
        Next
    Next
    
    '算fold-4與iteration_4_training之歐式距離
    euclidean = 0
    tempclass = ""
    
    For X = 0 To 296
        For Y = 0 To 1186
            For col = 1 To 8
                euclidean = euclidean + (CDbl(fold_4(X, col)) - CDbl(iteration_4_training(Y, col))) ^ 2
            Next col
            euclidean = euclidean ^ 0.5
            distance_4(X, Y) = euclidean
            tempclass = iteration_4_training(Y, 9)
            class_predict_candidate_4(X, Y) = tempclass
            euclidean = 0
        Next
    Next
    '--------------------------------------------------------------------------------------
    'iteration5:------------------------------------------------------------------------------------
    'fold-5當testing，其餘當training
    '先將fold資料合併
    Dim iteration_5_training(1187, 10) As String
    For X = 0 To 295
        For Y = 0 To 9
            iteration_5_training(X, Y) = fold_1(X, Y)
        Next
    Next
    For X = 0 To 296
        For Y = 0 To 9
            iteration_5_training(X + 296, Y) = fold_2(X, Y)
            iteration_5_training(X + 593, Y) = fold_3(X, Y)
            iteration_5_training(X + 890, Y) = fold_4(X, Y)
        Next
    Next
    
    '算fold-5與iteration_5_training之歐式距離
    euclidean = 0
    tempclass = ""
    
    For X = 0 To 296
        For Y = 0 To 1186
            For col = 1 To 8
                euclidean = euclidean + (CDbl(fold_5(X, col)) - CDbl(iteration_5_training(Y, col))) ^ 2
            Next col
            euclidean = euclidean ^ 0.5
            distance_5(X, Y) = euclidean
            tempclass = iteration_5_training(Y, 9)
            class_predict_candidate_5(X, Y) = tempclass
            euclidean = 0
        Next
    Next
    '--------------------------------------------------------------------------------------
    'For X = 0 To 1187
       'Print #2, class_predict_candidate(186, X)
    'Next
End Sub

Private Sub Command1_click() '紀錄k=3
    List1.Clear
    Dim mode1 As Integer
    Dim mode2 As Integer
    Dim mode3 As Integer
    Dim mode4 As Integer
    Dim mode5 As Integer
    Dim min_1(1188) As Double
    Dim min_2(1188) As Double
    Dim min_3(1188) As Double
    Dim class_1(1188) As String
    Dim class_2(1188) As String
    Dim class_3(1188) As String
    Dim accuracy(5) As Double
    
    'fold 1-----------------------------------------------
    For X = 0 To 295
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        For Y = 0 To 1187
            If distance(X, Y) <= min_1(X) Then
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_1(X) And distance(X, Y) <= min_2(X) And distance(X, Y) <= min_3(X) Then
                min_3(X) = min_2(X)
                min_2(X) = distance(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_2(X) And distance(X, Y) <= min_3(X) Then
                min_3(X) = distance(X, Y)
                class_3(X) = class_predict_candidate(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold1(296) As String
    Dim rnd_class As Integer
    Dim ctr_accuracy As Integer
    For X = 0 To 295
        If class_1(X) = class_2(X) Then
            class_predict_fold1(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) Then
            class_predict_fold1(X) = class_2(X)
        ElseIf class_3(X) = class_1(X) Then
            class_predict_fold1(X) = class_1(X)
        Else
            mode1 = X Mod 3
            Select Case mode1
                Case 1
                    class_predict_fold1(X) = class_1(X)
                Case 2
                    class_predict_fold1(X) = class_2(X)
                Case Else
                    class_predict_fold1(X) = class_3(X)
            End Select
        End If
    Next
    
    For X = 0 To 295
        If class_predict_fold1(X) = fold_1(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(0) = (ctr_accuracy / 296) * 100
    
    
    List1.AddItem "i_th fold" & vbTab & "#data" & vbTab & "#accurate data" & vbTab & "  accuracy"
    List1.AddItem "1st  fold" & vbTab & "296" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(0), 6) & "%"
    
    
    'fold 2-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        For Y = 0 To 1186
            If distance_2(X, Y) <= min_1(X) Then
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_2(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_1(X) And distance_2(X, Y) <= min_2(X) And distance_2(X, Y) <= min_3(X) Then
                min_3(X) = min_2(X)
                min_2(X) = distance_2(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_2(X) And distance_2(X, Y) <= min_3(X) Then
                min_3(X) = distance_2(X, Y)
                class_3(X) = class_predict_candidate_2(X, Y)
            End If
        Next Y
    Next X
    
    
    Dim class_predict_fold2(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) Then
            class_predict_fold2(X) = class_1(X)
            
        ElseIf class_2(X) = class_3(X) Then
            class_predict_fold2(X) = class_2(X)
            
        ElseIf class_3(X) = class_1(X) Then
            class_predict_fold2(X) = class_1(X)
            
        Else
            mode2 = X Mod 3
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_1(X)
                Case 2
                    class_predict_fold2(X) = class_2(X)
                Case Else
                    class_predict_fold2(X) = class_3(X)
            End Select
        End If
    Next
    
    For X = 0 To 296
        If class_predict_fold2(X) = fold_2(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(1) = (ctr_accuracy / 297) * 100
    
    List1.AddItem "2th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(1), 6) & "%"
    
    'fold 3-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        For Y = 0 To 1186
            If distance_3(X, Y) <= min_1(X) Then
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_3(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_1(X) And distance_3(X, Y) <= min_2(X) And distance_3(X, Y) <= min_3(X) Then
                min_3(X) = min_2(X)
                min_2(X) = distance_3(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_2(X) And distance_3(X, Y) <= min_3(X) Then
                min_3(X) = distance_3(X, Y)
                class_3(X) = class_predict_candidate_3(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold3(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) Then
            class_predict_fold3(X) = class_1(X)
            
        ElseIf class_2(X) = class_3(X) Then
            class_predict_fold3(X) = class_2(X)
            
        ElseIf class_3(X) = class_1(X) Then
            class_predict_fold3(X) = class_1(X)
            
        Else
            mode3 = X Mod 3
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_1(X)
                Case 2
                    class_predict_fold3(X) = class_2(X)
                Case Else
                    class_predict_fold3(X) = class_3(X)
            End Select
        End If
    Next
    
    For X = 0 To 296
        If class_predict_fold3(X) = fold_3(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(2) = (ctr_accuracy / 297) * 100
    
    List1.AddItem "3rd  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(2), 6) & "%"
    
    'fold 4-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        For Y = 0 To 1186
            If distance_4(X, Y) <= min_1(X) Then
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_4(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_1(X) And distance_4(X, Y) <= min_2(X) And distance_4(X, Y) <= min_3(X) Then
                min_3(X) = min_2(X)
                min_2(X) = distance_4(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_2(X) And distance_4(X, Y) <= min_3(X) Then
                min_3(X) = distance_4(X, Y)
                class_3(X) = class_predict_candidate_4(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold4(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) Then
            class_predict_fold4(X) = class_1(X)
            
        ElseIf class_2(X) = class_3(X) Then
            class_predict_fold4(X) = class_2(X)
            
        ElseIf class_3(X) = class_1(X) Then
            class_predict_fold4(X) = class_1(X)
            
        Else
            mode4 = X Mod 3
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_1(X)
                Case 2
                    class_predict_fold4(X) = class_2(X)
                Case Else
                    class_predict_fold4(X) = class_3(X)
            End Select
        End If
    Next
    
    For X = 0 To 296
        If class_predict_fold4(X) = fold_4(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(3) = (ctr_accuracy / 297) * 100
    
    List1.AddItem "4th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(3), 6) & "%"
 
    'fold 5-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        For Y = 0 To 1186
            If distance_5(X, Y) <= min_1(X) Then
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_5(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_1(X) And distance_5(X, Y) <= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_3(X) = min_2(X)
                min_2(X) = distance_5(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_3(X) = distance_5(X, Y)
                class_3(X) = class_predict_candidate_5(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold5(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) Then
            class_predict_fold5(X) = class_1(X)
            
        ElseIf class_2(X) = class_3(X) Then
            class_predict_fold5(X) = class_2(X)
            
        ElseIf class_3(X) = class_1(X) Then
            class_predict_fold5(X) = class_1(X)
            
        Else
            mode5 = X Mod 3
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_1(X)
                Case 2
                    class_predict_fold5(X) = class_2(X)
                Case Else
                    class_predict_fold5(X) = class_3(X)
            End Select
        End If
    Next
    
    For X = 0 To 296
        If class_predict_fold5(X) = fold_5(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(4) = (ctr_accuracy / 297) * 100
    List1.AddItem "5th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(4), 6) & "%"
    List1.AddItem "-----------------------------------------------------"
    List1.AddItem "average accuracy: " & FormatNumber(((accuracy(0) + accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4)) / 5), 6) & "%"
End Sub

Private Sub Command2_click() '紀錄k=4
    List1.Clear
    Dim min_1(1188) As Double
    Dim min_2(1188) As Double
    Dim min_3(1188) As Double
    Dim min_4(1188) As Double
    Dim class_1(1188) As String
    Dim class_2(1188) As String
    Dim class_3(1188) As String
    Dim class_4(1188) As String
    Dim accuracy(5) As Double
    'fold 1-----------------------------------------------
    For X = 0 To 295
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        For Y = 0 To 1187
            If distance(X, Y) <= min_1(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_1(X) And distance(X, Y) <= min_2(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_2(X) And distance(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = distance(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_3(X) And distance(X, Y) <= min_4(X) Then
                min_4(X) = distance(X, Y)
                class_3(X) = class_predict_candidate(X, Y)
            End If
        Next Y
    Next X
    
    Dim mode As Integer
    Dim class_predict_fold1(296) As String
    Dim rnd_class As Integer
    Dim ctr_accuracy As Integer
    For X = 0 To 295
        If class_1(X) = class_2(X) And class_1(X) = class_3(X) Then '123
            class_predict_fold1(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) And class_3(X) = class_4(X) Then '234
            class_predict_fold1(X) = class_2(X)
        ElseIf class_1(X) = class_3(X) And class_3(X) = class_4(X) Then '134
            class_predict_fold1(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) And class_2(X) = class_4(X) Then '124
            class_predict_fold1(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) Then '12
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_1(X)
                Case Else
                    class_predict_fold1(X) = class_2(X)
            End Select
        ElseIf class_1(X) = class_3(X) Then '13
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_1(X)
                Case Else
                    class_predict_fold1(X) = class_3(X)
            End Select
        ElseIf class_1(X) = class_4(X) Then '14
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_1(X)
                Case Else
                    class_predict_fold1(X) = class_4(X)
            End Select
         ElseIf class_2(X) = class_3(X) Then '23
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_2(X)
                Case Else
                    class_predict_fold1(X) = class_3(X)
            End Select
        ElseIf class_2(X) = class_4(X) Then '24
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_2(X)
                Case Else
                    class_predict_fold1(X) = class_4(X)
            End Select
        ElseIf class_3(X) = class_4(X) Then '34
            mode = X Mod 2
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_3(X)
                Case Else
                    class_predict_fold1(X) = class_4(X)
            End Select
        Else
            mode = X Mod 4
            Select Case mode
                Case 1
                    class_predict_fold1(X) = class_1(X)
                Case 2
                    class_predict_fold1(X) = class_2(X)
                Case 3
                    class_predict_fold1(X) = class_3(X)
                Case Else
                    class_predict_fold1(X) = class_4(X)
            End Select
        End If
    Next

    For X = 0 To 295
        If class_predict_fold1(X) = fold_1(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(0) = (ctr_accuracy / 296) * 100
    List1.AddItem "i_th fold" & vbTab & "#data" & vbTab & "#accurate data" & vbTab & "  accuracy"
    List1.AddItem "1st  fold" & vbTab & "296" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(0), 6) & "%"
    
    'fold 2-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        For Y = 0 To 1186
            If distance_2(X, Y) <= min_1(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_2(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_1(X) And distance_2(X, Y) <= min_2(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_2(X, Y)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_2(X) And distance_2(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = distance_2(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_3(X) And distance_2(X, Y) <= min_4(X) Then
                min_4(X) = distance_2(X, Y)
                class_3(X) = class_predict_candidate_2(X, Y)
            End If
        Next Y
    Next X


    Dim class_predict_fold2(297) As String
    Dim mode2 As Integer
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) And class_1(X) = class_3(X) Then '123
            class_predict_fold2(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) And class_3(X) = class_4(X) Then '234
            class_predict_fold2(X) = class_2(X)
        ElseIf class_1(X) = class_3(X) And class_3(X) = class_4(X) Then '134
            class_predict_fold2(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) And class_2(X) = class_4(X) Then '124
            class_predict_fold2(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) Then '12
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_1(X)
                Case Else
                    class_predict_fold2(X) = class_2(X)
            End Select
        ElseIf class_1(X) = class_3(X) Then '13
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_1(X)
                Case Else
                    class_predict_fold2(X) = class_3(X)
            End Select
        ElseIf class_1(X) = class_4(X) Then '14
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_1(X)
                Case Else
                    class_predict_fold2(X) = class_4(X)
            End Select
         ElseIf class_2(X) = class_3(X) Then '23
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_2(X)
                Case Else
                    class_predict_fold2(X) = class_3(X)
            End Select
        ElseIf class_2(X) = class_4(X) Then '24
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_2(X)
                Case Else
                    class_predict_fold2(X) = class_4(X)
            End Select
        ElseIf class_3(X) = class_4(X) Then '34
            mode2 = X Mod 2
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_3(X)
                Case Else
                    class_predict_fold2(X) = class_4(X)
            End Select
        Else
            mode2 = X Mod 4
            Select Case mode2
                Case 1
                    class_predict_fold2(X) = class_1(X)
                Case 2
                    class_predict_fold2(X) = class_2(X)
                Case 3
                    class_predict_fold2(X) = class_3(X)
                Case Else
                    class_predict_fold2(X) = class_4(X)
            End Select
        End If
    Next

    For X = 0 To 296
        If class_predict_fold2(X) = fold_2(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(1) = (ctr_accuracy / 297) * 100
    List1.AddItem "2th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(1), 6) & "%"

    'fold 3-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        For Y = 0 To 1186
            If distance_3(X, Y) <= min_1(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_3(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_1(X) And distance_3(X, Y) <= min_2(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_3(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_2(X) And distance_3(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = distance_3(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_3(X) And distance_3(X, Y) <= min_4(X) Then
                min_4(X) = distance_3(X, Y)
                class_3(X) = class_predict_candidate_3(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold3(297) As String
    Dim mode3 As Integer
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) And class_1(X) = class_3(X) Then '123
            class_predict_fold3(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) And class_3(X) = class_4(X) Then '234
            class_predict_fold3(X) = class_2(X)
        ElseIf class_1(X) = class_3(X) And class_3(X) = class_4(X) Then '134
            class_predict_fold3(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) And class_2(X) = class_4(X) Then '124
            class_predict_fold3(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) Then '12
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_1(X)
                Case Else
                    class_predict_fold3(X) = class_2(X)
            End Select
        ElseIf class_1(X) = class_3(X) Then '13
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_1(X)
                Case Else
                    class_predict_fold3(X) = class_3(X)
            End Select
        ElseIf class_1(X) = class_4(X) Then '14
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_1(X)
                Case Else
                    class_predict_fold3(X) = class_4(X)
            End Select
         ElseIf class_2(X) = class_3(X) Then '23
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_2(X)
                Case Else
                    class_predict_fold3(X) = class_3(X)
            End Select
        ElseIf class_2(X) = class_4(X) Then '24
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_2(X)
                Case Else
                    class_predict_fold3(X) = class_4(X)
            End Select
        ElseIf class_3(X) = class_4(X) Then '34
            mode3 = X Mod 2
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_3(X)
                Case Else
                    class_predict_fold3(X) = class_4(X)
            End Select
        Else
            mode3 = X Mod 4
            Select Case mode3
                Case 1
                    class_predict_fold3(X) = class_1(X)
                Case 2
                    class_predict_fold3(X) = class_2(X)
                Case 3
                    class_predict_fold3(X) = class_3(X)
                Case Else
                    class_predict_fold3(X) = class_4(X)
            End Select
        End If
    Next

    For X = 0 To 296
        If class_predict_fold3(X) = fold_3(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(2) = (ctr_accuracy / 297) * 100
    List1.AddItem "3rd  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(2), 6) & "%"

    'fold 4-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_3(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        For Y = 0 To 1186
            If distance_4(X, Y) <= min_1(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_4(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_1(X) And distance_4(X, Y) <= min_2(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_4(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_2(X) And distance_4(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = distance_4(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_3(X) And distance_4(X, Y) <= min_4(X) Then
                min_4(X) = distance_4(X, Y)
                class_3(X) = class_predict_candidate_4(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold4(297) As String
    Dim mode4 As Integer
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) And class_1(X) = class_3(X) Then '123
            class_predict_fold4(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) And class_3(X) = class_4(X) Then '234
            class_predict_fold4(X) = class_2(X)
        ElseIf class_1(X) = class_3(X) And class_3(X) = class_4(X) Then '134
            class_predict_fold4(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) And class_2(X) = class_4(X) Then '124
            class_predict_fold4(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) Then '12
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_1(X)
                Case Else
                    class_predict_fold4(X) = class_2(X)
            End Select
        ElseIf class_1(X) = class_3(X) Then '13
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_1(X)
                Case Else
                    class_predict_fold4(X) = class_3(X)
            End Select
        ElseIf class_1(X) = class_4(X) Then '14
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_1(X)
                Case Else
                    class_predict_fold4(X) = class_4(X)
            End Select
         ElseIf class_2(X) = class_3(X) Then '23
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_2(X)
                Case Else
                    class_predict_fold4(X) = class_3(X)
            End Select
        ElseIf class_2(X) = class_4(X) Then '24
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_2(X)
                Case Else
                    class_predict_fold4(X) = class_4(X)
            End Select
        ElseIf class_3(X) = class_4(X) Then '34
            mode4 = X Mod 2
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_3(X)
                Case Else
                    class_predict_fold4(X) = class_4(X)
            End Select
        Else
            mode4 = X Mod 4
            Select Case mode4
                Case 1
                    class_predict_fold4(X) = class_1(X)
                Case 2
                    class_predict_fold4(X) = class_2(X)
                Case 3
                    class_predict_fold4(X) = class_3(X)
                Case Else
                    class_predict_fold4(X) = class_4(X)
            End Select
        End If
    Next

    For X = 0 To 296
        If class_predict_fold4(X) = fold_4(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(3) = (ctr_accuracy / 297) * 100

    List1.AddItem "4th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(3), 6) & "%"

    'fold 5-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        For Y = 0 To 1186
            If distance_5(X, Y) <= min_1(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_5(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_1(X) And distance_5(X, Y) <= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_5(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_4(X) = min_3(X)
                min_3(X) = distance_5(X, Y)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_5(X, Y)
             ElseIf distance_5(X, Y) >= min_3(X) And distance_5(X, Y) <= min_4(X) Then
                min_4(X) = distance_5(X, Y)
                class_3(X) = class_predict_candidate_5(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold5(297) As String
    Dim mode5 As Integer
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        If class_1(X) = class_2(X) And class_1(X) = class_3(X) Then '123
            class_predict_fold5(X) = class_1(X)
        ElseIf class_2(X) = class_3(X) And class_3(X) = class_4(X) Then '234
            class_predict_fold5(X) = class_2(X)
        ElseIf class_1(X) = class_3(X) And class_3(X) = class_4(X) Then '134
            class_predict_fold5(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) And class_2(X) = class_4(X) Then '124
            class_predict_fold5(X) = class_1(X)
        ElseIf class_1(X) = class_2(X) Then '12
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_1(X)
                Case Else
                    class_predict_fold5(X) = class_2(X)
            End Select
        ElseIf class_1(X) = class_3(X) Then '13
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_1(X)
                Case Else
                    class_predict_fold5(X) = class_3(X)
            End Select
        ElseIf class_1(X) = class_4(X) Then '14
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_1(X)
                Case Else
                    class_predict_fold5(X) = class_4(X)
            End Select
         ElseIf class_2(X) = class_3(X) Then '23
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_2(X)
                Case Else
                    class_predict_fold5(X) = class_3(X)
            End Select
        ElseIf class_2(X) = class_4(X) Then '24
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_2(X)
                Case Else
                    class_predict_fold5(X) = class_4(X)
            End Select
        ElseIf class_3(X) = class_4(X) Then '34
            mode5 = X Mod 2
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_3(X)
                Case Else
                    class_predict_fold5(X) = class_4(X)
            End Select
        Else
            mode5 = X Mod 5
            Select Case mode5
                Case 1
                    class_predict_fold5(X) = class_1(X)
                Case 2
                    class_predict_fold5(X) = class_2(X)
                Case 3
                    class_predict_fold5(X) = class_3(X)
                Case Else
                    class_predict_fold5(X) = class_4(X)
            End Select
        End If
    Next

    For X = 0 To 296
        If class_predict_fold5(X) = fold_5(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(4) = (ctr_accuracy / 297) * 100
    List1.AddItem "5th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(4), 6) & "%"
    List1.AddItem "-----------------------------------------------------"
    List1.AddItem "average accuracy: " & FormatNumber(((accuracy(0) + accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4)) / 5), 6) & "%"
End Sub

Private Sub Command3_click() '紀錄k=5
    List1.Clear
    Dim class_name(11) As String
    class_name(1) = "CYT"
    class_name(2) = "NUC"
    class_name(3) = "MIT"
    class_name(4) = "ME3"
    class_name(5) = "ME2"
    class_name(6) = "ME1"
    class_name(7) = "EXC"
    class_name(8) = "VAC"
    class_name(9) = "POX"
    class_name(10) = "ERL"
    Dim average(5) As Double
    Dim min_1(1188) As Double
    Dim min_2(1188) As Double
    Dim min_3(1188) As Double
    Dim min_4(1188) As Double
    Dim min_5(1188) As Double
    Dim class_1(1188) As String
    Dim class_2(1188) As String
    Dim class_3(1188) As String
    Dim class_4(1188) As String
    Dim class_5(1188) As String
    
    'fold 1-----------------------------------------------
    For X = 0 To 295
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        For Y = 0 To 1187
            If distance(X, Y) <= min_1(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_1(X) And distance(X, Y) <= min_2(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_2(X) And distance(X, Y) <= min_3(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_3(X) And distance(X, Y) <= min_4(X) Then
                min_5(X) = min_4(X)
                min_4(X) = distance(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_4(X) And distance(X, Y) <= min_5(X) Then
                min_5(X) = distance(X, Y)
                class_5(X) = class_predict_candidate(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold1(296) As String
    Dim class_count_fold1(11) As Integer
    Dim class_count_fold2(11) As Integer
    Dim class_count_fold3(11) As Integer
    Dim class_count_fold4(11) As Integer
    Dim class_count_fold5(11) As Integer
    Dim rnd_class As Integer
    Dim ctr_accuracy As Integer
    Dim accuracy(5) As Double
    Dim tmpctr As Integer
    Dim tmpctr_cand221(2) As String
    Dim tmpctr_cand11111(5) As String
    Dim mode As Integer
    Dim mode2 As Integer
    Dim mode3 As Integer
    Dim mode4 As Integer
    Dim mode5 As Integer
    
    For X = 0 To 295
        For Y = 1 To 10
            class_count_fold1(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold1(Y) = 5 Then
                class_predict_fold1(X) = class_name(Y)
            ElseIf class_count_fold1(Y) = 4 Then
                class_predict_fold1(X) = class_name(Y)
            ElseIf class_count_fold1(Y) = 3 Then
                class_predict_fold1(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand221(0) = 0
        tmpctr_cand221(1) = 0
        tmpctr_cand11111(0) = 0
        tmpctr_cand11111(1) = 0
        tmpctr_cand11111(2) = 0
        tmpctr_cand11111(3) = 0
        tmpctr_cand11111(4) = 0
        
        If class_predict_fold1(X) = "" Then
            For Y = 1 To 10
                If class_count_fold1(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            If tmpctr = 1 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 2 Then
                        If tmpctr_cand221(0) = "0" Then
                            tmpctr_cand221(0) = class_name(Y)
                        ElseIf tmpctr_cand221(0) <> "0" Then
                            tmpctr_cand221(1) = class_name(Y)
                        End If
                    End If
                Next

                mode = X Mod 2
                Select Case mode
                    Case 1
                        class_predict_fold1(X) = tmpctr_cand221(0)
                    Case Else
                        class_predict_fold1(X) = tmpctr_cand221(1)
                End Select
            End If
            
            If tmpctr = 3 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 2 Then
                        class_predict_fold1(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 5 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 1 Then
                        If tmpctr_cand11111(0) = "0" Then
                            tmpctr_cand11111(0) = class_name(Y)
                        ElseIf tmpctr_cand11111(1) = "0" Then
                            tmpctr_cand11111(1) = class_name(Y)
                        ElseIf tmpctr_cand11111(2) = "0" Then
                            tmpctr_cand11111(2) = class_name(Y)
                        ElseIf tmpctr_cand11111(3) = "0" Then
                            tmpctr_cand11111(3) = class_name(Y)
                        ElseIf tmpctr_cand11111(4) = "0" Then
                            tmpctr_cand11111(4) = class_name(Y)
                        End If
                    End If
                Next
                mode = X Mod 5
                Select Case mode
                    Case 1
                        class_predict_fold1(X) = tmpctr_cand11111(0)
                    Case 2
                        class_predict_fold1(X) = tmpctr_cand11111(1)
                    Case 3
                        class_predict_fold1(X) = tmpctr_cand11111(2)
                    Case 4
                        class_predict_fold1(X) = tmpctr_cand11111(3)
                    Case Else
                        class_predict_fold1(X) = tmpctr_cand11111(4)
                End Select
            End If
        End If
        
        Next
    
    For X = 0 To 295
        If class_predict_fold1(X) = fold_1(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(0) = (ctr_accuracy / 296) * 100
    List1.AddItem "i_th fold" & vbTab & "#data" & vbTab & "#accurate data" & vbTab & "  accuracy"
    List1.AddItem "1st  fold" & vbTab & "296" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(0), 6) & "%"
    
    
    'fold 2-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        For Y = 0 To 1186
            If distance_2(X, Y) <= min_1(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_2(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_1(X) And distance_2(X, Y) <= min_2(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_2(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_2(X) And distance_2(X, Y) <= min_3(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_2(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_3(X) And distance_2(X, Y) <= min_4(X) Then
                min_5(X) = min_4(X)
                min_4(X) = distance_2(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_4(X) And distance_2(X, Y) <= min_5(X) Then
                min_5(X) = distance_2(X, Y)
                class_5(X) = class_predict_candidate_2(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold2(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold2(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold2(Y) = 5 Then
                class_predict_fold2(X) = class_name(Y)
            ElseIf class_count_fold2(Y) = 4 Then
                class_predict_fold2(X) = class_name(Y)
            ElseIf class_count_fold2(Y) = 3 Then
                class_predict_fold2(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand221(0) = 0
        tmpctr_cand221(1) = 0
        tmpctr_cand11111(0) = 0
        tmpctr_cand11111(1) = 0
        tmpctr_cand11111(2) = 0
        tmpctr_cand11111(3) = 0
        tmpctr_cand11111(4) = 0
        
        If class_predict_fold2(X) = "" Then
            For Y = 1 To 10
                If class_count_fold2(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            If tmpctr = 1 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 2 Then
                        If tmpctr_cand221(0) = "0" Then
                            tmpctr_cand221(0) = class_name(Y)
                        ElseIf tmpctr_cand221(0) <> "0" Then
                            tmpctr_cand221(1) = class_name(Y)
                        End If
                    End If
                Next

                mode2 = X Mod 2
                Select Case mode2
                    Case 1
                        class_predict_fold2(X) = tmpctr_cand221(0)
                    Case Else
                        class_predict_fold2(X) = tmpctr_cand221(1)
                End Select
            End If
            
            If tmpctr = 3 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 2 Then
                        class_predict_fold2(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 5 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 1 Then
                        If tmpctr_cand11111(0) = "0" Then
                            tmpctr_cand11111(0) = class_name(Y)
                        ElseIf tmpctr_cand11111(1) = "0" Then
                            tmpctr_cand11111(1) = class_name(Y)
                        ElseIf tmpctr_cand11111(2) = "0" Then
                            tmpctr_cand11111(2) = class_name(Y)
                        ElseIf tmpctr_cand11111(3) = "0" Then
                            tmpctr_cand11111(3) = class_name(Y)
                        ElseIf tmpctr_cand11111(4) = "0" Then
                            tmpctr_cand11111(4) = class_name(Y)
                        End If
                    End If
                Next
                mode2 = X Mod 5
                Select Case mode2
                    Case 1
                        class_predict_fold2(X) = tmpctr_cand11111(0)
                    Case 2
                        class_predict_fold2(X) = tmpctr_cand11111(1)
                    Case 3
                        class_predict_fold2(X) = tmpctr_cand11111(2)
                    Case 4
                        class_predict_fold2(X) = tmpctr_cand11111(3)
                    Case Else
                        class_predict_fold2(X) = tmpctr_cand11111(4)
                End Select
            End If
        End If
    Next

    For X = 0 To 296
        If class_predict_fold2(X) = fold_2(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(1) = (ctr_accuracy / 297) * 100
    List1.AddItem "2th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(1), 6) & "%"

    'fold 3-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        For Y = 0 To 1186
            If distance_3(X, Y) <= min_1(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_3(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_1(X) And distance_3(X, Y) <= min_2(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_3(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_2(X) And distance_3(X, Y) <= min_3(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_3(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_3(X) And distance_3(X, Y) <= min_4(X) Then
                min_5(X) = min_4(X)
                min_4(X) = distance_3(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_4(X) And distance_3(X, Y) <= min_5(X) Then
                min_5(X) = distance_3(X, Y)
                class_5(X) = class_predict_candidate_3(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold3(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold3(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold3(Y) = 5 Then
                class_predict_fold3(X) = class_name(Y)
            ElseIf class_count_fold3(Y) = 4 Then
                class_predict_fold3(X) = class_name(Y)
            ElseIf class_count_fold3(Y) = 3 Then
                class_predict_fold3(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand221(0) = 0
        tmpctr_cand221(1) = 0
        tmpctr_cand11111(0) = 0
        tmpctr_cand11111(1) = 0
        tmpctr_cand11111(2) = 0
        tmpctr_cand11111(3) = 0
        tmpctr_cand11111(4) = 0
        
        If class_predict_fold3(X) = "" Then
            For Y = 1 To 10
                If class_count_fold3(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            If tmpctr = 1 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 2 Then
                        If tmpctr_cand221(0) = "0" Then
                            tmpctr_cand221(0) = class_name(Y)
                        ElseIf tmpctr_cand221(0) <> "0" Then
                            tmpctr_cand221(1) = class_name(Y)
                        End If
                    End If
                Next

                mode3 = X Mod 2
                Select Case mode3
                    Case 1
                        class_predict_fold3(X) = tmpctr_cand221(0)
                    Case Else
                        class_predict_fold3(X) = tmpctr_cand221(1)
                End Select
            End If
            
            If tmpctr = 3 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 2 Then
                        class_predict_fold3(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 5 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 1 Then
                        If tmpctr_cand11111(0) = "0" Then
                            tmpctr_cand11111(0) = class_name(Y)
                        ElseIf tmpctr_cand11111(1) = "0" Then
                            tmpctr_cand11111(1) = class_name(Y)
                        ElseIf tmpctr_cand11111(2) = "0" Then
                            tmpctr_cand11111(2) = class_name(Y)
                        ElseIf tmpctr_cand11111(3) = "0" Then
                            tmpctr_cand11111(3) = class_name(Y)
                        ElseIf tmpctr_cand11111(4) = "0" Then
                            tmpctr_cand11111(4) = class_name(Y)
                        End If
                    End If
                Next
                mode3 = X Mod 5
                Select Case mode3
                    Case 1
                        class_predict_fold3(X) = tmpctr_cand11111(0)
                    Case 2
                        class_predict_fold3(X) = tmpctr_cand11111(1)
                    Case 3
                        class_predict_fold3(X) = tmpctr_cand11111(2)
                    Case 4
                        class_predict_fold3(X) = tmpctr_cand11111(3)
                    Case Else
                        class_predict_fold3(X) = tmpctr_cand11111(4)
                End Select
            End If
        End If
    Next

    For X = 0 To 296
        If class_predict_fold3(X) = fold_3(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(2) = (ctr_accuracy / 297) * 100
    List1.AddItem "3rd  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(2), 6) & "%"

    'fold 4-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        For Y = 0 To 1186
            If distance_4(X, Y) <= min_1(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_4(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_1(X) And distance_4(X, Y) <= min_2(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_4(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_2(X) And distance_4(X, Y) <= min_3(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_4(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_3(X) And distance_4(X, Y) <= min_4(X) Then
                min_5(X) = min_4(X)
                min_4(X) = distance_4(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_4(X) And distance_4(X, Y) <= min_5(X) Then
                min_5(X) = distance_4(X, Y)
                class_5(X) = class_predict_candidate_4(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold4(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold4(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_3(X)
           Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold4(Y) = 5 Then
                class_predict_fold4(X) = class_name(Y)
            ElseIf class_count_fold4(Y) = 4 Then
                class_predict_fold4(X) = class_name(Y)
            ElseIf class_count_fold4(Y) = 3 Then
                class_predict_fold4(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand221(0) = 0
        tmpctr_cand221(1) = 0
        tmpctr_cand11111(0) = 0
        tmpctr_cand11111(1) = 0
        tmpctr_cand11111(2) = 0
        tmpctr_cand11111(3) = 0
        tmpctr_cand11111(4) = 0
        
        If class_predict_fold4(X) = "" Then
            For Y = 1 To 10
                If class_count_fold4(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            If tmpctr = 1 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 2 Then
                        If tmpctr_cand221(0) = "0" Then
                            tmpctr_cand221(0) = class_name(Y)
                        ElseIf tmpctr_cand221(0) <> "0" Then
                            tmpctr_cand221(1) = class_name(Y)
                        End If
                    End If
                Next

                mode4 = X Mod 2
                Select Case mode4
                    Case 1
                        class_predict_fold4(X) = tmpctr_cand221(0)
                    Case Else
                        class_predict_fold4(X) = tmpctr_cand221(1)
                End Select
            End If
            
            If tmpctr = 3 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 2 Then
                        class_predict_fold4(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 5 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 1 Then
                        If tmpctr_cand11111(0) = "0" Then
                            tmpctr_cand11111(0) = class_name(Y)
                        ElseIf tmpctr_cand11111(1) = "0" Then
                            tmpctr_cand11111(1) = class_name(Y)
                        ElseIf tmpctr_cand11111(2) = "0" Then
                            tmpctr_cand11111(2) = class_name(Y)
                        ElseIf tmpctr_cand11111(3) = "0" Then
                            tmpctr_cand11111(3) = class_name(Y)
                        ElseIf tmpctr_cand11111(4) = "0" Then
                            tmpctr_cand11111(4) = class_name(Y)
                        End If
                    End If
                Next
                mode4 = X Mod 5
                Select Case mode4
                    Case 1
                        class_predict_fold4(X) = tmpctr_cand11111(0)
                    Case 2
                        class_predict_fold4(X) = tmpctr_cand11111(1)
                    Case 3
                        class_predict_fold4(X) = tmpctr_cand11111(2)
                    Case 4
                        class_predict_fold4(X) = tmpctr_cand11111(3)
                    Case Else
                        class_predict_fold4(X) = tmpctr_cand11111(4)
                End Select
            End If
        End If
    Next

    For X = 0 To 296
        If class_predict_fold4(X) = fold_4(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(3) = (ctr_accuracy / 297) * 100
    List1.AddItem "4th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(3), 6) & "%"

    'fold 5-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        For Y = 0 To 1186
            If distance_5(X, Y) <= min_1(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_5(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_1(X) And distance_5(X, Y) <= min_2(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_5(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_5(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_3(X) And distance_5(X, Y) <= min_4(X) Then
                min_5(X) = min_4(X)
                min_4(X) = distance_5(X, Y)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_4(X) And distance_5(X, Y) <= min_5(X) Then
                min_5(X) = distance_5(X, Y)
                class_5(X) = class_predict_candidate_5(X, Y)
            End If
        Next Y
    Next X

    Dim class_predict_fold5(297) As String
    rnd_class = 0
    ctr_accuracy = 0
    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold5(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold5(Y) = 5 Then
                class_predict_fold5(X) = class_name(Y)
            ElseIf class_count_fold5(Y) = 4 Then
                class_predict_fold5(X) = class_name(Y)
            ElseIf class_count_fold5(Y) = 3 Then
                class_predict_fold5(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand221(0) = 0
        tmpctr_cand221(1) = 0
        tmpctr_cand11111(0) = 0
        tmpctr_cand11111(1) = 0
        tmpctr_cand11111(2) = 0
        tmpctr_cand11111(3) = 0
        tmpctr_cand11111(4) = 0
        
        If class_predict_fold5(X) = "" Then
            For Y = 1 To 10
                If class_count_fold5(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            If tmpctr = 1 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 2 Then
                        If tmpctr_cand221(0) = "0" Then
                            tmpctr_cand221(0) = class_name(Y)
                        ElseIf tmpctr_cand221(0) <> "0" Then
                            tmpctr_cand221(1) = class_name(Y)
                        End If
                    End If
                Next

                mode5 = X Mod 2
                Select Case mode5
                    Case 1
                        class_predict_fold5(X) = tmpctr_cand221(0)
                    Case Else
                        class_predict_fold5(X) = tmpctr_cand221(1)
                End Select
            End If
            
            If tmpctr = 3 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 2 Then
                        class_predict_fold5(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 5 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 1 Then
                        If tmpctr_cand11111(0) = "0" Then
                            tmpctr_cand11111(0) = class_name(Y)
                        ElseIf tmpctr_cand11111(1) = "0" Then
                            tmpctr_cand11111(1) = class_name(Y)
                        ElseIf tmpctr_cand11111(2) = "0" Then
                            tmpctr_cand11111(2) = class_name(Y)
                        ElseIf tmpctr_cand11111(3) = "0" Then
                            tmpctr_cand11111(3) = class_name(Y)
                        ElseIf tmpctr_cand11111(4) = "0" Then
                            tmpctr_cand11111(4) = class_name(Y)
                        End If
                    End If
                Next
                mode5 = X Mod 5
                Select Case mode5
                    Case 1
                        class_predict_fold5(X) = tmpctr_cand11111(0)
                    Case 2
                        class_predict_fold5(X) = tmpctr_cand11111(1)
                    Case 3
                        class_predict_fold5(X) = tmpctr_cand11111(2)
                    Case 4
                        class_predict_fold5(X) = tmpctr_cand11111(3)
                    Case Else
                        class_predict_fold5(X) = tmpctr_cand11111(4)
                End Select
            End If
        End If
    Next

    For X = 0 To 296
        If class_predict_fold5(X) = fold_5(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(4) = (ctr_accuracy / 297) * 100
    List1.AddItem "5th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(4), 6) & "%"
    List1.AddItem "-----------------------------------------------------"
    List1.AddItem "average accuracy: " & FormatNumber(((accuracy(0) + accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4)) / 5), 6) & "%"
End Sub
Private Sub Command4_click() '紀錄k=6
    List1.Clear
    Dim class_name(11) As String
    class_name(1) = "CYT"
    class_name(2) = "NUC"
    class_name(3) = "MIT"
    class_name(4) = "ME3"
    class_name(5) = "ME2"
    class_name(6) = "ME1"
    class_name(7) = "EXC"
    class_name(8) = "VAC"
    class_name(9) = "POX"
    class_name(10) = "ERL"
    Dim average(5) As Double
    Dim min_1(1188) As Double
    Dim min_2(1188) As Double
    Dim min_3(1188) As Double
    Dim min_4(1188) As Double
    Dim min_5(1188) As Double
    Dim min_6(1188) As Double
    Dim class_1(1188) As String
    Dim class_2(1188) As String
    Dim class_3(1188) As String
    Dim class_4(1188) As String
    Dim class_5(1188) As String
    Dim class_6(1188) As String
    
    'fold 1-----------------------------------------------
    For X = 0 To 295
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        class_6(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        min_6(X) = 10000
        For Y = 0 To 1187
            If distance(X, Y) <= min_1(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_1(X) And distance(X, Y) <= min_2(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_2(X) And distance(X, Y) <= min_3(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_3(X) And distance(X, Y) <= min_4(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = distance(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_4(X) And distance(X, Y) <= min_5(X) Then
                min_6(X) = min_5(X)
                min_5(X) = distance(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_predict_candidate(X, Y)
            ElseIf distance(X, Y) >= min_5(X) And distance(X, Y) <= min_6(X) Then
                min_6(X) = distance(X, Y)
                class_6(X) = class_predict_candidate(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold1(296) As String
    Dim class_count_fold1(11) As Integer
    Dim class_count_fold2(11) As Integer
    Dim class_count_fold3(11) As Integer
    Dim class_count_fold4(11) As Integer
    Dim class_count_fold5(11) As Integer
    Dim class_count_fold6(11) As Integer
    Dim rnd_class As Integer
    Dim ctr_accuracy As Integer
    Dim accuracy(5) As Double
    Dim tmpctr As Integer
    Dim tmpctr_cand33(2) As String
    Dim tmpctr_cand222(3) As String
    Dim tmpctr_cand2211(2) As String
    Dim tmpctr_cand111111(6) As String
    Dim mode As Integer
    Dim mode2 As Integer
    Dim mode3 As Integer
    Dim mode4 As Integer
    Dim mode5 As Integer
    Dim mode6 As Integer
    
    For X = 0 To 295
        For Y = 1 To 10
            class_count_fold1(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        Select Case class_6(X)
            Case "CYT"
                class_count_fold1(1) = class_count_fold1(1) + 1
            Case "NUC"
                class_count_fold1(2) = class_count_fold1(2) + 1
            Case "MIT"
                class_count_fold1(3) = class_count_fold1(3) + 1
            Case "ME3"
                class_count_fold1(4) = class_count_fold1(4) + 1
            Case "ME2"
                class_count_fold1(5) = class_count_fold1(5) + 1
            Case "ME1"
                class_count_fold1(6) = class_count_fold1(6) + 1
            Case "EXC"
                class_count_fold1(7) = class_count_fold1(7) + 1
            Case "VAC"
                class_count_fold1(8) = class_count_fold1(8) + 1
            Case "POX"
                class_count_fold1(9) = class_count_fold1(9) + 1
            Case Else
                class_count_fold1(10) = class_count_fold1(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold1(Y) = 6 Then
                class_predict_fold1(X) = class_name(Y)
            ElseIf class_count_fold1(Y) = 5 Then
                class_predict_fold1(X) = class_name(Y)
            ElseIf class_count_fold1(Y) = 4 Then
                class_predict_fold1(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand33(0) = 0
        tmpctr_cand33(1) = 0
        tmpctr_cand2211(0) = 0
        tmpctr_cand2211(1) = 0
        For Y = 0 To 2
            tmpctr_cand222(Y) = 0
        Next
        For Y = 0 To 5
            tmpctr_cand111111(Y) = 0
        Next
        
        If class_predict_fold1(X) = "" Then
            For Y = 1 To 10
                If class_count_fold1(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            
            If tmpctr = 0 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 3 Then '33的情況
                        If tmpctr_cand33(0) = "0" Then
                                tmpctr_cand33(0) = class_name(Y)
                            Else: tmpctr_cand33(1) = class_name(Y)
                        End If
                    ElseIf class_count_fold1(Y) = 2 Then '222的情況
                        If tmpctr_cand222(0) = "0" Then
                                tmpctr_cand222(0) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(1) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(2) = class_name(Y)
                        End If
                    End If
                Next
                For Y = 1 To 10
                    If class_count_fold1(Y) = 3 Then
                        mode = X Mod 2
                        Select Case mode
                            Case 1
                                class_predict_fold1(X) = tmpctr_cand33(0)
                            Case Else
                                class_predict_fold1(X) = tmpctr_cand33(1)
                        End Select
                    End If
                    If class_count_fold1(Y) = 2 Then
                        mode = X Mod 3
                        Select Case mode
                            Case 1
                                class_predict_fold1(X) = tmpctr_cand222(0)
                            Case 2
                                class_predict_fold1(X) = tmpctr_cand222(1)
                            Case Else
                                class_predict_fold1(X) = tmpctr_cand222(2)
                        End Select
                    End If
                Next
            End If

            If tmpctr = 2 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 2 Then
                        If tmpctr_cand2211(0) = "0" Then
                            tmpctr_cand2211(0) = class_name(Y)
                        ElseIf tmpctr_cand2211(0) <> "0" Then
                            tmpctr_cand2211(1) = class_name(Y)
                        End If
                    End If
                Next

                mode = X Mod 2
                Select Case mode
                    Case 1
                        class_predict_fold1(X) = tmpctr_cand2211(0)
                    Case Else
                        class_predict_fold1(X) = tmpctr_cand2211(1)
                End Select
            End If
            
            If (tmpctr = 3 Or tmpctr = 1) Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 3 Then
                        class_predict_fold1(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 4 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 2 Then
                        class_predict_fold1(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 6 Then
                For Y = 1 To 10
                    If class_count_fold1(Y) = 1 Then
                        If tmpctr_cand111111(0) = "0" Then
                            tmpctr_cand111111(0) = class_name(Y)
                        ElseIf tmpctr_cand111111(1) = "0" Then
                            tmpctr_cand111111(1) = class_name(Y)
                        ElseIf tmpctr_cand111111(2) = "0" Then
                            tmpctr_cand111111(2) = class_name(Y)
                        ElseIf tmpctr_cand111111(3) = "0" Then
                            tmpctr_cand111111(3) = class_name(Y)
                        ElseIf tmpctr_cand111111(4) = "0" Then
                            tmpctr_cand111111(4) = class_name(Y)
                        ElseIf tmpctr_cand111111(5) = "0" Then
                            tmpctr_cand111111(5) = class_name(Y)
                        End If
                    End If
                Next
                mode = X Mod 6
                Select Case mode
                    Case 1
                        class_predict_fold1(X) = tmpctr_cand111111(0)
                    Case 2
                        class_predict_fold1(X) = tmpctr_cand111111(1)
                    Case 3
                        class_predict_fold1(X) = tmpctr_cand111111(2)
                    Case 4
                        class_predict_fold1(X) = tmpctr_cand111111(3)
                    Case 5
                        class_predict_fold1(X) = tmpctr_cand111111(4)
                    Case Else
                        class_predict_fold1(X) = tmpctr_cand111111(5)
                End Select
            End If
        End If
    Next
    For X = 0 To 295
        If class_predict_fold1(X) = fold_1(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(0) = (ctr_accuracy / 296) * 100
    List1.AddItem "i_th fold" & vbTab & "#data" & vbTab & "#accurate data" & vbTab & "  accuracy"
    List1.AddItem "1st  fold" & vbTab & "296" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(0), 6) & "%"
    
'    fold 2-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        class_6(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        min_6(X) = 10000
        For Y = 0 To 1186
            If distance_2(X, Y) <= min_1(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_2(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_1(X) And distance_2(X, Y) <= min_2(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_2(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_2(X) And distance_2(X, Y) <= min_3(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_2(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_3(X) And distance_2(X, Y) <= min_4(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = distance_2(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_4(X) And distance_2(X, Y) <= min_5(X) Then
                min_6(X) = min_5(X)
                min_5(X) = distance_2(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_predict_candidate_2(X, Y)
            ElseIf distance_2(X, Y) >= min_5(X) And distance_2(X, Y) <= min_6(X) Then
                min_6(X) = distance_2(X, Y)
                class_6(X) = class_predict_candidate_2(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold2(297) As String

   rnd_class = 0
   ctr_accuracy = 0

    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold2(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        Select Case class_6(X)
            Case "CYT"
                class_count_fold2(1) = class_count_fold2(1) + 1
            Case "NUC"
                class_count_fold2(2) = class_count_fold2(2) + 1
            Case "MIT"
                class_count_fold2(3) = class_count_fold2(3) + 1
            Case "ME3"
                class_count_fold2(4) = class_count_fold2(4) + 1
            Case "ME2"
                class_count_fold2(5) = class_count_fold2(5) + 1
            Case "ME1"
                class_count_fold2(6) = class_count_fold2(6) + 1
            Case "EXC"
                class_count_fold2(7) = class_count_fold2(7) + 1
            Case "VAC"
                class_count_fold2(8) = class_count_fold2(8) + 1
            Case "POX"
                class_count_fold2(9) = class_count_fold2(9) + 1
            Case Else
                class_count_fold2(10) = class_count_fold2(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold2(Y) = 6 Then
                class_predict_fold2(X) = class_name(Y)
            ElseIf class_count_fold2(Y) = 5 Then
                class_predict_fold2(X) = class_name(Y)
            ElseIf class_count_fold2(Y) = 4 Then
                class_predict_fold2(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand33(0) = 0
        tmpctr_cand33(1) = 0
        tmpctr_cand2211(0) = 0
        tmpctr_cand2211(1) = 0
        For Y = 0 To 2
            tmpctr_cand222(Y) = 0
        Next
        For Y = 0 To 5
            tmpctr_cand111111(Y) = 0
        Next
        
        If class_predict_fold2(X) = "" Then
            For Y = 1 To 10
                If class_count_fold2(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            
            If tmpctr = 0 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 3 Then '33的情況
                        If tmpctr_cand33(0) = "0" Then
                                tmpctr_cand33(0) = class_name(Y)
                            Else: tmpctr_cand33(1) = class_name(Y)
                        End If
                    ElseIf class_count_fold2(Y) = 2 Then '222的情況
                        If tmpctr_cand222(0) = "0" Then
                                tmpctr_cand222(0) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(1) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(2) = class_name(Y)
                        End If
                    End If
                Next
                For Y = 1 To 10
                    If class_count_fold2(Y) = 3 Then
                        mode2 = X Mod 2
                        Select Case mode2
                            Case 1
                                class_predict_fold2(X) = tmpctr_cand33(0)
                            Case Else
                                class_predict_fold2(X) = tmpctr_cand33(1)
                        End Select
                    End If
                    If class_count_fold2(Y) = 2 Then
                        mode2 = X Mod 3
                        Select Case mode2
                            Case 1
                                class_predict_fold2(X) = tmpctr_cand222(0)
                            Case 2
                                class_predict_fold2(X) = tmpctr_cand222(1)
                            Case Else
                                class_predict_fold2(X) = tmpctr_cand222(2)
                        End Select
                    End If
                Next
            End If

            If tmpctr = 2 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 2 Then
                        If tmpctr_cand2211(0) = "0" Then
                            tmpctr_cand2211(0) = class_name(Y)
                        ElseIf tmpctr_cand2211(0) <> "0" Then
                            tmpctr_cand2211(1) = class_name(Y)
                        End If
                    End If
                Next

                mode2 = X Mod 2
                Select Case mode2
                    Case 1
                        class_predict_fold2(X) = tmpctr_cand2211(0)
                    Case Else
                        class_predict_fold2(X) = tmpctr_cand2211(1)
                End Select
            End If
            
            If (tmpctr = 3 Or tmpctr = 1) Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 3 Then
                        class_predict_fold2(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 4 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 2 Then
                        class_predict_fold2(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 6 Then
                For Y = 1 To 10
                    If class_count_fold2(Y) = 1 Then
                        If tmpctr_cand111111(0) = "0" Then
                            tmpctr_cand111111(0) = class_name(Y)
                        ElseIf tmpctr_cand111111(1) = "0" Then
                            tmpctr_cand111111(1) = class_name(Y)
                        ElseIf tmpctr_cand111111(2) = "0" Then
                            tmpctr_cand111111(2) = class_name(Y)
                        ElseIf tmpctr_cand111111(3) = "0" Then
                            tmpctr_cand111111(3) = class_name(Y)
                        ElseIf tmpctr_cand111111(4) = "0" Then
                            tmpctr_cand111111(4) = class_name(Y)
                        ElseIf tmpctr_cand111111(5) = "0" Then
                            tmpctr_cand111111(5) = class_name(Y)
                        End If
                    End If
                Next
                mode2 = X Mod 6
                Select Case mode2
                    Case 1
                        class_predict_fold2(X) = tmpctr_cand111111(0)
                    Case 2
                        class_predict_fold2(X) = tmpctr_cand111111(1)
                    Case 3
                        class_predict_fold2(X) = tmpctr_cand111111(2)
                    Case 4
                        class_predict_fold2(X) = tmpctr_cand111111(3)
                    Case 5
                        class_predict_fold2(X) = tmpctr_cand111111(4)
                    Case Else
                        class_predict_fold2(X) = tmpctr_cand111111(5)
                End Select
            End If
        End If
    Next
    For X = 0 To 296
        If class_predict_fold2(X) = fold_2(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(1) = (ctr_accuracy / 297) * 100
    List1.AddItem "2nd  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(1), 6) & "%"
'    fold 3-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        class_6(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        min_6(X) = 10000
        For Y = 0 To 1186
            If distance_3(X, Y) <= min_1(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_3(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_1(X) And distance_3(X, Y) <= min_2(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_3(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_2(X) And distance_3(X, Y) <= min_3(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_3(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_3(X) And distance_3(X, Y) <= min_4(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = distance_3(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_4(X) And distance_3(X, Y) <= min_5(X) Then
                min_6(X) = min_5(X)
                min_5(X) = distance_3(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_predict_candidate_3(X, Y)
            ElseIf distance_3(X, Y) >= min_5(X) And distance_3(X, Y) <= min_6(X) Then
                min_6(X) = distance_3(X, Y)
                class_6(X) = class_predict_candidate_3(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold3(297) As String

   rnd_class = 0
   ctr_accuracy = 0

    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold3(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        Select Case class_6(X)
            Case "CYT"
                class_count_fold3(1) = class_count_fold3(1) + 1
            Case "NUC"
                class_count_fold3(2) = class_count_fold3(2) + 1
            Case "MIT"
                class_count_fold3(3) = class_count_fold3(3) + 1
            Case "ME3"
                class_count_fold3(4) = class_count_fold3(4) + 1
            Case "ME2"
                class_count_fold3(5) = class_count_fold3(5) + 1
            Case "ME1"
                class_count_fold3(6) = class_count_fold3(6) + 1
            Case "EXC"
                class_count_fold3(7) = class_count_fold3(7) + 1
            Case "VAC"
                class_count_fold3(8) = class_count_fold3(8) + 1
            Case "POX"
                class_count_fold3(9) = class_count_fold3(9) + 1
            Case Else
                class_count_fold3(10) = class_count_fold3(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold3(Y) = 6 Then
                class_predict_fold3(X) = class_name(Y)
            ElseIf class_count_fold3(Y) = 5 Then
                class_predict_fold3(X) = class_name(Y)
            ElseIf class_count_fold3(Y) = 4 Then
                class_predict_fold3(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand33(0) = 0
        tmpctr_cand33(1) = 0
        tmpctr_cand2211(0) = 0
        tmpctr_cand2211(1) = 0
        For Y = 0 To 2
            tmpctr_cand222(Y) = 0
        Next
        For Y = 0 To 5
            tmpctr_cand111111(Y) = 0
        Next
        
        If class_predict_fold3(X) = "" Then
            For Y = 1 To 10
                If class_count_fold3(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            
            If tmpctr = 0 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 3 Then '33的情況
                        If tmpctr_cand33(0) = "0" Then
                                tmpctr_cand33(0) = class_name(Y)
                            Else: tmpctr_cand33(1) = class_name(Y)
                        End If
                    ElseIf class_count_fold3(Y) = 2 Then '222的情況
                        If tmpctr_cand222(0) = "0" Then
                                tmpctr_cand222(0) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(1) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(2) = class_name(Y)
                        End If
                    End If
                Next
                For Y = 1 To 10
                    If class_count_fold3(Y) = 3 Then
                        mode3 = X Mod 2
                        Select Case mode3
                            Case 1
                                class_predict_fold3(X) = tmpctr_cand33(0)
                            Case Else
                                class_predict_fold3(X) = tmpctr_cand33(1)
                        End Select
                    End If
                    If class_count_fold3(Y) = 2 Then
                        mode3 = X Mod 3
                        Select Case mode3
                            Case 1
                                class_predict_fold3(X) = tmpctr_cand222(0)
                            Case 2
                                class_predict_fold3(X) = tmpctr_cand222(1)
                            Case Else
                                class_predict_fold3(X) = tmpctr_cand222(2)
                        End Select
                    End If
                Next
            End If

            If tmpctr = 2 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 2 Then
                        If tmpctr_cand2211(0) = "0" Then
                            tmpctr_cand2211(0) = class_name(Y)
                        ElseIf tmpctr_cand2211(0) <> "0" Then
                            tmpctr_cand2211(1) = class_name(Y)
                        End If
                    End If
                Next

                mode3 = X Mod 2
                Select Case mode3
                    Case 1
                        class_predict_fold3(X) = tmpctr_cand2211(0)
                    Case Else
                        class_predict_fold3(X) = tmpctr_cand2211(1)
                End Select
            End If
            
            If (tmpctr = 3 Or tmpctr = 1) Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 3 Then
                        class_predict_fold3(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 4 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 2 Then
                        class_predict_fold3(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 6 Then
                For Y = 1 To 10
                    If class_count_fold3(Y) = 1 Then
                        If tmpctr_cand111111(0) = "0" Then
                            tmpctr_cand111111(0) = class_name(Y)
                        ElseIf tmpctr_cand111111(1) = "0" Then
                            tmpctr_cand111111(1) = class_name(Y)
                        ElseIf tmpctr_cand111111(2) = "0" Then
                            tmpctr_cand111111(2) = class_name(Y)
                        ElseIf tmpctr_cand111111(3) = "0" Then
                            tmpctr_cand111111(3) = class_name(Y)
                        ElseIf tmpctr_cand111111(4) = "0" Then
                            tmpctr_cand111111(4) = class_name(Y)
                        ElseIf tmpctr_cand111111(5) = "0" Then
                            tmpctr_cand111111(5) = class_name(Y)
                        End If
                    End If
                Next
                mode3 = X Mod 6
                Select Case mode3
                    Case 1
                        class_predict_fold3(X) = tmpctr_cand111111(0)
                    Case 2
                        class_predict_fold3(X) = tmpctr_cand111111(1)
                    Case 3
                        class_predict_fold3(X) = tmpctr_cand111111(2)
                    Case 4
                        class_predict_fold3(X) = tmpctr_cand111111(3)
                    Case 5
                        class_predict_fold3(X) = tmpctr_cand111111(4)
                    Case Else
                        class_predict_fold3(X) = tmpctr_cand111111(5)
                End Select
            End If
        End If
    Next
    For X = 0 To 296
        If class_predict_fold3(X) = fold_3(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(2) = (ctr_accuracy / 297) * 100
    List1.AddItem "3rd  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(2), 6) & "%"

'    fold 4-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        class_6(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        min_6(X) = 10000
        For Y = 0 To 1186
            If distance_4(X, Y) <= min_1(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_4(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_1(X) And distance_4(X, Y) <= min_2(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_4(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_2(X) And distance_4(X, Y) <= min_3(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_4(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_3(X) And distance_4(X, Y) <= min_4(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = distance_4(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_4(X) And distance_4(X, Y) <= min_5(X) Then
                min_6(X) = min_5(X)
                min_5(X) = distance_4(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_predict_candidate_4(X, Y)
            ElseIf distance_4(X, Y) >= min_5(X) And distance_4(X, Y) <= min_6(X) Then
                min_6(X) = distance_4(X, Y)
                class_6(X) = class_predict_candidate_4(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold4(297) As String

   rnd_class = 0
   ctr_accuracy = 0

    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold4(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        Select Case class_6(X)
            Case "CYT"
                class_count_fold4(1) = class_count_fold4(1) + 1
            Case "NUC"
                class_count_fold4(2) = class_count_fold4(2) + 1
            Case "MIT"
                class_count_fold4(3) = class_count_fold4(3) + 1
            Case "ME3"
                class_count_fold4(4) = class_count_fold4(4) + 1
            Case "ME2"
                class_count_fold4(5) = class_count_fold4(5) + 1
            Case "ME1"
                class_count_fold4(6) = class_count_fold4(6) + 1
            Case "EXC"
                class_count_fold4(7) = class_count_fold4(7) + 1
            Case "VAC"
                class_count_fold4(8) = class_count_fold4(8) + 1
            Case "POX"
                class_count_fold4(9) = class_count_fold4(9) + 1
            Case Else
                class_count_fold4(10) = class_count_fold4(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold4(Y) = 6 Then
                class_predict_fold4(X) = class_name(Y)
            ElseIf class_count_fold4(Y) = 5 Then
                class_predict_fold4(X) = class_name(Y)
            ElseIf class_count_fold4(Y) = 4 Then
                class_predict_fold4(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand33(0) = 0
        tmpctr_cand33(1) = 0
        tmpctr_cand2211(0) = 0
        tmpctr_cand2211(1) = 0
        For Y = 0 To 2
            tmpctr_cand222(Y) = 0
        Next
        For Y = 0 To 5
            tmpctr_cand111111(Y) = 0
        Next
        
        If class_predict_fold4(X) = "" Then
            For Y = 1 To 10
                If class_count_fold4(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            
            If tmpctr = 0 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 3 Then '33的情況
                        If tmpctr_cand33(0) = "0" Then
                                tmpctr_cand33(0) = class_name(Y)
                            Else: tmpctr_cand33(1) = class_name(Y)
                        End If
                    ElseIf class_count_fold4(Y) = 2 Then '222的情況
                        If tmpctr_cand222(0) = "0" Then
                                tmpctr_cand222(0) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(1) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(2) = class_name(Y)
                        End If
                    End If
                Next
                For Y = 1 To 10
                    If class_count_fold4(Y) = 3 Then
                        mode4 = X Mod 2
                        Select Case mode4
                            Case 1
                                class_predict_fold4(X) = tmpctr_cand33(0)
                            Case Else
                                class_predict_fold4(X) = tmpctr_cand33(1)
                        End Select
                    End If
                    If class_count_fold4(Y) = 2 Then
                        mode4 = X Mod 3
                        Select Case mode4
                            Case 1
                                class_predict_fold4(X) = tmpctr_cand222(0)
                            Case 2
                                class_predict_fold4(X) = tmpctr_cand222(1)
                            Case Else
                                class_predict_fold4(X) = tmpctr_cand222(2)
                        End Select
                    End If
                Next
            End If

            If tmpctr = 2 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 2 Then
                        If tmpctr_cand2211(0) = "0" Then
                            tmpctr_cand2211(0) = class_name(Y)
                        ElseIf tmpctr_cand2211(0) <> "0" Then
                            tmpctr_cand2211(1) = class_name(Y)
                        End If
                    End If
                Next

                mode4 = X Mod 2
                Select Case mode4
                    Case 1
                        class_predict_fold4(X) = tmpctr_cand2211(0)
                    Case Else
                        class_predict_fold4(X) = tmpctr_cand2211(1)
                End Select
            End If
            
            If (tmpctr = 3 Or tmpctr = 1) Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 3 Then
                        class_predict_fold4(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 4 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 2 Then
                        class_predict_fold4(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 6 Then
                For Y = 1 To 10
                    If class_count_fold4(Y) = 1 Then
                        If tmpctr_cand111111(0) = "0" Then
                            tmpctr_cand111111(0) = class_name(Y)
                        ElseIf tmpctr_cand111111(1) = "0" Then
                            tmpctr_cand111111(1) = class_name(Y)
                        ElseIf tmpctr_cand111111(2) = "0" Then
                            tmpctr_cand111111(2) = class_name(Y)
                        ElseIf tmpctr_cand111111(3) = "0" Then
                            tmpctr_cand111111(3) = class_name(Y)
                        ElseIf tmpctr_cand111111(4) = "0" Then
                            tmpctr_cand111111(4) = class_name(Y)
                        ElseIf tmpctr_cand111111(5) = "0" Then
                            tmpctr_cand111111(5) = class_name(Y)
                        End If
                    End If
                Next
                mode4 = X Mod 6
                Select Case mode4
                    Case 1
                        class_predict_fold4(X) = tmpctr_cand111111(0)
                    Case 2
                        class_predict_fold4(X) = tmpctr_cand111111(1)
                    Case 3
                        class_predict_fold4(X) = tmpctr_cand111111(2)
                    Case 4
                        class_predict_fold4(X) = tmpctr_cand111111(3)
                    Case 5
                        class_predict_fold4(X) = tmpctr_cand111111(4)
                    Case Else
                        class_predict_fold4(X) = tmpctr_cand111111(5)
                End Select
            End If
        End If
    Next
    For X = 0 To 296
        If class_predict_fold4(X) = fold_4(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(3) = (ctr_accuracy / 297) * 100
    List1.AddItem "4th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(3), 6) & "%"

    '    fold 5-----------------------------------------------
    For X = 0 To 296
        class_1(X) = "10000"
        class_2(X) = "10000"
        class_3(X) = "10000"
        class_4(X) = "10000"
        class_5(X) = "10000"
        class_6(X) = "10000"
        min_1(X) = 10000
        min_2(X) = 10000
        min_3(X) = 10000
        min_4(X) = 10000
        min_5(X) = 10000
        min_6(X) = 10000
        For Y = 0 To 1186
            If distance_5(X, Y) <= min_1(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = min_1(X)
                min_1(X) = distance_5(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_1(X)
                class_1(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_1(X) And distance_5(X, Y) <= min_2(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = min_2(X)
                min_2(X) = distance_5(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_2(X)
                class_2(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_2(X) And distance_5(X, Y) <= min_3(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = min_3(X)
                min_3(X) = distance_5(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_3(X)
                class_3(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_3(X) And distance_5(X, Y) <= min_4(X) Then
                min_6(X) = min_5(X)
                min_5(X) = min_4(X)
                min_4(X) = distance_5(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_4(X)
                class_4(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_4(X) And distance_5(X, Y) <= min_5(X) Then
                min_6(X) = min_5(X)
                min_5(X) = distance_5(X, Y)
                class_6(X) = class_5(X)
                class_5(X) = class_predict_candidate_5(X, Y)
            ElseIf distance_5(X, Y) >= min_5(X) And distance_5(X, Y) <= min_6(X) Then
                min_6(X) = distance_5(X, Y)
                class_6(X) = class_predict_candidate_5(X, Y)
            End If
        Next Y
    Next X
    
    Dim class_predict_fold5(297) As String

   rnd_class = 0
   ctr_accuracy = 0

    For X = 0 To 296
        For Y = 1 To 10
            class_count_fold5(Y) = 0
        Next
        Select Case class_1(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_2(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_3(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_4(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_5(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        Select Case class_6(X)
            Case "CYT"
                class_count_fold5(1) = class_count_fold5(1) + 1
            Case "NUC"
                class_count_fold5(2) = class_count_fold5(2) + 1
            Case "MIT"
                class_count_fold5(3) = class_count_fold5(3) + 1
            Case "ME3"
                class_count_fold5(4) = class_count_fold5(4) + 1
            Case "ME2"
                class_count_fold5(5) = class_count_fold5(5) + 1
            Case "ME1"
                class_count_fold5(6) = class_count_fold5(6) + 1
            Case "EXC"
                class_count_fold5(7) = class_count_fold5(7) + 1
            Case "VAC"
                class_count_fold5(8) = class_count_fold5(8) + 1
            Case "POX"
                class_count_fold5(9) = class_count_fold5(9) + 1
            Case Else
                class_count_fold5(10) = class_count_fold5(10) + 1
        End Select
        
        For Y = 1 To 10
            If class_count_fold5(Y) = 6 Then
                class_predict_fold5(X) = class_name(Y)
            ElseIf class_count_fold5(Y) = 5 Then
                class_predict_fold5(X) = class_name(Y)
            ElseIf class_count_fold5(Y) = 4 Then
                class_predict_fold5(X) = class_name(Y)
            End If
        Next
        
        tmpctr = 0
        tmpctr_cand33(0) = 0
        tmpctr_cand33(1) = 0
        tmpctr_cand2211(0) = 0
        tmpctr_cand2211(1) = 0
        For Y = 0 To 2
            tmpctr_cand222(Y) = 0
        Next
        For Y = 0 To 5
            tmpctr_cand111111(Y) = 0
        Next
        
        If class_predict_fold5(X) = "" Then
            For Y = 1 To 10
                If class_count_fold5(Y) = 1 Then
                    tmpctr = tmpctr + 1
                End If
            Next
            
            If tmpctr = 0 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 3 Then '33的情況
                        If tmpctr_cand33(0) = "0" Then
                                tmpctr_cand33(0) = class_name(Y)
                            Else: tmpctr_cand33(1) = class_name(Y)
                        End If
                    ElseIf class_count_fold5(Y) = 2 Then '222的情況
                        If tmpctr_cand222(0) = "0" Then
                                tmpctr_cand222(0) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(1) = class_name(Y)
                            ElseIf tmpctr_cand222(0) <> "0" Then
                                tmpctr_cand222(2) = class_name(Y)
                        End If
                    End If
                Next
                For Y = 1 To 10
                    If class_count_fold5(Y) = 3 Then
                        mode5 = X Mod 2
                        Select Case mode5
                            Case 1
                                class_predict_fold5(X) = tmpctr_cand33(0)
                            Case Else
                                class_predict_fold5(X) = tmpctr_cand33(1)
                        End Select
                    End If
                    If class_count_fold5(Y) = 2 Then
                        mode5 = X Mod 3
                        Select Case mode5
                            Case 1
                                class_predict_fold5(X) = tmpctr_cand222(0)
                            Case 2
                                class_predict_fold5(X) = tmpctr_cand222(1)
                            Case Else
                                class_predict_fold5(X) = tmpctr_cand222(2)
                        End Select
                    End If
                Next
            End If

            If tmpctr = 2 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 2 Then
                        If tmpctr_cand2211(0) = "0" Then
                            tmpctr_cand2211(0) = class_name(Y)
                        ElseIf tmpctr_cand2211(0) <> "0" Then
                            tmpctr_cand2211(1) = class_name(Y)
                        End If
                    End If
                Next

                mode5 = X Mod 2
                Select Case mode5
                    Case 1
                        class_predict_fold5(X) = tmpctr_cand2211(0)
                    Case Else
                        class_predict_fold5(X) = tmpctr_cand2211(1)
                End Select
            End If
            
            If (tmpctr = 3 Or tmpctr = 1) Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 3 Then
                        class_predict_fold5(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 4 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 2 Then
                        class_predict_fold5(X) = class_name(Y)
                    End If
                Next
            End If
            
            If tmpctr = 6 Then
                For Y = 1 To 10
                    If class_count_fold5(Y) = 1 Then
                        If tmpctr_cand111111(0) = "0" Then
                            tmpctr_cand111111(0) = class_name(Y)
                        ElseIf tmpctr_cand111111(1) = "0" Then
                            tmpctr_cand111111(1) = class_name(Y)
                        ElseIf tmpctr_cand111111(2) = "0" Then
                            tmpctr_cand111111(2) = class_name(Y)
                        ElseIf tmpctr_cand111111(3) = "0" Then
                            tmpctr_cand111111(3) = class_name(Y)
                        ElseIf tmpctr_cand111111(4) = "0" Then
                            tmpctr_cand111111(4) = class_name(Y)
                        ElseIf tmpctr_cand111111(5) = "0" Then
                            tmpctr_cand111111(5) = class_name(Y)
                        End If
                    End If
                Next
                mode5 = X Mod 6
                Select Case mode5
                    Case 1
                        class_predict_fold5(X) = tmpctr_cand111111(0)
                    Case 2
                        class_predict_fold5(X) = tmpctr_cand111111(1)
                    Case 3
                        class_predict_fold5(X) = tmpctr_cand111111(2)
                    Case 4
                        class_predict_fold5(X) = tmpctr_cand111111(3)
                    Case 5
                        class_predict_fold5(X) = tmpctr_cand111111(4)
                    Case Else
                        class_predict_fold5(X) = tmpctr_cand111111(5)
                End Select
            End If
        End If
    Next
    For X = 0 To 296
        If class_predict_fold5(X) = fold_5(X, 9) Then
        ctr_accuracy = ctr_accuracy + 1
        End If
    Next
    accuracy(4) = (ctr_accuracy / 297) * 100
    List1.AddItem "5th  fold" & vbTab & "297" & vbTab & ctr_accuracy & vbTab & vbTab & "  " & FormatNumber(accuracy(4), 6) & "%"
    List1.AddItem "-----------------------------------------------------"
    List1.AddItem "average accuracy: " & FormatNumber(((accuracy(0) + accuracy(1) + accuracy(2) + accuracy(3) + accuracy(4)) / 5), 6) & "%"
End Sub
