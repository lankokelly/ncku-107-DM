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
   Begin VB.CommandButton backward 
      Caption         =   "Backward"
      Height          =   1095
      Left            =   10440
      TabIndex        =   10
      Top             =   1800
      Width           =   1920
   End
   Begin VB.CommandButton forward 
      Caption         =   "Forward"
      Height          =   1095
      Left            =   8040
      TabIndex        =   9
      Top             =   1800
      Width           =   2040
   End
   Begin VB.CommandButton EntropyBase 
      Caption         =   "Entropy-Based"
      Height          =   1095
      Left            =   4800
      TabIndex        =   8
      Top             =   1800
      Width           =   1915
   End
   Begin VB.CommandButton EqualFrequency 
      Caption         =   "Equal-Frequency discretization"
      Height          =   1095
      Left            =   2640
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton EqualWidth 
      Caption         =   "Equal-Width discretization"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
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
      Height          =   5475
      Left            =   480
      TabIndex        =   4
      Top             =   3600
      Width           =   12375
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
      Left            =   1920
      TabIndex        =   1
      Text            =   "yeast.txt"
      Top             =   330
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
      Left            =   4080
      TabIndex        =   0
      Top             =   340
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "-  按下特徵選取方法後，如要重新選擇離散化手法或特徵選取方式，請老師關掉視窗重新run"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   720
      Width           =   7095
   End
   Begin VB.Label Label4 
      Caption         =   "-  按Read後，選擇離散化方法，再選擇特徵選取方法"
      BeginProperty Font 
         Name            =   "微軟正黑體"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   5295
   End
   Begin VB.Line Line1 
      X1              =   7440
      X2              =   7440
      Y1              =   1320
      Y2              =   2880
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
      Left            =   8040
      TabIndex        =   11
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Discretization Method : "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   "Data"
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
      Top             =   3120
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
      Left            =   240
      TabIndex        =   2
      Top             =   360
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
Dim class As String
Dim X As Integer
Dim Y As Integer
Dim col As Integer

Dim data(1484, 10) As String  'data1 used in equal-width,changing value after discretization
Dim data2(1484, 10) As String 'data2 is a sorted matrix, used in equal-frequency
Dim data3(1484, 10) As String 'data3 used in equal-frequency, store value after discretization
Dim data4(1484, 10) As String 'data4 used in entropy-based

Dim frequency_cutpoint(8, 9) As String '(attribute,splitting point)
Dim max(8) As Double 'save 8 attribute max value
Dim min(8) As Double 'save 8 attribute min value
Dim equal_width(8) As Double 'save interval of each attribute
Dim interval As Integer 'interval is 10
Dim frequency As Integer 'frequency is 148

Dim freq_probability(11, 10) As String
Dim width_probability(11, 10) As Double
Dim pabmatrix(11, 11, 10, 10) As Double
Dim habmatrix(10, 10) As Double
Dim uabmatrix(10, 10) As Double
Dim h_value_att(10) As Double 'save each attribute's h_value

Dim selectset(9) As Integer



Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Load()

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
            'Open App.Path & "\test.txt" For Output As #2
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
                        data2(Y, col) = each_row_output(X)
                        data3(Y, col) = each_row_output(X)
                        data4(Y, col) = each_row_output(X)
                    col = col + 1
                    End If
                Next
                Y = Y + 1 'goto next row
            Loop
            
            MsgBox "Input file has prepared successfully !"
            MsgBox "please choose discretization method and feature selection direction!"
            
            Close #1
        End If
    End If
End Sub

Private Sub EqualWidth_click()
    List1.Clear
    interval = 10
    
    'count width of each attribute
    For col = 1 To 8
        min(col) = 1
            For X = 0 To 1483
                If CDbl(data(X, col)) > max(col) Then
                    max(col) = CDbl(data(X, col))
                ElseIf CDbl(data(X, col)) < min(col) Then
                    min(col) = CDbl(data(X, col))
                End If
            Next
        equal_width(col) = (max(col) - min(col)) / interval
    Next
    
    'print splitting point
    For col = 1 To 8
        List1.AddItem "Attribute : " & col & " " & "    Width= " & equal_width(col)
        For X = 1 To (interval - 1)
            List1.AddItem "splitting point = " & min(col) + X * equal_width(col)
        Next
        List1.AddItem ""
    Next
    
    'discretization to 1-10, each interval includes splitting point(lower bound)
    For col = 1 To 8
        For X = 0 To 1483
            If CDbl(data(X, col)) >= (min(col) + (interval - 1) * equal_width(col)) Then 'the 10th interval
                data(X, col) = 10
            ElseIf CDbl(data(X, col)) <> -1 Then
                For Y = 1 To interval - 1 'the 1th to 9th
                    If data(X, col) < (min(col) + Y * equal_width(col)) Then
                        data(X, col) = Y
                    End If
                Next
            End If
        Next
    Next
    
    'Print
    'For col = 1 To 8
        'For x = 0 To 1483
            'Print #2, data(x, col)
        'Next
    'Next
    
    'initialize the matrix
    For col = 1 To 9
        For X = 1 To 10
            width_probability(X, col) = 0
        Next
    Next
    
    
    'count the total appear times of each attribute's discrete value
    For col = 1 To 8
        For X = 0 To 1483
            Select Case CDbl(data(X, col))
                Case 1
                    width_probability(1, col) = width_probability(1, col) + 1
                Case 2
                    width_probability(2, col) = width_probability(2, col) + 1
                Case 3
                    width_probability(3, col) = width_probability(3, col) + 1
                Case 4
                    width_probability(4, col) = width_probability(4, col) + 1
                Case 5
                    width_probability(5, col) = width_probability(5, col) + 1
                Case 6
                    width_probability(6, col) = width_probability(6, col) + 1
                Case 7
                    width_probability(7, col) = width_probability(7, col) + 1
                Case 8
                    width_probability(8, col) = width_probability(8, col) + 1
                Case 9
                    width_probability(9, col) = width_probability(9, col) + 1
                Case Else
                    width_probability(10, col) = width_probability(10, col) + 1
            End Select
        Next
    Next
    
    'count class's
    For X = 0 To 1483
        Select Case data(X, 9)
            Case "CYT"
                width_probability(1, 9) = width_probability(1, 9) + 1
                data(X, 9) = 1
            Case "NUC"
                width_probability(2, 9) = width_probability(2, 9) + 1
                data(X, 9) = 2
            Case "MIT"
                width_probability(3, 9) = width_probability(3, 9) + 1
                data(X, 9) = 3
            Case "ME3"
                width_probability(4, 9) = width_probability(4, 9) + 1
                data(X, 9) = 4
            Case "ME2"
                width_probability(5, 9) = width_probability(5, 9) + 1
                data(X, 9) = 5
            Case "ME1"
                width_probability(6, 9) = width_probability(6, 9) + 1
                data(X, 9) = 6
            Case "EXC"
                width_probability(7, 9) = width_probability(7, 9) + 1
                data(X, 9) = 7
            Case "VAC"
                width_probability(8, 9) = width_probability(8, 9) + 1
                data(X, 9) = 8
            Case "POX"
                width_probability(9, 9) = width_probability(9, 9) + 1
                data(X, 9) = 9
            Case Else
                width_probability(10, 9) = width_probability(10, 9) + 1
                data(X, 9) = 10
        End Select
    Next
    
    For col = 1 To 9
        For X = 1 To 10
            width_probability(X, col) = width_probability(X, col) / 1484
        Next
    Next
    
    For X = 1 To 9
        h_value_att(X) = h_value(width_probability, X)
    Next
    
    probability_ab_matrix (data)
    
    Dim a As Integer
    Dim b As Integer
    Dim i As Integer
    Dim j As Integer
    
    For a = 1 To 9
        For b = 1 To 9
            For i = 1 To 10
                For j = 1 To 10
                        pabmatrix(i, j, a, b) = pabmatrix(i, j, a, b) / 1484
                Next j
            Next i
        Next b
    Next a
    
    For a = 1 To 9
        For b = 1 To 9
            habmatrix(a, b) = h_ab_value(pabmatrix, a, b)
        Next
    Next
    
    For a = 1 To 9
        For b = 1 To 9
            uabmatrix(a, b) = u_ab_value(habmatrix, a, b)
        Next
    Next
    
End Sub

Private Sub EqualFrequency_click()
    'use bubble-sort to sort attribute and renew data2
    Dim Tag As Double
    List1.Clear
    For col = 1 To 8
        For X = 0 To 1483
            For Y = X To 1483
                If CDbl(data2(X, col)) > CDbl(data2(Y, col)) Then
                    'swift
                    Tag = data2(X, col)
                    data2(X, col) = data2(Y, col)
                    data2(Y, col) = Tag
                    'change class
                    class = data2(X, 0)
                    data2(X, 0) = data2(Y, 0)
                    data2(Y, 0) = class
                End If
            Next Y
        Next X
    Next col
            
    interval = 10
    frequency = CInt(1484 / interval) '148
    For col = 1 To 8
        For X = 1 To 9
            frequency_cutpoint(col, X) = (CDbl(data2(X * frequency - 1, col)) + CDbl(data2(X * frequency, col))) / 2
        Next
    Next

    For col = 1 To 8
        List1.AddItem "Attribute : " & col
        For X = 1 To interval - 1
            List1.AddItem "splitting point = " & frequency_cutpoint(col, X)
        Next
        List1.AddItem ""
    Next
    
    For col = 1 To 8
        For X = 0 To 1483
            If CDbl(data3(X, col)) >= frequency_cutpoint(col, interval - 1) Then 'the 10th interval
                data3(X, col) = 10
            ElseIf CDbl(data3(X, col)) <> -1 Then
                For Y = 1 To interval - 1 'the 1th to 9th
                    If data3(X, col) < frequency_cutpoint(col, Y) Then
                        data3(X, col) = Y
                    End If
                Next
            End If
        Next
    Next
    
    
    'initialize the matrix
    For col = 1 To 9
        For X = 1 To 10
            freq_probability(X, col) = 0
        Next
    Next
    
    
    
    'count the total appear times of each attribute's discrete value
    For col = 1 To 8
        For X = 0 To 1483
            Select Case CDbl(data3(X, col))
                Case 1
                    freq_probability(1, col) = freq_probability(1, col) + 1
                Case 2
                    freq_probability(2, col) = freq_probability(2, col) + 1
                Case 3
                    freq_probability(3, col) = freq_probability(3, col) + 1
                Case 4
                    freq_probability(4, col) = freq_probability(4, col) + 1
                Case 5
                    freq_probability(5, col) = freq_probability(5, col) + 1
                Case 6
                    freq_probability(6, col) = freq_probability(6, col) + 1
                Case 7
                    freq_probability(7, col) = freq_probability(7, col) + 1
                Case 8
                    freq_probability(8, col) = freq_probability(8, col) + 1
                Case 9
                    freq_probability(9, col) = freq_probability(9, col) + 1
                Case Else
                    freq_probability(10, col) = freq_probability(10, col) + 1
            End Select
        Next
    Next
    
    'discretize class(9th attribute)
    For X = 0 To 1483
        Select Case data3(X, 9)
            Case "CYT"
                freq_probability(1, 9) = freq_probability(1, 9) + 1
                data3(X, 9) = 1
            Case "NUC"
                freq_probability(2, 9) = freq_probability(2, 9) + 1
                data3(X, 9) = 2
            Case "MIT"
                freq_probability(3, 9) = freq_probability(3, 9) + 1
                data3(X, 9) = 3
            Case "ME3"
                freq_probability(4, 9) = freq_probability(4, 9) + 1
                data3(X, 9) = 4
            Case "ME2"
                freq_probability(5, 9) = freq_probability(5, 9) + 1
                data3(X, 9) = 5
            Case "ME1"
                freq_probability(6, 9) = freq_probability(6, 9) + 1
                data3(X, 9) = 6
            Case "EXC"
                freq_probability(7, 9) = freq_probability(7, 9) + 1
                data3(X, 9) = 7
            Case "VAC"
                freq_probability(8, 9) = freq_probability(8, 9) + 1
                data3(X, 9) = 8
            Case "POX"
                freq_probability(9, 9) = freq_probability(9, 9) + 1
                data3(X, 9) = 9
            Case Else
                freq_probability(10, 9) = freq_probability(10, 9) + 1
                data3(X, 9) = 10
        End Select
    Next
        
    For col = 1 To 9
        For X = 1 To 10
            freq_probability(X, col) = freq_probability(X, col) / 1484
        Next
    Next
    
    'print
    'For col = 5 To 5
        'For X = 1 To 10
            'Print #3, freq_probability(X, col)
        'Next
    'Next
    
    For X = 1 To 9
        h_value_att(X) = h_value(freq_probability, X)
    Next
    
    'calculate pab with function
    probability_ab_matrix (data3)
    'List1.AddItem pabmatrix(8, 3, 3, 9)
    Dim a As Integer
    Dim b As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    For a = 1 To 9
        For b = 1 To 9
            For i = 1 To 10
                For j = 1 To 10
                        pabmatrix(i, j, a, b) = pabmatrix(i, j, a, b) / 1484
                Next j
            Next i
        Next b
    Next a
    
    For a = 1 To 9
        For b = 1 To 9
            habmatrix(a, b) = h_ab_value(pabmatrix, a, b)
        Next
    Next
    
    For a = 1 To 9
        For b = 1 To 9
            uabmatrix(a, b) = u_ab_value(habmatrix, a, b)
        Next
    Next

End Sub

Private Sub EntropyBase_click()
    List1.Clear
    data
End Sub

Function h_value(prob, ByVal attr As Integer) As Double
    Dim temp As Double
    temp = 0
    Dim i As Integer
    Dim log_value As Double
    Dim prob_temp As Double
    
    For i = 1 To 10
        prob_temp = prob(i, attr)
        log_value = log2(prob_temp)
        temp = temp - prob(i, attr) * log_value
    Next
    h_value = temp
End Function

'pab = a & b appear at the same time, count it
'this func will calculate a
'           pab(10(# of discrete),10(# of discrete),9(# of attribute),9(# of attribute))

Function probability_ab_matrix(after_dis_mat)
    'variable declaration
    Dim a As Integer
    Dim b As Integer
    Dim k As Integer
    Dim i As Integer
    Dim j As Integer
    
    For k = 0 To 1483 'fix row
        For a = 1 To 9 'attribute(a)
            For b = 1 To 9 'attribute(b)
                For i = 1 To 10 'discrete(i)
                    For j = 1 To 10 'discrete(j)
                        'record each attribute(a)'s value = i and attribute(b)'s value = j
                        If after_dis_mat(k, a) = i And after_dis_mat(k, b) = j Then
                            pabmatrix(i, j, a, b) = pabmatrix(i, j, a, b) + 1 'counter++
                        End If
                    Next
                Next
            Next
        Next
    Next
End Function

Function h_ab_value(pabmat, ByVal attr1, ByVal attr2) As Double
    Dim i As Integer
    Dim j As Integer
    Dim hab_temp As Double
    Dim pab_temp As Double

    For i = 1 To 10
        For j = 1 To 10
            pab_temp = pabmat(i, j, attr1, attr2)
            hab_temp = hab_temp - pabmat(i, j, attr1, attr2) * log2(pab_temp)
        Next
    Next
    h_ab_value = hab_temp
End Function

Function u_ab_value(habmat, ByVal attr1, ByVal attr2) As Double
    Dim uab_temp As Double
    If attr1 = attr2 Then 'definition of u(a,a)=1
        u_ab_value = 1
    ElseIf h_value_att(attr1) = 0 And h_value_att(attr2) = 0 Then 'design for this dataset
        Exit Function
    Else
        uab_temp = 2 * (h_value_att(attr1) + h_value_att(attr2) - habmat(attr1, attr2)) / (h_value_att(attr1) + h_value_att(attr2))
        u_ab_value = uab_temp
    End If
End Function

Private Sub forward_click()
    List1.Clear
    Dim temp_goodness As Double
    Dim i As Integer
    
    temp_goodness = 0
    
    'attribute name
    Dim attribute_name(9) As String
    attribute_name(1) = "mcg"
    attribute_name(2) = "gvh"
    attribute_name(3) = "alm"
    attribute_name(4) = "mit"
    attribute_name(5) = "erl"
    attribute_name(6) = "pox"
    attribute_name(7) = "vac"
    attribute_name(8) = "nuc"
    
    'before start selection
    List1.AddItem "Select 0 Attritube: {}"
    List1.AddItem "Goodness : " & temp_goodness
    List1.AddItem ""
    
    'select one attritube
    Dim max_1_attr As Integer
    Dim max_1_goodvalue As Double
    max_1_goodvalue = 0
    max_1_attr = 0
    For i = 1 To 8
        temp_goodness = uabmatrix(i, 9)
        If temp_goodness > max_1_goodvalue Then
            max_1_goodvalue = temp_goodness
            max_1_attr = i
        End If
    Next
    
    If max_1_goodvalue > 0 Then
        List1.AddItem "Select 1 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & ") }"
        List1.AddItem "Goodness : " & max_1_goodvalue
        List1.AddItem ""
    End If
    
    'select two attritube
    Dim max_2_attr As Integer
    Dim max_2_goodvalue As Double
    Dim denominator As Double
    max_2_goodvalue = 0
    max_2_attr = 0
    For i = 1 To 8
        denominator = (uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, i) + uabmatrix(i, max_1_attr) + uabmatrix(i, i)) ^ 0.5
        temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(i, 9)) / denominator
        If temp_goodness > max_2_goodvalue Then
            max_2_goodvalue = temp_goodness
            max_2_attr = i
        End If
    Next
    
    If max_2_goodvalue > max_1_goodvalue Then
        List1.AddItem "Select 2 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & ") }"
        List1.AddItem "Goodness : " & max_2_goodvalue
        List1.AddItem ""
    End If
    
    'select three attritube
    Dim max_3_attr As Integer
    Dim max_3_goodvalue As Double
    max_3_goodvalue = 0
    max_3_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Then
            max_3_goodvalue = max_3_goodvalue + 0
        Else
            denominator = (uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, i) + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, i) + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, i)) ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_3_goodvalue Then
                    max_3_goodvalue = temp_goodness
                    max_3_attr = i
                End If
        End If
    Next
    
    If max_3_goodvalue > max_2_goodvalue Then
        List1.AddItem "Select 3 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ") }"
        List1.AddItem "Goodness : " & max_3_goodvalue
        List1.AddItem ""
    End If
    
    'select four attritube
    Dim max_4_attr As Integer
    Dim max_4_goodvalue As Double
    max_4_goodvalue = 0
    max_4_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Or (i = max_3_attr) Then
            max_4_goodvalue = max_4_goodvalue + 0
        Else
            denominator = (uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, max_3_attr) + uabmatrix(max_1_attr, i) + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, max_3_attr) + uabmatrix(max_2_attr, i) + uabmatrix(max_3_attr, max_1_attr) + uabmatrix(max_3_attr, max_2_attr) + uabmatrix(max_3_attr, max_3_attr) + uabmatrix(max_3_attr, i) + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, max_3_attr) + uabmatrix(i, i)) ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(max_3_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_4_goodvalue Then
                    max_4_goodvalue = temp_goodness
                    max_4_attr = i
                End If
        End If
    Next
    
    If max_4_goodvalue > max_3_goodvalue Then
        List1.AddItem "Select 4 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ")" & ",A" & max_4_attr & "(" & attribute_name(max_4_attr) & ") }"
        List1.AddItem "Goodness : " & max_4_goodvalue
        List1.AddItem ""
    End If
    
    'select five attritube
    Dim max_5_attr As Integer
    Dim max_5_goodvalue As Double
    max_5_goodvalue = 0
    max_5_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Or (i = max_3_attr) Or (i = max_4_attr) Then
            max_5_goodvalue = max_5_goodvalue + 0
        Else
            denominator = (uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, max_3_attr) + uabmatrix(max_1_attr, max_4_attr) + uabmatrix(max_1_attr, i) + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, max_3_attr) + uabmatrix(max_2_attr, max_4_attr) + uabmatrix(max_2_attr, i) + uabmatrix(max_3_attr, max_1_attr) + uabmatrix(max_3_attr, max_2_attr) + uabmatrix(max_3_attr, max_3_attr) + uabmatrix(max_3_attr, max_4_attr) + uabmatrix(max_3_attr, i) + uabmatrix(max_4_attr, max_1_attr) + uabmatrix(max_4_attr, max_2_attr) + uabmatrix(max_4_attr, max_3_attr) + uabmatrix(max_4_attr, max_4_attr) + uabmatrix(max_4_attr, i) + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, max_3_attr) + uabmatrix(i, max_4_attr) + uabmatrix(i, i)) ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(max_3_attr, 9) + uabmatrix(max_4_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_5_goodvalue Then
                    max_5_goodvalue = temp_goodness
                    max_5_attr = i
                End If
        End If
    Next
    
    If max_5_goodvalue > max_4_goodvalue Then
        List1.AddItem "Select 5 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ")" & ",A" & max_4_attr & "(" & attribute_name(max_4_attr) & ")" & ",A" & max_5_attr & "(" & attribute_name(max_5_attr) & ") }"
        List1.AddItem "Goodness : " & max_5_goodvalue
        List1.AddItem ""
    End If
    
    'select six attritube
    Dim max_6_attr As Integer
    Dim max_6_goodvalue As Double
    max_6_goodvalue = 0
    max_6_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Or (i = max_3_attr) Or (i = max_4_attr) Or (i = max_5_attr) Then
            max_6_goodvalue = max_6_goodvalue + 0
        Else
            denominator = uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, max_3_attr) + uabmatrix(max_1_attr, max_4_attr) + uabmatrix(max_1_attr, max_5_attr) + uabmatrix(max_1_attr, i)
            denominator = denominator + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, max_3_attr) + uabmatrix(max_2_attr, max_4_attr) + uabmatrix(max_2_attr, max_5_attr) + uabmatrix(max_2_attr, i)
            denominator = denominator + uabmatrix(max_3_attr, max_1_attr) + uabmatrix(max_3_attr, max_2_attr) + uabmatrix(max_3_attr, max_3_attr) + uabmatrix(max_3_attr, max_4_attr) + uabmatrix(max_3_attr, max_5_attr) + uabmatrix(max_3_attr, i)
            denominator = denominator + uabmatrix(max_4_attr, max_1_attr) + uabmatrix(max_4_attr, max_2_attr) + uabmatrix(max_4_attr, max_3_attr) + uabmatrix(max_4_attr, max_4_attr) + uabmatrix(max_4_attr, max_5_attr) + uabmatrix(max_4_attr, i)
            denominator = denominator + uabmatrix(max_5_attr, max_1_attr) + uabmatrix(max_5_attr, max_2_attr) + uabmatrix(max_5_attr, max_3_attr) + uabmatrix(max_5_attr, max_4_attr) + uabmatrix(max_5_attr, max_5_attr) + uabmatrix(max_5_attr, i)
            denominator = denominator + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, max_3_attr) + uabmatrix(i, max_4_attr) + uabmatrix(i, max_5_attr) + uabmatrix(i, i)
            denominator = denominator ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(max_3_attr, 9) + uabmatrix(max_4_attr, 9) + uabmatrix(max_5_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_6_goodvalue Then
                    max_6_goodvalue = temp_goodness
                    max_6_attr = i
                End If
        End If
    Next
    
    If max_6_goodvalue > max_5_goodvalue Then
        List1.AddItem "Select 6 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ")" & ",A" & max_4_attr & "(" & attribute_name(max_4_attr) & ")" & ",A" & max_5_attr & "(" & attribute_name(max_5_attr) & "),A" & max_6_attr & "(" & attribute_name(max_6_attr) & ") }"
        List1.AddItem "Goodness : " & max_6_goodvalue
        List1.AddItem ""
    End If
    
    'select seven attritube
    Dim max_7_attr As Integer
    Dim max_7_goodvalue As Double
    max_7_goodvalue = 0
    max_7_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Or (i = max_3_attr) Or (i = max_4_attr) Or (i = max_5_attr) Or (i = max_6_attr) Then
            max_7_goodvalue = max_7_goodvalue + 0
        Else
            denominator = uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, max_3_attr) + uabmatrix(max_1_attr, max_4_attr) + uabmatrix(max_1_attr, max_5_attr) + uabmatrix(max_1_attr, max_6_attr) + uabmatrix(max_1_attr, i)
            denominator = denominator + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, max_3_attr) + uabmatrix(max_2_attr, max_4_attr) + uabmatrix(max_2_attr, max_5_attr) + uabmatrix(max_2_attr, max_6_attr) + uabmatrix(max_2_attr, i)
            denominator = denominator + uabmatrix(max_3_attr, max_1_attr) + uabmatrix(max_3_attr, max_2_attr) + uabmatrix(max_3_attr, max_3_attr) + uabmatrix(max_3_attr, max_4_attr) + uabmatrix(max_3_attr, max_5_attr) + uabmatrix(max_3_attr, max_6_attr) + uabmatrix(max_3_attr, i)
            denominator = denominator + uabmatrix(max_4_attr, max_1_attr) + uabmatrix(max_4_attr, max_2_attr) + uabmatrix(max_4_attr, max_3_attr) + uabmatrix(max_4_attr, max_4_attr) + uabmatrix(max_4_attr, max_5_attr) + uabmatrix(max_4_attr, max_6_attr) + uabmatrix(max_4_attr, i)
            denominator = denominator + uabmatrix(max_5_attr, max_1_attr) + uabmatrix(max_5_attr, max_2_attr) + uabmatrix(max_5_attr, max_3_attr) + uabmatrix(max_5_attr, max_4_attr) + uabmatrix(max_5_attr, max_5_attr) + uabmatrix(max_5_attr, max_6_attr) + uabmatrix(max_5_attr, i)
            denominator = denominator + uabmatrix(max_6_attr, max_1_attr) + uabmatrix(max_6_attr, max_2_attr) + uabmatrix(max_6_attr, max_3_attr) + uabmatrix(max_6_attr, max_4_attr) + uabmatrix(max_6_attr, max_5_attr) + uabmatrix(max_6_attr, max_6_attr) + uabmatrix(max_6_attr, i)
            denominator = denominator + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, max_3_attr) + uabmatrix(i, max_4_attr) + uabmatrix(i, max_5_attr) + uabmatrix(i, max_6_attr) + uabmatrix(i, i)
            denominator = denominator ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(max_3_attr, 9) + uabmatrix(max_4_attr, 9) + uabmatrix(max_5_attr, 9) + uabmatrix(max_6_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_7_goodvalue Then
                    max_7_goodvalue = temp_goodness
                    max_7_attr = i
                End If
        End If
    Next
    
    If max_7_goodvalue > max_6_goodvalue Then
        List1.AddItem "Select 7 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ")" & ",A" & max_4_attr & "(" & attribute_name(max_4_attr) & ")" & ",A" & max_5_attr & "(" & attribute_name(max_5_attr) & "),A" & max_6_attr & "(" & attribute_name(max_6_attr) & "),A" & max_7_attr & "(" & attribute_name(max_7_attr) & ") }"
        List1.AddItem "Goodness : " & max_7_goodvalue
        List1.AddItem ""
    End If
    
    'select eight attritube
    Dim max_8_attr As Integer
    Dim max_8_goodvalue As Double
    max_8_goodvalue = 0
    max_8_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Or (i = max_3_attr) Or (i = max_4_attr) Or (i = max_5_attr) Or (i = max_6_attr) Or (i = max_7_attr) Then
            max_8_goodvalue = max_8_goodvalue + 0
        Else
            denominator = uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, max_2_attr) + uabmatrix(max_1_attr, max_3_attr) + uabmatrix(max_1_attr, max_4_attr) + uabmatrix(max_1_attr, max_5_attr) + uabmatrix(max_1_attr, max_6_attr) + uabmatrix(max_1_attr, max_7_attr) + uabmatrix(max_1_attr, i)
            denominator = denominator + uabmatrix(max_2_attr, max_1_attr) + uabmatrix(max_2_attr, max_2_attr) + uabmatrix(max_2_attr, max_3_attr) + uabmatrix(max_2_attr, max_4_attr) + uabmatrix(max_2_attr, max_5_attr) + uabmatrix(max_2_attr, max_6_attr) + uabmatrix(max_2_attr, max_7_attr) + uabmatrix(max_2_attr, i)
            denominator = denominator + uabmatrix(max_3_attr, max_1_attr) + uabmatrix(max_3_attr, max_2_attr) + uabmatrix(max_3_attr, max_3_attr) + uabmatrix(max_3_attr, max_4_attr) + uabmatrix(max_3_attr, max_5_attr) + uabmatrix(max_3_attr, max_6_attr) + uabmatrix(max_3_attr, max_7_attr) + uabmatrix(max_3_attr, i)
            denominator = denominator + uabmatrix(max_4_attr, max_1_attr) + uabmatrix(max_4_attr, max_2_attr) + uabmatrix(max_4_attr, max_3_attr) + uabmatrix(max_4_attr, max_4_attr) + uabmatrix(max_4_attr, max_5_attr) + uabmatrix(max_4_attr, max_6_attr) + uabmatrix(max_4_attr, max_7_attr) + uabmatrix(max_4_attr, i)
            denominator = denominator + uabmatrix(max_5_attr, max_1_attr) + uabmatrix(max_5_attr, max_2_attr) + uabmatrix(max_5_attr, max_3_attr) + uabmatrix(max_5_attr, max_4_attr) + uabmatrix(max_5_attr, max_5_attr) + uabmatrix(max_5_attr, max_6_attr) + uabmatrix(max_5_attr, max_7_attr) + uabmatrix(max_5_attr, i)
            denominator = denominator + uabmatrix(max_6_attr, max_1_attr) + uabmatrix(max_6_attr, max_2_attr) + uabmatrix(max_6_attr, max_3_attr) + uabmatrix(max_6_attr, max_4_attr) + uabmatrix(max_6_attr, max_5_attr) + uabmatrix(max_6_attr, max_6_attr) + uabmatrix(max_6_attr, max_7_attr) + uabmatrix(max_6_attr, i)
            denominator = denominator + uabmatrix(max_7_attr, max_1_attr) + uabmatrix(max_7_attr, max_2_attr) + uabmatrix(max_7_attr, max_3_attr) + uabmatrix(max_7_attr, max_4_attr) + uabmatrix(max_7_attr, max_5_attr) + uabmatrix(max_7_attr, max_6_attr) + uabmatrix(max_7_attr, max_7_attr) + uabmatrix(max_7_attr, i)
            denominator = denominator + uabmatrix(i, max_1_attr) + uabmatrix(i, max_2_attr) + uabmatrix(i, max_3_attr) + uabmatrix(i, max_4_attr) + uabmatrix(i, max_5_attr) + uabmatrix(i, max_6_attr) + uabmatrix(i, max_7_attr) + uabmatrix(i, i)
            denominator = denominator ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(max_3_attr, 9) + uabmatrix(max_4_attr, 9) + uabmatrix(max_5_attr, 9) + uabmatrix(max_6_attr, 9) + uabmatrix(max_7_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_8_goodvalue Then
                    max_8_goodvalue = temp_goodness
                    max_8_attr = i
                End If
        End If
    Next
    
    If max_8_goodvalue > max_7_goodvalue Then
        List1.AddItem "Select 8 Attritube: " & "{ A" & max_1_attr & "(" & attribute_name(max_1_attr) & "),A" & max_2_attr & "(" & attribute_name(max_2_attr) & "),A" & max_3_attr & "(" & attribute_name(max_3_attr) & ")" & ",A" & max_4_attr & "(" & attribute_name(max_4_attr) & ")" & ",A" & max_5_attr & "(" & attribute_name(max_5_attr) & "),A" & max_6_attr & "(" & attribute_name(max_6_attr) & "),A" & max_7_attr & "(" & attribute_name(max_7_attr) & "),A" & max_8_attr & "(" & attribute_name(max_8_attr) & ") }"
        List1.AddItem "Goodness : " & max_8_goodvalue
    End If
    
End Sub

Private Sub backward_click()
    List1.Clear
    Dim temp_goodness As Double
    Dim i As Integer
    Dim j As Integer
    Dim denominator As Double
    denominator = 0
    temp_goodness = 0
    
    'attribute name
    Dim attribute_name(9) As String
    attribute_name(1) = "mcg"
    attribute_name(2) = "gvh"
    attribute_name(3) = "alm"
    attribute_name(4) = "mit"
    attribute_name(5) = "erl"
    attribute_name(6) = "pox"
    attribute_name(7) = "vac"
    attribute_name(8) = "nuc"
    
    'select eight attritube
    Dim del_0_attr As Integer
    Dim del_0_goodvalue As Double
    del_0_goodvalue = 0
    For i = 1 To 8
        For j = 1 To 8
            denominator = denominator + uabmatrix(i, j)
        Next
        temp_goodness = temp_goodness + uabmatrix(i, 9)
    Next
    denominator = denominator ^ 0.5
    temp_goodness = temp_goodness / denominator
    del_0_goodvalue = temp_goodness
    
    If del_0_goodvalue > 0 Then
        List1.AddItem "Select all Attritube: { A1,A2,A3,A4,A5,A6,A7,A8 }"
        List1.AddItem "Goodness : " & del_0_goodvalue
    End If
    
    'delete one attritube
    Dim temp_temp_goodness As Double
    Dim remove_1_attr As Integer
    Dim del_1_attr As Integer
    Dim del_1_goodvalue As Double
    Dim temp_denominator As Double
    del_1_goodvalue = 0
    del_1_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    temp_denominator = 0
    denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_1_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(del_1_attr, del_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(del_1_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, del_1_attr)
        Next
        denominator = denominator ^ 0.5
        temp_goodness = temp_goodness - uabmatrix(del_1_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_1_goodvalue Then
            del_1_goodvalue = temp_goodness
            remove_1_attr = del_1_attr
        End If
    Next
    
    If del_1_goodvalue > del_0_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 1 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")"
        List1.AddItem "Goodness : " & del_1_goodvalue
    End If
        
    'delete two attritube
    Dim remove_2_attr As Integer
    Dim del_2_attr As Integer
    Dim del_2_goodvalue As Double
    del_2_goodvalue = 0
    del_2_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_2_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(del_2_attr, del_2_attr) + uabmatrix(remove_1_attr, del_2_attr) + uabmatrix(del_2_attr, remove_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(del_2_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, del_2_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(del_2_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_2_goodvalue Then
            del_2_goodvalue = temp_goodness
            remove_2_attr = del_2_attr
        End If
    Next

    If del_2_goodvalue > del_1_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 2 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")"
        List1.AddItem "Goodness : " & del_2_goodvalue
    End If
    
    'delete three attritube
    Dim remove_3_attr As Integer
    Dim del_3_attr As Integer
    Dim del_3_goodvalue As Double
    del_3_goodvalue = 0
    del_3_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_3_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(remove_2_attr, remove_2_attr) + uabmatrix(remove_1_attr, remove_2_attr) + uabmatrix(remove_2_attr, remove_1_attr) + uabmatrix(del_3_attr, remove_1_attr) + uabmatrix(remove_1_attr, del_3_attr) + uabmatrix(remove_2_attr, del_3_attr) + uabmatrix(del_3_attr, remove_2_attr) + uabmatrix(del_3_attr, del_3_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(remove_2_attr, i)
            denominator = denominator - uabmatrix(del_3_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, remove_2_attr)
            denominator = denominator - uabmatrix(i, del_3_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(remove_2_attr, 9) - uabmatrix(del_3_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_3_goodvalue Then
            del_3_goodvalue = temp_goodness
            remove_3_attr = del_3_attr
        End If
    Next

    If del_3_goodvalue > del_2_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 3 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")," & "A" & remove_3_attr & "(" & attribute_name(remove_3_attr) & ")"
        List1.AddItem "Goodness : " & del_3_goodvalue
    End If
    
    'delete four attritube
    Dim remove_4_attr As Integer
    Dim del_4_attr As Integer
    Dim del_4_goodvalue As Double
    del_4_goodvalue = 0
    del_4_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_4_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(remove_2_attr, remove_2_attr) + uabmatrix(remove_1_attr, remove_2_attr) + uabmatrix(remove_2_attr, remove_1_attr) + uabmatrix(remove_3_attr, remove_1_attr) + uabmatrix(remove_1_attr, remove_3_attr) + uabmatrix(remove_2_attr, remove_3_attr) + uabmatrix(remove_3_attr, remove_2_attr) + uabmatrix(remove_3_attr, remove_3_attr)
        denominator = denominator + uabmatrix(remove_1_attr, del_4_attr) + uabmatrix(remove_2_attr, del_4_attr) + uabmatrix(remove_3_attr, del_4_attr) + uabmatrix(del_4_attr, del_4_attr) + uabmatrix(del_4_attr, remove_3_attr) + uabmatrix(del_4_attr, remove_2_attr) + uabmatrix(del_4_attr, remove_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(remove_2_attr, i)
            denominator = denominator - uabmatrix(remove_3_attr, i)
            denominator = denominator - uabmatrix(del_4_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, remove_2_attr)
            denominator = denominator - uabmatrix(i, remove_3_attr)
            denominator = denominator - uabmatrix(i, del_4_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(remove_2_attr, 9) - uabmatrix(remove_3_attr, 9) - uabmatrix(del_4_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_4_goodvalue Then
            del_4_goodvalue = temp_goodness
            remove_4_attr = del_4_attr
        End If
    Next

    If del_4_goodvalue > del_3_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 4 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")," & "A" & remove_3_attr & "(" & attribute_name(remove_3_attr) & ")," & "A" & remove_4_attr & "(" & attribute_name(remove_4_attr) & ")"
        List1.AddItem "Goodness : " & del_4_goodvalue
    End If
    
    'delete five attritube
    Dim remove_5_attr As Integer
    Dim del_5_attr As Integer
    Dim del_5_goodvalue As Double
    del_5_goodvalue = 0
    del_5_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_5_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(remove_2_attr, remove_2_attr) + uabmatrix(remove_1_attr, remove_2_attr) + uabmatrix(remove_2_attr, remove_1_attr) + uabmatrix(remove_3_attr, remove_1_attr) + uabmatrix(remove_1_attr, remove_3_attr) + uabmatrix(remove_2_attr, remove_3_attr) + uabmatrix(remove_3_attr, remove_2_attr) + uabmatrix(remove_3_attr, remove_3_attr)
        denominator = denominator + uabmatrix(remove_1_attr, remove_4_attr) + uabmatrix(remove_2_attr, remove_4_attr) + uabmatrix(remove_3_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_3_attr) + uabmatrix(remove_4_attr, remove_2_attr) + uabmatrix(remove_4_attr, remove_1_attr)
        denominator = denominator + uabmatrix(remove_1_attr, del_5_attr) + uabmatrix(remove_2_attr, del_5_attr) + uabmatrix(remove_3_attr, del_5_attr) + uabmatrix(remove_4_attr, del_5_attr) + uabmatrix(del_5_attr, del_5_attr) + uabmatrix(del_5_attr, remove_1_attr) + uabmatrix(del_5_attr, remove_2_attr) + uabmatrix(del_5_attr, remove_3_attr) + uabmatrix(del_5_attr, remove_2_attr) + uabmatrix(del_5_attr, remove_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(remove_2_attr, i)
            denominator = denominator - uabmatrix(remove_3_attr, i)
            denominator = denominator - uabmatrix(remove_4_attr, i)
            denominator = denominator - uabmatrix(del_5_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, remove_2_attr)
            denominator = denominator - uabmatrix(i, remove_3_attr)
            denominator = denominator - uabmatrix(i, remove_4_attr)
            denominator = denominator - uabmatrix(i, del_5_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(remove_2_attr, 9) - uabmatrix(remove_3_attr, 9) - uabmatrix(del_4_attr, 9) - uabmatrix(del_5_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_5_goodvalue Then
            del_5_goodvalue = temp_goodness
            remove_5_attr = del_5_attr
        End If
    Next

    If del_5_goodvalue > del_4_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 5 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")," & "A" & remove_3_attr & "(" & attribute_name(remove_3_attr) & ")," & "A" & remove_4_attr & "(" & attribute_name(remove_4_attr) & ")" & "A" & remove_5_attr & "(" & attribute_name(remove_5_attr) & ")"
        List1.AddItem "Goodness : " & del_5_goodvalue
    End If
    
    'delete six attritube
    Dim remove_6_attr As Integer
    Dim del_6_attr As Integer
    Dim del_6_goodvalue As Double
    del_6_goodvalue = 0
    del_6_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_6_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(remove_2_attr, remove_2_attr) + uabmatrix(remove_1_attr, remove_2_attr) + uabmatrix(remove_2_attr, remove_1_attr) + uabmatrix(remove_3_attr, remove_1_attr) + uabmatrix(remove_1_attr, remove_3_attr) + uabmatrix(remove_2_attr, remove_3_attr) + uabmatrix(remove_3_attr, remove_2_attr) + uabmatrix(remove_3_attr, remove_3_attr)
        denominator = denominator + uabmatrix(remove_1_attr, remove_4_attr) + uabmatrix(remove_2_attr, remove_4_attr) + uabmatrix(remove_3_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_3_attr) + uabmatrix(remove_4_attr, remove_2_attr) + uabmatrix(remove_4_attr, remove_1_attr)
        denominator = denominator + uabmatrix(remove_1_attr, del_6_attr) + uabmatrix(remove_2_attr, del_6_attr) + uabmatrix(remove_3_attr, del_6_attr) + uabmatrix(remove_4_attr, del_6_attr) + uabmatrix(del_5_attr, del_6_attr) + uabmatrix(del_6_attr, remove_1_attr) + uabmatrix(del_6_attr, remove_2_attr) + uabmatrix(del_6_attr, remove_3_attr) + uabmatrix(del_6_attr, remove_2_attr) + uabmatrix(del_6_attr, remove_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(remove_2_attr, i)
            denominator = denominator - uabmatrix(remove_3_attr, i)
            denominator = denominator - uabmatrix(remove_4_attr, i)
            denominator = denominator - uabmatrix(remove_5_attr, i)
            denominator = denominator - uabmatrix(del_6_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, remove_2_attr)
            denominator = denominator - uabmatrix(i, remove_3_attr)
            denominator = denominator - uabmatrix(i, remove_4_attr)
            denominator = denominator - uabmatrix(i, remove_5_attr)
            denominator = denominator - uabmatrix(i, del_6_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(remove_2_attr, 9) - uabmatrix(remove_3_attr, 9) - uabmatrix(del_4_attr, 9) - uabmatrix(del_5_attr, 9) - uabmatrix(del_6_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_6_goodvalue Then
            del_6_goodvalue = temp_goodness
            remove_6_attr = del_6_attr
        End If
    Next

    If del_6_goodvalue > del_5_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 6 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")," & "A" & remove_3_attr & "(" & attribute_name(remove_3_attr) & ")," & "A" & remove_4_attr & "(" & attribute_name(remove_4_attr) & ")" & "A" & remove_5_attr & "(" & attribute_name(remove_5_attr) & ")"
        List1.AddItem "Goodness : " & del_6_goodvalue
    End If
    
    'delete seven attritube
    Dim remove_7_attr As Integer
    Dim del_7_attr As Integer
    Dim del_7_goodvalue As Double
    del_7_goodvalue = 0
    del_7_attr = 0
    temp_goodness = 0
    temp_temp_goodness = 0
    denominator = 0
    temp_denominator = 0
    For i = 1 To 8
        For j = 1 To 8
            temp_denominator = temp_denominator + uabmatrix(i, j)
        Next
        temp_temp_goodness = temp_temp_goodness + uabmatrix(i, 9)
    Next

    For del_7_attr = 1 To 8
        denominator = temp_denominator
        temp_goodness = temp_temp_goodness
        denominator = denominator + uabmatrix(remove_1_attr, remove_1_attr) + uabmatrix(remove_2_attr, remove_2_attr) + uabmatrix(remove_1_attr, remove_2_attr) + uabmatrix(remove_2_attr, remove_1_attr) + uabmatrix(remove_3_attr, remove_1_attr) + uabmatrix(remove_1_attr, remove_3_attr) + uabmatrix(remove_2_attr, remove_3_attr) + uabmatrix(remove_3_attr, remove_2_attr) + uabmatrix(remove_3_attr, remove_3_attr)
        denominator = denominator + uabmatrix(remove_1_attr, remove_4_attr) + uabmatrix(remove_2_attr, remove_4_attr) + uabmatrix(remove_3_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_4_attr) + uabmatrix(remove_4_attr, remove_3_attr) + uabmatrix(remove_4_attr, remove_2_attr) + uabmatrix(remove_4_attr, remove_1_attr)
        denominator = denominator + uabmatrix(remove_1_attr, del_7_attr) + uabmatrix(remove_2_attr, del_7_attr) + uabmatrix(remove_3_attr, del_7_attr) + uabmatrix(remove_4_attr, del_7_attr) + uabmatrix(del_7_attr, del_7_attr) + uabmatrix(del_7_attr, remove_1_attr) + uabmatrix(del_7_attr, remove_2_attr) + uabmatrix(del_7_attr, remove_3_attr) + uabmatrix(del_7_attr, remove_2_attr) + uabmatrix(del_7_attr, remove_1_attr)
        For i = 1 To 8
            denominator = denominator - uabmatrix(remove_1_attr, i)
            denominator = denominator - uabmatrix(remove_2_attr, i)
            denominator = denominator - uabmatrix(remove_3_attr, i)
            denominator = denominator - uabmatrix(remove_4_attr, i)
            denominator = denominator - uabmatrix(remove_5_attr, i)
            denominator = denominator - uabmatrix(remove_6_attr, i)
            denominator = denominator - uabmatrix(del_7_attr, i)
        Next
        For i = 1 To 8
            denominator = denominator - uabmatrix(i, remove_1_attr)
            denominator = denominator - uabmatrix(i, remove_2_attr)
            denominator = denominator - uabmatrix(i, remove_3_attr)
            denominator = denominator - uabmatrix(i, remove_4_attr)
            denominator = denominator - uabmatrix(i, remove_5_attr)
            denominator = denominator - uabmatrix(i, remove_6_attr)
            denominator = denominator - uabmatrix(i, del_7_attr)
        Next
        denominator = denominator ^ 0.5
        
        temp_goodness = temp_goodness - uabmatrix(remove_1_attr, 9) - uabmatrix(remove_2_attr, 9) - uabmatrix(remove_3_attr, 9) - uabmatrix(del_4_attr, 9) - uabmatrix(del_5_attr, 9) - uabmatrix(del_6_attr, 9) - uabmatrix(del_7_attr, 9)
        temp_goodness = temp_goodness / denominator
        If temp_goodness > del_7_goodvalue Then
            del_7_goodvalue = temp_goodness
            remove_7_attr = del_7_attr
        End If
    Next

    If del_7_goodvalue > del_6_goodvalue Then
        List1.AddItem ""
        List1.AddItem "Remove 7 Attritube: " & "A" & remove_1_attr & "(" & attribute_name(remove_1_attr) & ")," & "A" & remove_2_attr & "(" & attribute_name(remove_2_attr) & ")," & "A" & remove_3_attr & "(" & attribute_name(remove_3_attr) & ")," & "A" & remove_4_attr & "(" & attribute_name(remove_4_attr) & ")" & "A" & remove_5_attr & "(" & attribute_name(remove_5_attr) & ")" & "A" & remove_6_attr & "(" & attribute_name(remove_6_attr) & ")" & "A" & remove_7_attr & "(" & attribute_name(remove_7_attr) & ")"
        List1.AddItem "Goodness : " & del_7_goodvalue
    End If
End Sub

Function log2(X As Double) As Double
    If X = 0 Then
        log2 = 0
    Else
        log2 = Log(X) / Log(2)
    End If
End Function

Private Sub Text1_Change()

End Sub
