VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12990
   LinkTopic       =   "Form2"
   ScaleHeight     =   10155
   ScaleWidth      =   12990
   Begin VB.CommandButton backward 
      Caption         =   "Backward"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   10
      Top             =   3840
      Width           =   3000
   End
   Begin VB.CommandButton forward 
      Caption         =   "Forward"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   3840
      Width           =   3000
   End
   Begin VB.CommandButton EntropyBase 
      Caption         =   "Entropy-Based"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10680
      TabIndex        =   8
      Top             =   1900
      Width           =   1915
   End
   Begin VB.CommandButton EqualFrequency 
      Caption         =   "Equal-Frequency discretization"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      TabIndex        =   6
      Top             =   1900
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
      Left            =   6360
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   5
      Top             =   1900
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
      Height          =   8040
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
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
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
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
      Left            =   6360
      TabIndex        =   7
      Top             =   1440
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1080
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
Dim x As Integer
Dim y As Integer
Dim col As Integer

Dim data(1484, 10) As String  'data1 used in equal-width,changing value after discretization
Dim data2(1484, 10) As String 'data2 is a sorted matrix, used in equal-frequency
Dim data3(1484, 10) As String 'data3 used in equal-frequency, store value after discretization

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
            Open App.Path & "\test-freq.txt" For Output As #3
            
            y = 0 'from row 0
            
            'read each line until end of file
            Do While Not EOF(1)
                Line Input #1, atts
                each_row_output = Split(atts, " ")
                col = 0
                For x = 0 To UBound(each_row_output)
                    If each_row_output(x) <> "" Then 'if space,ignore
                        data(y, col) = each_row_output(x)
                        data2(y, col) = each_row_output(x)
                        data3(y, col) = each_row_output(x)
                    col = col + 1
                    End If
                Next
                y = y + 1 'goto next row
            Loop
            
            MsgBox "Input file has prepared successfully !"
            MsgBox "please choose discretization method the right hand side!"
            
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
            For x = 0 To 1483
                If CDbl(data(x, col)) > max(col) Then
                    max(col) = CDbl(data(x, col))
                ElseIf CDbl(data(x, col)) < min(col) Then
                    min(col) = CDbl(data(x, col))
                End If
            Next
        equal_width(col) = (max(col) - min(col)) / interval
    Next
    
    'print splitting point
    For col = 1 To 8
        List1.AddItem "Attribute : " & col & " " & "    Width= " & equal_width(col)
        For x = 1 To (interval - 1)
            List1.AddItem "splitting point = " & min(col) + x * equal_width(col)
        Next
        List1.AddItem ""
    Next
    
    'discretization to 1-10, each interval includes splitting point(lower bound)
    For col = 1 To 8
        For x = 0 To 1483
            If CDbl(data(x, col)) >= (min(col) + (interval - 1) * equal_width(col)) Then 'the 10th interval
                data(x, col) = 10
            ElseIf CDbl(data(x, col)) <> -1 Then
                For y = 1 To interval - 1 'the 1th to 9th
                    If data(x, col) < (min(col) + y * equal_width(col)) Then
                        data(x, col) = y
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
        For x = 1 To 10
            width_probability(x, col) = 0
        Next
    Next
    
    
    'count the total appear times of each attribute's discrete value
    For col = 1 To 8
        For x = 0 To 1483
            Select Case CDbl(data(x, col))
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
    For x = 0 To 1483
        Select Case data(x, 9)
            Case "CYT"
                width_probability(1, 9) = width_probability(1, 9) + 1
                data(x, 9) = 1
            Case "NUC"
                width_probability(2, 9) = width_probability(2, 9) + 1
                data(x, 9) = 2
            Case "MIT"
                width_probability(3, 9) = width_probability(3, 9) + 1
                data(x, 9) = 3
            Case "ME3"
                width_probability(4, 9) = width_probability(4, 9) + 1
                data(x, 9) = 4
            Case "ME2"
                width_probability(5, 9) = width_probability(5, 9) + 1
                data(x, 9) = 5
            Case "ME1"
                width_probability(6, 9) = width_probability(6, 9) + 1
                data(x, 9) = 6
            Case "EXC"
                width_probability(7, 9) = width_probability(7, 9) + 1
                data(x, 9) = 7
            Case "VAC"
                width_probability(8, 9) = width_probability(8, 9) + 1
                data(x, 9) = 8
            Case "POX"
                width_probability(9, 9) = width_probability(9, 9) + 1
                data(x, 9) = 9
            Case Else
                width_probability(10, 9) = width_probability(10, 9) + 1
                data(x, 9) = 10
        End Select
    Next
    
    For col = 1 To 9
        For x = 1 To 10
            width_probability(x, col) = width_probability(x, col) / 1484
        Next
    Next
    
    For x = 1 To 9
        h_value_att(x) = h_value(width_probability, x)
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
        For x = 0 To 1483
            For y = x To 1483
                If CDbl(data2(x, col)) > CDbl(data2(y, col)) Then
                    'swift
                    Tag = data2(x, col)
                    data2(x, col) = data2(y, col)
                    data2(y, col) = Tag
                    'change class
                    class = data2(x, 0)
                    data2(x, 0) = data2(y, 0)
                    data2(y, 0) = class
                End If
            Next y
        Next x
    Next col
            
    interval = 10
    frequency = CInt(1484 / interval) '148
    For col = 1 To 8
        For x = 1 To 9
            frequency_cutpoint(col, x) = (CDbl(data2(x * frequency - 1, col)) + CDbl(data2(x * frequency, col))) / 2
        Next
    Next

    For col = 1 To 8
        List1.AddItem "Attribute : " & col
        For x = 1 To interval - 1
            List1.AddItem "splitting point = " & frequency_cutpoint(col, x)
        Next
        List1.AddItem ""
    Next
    
    For col = 1 To 8
        For x = 0 To 1483
            If CDbl(data3(x, col)) >= frequency_cutpoint(col, interval - 1) Then 'the 10th interval
                data3(x, col) = 10
            ElseIf CDbl(data3(x, col)) <> -1 Then
                For y = 1 To interval - 1 'the 1th to 9th
                    If data3(x, col) < frequency_cutpoint(col, y) Then
                        data3(x, col) = y
                    End If
                Next
            End If
        Next
    Next
    
    
    'initialize the matrix
    For col = 1 To 9
        For x = 1 To 10
            freq_probability(x, col) = 0
        Next
    Next
    
    
    
    'count the total appear times of each attribute's discrete value
    For col = 1 To 8
        For x = 0 To 1483
            Select Case CDbl(data3(x, col))
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
    For x = 0 To 1483
        Select Case data3(x, 9)
            Case "CYT"
                freq_probability(1, 9) = freq_probability(1, 9) + 1
                data3(x, 9) = 1
            Case "NUC"
                freq_probability(2, 9) = freq_probability(2, 9) + 1
                data3(x, 9) = 2
            Case "MIT"
                freq_probability(3, 9) = freq_probability(3, 9) + 1
                data3(x, 9) = 3
            Case "ME3"
                freq_probability(4, 9) = freq_probability(4, 9) + 1
                data3(x, 9) = 4
            Case "ME2"
                freq_probability(5, 9) = freq_probability(5, 9) + 1
                data3(x, 9) = 5
            Case "ME1"
                freq_probability(6, 9) = freq_probability(6, 9) + 1
                data3(x, 9) = 6
            Case "EXC"
                freq_probability(7, 9) = freq_probability(7, 9) + 1
                data3(x, 9) = 7
            Case "VAC"
                freq_probability(8, 9) = freq_probability(8, 9) + 1
                data3(x, 9) = 8
            Case "POX"
                freq_probability(9, 9) = freq_probability(9, 9) + 1
                data3(x, 9) = 9
            Case Else
                freq_probability(10, 9) = freq_probability(10, 9) + 1
                data3(x, 9) = 10
        End Select
    Next
        
    For col = 1 To 9
        For x = 1 To 10
            freq_probability(x, col) = freq_probability(x, col) / 1484
        Next
    Next
    
    'print
    For col = 5 To 5
        For x = 1 To 10
            Print #3, freq_probability(x, col)
        Next
    Next
    
    For x = 1 To 9
        h_value_att(x) = h_value(freq_probability, x)
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
    List1.AddItem "nothing"
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
    If h_value_att(attr1) + h_value_att(attr2) = 0 Then
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
    List1.AddItem "Select 1 Attritube : " & "A" & max_1_attr
    List1.AddItem "Goodness : " & max_1_goodvalue
    
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
    List1.AddItem "Select 2 Attritube : " & "A" & max_1_attr & " , A" & max_2_attr
    List1.AddItem "Goodness : " & max_2_goodvalue
    
    'select three attritube
    Dim max_3_attr As Integer
    Dim max_3_goodvalue As Double
    max_3_goodvalue = 0
    max_3_attr = 0
    For i = 1 To 8
        If (i = max_1_attr) Or (i = max_2_attr) Then
            max_3_goodvalue = max_3_goodvalue + 0
        Else
            'denominator = (uabmatrix(max_1_attr, max_1_attr) + uabmatrix(max_1_attr, i) + uabmatrix(i, max_1_attr) + uabmatrix(i, i)) ^ 0.5
            temp_goodness = (uabmatrix(max_1_attr, 9) + uabmatrix(max_2_attr, 9) + uabmatrix(i, 9)) / denominator
                If temp_goodness > max_3_goodvalue Then
                    max_3_goodvalue = temp_goodness
                    max_3_attr = i
                End If
        End If
    Next
    List1.AddItem "Select 3 Attritube : " & "A" & max_1_attr & " , A" & max_2_attr & ", A" & max_3_attr
    List1.AddItem "Goodness : " & max_3_goodvalue
End Sub

Function feature_selection(selectsets, i) As Double
    Dim maxgood As Double
    Dim tempgood As Double
    maxgood = 0
    For i = 1 To 8
        tempgood = goodness(selectsets, i)
        If tempgood > maxgood Then
            maxgood = tempgood
            selectsets(i) = 1
        End If
    Next
    feature_selection = maxgood
End Function

Function goodness(selectset, i) As Double
    'For i = 1 To 8
        'If selectset(i) = 1 Then
            goodness = uabmatrix(i, 9) / (uabmatrix(i, i)) ^ (0.5)
        'End If
    'Next
End Function

Function log2(x As Double) As Double
    If x = 0 Then
        log2 = 0
    Else
        log2 = Log(x) / Log(2)
    End If
End Function
