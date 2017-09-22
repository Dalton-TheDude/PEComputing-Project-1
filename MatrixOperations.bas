Attribute VB_Name = "MatrixOperations"
Option Explicit

Function MatrixMult(mata() As Double, matb() As Double) As Variant


    Dim nrowsa As Integer, ncolsa As Integer, nrowsb As Integer, ncolsb As Integer
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    nrowsa = UBound(mata, 1) - LBound(mata, 1) + 1
    nrowsb = UBound(matb, 1) - LBound(matb, 1) + 1
    ncolsa = UBound(mata, 2) - LBound(mata, 2) + 1
    ncolsb = UBound(matb, 2) - LBound(matb, 2) + 1
    
    If ncolsa = nrowsb Then
        Dim matc() As Double
        ReDim matc(1 To nrowsa, 1 To ncolsb)
        
        For k = 1 To ncolsb Step 1
            For i = 1 To nrowsa Step 1
                For j = 1 To ncolsa Step 1
                    matc(i, k) = matc(i, k) + (mata(i, j) * matb(j, k))
                            
                Next j
            Next i
        Next k
        
        MatrixMult = matc
    
    k = 2
    Do While Range("b1").Cells(k, 1).Value <> ""
        k = k + 1
    Loop
    
    Range("b1").Cells(k + i - 1, 1).Value = "Matrix Multiplication"
    
    For i = 1 To nrowsa Step 1
        For j = 1 To ncolsb Step 1
            Range("b1").Cells(k + i, j).Value = matc(i, j)
        Next j
    Next i
    
    Else
        MsgBox ("The matrices aren't the correct dimensions for multiplication.")
        Exit Function
    
    End If


End Function


Function MatrixAdd(mata() As Double, matb() As Double) As Variant


    Dim nrowsa As Integer, ncolsa As Integer, nrowsb As Integer, ncolsb As Integer
    Dim i As Integer, j As Integer, k As Integer
    nrowsa = UBound(mata, 1) - LBound(mata, 1) + 1
    nrowsb = UBound(matb, 1) - LBound(matb, 1) + 1
    ncolsa = UBound(mata, 2) - LBound(mata, 2) + 1
    ncolsb = UBound(matb, 2) - LBound(matb, 2) + 1
    
    If nrowsa = nrowsb And ncolsa = ncolsb Then
        Dim matc() As Double
        ReDim matc(1 To nrowsa, 1 To ncolsa)
        
        For i = 1 To nrowsa Step 1
            For j = 1 To ncolsa Step 1
                matc(i, j) = mata(i, j) + matb(i, j)
     '           Range("b1").Cells(i, j).Value = matc(i, j)
     
            Next j
        Next i
        
        MatrixAdd = matc
    
    k = 2
    Do While Range("b1").Cells(k, 1).Value <> ""
        k = k + 1
    Loop
        
    Range("b1").Cells(k - 1, 1).Value = "Matrix Addition"
    
    For i = 1 To nrowsa Step 1
        For j = 1 To ncolsb Step 1
            Range("b1").Cells(k + i, j).Value = matc(i, j)
        Next j
    Next i
    
    Else
        MsgBox ("The matrices aren't the correct dimensions for addition")
        Exit Function
    
    End If
    
            
End Function

Sub MatrixTest()

    ' Parameters

    Dim matr1() As Double, matr2() As Double
    Dim i As Integer, j As Integer, rows1 As Integer, rows2 As Integer, cols1 As Integer, cols2 As Integer
    rows1 = 5
    cols1 = 5
    rows2 = 5
    cols2 = 5
    ReDim matr1(1 To rows1, 1 To cols1)
    ReDim matr2(1 To rows2, 1 To cols2)
    
    'Filling the matrices
    
    For i = 1 To rows1 Step 1
        For j = 1 To cols1 Step 1
            matr1(i, j) = 1
            'matr1(i, j) = Application.WorksheetFunction.RandBetween(0, 10)
            
        Next j
    Next i
        
    
    For i = 1 To rows2 Step 1
        For j = 1 To cols2 Step 1
            matr2(i, j) = 1
            'matr2(i, j) = Application.WorksheetFunction.RandBetween(0, 10)
            
        Next j
    Next i

            
    Call MatrixAdd(matr1, matr2)
            

    Call MatrixMult(matr1, matr2)
        

End Sub
