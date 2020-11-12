' Bjontegaard Metric implementation for Excel.
'
' Provides two functions as Add-In:
'   BDSNR(BR1, PSNR1, BR2, PSNR2)
'     Returns the delta-SNR Bjontegaard (in dB)
'
'   BDBR(BR1, PSNR1, BR2, PSNR2)
'     Returns the delta-rate Bjontegaard (in %)
'
' Author:
'   Tim Bruylants, ETRO, Vrije Universiteit Brussel
'
' References:
'   [1] G. Bjontegaard, Calculation of average PSNR differences between RD-curves (VCEG-M33)
'   [2] S. Pateux, J. Jung, An excel add-in for computing Bjontegaard metric and its evolution
'
' The MIT License (MIT)
' Copyright (c) 2013 Tim Bruylants, ETRO, Vrije Universiteit Brussel
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software
' and associated documentation files (the "Software"), to deal in the Software without
' restriction, including without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom
' the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.

' Calculates Y(x), where the coefficients of polynomial are given by P
Private Function bjontegaard_polyval(P As Variant, X As Double)
    Dim xPow As Double, result As Double
    
    result = 0
    xPow = 1
    For i = UBound(P) To 0 Step -1
        result = result + xPow * P(i)
        xPow = xPow * X
    Next i
    
    bjontegaard_polyval = result
End Function

' Fits the curve and calculates the integral
Private Function bjontegaard_polyfit_and_integrate(X As Variant, Y As Variant, lowX As Double, highX As Double)
    ' Constants (yes, the array is not a const, thanks to the great VBA language)
    Const order As Integer = 3
    Dim powerArray() As Variant
    powerArray = Array(1, 2, 3) ' match this array with the order const
    
    ' Get the number of points to use and check validity
    Dim noPoints As Integer
    noPoints = UBound(X)
    If noPoints <> UBound(Y) Then
        Err.Raise vbObjectError + 1, "avsnr_polyfit_and_integrate", "Number of X-values does not match the number of Y-values."
    End If
    
    ' Polyfit 3rd order and calculate polynomial coefficients
    Dim P
    P = Application.WorksheetFunction.LinEst(Y, Application.Power(X, Application.WorksheetFunction.Transpose(powerArray)))
    
    ' Integrate the polynomial
    For i = 1 To order
        P(i) = P(i) / (order + 2 - i) ' + 2 to integrate and compensate for initial index (start from 1 vs 0 stuff)
    Next i
    ReDim Preserve P(UBound(P))
    P(order + 1) = 0
    ' At this point, P contains the integrated polynomial coefficients
    
    ' Use polynomial function to calculate the numerical integral between low and high (and return)
    bjontegaard_polyfit_and_integrate = bjontegaard_polyval(P, highX) - bjontegaard_polyval(P, lowX)
End Function

' Calculate a Bjontegaard difference value
Private Function bjontegaard_diff(X1 As Variant, Y1 As Variant, X2 As Variant, Y2 As Variant)
    Dim lowX As Double, highX As Double
    highX = Application.WorksheetFunction.Min(Application.WorksheetFunction.Max(X1), Application.WorksheetFunction.Max(X2))
    lowX = Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(X1), Application.WorksheetFunction.Min(X2))
    
    Dim int1, int2
    int1 = bjontegaard_polyfit_and_integrate(X1, Y1, lowX, highX)
    int2 = bjontegaard_polyfit_and_integrate(X2, Y2, lowX, highX)
    
    ' Return the difference value
    bjontegaard_diff = (int2 - int1) / (highX - lowX)
End Function

' Bjontegaard delta-SNR metric (in dB)
Function BDSNR(BR1 As Range, PSNR1 As Range, BR2 As Range, PSNR2 As Range)
    ' Error checking
    If BR1.Count <> PSNR1.Count Or BR2.Count <> PSNR2.Count Or BR1.Count < 4 Or BR2.Count < 4 Then
        BDSNR = CVErr(xlErrRef)
        Return
    End If

    ' Get data for two curves
    Dim BR1data As Variant, PSNR1data As Variant, BR2data As Variant, PSNR2data As Variant
    BR1data = WorksheetFunction.Transpose(BR1)
    PSNR1data = WorksheetFunction.Transpose(PSNR1)
    BR2data = WorksheetFunction.Transpose(BR2)
    PSNR2data = WorksheetFunction.Transpose(PSNR2)
    
    ' Put rates in logarithmic scale
    For i = 1 To BR1.Count
        BR1data(i) = Application.WorksheetFunction.Ln(BR1data(i))
    Next i
    For i = 1 To BR2.Count
        BR2data(i) = Application.WorksheetFunction.Ln(BR2data(i))
    Next i
    
    ' Calculate the Bjontegaard difference
    BDSNR = bjontegaard_diff(BR1data, PSNR1data, BR2data, PSNR2data)
End Function

' Bjontegaard delta-BR metric (in %)
Function BDBR(BR1 As Range, PSNR1 As Range, BR2 As Range, PSNR2 As Range)
    ' Error checking
    If BR1.Count <> PSNR1.Count Or BR2.Count <> PSNR2.Count Or BR1.Count < 4 Or BR2.Count < 4 Then
        BDBR = CVErr(xlErrRef)
        Return
    End If

    ' Get data for two curves
    Dim BR1data As Variant, PSNR1data As Variant, BR2data As Variant, PSNR2data As Variant
    BR1data = WorksheetFunction.Transpose(BR1)
    PSNR1data = WorksheetFunction.Transpose(PSNR1)
    BR2data = WorksheetFunction.Transpose(BR2)
    PSNR2data = WorksheetFunction.Transpose(PSNR2)
    
    ' Put rates in logarithmic scale
    For i = 1 To BR1.Count
        BR1data(i) = Application.WorksheetFunction.Ln(BR1data(i))
    Next i
    For i = 1 To BR2.Count
        BR2data(i) = Application.WorksheetFunction.Ln(BR2data(i))
    Next i
    
    ' Calculate the Bjontegaard difference
    BDBR = (Exp(bjontegaard_diff(PSNR1data, BR1data, PSNR2data, BR2data)) - 1) * 100
End Function
