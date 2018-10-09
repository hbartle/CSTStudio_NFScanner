'#Language "WWB-COM"

Option Explicit

Sub Main

' Definitions

Dim NumberOfProbes As Integer
Dim ScannerDistance As Double
Dim methods$(1)
methods$(0) = "Realistic"
methods$(1) = "Golden Spiral"
Dim Method As Integer
Dim axisoptions$(2)
axisoptions$(0) = "X-Axis"
axisoptions$(1) = "Y-Axis"
axisoptions$(2) = "Z-Axis"
Dim Axis As Integer

NumberOfProbes = 100
ScannerDistance = 300.0


'User Dialog For Planar Scanner Parameters
	Begin Dialog UserDialog 520,130 ' %GRID:20,10,1,1
		TextBox 160,70,160,20,.NumberOfProbes
		TextBox 160,100,160,20,.ScannerDistance
		OKButton 380,60,120,20
		CancelButton 380,90,120,20
		Text 20,70,140,20,"Number Of Probes",.Text1
		Text 20,100,140,20,"Scanner Distance",.Text2
		DropListBox 20,30,160,20,methods(),.Method
		DropListBox 200,30,160,20,axisoptions(),.Axis
		Text 20,10,140,10,"Placing Method",.Text3
		Text 200,10,140,10,"Reference Axis",.Text4
	End Dialog
Dim dlg As UserDialog
dlg.NumberOfProbes=CStr(NumberOfProbes)
dlg.ScannerDistance=CStr(ScannerDistance)
If (Dialog(dlg) = 0) Then Exit All

' Save User Input
NumberOfProbes = Evaluate(dlg.NumberOfProbes)
ScannerDistance =  Evaluate(dlg.ScannerDistance)
Method = Evaluate(dlg.Method)
Axis = Evaluate(dlg.Axis)

' Calculate Probe Positions
Dim ThetaAngles() As Double
Dim PhiAngles() As Double

Dim i As Integer
Dim k As Integer
Dim idx As Integer

If Method = 0 Then
	'Realistic Placing Method
	ReDim ThetaAngles(NumberOfProbes)
	ReDim PhiAngles(NumberOfProbes)

	Dim n As Integer
	n = Sqr(NumberOfProbes)
	idx = 0
	For i = 1 To n
		For k = 1 To n
			ThetaAngles(idx) = k*pi/n - pi/(2*n)
			PhiAngles(idx) = i*2*pi/n - pi/n
			Debug.Print ThetaAngles(idx)*180/pi;PhiAngles(idx)*180/pi
			idx = idx + 1
		Next k
	Next i

ElseIf Method = 1 Then
	'Golden Spiral Method
	ReDim ThetaAngles(NumberOfProbes)
	ReDim PhiAngles(NumberOfProbes)
	For i = 0 To NumberOfProbes-1
		ThetaAngles(i) = ACos(1 - 2*(i+0.5)/NumberOfProbes)
		PhiAngles(i) = pi * (1 + Sqr(5)) * (i+0.5)
	Next i
End If


'Clear Old Probes
'ClearProbes()
' Create Probes
CreateProbes(ThetaAngles,PhiAngles,ScannerDistance,NumberOfProbes,Method,Axis)



End Sub



Private Function ACos (x As Variant) As Variant

    Select Case x
        Case -1
            ACos = 4 * Atn(1)

        Case 0:
            ACos = 2 * Atn(1)

        Case 1:
            ACos = 0

        Case Else:
            ACos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
End Function

Private Function ClearProbes()
	While Probe.GetFirst
		Debug.Print Probe.GetCaption
		Probe.Delete (Probe.GetCaption)
		Probe.GetNext
	Wend

End Function


Private Function CreateProbes(ThetaAngles() As Double, PhiAngles() As Double,ScannerDistance As Double,NumberOfProbes As Integer, Method As Integer, Axis As Integer) As Boolean
	Dim i As Integer
	Dim IndexEnd As Integer
	If Method = 0 Then
		IndexEnd = NumberOfProbes-1
	ElseIf Method = 1 Then
		IndexEnd = NumberOfProbes
	End If
	For i = 0 To IndexEnd
		With Probe
			.Reset
			.AutoLabel 1
			.Field ("efield")
			.SetCoordinateSystemType ("cartesian")
			If Axis = 0 Then
				' X-Axis Reference
				.SetPosition2 (ScannerDistance*Sin(ThetaAngles(i))*Cos(PhiAngles(i)))
				.SetPosition3 (ScannerDistance*Sin(ThetaAngles(i))*Sin(PhiAngles(i)))
				.SetPosition1 (ScannerDistance*Cos(ThetaAngles(i)))

			ElseIf Axis = 1 Then
				' Y-Axis Reference
				.SetPosition3 (ScannerDistance*Sin(ThetaAngles(i))*Cos(PhiAngles(i)))
				.SetPosition1 (ScannerDistance*Sin(ThetaAngles(i))*Sin(PhiAngles(i)))
				.SetPosition2 (ScannerDistance*Cos(ThetaAngles(i)))

			ElseIf Axis = 2 Then
				' Z-Axis Reference

				.SetPosition1 (ScannerDistance*Sin(ThetaAngles(i))*Cos(PhiAngles(i)))
				.SetPosition2 (ScannerDistance*Sin(ThetaAngles(i))*Sin(PhiAngles(i)))
				.SetPosition3 (ScannerDistance*Cos(ThetaAngles(i)))
			End If
			.Orientation ("All")
			.Origin ("zero")
			.Create
		End With
	Next

End Function



