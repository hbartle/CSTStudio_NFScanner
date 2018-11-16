'#Language "WWB-COM"

Option Explicit

Sub Main

' Definitions
Dim ScannerRadius As Double
Dim NumberOfProbesPhi As Integer
Dim NumberOfProbesHeight As Integer
Dim ScannerHeightSpacing As Double
Dim axisoptions$(2)
axisoptions$(0) = "Y-Axis"
axisoptions$(1) = "X-Axis"
axisoptions$(2) = "Z-Axis"
Dim Axis As Integer

ScannerRadius = 200
NumberOfProbesPhi = 10
ScannerHeightSpacing = 20
NumberOfProbesHeight = 20

'User Dialog For Planar Scanner Parameters
	Begin Dialog UserDialog 540,210 ' %GRID:10,10,1,1
		TextBox 160,80,160,20,.ScannerRadius
		TextBox 160,110,160,20,.NumberOfProbesPhi
		TextBox 160,140,160,20,.ScannerHeightSpacing
		TextBox 160,170,160,20,.NumberOfProbesHeight
		OKButton 380,60,120,20
		CancelButton 380,90,120,20
		Text 20,80,140,20,"Scanner Radius",.Text1
		Text 20,110,140,30,"Number of Probes in Phi Dimension",.Text2
		Text 20,140,140,20,"Height Spacing",.Text3
		Text 20,170,140,30,"Number of Probes in Height Dimension",.Text4
		DropListBox 20,30,160,20,axisoptions(),.Axis
		Text 20,10,140,10,"Reference Axis",.Text5
	End Dialog
Dim dlg As UserDialog
dlg.ScannerRadius=CStr(ScannerRadius)
dlg.NumberOfProbesPhi=CStr(NumberOfProbesPhi)
dlg.ScannerHeightSpacing=CStr(ScannerHeightSpacing)
dlg.NumberOfProbesHeight=CStr(NumberOfProbesHeight)
If (Dialog(dlg) = 0) Then Exit All

' Save User Input
ScannerRadius = Evaluate(dlg.ScannerRadius)
NumberOfProbesPhi =  Evaluate(dlg.NumberOfProbesPhi)
ScannerHeightSpacing =  Evaluate(dlg.ScannerHeightSpacing)
NumberOfProbesHeight = Evaluate(dlg.NumberOfProbesHeight)
Axis = Evaluate(dlg.Axis)

' Calculate Probe Positions
Dim PhiAngles() As Double
Dim HeightValues() As Double

ReDim PhiAngles(NumberOfProbesPhi*NumberOfProbesHeight)
ReDim HeightValues(NumberOfProbesPhi*NumberOfProbesHeight)

Dim i As Integer
Dim k As Integer
Dim idx As Integer


idx = 0
For i = 1 To NumberOfProbesHeight
	For k = 1 To NumberOfProbesPhi
		HeightValues(idx) = (i-NumberOfProbesHeight/2)*ScannerHeightSpacing
		PhiAngles(idx) = k*2*pi/NumberOfProbesPhi
		'Debug.Print ThetaAngles(idx)*180/pi;PhiAngles(idx)*180/pi
		idx = idx + 1
	Next k
Next i

CreateProbes(PhiAngles,HeightValues,ScannerRadius,NumberOfProbesPhi,NumberOfProbesHeight,Axis)

End Sub


Private Function CreateProbes(PhiAngles() As Double, HeightValues() As Double, ScannerRadius As Double, NumberOfProbesPhi As Integer, NumberOfProbesHeight As Integer, Axis As Integer) As Boolean
	Dim i As Integer
	For i = 0 To NumberOfProbesPhi*NumberOfProbesHeight-1
		With Probe
			.Reset
			.AutoLabel 1
			.Field ("efield")
			.SetCoordinateSystemType ("cartesian")
			If Axis = 0 Then
				' Y-Axis Reference
				.SetPosition3 (ScannerRadius*Cos(PhiAngles(i)))
				.SetPosition1 (ScannerRadius*Sin(PhiAngles(i)))
				.SetPosition2 (HeightValues(i))

			ElseIf Axis = 1 Then
				' X-Axis Reference
				.SetPosition2 (ScannerRadius*Cos(PhiAngles(i)))
				.SetPosition3 (ScannerRadius*Sin(PhiAngles(i)))
				.SetPosition1 (HeightValues(i))

			ElseIf Axis = 2 Then
				' Z-Axis Reference

				.SetPosition1 (ScannerRadius*Cos(PhiAngles(i)))
				.SetPosition2 (ScannerRadius*Sin(PhiAngles(i)))
				.SetPosition3 (HeightValues(i))
			End If
			.Orientation ("All")
			.Origin ("zero")
			.Create
		End With
	Next

End Function
