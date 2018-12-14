'#Language "WWB-COM"

Option Explicit

Sub Main

' Definitions

Dim NumberOfProbesTheta As Integer
Dim NumberOfProbesPhi As Integer
Dim ScannerDistance As Double
Dim axisoptions$(2)
axisoptions$(0) = "X-Axis"
axisoptions$(1) = "Y-Axis"
axisoptions$(2) = "Z-Axis"
Dim Axis As Integer

NumberOfProbesTheta = 30
NumberOfProbesPhi = 30
ScannerDistance = 300.0


'User Dialog For Planar Scanner Parameters
	Begin Dialog UserDialog 540,210 ' %GRID:20,10,1,1
		TextBox 160,110,160,20,.NumberOfProbesTheta
		TextBox 160,150,160,20,.NumberOfProbesPhi
		TextBox 160,70,160,20,.ScannerDistance
		OKButton 380,60,120,20
		CancelButton 380,90,120,20
		Text 20,100,140,30,"Number Of Probes in Theta Dimension",.Text1
		Text 20,140,140,30,"Number Of Probes in Phi Dimension",.Text5
		Text 20,70,140,20,"Scanner Distance",.Text2
		DropListBox 20,30,160,20,axisoptions(),.Axis
		Text 20,10,140,10,"Reference Axis",.Text4
	End Dialog
Dim dlg As UserDialog
dlg.NumberOfProbesTheta=CStr(NumberOfProbesTheta)
dlg.NumberOfProbesPhi=CStr(NumberOfProbesPhi)
dlg.ScannerDistance=CStr(ScannerDistance)
If (Dialog(dlg) = 0) Then Exit All

' Save User Input
NumberOfProbesTheta = Evaluate(dlg.NumberOfProbesTheta)
NumberOfProbesPhi = Evaluate(dlg.NumberOfProbesPhi)
ScannerDistance =  Evaluate(dlg.ScannerDistance)
Axis = Evaluate(dlg.Axis)

' Calculate Probe Positions
Dim ThetaAngles() As Double
Dim PhiAngles() As Double

Dim i As Integer
Dim k As Integer
Dim idx As Integer


'Realistic Placing Method
ReDim ThetaAngles(NumberOfProbesPhi*NumberOfProbesTheta)
ReDim PhiAngles(NumberOfProbesPhi*NumberOfProbesTheta)


idx = 0
For i = 0 To NumberOfProbesPhi - 1
	For k = 0 To NumberOfProbesTheta - 1
		ThetaAngles(idx) = (k+1/2) * pi/NumberOfProbesTheta
		PhiAngles(idx) = i*2*pi/NumberOfProbesPhi
		Debug.Print ThetaAngles(idx)*180/pi;PhiAngles(idx)*180/pi
		idx = idx + 1
	Next k
Next i


'Clear Old Probes
'ClearProbes()
' Create Probes
CreateProbes(ThetaAngles,PhiAngles,ScannerDistance,NumberOfProbesTheta,NumberOfProbesPhi,Axis)



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


Private Function CreateProbes(ThetaAngles() As Double, PhiAngles() As Double,ScannerDistance As Double,NumberOfProbesTheta As Integer, NumberOfProbesPhi As Integer, Axis As Integer) As Boolean
	Dim i As Integer
	For i = 0 To NumberOfProbesTheta*NumberOfProbesPhi-1
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



