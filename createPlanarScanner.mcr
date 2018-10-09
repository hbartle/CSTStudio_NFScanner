'#Language "WWB-COM"

Option Explicit

Sub Main

' Definitions
Dim axisoptions$(2)
axisoptions$(0) = "X-Axis"
axisoptions$(1) = "Y-Axis"
axisoptions$(2) = "Z-Axis"
Dim Axis As Integer

Dim NumberOfProbesInWidthDim As Integer
Dim NumberOfProbesInHeightDim As Integer
Dim ScannerWidthSpacing As Double
Dim ScannerHeightSpacing As Double
Dim ScannerDistance As Double

NumberOfProbesInWidthDim = 10
NumberOfProbesInHeightDim = 10
ScannerWidthSpacing = 10.0
ScannerHeightSpacing = 10.0
ScannerDistance = 300.0


'User Dialog For Planar Scanner Parameters
	Begin Dialog UserDialog 520,250 ' %GRID:20,10,1,1
		TextBox 160,50,160,20,.NumberOfProbesInWidthDim
		TextBox 160,90,160,20,.NumberOfProbesInHeightDim
		TextBox 160,130,160,20,.ScannerWidthSpacing
		TextBox 160,170,160,20,.ScannerHeightSpacing
		TextBox 160,210,160,20,.ScannerDistance
		DropListBox 20,21,150,28,axisoptions(),.Axis
		OKButton 380,180,120,20
		CancelButton 380,210,120,20
		Text 20,50,130,30,"Number Of Probes in Width Dimension",.Text1
		Text 20,90,130,30,"Number Of Probes in Height Dimension",.Text2
		Text 20,130,130,30,"Spacing in Width Dimension",.Text3
		Text 20,170,130,30,"Spacing in Width Dimension",.Text4
		Text 20,210,130,30,"Scanner Distance",.Text5
	End Dialog
Dim dlg As UserDialog
dlg.NumberOfProbesInWidthDim=CStr(NumberOfProbesInWidthDim)
dlg.NumberOfProbesInHeightDim=CStr(NumberOfProbesInHeightDim)
dlg.ScannerWidthSpacing=CStr(ScannerWidthSpacing)
dlg.ScannerHeightSpacing=CStr(ScannerHeightSpacing)
dlg.ScannerDistance=CStr(ScannerDistance)


If (Dialog(dlg) = 0) Then Exit All



' Save User Input
NumberOfProbesInWidthDim = Evaluate(dlg.NumberOfProbesInWidthDim) - 1
NumberOfProbesInHeightDim =  Evaluate(dlg.NumberOfProbesInHeightDim) -1
ScannerWidthSpacing = Evaluate(dlg.ScannerWidthSpacing)
ScannerHeightSpacing =  Evaluate(dlg.ScannerHeightSpacing)
ScannerDistance =  Evaluate(dlg.ScannerDistance)
Axis = Evaluate(dlg.Axis)


' Calculate Probe Positions
Dim WidthCoordinates() As Double
Dim HeightCoordinates() As Double

ReDim WidthCoordinates(NumberOfProbesInWidthDim)
ReDim HeightCoordinates(NumberOfProbesInHeightDim)

Dim i As Integer
For i = 0 To NumberOfProbesInWidthDim
	WidthCoordinates(i) = (i* NumberOfProbesInWidthDim - NumberOfProbesInWidthDim^2/2 )*ScannerWidthSpacing'
Next i
For i  = 0 To NumberOfProbesInHeightDim
	HeightCoordinates(i) = (i* NumberOfProbesInHeightDim -NumberOfProbesInHeightDim^2/2)*ScannerHeightSpacing
Next i



' Create Probes
CreateProbes(WidthCoordinates,HeightCoordinates,ScannerDistance, Axis)



End Sub

Private Function ClearProbes()
	While Probe.GetFirst
		Debug.Print Probe.GetCaption
		Probe.Delete (Probe.GetCaption)
		Probe.GetNext
	Wend

End Function


Private Function CreateProbes(WidthCoordinates() As Double, HeightCoordinates() As Double,ScannerDistance As Double,Axis As Integer) As Boolean
	Dim W As Double
	Dim h As Double
	For Each W In WidthCoordinates
		For Each h In HeightCoordinates
			With Probe
				.Reset
				.AutoLabel 1
				.Field ("efield")
				.SetCoordinateSystemType ("cartesian")
				If Axis = 0 Then
					.SetPosition1 (ScannerDistance)
					.SetPosition2 (W)
					.SetPosition3 (h)
				ElseIf Axis = 1 Then
					.SetPosition1 (W)
					.SetPosition2 (ScannerDistance)
					.SetPosition3 (h)
				ElseIf Axis = 2 Then
					.SetPosition1 (W)
					.SetPosition2 (h)
					.SetPosition3 (ScannerDistance)
				End If
				.Orientation ("All")
				.Origin ("zero")
				.Create
			End With
		Next
	Next
End Function



