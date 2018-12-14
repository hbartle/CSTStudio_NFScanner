'#Language "WWB-COM"

Option Explicit

Sub Main

' Definitions
Dim axisoptions$(2)
axisoptions$(0) = "Z-Axis"
axisoptions$(1) = "Y-Axis"
axisoptions$(2) = "X-Axis"
Dim Axis As Integer

Dim NumberOfProbesInWidthDim As Integer
Dim NumberOfProbesInHeightDim As Integer
Dim ScannerWidthSpacing As Double
Dim ScannerHeightSpacing As Double
Dim ScannerDistance As Double

' Variances
Dim WidthVariance As Double
Dim HeightVariance As Double
Dim DistanceVariance As Double

NumberOfProbesInWidthDim = 30
NumberOfProbesInHeightDim = 30
ScannerWidthSpacing = 30.0
ScannerHeightSpacing = 30.0
ScannerDistance = 300.0
WidthVariance = 0.0
HeightVariance = 0.0
DistanceVariance = 0.0

'User Dialog For Planar Scanner Parameters
	Begin Dialog UserDialog 800,310 ' %GRID:20,10,1,1
		GroupBox 380,40,320,130,"Position Variance",.GroupBox2
		GroupBox 20,40,340,250,"Dimensions",.GroupBox1
		TextBox 180,90,160,20,.NumberOfProbesInWidthDim
		TextBox 180,130,160,20,.NumberOfProbesInHeightDim
		TextBox 180,170,160,20,.ScannerWidthSpacing
		TextBox 180,210,160,20,.ScannerHeightSpacing
		TextBox 180,250,160,20,.ScannerDistance
		TextBox 520,70,160,20,.WidthVariance
		TextBox 520,100,160,20,.HeightVariance
		TextBox 520,130,160,20,.DistanceVariance
		DropListBox 40,60,160,30,axisoptions(),.Axis
		OKButton 600,240,120,20
		CancelButton 600,270,120,20
		Text 40,90,140,30,"Number Of Probes in Width Dimension",.Text1
		Text 400,70,120,20,"Width Dimension",.Text6
		Text 400,100,120,20,"Height Dimension",.Text7
		Text 400,130,120,30,"Distance Dimension",.Text8
		Text 40,130,140,30,"Number Of Probes in Height Dimension",.Text2
		Text 40,170,140,30,"Spacing in Width Dimension ",.Text3
		Text 40,210,140,30,"Spacing in Height Dimension ",.Text4
		Text 40,250,140,30,"Scanner Distance ",.Text5
		Text 20,10,180,20,"Planar NF Scanner Setup",.Text9
	End Dialog
Dim dlg As UserDialog
dlg.NumberOfProbesInWidthDim=CStr(NumberOfProbesInWidthDim)
dlg.NumberOfProbesInHeightDim=CStr(NumberOfProbesInHeightDim)
dlg.ScannerWidthSpacing=CStr(ScannerWidthSpacing)
dlg.ScannerHeightSpacing=CStr(ScannerHeightSpacing)
dlg.ScannerDistance=CStr(ScannerDistance)
dlg.WidthVariance=CStr(WidthVariance)
dlg.HeightVariance=CStr(HeightVariance)
dlg.DistanceVariance=CStr(DistanceVariance)

If (Dialog(dlg) = 0) Then Exit All



' Save User Input
NumberOfProbesInWidthDim = Evaluate(dlg.NumberOfProbesInWidthDim) - 1
NumberOfProbesInHeightDim =  Evaluate(dlg.NumberOfProbesInHeightDim) -1
ScannerWidthSpacing = Evaluate(dlg.ScannerWidthSpacing)
ScannerHeightSpacing =  Evaluate(dlg.ScannerHeightSpacing)
ScannerDistance =  Evaluate(dlg.ScannerDistance)
Axis = Evaluate(dlg.Axis)
WidthVariance = Evaluate(dlg.WidthVariance)
HeightVariance = Evaluate(dlg.HeightVariance)
DistanceVariance = Evaluate(dlg.DistanceVariance)

' Calculate Probe Positions
Dim WidthCoordinates() As Double
Dim HeightCoordinates() As Double

ReDim WidthCoordinates(NumberOfProbesInWidthDim)
ReDim HeightCoordinates(NumberOfProbesInHeightDim)

Dim i As Integer
For i = 0 To NumberOfProbesInWidthDim
	WidthCoordinates(i) = (i - NumberOfProbesInWidthDim/2 )*ScannerWidthSpacing'
Next i
For i  = 0 To NumberOfProbesInHeightDim
	HeightCoordinates(i) = (i - NumberOfProbesInHeightDim/2)*ScannerHeightSpacing
Next i




' Create Probes
CreateProbes(WidthCoordinates,HeightCoordinates,ScannerDistance, Axis, WidthVariance,HeightVariance,DistanceVariance)



End Sub

Private Function ClearProbes()
	While Probe.GetFirst
		Debug.Print Probe.GetCaption
		Probe.Delete (Probe.GetCaption)
		Probe.GetNext
	Wend

End Function


Private Function CreateProbes(WidthCoordinates() As Double, HeightCoordinates() As Double,ScannerDistance As Double,Axis As Integer, WidthVariance As Double,HeightVariance As Double,DistanceVariance As Double) As Boolean

	Dim w As Double
	Dim h As Double
	Dim w_noise As Double
	Dim h_noise As Double
	Dim d_noise As Double
	Dim u1 As Double
	Dim u2 As Double
	Dim z As Double
	Dim toggle_file_write As Boolean

	' Check if normal distribution should be applied
	If WidthVariance <> 0.0 Or HeightVariance <> 0.0 Or DistanceVariance <> 0.0 Then
		toggle_file_write = True
	End If

	' Create Export Directory for real probe positions txt file
	Dim exportDir As String
	exportDir = GetProjectPath("Project") & "\Export\ProbedNearField"

	If Len(Dir(exportDir,vbDirectory)) = 0 Then
		MkDir exportDir
	End If

	If toggle_file_write = True Then
		Open exportDir & "\ProbePositions.txt" For Output As #1
		' File Header
		Print #1,"X";"	";"Y";"	";"Z"
	End If

	For Each w In WidthCoordinates
		For Each h In HeightCoordinates
			' Apply Normal Distribution to coordinates
			If WidthVariance = 0.0 Then
				w_noise = 0
			Else
				' Create normal distributed noise using Box-Muller transform
				Randomize
				u1 = Rnd()
				u2 = Rnd()
				z = Sqr(-2*Log(u1))*Cos(2*pi*u2)
				w_noise = WidthVariance*z
			End If
			If HeightVariance = 0.0 Then
				h_noise = 0
			Else
				' Create normal distributed noise using Box-Muller transform
				Randomize
				u1 = Rnd()
				u2 = Rnd()
				z = Sqr(-2*Log(u1))*Cos(2*pi*u2)
				h_noise = HeightVariance*z
			End If
			If DistanceVariance =0.0 Then
				d_noise = 0
			Else
				' Create normal distributed noise using Box-Muller transform
				Randomize
				u1 = Rnd()
				u2 = Rnd()
				z = Sqr(-2*Log(u1))*Cos(2*pi*u2)
				d_noise = DistanceVariance*z
			End If

			' Create the Probe
			With Probe
				.Reset
				.AutoLabel 1
				.Field ("efield")
				.SetCoordinateSystemType ("cartesian")
				If Axis = 0 Then
					.SetPosition1 (w+w_noise)
					.SetPosition2 (h+h_noise)
					.SetPosition3 (ScannerDistance+d_noise)
				ElseIf Axis = 1 Then
					.SetPosition1 (w+w_noise)
					.SetPosition2 (ScannerDistance+d_noise)
					.SetPosition3 (h+h_noise)
				ElseIf Axis = 2 Then
					.SetPosition1 (ScannerDistance+d_noise)
					.SetPosition2 (w+w_noise)
					.SetPosition3 (h+h_noise)
				End If
				.Orientation ("All")
				.Origin ("zero")
				.Create
			End With

			' Save actual position to text file
			If toggle_file_write = True Then
				Print #1,w;"	";h;"	";ScannerDistance;"
			End If
		Next
	Next
	If toggle_file_write = True Then
		Close #1
	End If

End Function



