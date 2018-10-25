'#Language "WWB-COM"

Option Explicit

Sub Main

Dim x() As Double
Dim y() As Double
Dim z() As Double
Dim phi() As Double
Dim theta() As Double
Dim r() As Double

Dim idx As Integer
idx = 0

Dim NumberOfProbes As Integer
NumberOfProbes = 1

'Get Number of Probes
Probe.GetFirst
While Probe.GetNext
	NumberOfProbes = NumberOfProbes + 1
Wend


' Get Coordinates and E Field of all Probes
ReDim x(NumberOfProbes-1)
ReDim y(NumberOfProbes-1)
ReDim z(NumberOfProbes-1)
ReDim phi(NumberOfProbes-1)
ReDim theta(NumberOfProbes-1)
ReDim r(NumberOfProbes-1)

Dim dataAbs As Variant
Dim dataX As Variant
Dim dataY As Variant
Dim dataZ As Variant
Dim caption As String

' Extract Near Field at Frequencies where we have a far-field monitor
Dim NumberOfFrequencies As Integer
NumberOfFrequencies = Monitor.GetNumberOfMonitors

Dim frequency() As Double
ReDim frequency(NumberOfFrequencies-1)
Dim k As Integer
For k = 0 To NumberOfFrequencies-1
	frequency(k) = Monitor.GetMonitorFrequencyFromIndex(k)
Next k

Dim exportDir As String
exportDir = GetProjectPath("Project") & "\Export\ProbedNearField"

If Len(Dir(exportDir,vbDirectory)) = 0 Then
	MkDir exportDir
End If




' Iterate through all frequencies
Dim f As Double
For Each f In frequency
	'''''''''''''''''''''''''''
	' Export Near Field Results
	'''''''''''''''''''''''''''

	' Text File to Store it into
	Open exportDir & "\NearFieldProbeResults" & f & "GHz.txt" For Output As #1
	' File Header
	'Print #1,"X";"	";"Y";"	";"Z";"	";"R";"	";"Theta";"	";"Phi";"	";"ExReal";"	";"ExImg";"	";"EyReal";"	";"EyImg";"	";"EzReal";"	";"EzImg";"	";"EabsReal";"	";"EabsImg"
	Print #1,"X";"	";"Y";"	";"Z";"	";"ExReal";"	";"ExImg";"	";"EyReal";"	";"EyImg";"	";"EzReal";"	";"EzImg";"	";"EabsReal";"	";"EabsImg"

	' Iterate through Probes
	Probe.GetFirst
	For idx = 0 To NumberOfProbes-1
		' Coordinates
		x(idx) = Probe.GetPosition1
		y(idx) = Probe.GetPosition2
		z(idx) = Probe.GetPosition3
		caption = Probe.GetCaption

		' Calculate Spherical Coordinates
		'r(idx) = Sqr(x(idx)^2 + y(idx)^2 + z(idx)^2)
		'theta(idx) = ACos(z(idx)/r(idx))
		'phi(idx) = Atn2(y(idx),x(idx))

		' Read E-Field Data at given frequency
		dataAbs = Resulttree.GetResultFromTreeItem("1D Results\Probes\E-Field\" & caption & "(Abs) [1]","3D:RunID:0")
		Dim n As Integer
		While Abs(dataAbs.GetX(n)-f)>= 0.002
			n = n+1
		Wend
		dataX = Resulttree.GetResultFromTreeItem("1D Results\Probes\E-Field\" & caption & "(X) [1]","3D:RunID:0")
		dataY = Resulttree.GetResultFromTreeItem("1D Results\Probes\E-Field\" & caption & "(Y) [1]","3D:RunID:0")
		dataZ = Resulttree.GetResultFromTreeItem("1D Results\Probes\E-Field\" & caption & "(Z) [1]","3D:RunID:0")

		' Print Values to File
		'Print #1,x(idx);"	";y(idx);"	";z(idx);"	";r(idx);"	";theta(idx);"	";phi(idx);"	";dataX.GetYRe(n);"	";dataX.GetYIm(n);"	";dataY.GetYRe(n);"	";dataY.GetYIm(n);"	";dataZ.GetYRe(n);"	";dataZ.GetYIm(n);"	";dataAbs.GetYRe(n);"	";dataAbs.GetYIm(n)
		Print #1,x(idx);"	";y(idx);"	";z(idx);"	";dataX.GetYRe(n);"	";dataX.GetYIm(n);"	";dataY.GetYRe(n);"	";dataY.GetYIm(n);"	";dataZ.GetYRe(n);"	";dataZ.GetYIm(n);"	";dataAbs.GetYRe(n);"	";dataAbs.GetYIm(n)
		Probe.GetNext
	Next idx
	Close #1

	'''''''''''''''''''''''''''
	' Export Far Field Results
	'''''''''''''''''''''''''''
	'Dim farFieldData As Variant
	'farFieldData = Resulttree.GetResultFromTreeItem("Farfields\farfield (f=" & f & ")[1]\Abs","3D:RunID:0")
	'farFieldData = Resulttree.GetResultFromTreeItem("Farfields\farfield (f=" & f & ") [1]\Abs","3D:RunID:0")

Next f




	
End Sub

Public Function Atn2(y As Double, x As Double) As Double

  If x > 0 Then

    Atn2 = Atn(y / x)

  ElseIf x < 0 Then

    Atn2 = Sgn(y) * (Pi - Atn(Abs(y / x)))

  ElseIf y = 0 Then

    Atn2 = 0

  Else

    Atn2 = Sgn(y) * Pi / 2

  End If

End Function

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
