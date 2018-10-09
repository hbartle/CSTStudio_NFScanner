'#Language "WWB-COM"

Option Explicit

Sub Main

Dim x() As Double
Dim y() As Double
Dim z() As Double
Dim idx As Integer
idx = 0
Dim NumberOfProbes As Integer
NumberOfProbes = 1

'Get Number of Probes
Probe.GetFirst
While Probe.GetNext
	NumberOfProbes = NumberOfProbes + 1
Wend


' Get Cartesian Coordinates and E Field of all Probes
ReDim x(NumberOfProbes-1)
ReDim y(NumberOfProbes-1)
ReDim z(NumberOfProbes-1)

Dim p As Object
Dim filename As String

Probe.GetFirst

For idx = 0 To NumberOfProbes-1
	' Coordinates
	x(idx) = Probe.GetPosition1
	y(idx) = Probe.GetPosition2
	z(idx) = Probe.GetPosition3
	Debug.Print Probe.GetCaption
	filename = Resulttree.GetFileFromTreeItem("1D Results\Probes")

	If filename = "" Then
		ReportInformationToWindow("Result does not exist.")
	Else
	Set p = Result1D("E-Field (-80.0547 110.186 267.302)(Abs)[1]")
	'Debug.Print p
	Probe.GetNext
Next idx



	
End Sub
