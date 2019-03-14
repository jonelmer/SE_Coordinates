Imports SolidEdgeCommunity
Imports SolidEdgeCommunity.Extensions ' https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Friend Class Program
    <STAThread>
    Shared Sub Main(ByVal args() As String)
        Dim application As SolidEdgeFramework.Application = Nothing
        'Dim assemblyDocument As SolidEdgeAssembly.AssemblyDocument = Nothing
        Dim assemblyDocument As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim occurrences As SolidEdgePart.CoordinateSystems = Nothing

        Try
            ' Register with OLE to handle concurrency issues on the current thread.
            SolidEdgeCommunity.OleMessageFilter.Register()

            ' Connect to or start Solid Edge.
            application = SolidEdgeCommunity.SolidEdgeUtils.Connect(True, True)

            ' Get the active document.
            assemblyDocument = application.GetActiveDocument(Of SolidEdgeFramework.SolidEdgeDocument)(False)

            If assemblyDocument IsNot Nothing Then
                ' Get a reference to the Occurrences collection.
                occurrences = assemblyDocument.CoordinateSystems

                Dim output As String = "Name, XOffset, YOffset, ZOffset, XRotation, YRotation, ZRotation" + vbNewLine

                For Each occurrence In occurrences.OfType(Of SolidEdgePart.CoordinateSystem)()
                    ' Allocate a new array to hold transform.
                    Dim transform(5) As Double

                    ' Get the occurrence transform.
                    occurrence.GetOrientation(transform(0), transform(1), transform(2), transform(3), transform(4), transform(5))

                    ' Report the occurrence transform.
                    Console.WriteLine("'{0}'", occurrence.Name)
                    Console.WriteLine("XOffset (m):      {0}", transform(0))
                    Console.WriteLine("YOffset (m):      {0}", transform(1))
                    Console.WriteLine("ZOffset (m):      {0}", transform(2))
                    Console.WriteLine("XRotation (rads): {0}", transform(3))
                    Console.WriteLine("YRotation (rads): {0}", transform(4))
                    Console.WriteLine("ZRotation (rads): {0}", transform(5))
                    Console.WriteLine()

                    output = output + """" + occurrence.Name + """, "

                    output = output + transform(0).ToString + ", "
                    output = output + transform(1).ToString + ", "
                    output = output + transform(2).ToString + ", "
                    output = output + transform(3).ToString + ", "
                    output = output + transform(4).ToString + ", "
                    output = output + transform(5).ToString + vbNewLine
                    ' Console.WriteLine(output)

                Next occurrence

                My.Computer.Clipboard.SetText(output)

                MsgBox("Coordinate system data copied to clipboard")

            Else
                Throw New System.Exception("No active document.")
            End If
        Catch ex As System.Exception
            Console.WriteLine(ex.Message)
        Finally
            SolidEdgeCommunity.OleMessageFilter.Unregister()
        End Try
    End Sub
End Class
