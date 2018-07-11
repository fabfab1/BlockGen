'Class for handeling surface conversions

Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Runtime.InteropServices
Imports Autodesk.AutoCAD.ApplicationServices

Public Class Surfaces
    Private TopPointList As Point3dCollection
    Private BottomPointList As Point3dCollection
    Private SurfaceObjs As New List(Of ObjectId)

    '<DllImport("acad.exe", CallingConvention:=CallingConvention.Cdecl, EntryPoint:="acedCmd")> _
    'Private Shared Function acedCmd(pResbuf As IntPtr) As Integer
    'End Function

    Public Sub New(ByVal pTopPointList As Point3dCollection, ByVal pBottomPointList As Point3dCollection)
        TopPointList = pTopPointList
        BottomPointList = pBottomPointList
    End Sub

    Public Sub GenerateEdgeSurface(ByVal aIndex As Integer, ByVal aTransaction As Transaction, ByVal aModelSpace As BlockTableRecord)
        Dim f As New Face(BottomPointList(aIndex), BottomPointList(aIndex + 1), TopPointList(aIndex + 1), TopPointList(aIndex), True, True, True, True)
        Dim surf As Autodesk.AutoCAD.DatabaseServices.Surface = Autodesk.AutoCAD.DatabaseServices.Surface.CreateFrom(f)

        SurfaceObjs.Add(aModelSpace.AppendEntity(surf))
        aTransaction.AddNewlyCreatedDBObject(surf, True)
    End Sub

    Public Sub GenerateSurface(ByVal aTransaction As Transaction, ByVal aModelSpace As BlockTableRecord)
        For index As Integer = 0 To TopPointList.Count - 2
            GenerateEdgeSurface(index, aTransaction, aModelSpace)
        Next

        Dim topTriangles As List(Of Triangle) = Tessellate(TopPointList)
        Dim bottomTriangles As List(Of Triangle) = Tessellate(BottomPointList)

        For Each t As Triangle In topTriangles
            t.GenerateSurface(aTransaction, aModelSpace, SurfaceObjs)
        Next

        For Each t As Triangle In bottomTriangles
            t.GenerateSurface(aTransaction, aModelSpace, SurfaceObjs)
        Next

        ' best method for creating a solid from multiple surfaces
        Using sol As New Solid3d()
            Dim surfaces As New List(Of Entity)
            For Each oid As ObjectId In SurfaceObjs
                Dim ent As Entity = aTransaction.GetObject(oid, OpenMode.ForWrite)
                If Not ent = Nothing Then
                    surfaces.Add(ent)
                End If
            Next

            Dim flags As New IntegerCollection()
            sol.CreateSculptedSolid(surfaces.ToArray(), flags)

            For Each ent As Entity In surfaces
                ent.Erase()
            Next
            aModelSpace.AppendEntity(sol)
            aTransaction.AddNewlyCreatedDBObject(sol, True)
        End Using

        ' Alternate method that works, but not quite as well
        'SurfaceObjs.Reverse()

        'Dim surf As Autodesk.AutoCAD.DatabaseServices.Surface = aTransaction.GetObject(SurfaceObjs(0), OpenMode.ForWrite, False)
        'Dim nsurf As Autodesk.AutoCAD.DatabaseServices.Surface = Nothing

        'SurfaceObjs.RemoveAt(0)

        'Try
        '    For Each oid As ObjectId In SurfaceObjs
        '        Dim s As Autodesk.AutoCAD.DatabaseServices.Surface = aTransaction.GetObject(oid, OpenMode.ForWrite, False)
        '        nsurf = surf.BooleanUnion(s)
        '        s.Erase()
        '        If nsurf = Nothing Then
        '            Continue For
        '        End If
        '        aModelSpace.AppendEntity(nsurf)
        '        aTransaction.AddNewlyCreatedDBObject(nsurf, True)
        '        surf.Erase()
        '        surf = nsurf
        '    Next

        '    Try
        '        Dim sol As New Solid3d()
        '        sol.CreateFrom(surf)
        '        aModelSpace.AppendEntity(sol)
        '        aTransaction.AddNewlyCreatedDBObject(sol, True)
        '        surf.Erase()
        '    Catch ex As Exception
        '        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf + "WallGen ERROR: Unable to convert surface to solid. Look for self intersection.")
        '    End Try
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    Private Function asPnt2d(ByVal aPoint As Point3d) As Point2d
        Return New Point2d(aPoint.X, aPoint.Y)
    End Function

    Private Function ContainsEdge(ByVal aLine As LineSegment2d, ByVal aLineList As List(Of LineSegment2d)) As Boolean
        Try
            For Each ln As LineSegment2d In aLineList
                If aLine.StartPoint.IsEqualTo(ln.StartPoint) AndAlso aLine.EndPoint.IsEqualTo(ln.EndPoint) Then
                    Return True
                End If
                If aLine.StartPoint.IsEqualTo(ln.EndPoint) AndAlso aLine.EndPoint.IsEqualTo(ln.StartPoint) Then
                    Return True
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return False
    End Function

    Private Function IsLineInside(ByVal aLine As LineSegment2d, ByVal aLineList As List(Of LineSegment2d), ByVal aPointList As Point3dCollection, ByRef isEdge As Boolean) As Boolean
        Dim tol As New Tolerance(0.1, 0.1)
        Dim mid As Point2d = aLine.MidPoint
        isEdge = False

        ' if the test line is an edge segment, it's not inside
        If ContainsEdge(aLine, aLineList) Then
            isEdge = True
            Return False
        End If

        ' if the above point isn't contained within the polygon, then it's assured the line is
        ' at least partially outside the polygon
        If Not PointInPolygon(mid, aPointList) Then
            Return False
        End If

        For Each ln As LineSegment2d In aLineList
            Dim ipoints() As Point2d = ln.IntersectWith(aLine, tol)

            If ipoints Is Nothing Then
                Continue For
            End If

            ' count non end points
            Dim icount As Integer = 0
            For Each pt As Point2d In ipoints
                If pt.IsEqualTo(aLine.EndPoint, tol) OrElse pt.IsEqualTo(aLine.StartPoint, tol) Then
                    Continue For
                End If
                icount += 1
            Next

            ' if the test line intersects with any other line it can't be used since at least
            ' part of it is outside.
            If icount > 0 Then
                Return False
            End If
        Next

        Return True
    End Function

    Private Function Tessellate(ByVal aPointList As Point3dCollection) As List(Of Triangle)
        Dim lines As New List(Of LineSegment2d)
        For i As Integer = 0 To aPointList.Count - 1
            lines.Add(New LineSegment2d(asPnt2d(aPointList(i)), asPnt2d(aPointList(i + 1 Mod aPointList.Count))))
        Next

        Dim pivot As Integer = 0
        Dim epoint As Integer = 1
        Dim i2 As Integer = 0
        Dim i3 As Integer = 0
        Dim isEdge = False
        Dim Triangles As New List(Of Triangle)
        Dim stopAt As Integer = -1
        Dim stopAts As New Dictionary(Of Integer, Integer)

        While pivot >= 0
            Dim testline As New LineSegment2d(asPnt2d(aPointList(pivot)), asPnt2d(aPointList(epoint Mod aPointList.Count)))
            If Not IsLineInside(testline, lines, aPointList, isEdge) Then
                If isEdge AndAlso i2 = pivot Then ' represents the first side of the triangle
                    i2 = epoint
                ElseIf isEdge AndAlso i2 = i3 Then ' represents the second side of the triangle
                    i3 = epoint
                    epoint -= 1 ' need to cancel out indexing of epoint next cycle will take care of it
                End If
                epoint += 1
                If epoint >= aPointList.Count Then
                    Exit While
                ElseIf epoint < aPointList.Count - 2 Then ' only continue while not at the last triangle
                    Continue While
                End If
            ElseIf i2 = pivot Then ' test line is inside polygon but don't have a 3rd point yet so keep going
                i2 = epoint
                epoint += 1
                If epoint > aPointList.Count Then
                    Exit While
                End If
                Continue While
            End If

            ' if we get here, epoint should close the triangle
            i3 = epoint
            Triangles.Add(New Triangle(aPointList(pivot), aPointList(i2), aPointList(i3)))
            If (i3 - i2) > 1 Then 'means there was a gap and we have missing triangles
                stopAt = i3
                If Not stopAts.ContainsKey(stopAt) Then
                    stopAts(stopAt) = pivot
                End If
                pivot = i2 'temporarily shift pivot until reaching stop at
                epoint = pivot + 1
                If i3 = aPointList.Count Then
                    epoint = pivot + 1
                    i2 = pivot
                    stopAt = -1 'we are at the end of the list so this no longer applies
                End If
            Else
                i2 = i3
                epoint = i3 + 1

                If i3 = stopAt Then
                    If Not stopAts.TryGetValue(stopAt, pivot) Then ' shift pivot back
                        stopAt = -1
                    End If
                End If
            End If
            If epoint >= aPointList.Count - 1 Then
                Exit While
            End If
        End While
        Return Triangles
    End Function

    'Only works for flat polygons parallel to the world plane
    Private Function PointInPolygon(ByVal apoint As Point3d, ByVal aPointList As Point3dCollection) As Boolean
        Return PointInPolygon(New Point2d(apoint.X, apoint.Y), aPointList)
    End Function

    Private Function PointInPolygon(ByVal aPoint As Point2d, ByVal aPointList As Point3dCollection) As Boolean
        Dim polySides As Integer = aPointList.Count
        Dim polyX(polySides) As Double
        Dim polyY(polySides) As Double
        Dim i As Integer = 0
        Dim j As Integer = polySides - 1
        Dim oddNodes As Boolean = False
        Dim x As Double = aPoint.X
        Dim y As Double = aPoint.Y

        For Each vtx As Point3d In aPointList
            polyX(i) = vtx.X
            polyY(i) = vtx.Y
            i += 1
        Next

        For i = 0 To polySides - 1
            If (polyY(i) < y AndAlso polyY(j) >= y) OrElse (polyY(j) < y AndAlso polyY(i) >= y) Then
                If polyX(i) + (y - polyY(i)) / (polyY(j) - polyY(i)) * (polyX(j) - polyX(i)) < x Then
                    oddNodes = Not oddNodes
                End If
            End If
            j = i
        Next

        Return oddNodes
    End Function

End Class
