'Class for handeling face information and cutting

Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices

Public Class Faces
    Private m_Bottom As Plane
    Private m_Top As Plane
    Private m_Left As Plane
    Private m_Right As Plane
    Private m_Front As Plane
    Private m_Back As Plane
    Private m_PointList As Point3dCollection

    'The following is a translation from the entered points as described in
    'the "plug-in requriements.pdf" received from Federico Buccellati on July 3, 2007
    'to the individual faces of the solid with the "bottom" and "top" being
    'parallel to the world plane in AutoCAD.
    Private m_BottomPointIndexes() As Integer = {4, 0, 1, 5}
    Private m_TopPointIndexes() As Integer = {2, 3, 7, 6}
    Private m_LeftPointIndexes() As Integer = {3, 0, 4, 7}
    Private m_RightPointIndexes() As Integer = {5, 1, 2, 6}
    Private m_FrontPointIndexes() As Integer = {1, 0, 3, 2}
    Private m_BackPointIndexes() As Integer = {7, 4, 5, 6}

    Public Sub New(ByVal aPointList As Point3dCollection)
        m_PointList = aPointList
        m_Bottom = New Plane(aPointList(m_BottomPointIndexes(0)), aPointList(m_BottomPointIndexes(1)), aPointList(m_BottomPointIndexes(2)))
        m_Top = New Plane(aPointList(m_TopPointIndexes(0)), aPointList(m_TopPointIndexes(1)), aPointList(m_TopPointIndexes(2)))
        m_Left = New Plane(aPointList(m_LeftPointIndexes(0)), aPointList(m_LeftPointIndexes(1)), aPointList(m_LeftPointIndexes(2)))
        m_Right = New Plane(aPointList(m_RightPointIndexes(0)), aPointList(m_RightPointIndexes(1)), aPointList(m_RightPointIndexes(2)))
        m_Front = New Plane(aPointList(m_FrontPointIndexes(0)), aPointList(m_FrontPointIndexes(1)), aPointList(m_FrontPointIndexes(2)))
        m_Back = New Plane(aPointList(m_BackPointIndexes(0)), aPointList(m_BackPointIndexes(1)), aPointList(m_BackPointIndexes(2)))
    End Sub

    Public Sub CutFaces(ByRef aBox As Solid3d, ByVal aTransaction As Transaction, ByVal aModelSpace As BlockTableRecord, _
                        ByVal aExtents As Extents)
        Dim PlaneCoords As CoordinateSystem3d
        Dim cutterBase As Circle
        Dim cutter As Solid3d
        Dim CurveList As DBObjectCollection
        Dim NewObjects As DBObjectCollection
        Dim uRegion As Region
        Dim FaceList As New List(Of Plane)

        FaceList.Add(m_Bottom)
        FaceList.Add(m_Top)
        FaceList.Add(m_Left)
        FaceList.Add(m_Right)
        FaceList.Add(m_Front)
        FaceList.Add(m_Back)

        For Each uPlane As Plane In FaceList
            PlaneCoords = uPlane.GetCoordinateSystem()
            cutterBase = New Circle(PlaneCoords.Origin, uPlane.Normal, aExtents.LongestLength * 2.0)

            CurveList = New DBObjectCollection
            CurveList.Add(cutterBase)
            NewObjects = Region.CreateFromCurves(CurveList)
            uRegion = DirectCast(NewObjects(0), Region)

            cutter = New Solid3d
            If uPlane.Normal.Equals(uRegion.Normal) Then
                cutter.Extrude(uRegion, aExtents.LongestLength * 0.5, 0.0)
            Else
                'if the normals of the new region and the original plane don't match
                'we have to use a negative value for the extrusion distance. This was added
                'to deal with an AutoCAD bug.
                cutter.Extrude(uRegion, -aExtents.LongestLength * 0.5, 0.0)
            End If

            aModelSpace.AppendEntity(cutter)
            aTransaction.AddNewlyCreatedDBObject(cutter, True)
            aBox.BooleanOperation(BooleanOperationType.BoolSubtract, cutter)
        Next

    End Sub
End Class
