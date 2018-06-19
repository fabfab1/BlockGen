'Class for handeling triangles

Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices

Public Class Triangle
    Private Point1 As Point3d
    Private Point2 As Point3d
    Private Point3 As Point3d

    Public Sub New(ByVal aPt1 As Point3d, ByVal aPt2 As Point3d, ByVal aPt3 As Point3d)
        Point1 = aPt1
        Point2 = aPt2
        Point3 = aPt3
    End Sub

    Public Sub GenerateSurface(ByVal aTransaction As Transaction, ByVal aModelSpace As BlockTableRecord, ByRef aSurfaceObjs As List(Of ObjectId))
        Dim f As New Face(Point1, Point2, Point3, True, True, True, True)
        Dim surf As Autodesk.AutoCAD.DatabaseServices.Surface = Autodesk.AutoCAD.DatabaseServices.Surface.CreateFrom(f)

        aSurfaceObjs.Add(aModelSpace.AppendEntity(surf))
        aTransaction.AddNewlyCreatedDBObject(surf, True)
    End Sub
End Class
