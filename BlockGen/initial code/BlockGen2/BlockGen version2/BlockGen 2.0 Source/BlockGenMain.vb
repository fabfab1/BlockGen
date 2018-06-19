'Main class contains functions needed for connecting to AutoCAD

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices

Public Class BlockGenMain
    Implements Autodesk.AutoCAD.Runtime.IExtensionApplication 'Ties this dll to AutoCAD

    'The following functions are required by AutoCAD extensions
    Public Sub Initialize() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Initialize
        'Report to the user that this dll has been loaded
        Dim uEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        uEditor.WriteMessage(vbNewLine & My.Application.Info.Title & " " & _
        My.Application.Info.Version.ToString & " Loaded" & vbNewLine)
    End Sub

    Public Sub Terminate() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Terminate

    End Sub
    'End of required functions

    'AutoCAD command definition for "BlockGen"
    'On execution, the BlockGen command will prompt for 8 points
    'Example script (sample1.scr):
    '   BlockGen 0,0,0 5,0,0 5,0,5 0,0,5 0,5,0 5,5,0 5,5,5 0,5,5
    'Or each point can be on its own line like this (sample2.scr):
    '   BlockGen
    '   0,0,0
    '   5,0,0
    '   5,0,5
    '   0,0,5
    '   0,5,0
    '   5,5,0
    '   5,5,5
    '   0,5,5
    <CommandMethod("BlockGen")> _
    Public Shared Sub CBlockGen()
        Dim uEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim uDatabase As Database = Application.DocumentManager.MdiActiveDocument.Database
        Dim uTransactionManager As Autodesk.AutoCAD.DatabaseServices.TransactionManager = uDatabase.TransactionManager

        Dim PointList As New Point3dCollection
        Dim origon As New Point3d(0.0, 0.0, 0.0)
        Dim XDirection As New Vector3d(1.0, 0.0, 0.0)
        Dim YDirection As New Vector3d(0.0, 1.0, 0.0)
        Dim ZDirection As New Vector3d(0.0, 0.0, 1.0)

        'Get the points
        Dim OldOsMode As Object = Application.GetSystemVariable("osmode")
        Application.SetSystemVariable("osmode", 0) 'Turn off object snap modes during point input
        For index As Integer = 1 To 8
            Dim result As PromptPointResult = uEditor.GetPoint(vbLf & "Enter Point " & index.ToString & ": ")
            If Not result.Status = PromptStatus.OK Then
                Exit Sub
            End If
            PointList.Add(result.Value)
        Next
        Application.SetSystemVariable("osmode", OldOsMode)

        'Define all 6 faces (planes) used
        Dim uFaces As New Faces(PointList)

        'Find overal extents of entered points.
        Dim uExtents As New Extents(PointList)

        'Grow extents by 25% to make sure we have enough "material" to cut away from
        uExtents.Grow(0.25)

        Using uTransaction As Transaction = uTransactionManager.StartTransaction()
            Dim uBlockTable As BlockTable = DirectCast(uTransaction.GetObject(uDatabase.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ModelSpace As BlockTableRecord = DirectCast(uTransaction.GetObject(uBlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

            'Create a base solid from which to cut the faces based on the actual entered points.
            'This will accomodate any reasonable angle that each face plane creates.
            Dim box As New Solid3d
            box.CreateBox(uExtents.DeltaX, uExtents.DeltaY, uExtents.DeltaZ)

            Dim trans As Matrix3d = Matrix3d.AlignCoordinateSystem(origon, XDirection, YDirection, ZDirection, _
                                                                   uExtents.Center, XDirection, YDirection, ZDirection)

            box.TransformBy(trans)
            ModelSpace.AppendEntity(box)
            uTransaction.AddNewlyCreatedDBObject(box, True)

            'Cut the base solid down to the given points
            uFaces.CutFaces(box, uTransaction, ModelSpace, uExtents)

            uTransaction.Commit()
        End Using

    End Sub

    <CommandMethod("WallGen")> _
    Public Shared Sub WallGen()
        Dim uEditor As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim uDatabase As Database = Application.DocumentManager.MdiActiveDocument.Database
        Dim uTransactionManager As Autodesk.AutoCAD.DatabaseServices.TransactionManager = uDatabase.TransactionManager

        Dim PointList As New Point3dCollection
        Dim origon As New Point3d(0.0, 0.0, 0.0)
        Dim XDirection As New Vector3d(1.0, 0.0, 0.0)
        Dim YDirection As New Vector3d(0.0, 1.0, 0.0)
        Dim ZDirection As New Vector3d(0.0, 0.0, 1.0)

        Dim ppo As New PromptPointOptions("")
        ppo.AllowNone = True

        'Get the points
        Dim OldOsMode As Object = Application.GetSystemVariable("osmode")
        Application.SetSystemVariable("osmode", 0) 'Turn off object snap modes during point input
        For index As Integer = 1 To 100
            ppo.Message = vbLf & "Enter Point " & index.ToString & "<enter to build>: "
            Dim result As PromptPointResult = uEditor.GetPoint(ppo)
            If Not result.Status = PromptStatus.OK Then
                Exit For
            End If
            PointList.Add(result.Value)
        Next
        Application.SetSystemVariable("osmode", OldOsMode)

        'Check for even number of points and split into two groups.
        If Not PointList.Count Mod 2 = 0 OrElse PointList.Count < 4 Then
            uEditor.WriteMessage(vbLf & "Number of supplied points is not divisible by 2, Exiting")
            Exit Sub
        End If

        Dim PointList1 As New Point3dCollection
        Dim PointList2 As New Point3dCollection
        For index As Integer = 0 To (PointList.Count / 2) - 1
            PointList1.Add(PointList(index))
        Next
        For index As Integer = (PointList.Count / 2) To PointList.Count - 1
            PointList2.Add(PointList(index))
        Next

        'Initiate surface class
        Dim smanager As Surfaces = Nothing
        If PointList1(0).Z > PointList2(0).Z Then
            smanager = New Surfaces(PointList1, PointList2)
        Else
            smanager = New Surfaces(PointList2, PointList1)
        End If

        'Build Surface
        Try
            Using uTransaction As Transaction = uTransactionManager.StartTransaction()
                Dim uBlockTable As BlockTable = DirectCast(uTransaction.GetObject(uDatabase.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ModelSpace As BlockTableRecord = DirectCast(uTransaction.GetObject(uBlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)

                smanager.GenerateSurface(uTransaction, ModelSpace)

                uTransaction.Commit()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

End Class
