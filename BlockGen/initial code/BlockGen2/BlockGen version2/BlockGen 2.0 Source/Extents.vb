'Class for handeling extents information

Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices

Public Class Extents
    Private m_maxx As Double
    Private m_maxy As Double
    Private m_maxz As Double
    Private m_minx As Double
    Private m_miny As Double
    Private m_minz As Double
    Private m_Extents As Extents3d
    Private m_PointList As Point3dCollection

    Public Sub New(ByVal aPointList As Point3dCollection)
        m_PointList = aPointList

        GetMaxMin()

        m_Extents = New Extents3d(New Point3d(m_minx, m_miny, m_minz), New Point3d(m_maxx, m_maxy, m_maxz))

    End Sub

    Public ReadOnly Property DeltaX() As Double
        Get
            Return m_maxx - m_minx
        End Get
    End Property

    Public ReadOnly Property DeltaY() As Double
        Get
            Return m_maxy - m_miny
        End Get
    End Property

    Public ReadOnly Property DeltaZ() As Double
        Get
            Return m_maxz - m_minz
        End Get
    End Property

    Public ReadOnly Property Center() As Point3d
        Get
            Return New Point3d(m_minx + (m_maxx - m_minx) / 2.0, _
                               m_miny + (m_maxy - m_miny) / 2.0, _
                               m_minz + (m_maxz - m_minz) / 2.0)
        End Get
    End Property

    Public ReadOnly Property LongestLength() As Double
        Get
            Return Math.Max(Math.Max(m_maxx - m_minx, m_maxy - m_miny), m_maxz - m_minz)
        End Get
    End Property

    Private Sub GetMaxMin()
        'Must use this method due to a bug in AutoCAD's Extents3d functionality that 
        'doesn't reset the minimum point of the extents correctly.
        m_maxx = Double.MinValue
        m_maxy = Double.MinValue
        m_maxz = Double.MinValue
        m_minx = Double.MaxValue
        m_miny = Double.MaxValue
        m_minz = Double.MaxValue

        For index As Integer = 0 To 7
            If m_PointList(index).X > m_maxx Then m_maxx = m_PointList(index).X
            If m_PointList(index).Y > m_maxy Then m_maxy = m_PointList(index).Y
            If m_PointList(index).Z > m_maxz Then m_maxz = m_PointList(index).Z

            If m_PointList(index).X < m_minx Then m_minx = m_PointList(index).X
            If m_PointList(index).Y < m_miny Then m_miny = m_PointList(index).Y
            If m_PointList(index).Z < m_minz Then m_minz = m_PointList(index).Z
        Next
    End Sub

    Public Sub Grow(ByVal aPrecentage As Double)
        Dim ExpansionLength As Double = LongestLength * aPrecentage

        m_maxx += ExpansionLength
        m_maxy += ExpansionLength
        m_maxz += ExpansionLength

        m_minx -= ExpansionLength
        m_miny -= ExpansionLength
        m_minz -= ExpansionLength

        m_Extents = New Extents3d(New Point3d(m_minx, m_miny, m_minz), New Point3d(m_maxx, m_maxy, m_maxz))

    End Sub

End Class
