VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Window Shaper  - Window being shaped is PictureBox"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9690
   Begin VB.CommandButton Command3 
      Caption         =   "Test this One"
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Shape Me"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   7  'Diagonal Cross
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   60
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   5565
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   45
      Width           =   5565
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   5760
      Picture         =   "Form1.frx":4DAAE
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   3840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset Me"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   3510
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   $"Form1.frx":55B10
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   135
      TabIndex        =   6
      Top             =   4395
      Visible         =   0   'False
      Width           =   5160
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":55BE5
      Height          =   1095
      Left            =   5760
      TabIndex        =   4
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Some Fun with Regions"
      Index           =   0
      Begin VB.Menu mnuRgn 
         Caption         =   "Rectangle"
         Index           =   0
      End
      Begin VB.Menu mnuRgn 
         Caption         =   "Elliptical"
         Index           =   1
      End
      Begin VB.Menu mnuRgn 
         Caption         =   "Polygon"
         Index           =   2
      End
      Begin VB.Menu mnuRgn 
         Caption         =   "Multi-Polygon"
         Index           =   3
      End
      Begin VB.Menu mnuRgn 
         Caption         =   "Rounded Rectangle"
         Index           =   4
      End
      Begin VB.Menu mnuRgn 
         Caption         =   "Shaped"
         Index           =   5
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Loading Pre-Saved Regions"
      Index           =   1
      Begin VB.Menu mnuImport 
         Caption         =   "Apply Region from File"
         Index           =   0
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Apply Region from RES File"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' note that the class has almost every region-related function there is.
' only a few are actually used in this sample project

'// UDTs used in the region menu samples; highlighting flexibility
' of class's CreateRegion function
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private amCompiled As Boolean

Private Sub Command1_Click()
Dim X As Integer
Dim cRegion As New clsRegions
Dim tStart As Long

' simple, simple tests. I will not compare the novice approach of GetPixel/SetPixel
' as that old method is super slow; kinda akin to copying and pasting a text document
' one character at a time :)

' For a real appreciation, run compiled vs IDE
' Speed difference when compiled? Routines 2 to 3 times faster at least

' Using no optional exclusion rectangles (see remarks a bit further down about these):
' A comparison to VBAccelerator? For Picture1 image (few regions) theirs only slightly
' slower (.062 vs .047 seconds). On the sandstone bmp (1000's regions, Picture2)...
' theirs +100x slower (6.7 vs .062 seconds)

' should you want to compare mine to the code found on VBAccelerator, the link:
' http://www.vbaccelerator.com/home/VB/Code/Libraries/Graphics_and_GDI/Changing_Window_Shapes/Creating_Window_Shapes_from_Bitmaps/Bitmap_Shape_Project_Files.asp
' Download their zip and extract the files: cDibSection.cls & cDIBSectionRegion.cls
' then add files to this project and un-rem the following section.
'' ////////////
'    ' using the sandstone bmp for the test to highlight speed differences
'    Dim vbAccelDIB As New cDIBSection
'    Dim vbAccelRgn As New cDIBSectionRegion
' tStart = GetTickCount
'    vbAccelDIB.CreateFromPicture Picture2.Picture
'    vbAccelRgn.Create vbAccelDIB, 8421504 ' rgb(128,128,128) top left pixel color
' Debug.Print "vbAccelerator method (complex region) : " & (GetTickCount - tStart) & " ms"
'    vbAccelRgn.Destroy
'    Set vbAccelRgn = Nothing
'    Set vbAccelDIB = Nothing
'    Call Command3_Click         ' call my routine for same bitmap
'' ////////////
' don't forget to rem out the above section again when done testing the 2 projects.


' What I claim to be the fastest VB shaper on the planet....
tStart = GetTickCount
    ' minor modifications can improve speed on large to very large bitmaps
    ' next example is passing an exclusion rectangle. All pixels in that rectangle
    ' are not processed and simply, blindly added to the shaped region
    cRegion.RegionFromBitmap Picture1.Picture, Picture1.hWnd, , , 109, 44, 237, 237
    Debug.Print "with exclusion rectangle: " & (GetTickCount - tStart) & " ms"
    ' when compiled this example is almost instantaneous: < 20 ms
    DoEvents
    
    ' How did I get the rectangle measurements to use in the example above?
    ' Brought bitmap into paint & drew a rectangle, then wrote down the coords
    ' and added them to the function call above. Or, a container with scalemode=Pixels,
    ' you can size a Shape/Label control over the image to get the coordinates.
    
tStart = GetTickCount
    cRegion.RegionFromBitmap Picture1.Picture, Picture1.hWnd
    Debug.Print "w/o exclusion rectangle: " & (GetTickCount - tStart) & " ms"
    If amCompiled Then MsgBox "Region created and applied in " & (GetTickCount - tStart) & " ms.", vbInformation + vbOKOnly

'^^ because VB doesn't load APIs until they are first used, the first
' time you run any application that calls APIs, it might be a tad bit slower.

' On faster computers, the differences between the above two tests may be nil.
' This is a limitation of using GetTickCount for timing. Again, simple tests really.

End Sub

Private Sub Command2_Click()
    ' clear any custom window region
    Dim cRegion As New clsRegions
    cRegion.SetRegionToWindow 0, Picture1.hWnd
    
End Sub

Private Sub Command3_Click()

    ' process the 18,000 region rectangle SandStone.bmp
    ' This button only available when you manually enlarge the test form

    Dim cRegion As New clsRegions
    Dim tStart As Long
    
tStart = GetTickCount
    cRegion.RegionFromBitmap Picture2.Picture, Picture2.hWnd
    Debug.Print "Complex Region (no exclusion rectangle): " & (GetTickCount - tStart) & " ms"
    If amCompiled Then MsgBox "Region created and applied in " & (GetTickCount - tStart) & " ms.", vbInformation + vbOKOnly
    
'    Dim hRgn As Long
'    hRgn = cRegion.ImportRegion(rgn_FromWindow, Picture2.hWnd)
'    Debug.Print "Region Rectangle Count is "; (cRegion.RegionSizeBytes(hRgn) - 32) \ 16
'    cRegion.DestroyRegion hRgn

End Sub

Private Sub Form_Load()
    Me.Move (Screen.Width - 5775) \ 2, (Screen.Height - 4740) \ 2, 5775, 5140
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then Err.Clear Else amCompiled = True
End Sub

Private Sub mnuImport_Click(Index As Integer)
    Dim hRgn As Long
    Dim cRegion As New clsRegions
    
    Select Case Index
        Case 0: ' apply from disk file
        
            ' This is a step by step (no shortcuts) how to save a region to file
            ' and also how to apply a region that is saved in a file
        
            ' first ensure our sample image is not already shaped
            cRegion.SetRegionToWindow 0, Picture1.hWnd
            If MsgBox("Will create and save a region to a file.", vbYesNo, "Continue?") = vbNo Then Exit Sub
            
            ' create a region from the bitmap
            hRgn = cRegion.RegionFromBitmap(Picture1.Picture.Handle)
            
            ' now save it to a file & also delete the region
            cRegion.SaveRegionToFile hRgn, App.Path & "\_testRgn.txt", , True
            MsgBox "The region was saved as '_testRgn.txt' in the App.Path" & vbCrLf & _
                "Will now apply it to the sample image.", vbOKOnly
            
            ' get the region from the file and then apply it to the window
            hRgn = cRegion.ImportRegion(rgn_FromFile, App.Path & "\_testRgn.txt")
            cRegion.SetRegionToWindow hRgn, Picture1.hWnd
            
            
            ' now for some shortcuts
            
            'saving a region to file created from a bitmap
            ' next statements create a region from bitmap, saves to file & then destroys the region
'            With cRegion
'                .SaveRegionToFile .RegionFromBitmap(Picture2.Picture.Handle), "c:\testRgn.dat", , True
'            End With

            'applying a region saved in a file to a window
            ' next statements get region from a text file & then applies to a window
'            With cRegion
'                .SetRegionToWindow .ImportRegion(rgn_FromFile, "c:\testRgn.dat"), Picture2.hWnd
'            End With
            ' note we don't destroy regions applied to windows
    
        Case 1: ' apply from a RES file
        
            ' if you saved a region to file, simply upload that file into
            ' your RES file as a "Custom" resource.
            
            ' basically the same procedure as above
            ' first ensure our sample image is not already shaped
            cRegion.SetRegionToWindow 0, Picture1.hWnd
            MsgBox "Will now import a region saved within the project's resource file.", vbOKOnly
            
            ' get the region from the resource and then apply it to the window
            hRgn = cRegion.ImportRegion(rgn_FromResource, "Custom", 101)
            cRegion.SetRegionToWindow hRgn, Picture1.hWnd
            
            ' want a short cut for this one too?
            ' next statements create a region from RES file & apply it to a window
'            With cRegion
'                .SetRegionToWindow .ImportRegion(rgn_FromResource, "Custom", 101), Picture1.hWnd
'            End With
            ' note we don't destroy regions applied to windows
        
    End Select
End Sub

Private Sub mnuRgn_Click(Index As Integer)

    ' The CreateRegion function within the class is pretty flexible. It will
    ' allow creation using all the common GDI functions from a single function:
    '   CreateRectRgn, CreateRectRgnIndirect
    '   CreateEllipticRgn, CreateEllipticRgnIndirect
    '   CreatePolygonRgn, CreatePolyPolygonRgn
    '   CreateRoundRectRgn
    ' additionally, the class was designed to use a Variant parameter array which
    ' allows you to pass pointers to UDTs, arrays, or individual values. Although,
    ' that Variant paramArray could be converted to Longs it was not. This is
    ' because the function must accept anywhere from 1 to 6 parameters depending on
    ' which type of region you are creating and how you want to pass the data.

    Dim Vertex() As POINTAPI
    Dim tArray(0 To 8) As Long  ' RECT structure in an array for some tests
    Dim polypolyCount(0 To 1) As Long ' used only for the MultiPolygon sample
    Dim tRect As RECT
    Dim hRgn As Long
    Dim cRegion As New clsRegions

    ' cache the target window's width/height for samples below
    tRect.Right = Picture1.Width \ Screen.TwipsPerPixelX
    tRect.Bottom = Picture1.Height \ Screen.TwipsPerPixelY

    With cRegion
    
        Select Case Index
            Case 0 ' Rectangles. 3 different parameter options are allowed.
                   ' UnRem whichever you want test
               tRect.Right = tRect.Right \ 2
               tRect.Bottom = tRect.Bottom \ 2 + 50
               tRect.Top = 100: tRect.Left = 50 ' shift some settings for testing
               
               ' passing UDT
               hRgn = .CreateRegion(rgn_Rectangle, VarPtr(tRect))
               
               ' passing individual RECT members
               'hRgn = .CreateRegion(rgn_Rectangle, 0, 0, tRect.Right, tRect.Bottom)
               
               ' passing from an array
               'tArray(2) = tRect.Right
               'tArray(3) = tRect.Bottom
               'hRgn = .CreateRegion(rgn_Rectangle, VarPtr(tArray(0)))

            Case 1 ' Ellipticals. Same 3 different approaches as above.
                   ' Only one is shown here, the others are identical to above
               ' passing UDT
               hRgn = .CreateRegion(rgn_Elliptical, VarPtr(tRect))
               
            Case 2 ' Single Polygons. Only one way to pass & same as the API
                   ' Optional: PointAPI array can be substituted with a Long array
                ReDim Vertex(0 To 5)
                Vertex(0).X = 50: Vertex(0).Y = 120
                Vertex(1).X = 300: Vertex(1).Y = 120
                Vertex(2).X = 100: Vertex(2).Y = 270
                Vertex(3).X = 175: Vertex(3).Y = 20
                Vertex(4).X = 250: Vertex(4).Y = 270
                Vertex(5).X = 50: Vertex(5).Y = 120
                ' ParamArray: pointer for PointAPI array, nr of array entries, fill mode
                hRgn = .CreateRegion(rgn_Polygon, VarPtr(Vertex(0)), 6, 2)
                
            Case 3 ' Multiple Polygons. Only one way to pass & same as API
                   ' Optional: PointAPI array can be substituted with a Long array
                ReDim Vertex(0 To 6)
                Vertex(0).X = 25: Vertex(0).Y = 25  ' 1st point: (25,25)
                Vertex(1).X = 100: Vertex(1).Y = 100  ' 2nd point: (50,50)
                Vertex(2).X = 25: Vertex(2).Y = 100  ' 3rd point: (25,50)
                polypolyCount(0) = 3  ' three vertices for the triangle
            
                ' Load the vertices of the diamond into the vertex array.
                Vertex(3).X = 150: Vertex(3).Y = 150  ' 1st point: (150,150)
                Vertex(4).X = 200: Vertex(4).Y = 200  ' 2nd point: (200,200)
                Vertex(5).X = 150: Vertex(5).Y = 250  ' 3rd point: (150,250)
                Vertex(6).X = 100: Vertex(6).Y = 200  ' 4th point: (100,200)
                polypolyCount(1) = 4  ' four vertices for the diamond
            
                ' ParamArray: pointer for PointAPI array, array of poly vertice counts, nr items in that array, fill mode
                hRgn = .CreateRegion(rgn_PolyPolygon, VarPtr(Vertex(0)), VarPtr(polypolyCount(0)), 2, 1)
                
            Case 4 ' Rounded Rectangles. 3 different parameter options are allowed.
               ' passing UDT
               
                ' ParamArray: pointer for RECT, corner width, corner height
               hRgn = .CreateRegion(rgn_RoundRectangle, VarPtr(tRect), 30, 30)
               
               ' passing individual RECT members
               ' ParamArray: individual RECT members, corner width, corner height
               'hRgn = .CreateRegion(rgn_RoundRectangle, 0, 0, tRect.Right, tRect.Bottom, 30, 30)
               
               ' passing from an array
               'tArray(5) = tRect.Right
               'tArray(4) = tRect.Bottom
               ' ParamArray: pointer for Long Array, corner width, corner height
               'hRgn = .CreateRegion(rgn_RoundRectangle, VarPtr(tArray(2)), 30, 30)
               
            Case 5 ' shaped
            
                hRgn = .RegionFromBitmap(Picture1.Picture.Handle)
            
        End Select
        
        .SetRegionToWindow hRgn, Picture1.hWnd, True
    
    End With
End Sub
