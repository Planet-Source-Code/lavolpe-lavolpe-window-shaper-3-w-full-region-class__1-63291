VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' A general purpose Regions class with 99% of all known region functions I can
' think of and some that are personally created that are not provided by API.
' Most are easy, some are advanced. Read thru the remarks in this section for
' an overview of the entire class and the functions that this class offers.

' Ref for Regions: http://msdn.microsoft.com/library/en-us/gdi/regions_6o6r.asp

' BRIEF HISTORY

' \/\/\/\/\/\/\
' *** VERSION 1 - used GetPixel & CombineRgn, just like everyone else
' *** VERSION 2 - use of DIBs  vs GetPixel & my unique method (Jun 04)
' UPDATED to include using DIB sections. The increase in speed is
' truly amazing over earlier version.....
' Over a 300% increase in speed noted on smallish bitmaps (96x96) &
' Over a 425% increase noted on mid-sized bitmaps (265x265)

' This approach should be somewhat faster than typical approaches that
' create a region by using CreateRectRgn, CombineRgn & DeleteObject.
' That is 'cause this approach does not use any of those functions to
' create regions (except in rare cases on Win98 and when stretching a
' region. Even in those rare cases, those functions are only called once
' for every 2000 region parts). Like typical approaches a rectangle format
' is used to "add to a region", however, this approach directly creates
' the region header & region structure and passes that to a single API
' to finish the job of creating the end result. For those that play around
' with such things, I think you will recognize the difference.

' EDITED: The ExtCreateRegion seems to have an undocumented restriction: it
' won't create regions comprising of more than 4K rectangles (Win98). So to
' get around this for very complex bitmaps, I rewrote the function to create
' regions of 2K rectangles at a time if needed. This is still extremely
' fast. I compared the window shaping code from vbAccelerator with Win98's
' SandStone.bmp (15,000 rects) & using Win98. vbAccelerator's code averaged
' 6,900 ms. My routines averaged 77 ms, uncompiled, & that's not a typo!

' EDITED: I was allowing default error trapping on UBound() to resize the
' rectangle array: when trying to update array element beyond UBound, error
' would occur & be redirected to resize the array. However, thanx to
' Robert Rayment, if the UBound checks are disabled in compile optimizations,
' then we get a crash. Therefore, checks made appropriately & a tiny loss of
' speed is the trade-off for safety.

' *** VERSION 3.1 - Anti Regions & speed modifications (10 Jan 05)
' Function has an optional parameter to return the anti-Region.
' That is the region of only transparent pixels. This could be used, for
' example, with APIs like FillRgn to replace the "transparent" color with
' another color.

' **** VERSION 4 (Oct 05) - Accessed bitmap bytes directly within memory.
' This version was tossed. Although the time saved on very large bitmaps was
' somewhat significant, the routines were not robust enough to handle all the
' various exceptions within bitmap structures (4bpp, 8pbb, compression, etc)
' Something on the backburner to possibly attack at a later time.

' -- Converted from Module to Class and added other methods
'    Note: This class does not store any regions or pointers. Regions that
'          you create from this class must be destroyed by you. There is only
'          one exception: Should you apply/assign a region to a window, then
'          you must not destroy that region; Windows O/S owns it thereafter

'   About some functions. Not all region-related functions are included here.
'   Feel free to add others as you find you want them. Some of the functions
'   are one-liners, meaning all they do is call an API and no other processing.
'   Those functions may not be of common-knowledge or not often used. Adding
'   them to this class was just a matter of courtesy vs necessity.

' *** VERSION 4.1 (Nov 05) - Included exclusion rectangles
' By excluding a portion of your bitmap that will never have "transparent" pixels
' in it, you can help speed up the process even more. This is simply because the
' routines will just blindly add that rectangular area to the shaped region without
' the need to compare any pixel in that area. The time savings is directly
' proportional to the area of the rectangle & size of the bitmap. The larger
' the bitmap and the larger the exclusion rectangle; the larger the time savings.
' Example: a 200x200 pixel rectangular exception area: 120,000 fewer bytes processed
'            200 x 200 x 3 bytes per pixel
' -- modified some core routines to accept an XForm structure (for stretching regions)

' Feb 1 & 2. Added following functions to help make this a complete Regions class:
'   CombineRegions, CreateRegion, IsNullRegion, PaintRegion
' -- also added capability for user to pass RECT & POINTAPI structures to applicable
'       functions when the undocumented VB VarPtr() function is used
' -- Added capability of allowing class to store a region.
'       This way you don't need to maintain a separate variable for that purpose.
'       When the class is destroyed, if the region still exists, it will be destroyed too.
'       See SetCachedRegion function & GetCachedRegion property
' -- for the fun of it. Added a good, flexible, region intersection routine


'========================================================================================
'                 PUBLIC FUNCTION LISTING FOR THIS CLASS
'                 see each function for more information
'========================================================================================
' AreReqionsEqual :: compares 2 regions and returns if they are equal or not **
' CombineRegions :: merges 2 regions with various options **
' ConvertRgnToRectangle :: changes any region into a simple rectangular region **
' CreateRegion :: mimics CreateRectRgn,CreateEllipticRgn,CreatePolygonRgn, & more ++
' DestroyRegion :: removes a region from memory **
' ExtractRegionRectangles :: returns region's internal rectangles as a byte array
' GetCachedRegion (Property) :: returns region handle that was set via SetCachedRegion
' GetRegionBytes :: returns the entire region (header too) as a byte array
' ImportRegion :: offers 8 ways of creating a region. Also see RegionFromBitmap ++
' IsNullRegion :: returns if a region contains a valid bounding rectangle
' IsRegion :: returns whether or not a passed handle is a region
' MoveRegion :: relocates/offsets a region or clipping region
' PaintRegion :: fills, frames and/or inverts a DC based off of a region
' PointInRegion :: returns if a passed X,Y is within a region or clipping region **
' PolyFillMode :: sets a region's polygon fill mode. See MSDN for more info.
' RectInRegion :: returns if a rectangle is within a region or clipping region **
' RegionBounds :: returns the bounding rectangle of a region **
' RegionFromBitmap :: returns a region based off of a bitmap's image ++
' RegionIntersect :: tests 2 regions for intersection/collision ++
' RegionSizeBytes :: returns the size, in bytes, of a region **
' SaveRegionToFile :: saves the region data (header too) to a file
' SetCachedRegion :: allows user to have class save & manage one region handle
' SetClipRegionToDC :: applies or removes a clipping region to/from a DC ** &&
' SetRegionToWindow :: applies or removes a region to/from a window **
' StretchRegion :: stretches a region much like StretchBlt does to bitmaps ++
' ValidateInvalidateRgn :: validates or invalidates a window based of an optional region **

' ** Those functions ending with ** are basically one liners
'    that can be called directly from APIs, should you choose
' ++ May return a region that must be destroyed at some point, unless
'    applied to a window via SetWindowRgn API or class's SetRegionToWindow
' && Regions applied as clipping regions are not actually selected into DC,
'    a copy is; therefore, originals must be destroyed at some point
'========================================================================================

' FYI: What is a region? Regions are, memory-wise, stacked rows of rectangles where
' each rectangle can be of any width, but any on a single row are all the same height,
' no row of rectangles will overlap another row of rectangles and the rows of
' rectangles are top to bottom, left to right, but each row can be of any height.
' Regional rectangles are like the bricks of a house and the house is the region.
' The method employed here builds those rows of bricks and then passes all the rows
' (the entire blueprint) to a single API to create a region/house in a single call.
' This is much faster than others out there that basically build a single brick and
' then add it to the house, a brick at a time (slow) or a row at a time (better).

' Last note: Should you want to strip out only the RegionFromBitmap routine, include:
'   RegionFromBitmap
'   c_CreatePartialRegion
'   c_CreateWin9xRegion
'   c_ScanSection
'   appropriate APIs, Constants, & User Defined Types (UDTs) for those routines

' GDI32 APIs
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateEllipticRgnIndirect Lib "gdi32.dll" (ByVal lpRect As Long) As Long
'^^ modified lpRect parameter to accept a RECT pointer ByVal vs a RECT structure ByRef
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByVal lpPoints As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'^^ modified lpPoints parameter to accept a POINTAPI pointer ByVal vs a POINTAPI structure ByRef
Private Declare Function CreatePolyPolygonRgn Lib "gdi32.dll" (ByVal lpPoints As Long, ByRef lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
'^^ modified lpPoints parameter to accept a POINTAPI pointer ByVal vs a POINTAPI structure ByRef
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (ByVal lpRect As Long) As Long
'^^ modified lpRect parameter to accept a RECT pointer ByVal vs a RECT structure ByRef
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function EqualRgn Lib "gdi32.dll" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Private Declare Function ExtCreateRegion Lib "gdi32.dll" (ByRef lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long
' ^^ modified lpXform parameter to accept an XForm or not
Private Declare Function ExtSelectClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetPolyFillMode Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetRandomRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByValt As Long) As Long
Private Declare Function GetRegionData Lib "gdi32.dll" (ByVal hRgn As Long, ByVal dwCount As Long, ByRef lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32.dll" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetUpdateRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal fErase As Long) As Long
Private Declare Function InvalidateRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Private Declare Function InvertRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PaintRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function PtInRegion Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function PtVisible Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function RectInRgn Lib "gdi32.dll" Alias "RectInRegion" (ByVal hRgn As Long, ByRef lpRect As RECT) As Long
' ^^ modified API name so it won't conflict with local function
Private Declare Function RectVisible Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function SetPolyFillMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetRectRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ValidateRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long

' Kernel32 APIs
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function IsBadReadPtr Lib "kernel32.dll" (ByVal lp As Long, ByVal ucb As Long) As Long
'^^ handy little function that can help prevent crashes when playing with memory pointers

' User32 APIs
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long 'also used in clsBarColors
Private Declare Function GetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long) As Long
Private Declare Function IntersectRect Lib "user32.dll" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Function IsRectEmpty Lib "user32.dll" (ByRef lpRect As RECT) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

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
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As Long
End Type
Private Type XFORM  ' used for stretching/skewing a region
    eM11 As Single  ' note: some versions of this UDT have
    eM12 As Single  ' the elements as double -- wrong!
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

' Custom, class enumerations:

' Options for use with SaveRegionToFile function
Public Enum FileAttributes
    attr_Normal = vbNormal
    attr_ReadOnly = vbReadOnly
    attr_Hidden = vbHidden
    attr_System = vbSystem
End Enum
' Options for use with ImportRegion function
Public Enum RegionImportLoc
    rgn_FromWindow = 0
    rgn_FromFile = 1
    rgn_FromResource = 2
    rgn_FromByteArray = 3
    rgn_CopyRegion = 4
    rgn_dcClipRegion = 5
    rgn_SysClipRegion = 6
    rgn_dcUpdateRegion = 7
End Enum
' options for use with CombineRegions function
Public Enum RegionOps
    RGN_NoAction = 0
    RGN_AND = 1
    RGN_COPY = 5
    RGN_DIFF = 4
    RGN_OR = 2
    RGN_XOR = 3
End Enum
' options for use with CreateRegion function
Public Enum RegionCreations
    rgn_Rectangle = 0
    rgn_RoundRectangle = 1
    rgn_Elliptical = 2
    rgn_Polygon = 3
    rgn_PolyPolygon = 4
End Enum
' options for use with PaintRegion function
Public Enum RegionFills
    rgn_Fill = 1
    rgn_Frame = 2
    rgn_Invert = 4
End Enum
' options for use with PolyFillMode function
Public Enum GetSetAction
    mode_Get = 0
    mode_Set = 1
End Enum

' Constants used
Private Const BI_RGB As Long = 0&
Private Const OBJ_REGION As Long = 8&
Private Const OBJ_BRUSH As Long = 2&
Private Const SYSRGN As Long = 4&

' common to a few sub-functions, these are erased immediately after use/reuse:
Private rgnRects() As RECT  ' array of rectangles comprising region
Private gpBytes() As Byte   ' the DIB byte array
Private c_hRgn As Long      ' optional region maintained by class: seet CachedRegion function


Public Function AreReqionsEqual(ByVal hRgn1 As Long, ByVal hRgn2 As Long) As Long

    ' compares 2 regions & returns non-zero if they are the same
    ' Regions are the same if they are the same size & shape
    
    AreReqionsEqual = EqualRgn(hRgn1, hRgn2)
    
End Function

Public Function CombineRegions(ByVal hRgnDest As Long, ByVal hRgnSrc1 As Long, _
                        ByVal hRgnSrc2 As Long, ByVal Action As RegionOps, _
                        Optional ByVal bDestroyRgn As Boolean = False) As Long
                        
    ' function is basically the CombineRgn API with option to destroy
    ' hRgnSrc1 and/or hRgnSrc2 if they are not the same as hRgnDest
    
    ' See MSDN if not familiar with the CombineRgn API
    
    CombineRgn hRgnDest, hRgnSrc1, hRgnSrc2, Action
    
    If bDestroyRgn Then
        If hRgnSrc1 <> hRgnDest Then DeleteObject hRgnSrc1
        If hRgnSrc2 <> hRgnDest Then DeleteObject hRgnSrc2
    End If

End Function

Public Function ConvertRgnToRectangle(ByVal hRgn As Long, _
                Optional ByVal RectVarPtr As Long, _
                Optional ByVal Left As Long, Optional ByVal Top As Long, _
                Optional ByVal Right As Long, Optional ByVal Bottom As Long) As Long
    
    ' function changes any existing region to a regular rectangular region
    ' non-zero indicates success.
    
    ' About the RectVarPtr parameter (optional).
    ' You can pass a RECT structure if you choose vs the individual member values
    ' -- To do this, pass the structure like so:  VarPtr(myRect)
    
    'The region does not include the lower and right boundaries of the rectangle
    
    If IsRegion(hRgn) Then
        Dim rRect As RECT
        If RectVarPtr = 0& Then
            ConvertRgnToRectangle = SetRectRgn(hRgn, Top, Left, Right, Bottom)
        Else
            If IsBadReadPtr(RectVarPtr, 16&) = 0& Then
                CopyMemory rRect, ByVal RectVarPtr, 16&
                ConvertRgnToRectangle = SetRectRgn(hRgn, rRect.Top, rRect.Left, rRect.Right, rRect.Bottom)
            End If
        End If
    End If
    
    ' FYI: this little-used API would be faster in creating a region, one subregion
    ' at a time. Why? Compare to the typical logic:
    
    ' Typical: 100's of times: Create temp region, Combine it, Destroy It
    ' With this API: Create temp region just once
    '          100's of times: convert temp region (SetRectRgn), Combine it
    '          When all done, then destroy the temp region, just once
    
End Function

Public Function CreateRegion(ByVal RgnType As RegionCreations, ParamArray vData() As Variant) As Long

    ' function creates five types of regions. See also RegionFromBitmap & ImportRegion
    
    ' This function allows passing of array or UDT pointers. In VB it is very easy
    '   to pass this with the undocumented VarPtr() function. Examples below.
    
    ' Caution: Playing with pointers can lead to crashes if invalid pointers are
    ' passed. This routine attempts to prevent that via use of the IsBadReadPtr API.
    
    ' The vData() parameter array is used as follows & depends on
    ' (RgnType) the type of region being created :

    ' for Rgn_Rectangle & Rgn_Elliptical. 3 options of supplying your values
        ' Option 1: RECT structure pointer (1 param entry only)
            ' 1st param entry: VarPtr(yourRECTvariable)
            ' Ex: myRgn=CreateRegion(rgn_Rectangle, VarPtr(myRect))
            ' Ex: myRgn=CreateRegion(rgn_Elliptical, VarPtr(myRect))
        ' Option 2: Long array pointer that contains at least 4 entries (1 param entry only)
            ' 1st param entry: VarPtr(yourArray(item))
            ' Ex: myRgn=CreateRegion(rgn_Rectangle, VarPtr(myArray(0)))
            ' Ex: myRgn=CreateRegion(rgn_Elliptical, VarPtr(myArray(0)))
        ' Option 3: Individual RECT members (4 param entries)
            ' 1st param entry: Left value
            ' 2nd param entry: Top value
            ' 3rd param entry: Right value
            ' 4th param entry: Bottom value
            ' Ex: myRgn=CreateRegion(rgn_Rectangle, 0, 0, 100, 200)
            ' Ex: myRgn=CreateRegion(rgn_Elliptical, 0, 0, 100, 200)
    
    ' for Rgn_RoundRectangle. 3 options of supplying your values
        ' Option 1: RECT structure pointer (3 param entries)
            ' 1st param entry: VarPtr(yourRECTvariable)
            ' 2nd param entry: rounded corner width
            ' 3rd param entry: rounded corner height
            ' Ex: myRgn=CreateRegion(rgn_RoundRectangle, VarPtr(myRect), 20, 20)
        ' Option 2: Long array pointer that contains at least 4 entries (3 param entries)
            ' 1st param entry: VarPtr(yourArray(item))
            ' 2nd param entry: rounded corner width
            ' 3rd param entry: rounded corner height
            ' Ex: myRgn=CreateRegion(rgn_RoundRectangle, VarPtr(myArray(0)), 20, 20)
        ' Option 3: Individual RECT members (6 param entries)
            ' 1st param entry: Left value
            ' 2nd param entry: Top value
            ' 3rd param entry: Right value
            ' 4th param entry: Bottom value
            ' 5th param entry: rounded corner width
            ' 6th param entry: rounded corner height
            ' Ex: myRgn=CreateRegion(rgn_Rectangle, 0, 0, 100, 200, 20, 20)
    
    ' for Rgn_Polygon. 2 options of providing your values
        ' Option 1: POINTAPI array pointer (3 param entries)
            ' 1st param entry: VarPtr(yourPOINTAPIarray(item))
            ' 2nd param entry: nr of array items used in 1st param above
            ' 3rd param entry: region fill mode. Must be 1 or 2
            ' Ex: myRgn=CreateRegion(rgn_Polygon, VarPtr(myPointAPIs(0)), 6, 2)
        ' Option 2: Long array pointer containing pairs of X,Y points (3 param entries)
            ' 1st param entry: VarPtr(yourArray(item))
            ' 2nd param entry: nr of array "pairs" used in 1st param above
            ' 3rd param entry: region fill mode. Must be 1 or 2
            ' Ex: myRgn=CreateRegion(rgn_Polygon, VarPtr(myArray(0)), 6, 2)
    
    ' for Rgn_PolyPolygon. 2 options of providing your values
        ' Note: the 2nd parameter is a long array where each entry in the array
        '       is the count of "point pairs" of each polygon you are supplying
        ' Option 1: POINTAPI array pointer (4 param entries)
            ' 1st param entry: VarPtr(yourPOINTAPIarray(item))
            ' 2nd param entry: VarPtr(yourPointCountArray(item))
            ' 3rd param entry: nr of array items used in 2nd param above
            ' 4th param entry: region fill mode. Must be 1 or 2
            ' Ex: myRgn=CreateRegion(rgn_PolyPolygon, VarPtr(myPointAPIs(0)), VarPtr(myPolyEntries(0)), 2, 2)
        ' Option 2: Long array pointer containing pairs of X,Y points (4 param entries)
            ' 1st param entry: VarPtr(yourArray(item))
            ' 2nd param entry: VarPtr(yourPointCountArray(item))
            ' 3rd param entry: nr of array items used in 2nd param above
            ' 4th param entry: region fill mode. Must be 1 or 2
            ' Ex: myRgn=CreateRegion(rgn_PolyPolygon, VarPtr(myArray(0)), VarPtr(myPolyEntries(0)), 2, 2)
    
    Dim hRgn As Long
    
    On Error Resume Next
    Select Case RgnType
    ' Note: if no parameters are passed in the paramArray, the UBound will be -1
        
        Case rgn_Rectangle
            If UBound(vData) = 0& Then ' minimum is 1 param entry (RECT pointer)
                If (VarType(vData(0)) = vbLong) Then
                    If IsBadReadPtr(vData(0), 16&) = 0& Then 'ensure we can read 16bytes
                        hRgn = CreateRectRgnIndirect(vData(0))
                    End If
                End If
            ElseIf UBound(vData) > 2& Then ' requires at least 4 param entries
                hRgn = CreateRectRgn(vData(0), vData(1), vData(2), vData(3))
            End If
            
        Case rgn_RoundRectangle
            If UBound(vData) = 2& Then 'minimum is 3 param entries
                If (VarType(vData(0)) = vbLong) Then
                    If IsBadReadPtr(vData(0), 16&) = 0& Then 'ensure we can read 16bytes
                        Dim rRect As RECT
                        CopyMemory rRect, ByVal CLng(vData(0)), 16&
                        hRgn = CreateRoundRectRgn(rRect.Left, rRect.Top, rRect.Right, rRect.Bottom, vData(1), vData(2))
                    End If
                End If
            ElseIf UBound(vData) > 4& Then ' requires at least 6 param entries
                hRgn = CreateRoundRectRgn(vData(0), vData(1), vData(2), vData(3), vData(4), vData(5))
            End If
            
        Case rgn_Elliptical
            If UBound(vData) = 0& Then
                If (VarType(vData(0)) = vbLong) Then
                    If IsBadReadPtr(vData(0), 16&) = 0& Then 'ensure we can read 16bytes
                        hRgn = CreateEllipticRgnIndirect(vData(0))
                    End If
                End If
            ElseIf UBound(vData) > 2& Then ' requires at least 4 param entries
                hRgn = CreateEllipticRgn(vData(0), vData(1), vData(2), vData(3))
            End If
            
        Case rgn_Polygon
            If UBound(vData) > 1& Then 'requires at least 3 param entries
                If vData(1) > 0& Then ' number of array elements
                    ' ensure we can read the number of points; each point is 8 bytes
                    If IsBadReadPtr(vData(0), vData(1) * 8&) = 0& Then
                        hRgn = CreatePolygonRgn(vData(0), vData(1), vData(2))
                    End If
                End If
            End If
    
        Case rgn_PolyPolygon
            If UBound(vData) > 2& Then ' requires at least 4 param entries
                If vData(2) > 0& Then ' number of 2nd parameter array elements
                    ' ensure we can read the number of 2nd parameter array items
                    If IsBadReadPtr(vData(1), vData(2) * 4&) = 0& Then
                
                        Dim I As Integer, ttlPoints As Long, tArray() As Long
                        ' copy to local array the array items from the passed pointer
                        ReDim tArray(0 To vData(2) - 1&)
                        CopyMemory tArray(0), ByVal CLng(vData(1)), vData(2) * 4&
                        For I = 0 To vData(2) - 1
                            ttlPoints = ttlPoints + tArray(I) 'tally expected points
                        Next
                        ' ensure we can read the number of points; each point is 8 bytes
                        If IsBadReadPtr(vData(0), ttlPoints * 8&) = 0& Then
                            hRgn = CreatePolyPolygonRgn(vData(0), tArray(0), vData(2), vData(3))
                        End If
                    
                    End If
                End If
            End If
        
    End Select
    
    If Err Then
        Err.Clear
        If hRgn Then DeleteObject hRgn
        hRgn = 0&
    End If
    
    CreateRegion = hRgn
    
End Function

Public Function DestroyRegion(ByRef hRgn As Long) As Long
    
    ' Simply destroy an existing region; returns non-zero if successful
    ' DeleteObject destroys much more than regions; verify we are passed a region.
    If IsRegion(hRgn) Then DestroyRegion = DeleteObject(hRgn)
    
End Function

Public Function ExtractRegionRectangles(ByVal hRgn As Long, Optional ByRef nrRectangles As Long) As Byte()

' This function extracts the rectangles within the region
' Should you need to use the function, you will need to use a bit of memory shifting
' since classes don't allow returning of UDTs (i.e., RECT structures)
' Use the following & un-Rem when pasting in your routine.
' Remarks are prefixed with slashes

'' /Create 2 arrays first:
' Dim rgnRects() As RECT, rgnBytes() As Byte
'' /Create a long variable (Optional)
' Dim nrRects As Long
'' /Call function:
' rgnBytes = ExtractRegionRectangles(hRgn, nrRects) ' nrRects is optional
'' /Ensure array returned ok:
' If UBound(rgnBytes) > -1 Then
''       /Resize your Rect array...
''       /If you used the optional nrRects variable then
'           ReDim rgnRects(0 To nrRects - 1)
''       /Otherwise, do a little math...each Rect is 16 bytes
'           ReDim rgnRects(0 To (UBound(rgnBytes) + 1) \ 16 - 1)
''       /Now copy the bytes into your rectangle array
'       CopyMemory rgnRects(0), rgnBytes(0), UBound(rgnBytes) + 1
'       Erase rgnBytes()
' End If


    Dim b() As Byte
    Dim nBytes As Long
    
    ' get the region in bytes
    b() = GetRegionBytes(hRgn)
    nBytes = UBound(b) + 1&
    If nBytes > 32& Then
        nrRectangles = (nBytes - 32&) \ 16
        ' shift rectangles left 2 places to remove the region header info
        CopyMemory b(0), b(32), nBytes - 32&
        ReDim Preserve b(0 To nBytes - 33&)
    Else
        ReDim b(-1 To -1)
    End If
    
    ExtractRegionRectangles = b()

End Function

Public Property Get GetCachedRegion() As Long
    ' this class doesn't cache any thing it creates.
    ' However, you can use the class to cache one region handle via the
    ' SetCachedRegion function and this property will retrieve that region handle
    GetCachedRegion = c_hRgn
End Property

Public Function GetRegionBytes(ByVal hRgn As Long, Optional ByVal bDestroyRegion As Boolean) As Byte()

' function returns actual bytes of a region structure, including the header
' if UBound() of function's return array < 1 then function failed

    Dim nBytes As Long
    Dim b() As Byte
    
    ReDim b(-1 To -1)   ' set failure return value
    If IsRegion(hRgn) Then
        ' get nr bytes in the array
        nBytes = GetRegionData(hRgn, 0&, ByVal 0&)
        If nBytes > 32& Then
            ' size our array to retrieve the bytes
             ReDim b(0 To nBytes - 1&) As Byte
             ' call function again, filling the array
             If nBytes <> GetRegionData(hRgn, nBytes, b(0)) Then ReDim b(-1 To -1)
        End If
        If bDestroyRegion Then DeleteObject hRgn
    End If
    
    GetRegionBytes = b()
    
End Function

Public Function ImportRegion(ByVal RgnType As RegionImportLoc, _
            ByVal RgnLocation As Variant, Optional ByVal lParam As Long) As Long

' The RegionFromBitmap is one way to create a region; here are 8 others

' See also CreateRegion for other ways to create a region from scratch

' This function retrieves/creates a region from one of these specific locations
' If function returns zero, then no region was created

' 1. From another window
'    RgnType = rgn_FromWindow
'    RgnLocation = hWnd of the window containing the region
'    lParam not used

' 2. From a file that contains region data
'    RgnType = rgn_FromFile
'    RgnLocation = the existing filename containing the data
'    lParam not used

' 3. From a .RES file's resource. The passed array must include the region header
'    RgnType = rgn_FromResource
'    RgnLocation = string name of resource section, usually "Custom"
'    lParam = The numerical ID of the data in that custom resource section

' 4. From a byte array you pass. The passed array must include the region header
'    RgnType = rgn_FromByteArray
'    RgnLocation = the byte array
'    lParam not used

' 5. From another region handle
'    RgnType = rgn_CopyRegion
'    RgnLocation = the handle to the region to be copied
'    lParam = if non-zero then passed region will be destroyed

' 6. The clipping region from a DC
'    RgnType = rgn_dcClipRegion
'    RgnLocation = the hDC to get the clipping region from
'    lParam = not used

' 7. The system clipping region. See: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/clipping_4q0e.asp
'    RgnType = rgn_SysClipRegion
'    RgnLocation = the hDC to get the clipping region from
'    lParam = not used

' 8. The DCs current Update Region
'    RgnType = rgn_dcUpdateRegion
'    RgnLocation = the hDC to get the update region from
'    lParam = if non-zero, window bkg is erased & child windows are redrawn. If zero, no drawing takes place
    
    
    Dim hRgn As Long
    Dim cRgn As Long
    
    On Error Resume Next
    
    Select Case RgnType
'1.
    Case rgn_FromWindow
        Dim hWnd As Long
        hWnd = Val(RgnLocation)
        If hWnd <> 0& Then
            If IsWindow(hWnd) Then
                hRgn = CreateRectRgn(0&, 0&, 0&, 0&)
                If GetWindowRgn(hWnd, hRgn) = 0& Then
                    DeleteObject hRgn
                    hRgn = 0&
                End If
            End If
        End If
'2.
    Case rgn_FromFile
        Dim fnr As Integer
        If Len(Dir$(CStr(RgnLocation), vbHidden Or vbArchive Or vbReadOnly Or vbSystem)) > 0& Then
            If Not Err Then
                fnr = FreeFile()
                Open CStr(RgnLocation) For Binary Access Read As #fnr
                If Not Err Then
                    ' gpBytes will be used to create region near end of this routine
                    ReDim gpBytes(0 To LOF(fnr) - 1&)
                    Get #fnr, , gpBytes()
                    If Err Then RgnType = -1& ' flag indicating failure
                End If
                Close #fnr
            End If
        End If
'3.
    Case rgn_FromResource
        ' gpBytes will be used to create region near end of this routine
        gpBytes() = LoadResData(lParam, CStr(RgnLocation))
        If Err Then RgnType = -1& ' flag indicating failure
'4.
    Case rgn_FromByteArray
        If (VarType(RgnLocation) And (vbArray Or vbByte)) > 0& Then
            ' gpBytes will be used to create region near end of this routine
            gpBytes = RgnLocation
            If Err Then RgnType = -1& ' flag indicating failure
        Else
            RgnType = -1& ' flag indicating failure
        End If
'5.
    Case rgn_CopyRegion
        cRgn = Val(RgnLocation)
        If IsRegion(cRgn) Then
            hRgn = CreateRectRgn(0&, 0&, 0&, 0&)
            CombineRegions hRgn, cRgn, hRgn, RGN_COPY, (lParam <> 0&)
        End If
'6.
    Case rgn_dcClipRegion
        If IsNumeric(RgnLocation) Then
            hRgn = CreateRectRgn(0&, 0&, 0&, 0&)
            If GetClipRgn(Val(RgnLocation), hRgn) < 1& Then
                DeleteObject hRgn
                hRgn = 0&
            End If
        End If
'7.
    Case rgn_SysClipRegion
        If IsNumeric(RgnLocation) Then
            hRgn = CreateRectRgn(0&, 0&, 0&, 0&)
            If GetRandomRgn(Val(RgnLocation), hRgn, SYSRGN) = 0& Then
                DeleteObject hRgn
                hRgn = 0&
            End If
        End If
'8.
    Case rgn_dcUpdateRegion
        If IsNumeric(RgnLocation) Then
            hRgn = CreateRectRgn(0&, 0&, 0&, 0&)
            If GetUpdateRgn(Val(RgnLocation), hRgn, lParam) = 0& Then
                DeleteObject hRgn
                hRgn = 0&
            End If
        End If
        
    Case Else
        Exit Function
        
    End Select
    
    Select Case RgnType
        Case rgn_FromFile, rgn_FromResource, rgn_FromByteArray
        ' use our custom region creation logic vs the ExtCreateRegion API since Win98
        ' has issues with complex regions. This way we create the region always, and
        ' w/o errors caused by O/S restrictions
            If UBound(gpBytes) > 32& Then
                ReDim rgnRects(0 To (UBound(gpBytes) + 1&) \ 16)
                CopyMemory rgnRects(0), gpBytes(0), UBound(gpBytes) + 1&
                hRgn = c_CreateWin9xRegion(UBound(rgnRects), rgnRects(1).Left, rgnRects(1).Right)
                Erase rgnRects
            End If
            
    End Select
    
    If Err Then
        Err.Clear
        If hRgn <> 0& Then
            DeleteObject hRgn
            hRgn = 0&
        End If
    End If
    
    Erase gpBytes
    ImportRegion = hRgn

End Function

Public Function IsNullRegion(ByVal hRgn As Long) As Long
    
    ' returns whether or not the passed region is Null
    ' Null is qualified by one of these 2 conditions:
    '   1. The region has no regional rectangles but is still a valid region otherwise
    '   2. The region has no area (Bounding rectangle's Right<=Left or Bottom<=Top)
    
    ' 0   indicates not Null
    ' > 0 indicates Null
    ' < 0 indicates invalid region
    
    Dim rRect As RECT
    If IsRegion(hRgn) Then
        If GetRegionData(hRgn, 0&, ByVal 0&) = 32& Then
            ' region has no rectangles; null by default
            IsNullRegion = 32&
        Else
            GetRgnBox hRgn, rRect
            IsNullRegion = Abs(IsRectEmpty(rRect))
        End If
    Else
        IsNullRegion = -1&
    End If

End Function

Public Function IsRegion(ByVal hRgn As Long) As Boolean
    
    ' returns whether or not the passed pointer is a region
    IsRegion = (GetObjectType(hRgn) = OBJ_REGION)

End Function

Public Function MoveRegion(ByVal hRgnOrDC As Long, ByVal X As Long, ByVal Y As Long, _
                            Optional isClipRegion As Boolean = False, _
                            Optional ByVal RectVarPtr As Long) As Boolean

    ' this function moves a region
    
    ' if isClipRegion = True, then the movement is incremental and only applies
    '   to a clipping region. hRgnOrDC is the DC containing the clipping region
    
    ' if isClipRegion = False, then the movement is absolute. X & Y are the new
    '   coordinates. hRgnOrDC is the handle to the region to move.
    '   If desired, function will return the new coordinates via the RectVarPtr
    
    ' About the RectVarPtr parameter (optional).
    ' To pass a RECT structure, use:  VarPtr(myRect)
    
    ' return value indicates success
    
    If isClipRegion Then
    
        MoveRegion = (OffsetClipRgn(hRgnOrDC, X, Y) <> 0&)
    
    Else
    
        If IsRegion(hRgnOrDC) Then
        
            Dim rgnRect As RECT
            GetRgnBox hRgnOrDC, rgnRect
            OffsetRgn hRgnOrDC, -rgnRect.Left + X, -rgnRect.Top + Y
            
            If RectVarPtr <> 0& Then
                ' user wants the new coordinates passed back too
                If IsBadReadPtr(RectVarPtr, 16&) = 0& Then
                    OffsetRect rgnRect, -rgnRect.Left + X, -rgnRect.Top + Y
                    CopyMemory ByVal RectVarPtr, rgnRect, 16&
                End If
            End If
            
            MoveRegion = True
            
        End If
        
    End If
End Function

Public Function PaintRegion(ByVal hRgn As Long, ByVal hDC As Long, _
                        ByVal Action As RegionFills, _
                        Optional ByVal FrameColor As Long, _
                        Optional ByVal FillColor As Long, _
                        Optional ByVal FrameBrush As Long, _
                        Optional ByVal FillBrush As Long, _
                        Optional ByVal bDeleteBrushes As Boolean = False) As Long
    
    ' Functions applies color or texture (brushes) to a region.
    ' Non-zero return indicates success
    ' non-zero Brush parameters used over non-zero Color parameters
    
    ' hRgn :: region to use for the painting. A null region will abort this routine
    ' hDC :: the DC to update
    ' Action :: all can be combined with OR
    '   rgn_Invert will invert the region using system functions
    '   rgn_Fill will fill the region with passed brush/color
    '   rgn_Frame will frame region with passed brush/color
    ' FrameColor is color to draw a single pixel frame around the region
    '   To add a thicker frame, you need to create the brush & pass its handle as FrameBrush
    ' FillColor is color to fill the region with
    '   -- pass -1 to use the brush currently selected into the DC
    '   -- Note that -1 uses the FillColor of VB DCs. But ensure FillStyle is not Transparent
    ' FrameBrush is a handle to an existing brush to use for framing the region
    '   To use patterned brushes, create one and pass the handle in this parameter
    ' FillBrush is a handle to an existing brush to use for filling the region
    '   To use patterned brushes, create one and pass the handle in this parameter
    ' bDeleteBrushes :: if True, any passed brush handles will be deleted
    '                   any passed brush handles should not be selected into any DCs
                        
    Dim bDeleteObj As Integer
    
    If IsRegion(hRgn) = False Then Exit Function
    If hDC = 0& Then Exit Function
    
    ' validate and/or create the needed brush for filling the region
    If (Action And rgn_Fill) = rgn_Fill Then
        If GetObjectType(FillBrush) = OBJ_BRUSH Then
            bDeleteObj = Abs(bDeleteBrushes)
        Else
            If FillColor = -1& Then
                ' value of -1 means use whatever is selected into the DC
                FillBrush = FillColor
            Else
                ' convert VB system colors if needed
                If FillColor < 0& Then FillColor = GetSysColor(FillColor And &HFF&)
                FillBrush = CreateSolidBrush(FillColor)
                bDeleteObj = Abs(FillBrush <> 0&)
            End If
        End If
    End If
    
    ' validate and/or create the needed brush for framing the region
    If (Action And rgn_Frame) = rgn_Frame Then
        If GetObjectType(FrameBrush) = OBJ_BRUSH Then
            If bDeleteBrushes Then bDeleteObj = bDeleteObj Or 2
        Else
            ' convert VB system colors if needed
            If FrameColor < 0& Then FrameColor = GetSysColor(FrameColor And &HFF&)
            FrameBrush = CreateSolidBrush(FrameColor)
            If FrameBrush <> 0& Then bDeleteObj = bDeleteObj Or 2
        End If
    End If
        
    ' fill then frame as needed
    If FillBrush <> 0& Then
        If FillColor = -1& Then
            PaintRegion = PaintRgn(hDC, hRgn)
        Else
            PaintRegion = FillRgn(hDC, hRgn, FillBrush)
            If (bDeleteObj And 1) = 1 Then DeleteObject FillBrush
        End If
    End If
    
    ' Invert if that action was passed.
    If (Action And rgn_Invert) = rgn_Invert Then PaintRegion = InvertRgn(hDC, hRgn)
    
    ' now frame if needed
    If FrameBrush <> 0& Then
        PaintRegion = FrameRgn(hDC, hRgn, FrameBrush, 1&, 1&)
        If (bDeleteObj And 2) = 2 Then DeleteObject FrameBrush
    End If
    
End Function

Public Function PointInRegion(ByVal hRgnOrDC As Long, ByVal X As Long, ByVal Y As Long, _
                                Optional ByVal bIsDC As Boolean = False) As Long
    
    ' if the point is within a region or clipping region, non-zero will be returned
    
    ' bIsDC = True, then hRgnOrDC is a DC containing a clipping region
    ' bIsDC = False, then hRgnOrDC is a region handle
    
    If hRgnOrDC = 0& Then Exit Function
    
    If bIsDC Then
        PointInRegion = PtVisible(hRgnOrDC, X, Y)
    Else
        PointInRegion = PtInRegion(hRgnOrDC, X, Y)
    End If

    '^^ what's the difference? If you have the clipping region cached, you can use
    '   either function. But if you don't have it cached, then the only way to check
    '   against a clipping region is PtVisible or extracting the region from the DC
End Function

Public Function PolyFillMode(ByVal Action As GetSetAction, ByVal hDC As Long, _
                            Optional SetFillValue As Long = 2&) As Long

    ' return value is the DC polygon fill mode if action = mode_Get
    ' else the return value is the result of setting the mode: non-zero indicates success
    
    ' From MSDN: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/regions_2z8l.asp
    ' When the fill mode is ALTERNATE, GDI fills the area between odd-numbered and even-numbered
    '   polygon sides on each scan line. That is, GDI fills the area between the first and second
    '   side, between the third and fourth side, and so on.
    ' When the fill mode is WINDING, GDI fills any region that has a nonzero winding value. This
    '   value is defined as the number of times a pen used to draw the polygon would go around the
    '   region. The direction of each edge of the polygon is important.

    If hDC = 0& Then Exit Function
    
    If Action = mode_Set Then
        PolyFillMode = SetPolyFillMode(hDC, SetFillValue)
    Else
        PolyFillMode = GetPolyFillMode(hDC)
    End If
    
End Function

Public Function RectInRegion(ByVal hRgnOrDC As Long, _
                Optional ByVal bIsDC As Boolean = False, Optional ByVal RectVarPtr As Long, _
                Optional ByVal Left As Long, Optional ByVal Top As Long, _
                Optional ByVal Right As Long, Optional ByVal Bottom As Long) As Long
                
    ' If any part of the rectangle exists in the region, the return value is non-zero
    
    ' bIsDC = True, then hRgnOrDC is a DC containing a clipping region
    ' bIsDC = False, then hRgnOrDC is a region handle
    
    ' About the RectVarPtr parameter (optional).
    ' You can pass a RECT structure if you choose vs the individual member values
    ' -- To do this, pass the structure like so:  VarPtr(myRect)
    
    If hRgnOrDC = 0& Then Exit Function
    
    Dim tRect As RECT
    If RectVarPtr = 0& Then
        SetRect tRect, Left, Top, Right, Bottom
    Else
        If IsBadReadPtr(RectVarPtr, 16&) = 0& Then CopyMemory tRect, ByVal RectVarPtr, 16&
    End If
    
    If bIsDC Then
        RectInRegion = RectVisible(hRgnOrDC, tRect)
    Else
        RectInRegion = RectInRgn(hRgnOrDC, tRect)
    End If
    
    '^^ what's the difference? If you have the clipping region cached, you can use
    '   either function. But if you don't have it cached, then the only way to check
    '   against a clipping region is RectVisible or extracting the region from the DC

End Function

Public Function RegionBounds(ByVal hRgn As Long, Optional ByVal RectVarPtr As Long, _
                    Optional ByRef Left As Long, Optional ByRef Top As Long, _
                    Optional ByRef Right As Long, Optional ByRef Bottom As Long) As Long

    ' this function returns the bounding rectangle of the passed region
    ' by updating the passed parameters or the passed RECT structure
    
    ' About the RectVarPtr parameter (optional).
    ' You can pass a RECT structure if you choose vs the individual member values
    ' -- To do this, pass the structure like so:  VarPtr(myRect)
    
    ' Note: this is also bytes 16-31 of the header returned from GetRegionData
    
    Dim rgnRect As RECT
    If IsRegion(hRgn) Then
        RegionBounds = GetRgnBox(hRgn, rgnRect)
        If RectVarPtr <> 0& Then
            If IsBadReadPtr(RectVarPtr, 16&) = 0& Then
                CopyMemory ByVal RectVarPtr, rgnRect, 16&
            End If
        Else
            Left = rgnRect.Left
            Top = rgnRect.Top
            Right = rgnRect.Right
            Bottom = rgnRect.Bottom
        End If
    End If
End Function

Public Function RegionFromBitmap(ByVal hBitmap As Long, _
        Optional ByVal hWndToApply As Long, _
        Optional ByVal transColor As Long = -1&, _
        Optional ByVal returnAntiRegion As Boolean, _
        Optional ByVal ExcludeL As Long, Optional ByVal ExcludeT As Long, _
        Optional ByVal ExcludeR As Long, Optional ByVal ExcludeB As Long) As Long

'*******************************************************
' FUNCTION RETURNS.
' 1. If hWndToApply is zero, the shaped region handle is returned
' 2. Otherwise, a non-zero value indicates region created & applied
'    to the passed window handle
'*******************************************************

' Also see CreateRegion & ImportRegion for other ways to create/re-use regions

' PARAMETERS
'=============
' hBitmap (Required) : handle to a bitmap to be used to create the region
' Optional hWndToApply : hWnd to assign the shaped region to
' Optional transColor : the transparent color, if -1, then top/left corner is used
' Optional returnAntiRegion : If False (default) then the region excluding transparent
'       pixels will be used/returned.  If True, then the region including only
'       transparent pixels will be used/returned
' Optional ExcludeL, ExcludeT, ExcludeR, ExcludeB are exclusion rectangle coordinates
'   If provided, the pixels in those coordinates will not be examined and
'   that rectangle's area will be added to the shaped region as is. This exclusion
'   rectangle should be as large of an area of the image as possible
'   which contains no transparency

    ' test for required variable first
    If hBitmap = 0& Then Exit Function
    ' if applying to a window, ensure the value passed is a window
    If hWndToApply <> 0& Then
        If IsWindow(hWndToApply) = 0& Then Exit Function
    End If
    
    ' now ensure hBitmap handle passed is a usable bitmap
    Dim bmpInfo As BITMAPINFO
    If GetGDIObject(hBitmap, Len(bmpInfo), bmpInfo) = 0& Then Exit Function
    
    ' declare bunch of variables...
    Dim rectCount As Long ' number of rectangles & used to increment above array
    Dim lScanLines As Long ' used to size the DIB bit array
    Dim rtnRegion As Long ' region handle returned by this function if appropriate
    Dim dibDC As Long ' DC to use for GetDIBits
   
    On Error GoTo CleanUp
      
    With bmpInfo.bmiHeader
    
        .biHeight = Abs(.biHeight) ' per msdn the .biHeight may be negative already
        
        ' Scans must align on dword boundaries:
        lScanLines = (.biWidth * 3& + 3&) And &HFFFFFFFC
        ReDim gpBytes(0 To lScanLines - 1&, 0 To .biHeight - 1&)
        
        ' validate passed exclusion rectangle dimensions
        If ExcludeL < 0& Then ExcludeL = 0&
        If ExcludeT < 0& Then ExcludeT = 0&
        If ExcludeB > .biHeight Then ExcludeB = .biHeight
        If ExcludeR > .biWidth Then ExcludeT = .biWidth
        If ExcludeB <= ExcludeT Or ExcludeR <= ExcludeL Then
            ExcludeR = ExcludeL
            ExcludeB = ExcludeT
        End If
       
        ' build the DIB header
       .biSize = Len(bmpInfo.bmiHeader)
       .biBitCount = 24
       .biPlanes = 1
       .biCompression = BI_RGB
       .biClrUsed = 0&
       .biClrImportant = 0&
       .biSizeImage = 0&
       .biHeight = -.biHeight
        ' ^^ most DIBs are bottom:top, by using negative Height it will load top:bottom
    End With
    
    ' get the image into DIB bits,
    dibDC = GetDC(0&)
    ' note that biHeight above was changed to negative so we reverse it form here
    If GetDIBits(dibDC, hBitmap, 0&, -bmpInfo.bmiHeader.biHeight, gpBytes(0, 0), bmpInfo, 0&) = 0& Then
        Erase gpBytes()
        ReleaseDC 0&, dibDC
        Exit Function
    End If
    ReleaseDC 0&, dibDC 'failure to release DC can leak or removes limited shared DCs
        
    ' now calculate the transparent color as needed
    If transColor = -1 Then
        ' when -1 is passed, use top left corner pixel color
        transColor = gpBytes(0, 0) Or (gpBytes(1, 0) * 256&) Or (gpBytes(2, 0) * 65536)
    Else
        ' convert vbSystemColor if possible
        If transColor < 0& Then transColor = GetSysColor(transColor And &HFF&)
        ' typical DIBs are stored as BGR vs RGB
        ' convert to BGR vs converting each bitmap pixel to RGB for comparison
        transColor = ((transColor And &HFF&) * &H10000) Or _
            (((transColor And &HFF00&) \ &H100&) * &H100&) Or _
            ((transColor And &HFF0000) \ &H10000)
    End If
        
    ' Process the bitmap bytes
    With bmpInfo.bmiHeader
     
        .biHeight = Abs(.biHeight)
         ' start with an arbritray number of rectangles
        ReDim rgnRects(0 To .biWidth * 3)
            
        If ExcludeR > ExcludeL Then
            ' when an exclusion rectangle is provided, we process the bitmap like so:
            
            ' Do from the top of bitmap to top of exclusion rectangle
            c_ScanSection 0&, .biWidth, .biWidth, .biWidth, ExcludeT, transColor, returnAntiRegion, rectCount
            
            ' Do the exclusion rectangle, add estimated number of unexamined rows
            ReDim Preserve rgnRects(0 To UBound(rgnRects) + (ExcludeB - ExcludeT))
            c_ScanSection ExcludeT, ExcludeL, ExcludeR, .biWidth, ExcludeB, transColor, returnAntiRegion, rectCount
            
            ' Do from bottom of the exclusion rectangle to the bottom of the bitmap
            c_ScanSection ExcludeB, .biWidth, .biWidth, .biWidth, .biHeight, transColor, returnAntiRegion, rectCount
    
        Else
            ' No exclusion rectangle, process from top to bottom of the bitmap
            c_ScanSection 0&, .biWidth, 0&, .biWidth, .biHeight, transColor, returnAntiRegion, rectCount
        
        End If
        
    End With
    
    On Error Resume Next
    ' check for failure & engage backup plan if needed
    If rectCount Then
    
        ' there were regional rectangles identified, try to create the region
        rtnRegion = c_CreatePartialRegion(2&, rectCount + 1&, 0&, bmpInfo.bmiHeader.biWidth)

        ' ok, now to test whether or not we are good to go...
        ' if less than 2000 rectangles, function should have worked & if it didn't
        ' it wasn't due O/S restrictions -- failure

        If rtnRegion = 0& And rectCount > 2000& Then
            rtnRegion = c_CreateWin9xRegion(rectCount + 1&, 0&, bmpInfo.bmiHeader.biWidth)
            ' ^^ if rtnRegion is zero; windows could not create the region
        End If
    
    End If
    
CleanUp:
    
    Erase gpBytes  ' no longer needed; we can purge it now
    Erase rgnRects() ' no longer needed; we can purge it now
    
    If Err Then
        If rtnRegion <> 0& Then DeleteObject rtnRegion
'unRem next Debug line if the region is not being created.
'Debug.Print "Shaped Region failed. Windows could not create the region." & vbCrLf & "  Error: "; Err.Description
        Err.Clear
        ' return value for this function is now zero
    Else
        If hWndToApply <> 0& Then
            RegionFromBitmap = SetWindowRgn(hWndToApply, rtnRegion, True)
            ' ^^ if above API call fails, return value will be zero
        Else
            RegionFromBitmap = rtnRegion
            ' ^^ return handle to the shaped region
        End If
    End If

End Function

Public Function RegionIntersect(ByVal Handle1 As Long, ByVal Handle2 As Long, _
                    Optional ByVal RectVarPtr1 As Long, Optional ByVal RectVarPtr2 As Long, _
                    Optional ByVal xOffset1 As Long, Optional ByVal yOffset1 As Long, _
                    Optional ByVal xOffset2 As Long, Optional ByVal yOffset2 As Long, _
                    Optional ByVal bReturnAsRegion As Boolean, _
                    Optional ByRef RtnRgnHandle As Long) As Boolean
                    
    ' this function offers one way to determine if 2 regions intersect and will
    ' return True or False and optionally the region containing the intersection
    ' if bReturnAsRegion is True. Then that region will be passed as RtnRgnHandle.
    ' You are responsible for destroying that region if it is passed back.
    
    ' Handle1 and Handle2 are region handles. However, if you are trying to detect
    ' collision/intersection of shaped windows, those parameters can be window handles
    ' and this routine will extract the regions from those Windows & destroy
    ' the region copies when done.
    
    ' If you cached either of the region's bounding rectangles, you can provide them
    ' via VarPtr(myRect)in the RectVarPtr1 & RectVarPtr2 parameters which would negate
    ' providing any offsets. If you do cache them, pass them, it can speed up this
    ' routine too.
    
    ' About the xOffsets and yOffsets. When a region is applied to a window, the
    ' top/left coordinates of that region are almost always 0,0. Therefore, to
    ' detect shaped windows interesecting, you should supply offsets that are
    ' common to both regions. Otherwise even though the windows may be on
    ' opposites of the screen, their region coords will each be near 0,0 and an
    ' intersection is True which is really False. Generally, the offsets
    ' would be the window position in relation to the desktop.
    
    ' ^^ To get non-child window screen locations, use GetWindowRect as one solution
    '       or Me.Left\Screen.TwipsPerPixelX for example
    ' ^^ To get child window screen locations, use ClientToScreen API
    
    ' When both regions are memory only (not applied to a window) the offsets
    ' should not be needed.
    
    Dim hRect1 As RECT, hRect2 As RECT
    Dim rtnRgn As Long, iRgn As Long, bKillRgn As Integer
    
    Dim bRtn As Boolean
    
    RtnRgnHandle = 0&
    
    ' validate passed parameters
    If IsWindow(Handle1) Then
        Handle1 = ImportRegion(rgn_FromWindow, Handle1)
        If Handle1 = 0& Then Exit Function
        bKillRgn = 1 ' flag to destroy: we created it, we need to destroy it
    Else
        If IsRegion(Handle1) = False Then Exit Function ' invalid region handle
    End If
    If IsWindow(Handle2) Then
        Handle2 = ImportRegion(rgn_FromWindow, Handle2)
        ' flag to destroy: we created it, we need to destroy it
        If Handle2 <> 0& Then bKillRgn = bKillRgn Or 2
    Else
        If IsRegion(Handle2) = False Then Handle2 = 0& ' invalid region handle
    End If
    
    If Handle1 <> 0& And Handle2 <> 0& Then

        ' a quicker test first. If the region's bounding rects don't intersect
        ' then the regions don't intersect. Overkill for small regions, but can
        ' save time when testing large regions for intersection
        If RectVarPtr1 = 0& Then
            GetRgnBox Handle1, hRect1
            OffsetRect hRect1, xOffset1, yOffset1
        Else
            If IsBadReadPtr(RectVarPtr1, 16&) = 0& Then CopyMemory hRect1, ByVal RectVarPtr1, 16&
        End If
        If RectVarPtr2 = 0& Then
            GetRgnBox Handle2, hRect2
            OffsetRect hRect2, xOffset2, yOffset2
        Else
            If IsBadReadPtr(RectVarPtr2, 16&) = 0& Then CopyMemory hRect2, ByVal RectVarPtr2, 16&
        End If
        
        If IntersectRect(hRect1, hRect1, hRect2) <> 0& Then
        
            ' the above test only tells us if the region's bounding rectangles intersect.
            ' If they do, because of potential non-rectangular shapes in either of the
            ' regions, the pixels within the regions may not be interesecting...
            ' Test for that now.
        
            ' we need a region for the test, but we may not need to create it if we
            ' already got one from a Window handle. Reuse it; saves a little time
            
            If (bKillRgn And 1) = 1 Then    ' we created one from window #1
                rtnRgn = Handle1                ' use it
                bKillRgn = bKillRgn And Not 1 ' remove the destruction flag
                iRgn = Handle2                  ' region to test for intersection
            ElseIf (bKillRgn And 2) = 2 Then ' we created one from window #2
                rtnRgn = Handle2                 ' use it
                bKillRgn = bKillRgn And Not 2 ' remove destruction flag
                iRgn = Handle1                  ' region to test for intersection
            Else
                ' can't reuse any, so we create one & will copy Handle1
                rtnRgn = CreateRectRgn(0&, 0&, 0&, 0&)
                CombineRgn rtnRgn, Handle1, rtnRgn, RGN_COPY
                iRgn = Handle2                  ' region to test for intersection
            End If
        
            ' overlay the regions and see what we get back
            CombineRgn rtnRgn, rtnRgn, iRgn, RGN_AND
            
            ' if the total region size is 32bytes, this can be interpreted
            ' that 1) we have a region & 2) the region has no region rectangles
            ' since the 1st 32bytes of any region is the header information
            ' and the remaining bytes are the region rectangles.
            ' No rectangles, no collision
            If GetRegionData(rtnRgn, 0&, ByVal 0&) = 32& Then
            
                ' no true intersection, kill the region & set return values
                DeleteObject rtnRgn
                rtnRgn = 0&
                
            Else    ' we have a collision/intersection
            
                bRtn = True                 ' set return values
                If Not bReturnAsRegion Then ' destroy region if user didn't want one
                    DeleteObject rtnRgn
                    rtnRgn = 0&
                Else
                    RtnRgnHandle = rtnRgn
                End If
            
            End If
        End If
    End If
    
    ' destroy any regions this routine may have created
    If (bKillRgn And 1) = 1 Then DeleteObject Handle1
    If (bKillRgn And 2) = 2 Then DeleteObject Handle2
    
    RegionIntersect = bRtn
    
    ' a few last notes regarding the above routine. It should be quick enough
    ' for games but I would make some big modifications:
    '   1) Remove all the safety checks. You'll know whether the regions are
    '           from window handles or not. You won't have to test for valid regions
    '   2) If regions are always small, say 64x64 pixel images or smaller, I wouldn't
    '           use the IntersectRect test. Skip it or test with & without, for speed
    '   3) If #1 & #2 mods made, your code would shrink down to something like:
    '       Do the IntersectRect test if you are going to use it
    '       Copy Region1 to a new tempRgn otherwise you will permanently change Region1
    '       CombineRgn tempRgn, tempRgn, Region2, Rgn_And
    '       Test for non-intersection with GetRegionData(tempRgn,0,ByVal 0&)=32&
    '       Destroy tempRgn unless needed for some other purposes
    '   4) Of course: don't recreate the region every time your game piece moves,
    '           create region once & move the region with your game piece: OffsetRgn API
    '           By moving the region, you will be extracting the bounding RECT anyway
    '           and caching that Rect would allow you to use the IntersectRect test
    '           above without any additional overhead. Also see Me.MoveRegion function
    
End Function

Public Function RegionSizeBytes(ByVal hRgn As Long) As Long

    ' function simply returns the number of bytes in the region structure
    ' The number of individual rectangles is an easy calculation:
    
    ' NrRects = (NrBytes - 32)\16
    ' ^^ The 1st 32 bytes contains the region header, & 1st RECT starts on next byte
    
    If IsRegion(hRgn) Then RegionSizeBytes = GetRegionData(hRgn, 0&, ByVal 0&)

End Function

Public Function SaveRegionToFile(ByVal hRgn As Long, ByVal sFileName As String, _
        Optional ByVal fileAttrs As FileAttributes, Optional ByVal bDestroyRegion As Boolean) As Long

' function saves a region structure to a filename of your choice.
' Zero return value indicates success, non-zero indicates failure

' hRgn :: region handle to save
' sFileName :: full path & filename of file; path must exist
'   - if file already exists it will be overwritten
' fileAttrs :: optional file attributes for the file after it is written
' bDestroyRegion :: remove region from memory when function exits

    Dim iFile As Integer
    Dim fAttr As Long
    Dim rtnVal As Long
          
    If Not IsRegion(hRgn) Then
        ' passed handle is zero or not a region
        SaveRegionToFile = 380&   ' invalid property value
        Exit Function
    End If
    sFileName = Trim$(sFileName)
    If Len(sFileName) = 0& Then
        ' no filename, can't do anything
        SaveRegionToFile = 76&   ' path not found
        If bDestroyRegion Then DeleteObject hRgn
        Exit Function
    End If
          
         
    On Error Resume Next
    ' any errors here & we bug out
    
    If Len(Dir$(sFileName, vbSystem Or vbHidden Or vbArchive Or vbReadOnly)) > 0& Then
        ' drive & path exists?  if not can't continue
        If Err Then
            rtnVal = Err.Number
            Err.Clear
        Else
            ' try to kill existing file; if not can't continue
            SetAttr sFileName, vbNormal ' can't use Kill if file is readonly or hidden
            Kill sFileName
            If Err Then
                ' can't delete the file, therefore; can't overwrite it
                rtnVal = Err.Number
                Err.Clear
            End If
        End If
    End If
    
    ' if no errors to this point, continue on
    If rtnVal = 0& Then
        ' call helper function to get bytes in the region
        gpBytes() = GetRegionBytes(hRgn, bDestroyRegion)
        If UBound(gpBytes) > -1& Then
            
            ' prevent bogus values passed by user
            For iFile = 1 To 4
                fAttr = fAttr Or (fileAttrs And iFile)
            Next
            iFile = FreeFile
            
            Open sFileName For Binary Access Write Lock Read As #iFile
            If Not Err Then
                ' write the region data to the file; hopefully w/o errors
                Put #iFile, , gpBytes()
                Close #iFile
                If Not Err Then
                    ' try to set the file attributes. If error occurs, no biggie
                    SetAttr sFileName, fAttr
                    If Err Then Err.Clear ' don't abort on this error
                End If
            End If
            If Err Then
                rtnVal = Err.Number
                Err.Clear
            End If
        Else
            rtnVal = 17& ' can't perform requested operation
        End If
        Erase gpBytes
    Else
        If bDestroyRegion Then DestroyRegion hRgn
    End If
    
    SaveRegionToFile = rtnVal
    
End Function

Public Function SetCachedRegion(ByVal hRgn As Long, Optional ByVal bDestroyPrevious As Boolean = False) As Boolean

    ' this class doesn't cache anything. However you may use this function to cache
    ' one individual region handle so you don't have to maintain a separate variable
    ' for some "public" handle.
    
    ' Any passed handle here will be destroyed when the class is destroyed.
    ' To remove the region from the cache without destroying it,
    ' simply pass 0& as hRgn and False for bDestroyPrevious before class is terminated
    
    ' If bDestroyPrevious is True, any previously cached handle is destroyed
    
    ' Function fails if the passed handle is not a region

    If bDestroyPrevious = True Then
        If c_hRgn <> 0& Then DestroyRegion c_hRgn
    End If
    If IsRegion(hRgn) Then
        c_hRgn = hRgn
        SetCachedRegion = True
    Else
        c_hRgn = 0&
    End If
    
End Function

Public Function SetClipRegionToDC(ByVal hRgn As Long, ByVal hDC As Long, _
                            Optional ByVal bDestroyRgn As Boolean, _
                            Optional ByVal RgnAction As RegionOps = RGN_NoAction) As Long
                            
    ' this function applies a clipping region of a DC
    ' Passing zero as the hRgn, removes any custom clipping region from the DC
    ' Non-zero return value indicates success
    
    '   see MSDN:SelectClipRgn & ExtSelectClipRgn for possible return values
    
    If hDC <> 0& Then
        If RgnAction = RGN_NoAction Or (hRgn = 0& And RgnAction <> RGN_COPY) Then
            SetClipRegionToDC = SelectClipRgn(hDC, hRgn)
        Else
            ' if hRgn is Null then RgnAction must = RGN_COPY else error; therefore
            ' that test is made above.
            SetClipRegionToDC = ExtSelectClipRgn(hDC, hRgn, RgnAction)
        End If
    End If
    
    If bDestroyRgn Then DestroyRegion hRgn
    
End Function

Public Function SetRegionToWindow(ByVal hRgn As Long, ByVal hWnd As Long, _
                        Optional ByVal bRedrawWindow As Boolean = True) As Long
    ' this function applies a region to a window
    ' Passing zero as the hRgn, removes any custom region from the window
    ' non-zero return value indicates success
    If hWnd <> 0& Then
        If IsWindow(hWnd) Then SetRegionToWindow = SetWindowRgn(hWnd, hRgn, bRedrawWindow)
    End If

End Function

Public Function StretchRegion(ByVal hRgn As Long, ByVal GrowCx As Single, ByVal GrowCy As Single, _
                        Optional ByVal bDestroyOriginalRgn As Boolean = False) As Long
    
    ' function returns a stretched region if successful
    
    ' To stretch a region, supply the GrowCx and GrowCy with
    ' decimal values where 1 is the same size as current
    ' values btwn .1 and .99 reduce size and
    ' values from 1.1 and > increase size
    
    If IsRegion(hRgn) Then
    
        If GrowCx > 0# And GrowCy > 0# Then
    
            Dim nBytes As Long, rtnRgn As Long
            Dim xFrm As XFORM
            
            ' apply the XForm values (these are only XForm members compatible with all O/S)
            xFrm.eM11 = GrowCx
            xFrm.eM22 = GrowCy
            
            ' Get the number of bytes in the region
            nBytes = GetRegionData(hRgn, 0&, ByVal 0&)
            
            ' ^^ if curious why not using Me.GetRegionBytes() function to return the
            ' bytes, it is because I want to get the bytes directly into the rgnRects()
            ' array. Otherwise, I would have to transfer the bytes from a byte() array &
            ' then use copymemory to transfer that array into the rgnRects() array.
            ' Just a bit faster & less memory intensive....
            
            If nBytes > 32& Then
                ' resize the rectangle array to receive the bytes, 16bytes per rectangle
                ReDim rgnRects(0 To nBytes \ 16 - 1&)
                ' get the bytes & put directly into our RECT array
                If nBytes = GetRegionData(hRgn, nBytes, rgnRects(0)) Then
                    ' now step thru the stretching 2K at a time; very fast
                    ' ^^ thru trial and error it is actually faster to break the region
                    ' apart and do it in sections than to send a large region to be stretched
                    ' all at once. The time savings can be literally seconds...
                    rtnRgn = c_CreateWin9xRegion(UBound(rgnRects), rgnRects(1).Left, rgnRects(1).Right, VarPtr(xFrm))
                    Erase rgnRects()
                    StretchRegion = rtnRgn
                End If
            End If
    
        End If
    
        If bDestroyOriginalRgn Then DeleteObject hRgn
    
    End If
    
End Function

Public Function ValidateInvalidateRgn(ByVal hWnd As Long, Optional ByVal hRgn As Long, _
                                Optional ByVal bValidate = True, _
                                Optional ByVal bRedrawInvalidated As Boolean = True, _
                                Optional bDestroyRegion As Boolean = False) As Long
                                
    ' function either validates or invalidates a window using a region
    
    ' bValidate = True
    ' When Validating (removing a region from an Update Region), supplying a Null
    '       region will remove the entire client area from the update region.
    ' bValidate = False
    ' When InValidating (adding a region to an update region), supplying a Null
    '       region will add the entire client to the update region.
    '       Also, the bRedrawInvalidated parameter will force an immediate redraw if True
        
    If IsWindow(hWnd) = 0 Then Exit Function
    
    If bValidate Then ' removing region from update region
    
        ValidateInvalidateRgn = ValidateRgn(hWnd, hRgn)
    Else                ' adding region to update region
    
        ValidateInvalidateRgn = InvalidateRgn(hWnd, hRgn, bRedrawInvalidated)
    End If
    
    If bDestroyRegion Then DestroyRegion hRgn
                                
End Function

' /////////////////////////  CLASS ONLY FUNCTIONS \\\\\\\\\\\\\\\\\\\\
' all non-public class functions/routines are prefixed with c_ in their name

Private Function c_CreatePartialRegion(ByVal lIndex As Long, ByVal uIndex As Long, ByVal leftOffset As Long, ByVal Cx As Long, Optional ByVal xFrmPtr As Long) As Long

    ' Creates a region from a Rect() array and optionally stretches the region

    On Error Resume Next
    ' Note: Ideally, contiguous rows vertically of equal height & width should
    ' be combined into one larger row. However, thru trial & error I found
    ' that Windows does this for us and taking the extra time to do it ourselves
    ' is too cumbersome & slows down the results.
    
    ' the first 32 bytes of a region contain the header describing the region.
    ' Well, 32 bytes equates to 2 rectangles (16 bytes each), so I'll
    ' cheat a little & use rectangles to store the header
    With rgnRects(lIndex - 2&) ' bytes 0-15
        .Left = 32                      ' length of region header in bytes
        .Top = 1                        ' required cannot be anything else
        .Right = uIndex - lIndex + 1&   ' number of rectangles for the region
        .Bottom = .Right * 16&          ' byte size used by the rectangles;
    End With                            ' ^^ can be zero & Windows will calculate
    
    With rgnRects(lIndex - 1&) ' bytes 16-31 bounding rectangle identification
        .Left = leftOffset                  ' left
        .Top = rgnRects(lIndex).Top         ' top
        .Right = leftOffset + Cx            ' right
        .Bottom = rgnRects(uIndex).Bottom   ' bottom
    End With
    ' call function to create region from our byte (RECT) array
    c_CreatePartialRegion = ExtCreateRegion(ByVal xFrmPtr, (rgnRects(lIndex - 2&).Right + 2&) * 16&, rgnRects(lIndex - 2&))
    If Err Then Err.Clear

End Function

Private Function c_CreateWin9xRegion(ByVal rectCount As Long, ByVal leftOffset As Long, ByVal Cx As Long, Optional ByVal xFrmPtr As Long) As Long
' Pulled out of main routine 'cause now other routines share the same logic
' and we will simply share this part of the code

' Win98 has problems with regional rectangles over 4000
' So, we'll try again in case this is the prob with other systems too.
' We'll step it at 2000 at a time which is stil very fast

    Dim X As Long, Y As Long ' loop counters
    Dim win9xRgn As Long     ' partial region
    Dim rtnRegion As Long    ' combined region & return value of this function
    
    ' we start with 2 'cause first 2 RECTs are the header
    For X = 2& To rectCount Step 2000&
    
        If X + 2000& > rectCount Then
            Y = rectCount
        Else
            Y = X + 2000&
        End If
        
        ' attempt to create partial region
        win9xRgn = c_CreatePartialRegion(X, Y, leftOffset, Cx, xFrmPtr)
        
        If win9xRgn = 0& Then    ' failure
            ' clean up combined region if needed
            If rtnRegion Then DeleteObject rtnRegion
            rtnRegion = 0&
            Exit For ' abort
        Else
            If rtnRegion Then ' already started
                ' use combineRgn, but only every 2000th time
                CombineRgn rtnRegion, rtnRegion, win9xRgn, RGN_OR
                DeleteObject win9xRgn
            Else    ' first time thru
                rtnRegion = win9xRgn
            End If
        End If
    Next
    ' done; return result
    c_CreateWin9xRegion = rtnRegion

End Function

Private Sub c_ScanSection(ByVal Y As Long, ByVal Cx As Long, ByVal Cx2 As Long, ByVal Cx3 As Long, _
        ByVal Cy As Long, ByVal transColor As Long, ByVal returnAntiRegion As Boolean, ByRef rectCount As Long)

' function added as an extension of the main routine; therefore, some of the
' user passed variables are also passed here too

' This function will scan a specific number of bitmap lines and can skip over
' a continguous section of pixels identified by the exclusion rectangle.
' Note the complexity below. The exlusion rectangle may not be aligned with
' any of the edges of the bitmap so we may have to process pixels left, right,
' above and below the exclusion rectangle.

' See RegionFromBitmap additional remarks about the exclusion rectangle

    Dim scanX As Long, scanY As Long, X As Long ' simple loop variables
    Dim tgtColor As Long                ' a DIB pixel color (BGR format)
    Dim rStart As Long                  ' rectangle tracking started
    Dim bExludeDone As Boolean          ' exclusion rectangle processed horizontally
    Dim endX As Long                    ' last byte on line to scan
    Dim maxRectCount As Long            ' UBound(rgnRects)
    
    ' reset flag
    rStart = -1&
    endX = Cx
    maxRectCount = UBound(rgnRects)
    
    ' begin pixel by pixel comparisons
    For scanY = Y To Cy - 1&
    
        For scanX = X To endX - 1&
            ' my hack continued: we already saved a long as BGR, now
            ' get the current DIB pixel into a long (BGR also) & compare
            CopyMemory tgtColor, gpBytes(scanX * 3&, scanY), &H3
            
            ' test to see if next pixel is a target color
            If (transColor = tgtColor) Xor returnAntiRegion Then
            
                If rStart > -1& Then ' we're currently tracking a rectangle, so let's close it
                    ' see if array needs to be resized
                   If rectCount + 1& = maxRectCount Then
                        maxRectCount = UBound(rgnRects) + Cx3
                        ReDim Preserve rgnRects(0 To maxRectCount)
                    End If
                    
                    ' add the rectangle to our array
                    SetRect rgnRects(rectCount + 2&), rStart, scanY, scanX, scanY + 1&
                    rStart = -1& ' reset flag
                    rectCount = rectCount + 1&     ' keep track of nr in use
                End If
            
            Else
                ' not a target color
                If rStart = -1& Then rStart = scanX ' set start point
            
            End If
        Next scanX
        
        If rStart > -1& Then
            ' got to end of section without hitting another transparent pixel
            ' but we're tracking so we'll close rectangle now
           
            ' see if array needs to be resized
            If rectCount + 1& = maxRectCount Then
                 maxRectCount = UBound(rgnRects) + Cx3
                 ReDim Preserve rgnRects(0 To maxRectCount)
             End If
            ' add the rectangle to our array
            SetRect rgnRects(rectCount + 2&), rStart, scanY, scanX, scanY + 1&
            rStart = -1& ' reset flag
            rectCount = rectCount + 1&     ' keep track of nr in use
        End If
        
        ' exclusion rectangle checks here
        If bExludeDone Then
            ' when the left & right side of exclusion rect finished, reset variables
            bExludeDone = False
            X = 0&
            endX = Cx
        Else
            ' see if we have an exclusion rectangle being processed
            If Cx2 > 0& Then             ' exclusion rectangle passed?
                If Cx2 <> Cx3 Then      ' exclusion rectangle was passed
                    
                    ' add the excluded portion of the bitmap to our region
                    SetRect rgnRects(rectCount + 2&), Cx, scanY, Cx2 + 1&, scanY + 1&
                    rectCount = rectCount + 1&
                    ' set the next column of pixels to be checked (right side of rectangle)
                    X = Cx2
                    ' set the last column to be checked (left edge of bitmap)
                    endX = Cx3
                    ' decrement the loop variable so it will run same row again
                    scanY = scanY - 1&
                    ' toggle flag so X & endX can be restored after row is processed
                    bExludeDone = True
                End If
            End If
        End If
        
    Next scanY

End Sub

Private Sub Class_Terminate()

    ' if class was passed a valid region handle to cache,
    ' it gets destroyed when this class is destroyed

    If c_hRgn <> 0& Then DestroyRegion c_hRgn
    
End Sub
