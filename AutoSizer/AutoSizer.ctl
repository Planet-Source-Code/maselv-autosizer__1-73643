VERSION 5.00
Begin VB.UserControl AutoSizer 
   BackColor       =   &H000000FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   375
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   375
   ToolboxBitmap   =   "AutoSizer.ctx":0000
   Begin VB.Timer tmrint 
      Interval        =   10
      Left            =   1800
      Top             =   960
   End
   Begin VB.Image ImgAutoResizeIcon 
      Height          =   375
      Left            =   0
      Picture         =   "AutoSizer.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "AutoSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************************************
'*   Author:             Masika .S. Elvas                                                            *
'*   Gender:             Male                                                                        *
'*   Postal Address:     P.O Box 137, BUNGOMA 50200, KENYA                                           *
'*   Phone No:           (254) 724 688 172 / (254) 751 041 184                                       *
'*   E-mail Address:     maselv_e@yahoo.co.uk / masika_elvas@programmer.net / masika_elvas@live.com  *
'*   Location            BUNGOMA, KENYA                                                              *
'*****************************************************************************************************
'                                                                                                    -
'-----------------------------------------------------------------------------------------------------
'->  AutoSizer {Tue 14th Dec 2010}                                                                   -
'-----------------------------------------------------------------------------------------------------
'
'->  DESCRIPTION:                                                                                    -
'                                                                                                    -
'    AutoSizer is an ActiveX control that adds intelligent form resizing to your applications without
'    any code. AutoSizer provides simple proportional resizing just by adding it to your form. Unlike
'    other resizing controls, AutoSizer additionally provides intelligent Resizing. You can
'    individually adjust the size, position and font of each control on your form. Not case-sensitive
'-----------------------------------------------------------------------------------------------------
'
'    INTELLIGENT TAG EFFECT
'
'    Drop a AutoSizer control anywhere on the Form (it's invisible at runtime). Position all the
'    controls the way you want them to be. Then For each control (button, image, list, ...) you need
'    to set a rule to tell Resizer how to move, anchor or resize it. Setting a rule is done by entering
'    some characters in the Tag Property.
'
'    AutoSizer:X, AutoSizer:Y moves the controls when the form is resized, and AutoSizer:W, AutoSizer:H
'    resizes the controls when the form is resized
'
'    The following table shows what character to enter in the tag field and the effect it has:
'
'    AutoSizer:X   =>  Moves the control horizontally with the Right border of the form
'    AutoSizer:Y   =>  Moves the control vertically with the Lower Border of the form
'    AutoSizer:W   =>  Resizes the Width of the control with the Right border of the form
'    AutoSizer:H   =>  Resizes the Height of the control with the Lower Border of the form
'    AutoSizer:F   =>  Scales the font of the control proportionally to the new size of the form
'    AutoSizer:C   =>  Centers the control in the form
'    AutoSizer:WH  => Control's width and height resize with the form borders proportionally
'    AutoSizer:XH  => Control moves horizontally and Height resizes with the form border
'----------------------------------------------------------------------------------------------------
'
'    -> TIPS
'
'    1 - Resizing grid columns, managing other difficult controls:
'
'    AutoSizer can manage controls having standard Top, Left, Width Height and Tag properties.
'    However some controls such as the line and shape don't have those standard properties. In this
'    case AutoSizer gives you flexibility to handle them manually by adding code to the BeforeResize
'    and AfterResize events. The same apply for resizing grid rows and columns or controls contained
'    on a Tab control.
'
'    2 - Add AutoSizer to an existing project:
'
'    Adding AutoSizer to existing projects is very simple and easy. Just Drop your AutoSizer control
'    anywhere on your form (it's invisible at runtime). If you want to use advanced rule based resizing,
'    set the Tag properties of each control.
'
'    3 - Remove AutoSizer from a project:
'
'    Removing an AutoSizer control from your projects is also very easy. Just delete the AutoSizer control.
'    You don't need to remove the Tag Properties. Actually, don't remove the Tag Property of the controls
'    because if at a later time you decide to add the AutoSizer, all the resizing rules will be already set.
'
'    4 - Tips for Font Resizing:
'
'    If you use ResizeFonts Property, use scalable fonts. Usually True Type Fonts are easily scalable by
'    windows. The default visual basic font (Ms Sans Serif) is not! it has a minimum font size of 8 and
'    resizes to only the available font sizes (8, 10, 12, 14, 18, 24), it does not resize proportionally.
'    The font Arial for example can be resized to any font size proportionally (even to sizes not available
'    on the size list).
'
'    5 - Tips for Image Resizing:
'
'    If There is a picture box on your form, you need to set its property "Stretch = true" if you want the
'    image to be scaled.
'
'    For example to resize a command button horizontally 'AutoSizer:W' in its tag property. To only move the
'    button horizontally you would type 'AutoSizer:X'.
'
'    If you want a control to remain unchanged and in the same position, let its tag property not to start with
'    the word 'AutoSizer:' or leave it empty.
'
'    Of course to cover every possible way in the most complex interfaces, you can combine letters in the Tag
'    property. To move vertically and resize horizontally and scale the font you would type ‘AutoSizer:YWF’
'-----------------------------------------------------------------------------------------------------------
'    -> PRIORITY
'
'    Of course, you can't tell a control to move horizontally and resize horizontally at the same time.
'    There is a priority of Centering (C), over resizing (W,H) over moving (X,Y).
'-----------------------------------------------------------------------------------------------------------
'    -> LIMITATIONS
'
'    1 - AutoSizer uses the Tag property of the controls. Therefore you will not be able to use the tag property
'    starting with 'AutoSizer:' in your programming logic.
'
'    2 - Due to how Visual Basic enumerates controls on the form, The Zorder method should not be used. Changing
'    the Zorder property of any control on the form will cause unpredictable results.
'
'    3 - When you add a control programmatically at runtime, you need to inform AutoSizer by executing its
'    GetInitialPositions method
'
'-----------------------------------------------------------------------------------------------------------

Option Explicit

Private ParentDimensions$
Private FrmParentDimensions$
Private IsLoading As Boolean
Private FrmParentObj As Form
Private CtrlsPosition As New Collection
Private ParentPosition As New Collection

Private WithEvents FrmParent As Form
Attribute FrmParent.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'CONTROL PROPERTY VARIABLES
'----------------------------------------------------------------------------------------------------
Private auto_Enabled As Boolean
Private auto_MinWidth&, auto_MinHeight&, auto_MaxWidth&, auto_MaxHeight&
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'CONTROL EVENTS
'----------------------------------------------------------------------------------------------------
Public Event AfterResize()
Public Event BeforeResize()
'----------------------------------------------------------------------------------------------------

Public Function GetInitialPositions() As Boolean
On Local Error Resume Next
    
    If Not auto_Enabled Then Exit Function
    
    Dim MousePointerState%
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    Set FrmParent = FrmParentObj
    
    Dim iCnt&, i&, j&
    Dim FrmParentName$
    Dim Ctrl As Control
    
    'Clear initial entries in the CtrlsPosition Collection
    Do While CtrlsPosition.Count > &H0
        CtrlsPosition.Remove &H1 'Remove from the list
    Loop 'Move to the next entry
    
    Dim CtrlName$
    
    'Add each control's dimensions to the CtrlsPosition Collection
    For Each Ctrl In FrmParent.Controls
        
        If Ctrl.Tag <> VBA.vbNullString And VBA.Left$(Ctrl.Tag, VBA.Len("AutoSizer:")) = "AutoSizer:" Then
            
            iCnt = -&H1: iCnt = Ctrl.Index
            
            If iCnt <> -&H1 Then
                CtrlName = Ctrl.Name & "(" & Ctrl.Index & ")"
            Else
                CtrlName = Ctrl.Name
            End If
            
            Dim ObjContainer As Object
            
            If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "C") <> &H0 Then
                Set ObjContainer = Ctrl.Container
            Else
                Set ObjContainer = FrmParent
            End If
            
            j = -&H1: j = ObjContainer.Index
            
            If j <> -&H1 Then
                FrmParentName = ObjContainer.Name & "(" & ObjContainer.Index & ")"
            Else
                FrmParentName = ObjContainer.Name
            End If
            
            ParentPosition.Add ObjContainer.Left & ":" & ObjContainer.Width & ":" & ObjContainer.Top & ":" & ObjContainer.Height, CtrlName & " " & FrmParentName & ".Dimensions"
            
            If TypeOf Ctrl Is Line Then
                
                CtrlsPosition.Add Ctrl.X1, CtrlName & ".X1"
                CtrlsPosition.Add Ctrl.X2, CtrlName & ".X2"
                CtrlsPosition.Add Ctrl.Y1, CtrlName & ".Y1"
                CtrlsPosition.Add Ctrl.Y2, CtrlName & ".Y2"
                
            Else
                
                CtrlsPosition.Add Ctrl.Left, CtrlName & ".Left"
                CtrlsPosition.Add Ctrl.Width, CtrlName & ".Width"
                CtrlsPosition.Add Ctrl.Top, CtrlName & ".Top"
                CtrlsPosition.Add Ctrl.Height, CtrlName & ".Height"
                CtrlsPosition.Add Ctrl.FontSize / Parent.ScaleHeight, CtrlName & ".FontSize"
                
            End If
            
        End If
        
    Next Ctrl 'Move to the next control in the specified Form
    
    'Reset to the current Mouse Pointer state
    Screen.MousePointer = MousePointerState
    
End Function

Public Function AutoResize() As Boolean
On Local Error Resume Next
    
    If Not auto_Enabled Then Exit Function
    
    If FrmParent.Width > auto_MaxWidth And auto_MaxWidth > &H0 Then FrmParent.Width = auto_MaxWidth
    If FrmParent.Height > auto_MaxHeight And auto_MaxHeight > &H0 Then FrmParent.Height = auto_MaxHeight
    If FrmParent.Width < auto_MinWidth And auto_MinWidth > &H0 Then FrmParent.Width = auto_MinWidth
    If FrmParent.Height < auto_MinHeight And auto_MinHeight > &H0 Then FrmParent.Height = auto_MinHeight
    
    Dim i&, j&
    Dim Ctrl As Control
    Dim CtrlArrayFrm() As String
    Dim CtrlName$, FrmParentName$
    
    Dim MousePointerState%
    
    RaiseEvent BeforeResize
    
    'Get the current Mouse Pointer state
    MousePointerState = Screen.MousePointer
    
    'Indicate that a process or operation is in progress.
    Screen.MousePointer = vbHourglass
    
    For Each Ctrl In FrmParent.Controls
        
        If Ctrl.Tag <> VBA.vbNullString And VBA.Left$(Ctrl.Tag, VBA.Len("AutoSizer:")) = "AutoSizer:" Then
            
            i = -&H1: i = Ctrl.Index
            
            If i <> -&H1 Then
                CtrlName = Ctrl.Name & "(" & Ctrl.Index & ")"
            Else
                CtrlName = Ctrl.Name
            End If
            
            Dim ObjContainer As Object
            
            If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "C") <> &H0 Then
                Set ObjContainer = Ctrl.Container
            Else
                Set ObjContainer = FrmParent
            End If
            
            j = -&H1: j = ObjContainer.Index
            
            If j <> -&H1 Then
                FrmParentName = ObjContainer.Name & "(" & ObjContainer.Index & ")"
            Else
                FrmParentName = ObjContainer.Name
            End If
            
            CtrlArrayFrm = VBA.Split(ParentPosition(CtrlName & " " & FrmParentName & ".Dimensions"), ":")
            
            If TypeOf Ctrl Is Line Then
                
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "Y") <> &H0 Then Ctrl.Y1 = ObjContainer.Height - ((CtrlArrayFrm(&H3)) - (VBA.Val(CtrlsPosition(CtrlName & ".Y1")))): Ctrl.Y2 = ObjContainer.Height - ((CtrlArrayFrm(&H3)) - (VBA.Val(CtrlsPosition(CtrlName & ".Y2"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "X") <> &H0 Then Ctrl.X1 = ObjContainer.Height - ((CtrlArrayFrm(&H1)) - (VBA.Val(CtrlsPosition(CtrlName & ".X1")))): Ctrl.X2 = ObjContainer.Height - ((CtrlArrayFrm(&H1)) - (VBA.Val(CtrlsPosition(CtrlName & ".X2"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "H") <> &H0 Then Ctrl.Y2 = ObjContainer.Height - ((CtrlArrayFrm(&H3)) - (VBA.Val(CtrlsPosition(CtrlName & ".Y2"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "W") <> &H0 Then Ctrl.X2 = ObjContainer.Width - ((CtrlArrayFrm(&H1)) - (VBA.Val(CtrlsPosition(CtrlName & ".X2"))))
                
            Else
                
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "Y") <> &H0 Then Ctrl.Top = ObjContainer.Height - ((CtrlArrayFrm(&H3)) - (VBA.Val(CtrlsPosition(CtrlName & ".Top"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "X") <> &H0 Then Ctrl.Left = ObjContainer.Width - ((CtrlArrayFrm(&H1)) - (VBA.Val(CtrlsPosition(CtrlName & ".Left"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "H") <> &H0 Then Ctrl.Height = ObjContainer.Height - ((CtrlArrayFrm(&H3)) - (VBA.Val(CtrlsPosition(CtrlName & ".Height"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "W") <> &H0 Then Ctrl.Width = ObjContainer.Width - ((CtrlArrayFrm(&H1)) - (VBA.Val(CtrlsPosition(CtrlName & ".Width"))))
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "F") <> &H0 Then Ctrl.FontSize = VBA.Val(CtrlsPosition(CtrlName & ".FontSize")) * Parent.ScaleHeight
                
                If VBA.InStr(VBA.Replace(VBA.UCase$(Ctrl.Tag), "AUTOSIZER:", VBA.vbNullString), "C") <> &H0 Then
                    
                    Ctrl.Top = VBA.Val(CtrlsPosition(CtrlName & ".Top"))
                    Ctrl.Height = ObjContainer.Height - (CtrlArrayFrm(&H3) - VBA.Val(CtrlsPosition(CtrlName & ".Height")))
                    Ctrl.Width = (VBA.Val(CtrlsPosition(CtrlName & ".Width")) * Ctrl.Height) / VBA.Val(CtrlsPosition(CtrlName & ".Height"))
                    
                    Debug.Print ObjContainer.Name
                    
                    If Ctrl.Width > ObjContainer.Width Then
                        
                        Ctrl.Height = (Ctrl.Height * ObjContainer.Width) / Ctrl.Width
                        Ctrl.Width = ObjContainer.Width - 100
                        
                    End If
                    
                    Ctrl.Left = (ObjContainer.Width / &H2) - (Ctrl.Width / &H2) + &HA
                    
                End If
                
            End If
            
        End If
        
    Next Ctrl
    
    'Reset to the current Mouse Pointer state
    Screen.MousePointer = MousePointerState
    
    RaiseEvent AfterResize
    
End Function

Private Sub FrmParent_Load()
    If auto_Enabled Then Call GetInitialPositions
End Sub

Private Sub FrmParent_Resize()
    If auto_Enabled Then Call AutoResize
End Sub

Private Sub UserControl_InitProperties()
    Set FrmParentObj = UserControl.Parent
    auto_Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    
    Set FrmParentObj = PropBag.ReadProperty("ParentForm", UserControl.Parent)
    
    auto_Enabled = PropBag.ReadProperty("Enabled", True)
    
    auto_MaxHeight = PropBag.ReadProperty("MaxHeight", &H0)
    auto_MaxWidth = PropBag.ReadProperty("MaxWidth", &H0)
    auto_MinHeight = PropBag.ReadProperty("MinHeight", &H0)
    auto_MinWidth = PropBag.ReadProperty("MinWidth", &H0)
    
    Set FrmParent = FrmParentObj
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = ImgAutoResizeIcon.Width: UserControl.Height = ImgAutoResizeIcon.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("ParentForm", FrmParentObj, UserControl.Parent)
    Call PropBag.WriteProperty("Enabled", auto_Enabled, True)
    
    Call PropBag.WriteProperty("MinWidth", auto_MinWidth, &H0)
    Call PropBag.WriteProperty("MaxWidth", auto_MaxWidth, &H0)
    Call PropBag.WriteProperty("MinHeight", auto_MinHeight, &H0)
    Call PropBag.WriteProperty("MaxHeight", auto_MaxHeight, &H0)
    
End Sub

'----------------------------------------------------------------------------------------------------
'           CONTROL PROPERTIES
'----------------------------------------------------------------------------------------------------

Public Property Get About() As String
    About = "Masika Elvas : maselv_e@yahoo.co.uk"
End Property

Public Property Let About(ByVal vNewValue As String)
    vNewValue = "Masika Elvas : maselv_e@yahoo.co.uk"
End Property

Public Property Get Enabled() As Boolean
    Enabled = auto_Enabled: UserControl.Enabled = auto_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
On Local Error GoTo Handle_Error
    
    auto_Enabled = vNewValue
    UserControl.Enabled = auto_Enabled
    PropertyChanged "Enabled"
    
Handle_Error:
    
End Property

Public Property Get MaxHeight() As Variant
    MaxHeight = auto_MaxHeight
End Property

Public Property Let MaxHeight(ByVal vNewValue As Variant)
On Local Error GoTo Handle_Error
    
    auto_MaxHeight = vNewValue
    PropertyChanged "MaxHeight"
    
Handle_Error:
    
End Property

Public Property Get MaxWidth() As Variant
    MaxWidth = auto_MaxWidth
End Property

Public Property Let MaxWidth(ByVal vNewValue As Variant)
On Local Error GoTo Handle_Error
    
    auto_MaxWidth = vNewValue
    PropertyChanged "MaxWidth"
    
Handle_Error:
    
End Property

Public Property Get MinHeight() As Variant
    MinHeight = auto_MinHeight
End Property

Public Property Let MinHeight(ByVal vNewValue As Variant)
On Local Error GoTo Handle_Error
    
    auto_MinHeight = vNewValue
    PropertyChanged "MinHeight"
    
Handle_Error:
    
End Property

Public Property Get MinWidth() As Variant
    MinWidth = auto_MinWidth
End Property

Public Property Let MinWidth(ByVal vNewValue As Variant)
On Local Error GoTo Handle_Error
    
    auto_MinWidth = vNewValue
    PropertyChanged "MinWidth"
    
Handle_Error:
    
End Property

'----------------------------------------------------------------------------------------------------
