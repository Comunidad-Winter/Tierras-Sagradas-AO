Attribute VB_Name = "Resolucion"
Public Sub SetResolucion()
 
    Dim lRes As Long
    Dim MidevM As typDevMODE
    Dim CambiarResolucion As Boolean
   
    lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, MidevM)
   
    oldResWidth = Screen.Width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.Height \ Screen.TwipsPerPixelY
   
    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If
   
    If CambiarResolucion Then
       
        With MidevM
            oldDepth = .dmBitsPerPel
            oldFrequency = .dmDisplayFrequency
           
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
            .dmPelsWidth = 800
            .dmPelsHeight = 600
            .dmBitsPerPel = 16
        End With
       
        lRes = ChangeDisplaySettings(MidevM, CDS_TEST)
    Else
        bNoResChange = True
    End If
End Sub
 
Public Sub ResetResolucion()
 
    Dim typDevM As typDevMODE
    Dim lRes As Long
   
    If Not bNoResChange Then
   
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
       
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
       
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub
