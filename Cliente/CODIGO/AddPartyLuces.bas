Attribute VB_Name = "AddPartyLuces"
Public Sub AgregarParticulasyLuces(Map As Integer)

    If Map = 1 Then
        Call General_Particle_Create(1, 49, 71, -1)
        Call General_Particle_Create(6, 49, 85, -1)
        Call General_Particle_Create(6, 50, 85, -1)
        Call General_Particle_Create(9, 49, 57, -1)
        Call General_Particle_Create(6, 60, 58, -1)
        Call General_Particle_Create(6, 60, 65, -1)
        Call General_Particle_Create(6, 39, 58, -1)
        Call General_Particle_Create(6, 39, 65, -1)
        Call General_Particle_Create(6, 22, 77, -1)
        Call General_Particle_Create(6, 13, 41, -1)
        Call General_Particle_Create(6, 31, 41, -1)
        Call General_Particle_Create(6, 68, 41, -1)
        Call General_Particle_Create(6, 86, 41, -1)
        Light.Create_Light_To_Map 49, 85, 3, 255, 255, 255
        Light.Create_Light_To_Map 50, 85, 3, 255, 255, 255
        Light.Create_Light_To_Map 60, 58, 3, 255, 255, 255
        Light.Create_Light_To_Map 60, 65, 3, 255, 255, 255
        Light.Create_Light_To_Map 39, 58, 3, 255, 255, 255
        Light.Create_Light_To_Map 39, 65, 3, 255, 255, 255
        Light.Create_Light_To_Map 22, 77, 3, 255, 255, 255
        Light.Create_Light_To_Map 13, 41, 3, 255, 255, 255
        Light.Create_Light_To_Map 31, 41, 3, 255, 255, 255
        Light.Create_Light_To_Map 68, 41, 3, 255, 255, 255
        Light.Create_Light_To_Map 86, 41, 3, 255, 255, 255
    End If
    
    If Map = 81 Then
        Call General_Particle_Create(6, 42, 84, -1)
        Call General_Particle_Create(6, 56, 84, -1)
        Call General_Particle_Create(6, 46, 68, -1)
        Call General_Particle_Create(6, 52, 68, -1)
        Call General_Particle_Create(53, 16, 46, -1)
        Call General_Particle_Create(53, 36, 46, -1)
        Call General_Particle_Create(6, 46, 48, -1)
        Call General_Particle_Create(6, 52, 48, -1)
        Call General_Particle_Create(6, 56, 46, -1)
        Call General_Particle_Create(6, 81, 46, -1)
        Call General_Particle_Create(53, 41, 29, -1)
        Call General_Particle_Create(53, 57, 29, -1)
        Call General_Particle_Create(6, 61, 31, -1)
        Call General_Particle_Create(6, 83, 31, -1)
        Call General_Particle_Create(6, 83, 43, -1)
        Call General_Particle_Create(6, 61, 43, -1)
        Call General_Particle_Create(53, 59, 46, -1)
        Call General_Particle_Create(53, 81, 46, -1)
        Call General_Particle_Create(6, 44, 18, -1)
        Call General_Particle_Create(6, 54, 18, -1)
        Call General_Particle_Create(6, 37, 13, -1)
        Call General_Particle_Create(6, 24, 9, -1)
        Light.Create_Light_To_Map 42, 85, 3, 255, 255, 255
        Light.Create_Light_To_Map 56, 85, 3, 255, 255, 255
        Light.Create_Light_To_Map 46, 69, 3, 255, 255, 255
        Light.Create_Light_To_Map 52, 69, 3, 255, 255, 255
        Light.Create_Light_To_Map 16, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 36, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 46, 49, 3, 255, 255, 255
        Light.Create_Light_To_Map 52, 49, 3, 255, 255, 255
        Light.Create_Light_To_Map 56, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 81, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 41, 30, 3, 255, 255, 255
        Light.Create_Light_To_Map 57, 30, 3, 255, 255, 255
        Light.Create_Light_To_Map 61, 32, 3, 255, 255, 255
        Light.Create_Light_To_Map 83, 32, 3, 255, 255, 255
        Light.Create_Light_To_Map 83, 44, 3, 255, 255, 255
        Light.Create_Light_To_Map 61, 44, 3, 255, 255, 255
        Light.Create_Light_To_Map 59, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 81, 47, 3, 255, 255, 255
        Light.Create_Light_To_Map 44, 19, 3, 255, 255, 255
        Light.Create_Light_To_Map 54, 19, 3, 255, 255, 255
        Light.Create_Light_To_Map 37, 14, 3, 255, 255, 255
        Light.Create_Light_To_Map 61, 14, 3, 255, 255, 255
        Light.Create_Light_To_Map 24, 10, 3, 255, 255, 255
    End If
    
    If Map = 20 Then
        Call General_Particle_Create(6, 32, 71, -1)
        Light.Create_Light_To_Map 32, 71, 3, 255, 255, 255
    End If
    
    If Map = 54 Then
        Call General_Particle_Create(6, 42, 40, -1)
        Call General_Particle_Create(6, 61, 40, -1)
        Call General_Particle_Create(6, 61, 54, -1)
        Call General_Particle_Create(6, 42, 54, -1)
        Light.Create_Light_To_Map 52, 49, 20, 255, 255, 255
    End If
    
    If Map = 100 Then
        Call General_Particle_Create(6, 41, 56, -1)
        Call General_Particle_Create(6, 60, 56, -1)
        Call General_Particle_Create(6, 60, 39, -1)
        Call General_Particle_Create(6, 41, 39, -1)
        Light.Create_Light_To_Map 41, 58, 3, 255, 255, 255
        Light.Create_Light_To_Map 60, 58, 3, 255, 255, 255
        Light.Create_Light_To_Map 60, 41, 3, 255, 255, 255
        Light.Create_Light_To_Map 41, 41, 3, 255, 255, 255
        Light.Create_Light_To_Map 50, 50, 15, 255, 255, 255
        Light.Create_Light_To_Map 26, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 33, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 40, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 47, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 55, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 62, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 69, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 76, 16, 3, 255, 255, 255
        Light.Create_Light_To_Map 26, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 33, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 40, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 47, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 55, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 62, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 69, 23, 3, 255, 255, 255
        Light.Create_Light_To_Map 76, 23, 3, 255, 255, 255
    End If

If Map = 18 Then
    Call General_Particle_Create(6, 54, 18, -1)
    Call General_Particle_Create(6, 41, 13, -1)
    Call General_Particle_Create(6, 48, 11, -1)
    Call General_Particle_Create(6, 54, 11, -1)
    Call General_Particle_Create(6, 61, 13, -1)
    Call General_Particle_Create(6, 41, 18, -1)
    Call General_Particle_Create(6, 48, 18, -1)
    Call General_Particle_Create(6, 47, 29, -1)
    Call General_Particle_Create(6, 55, 29, -1)
    Call General_Particle_Create(6, 23, 21, -1)
    Call General_Particle_Create(6, 38, 33, -1)
    Call General_Particle_Create(6, 26, 39, -1)
    Call General_Particle_Create(6, 35, 39, -1)
    Call General_Particle_Create(6, 28, 49, -1)
    Call General_Particle_Create(6, 34, 49, -1)
    Call General_Particle_Create(53, 29, 57, -1)
    Call General_Particle_Create(53, 33, 57, -1)
    Call General_Particle_Create(6, 34, 66, -1)
    Call General_Particle_Create(6, 28, 66, -1)
    Call General_Particle_Create(43, 13, 77, -1)
    Call General_Particle_Create(6, 10, 83, -1)
    Call General_Particle_Create(53, 21, 80, -1)
    Call General_Particle_Create(53, 29, 80, -1)
    Call General_Particle_Create(57, 37, 80, -1)
    Call General_Particle_Create(6, 47, 84, -1)
    Call General_Particle_Create(6, 55, 84, -1)
    Call General_Particle_Create(6, 70, 88, -1)
    Call General_Particle_Create(53, 69, 57, -1)
    Call General_Particle_Create(53, 73, 57, -1)
    Call General_Particle_Create(53, 67, 39, -1)
    Call General_Particle_Create(53, 76, 39, -1)
    Call General_Particle_Create(6, 79, 33, -1)
    Call General_Particle_Create(6, 64, 21, -1)
    Call General_Particle_Create(53, 42, 46, -1)
    Call General_Particle_Create(6, 49, 47, -1)
    Call General_Particle_Create(6, 53, 47, -1)
    Call General_Particle_Create(53, 60, 46, -1)
    Call General_Particle_Create(53, 45, 55, -1)
    Call General_Particle_Create(6, 42, 57, -1)
    Call General_Particle_Create(53, 57, 55, -1)
    Call General_Particle_Create(6, 60, 57, -1)
    Call General_Particle_Create(53, 44, 65, -1)
    Call General_Particle_Create(53, 44, 72, -1)
    Call General_Particle_Create(53, 58, 72, -1)
    Call General_Particle_Create(53, 58, 65, -1)
    Call General_Particle_Create(53, 11, 49, -1)
    Call General_Particle_Create(53, 89, 51, -1)
    Light.Create_Light_To_Map 41, 14, 3, 255, 255, 255
    Light.Create_Light_To_Map 48, 13, 3, 255, 255, 255
    Light.Create_Light_To_Map 54, 13, 3, 255, 255, 255
    Light.Create_Light_To_Map 61, 14, 3, 255, 255, 255
    Light.Create_Light_To_Map 41, 19, 3, 255, 255, 255
    Light.Create_Light_To_Map 48, 19, 3, 255, 255, 255
    Light.Create_Light_To_Map 47, 31, 3, 255, 255, 255
    Light.Create_Light_To_Map 55, 31, 3, 255, 255, 255
    Light.Create_Light_To_Map 23, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 38, 35, 3, 255, 255, 255
    Light.Create_Light_To_Map 26, 40, 3, 255, 255, 255
    Light.Create_Light_To_Map 35, 40, 3, 255, 255, 255
    Light.Create_Light_To_Map 28, 50, 3, 255, 255, 255
    Light.Create_Light_To_Map 34, 50, 3, 255, 255, 255
    Light.Create_Light_To_Map 29, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 33, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 34, 67, 3, 255, 255, 255
    Light.Create_Light_To_Map 28, 67, 3, 255, 255, 255
    Light.Create_Light_To_Map 10, 84, 3, 255, 255, 255
    Light.Create_Light_To_Map 21, 81, 3, 255, 255, 255
    Light.Create_Light_To_Map 29, 81, 3, 255, 255, 255
    Light.Create_Light_To_Map 37, 81, 3, 255, 255, 255
    Light.Create_Light_To_Map 47, 85, 3, 255, 255, 255
    Light.Create_Light_To_Map 55, 85, 3, 255, 255, 255
    Light.Create_Light_To_Map 70, 90, 3, 255, 255, 255
    Light.Create_Light_To_Map 69, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 73, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 67, 40, 3, 255, 255, 255
    Light.Create_Light_To_Map 76, 40, 3, 255, 255, 255
    Light.Create_Light_To_Map 79, 35, 3, 255, 255, 255
    Light.Create_Light_To_Map 64, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 42, 47, 3, 255, 255, 255
    Light.Create_Light_To_Map 49, 49, 3, 255, 255, 255
    Light.Create_Light_To_Map 53, 49, 3, 255, 255, 255
    Light.Create_Light_To_Map 60, 47, 3, 255, 255, 255
    Light.Create_Light_To_Map 45, 56, 2, 255, 255, 255
    Light.Create_Light_To_Map 42, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 57, 56, 2, 255, 255, 255
    Light.Create_Light_To_Map 60, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 44, 67, 3, 255, 255, 255
    Light.Create_Light_To_Map 44, 74, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 74, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 67, 3, 255, 255, 255
    Light.Create_Light_To_Map 11, 51, 3, 255, 255, 255
    Light.Create_Light_To_Map 89, 53, 3, 255, 255, 255
    Light.Create_Light_To_Map 30, 29, 11, 255, 255, 255
    Light.Create_Light_To_Map 72, 29, 11, 255, 255, 255
End If

If Map = 101 Then
    Call General_Particle_Create(103, 22, 25, -1)
    Call General_Particle_Create(103, 45, 25, -1)
    Call General_Particle_Create(103, 45, 41, -1)
    Call General_Particle_Create(103, 22, 41, -1)
    Call General_Particle_Create(100, 22, 58, -1)
    Call General_Particle_Create(100, 45, 58, -1)
    Call General_Particle_Create(100, 45, 75, -1)
    Call General_Particle_Create(100, 22, 75, -1)
    Call General_Particle_Create(101, 58, 58, -1)
    Call General_Particle_Create(101, 81, 58, -1)
    Call General_Particle_Create(101, 81, 75, -1)
    Call General_Particle_Create(101, 58, 75, -1)
    Call General_Particle_Create(6, 58, 25, -1)
    Call General_Particle_Create(6, 81, 25, -1)
    Call General_Particle_Create(6, 81, 41, -1)
    Call General_Particle_Create(6, 58, 41, -1)
    Light.Create_Light_To_Map 22, 27, 3, 255, 255, 255
    Light.Create_Light_To_Map 45, 27, 3, 255, 255, 255
    Light.Create_Light_To_Map 45, 43, 3, 255, 255, 255
    Light.Create_Light_To_Map 22, 43, 3, 255, 255, 255
    Light.Create_Light_To_Map 22, 60, 3, 255, 255, 255
    Light.Create_Light_To_Map 45, 60, 3, 255, 255, 255
    Light.Create_Light_To_Map 45, 77, 3, 255, 255, 255
    Light.Create_Light_To_Map 22, 77, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 60, 3, 255, 255, 255
    Light.Create_Light_To_Map 81, 60, 3, 255, 255, 255
    Light.Create_Light_To_Map 81, 77, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 77, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 27, 3, 255, 255, 255
    Light.Create_Light_To_Map 81, 27, 3, 255, 255, 255
    Light.Create_Light_To_Map 81, 43, 3, 255, 255, 255
    Light.Create_Light_To_Map 58, 43, 3, 255, 255, 255
    Light.Create_Light_To_Map 33, 34, 15, 255, 255, 255
    Light.Create_Light_To_Map 69, 68, 15, 255, 255, 255
    Light.Create_Light_To_Map 33, 68, 15, 255, 255, 255
    Light.Create_Light_To_Map 69, 34, 15, 255, 255, 255
End If

If Map = 99 Then
    Call General_Particle_Create(6, 41, 56, -1)
    Call General_Particle_Create(6, 60, 56, -1)
    Call General_Particle_Create(6, 60, 39, -1)
    Call General_Particle_Create(6, 41, 39, -1)
    Light.Create_Light_To_Map 41, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 60, 58, 3, 255, 255, 255
    Light.Create_Light_To_Map 60, 41, 3, 255, 255, 255
    Light.Create_Light_To_Map 41, 41, 3, 255, 255, 255
    Light.Create_Light_To_Map 50, 50, 15, 255, 255, 255
    Light.Create_Light_To_Map 26, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 33, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 40, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 47, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 55, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 62, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 69, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 76, 16, 3, 255, 255, 255
    Light.Create_Light_To_Map 26, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 33, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 40, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 47, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 55, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 62, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 69, 23, 3, 255, 255, 255
    Light.Create_Light_To_Map 76, 23, 3, 255, 255, 255
End If
    
    If Map = 94 Then
        Call General_Particle_Create(53, 19, 38, -1)
        Call General_Particle_Create(53, 29, 38, -1)
        Call General_Particle_Create(53, 48, 6, -1)
        Call General_Particle_Create(53, 52, 6, -1)
        Call General_Particle_Create(33, 45, 45, -1)
        Call General_Particle_Create(60, 50, 45, -1)
        Call General_Particle_Create(53, 45, 64, -1)
        Call General_Particle_Create(53, 55, 64, -1)
        Call General_Particle_Create(53, 45, 73, -1)
        Call General_Particle_Create(53, 55, 73, -1)
        Call General_Particle_Create(6, 38, 76, -1)
        Call General_Particle_Create(6, 62, 76, -1)
        Call General_Particle_Create(6, 45, 83, -1)
        Call General_Particle_Create(6, 55, 83, -1)
        Light.Create_Light_To_Map 38, 77, 3, 255, 255, 255
        Light.Create_Light_To_Map 62, 77, 3, 255, 255, 255
        Light.Create_Light_To_Map 45, 84, 3, 255, 255, 255
        Light.Create_Light_To_Map 55, 84, 3, 255, 255, 255
        Light.Create_Light_To_Map 48, 7, 3, 255, 255, 255
        Light.Create_Light_To_Map 52, 7, 3, 255, 255, 255
    End If
    
    If Map = 95 Then
        Call General_Particle_Create(6, 46, 61, -1)
        Call General_Particle_Create(6, 54, 61, -1)
        Call General_Particle_Create(6, 46, 50, -1)
        Call General_Particle_Create(6, 54, 50, -1)
        Call General_Particle_Create(6, 46, 42, -1)
        Call General_Particle_Create(6, 54, 42, -1)
        Call General_Particle_Create(6, 54, 34, -1)
        Call General_Particle_Create(6, 46, 34, -1)
        Call General_Particle_Create(6, 47, 25, -1)
        Call General_Particle_Create(6, 53, 25, -1)
        Call General_Particle_Create(6, 42, 25, -1)
        Call General_Particle_Create(6, 36, 16, -1)
        Call General_Particle_Create(6, 43, 16, -1)
        Call General_Particle_Create(6, 58, 25, -1)
        Call General_Particle_Create(6, 57, 16, -1)
        Call General_Particle_Create(6, 64, 16, -1)
        Light.Create_Light_To_Map 46, 43, 3, 255, 255, 255
        Light.Create_Light_To_Map 54, 43, 3, 255, 255, 255
        Light.Create_Light_To_Map 46, 35, 3, 255, 255, 255
        Light.Create_Light_To_Map 54, 35, 3, 255, 255, 255
        Light.Create_Light_To_Map 47, 26, 3, 255, 255, 255
        Light.Create_Light_To_Map 53, 26, 3, 255, 255, 255
        Light.Create_Light_To_Map 58, 26, 3, 255, 255, 255
        Light.Create_Light_To_Map 42, 26, 3, 255, 255, 255
        Light.Create_Light_To_Map 46, 51, 3, 255, 255, 255
        Light.Create_Light_To_Map 46, 62, 3, 255, 255, 255
        Light.Create_Light_To_Map 54, 51, 3, 255, 255, 255
        Light.Create_Light_To_Map 54, 62, 3, 255, 255, 255
        Light.Create_Light_To_Map 43, 17, 3, 255, 255, 255
        Light.Create_Light_To_Map 36, 17, 3, 255, 255, 255
        Light.Create_Light_To_Map 57, 17, 3, 255, 255, 255
        Light.Create_Light_To_Map 64, 17, 3, 255, 255, 255
    End If
    
    If Map = 72 Then
        Call General_Particle_Create(100, 40, 48, -1)
        Call General_Particle_Create(100, 64, 48, -1)
        Call General_Particle_Create(100, 40, 28, -1)
        Call General_Particle_Create(100, 64, 28, -1)
        Light.Create_Light_To_Map 52, 39, 18, 255, 255, 255
    End If
    
End Sub
