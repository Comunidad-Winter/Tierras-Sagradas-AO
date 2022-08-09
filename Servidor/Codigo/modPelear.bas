Attribute VB_Name = "modPelear"
Option Explicit
 
Public Team1F As Boolean
Public kekeke As Byte
Public UsuarioPelea(1 To 8) As Integer
Public MapadelTorneo As Byte
Public Sub Pelear1vs1(ByVal userindex As Integer)

If UsuarioPelea(1) <> 0 And UsuarioPelea(2) <> 0 Then Call ResetearPeleas
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If

   If UsuarioPelea(1) = 0 Then
    UsuarioPelea(1) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
    Call SendData(SendTarget.toindex, userindex, 0, "||39")
   Else
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Then
        Call SendData(SendTarget.toindex, userindex, 0, "||40")
    Else
       UsuarioPelea(2) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
    End If
   End If
   
If UsuarioPelea(2) > 0 Then
    UserList(UsuarioPelea(1)).flags.MapaAnterior_dos = UserList(UsuarioPelea(1)).Pos.Map
    UserList(UsuarioPelea(1)).flags.XAnterior_dos = UserList(UsuarioPelea(1)).Pos.X
    UserList(UsuarioPelea(1)).flags.YAnterior_dos = UserList(UsuarioPelea(1)).Pos.Y
   
    UserList(UsuarioPelea(2)).flags.MapaAnterior_dos = UserList(UsuarioPelea(2)).Pos.Map
    UserList(UsuarioPelea(2)).flags.XAnterior_dos = UserList(UsuarioPelea(2)).Pos.X
    UserList(UsuarioPelea(2)).flags.YAnterior_dos = UserList(UsuarioPelea(2)).Pos.Y
    
    If UserList(UsuarioPelea(1)).Pos.Map = 100 And UserList(UsuarioPelea(2)).Pos.Map = 100 Then
      MapadelTorneo = 100
    ElseIf UserList(UsuarioPelea(1)).Pos.Map = 99 And UserList(UsuarioPelea(2)).Pos.Map = 99 Then
      MapadelTorneo = 99
    Else
      Call SendData(SendTarget.toindex, userindex, 0, "||41")
      Exit Sub
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, "||454@" & UserList(UsuarioPelea(1)).Name & "@" & UserList(UsuarioPelea(2)).Name)
    
    kekeke = 5
    MapaCont = MapadelTorneo
    Call SendData(SendTarget.ToMap, 0, MapaCont, "||455@" & kekeke)
    SendData SendTarget.ToMap, 0, MapaCont, "CU" & kekeke
    CuentaRegresiva = kekeke
   
    Call WarpUserChar(UsuarioPelea(1), MapadelTorneo, 41, 41)
    Call WarpUserChar(UsuarioPelea(2), MapadelTorneo, 60, 58)
End If
 
End Sub
Public Sub Pelear2vs2(ByVal userindex As Integer)

If UsuarioPelea(1) <> 0 And UsuarioPelea(2) <> 0 And UsuarioPelea(3) <> 0 And UsuarioPelea(4) <> 0 Then Call ResetearPeleas
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(2) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(3) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(4) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
   If Team1F = False Then
 
 
   
    If UsuarioPelea(1) = 0 Then
     UsuarioPelea(1) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||42")
    Else
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(2) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||44")
     Team1F = True
    End If
    End If
   
   Else
   
   If UsuarioPelea(3) = 0 Then
     UsuarioPelea(3) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||42")
    Else
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(3)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(4) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
    End If
    End If
   
   End If
   
If UsuarioPelea(4) > 0 Then
    UserList(UsuarioPelea(1)).flags.MapaAnterior_dos = UserList(UsuarioPelea(1)).Pos.Map
    UserList(UsuarioPelea(1)).flags.XAnterior_dos = UserList(UsuarioPelea(1)).Pos.X
    UserList(UsuarioPelea(1)).flags.YAnterior_dos = UserList(UsuarioPelea(1)).Pos.Y
    UserList(UsuarioPelea(2)).flags.MapaAnterior_dos = UserList(UsuarioPelea(2)).Pos.Map
    UserList(UsuarioPelea(2)).flags.XAnterior_dos = UserList(UsuarioPelea(2)).Pos.X
    UserList(UsuarioPelea(2)).flags.YAnterior_dos = UserList(UsuarioPelea(2)).Pos.Y
    UserList(UsuarioPelea(3)).flags.MapaAnterior_dos = UserList(UsuarioPelea(3)).Pos.Map
    UserList(UsuarioPelea(3)).flags.XAnterior_dos = UserList(UsuarioPelea(3)).Pos.X
    UserList(UsuarioPelea(3)).flags.YAnterior_dos = UserList(UsuarioPelea(3)).Pos.Y
    UserList(UsuarioPelea(4)).flags.MapaAnterior_dos = UserList(UsuarioPelea(4)).Pos.Map
    UserList(UsuarioPelea(4)).flags.XAnterior_dos = UserList(UsuarioPelea(4)).Pos.X
    UserList(UsuarioPelea(4)).flags.YAnterior_dos = UserList(UsuarioPelea(4)).Pos.Y
    
    If UserList(UsuarioPelea(1)).Pos.Map = 100 And UserList(UsuarioPelea(2)).Pos.Map = 100 And UserList(UsuarioPelea(3)).Pos.Map = 100 And UserList(UsuarioPelea(4)).Pos.Map = 100 Then
      MapadelTorneo = 100
    ElseIf UserList(UsuarioPelea(1)).Pos.Map = 99 And UserList(UsuarioPelea(2)).Pos.Map = 99 And UserList(UsuarioPelea(3)).Pos.Map = 99 And UserList(UsuarioPelea(4)).Pos.Map = 99 Then
      MapadelTorneo = 99
    Else
      Call SendData(SendTarget.toindex, userindex, 0, "||41")
      Exit Sub
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, "||456@" & UserList(UsuarioPelea(1)).Name & "@" & UserList(UsuarioPelea(2)).Name & "@" & UserList(UsuarioPelea(3)).Name & "@" & UserList(UsuarioPelea(4)).Name)
    
    kekeke = 5
    MapaCont = MapadelTorneo
    Call SendData(SendTarget.ToMap, 0, MapaCont, "||455@" & kekeke)
    SendData SendTarget.ToMap, 0, MapaCont, "CU" & kekeke
    CuentaRegresiva = kekeke
   
    Call WarpUserChar(UsuarioPelea(1), MapadelTorneo, 41, 41)
    Call WarpUserChar(UsuarioPelea(2), MapadelTorneo, 42, 42)
    Call WarpUserChar(UsuarioPelea(3), MapadelTorneo, 59, 57)
    Call WarpUserChar(UsuarioPelea(4), MapadelTorneo, 60, 58)
End If
   
End Sub
Public Sub Pelear3vs3(ByVal userindex As Integer)
 
If UsuarioPelea(1) <> 0 And UsuarioPelea(2) <> 0 And UsuarioPelea(3) <> 0 And UsuarioPelea(4) <> 0 And UsuarioPelea(5) <> 0 And UsuarioPelea(6) <> 0 Then Call ResetearPeleas
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(2) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(3) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(4) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(5) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(6) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
 
   If Team1F = False Then
    If UsuarioPelea(1) = 0 Then
     UsuarioPelea(1) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    ElseIf UsuarioPelea(2) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(2) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(3) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(2)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(3) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||44")
     Team1F = True
    End If
   End If
   
   Else
   
    If UsuarioPelea(4) = 0 Then
     UsuarioPelea(4) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    ElseIf UsuarioPelea(5) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(4)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(5) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(6) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(5)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(5)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(6) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
    End If
   End If
   
   End If
   
If UsuarioPelea(6) > 0 Then
   
    UserList(UsuarioPelea(1)).flags.MapaAnterior_dos = UserList(UsuarioPelea(1)).Pos.Map
    UserList(UsuarioPelea(1)).flags.XAnterior_dos = UserList(UsuarioPelea(1)).Pos.X
    UserList(UsuarioPelea(1)).flags.YAnterior_dos = UserList(UsuarioPelea(1)).Pos.Y
    UserList(UsuarioPelea(2)).flags.MapaAnterior_dos = UserList(UsuarioPelea(2)).Pos.Map
    UserList(UsuarioPelea(2)).flags.XAnterior_dos = UserList(UsuarioPelea(2)).Pos.X
    UserList(UsuarioPelea(2)).flags.YAnterior_dos = UserList(UsuarioPelea(2)).Pos.Y
    UserList(UsuarioPelea(3)).flags.MapaAnterior_dos = UserList(UsuarioPelea(3)).Pos.Map
    UserList(UsuarioPelea(3)).flags.XAnterior_dos = UserList(UsuarioPelea(3)).Pos.X
    UserList(UsuarioPelea(3)).flags.YAnterior_dos = UserList(UsuarioPelea(3)).Pos.Y
    UserList(UsuarioPelea(4)).flags.MapaAnterior_dos = UserList(UsuarioPelea(4)).Pos.Map
    UserList(UsuarioPelea(4)).flags.XAnterior_dos = UserList(UsuarioPelea(4)).Pos.X
    UserList(UsuarioPelea(4)).flags.YAnterior_dos = UserList(UsuarioPelea(4)).Pos.Y
    UserList(UsuarioPelea(5)).flags.MapaAnterior_dos = UserList(UsuarioPelea(5)).Pos.Map
    UserList(UsuarioPelea(5)).flags.XAnterior_dos = UserList(UsuarioPelea(5)).Pos.X
    UserList(UsuarioPelea(5)).flags.YAnterior_dos = UserList(UsuarioPelea(5)).Pos.Y
    UserList(UsuarioPelea(6)).flags.MapaAnterior_dos = UserList(UsuarioPelea(6)).Pos.Map
    UserList(UsuarioPelea(6)).flags.XAnterior_dos = UserList(UsuarioPelea(6)).Pos.X
    UserList(UsuarioPelea(6)).flags.YAnterior_dos = UserList(UsuarioPelea(6)).Pos.Y
    
    If UserList(UsuarioPelea(1)).Pos.Map = 100 And UserList(UsuarioPelea(2)).Pos.Map = 100 And UserList(UsuarioPelea(3)).Pos.Map = 100 And UserList(UsuarioPelea(4)).Pos.Map = 100 And UserList(UsuarioPelea(5)).Pos.Map = 100 And UserList(UsuarioPelea(6)).Pos.Map = 100 Then
      MapadelTorneo = 100
    ElseIf UserList(UsuarioPelea(1)).Pos.Map = 99 And UserList(UsuarioPelea(2)).Pos.Map = 99 And UserList(UsuarioPelea(3)).Pos.Map = 99 And UserList(UsuarioPelea(4)).Pos.Map = 99 And UserList(UsuarioPelea(5)).Pos.Map = 99 And UserList(UsuarioPelea(6)).Pos.Map = 99 Then
      MapadelTorneo = 99
    Else
      Call SendData(SendTarget.toindex, userindex, 0, "||41")
      Exit Sub
    End If
    
    Call SendData(SendTarget.ToAll, 0, 0, "||457@" & UserList(UsuarioPelea(1)).Name & "@" & UserList(UsuarioPelea(2)).Name & "@" & UserList(UsuarioPelea(3)).Name & "@" & UserList(UsuarioPelea(4)).Name & "@" & UserList(UsuarioPelea(5)).Name & "@" & UserList(UsuarioPelea(6)).Name)
    
    kekeke = 5
    MapaCont = MapadelTorneo
    Call SendData(SendTarget.ToMap, 0, MapaCont, "||455@" & kekeke)
    SendData SendTarget.ToMap, 0, MapaCont, "CU" & kekeke
    CuentaRegresiva = kekeke
   
    Call WarpUserChar(UsuarioPelea(1), MapadelTorneo, 41, 41)
    Call WarpUserChar(UsuarioPelea(2), MapadelTorneo, 42, 42)
    Call WarpUserChar(UsuarioPelea(3), MapadelTorneo, 41, 42)
    Call WarpUserChar(UsuarioPelea(4), MapadelTorneo, 59, 57)
    Call WarpUserChar(UsuarioPelea(5), MapadelTorneo, 60, 58)
    Call WarpUserChar(UsuarioPelea(6), MapadelTorneo, 60, 57)
End If
   
End Sub
Public Sub ResetearPeleas()
 
Team1F = False
UsuarioPelea(1) = 0
UsuarioPelea(2) = 0
UsuarioPelea(3) = 0
UsuarioPelea(4) = 0
UsuarioPelea(5) = 0
UsuarioPelea(6) = 0
UsuarioPelea(7) = 0
UsuarioPelea(8) = 0
 
End Sub
Public Sub Pelear4vs4(ByVal userindex As Integer)
 
If UsuarioPelea(1) <> 0 And UsuarioPelea(2) <> 0 And UsuarioPelea(3) <> 0 And UsuarioPelea(4) <> 0 And UsuarioPelea(5) <> 0 And UsuarioPelea(6) <> 0 And UsuarioPelea(7) <> 0 And UsuarioPelea(8) <> 0 Then Call ResetearPeleas
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(1) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(2) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(3) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(4) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(5) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
If NameIndex(UserList(UserList(userindex).flags.TargetUser).Name) = UsuarioPelea(6) Then
    Call SendData(SendTarget.toindex, userindex, 0, "||38")
Exit Sub
End If
 
 
   If Team1F = False Then
    If UsuarioPelea(1) = 0 Then
     UsuarioPelea(1) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    ElseIf UsuarioPelea(2) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(2) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(3) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(2)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(3) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(4) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(1)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(2)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(3)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(4) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||44")
     Team1F = True
    End If
   End If
   
   Else
   
    If UsuarioPelea(5) = 0 Then
     UsuarioPelea(5) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    ElseIf UsuarioPelea(6) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(5)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(6) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(7) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(5)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(6)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(7) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
     Call SendData(SendTarget.toindex, userindex, 0, "||45")
    End If
    ElseIf UsuarioPelea(8) = 0 Then
    If UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(5)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(6)).Name Or UserList(UserList(userindex).flags.TargetUser).Name = UserList(UsuarioPelea(7)).Name Then
     Call SendData(SendTarget.toindex, userindex, 0, "||43")
    Else
     UsuarioPelea(8) = NameIndex(UserList(UserList(userindex).flags.TargetUser).Name)
    End If
   End If
   
   End If
   
If UsuarioPelea(8) > 0 Then
   
    UserList(UsuarioPelea(1)).flags.MapaAnterior_dos = UserList(UsuarioPelea(1)).Pos.Map
    UserList(UsuarioPelea(1)).flags.XAnterior_dos = UserList(UsuarioPelea(1)).Pos.X
    UserList(UsuarioPelea(1)).flags.YAnterior_dos = UserList(UsuarioPelea(1)).Pos.Y
    UserList(UsuarioPelea(2)).flags.MapaAnterior_dos = UserList(UsuarioPelea(2)).Pos.Map
    UserList(UsuarioPelea(2)).flags.XAnterior_dos = UserList(UsuarioPelea(2)).Pos.X
    UserList(UsuarioPelea(2)).flags.YAnterior_dos = UserList(UsuarioPelea(2)).Pos.Y
    UserList(UsuarioPelea(3)).flags.MapaAnterior_dos = UserList(UsuarioPelea(3)).Pos.Map
    UserList(UsuarioPelea(3)).flags.XAnterior_dos = UserList(UsuarioPelea(3)).Pos.X
    UserList(UsuarioPelea(3)).flags.YAnterior_dos = UserList(UsuarioPelea(3)).Pos.Y
    UserList(UsuarioPelea(4)).flags.MapaAnterior_dos = UserList(UsuarioPelea(4)).Pos.Map
    UserList(UsuarioPelea(4)).flags.XAnterior_dos = UserList(UsuarioPelea(4)).Pos.X
    UserList(UsuarioPelea(4)).flags.YAnterior_dos = UserList(UsuarioPelea(4)).Pos.Y
    UserList(UsuarioPelea(5)).flags.MapaAnterior_dos = UserList(UsuarioPelea(5)).Pos.Map
    UserList(UsuarioPelea(5)).flags.XAnterior_dos = UserList(UsuarioPelea(5)).Pos.X
    UserList(UsuarioPelea(5)).flags.YAnterior_dos = UserList(UsuarioPelea(5)).Pos.Y
    UserList(UsuarioPelea(6)).flags.MapaAnterior_dos = UserList(UsuarioPelea(6)).Pos.Map
    UserList(UsuarioPelea(6)).flags.XAnterior_dos = UserList(UsuarioPelea(6)).Pos.X
    UserList(UsuarioPelea(6)).flags.YAnterior_dos = UserList(UsuarioPelea(6)).Pos.Y
    UserList(UsuarioPelea(7)).flags.MapaAnterior_dos = UserList(UsuarioPelea(7)).Pos.Map
    UserList(UsuarioPelea(7)).flags.XAnterior_dos = UserList(UsuarioPelea(7)).Pos.X
    UserList(UsuarioPelea(7)).flags.YAnterior_dos = UserList(UsuarioPelea(7)).Pos.Y
    UserList(UsuarioPelea(8)).flags.MapaAnterior_dos = UserList(UsuarioPelea(8)).Pos.Map
    UserList(UsuarioPelea(8)).flags.XAnterior_dos = UserList(UsuarioPelea(8)).Pos.X
    UserList(UsuarioPelea(8)).flags.YAnterior_dos = UserList(UsuarioPelea(8)).Pos.Y
    
    If UserList(UsuarioPelea(1)).Pos.Map = 100 And UserList(UsuarioPelea(2)).Pos.Map = 100 And UserList(UsuarioPelea(3)).Pos.Map = 100 And UserList(UsuarioPelea(4)).Pos.Map = 100 And UserList(UsuarioPelea(5)).Pos.Map = 100 And UserList(UsuarioPelea(6)).Pos.Map = 100 And UserList(UsuarioPelea(7)).Pos.Map = 100 And UserList(UsuarioPelea(8)).Pos.Map = 100 Then
      MapadelTorneo = 100
    ElseIf UserList(UsuarioPelea(1)).Pos.Map = 99 And UserList(UsuarioPelea(2)).Pos.Map = 99 And UserList(UsuarioPelea(3)).Pos.Map = 99 And UserList(UsuarioPelea(4)).Pos.Map = 99 And UserList(UsuarioPelea(5)).Pos.Map = 99 And UserList(UsuarioPelea(6)).Pos.Map = 99 And UserList(UsuarioPelea(7)).Pos.Map = 99 And UserList(UsuarioPelea(8)).Pos.Map = 99 Then
      MapadelTorneo = 99
    Else
      Call SendData(SendTarget.toindex, userindex, 0, "||41")
      Exit Sub
    End If
   
   Call SendData(SendTarget.ToAll, 0, 0, "||458@" & UserList(UsuarioPelea(1)).Name & "@" & UserList(UsuarioPelea(2)).Name & "@" & UserList(UsuarioPelea(3)).Name & "@" & UserList(UsuarioPelea(4)).Name & "@" & UserList(UsuarioPelea(5)).Name & "@" & UserList(UsuarioPelea(6)).Name & "@" & UserList(UsuarioPelea(7)).Name & "@" & UserList(UsuarioPelea(8)).Name)
   
    kekeke = 5
    MapaCont = MapadelTorneo
    Call SendData(SendTarget.ToMap, 0, MapaCont, "||455@" & kekeke)
    SendData SendTarget.ToMap, 0, MapaCont, "CU" & kekeke
    CuentaRegresiva = kekeke
   
    Call WarpUserChar(UsuarioPelea(1), MapadelTorneo, 41, 41)
    Call WarpUserChar(UsuarioPelea(2), MapadelTorneo, 42, 42)
    Call WarpUserChar(UsuarioPelea(3), MapadelTorneo, 41, 42)
    Call WarpUserChar(UsuarioPelea(4), MapadelTorneo, 42, 41)
    Call WarpUserChar(UsuarioPelea(5), MapadelTorneo, 60, 58)
    Call WarpUserChar(UsuarioPelea(6), MapadelTorneo, 60, 57)
    Call WarpUserChar(UsuarioPelea(7), MapadelTorneo, 60, 58)
    Call WarpUserChar(UsuarioPelea(8), MapadelTorneo, 59, 58)
End If
   
End Sub

