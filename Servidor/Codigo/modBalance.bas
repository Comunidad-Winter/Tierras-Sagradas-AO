Attribute VB_Name = "modBalance"
Option Explicit

Private Type balClases
    Guerrero As Single
    Cazador As Single
    Paladin As Single
    Asesino As Single
    Ladron As Single
    Bardo As Single
    Clerigo As Single
    Mago As Single
    Druida As Single
End Type

Private Type balModificaciones
    ModificadorEvasion As balClases
    ModificadorPoderAtaqueArmas As balClases
    ModificadorPoderAtaqueProyectiles As balClases
    ModicadorDañoClaseArmas As balClases
    ModicadorDañoClaseProyectiles As balClases
    ModEvasionDeEscudoClase As balClases
    AtaqueFisico As balClases
    AtaqueMagico As balClases
    DefensaFisica As balClases
    DefensaMagica As balClases
End Type


Public Balance As balModificaciones
Public Sub LoadBalance()

Dim l_file As clsIniReader

    Set l_file = New clsIniReader

    '@ load file
    l_file.Initialize App.Path & "\Dat\Balance.dat"
    
    'Evasion escudo
    Balance.ModEvasionDeEscudoClase.Asesino = l_file.GetValue("ModEvasionDeEscudoClase", "Asesino")
    Balance.ModEvasionDeEscudoClase.Bardo = l_file.GetValue("ModEvasionDeEscudoClase", "Bardo")
    Balance.ModEvasionDeEscudoClase.Cazador = l_file.GetValue("ModEvasionDeEscudoClase", "Cazador")
    Balance.ModEvasionDeEscudoClase.Clerigo = l_file.GetValue("ModEvasionDeEscudoClase", "Clerigo")
    Balance.ModEvasionDeEscudoClase.Druida = l_file.GetValue("ModEvasionDeEscudoClase", "Druida")
    Balance.ModEvasionDeEscudoClase.Guerrero = l_file.GetValue("ModEvasionDeEscudoClase", "Guerrero")
    Balance.ModEvasionDeEscudoClase.Ladron = l_file.GetValue("ModEvasionDeEscudoClase", "Ladron")
    Balance.ModEvasionDeEscudoClase.Mago = l_file.GetValue("ModEvasionDeEscudoClase", "Mago")
    Balance.ModEvasionDeEscudoClase.Paladin = l_file.GetValue("ModEvasionDeEscudoClase", "Paladin")
    
    'Daño clases
    Balance.ModicadorDañoClaseArmas.Asesino = l_file.GetValue("ModicadorDañoClaseArmas", "Asesino")
    Balance.ModicadorDañoClaseArmas.Bardo = l_file.GetValue("ModicadorDañoClaseArmas", "Bardo")
    Balance.ModicadorDañoClaseArmas.Cazador = l_file.GetValue("ModicadorDañoClaseArmas", "Cazador")
    Balance.ModicadorDañoClaseArmas.Clerigo = l_file.GetValue("ModicadorDañoClaseArmas", "Clerigo")
    Balance.ModicadorDañoClaseArmas.Druida = l_file.GetValue("ModicadorDañoClaseArmas", "Druida")
    Balance.ModicadorDañoClaseArmas.Guerrero = l_file.GetValue("ModicadorDañoClaseArmas", "Guerrero")
    Balance.ModicadorDañoClaseArmas.Ladron = l_file.GetValue("ModicadorDañoClaseArmas", "Ladron")
    Balance.ModicadorDañoClaseArmas.Mago = l_file.GetValue("ModicadorDañoClaseArmas", "Mago")
    Balance.ModicadorDañoClaseArmas.Paladin = l_file.GetValue("ModicadorDañoClaseArmas", "Paladin")
    
    'Daño proyectiles
    Balance.ModicadorDañoClaseProyectiles.Asesino = l_file.GetValue("ModicadorDañoClaseProyectiles", "Asesino")
    Balance.ModicadorDañoClaseProyectiles.Bardo = l_file.GetValue("ModicadorDañoClaseProyectiles", "Bardo")
    Balance.ModicadorDañoClaseProyectiles.Cazador = l_file.GetValue("ModicadorDañoClaseProyectiles", "Cazador")
    Balance.ModicadorDañoClaseProyectiles.Clerigo = l_file.GetValue("ModicadorDañoClaseProyectiles", "Clerigo")
    Balance.ModicadorDañoClaseProyectiles.Druida = l_file.GetValue("ModicadorDañoClaseProyectiles", "Druida")
    Balance.ModicadorDañoClaseProyectiles.Guerrero = l_file.GetValue("ModicadorDañoClaseProyectiles", "Guerrero")
    Balance.ModicadorDañoClaseProyectiles.Ladron = l_file.GetValue("ModicadorDañoClaseProyectiles", "Ladron")
    Balance.ModicadorDañoClaseProyectiles.Mago = l_file.GetValue("ModicadorDañoClaseProyectiles", "Mago")
    Balance.ModicadorDañoClaseArmas.Paladin = l_file.GetValue("ModicadorDañoClaseProyectiles", "Paladin")
    
    'Evasion clase
    Balance.ModificadorEvasion.Asesino = l_file.GetValue("ModificadorEvasion", "Asesino")
    Balance.ModificadorEvasion.Bardo = l_file.GetValue("ModificadorEvasion", "Bardo")
    Balance.ModificadorEvasion.Cazador = l_file.GetValue("ModificadorEvasion", "Cazador")
    Balance.ModificadorEvasion.Clerigo = l_file.GetValue("ModificadorEvasion", "Clerigo")
    Balance.ModificadorEvasion.Druida = l_file.GetValue("ModificadorEvasion", "Druida")
    Balance.ModificadorEvasion.Guerrero = l_file.GetValue("ModificadorEvasion", "Guerrero")
    Balance.ModificadorEvasion.Ladron = l_file.GetValue("ModificadorEvasion", "Ladron")
    Balance.ModificadorEvasion.Mago = l_file.GetValue("ModificadorEvasion", "Mago")
    Balance.ModificadorEvasion.Paladin = l_file.GetValue("ModificadorEvasion", "Paladin")
    
    'Ataque c/armas
    Balance.ModificadorPoderAtaqueArmas.Asesino = l_file.GetValue("ModificadorPoderAtaqueArmas", "Asesino")
    Balance.ModificadorPoderAtaqueArmas.Bardo = l_file.GetValue("ModificadorPoderAtaqueArmas", "Bardo")
    Balance.ModificadorPoderAtaqueArmas.Cazador = l_file.GetValue("ModificadorPoderAtaqueArmas", "Cazador")
    Balance.ModificadorPoderAtaqueArmas.Clerigo = l_file.GetValue("ModificadorPoderAtaqueArmas", "Clerigo")
    Balance.ModificadorPoderAtaqueArmas.Druida = l_file.GetValue("ModificadorPoderAtaqueArmas", "Druida")
    Balance.ModificadorPoderAtaqueArmas.Guerrero = l_file.GetValue("ModificadorPoderAtaqueArmas", "Guerrero")
    Balance.ModificadorPoderAtaqueArmas.Ladron = l_file.GetValue("ModificadorPoderAtaqueArmas", "Ladron")
    Balance.ModificadorPoderAtaqueArmas.Mago = l_file.GetValue("ModificadorPoderAtaqueArmas", "Mago")
    Balance.ModificadorPoderAtaqueArmas.Paladin = l_file.GetValue("ModificadorPoderAtaqueArmas", "Paladin")
    
    'Ataque c/proyectiles
    Balance.ModificadorPoderAtaqueProyectiles.Asesino = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Asesino")
    Balance.ModificadorPoderAtaqueProyectiles.Bardo = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Bardo")
    Balance.ModificadorPoderAtaqueProyectiles.Cazador = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Cazador")
    Balance.ModificadorPoderAtaqueProyectiles.Clerigo = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Clerigo")
    Balance.ModificadorPoderAtaqueProyectiles.Druida = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Druida")
    Balance.ModificadorPoderAtaqueProyectiles.Guerrero = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Guerrero")
    Balance.ModificadorPoderAtaqueProyectiles.Ladron = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Ladron")
    Balance.ModificadorPoderAtaqueProyectiles.Mago = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Mago")
    Balance.ModificadorPoderAtaqueProyectiles.Paladin = l_file.GetValue("ModificadorPoderAtaqueProyectiles", "Paladin")
    
    'Ataque físico
    Balance.AtaqueFisico.Asesino = l_file.GetValue("AtaqueFisico", "Asesino")
    Balance.AtaqueFisico.Bardo = l_file.GetValue("AtaqueFisico", "Bardo")
    Balance.AtaqueFisico.Cazador = l_file.GetValue("AtaqueFisico", "Cazador")
    Balance.AtaqueFisico.Clerigo = l_file.GetValue("AtaqueFisico", "Clerigo")
    Balance.AtaqueFisico.Druida = l_file.GetValue("AtaqueFisico", "Druida")
    Balance.AtaqueFisico.Guerrero = l_file.GetValue("AtaqueFisico", "Guerrero")
    Balance.AtaqueFisico.Ladron = l_file.GetValue("AtaqueFisico", "Ladron")
    Balance.AtaqueFisico.Mago = l_file.GetValue("AtaqueFisico", "Mago")
    Balance.AtaqueFisico.Paladin = l_file.GetValue("AtaqueFisico", "Paladin")
    
    'Ataque mágico
    Balance.AtaqueMagico.Asesino = l_file.GetValue("AtaqueMagico", "Asesino")
    Balance.AtaqueMagico.Bardo = l_file.GetValue("AtaqueMagico", "Bardo")
    Balance.AtaqueMagico.Cazador = l_file.GetValue("AtaqueMagico", "Cazador")
    Balance.AtaqueMagico.Clerigo = l_file.GetValue("AtaqueMagico", "Clerigo")
    Balance.AtaqueMagico.Druida = l_file.GetValue("AtaqueMagico", "Druida")
    Balance.AtaqueMagico.Guerrero = l_file.GetValue("AtaqueMagico", "Guerrero")
    Balance.AtaqueMagico.Ladron = l_file.GetValue("AtaqueMagico", "Ladron")
    Balance.AtaqueMagico.Mago = l_file.GetValue("AtaqueMagico", "Mago")
    Balance.AtaqueMagico.Paladin = l_file.GetValue("AtaqueMagico", "Paladin")

    'Defensa fisica
    Balance.DefensaFisica.Asesino = l_file.GetValue("DefensaFisica", "Asesino")
    Balance.DefensaFisica.Bardo = l_file.GetValue("DefensaFisica", "Bardo")
    Balance.DefensaFisica.Cazador = l_file.GetValue("DefensaFisica", "Cazador")
    Balance.DefensaFisica.Clerigo = l_file.GetValue("DefensaFisica", "Clerigo")
    Balance.DefensaFisica.Druida = l_file.GetValue("DefensaFisica", "Druida")
    Balance.DefensaFisica.Guerrero = l_file.GetValue("DefensaFisica", "Guerrero")
    Balance.DefensaFisica.Ladron = l_file.GetValue("DefensaFisica", "Ladron")
    Balance.DefensaFisica.Mago = l_file.GetValue("DefensaFisica", "Mago")
    Balance.DefensaFisica.Paladin = l_file.GetValue("DefensaFisica", "Paladin")
    
    'Defensa mágica
    Balance.DefensaMagica.Asesino = l_file.GetValue("DefensaMagica", "Asesino")
    Balance.DefensaMagica.Bardo = l_file.GetValue("DefensaMagica", "Bardo")
    Balance.DefensaMagica.Cazador = l_file.GetValue("DefensaMagica", "Cazador")
    Balance.DefensaMagica.Clerigo = l_file.GetValue("DefensaMagica", "Clerigo")
    Balance.DefensaMagica.Druida = l_file.GetValue("DefensaMagica", "Druida")
    Balance.DefensaMagica.Guerrero = l_file.GetValue("DefensaMagica", "Guerrero")
    Balance.DefensaMagica.Ladron = l_file.GetValue("DefensaMagica", "Ladron")
    Balance.DefensaMagica.Mago = l_file.GetValue("DefensaMagica", "Mago")
    Balance.DefensaMagica.Paladin = l_file.GetValue("DefensaMagica", "Paladin")
End Sub
Function ModificarAtaqueFisico(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Guerrero
        Case "CAZADOR"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Cazador
        Case "PALADIN"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Paladin
        Case "ASESINO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Asesino
        Case "LADRON"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Ladron
        Case "BARDO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Bardo
        Case "CLERIGO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Clerigo
        Case "MAGO"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Mago
        Case "DRUIDA"
            ModificarAtaqueFisico = Balance.AtaqueFisico.Druida
        Case Else
            ModificarAtaqueFisico = 0
    End Select
    
End Function
Function ModificarAtaqueMagico(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Guerrero
        Case "CAZADOR"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Cazador
        Case "PALADIN"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Paladin
        Case "ASESINO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Asesino
        Case "LADRON"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Ladron
        Case "BARDO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Bardo
        Case "CLERIGO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Clerigo
        Case "MAGO"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Mago
        Case "DRUIDA"
            ModificarAtaqueMagico = Balance.AtaqueMagico.Druida
        Case Else
            ModificarAtaqueMagico = 0
    End Select
    
End Function
Function ModificarDefensaFisica(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarDefensaFisica = Balance.DefensaFisica.Guerrero
        Case "CAZADOR"
            ModificarDefensaFisica = Balance.DefensaFisica.Cazador
        Case "PALADIN"
            ModificarDefensaFisica = Balance.DefensaFisica.Paladin
        Case "ASESINO"
            ModificarDefensaFisica = Balance.DefensaFisica.Asesino
        Case "LADRON"
            ModificarDefensaFisica = Balance.DefensaFisica.Ladron
        Case "BARDO"
            ModificarDefensaFisica = Balance.DefensaFisica.Bardo
        Case "CLERIGO"
            ModificarDefensaFisica = Balance.DefensaFisica.Clerigo
        Case "MAGO"
            ModificarDefensaFisica = Balance.DefensaFisica.Mago
        Case "DRUIDA"
            ModificarDefensaFisica = Balance.DefensaFisica.Druida
        Case Else
            ModificarDefensaFisica = 0
    End Select
    
End Function
Function ModificarDefensaMagica(ByVal clase As String) As Single

    Select Case UCase$(clase)
        Case "GUERRERO"
            ModificarDefensaMagica = Balance.DefensaMagica.Guerrero
        Case "CAZADOR"
            ModificarDefensaMagica = Balance.DefensaMagica.Cazador
        Case "PALADIN"
            ModificarDefensaMagica = Balance.DefensaMagica.Paladin
        Case "ASESINO"
            ModificarDefensaMagica = Balance.DefensaMagica.Asesino
        Case "LADRON"
            ModificarDefensaMagica = Balance.DefensaMagica.Ladron
        Case "BARDO"
            ModificarDefensaMagica = Balance.DefensaMagica.Bardo
        Case "CLERIGO"
            ModificarDefensaMagica = Balance.DefensaMagica.Clerigo
        Case "MAGO"
            ModificarDefensaMagica = Balance.DefensaMagica.Mago
        Case "DRUIDA"
            ModificarDefensaMagica = Balance.DefensaMagica.Druida
        Case Else
            ModificarDefensaMagica = 0
    End Select
    
End Function
