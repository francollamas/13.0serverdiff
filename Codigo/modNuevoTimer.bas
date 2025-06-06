Attribute VB_Name = "modIntervalos"
'Argentum Online 0.12.2
'Copyright (C) 2002 M�rquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim TActual As Long

TActual = timeGetTime

If TActual - UserList(Userindex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
    If actualizar Then
        UserList(Userindex).Counters.TimerLanzarSpell = TActual
    End If
    IntervaloPermiteLanzarSpell = True
Else
    IntervaloPermiteLanzarSpell = False
End If

End Function

Public Function IntervaloPermiteAtacar(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim TActual As Long

TActual = timeGetTime

If TActual - UserList(Userindex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
    If actualizar Then
        UserList(Userindex).Counters.TimerPuedeAtacar = TActual
        UserList(Userindex).Counters.TimerGolpeUsar = TActual
    End If
    IntervaloPermiteAtacar = True
Else
    IntervaloPermiteAtacar = False
End If
End Function

Public Function IntervaloPermiteGolpeUsar(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: ZaMa
'Checks if the time that passed from the last hit is enough for the user to use a potion.
'Last Modification: 06/04/2009
'***************************************************

Dim TActual As Long

TActual = timeGetTime

If TActual - UserList(Userindex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
    If actualizar Then
        UserList(Userindex).Counters.TimerGolpeUsar = TActual
    End If
    IntervaloPermiteGolpeUsar = True
Else
    IntervaloPermiteGolpeUsar = False
End If
End Function

Public Function NpcIntervaloGolpe(ByVal NpcIndex As Integer, _
                                  Optional ByVal actualizar As Boolean = True) As Boolean
    Dim TActual As Long
    
    TActual = timeGetTime
    
    If TActual - Npclist(NpcIndex).Contadores.intervaloGolpe >= IntervaloNPCPuedeAtacar Then
        If actualizar Then Npclist(NpcIndex).Contadores.intervaloGolpe = TActual
        NpcIntervaloGolpe = True
    Else
        NpcIntervaloGolpe = False
        Exit Function
    End If
End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    Dim TActual As Long
    
    With UserList(Userindex)
        If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
            Exit Function
        End If
        
        TActual = timeGetTime
        
        If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
            If actualizar Then
                .Counters.TimerMagiaGolpe = TActual
                .Counters.TimerPuedeAtacar = TActual
                .Counters.TimerGolpeUsar = TActual
            End If
            IntervaloPermiteMagiaGolpe = True
        Else
            IntervaloPermiteMagiaGolpe = False
        End If
    End With
End Function

Public Function IntervaloPermiteGolpeMagia(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim TActual As Long
    
    If UserList(Userindex).Counters.TimerGolpeMagia > UserList(Userindex).Counters.TimerPuedeAtacar Then
        Exit Function
    End If
    
    TActual = timeGetTime
    
    If TActual - UserList(Userindex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
        If actualizar Then
            UserList(Userindex).Counters.TimerGolpeMagia = TActual
            UserList(Userindex).Counters.TimerLanzarSpell = TActual
        End If
        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False
    End If
End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim TActual As Long
    
    TActual = timeGetTime
    
    If TActual - UserList(Userindex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
        If actualizar Then UserList(Userindex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False
    End If
End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 25/01/2010 (ZaMa)
'25/01/2010: ZaMa - General adjustments.
'***************************************************

    Dim TActual As Long
    
    TActual = timeGetTime
    
    If TActual - UserList(Userindex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
        If actualizar Then
            UserList(Userindex).Counters.TimerUsar = TActual
            'UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        
        'UserList(UserIndex).Counters.failedUsageAttempts = UserList(UserIndex).Counters.failedUsageAttempts + 1
        
        'Tolerancia arbitraria - 20 es MUY alta, la est� chiteando zarpado
        'If UserList(UserIndex).Counters.failedUsageAttempts = 20 Then
            'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(UserList(UserIndex).name & " kicked by the server por posible modificaci�n de intervalos.", FontTypeNames.FONTTYPE_FIGHT))
            'Call CloseSocket(UserIndex)
        'End If
    End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = True) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim TActual As Long
    
    TActual = timeGetTime
    
    If TActual - UserList(Userindex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
        If actualizar Then UserList(Userindex).Counters.TimerPuedeUsarArco = TActual
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
    End If

End Function

Public Function IntervaloPermiteSerAtacado(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: ZaMa
'Last Modify Date: 13/11/2009
'13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
'**************************************************************
    Dim TActual As Long
    
    TActual = timeGetTime
    
    With UserList(Userindex)
        ' Inicializa el timer
        If actualizar Then
            .Counters.TimerPuedeSerAtacado = TActual
            .flags.NoPuedeSerAtacado = True
            IntervaloPermiteSerAtacado = False
        Else
            If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
                .flags.NoPuedeSerAtacado = False
                IntervaloPermiteSerAtacado = True
            Else
                IntervaloPermiteSerAtacado = False
            End If
        End If
    End With

End Function

Public Function IntervaloPerdioNpc(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: ZaMa
'Last Modify Date: 13/11/2009
'13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
'**************************************************************
    Dim TActual As Long
    
    TActual = timeGetTime
    
    With UserList(Userindex)
        ' Inicializa el timer
        If actualizar Then
            .Counters.TimerPerteneceNpc = TActual
            IntervaloPerdioNpc = False
        Else
            If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
                IntervaloPerdioNpc = True
            Else
                IntervaloPerdioNpc = False
            End If
        End If
    End With

End Function

Public Function IntervaloEstadoAtacable(ByVal Userindex As Integer, Optional ByVal actualizar As Boolean = False) As Boolean
'**************************************************************
'Author: ZaMa
'Last Modify by: ZaMa
'Last Modify Date: 13/01/2010
'13/01/2010: ZaMa - Add the Timer which determines wether the user can be atacked by an user or not
'**************************************************************
    Dim TActual As Long
    
    TActual = timeGetTime
    
    With UserList(Userindex)
        ' Inicializa el timer
        If actualizar Then
            .Counters.TimerEstadoAtacable = TActual
            IntervaloEstadoAtacable = True
        Else
            If TActual - .Counters.TimerEstadoAtacable >= IntervaloAtacable Then
                IntervaloEstadoAtacable = False
            Else
                IntervaloEstadoAtacable = True
            End If
        End If
    End With

End Function
