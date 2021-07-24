#Include "FIVEWIN.Ch"
//Five Solutions Consultoria - 22/03/2010
//Validação do Campo ZZ2_CODBAR - Evitar digitação no campo, permite apenas leitura através de código de barras.
User Function ChckKey
   MsgInfo("Novo 2 ChckKey")
   lVldKey := .T.
   cUltKey := ""
   Inkey()
   cUltKey := u_COMLastKey()  //LastKey() //
   If Type("cUltKey") == "N"
     Alert("Ultima Tecla Digitada "+Alltrim(Str(cUltKey)))
   EndIf
   If Type("cUltKey") <> "U"
      Alert("Ultima Tecla Digitada "+cUltKey)
   Else
      Alert("Ultima Tecla Digitada tem conteúdo LOGICO")
   EndIf  
   
   If !Empty(cUltKey)
      lVldKey := .F.
      Alert("Não é permitido digitação neste campo, favor utilizar LEITOR DE CODIGO DE BARRAS, Esta operação será cancelada!")
   EndIf
Return(lVldKey)


//Function TerLastKey()
User Function COMLastKey()
Local nRet
/*
If TerProtocolo()=="GRADUAL"
    MsgInfo("GRADUAL")
	nRet := Asc(__cLastKey)
	If nRet == 127
		LESC := .t.
		nRet := 27
	EndIf
ElseIf TerProtocolo()=="VT100"*/
   MsgInfo("VT100")
	nRet := VTLastKey()
/*
Else
MsgInfo("opcao nebhuma")
EndIf*/
Return nRet

//Function TerProtocolo(cProtocolo)
User Function COMProtocolo(cProtocolo)
If cProtocolo==NIL
	Return _Protocolo
EndIf
_Protocolo:=Upper(Alltrim(cProtocolo))
Return ''