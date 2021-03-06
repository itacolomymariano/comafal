#include "TopConn.Ch"
User Function TrtSF4
   Private nAtu  := 0
   If MsgYESNO("Confirma tratamento do campo F4_NATOPER")
      Processa({||RunSF4()},"Atualizando SF4, Aguarde ...")
   EndIf
   MsgInfo("Atualiza??o SF4, Conclu?da com Sucesso !!! "+Alltrim(Str(nAtu))+" Registros Atualizados")
Return
Static Function RunSF4
   For i := 1 To 2
      If i == 1
         cQrySF4 := "SELECT COUNT(*) AS TOTREG"
      Else
         cQrySF4 := "SELECT SF4.F4_CODIGO,SF4.F4_CF,R_E_C_N_O_ AS RECSF4"
      EndIf
      cQrySF4 += " FROM "+RetSQLName("SF4")+" SF4"
      If i > 1
         cQrySF4 += " ORDER BY SF4.F4_CF"
      EndIf
      
      TCQuery cQrySF4 NEW ALIAS "TSF4"
      
      If i == 1
         DbSelectArea("TSF4")
         nReg := TSF4->TOTREG
         DbCloseArea()
      EndIf
      
   Next i
   DbSelectArea("TSF4")
   nCont := 1
   
   While !Eof()
      IncProc("Gerando Sequencial Nat. Opera??o "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
      cCF  := TSF4->F4_CF
      nSeq := 1
      While TSF4->F4_CF == cCF .And. !Eof()
         DbSelectArea("SF4")
         DbSetOrder(1)//F4_FILIAL+F4_CODIGO
         If DbSeek(xFilial("SF4")+TSF4->F4_CODIGO)
            RecLock("SF4",.F.)
               SF4->F4_NATOPER := Alltrim(TSF4->F4_CF)+StrZero(nSeq,2)
            MsUnLock()
            nSeq ++
            nAtu ++
         EndIf
         DbSelectArea("TSF4")
         DbSkip()
         nCont ++
      EndDo
   EndDo
   DbSelectArea("TSF4")
   DbCloseArea()
Return