#include "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCARGACTD  บAutor  ณFive Solutions      บ Data ณ  01/06/2010 บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณ Carga da tabela CTD - Itens Contabeis com base no cadastro บฑฑ
ฑฑบ          ณ de Clientes e Fornecedores. As entidades serใo criadas     บฑฑ
ฑฑบ          ณ precedidas de um prefixo padrใo 112-Cliente e 210-Fornecedoบฑฑ
ฑฑบ          ณ res.                                                       บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ COMAFAL :PE,SP e RS.                                      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function CargaCTD
   
   Private nCtt := 0
   Private nFtt := 0
   
   If MsgYesNo("Deseja Realizar Carga de Item Contabil de Cad.Clientes")
      Processa({||RCliCTD()},"Processando Cad. Clientes ...")
   EndIf
   If MsgYesNo("Deseja Realizar Carga de Item Contabil de Cad.Fornecedores")
      Processa({||RForCTD()},"Processando Cad. Fornecedores ...")
   EndIf
   
   MsgInfo("Processo Concluido!, Itens Incluidos: "+Alltrim(Str(nCtt))+" Clientes, "+Alltrim(Str(nFtt))+" Fornecedores")
   
Return

Static Function RCliCTD
   
   For nCli := 1 To 2
      
      If nCli == 1
         cQryCli := " SELECT COUNT(*) NTTCLI "
      Else
         cQryCli := " SELECT A1_COD,A1_NOME "
      EndIf
      
      cQryCli += "   FROM "+RetSQLName("SA1")
      cQryCli += "  WHERE A1_FILIAL = '"+xFilial("SA1")+"'"
      cQryCli += "    AND D_E_L_E_T_ <> '*'"
   
      TCQuery cQryCli NEW ALIAS "XSA1"
      
      If nCli == 1
         DbSelectArea("XSA1")
         nReg := XSA1->NTTCLI
         DbCloseArea()
      EndIf
      
   Next nCli
   ProcRegua(nReg)
   nCtt := 1
   DbSelectArea("XSA1")
   While XSA1->(!Eof())
      
      IncProc("Registrando Item Contabil[SA1] ... "+Alltrim(Str(nCtt))+" / "+Alltrim(Str(nReg)))
      
      cChkCli := " SELECT COUNT(*) NQTCTD"
      cChkCli += "   FROM "+RetSQLName("CTD")
      cChkCli += "  WHERE CTD_FILIAL = '  '"
      cChkCli += "    AND CTD_ITEM = '112"+XSA1->A1_COD+"'"
      cChkCli += "    AND D_E_L_E_T_ <> '*'"
      TCQuery cChkCli NEW ALIAS "XCTD"
      DbSelectArea("XCTD")
         nTemCTD := XCTD->NQTCTD
      DbCloseArea()
      
      If nTemCTD <= 0
         
         DbSelectArea("CTD")
         RecLock("CTD",.T.)
            CTD->CTD_FILIAL := SPACE(2)
            CTD->CTD_ITEM   := "112"+XSA1->A1_COD
            CTD->CTD_DESC01 := Alltrim(XSA1->A1_NOME)
            CTD->CTD_CLASSE := "2" //Analitico
            CTD->CTD_BLOQ   := "2" //Nใo Bloqueado
            CTD->CTD_DTEXIS := dDataBase
            //CTD->CTD_ITLP   := ""    //Item de Apura็ใo de Resultados - (Lucros/Perdas)
            CTD->CTD_CLOBRG := "2" //Obriga digita็ใo da Classe de Valor do respectivo Item Contแbil - 1=Sim;2=Nใo
            CTD->CTD_ACCLVL := "1" //Aceita digita็ใo da Classe de Valor - 1=Sim;2=Nใo
         
         MsUnLock()
      
      EndIf
      
      DbSelectArea("XSA1")
      DbSkip()
      nCtt ++
   
   EndDo
   DbSelectArea("XSA1")
   DbCloseArea()
   
Return
Static Function RForCTD
   For nFor := 1 To 2
      
      If nFor == 1
         cQryFor := " SELECT COUNT(*) NTTFor "
      Else
         cQryFor := " SELECT A2_COD,A2_NOME "
      EndIf
      
      cQryFor += "   FROM "+RetSQLName("SA2")
      cQryFor += "  WHERE A2_FILIAL = '"+xFilial("SA2")+"'"
      cQryFor += "    AND D_E_L_E_T_ <> '*'"
   
      TCQuery cQryFor NEW ALIAS "XSA2"
      
      If nFor == 1
         DbSelectArea("XSA2")
         nReg := XSA2->NTTFor
         DbCloseArea()
      EndIf
      
   Next nFor
   ProcRegua(nReg)
   nFtt := 1
   DbSelectArea("XSA2")
   While XSA2->(!Eof())
      IncProc("Registrando Item Contabil[SA2] ... "+Alltrim(Str(nFtt))+" / "+Alltrim(Str(nReg)))
      
      cChkFor := " SELECT COUNT(*) NQTCTD"
      cChkFor += "   FROM "+RetSQLName("CTD")
      cChkFor += "  WHERE CTD_FILIAL = '  '"
      cChkFor += "    AND CTD_ITEM = '210"+XSA2->A2_COD+"'"
      cChkFor += "    AND D_E_L_E_T_ <> '*'"
      TCQuery cChkFor NEW ALIAS "XCTD"
      DbSelectArea("XCTD")
         nTemCTD := XCTD->NQTCTD
      DbCloseArea()
      
      If nTemCTD <= 0
            
         DbSelectArea("CTD")
         RecLock("CTD",.T.)
            CTD->CTD_FILIAL := SPACE(2)
            CTD->CTD_ITEM   := "210"+XSA2->A2_COD
            CTD->CTD_DESC01 := Alltrim(XSA2->A2_NOME)
            CTD->CTD_CLASSE := "2" //Analitico
            CTD->CTD_BLOQ   := "2" //Nใo Bloqueado
            CTD->CTD_DTEXIS := dDataBase
            //CTD->CTD_ITLP   := ""    //Item de Apura็ใo de Resultados - (Lucros/Perdas)
            CTD->CTD_CLOBRG := "2" //Obriga digita็ใo da Classe de Valor do respectivo Item Contแbil - 1=Sim;2=Nใo
            CTD->CTD_ACCLVL := "1" //Aceita digita็ใo da Classe de Valor - 1=Sim;2=Nใo
         
         MsUnLock()
      
      EndIf
      
      DbSelectArea("XSA2")
      DbSkip()
      nFtt ++
   
   EndDo
   DbSelectArea("XSA2")
   DbCloseArea()
Return