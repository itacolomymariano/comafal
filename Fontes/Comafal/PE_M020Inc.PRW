#include "TopConn.Ch"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PE_M020IncºAutor  ³Five Solutions      º Data ³  02/07/2008 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ PE - Ponto de Entrada na Inclusão do Cadastro de Fornecedo-º±±
±±º          ³ res para possibilitar inclusão automática da conta contábilº±±
±±º          ³  no Plano de Contas.                                       º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ COMAFAL PE/SP/RS                                          º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function M020Inc
   cArea := GetArea()
   cAreaAntes := Alias()
   /************************
   Five Solutions Consultoria
   02/09/2010
   Alterado codigo para evitar inclusão de conta contabil
   
   cCtaCtb := SA2->A2_CONTA
   cAreaAntes := Alias()
   DbSelectArea("CT1")
   DbSetOrder(1)
   If !DbSeek(xFilial("CT1")+cCtaCtb)
	  DbSelectArea(cAreaAntes)
      IncSA2x()
   EndIf
   *************************************************************/
   /**********************
   Five Solutions Consultoria
   02/09/2010
   Implementado codigo para incluir item contabil referente ao fornecedor
   ***********************************************************************/
   
   cChkFor := " SELECT COUNT(*) NQTCTD"
   cChkFor += "   FROM "+RetSQLName("CTD")
   cChkFor += "  WHERE CTD_FILIAL = '  '"
   cChkFor += "    AND CTD_ITEM = '210"+SA2->A2_COD+"'"
   cChkFor += "    AND D_E_L_E_T_ <> '*'"
   TCQuery cChkFor NEW ALIAS "XCTD"
   DbSelectArea("XCTD")
      nTemCTD := XCTD->NQTCTD
   DbCloseArea()
      
   If nTemCTD <= 0
            
      DbSelectArea("CTD")
      RecLock("CTD",.T.)
         CTD->CTD_FILIAL := SPACE(2)
         CTD->CTD_ITEM   := "210"+SA2->A2_COD
         CTD->CTD_DESC01 := Alltrim(SA2->A2_NOME)
         CTD->CTD_CLASSE := "2" //Analitico
         CTD->CTD_BLOQ   := "2" //Não Bloqueado
         CTD->CTD_DTEXIS := dDataBase
         //CTD->CTD_ITLP   := ""    //Item de Apuração de Resultados - (Lucros/Perdas)
         CTD->CTD_CLOBRG := "2" //Obriga digitação da Classe de Valor do respectivo Item Contábil - 1=Sim;2=Não
         CTD->CTD_ACCLVL := "1" //Aceita digitação da Classe de Valor - 1=Sim;2=Não
         
      MsUnLock()
      
   EndIf
         
   DbSelectArea(cAreaAntes)
   RestArea(cArea)
Return
Static Function IncSA2x()
   lMsErroAuto := .F.
   cContaFor := SA2->A2_CONTA
   cDescCta  := Alltrim(SA2->A2_NOME)
   nOpc   := 3 
   aConta := {{"CT1_CONTA",cContaFor  ,Nil},;
              {"CT1_DESC01",cDescCta  ,Nil},;
              {"CT1_CLASSE","2"       ,Nil},; //2 - Analítica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","21102",Nil}}  //21102 - Alterado em 11/06/2009 - Antes, 21102001 - Conta Fixa para Conta Superior de Fornecedores.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
   If lMsErroAuto
      Alert("Ocorreu Erro Fatal na Inclusão da Conta Contábil: "+Alltrim(cContaFor)+" no Plano de Contas, Favor reportar ao Depto. de TI")
   EndIf
Return