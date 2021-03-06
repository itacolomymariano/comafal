#include "TopConn.Ch"
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PE_M020Inc�Autor  �Five Solutions      � Data �  02/07/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � PE - Ponto de Entrada na Inclus�o do Cadastro de Fornecedo-���
���          � res para possibilitar inclus�o autom�tica da conta cont�bil���
���          �  no Plano de Contas.                                       ���
���          �                                                            ���
���          �                                                            ���
���          �                                                            ���
���          �                                                            ���
���          �                                                            ���
���          �                                                            ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL PE/SP/RS                                          ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
User Function M020Inc
   cArea := GetArea()
   cAreaAntes := Alias()
   /************************
   Five Solutions Consultoria
   02/09/2010
   Alterado codigo para evitar inclus�o de conta contabil
   
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
         CTD->CTD_BLOQ   := "2" //N�o Bloqueado
         CTD->CTD_DTEXIS := dDataBase
         //CTD->CTD_ITLP   := ""    //Item de Apura��o de Resultados - (Lucros/Perdas)
         CTD->CTD_CLOBRG := "2" //Obriga digita��o da Classe de Valor do respectivo Item Cont�bil - 1=Sim;2=N�o
         CTD->CTD_ACCLVL := "1" //Aceita digita��o da Classe de Valor - 1=Sim;2=N�o
         
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
              {"CT1_CLASSE","2"       ,Nil},; //2 - Anal�tica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","21102",Nil}}  //21102 - Alterado em 11/06/2009 - Antes, 21102001 - Conta Fixa para Conta Superior de Fornecedores.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
   If lMsErroAuto
      Alert("Ocorreu Erro Fatal na Inclus�o da Conta Cont�bil: "+Alltrim(cContaFor)+" no Plano de Contas, Favor reportar ao Depto. de TI")
   EndIf
Return