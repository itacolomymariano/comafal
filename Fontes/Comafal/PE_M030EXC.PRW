/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PE_M030EXC�Autor  �Five Solutions      � Data �  02/07/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � PE - Ponto de Entrada na Exclus�o do Cadastro do Cliente   ���
���          � para possibilitar exclus�o autom�tica da conta cont�bil no ���
���          � Plano de Contas.                                           ���
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

User Function M030EXC
   cArea := GetArea()
   cAreaAntes := Alias()
   /************************
    Five Solutions Consultoria
    02/09/2010
    Alterado codigo para comentar exclus�o da conta contabil.
   cCtaCtb := SA1->A1_CONTA
   cAreaAntes := Alias()
   DbSelectArea("CT1")
   DbSetOrder(1)
   If DbSeek(xFilial("CT1")+cCtaCtb)
	  DbSelectArea(cAreaAntes)
      If MsgYESNO("Confirma Exclus�o da Conta Cont�bil "+cCtaCtb+" do Cliente "+Alltrim(SA1->A1_NOME))
         ExcSA1()
      EndIf
   EndIf
   ***************************************************************************************************/
   /************************
    * Five Solutions Consultoria
    * 02/09/2010
    * Implementado para excluir registro do item contabil referente ao cliente
    ***************************************************************************/
   cItCTB := "112"+SA1->A1_COD
   DbSelectArea("CTD") //Tabela de Itens Contabil
   DbSetOrder(1) //CTD_FILIAL+CTD_ITEM
   If DbSeek(xFilial("CT1")+cItCTB)
      If MsgYESNO("Confirma Exclus�o do Item Cont�bil "+cItCTB+" do Cliente "+Alltrim(SA1->A1_NOME))
         RecLock("CTD",.F.)
            Delete
         MsUnLock()
      EndIf
   EndIf
      
   DbSelectArea(cAreaAntes)
   RestArea(cArea)
Return
Static Function ExcSA1()
   lMsErroAuto := .F.
   cContaCli := SA1->A1_CONTA
   cDescCta  := Alltrim(SA1->A1_NOME)
   nOpc   := 4 
   aConta := {{"CT1_CONTA",cContaCli  ,Nil},;
              {"CT1_DESC01",cDescCta  ,Nil},;
              {"CT1_CLASSE","2"       ,Nil},; //2 - Anal�tica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","11201001",Nil}}  //11201001 - Conta Fixa para Conta Superior de Clientes.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
   If lMsErroAuto
      Alert("Ocorreu Erro Fatal na Exclus�o da Conta Cont�bil: "+Alltrim(cContaCli)+" no Plano de Contas, Favor reportar ao Depto. de TI")
   EndIf
Return