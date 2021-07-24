/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³PE_M030EXCºAutor  ³Five Solutions      º Data ³  02/07/2008 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ PE - Ponto de Entrada na Exclusão do Cadastro do Cliente   º±±
±±º          ³ para possibilitar exclusão automática da conta contábil no º±±
±±º          ³ Plano de Contas.                                           º±±
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

User Function M030EXC
   cArea := GetArea()
   cAreaAntes := Alias()
   /************************
    Five Solutions Consultoria
    02/09/2010
    Alterado codigo para comentar exclusão da conta contabil.
   cCtaCtb := SA1->A1_CONTA
   cAreaAntes := Alias()
   DbSelectArea("CT1")
   DbSetOrder(1)
   If DbSeek(xFilial("CT1")+cCtaCtb)
	  DbSelectArea(cAreaAntes)
      If MsgYESNO("Confirma Exclusão da Conta Contábil "+cCtaCtb+" do Cliente "+Alltrim(SA1->A1_NOME))
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
      If MsgYESNO("Confirma Exclusão do Item Contábil "+cItCTB+" do Cliente "+Alltrim(SA1->A1_NOME))
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
              {"CT1_CLASSE","2"       ,Nil},; //2 - Analítica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","11201001",Nil}}  //11201001 - Conta Fixa para Conta Superior de Clientes.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
   If lMsErroAuto
      Alert("Ocorreu Erro Fatal na Exclusão da Conta Contábil: "+Alltrim(cContaCli)+" no Plano de Contas, Favor reportar ao Depto. de TI")
   EndIf
Return