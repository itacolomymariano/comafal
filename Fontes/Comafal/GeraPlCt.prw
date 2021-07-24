#INCLUDE "TopConn.Ch"
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �GERAPLCT  �Autor  �Five Solutions      � Data �  06/06/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � Programa para execu��o �nica com o objetivo de gerar plano ���
���          � de contas com base nas contas cont�beis informadas nos     ���
���          � cadastros de fornecedores e clientes.                      ���
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
User Function GeraPlCt
   Processa({||RunPlCt()},"Gerando Plano de Contas, Aguarde ...")
   Alert("Processo Concluido com Exito !!!")
Return
Static Function RunPlCt
   
   For nPl := 1 To 2
      If nPl == 1
         cQryPl := "SELECT COUNT(*) AS TOTREG"
      Else
         cQryPl := "SELECT SA1.A1_NOME,SA1.A1_CONTA"
      EndIf
      cQryPl += " FROM "+RetSQLName("SA1")+" SA1"
      cQryPl += " WHERE SUBSTRING(SA1.A1_CONTA,1,3) <> '   '"
      cQryPl += "   AND SA1.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryPl NEW ALIAS "TSA1"
      
      If nPl == 1
         DbSelectArea("TSA1")
         nReg := TSA1->TOTREG
         DbCloseArea()
      EndIf
   Next nPl
   ProcRegua(nReg)
   nCont := 1
   DbSelectArea("TSA1")
   While !Eof()
      IncProc("Gerando Contas de Clientes ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
      DbSelectArea("CT1")
      DbSetOrder(1)//CT1_FILIAL+CT1_CONTA
      If !DbSeek(xFilial("CT1")+ALLTRIM(TSA1->A1_CONTA))
         IncSA1()
      EndIf
      DbSelectArea("TSA1")
      DbSkip()
      nCont ++
   EndDo
   DbSelectArea("TSA1")
   DbCloseArea()

   For nPl := 1 To 2
      If nPl == 1
         cQryPl := "SELECT COUNT(*) AS TOTREG"
      Else
         cQryPl := "SELECT SA2.A2_NOME,SA2.A2_CONTA"
      EndIf
      cQryPl += " FROM "+RetSQLName("SA2")+" SA2"
      cQryPl += " WHERE SUBSTRING(SA2.A2_CONTA,1,3) <> '   '"
      cQryPl += "   AND SA2.D_E_L_E_T_ <> '*'"
      
      TCQuery cQryPl NEW ALIAS "TSA2"
      
      If nPl == 1
         DbSelectArea("TSA2")
         nReg := TSA2->TOTREG
         DbCloseArea()
      EndIf
   Next nPl
   ProcRegua(nReg)
   nCont := 1
   DbSelectArea("TSA2")
   While !Eof()
      IncProc("Gerando Contas de Fornecedores ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg)))
      DbSelectArea("CT1")
      DbSetOrder(1)//CT1_FILIAL+CT1_CONTA
      If !DbSeek(xFilial("CT1")+ALLTRIM(TSA2->A2_CONTA))
         IncSA2()
      EndIf
      DbSelectArea("TSA2")
      DbSkip()
      nCont ++
   EndDo
   DbSelectArea("TSA2")
   DbCloseArea()
   
Return

Static Function IncSA1()
   cContaCli := TSA1->A1_CONTA
   cDescCta  := Alltrim(TSA1->A1_NOME)
   nOpc   := 3 
   aConta := {{"CT1_CONTA",cContaCli  ,Nil},;
              {"CT1_DESC01",cDescCta  ,Nil},;
              {"CT1_CLASSE","2"       ,Nil},; //2 - Anal�tica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","11201001",Nil}}  //11201001 - Conta Fixa para Conta Superior de Clientes.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
Return

Static Function IncSA2()
   cContaFor := TSA2->A2_CONTA
   cDescCta  := Alltrim(TSA2->A2_NOME)
   nOpc   := 3 
   aConta := {{"CT1_CONTA",cContaFor  ,Nil},;
              {"CT1_DESC01",cDescCta  ,Nil},;
              {"CT1_CLASSE","2"       ,Nil},; //2 - Anal�tica
              {"CT1_NORMAL","1"       ,Nil},; //1 - Devedora
              {"CT1_CTASUP","21102",Nil}}  //21102 - Alterado em 11/06/2009 - Antes, 21102001 - Conta Fixa para Conta Superior de Fornecedores.
   MsExecAuto({|x,y| CTBA020(x,y)},aConta,nOpc)
Return