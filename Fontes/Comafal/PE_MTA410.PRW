#include "Protheus.Ch"
/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PE_MTA410 �Autor  �Five Solutions      � Data �  28/09/2010 ���
�������������������������������������������������������������������������͹��
���Desc.     � Ponto de Entrada Ap�s ter confirmado e validado todo o     ���
���          � Pedido de Venda.                                           ���
���          � Estamos utilizando para validar a comiss�o no momento da   ���
���          � grava��o do Pedido de Venda conforme determina             ���
���          � do na Tabela de Pre�o(DA0 - DA0_COMIS)                     ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL: PE,SP e RS - Pedido de Vendas(MATA410)            ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function MTA410 //Ponto de Entrada
   cAreaKi   := GetArea()
   lRetVldPV := .T. 
   If Alltrim(FunName()) == "MATA415" .Or. Alltrim(FunName()) == "MATA416"
      Processa({||fVldCPV()},"Validando Comiss�es ...")
      RestArea(cAreaKi)
      Return(lRetVldPV)
   EndIf
   If l410Auto
      RestArea(cAreaKi) 
      Return(lRetVldPV)
   EndIf
   If M->C5_TIPO == "N"
      If xFilial("SC5") == "06" 
         Processa({||fVldCPV()},"Validando Comiss�es ...")
      EndIf
   EndIf
   RestArea(cAreaKi)
Return(lRetVldPV)
User Function COMISTB //Gatilho no C5_CLIENTE, C5_LOJACLI e C5_TABELA.
   nComTB    := Posicione("DA0",1,xFilial("DA0")+M->C5_TABELA,"DA0_COMIS")
   If M->C5_TIPO == "N" 
      cAreaKi   := GetArea()
      nComTB    := 0
      lRetVldPV := .T. 
      Processa({||fVldCPV()},"Validando Comiss�es ...")
      RestArea(cAreaKi)
   EndIf
Return(nComTB)

Static Function fVldCPV
   cTabPrc:= M->C5_TABELA
   DbSelectArea("DA0")
   DbSetOrder(1)//DA0_FILIAL+DA0_CODTAB
   If DbSeek(xFilial("DA0")+cTabPrc)
      nComTB := DA0->DA0_COMIS
   Else
      Alert("A tabela de pre�o "+cTabPrc+" n�o foi localizada no banco de dados do sistema(Tabela DA0)")
      Return(.F.)
   EndIf
   
   M->C5_COMIS1 := nComTB
   /* A Comiss�o da COMAFAL � atribu�da no Cadastro do Vendedor portanto n�o ser� tratada por item.
   
   nReg := Len(aCols)
   ProcRegua(nReg)
   For nPV := 1 To nReg
      IncProc("Informando Comiss�es dos Itens "+Alltrim(Str(nPV))+" / "+Alltrim(Str(nReg)))
      GdFieldPut("C6_COMIS1",nComTB,nPV)
   Next nPV
   */
Return

/*****************
* Gatilho do Produto para informa��o do pre�o no Pedido de Vendas
* Criado o Gatilho C6_QTDVEN Sequencia 003 Contra Dominio C6_PRCVEN 
* Regra u_PrcTABCMF() Condi��o SM0->M0_CODFIL == "06"
*********************************************************************/
User Function PrcTABCMF
   cAreaKi := GetArea()
   nPrcTAB := GdFieldGet("C6_PRCVEN",n) //M->C6_PRCVEN
   If l410Auto
      RestArea(cAreaKi)
      Return(nPrcTAB)
   EndIf
   
   If M->C5_TIPO == "N"
   
      If !l410Auto //Se n�o for Rotina Automatica
         cProdTB := GdFieldGet("C6_PRODUTO",n) //M->C6_PRODUTO
         cTabPd  := M->C5_TABELA
         nQtdVen := M->C6_QTDVEN
         dEmiPV  := M->C5_EMISSAO //Five - 28/03/2011
         nAlqIPI := Posicione("SB1",1,xFilial("SB1")+cProdTB,"B1_IPI")
         nVlrFrt := GdFieldGet("C6_VLRKM",n)     //Valor do Frete informado no Pedido de Vendas
         nPrcTAB := 0
         cNmeProd:= Alltrim(Posicione("SB1",1,xFilial("SB1")+cProdTB,"B1_DESC"))
         If Empty(cTabPd)
            Alert("Informe primeiro a tabela de pre�o do pedido de vendas")
            Return(nPrcTAB)
         EndIf
         DbSelectArea("DA1")
         DbSetOrder(1) //DA1_FILIAL+DA1_CODTAB+DA1_CODPRO+DA1_INDLOT+DA1_ITEM
         If DbSeek(xFilial("DA1")+cTabPd+cProdTB)
            //nVlrIPI := If(nAlqIPI>0, (DA1->DA1_PRCVEN * (nAlqIPI/100)),0)
            //Five 28/03/2011 - nUmPrc  := DA1->DA1_PRCVEN  //Five Solutions Consultoria - 23/12/2010
            If dEmiPV < DA1->DA1_DTALTP
               nUmPrc := DA1->DA1_PRCANT
            Else
               nUmPrc := DA1->DA1_PRCVEN
            EndIf            
            nClcIPI := (nAlqIPI / 100) + 1
            //Five 28/03/2011 - nPreC6  := (DA1->DA1_PRCVEN / nClcIPI) - nVlrFrt
            nPreC6  := (nUmPrc / nClcIPI) - nVlrFrt
            //MsgInfo("nPreC6 "+Alltrim(Str(nPreC6)))
      
            
             cCndPV  := M->C5_CONDPAG
             nPercDsc:= fPrzMedio(cCndPV)
             If nPercDsc > 0
                //Five - 23/12/2010 - nVlDcPrz := nPreC6 * (nPercDsc/100)
                nVlDcPrz := nUmPrc * (nPercDsc/100)
                nPrcTMP := nPreC6
                nPreC6   := nPreC6 - nVlDcPrz
                //MsgInfo("Pre�o do Produto(S/IPI e S/Frete "+Transform(nPrcTMP,"@E 999,999,999.99")+" Percentual de Desconto a ser aplicado "+Transform((nPercDsc/100),"@E 999.99")+" Valor do Desconto "+Transform(nVlDcPrz,"@E 999,999,999.99")+" Valor Liquido "+Transform(nPreC6,"@E 999,999,999.99"))
             EndIf
            
      
            nPrcTAB := a410Arred(nPreC6,"C6_PRCVEN")
            nTotPrd := a410Arred((nPrcTAB * nQtdVen),"C6_VALOR")
            GdFieldPut("C6_VALOR",nTotPrd,n)
         Else
            Alert("O Produto "+cProdTB+" - "+cNmeProd+" n�o foi localizado na Tabela de Pre�o "+cTabPd+" utilizada no Pedido, este item n�o pode ser utilizado, pois seu pre�o permanecer� zerado")
            Return(nPrcTAB)
         EndIf
   
      EndIf
   
   EndIf
   
   RestArea(cAreaKi)
   //MsgInfo("nPrcTAB "+Alltrim(Str(nPrcTAB)))
Return(nPrcTAB)

/***************************
* Valida permi��o da edi��o do campo Pre�o de Vendas - C6_PRCVEN de acordo
* com a Tabela de Pre�o informada.
* Valida��o da Edi��o do Campo C6_PRCVEN(X3_WHEN) = IF(xFilial("SC6")=="06",U_VldPrecoTB(),.T.)
* Para esta funcionalidade foram criados os seguintes campos na tabela DA0-Cabe�alho da Tab.Pre�o:
*   DA0_EDTPRC - Edita Preco no Pedido de Vendas possibilitando op��es S=Sim;N=Nao
*   DA0_TIPALT - Altera Pre�o para Maior ou Menor possibilitando op��es M=Maior;B=Menor
***********************************************************************************************/
User Function VldPrecoTB    //X3_WHEN do Campo C6_PRCVEN
   If l410Auto
      Return(.T.)
   EndIf
   cTabPd   := M->C5_TABELA 
   If M->C5_TIPO <> "N"
      cTabPd   := M->C5_TABELA 
      Return(.T.)
   EndIf
Return( Posicione("DA0",1,xFilial("DA0")+cTabPd,"DA0_EDTPRC") == "S" )

/********************
* Valida��o do Pre�o digitado em rela��o a tabela de pre�o
* informada no pedido de vendas
* Valida��o de usu�rio do campo C6_PRCVEN
*************************************************************/
User Function VldTabxPDg

   If l410Auto
      Return(.T.)
   EndIf
   
   If M->C5_TIPO <> "N"
      Return(.T.)
   EndIf
   
   cTabPd   := M->C5_TABELA 
   dEmiPV   := M->C5_EMISSAO //Five - 17/02/2011
   
   cProdTB  := GdFieldGet("C6_PRODUTO",n) //M->C6_PRODUTO //
   nVlrFrt  := GdFieldGet("C6_VLRKM",n)     //Valor do Frete informado no Pedido de Vendas
   nAlqIPI  := Posicione("SB1",1,xFilial("SB1")+cProdTB,"B1_IPI") 
   cNmeProd := Alltrim(Posicione("SB1",1,xFilial("SB1")+cProdTB,"B1_DESC"))
   nClcIPI  := (nAlqIPI / 100) + 1
   cTabPd   := M->C5_TABELA
   If Empty(cTabPd) 
      Alert("Informe a Tabela de Pre�o")
      Return(.F.)
   EndIf
   
   nPrcTba  := Posicione("DA1",1,xFilial("DA1")+cTabPd+cProdTB,"DA1_PRCVEN")

   If Empty(nPrcTba)
      Alert("O Produto "+cProdTB+" - "+cNmeProd+" n�o foi localizado na Tabela de Pre�o "+cTabPd+" utilizada no Pedido, este item n�o pode ser utilizado, pois seu pre�o permanecer� zerado")
      Return(.F.)
   EndIf
   
   nPrcDig  := M->C6_PRCVEN + nVlrFrt //GdFieldGet("C6_PRCVEN")  
   nVlrIPI  := (nPrcDig) * (nAlqIPI/100)
   nPrcDig  := a410Arred(nPrcDig + nVlrIPI)
   
   //MsgInfo(" M->C6_PRCVEN "+Alltrim(Str(M->C6_PRCVEN))+" nVlrIPI "+Alltrim(Str(nVlrIPI))+" nPrcDig "+Alltrim(Str(nPrcDig)))
   cNmeProd := Alltrim(Posicione("SB1",1,xFilial("SB1")+cProdTB,"B1_DESC"))
   nPrcTAB  := 0
   lRetOkPrc:= .T.
   If Empty(cTabPd)
      Alert("Informe primeiro a tabela de pre�o do pedido de vendas")
      Return(.F.)
   EndIf
   DbSelectArea("DA1")
   DbSetOrder(1) //DA1_FILIAL+DA1_CODTAB+DA1_CODPRO+DA1_INDLOT+DA1_ITEM
   If DbSeek(xFilial("DA1")+cTabPd+cProdTB)
      If dEmiPV < DA1->DA1_DTALTP
         nPrcTAB := DA1->DA1_PRCANT
      Else
         nPrcTAB := DA1->DA1_PRCVEN
      EndIf
   Else
      Alert("O Produto "+cProdTB+" - "+cNmeProd+" n�o foi localizado na Tabela de Pre�o "+cTabPd+" utilizada no Pedido, este item n�o pode ser utilizado")
      Return(.F.)
   EndIf
   cPermic  := Posicione("DA0",1,xFilial("DA0")+cTabPd,"DA0_TIPALT") //M=Maior;B=Menor
   
   /*
   If nPrcDig > nPrcTAB .And. cPermic <> "M"
      Alert("A Tabela "+cTabPd+" n�o permite pre�o MAIOR que "+Transform(nPrcTAB,"@E 9,999,999.99")+" para o produto "+cNmeProd+" o pre�o digitado "+Transform(nPrcDig,"@E 9,999,999.99")+" n�o ser� permitido")
      Return(.F.)
   EndIf
   */
   
   If nPrcDig < nPrcTAB .And. cPermic <> "B"
      Alert("A Tabela "+cTabPd+" n�o permite pre�o MENOR que "+Transform(nPrcTAB,"@E 9,999,999.99")+" para o produto "+cNmeProd+" o pre�o digitado "+Transform(nPrcDig,"@E 9,999,999.99")+" n�o ser� permitido")
      Return(.F.)
   EndIf
   
   //If (nPrcDig > nPrcTAB .And. cPermic == "M") .Or. (nPrcDig < nPrcTAB .And. cPermic == "B")
   If (nPrcDig < nPrcTAB .And. cPermic == "B")
      lRetOkPrc := .F.
      cCodPsw := SPACE(10)
      cAprovad:= SPACE(2)
      Define MsDialog oPswPrc Title "Valida��o do Pre�o Informado" From 001,015 To 250,350 Of oMainWnd Pixel
         
         nAprov := 1

         //@ 000,000 BITMAP oBmp RESNAME "LOGIN" Of oPswPrc SIZE 500,400 NOBORDER WHEN .F. PIXEL  //PROJETOAP,LOGIN,PMSEDT4,PMSEDT3,PMSEDT1.PMSDOC,PMSTASK1 - 4,6,ROTEIRO,CLOCK01,ENGRENAGEM,INSTRUME,CLIPS,TOOLBAR,BMPEDT1...4,BMP_TASK1...4
         @ 010,015 SAY "Codigo de Autoriza��o    " Of oPswPrc PIXEL SIZE 90 ,10 //FONT oBold
         @ 010,100 MSGET oCodPsw VAR cCodPsw Of oPswPrc PIXEL SIZE 50 ,10 PASSWORD
         @ 030,015 SAY "Favor solicitar codigo de altera��o de pre�o " Of oPswPrc PIXEL SIZE 150 ,10
         @ 045,015 SAY "a um dos gestores informados abaixo: " Of oPswPrc PIXEL SIZE 150 ,10 //FONT oBold 
         @ 060,015 TO 100,150 LABEL "Selecione o Aprovador" Of oPswPrc PIXEL

         @ 075,020 RADIO oAprov VAR nAprov 3D SIZE 100,10 PROMPT "Patricia Amaral [Auditoria ] ","Roberto Lira    [Financeiro] " Of oPswPrc PIXEL
         /*
         @ 060,015 SAY "  * Patricia Amaral [Auditoria ] " Of oPswPrc PIXEL SIZE 150 ,10 //FONT oBold 
         @ 075,015 SAY "  * Roberto Lira    [Financeiro] " Of oPswPrc PIXEL SIZE 150 ,10 //FONT oBold 
         */
         @ 115,090 BUTTON oCnfPsw PROMPT "&Confirma" SIZE 70,10 OF oPswPrc PIXEL ACTION(fChkPsw()) WHEN .T.
         
      Activate MsDialog oPswPrc Center 
   EndIf
      
Return(lRetOkPrc)

Static Function fChkPsw()
   cCdoVld := ""
   If nAprov == 1
      cCdoVld := Alltrim(GetMV("MV_COMPSW"))
      cAprovad:= "PAMARAL"
   Else
      cCdoVld := Alltrim(GetMV("MV_COMPS2"))
      cAprovad:= "RLIRA"
   EndIf
   //MsgInfo("Alltrim(cCodPsw): "+Alltrim(cCodPsw)+ " Alltrim(cCdoVld): "+Alltrim(cCdoVld))
   If Alltrim(cCodPsw) == Alltrim(cCdoVld)
      //GdFieldPut("C6_PRCVEN",nPrcDig,n)
      GdFieldPut("C6_LIBPRC",cAprovad,n)
      lRetOkPrc:= .T.
      oPswPrc:End()
      Return(.T.)
   Else
      Alert("C�digo Inv�lido !")
      lRetOkPrc:= .F.
      Return(.F.)
   EndIf
Return

/*
 Criado Valida��o de Usu�rio(X3_VLDUSER) no Campo C5_TABELA para validar altera��o da tabela de pre�o 
 ap�s j� ter sido digitado produtos no pedido de vendas.
 REGRA DA VALIDA��O: If(SM0->M0_CODFIL == "06" , u_AltTBPrc(), .T.)
*/
User Function AltTBPrc
   If l410Auto
      Return(.T.)
   EndIf
   If M->C5_TIPO <> "N"
      Return(.T.)
   EndIf
   If Len(aCols) > 0 .Or. !Empty(GdFieldGet("C6_PRODUTO",n))
      If MsgYesNo("A Altera��o da Tabela de Pre�o ap�s digita��o de produto implicar� na redigita��o das quantidades dos produtos para obter novamente os pre�o, CONFIRMA ESTA OPERA��O ?")
         For nTb := 1 To Len(aCols)
            //GdFieldPut("C6_PRODUTO","",nTb)
            GdFieldPut("C6_PRCVEN",0,nTb)
            GdFieldPut("C6_QTDVEN",0,nTb)
         Next nTb
         MsgInfo("Digite novamente as quantidades dos produtos para obter os pre�os da tabela")
         Return(.T.)
      Else
         Return(.F.)
      EndIf
   EndIf
Return(.T.)

User Function GrvCodPsw
   
   If Alltrim(cUserName) == "PAMARAL"
      cCodAnt := Alltrim(GetMV("MV_COMPSW"))
   Else
      cCodAnt := Alltrim(GetMV("MV_COMPS2"))
   EndIf
   
   cCodPsw := SPACE(10)
   Define MsDialog oPswTel Title "Codigo de Permiss�o p/Alterar Pre�o" From 001,015 To 150,350 Of oMainWnd Pixel
         
      @ 010,015 SAY "Codigo Anterior          " Of oPswTel PIXEL SIZE 90 ,10 //FONT oBold
      @ 010,100 MSGET oCodPsw VAR cCodPsw Of oPswTel PIXEL SIZE 50 ,10 PASSWORD 
      @ 025,100 BUTTON oCnfPsw PROMPT "&Ver Anterior" SIZE 50,10 OF oPswTel PIXEL ACTION(MsgInfo("Codigo Anterior: "+cCodAnt)) WHEN .T.
      @ 040,100 BUTTON oCnfPsw PROMPT "&Confirma" SIZE 50,10 OF oPswTel PIXEL ACTION(TstPsw()) WHEN .T.
         
   Activate MsDialog oPswTel Center 
   
Return
Static Function TstPsw
   If cCodPsw <> cCodAnt
      Alert("C�digo Inv�lido !")
      lRetOkPrc:= .F.
      Return(.F.)
   Else
      oPswTel:End()
      cNewPsw := SPACE(10)
      Define MsDialog oNewTel Title "Novo Codigo p/Alterar Pre�o" From 001,015 To 150,350 Of oMainWnd Pixel
         
         @ 010,015 SAY "Codigo Novo              " Of oNewTel PIXEL SIZE 90 ,10 //FONT oBold
         @ 010,100 MSGET oNewPsw VAR cNewPsw Of oNewTel PIXEL SIZE 50 ,10 PASSWORD 
         @ 025,100 BUTTON oCnfNew PROMPT "&Confirma" SIZE 50,10 OF oNewTel PIXEL ACTION(GrvPsw()) WHEN .T.
         
      Activate MsDialog oNewTel Center 

   EndIf
Return
Static Function GrvPsw()
   If Alltrim(cUserName) == "PAMARAL"
      PutMV("MV_COMPSW",cNewPsw)
   Else
      PutMV("MV_COMPS2",cNewPsw)
   EndIf
   MsgInfo("O Novo codigo para altera��o de pre�os foi gravado com sucesso!")
   oNewTel:End()
Return(.T.)

Static Function fPrzMedio(xCondPgt)

   
   Private nAutoriz  := GetMV("MV_PRZMED")
   Private cCondicao := Posicione("SE4",1,xFilial("SE4")+xCondPgt,"E4_COND") //cCndPV //Alltrim(xCondPgt) //Alltrim(M->E4_COND)
   Private cCondClc  := STRTRAN(cCondicao,",","+")
   Private nQtdPc    := RunPc() //(cCondicao)
   Private nPrzTot   := &cCondClc
   Private nMedPrazo := (nPrzTot / nQtdPc)
   Private nDescCMF  := (2 / nAutoriz)
   Private nDescPmd  := 0
   If nMedPrazo < nAutoriz
      nDiasPrz := nAutoriz - nMedPrazo
      nDescPmd := a410Arred(nDescCMF * nDiasPrz)
      //MsgInfo("Prazo Medio Limite Autorizado pela Diretoria "+Alltrim(Str(nAutoriz))+" Prazo Medio da Condicao "+cCondicao+" = "+Alltrim(Str(nMedPrazo))+" Dias de Diferen�a "+Alltrim(Str(nDiasPrz)))
   EndIf

Return(nDescPmd)

Static Function RunPc//(cCondicao)
   
   nQuantPC := 1//J� iniciamos considerando o m�nimo de uma parcela, ent�o se existir a primeira v�rgula
                //isto indica que temos duas parcelas, e se for encontrada uma pr�xima v�rgula temos tr�s
                // e assim por diante contaremos o n�mero de parcelas de acordo com o n�mero de v�rgulas
                // encontradas.
   nIncio  := 1
   For nPc := 1 To Len(cCondicao)
      cTexto := Substr(IIF(nPC==1,cCondicao,cTexto),nIncio,Len(cCondicao)) 
      //Alert("cTexto "+cTexto+" nIncio "+Alltrim(Str(nIncio)))
      nPosVirg := At(",",cTexto)
      //Alert("nPosVirg "+Alltrim(Str(nPosVirg))+" nPC "+Alltrim(Str(nPC))+" nQuantPC "+Alltrim(Str(nQuantPC)))
      nIncio  := nPosVirg + 1
      If nPosVirg > 0
         nQuantPC ++
      EndIf
   Next nPC
   
Return(nQuantPC)