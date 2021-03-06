#Include "Protheus.ch"
#Include "TopConn.Ch"
#Include "Font.Ch"
#Include "Colors.Ch"
#Include "COLOR.CH"

/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?MANTABPRC ?Autor  ?Five Solutions      ? Data ?  27/12/2010 ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? Tela de Manuten??o das Tabelas de Pre?o por Grupo/Subgrupo ???
???          ?                                                            ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL - PE, SP e RS - Faturamento.                       ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function ManTabPrc

   Private xPerg     := "PRGMTP"
   Private aProdTabs := {}
   Private nReajMTP  := 0
   Private oProds
   Private oFlagAmar   := LoadBitmap( GetResources(), "BR_AMARELO")
   Private oFlagAzul   := LoadBitmap( GetResources(), "BR_AZUL")
   Private oFlagMarr   := LoadBitmap( GetResources(), "BR_MARRON")
   Private oNovoPrc 
   Private nNovoPrc    := 0
   Private aTbMatriz   := {}

   Private cCpoPsq := Space(100)
   Private cIndice := Space(10)
   Private aIndice := {"Codigo","Descricao"}
   Private oCombo  := Nil
   Private dDataalt:= Date()//Por Helton
 
   MtPrgMTP()

   If Pergunte(xPerg,.T.)
      cGrpPrd   := MV_PAR01
      cSbGrpPrd := MV_PAR02
      cSbGrpAte := MV_PAR03
      nReajMTP  := MV_PAR04
      cProdIni  := MV_PAR05
      cProdFim  := MV_PAR06
      cDelProd  := Alltrim(MV_PAR07)
      cTabIni   := MV_PAR08
      cTabFim   := MV_PAR09
      cDelTab   := Alltrim(MV_PAR10)
      cBitola   := MV_PAR11
      cTabMat   := MV_PAR12
      lTabUm    := If(MV_PAR13==1,.T.,.F.)
      nPrcApl   := MV_PAR14
      cFilMed   := Alltrim(MV_PAR15)
      lJurSimp  := If(MV_PAR16==1,.T.,.F.)
      //lSoTbVen  := If(MV_PAR17==1,.T.,.F.)
      nNovoPrc  := nPrcApl
      
      cDesGrp   := If(!Empty(cGrpPrd),Alltrim(Posicione("SBM",1,xFilial("SBM")+cGrpPrd,"BM_DESC")),"Todos")
      cSGPrd    := Alltrim(Posicione("SZY",1,xFilial("SZY")+cSbGrpPrd,"ZY_DESC"))
      cSGAtPrd  := Alltrim(Posicione("SZY",1,xFilial("SZY")+cSbGrpAte,"ZY_DESC"))
      
      Processa({||fQryMTP()},"Selecionando Produtos ... ")
   Else
      Return
   EndIf
   
   DEFINE FONT oFntGrande NAME "Arial" SIZE 20,22 BOLD
   DEFINE FONT oFntMedia NAME "Arial" SIZE 14,16 BOLD
   DEFINE FONT oFntPeq NAME "Arial" SIZE 10,12 BOLD
   oFont1 := TFont():New("Arial",0,16,,.F.,,,,.T.,.T.,.F.) //Fonte ?talico Sublinhado
   
   Define MsDialog oDlgMTP Title "Manuten??o da Tabela de Pre?o" From 001,005 To 630,980 Of oMainWnd Pixel
   
   @010,005 SAY "Grupo    " SIZE 100,10 FONT oFntMedia COLOR CLR_HBLUE Of oDlgMTP PIXEL 
   @010,115 SAY cGrpPrd+" "+cDesGrp   SIZE 200,10 FONT oFntMedia Of oDlgMTP PIXEL 
      
   @025,005 SAY "Subgrupo " SIZE 100,10 FONT oFntMedia COLOR CLR_HBLUE Of oDlgMTP PIXEL 
/*
      cSbGrpPrd := MV_PAR02
      cSbGrpAte := MV_PAR03
      cSGPrd    := Alltrim(Posicione("SZY",1,xFilial("SZY")+cSbGrpPrd,"ZY_DESC"))
      cSGAtPrd  := Alltrim(Posicione("SZY",1,xFilial("SZY")+cSbGrpAte,"ZY_DESC"))
*/   
   If Alltrim(cSGPrd) == Alltrim(cSGAtPrd)
      @025,115 SAY cSbGrpPrd+" "+cSGPrd SIZE 200,10 FONT oFntMedia Of oDlgMTP PIXEL
   Else
      @025,115 SAY "de "+cSbGrpPrd+" "+cSGPrd+" At? "+cSbGrpAte+" "+cSGAtPrd SIZE 300,10 FONT oFntMedia Of oDlgMTP PIXEL 
   EndIf
   
   If lTabUm
      @040,005 SAY "Novo Pre?o " SIZE 100,10 FONT oFntPeq Of oDlgMTP PIXEL 
      @040,115 MSGET oNovoPrc VAR nNovoPrc SIZE 50,10 Picture "@E 999,999,999.99" WHEN .F. FONT oFntPeq Of oDlgMTP PIXEL 
   Else
      @040,005 SAY "Reajuste(%)" SIZE 100,10 FONT oFntPeq Of oDlgMTP PIXEL 
      //@040,115 MSGET oReajMTP VAR nReajMTP SIZE 50,10 Picture "@E 999.99" VALID(Processa({||fRjtTAB()})) FONT oFntPeq Of oDlgMTP PIXEL 
      @040,115 MSGET oReajMTP VAR nReajMTP SIZE 50,10 Picture "@E 999.999" WHEN .F. FONT oFntPeq Of oDlgMTP PIXEL 
   EndIf

   @ 040,200 SAY "Pesquisar"       SIZE 065,008 FONT oFntPeq PIXEL OF oDlgMTP
   @ 039,250 COMBOBOX oCombo  VAR cIndice ITEMS aIndice  SIZE 070,008    PIXEL OF oDlgMTP VALID fPsqProd(cIndice,aProdTabs,oProds)
   @ 039,320 MSGET oCpoPsq VAR cCpoPsq  PICTURE "@!" WHEN .T. SIZE 120,008    PIXEL OF oDlgMTP
   @ 039,450 BUTTON "Buscar"             SIZE 028,010 PIXEL OF oDlgMTP ACTION fBuscPrd(aProdTabs,cCpoPsq,cIndice,oProds)
     

   @ 055,007 LISTBOX oProds VAR cVarProds Fields HEADER "F","Produto","Descri??o","Tabela","Pre?o Atual","Pre?o Ajustado","Prc Tab Matriz" SIZE 470,220 NOSCROLL OF oDlgMTP PIXEL
   //@ 290,007 SAY "click duplo sobre no produto diferenciado ajusta pre?o manualmente" FONT oFont1 COLOR CLR_HBLUE OF oDlgMTP PIXEL
      
   oProds:SetArray(aProdTabs)
   oProds:bLine := { || { oFlagMarr,aProdTabs[oProds:nAt][1], aProdTabs[oProds:nAt,2],;
   aProdTabs[oProds:nAt,3],Transform(aProdTabs[oProds:nAt,4],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,5],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,7],"@E 999,999,999.99")}}
   /*
   oProds:blDblClick := { || RunAjtMan( aProdTabs[oProds:nAt][1], aProdTabs[oProds:nAt,2],;
                                        aProdTabs[oProds:nAt,3],aProdTabs[oProds:nAt,4],aProdTabs[oProds:nAt,5])}
   */                                     
   oCpoPsq:SetFocus()
   
   @ 290,320 BUTTON OemToAnsi("&A t u a l i z a   T a b e l a s") SIZE 150,12 OF oDlgMTP PIXEL ACTION Processa({|| fAtMTP()},"Atualizando Tabelas de Pre?os ...") WHEN .T.
   
   Activate MsDialog oDlgMTP Center
   
Return

Static function fPsqProd(cIndice,aProdTabs,oProds)

   Do Case
      Case cIndice=="Codigo"
           aSort(aProdTabs, ,,{|x,y| x[1]<y[1]}) 
      Case cIndice=="Descricao"
           aSort(aProdTabs, ,,{|x,y| x[2]<y[2]}) 
   End Do
    
   oProds:Refresh( )
   oProds:SetArray(aProdTabs)
   oProds:bLine := { || { oFlagMarr,aProdTabs[oProds:nAt][1], aProdTabs[oProds:nAt,2],;
   aProdTabs[oProds:nAt,3],Transform(aProdTabs[oProds:nAt,4],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,5],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,7],"@E 999,999,999.99")}}
  
Return()

// Fun??o que realiza a pesquisa 
Static Function fBuscPrd(aProdTabs,cCpoPsq,cIndice,oProds)
 
   Do Case 
      Case cIndice=="Codigo"
           If !Empty(cCpoPsq)
              nTamPsq := Len(Alltrim(cCpoPsq))
              If (nA := Ascan(aProdTabs, {|x| Substr(x[1],1,nTamPsq) $ AllTrim(cCpoPsq) })) > 0
                 oProds:nAt := nA
              Else
                 Alert("O Codigo "+cCpoPsq+" n?o foi localizado em nehuma das tabelas de pre?os")
              EndIf
           Else
              MsgInfo("Informe o codigo do produto.")
           EndIf  
      Case cIndice=="Descricao"
           If !Empty(cCpoPsq)
              nTamPsq := Len(Alltrim(cCpoPsq))
              If (nA := Ascan(aProdTabs, {|x| Substr(x[2],1,nTamPsq) $ AllTrim(cCpoPsq) })) > 0
                 oProds:nAt := nA
              Else
                 Alert("A descri??o "+cCpoPsq+" n?o foi localizada em nehuma das tabelas de pre?os")
              EndIf
           Else
              MsgInfo("Informe a descri??o do produto.")
           EndIf
   End Do
 
   oProds:Refresh( )
   oProds:SetArray(aProdTabs)
   oProds:bLine := { || { oFlagMarr,aProdTabs[oProds:nAt][1], aProdTabs[oProds:nAt,2],;
   aProdTabs[oProds:nAt,3],Transform(aProdTabs[oProds:nAt,4],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,5],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,7],"@E 999,999,999.99")}}
   oProds:SetFocus()
   
Return(.T.)


Static Function fAtMTP()

   nRegATP := Len(aProdTabs)
   ProcRegua(nRegATP)
   For nGrv := 1 To Len(aProdTabs)
      IncProc("Gravando novos pre?os reajustados ... "+Alltrim(Str(nGrv))+" / "+Alltrim(Str(nGrv)))
      nGoDA1 := aProdTabs[nGrv,6]
      DbSelectArea("DA1")
      DbGoTo(nGoDA1)
      RecLock("DA1",.F.)
         DA1->DA1_PRCANT := aProdTabs[nGrv,4] //Por Helton
         DA1->DA1_PRCVEN := aProdTabs[nGrv,5]
         DA1->DA1_DTALTP := dDataalt //Por Helton
      MsUnLock()
   Next nGrv
   MsgInfo("Processo Conclu?do, "+Alltrim(Str(nGrv))+" Pre?os ajustados")
   oDlgMTP:End()
Return

Static Function fRjtTAB()

   If MsgYesNo("Deseja Reajustar "+cDesGrp+" / "+cSGPrd+" em "+Transform(nReajMTP,"@E 999,99")+"% ?")
      nReg := Len(aProdTabs)
      For nRjt := 1 To Len(aProdTabs)
         
         IncProc("Reajustando Produtos ... "+Alltrim(Str(nRjt))+" / "+Alltrim(Str(nReg)))
         aProdTabs[nRjt,5] := A410Arred( ((aProdTabs[nRjt,4] * (nReajMTP)) + aProdTabs[nRjt,4]) )

      Next nRjt
      oProds:SetArray(aProdTabs)
      oProds:bLine := { || { oFlagMarr,aProdTabs[oProds:nAt][1], aProdTabs[oProds:nAt,2],;
      aProdTabs[oProds:nAt,3],Transform(aProdTabs[oProds:nAt,4],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,5],"@E 999,999,999.99"),Transform(aProdTabs[oProds:nAt,7],"@E 999,999,999.99")}}
      oProds:Refresh()
      
   EndIf
   
Return(.T.)

Static Function RunAjtMan
   MsgInfo("Ajuste manual de pre?os")
Return

Static Function fQryMTP()
   
   aProdTabs := {}
   //MsgInfo(" nReajMTP "+Alltrim(Str(nReajMTP))+" (nReajMTP/100) "+Transform((nReajMTP/100),"@E 999.999"))
   For nMTP := 1 To 2
      
      If nMTP == 1
         cQryMTP := " SELECT COUNT(*) REGMTP "
      Else
         cQryMTP := " SELECT DA1.DA1_CODPRO,SB1.B1_DESC,DA1.DA1_CODTAB,DA0.DA0_DESCRI,DA1.DA1_PRCVEN, DA1.R_E_C_N_O_ DA1REC "
         //cQryMTP += "        (DA1.DA1_PRCVEN * "+StrTran(Transform((nReajMTP/100),"@E 999.999"),",",".")+") + DA1.DA1_PRCVEN PRCREAJ, DA1.R_E_C_N_O_ DA1REC "
      EndIf
      
      cQryMTP += "   FROM "+RetSQLName("DA1")+" DA1,"+RetSQLName("DA0")+" DA0,"+RetSQLName("SB1")+" SB1 "
      cQryMTP += "  WHERE DA0.DA0_FILIAL = '"+xFilial("DA0")+"'"
      cQryMTP += "    AND DA1.DA1_FILIAL = '"+xFilial("DA1")+"'"
      cQryMTP += "    AND SB1.B1_FILIAL = '"+xFilial("SB1")+"'"
   
      cQryMTP += "    AND DA0.DA0_CODTAB = DA1.DA1_CODTAB "
      cQryMTP += "    AND DA1.DA1_CODPRO = SB1.B1_COD "

      cQryMTP += "    AND SB1.B1_COD BETWEEN '"+cProdIni+"' AND '"+cProdFim+"'"
      If !Empty(cDelProd)
         cQryMTP += "    AND SB1.B1_COD NOT IN "+FormatIn(cDelProd,",")
      EndIf

      If lTabUm
         If !Empty(cGrpPrd)
            cQryMTP += "    AND SB1.B1_GRUPO = '"+cGrpPrd+"'"
         EndIf
         If !Empty(cSbGrpPrd)
            //cQryMTP += "    AND SB1.B1_SUBGPO = '"+cSbGrpPrd+"'"
            cQryMTP += "    AND SB1.B1_SUBGPO BETWEEN '"+cSbGrpPrd+"' AND '"+cSbGrpAte+"'"
         EndIf
         cQryMTP += "    AND DA0.DA0_CODTAB = '"+cTabMat+"'"
      Else
         cQryMTP += "    AND DA0.DA0_CODTAB BETWEEN '"+cTabIni+"' AND '"+cTabFim+"'"
         If !Empty(cDelTab)
            cQryMTP += "    AND DA0.DA0_CODTAB NOT IN "+FormatIn(cDelTab,",")
         EndIf
         cQryMTP += "    AND DA1.DA1_CODTAB NOT IN ('"+cTabMat+"')"
         
         /*
         If lSoTbVen
            If SM0->M0_CODFIL == "02"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,1) = 'S'"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,3) NOT IN ('S00','S01')"
            ElseIf SM0->M0_CODFIL == "03"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,1) = 'R'"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,3) NOT IN ('R00','R01')"
            ElseIf SM0->M0_CODFIL == "06"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,1) = 'P'"
               cQryMTP += "    AND SUBSTRING(DA1.DA1_CODTAB,1,3) NOT IN ('P00','P01')"
            EndIf
         EndIf
         */
      EndIf
      
      If !Empty(cFilMed)
         cQryMTP += "    AND SUBSTRING(SB1.B1_COD,7,4) = '"+cFilMed+"'"
      EndIf
      
      If !Empty(cBitola)
         cBitola := If(Val(Alltrim(Substr(cBitola,1,1)))>0,Substr(cBitola,1,2),Substr(cBitola,2,2))
         cQryMTP += "    AND SUBSTRING(SB1.B1_COD,11,2) = '"+cBitola+"'"
      EndIf
   
      cQryMTP += "    AND DA0.D_E_L_E_T_ <> '*' "
      cQryMTP += "    AND DA1.D_E_L_E_T_ <> '*' "
      cQryMTP += "    AND SB1.D_E_L_E_T_ <> '*' "
      
      If nMTP > 1
         cQryMTP += " ORDER BY DA1.DA1_CODTAB,DA1.DA1_CODPRO,SB1.B1_DESC "
      EndIf
      
      MemoWrite("ManutencaoTabPreco.SQL",cQryMTP)
      MemoWrite("C:\TEMP\ManutencaoTabPreco.SQL",cQryMTP)
      
      TCQuery cQryMTP NEW ALIAS "YMTP"
      
      If nMTP == 1
         DbSelectArea("YMTP")
         nRegYMTP := YMTP->REGMTP
         DbCloseArea()
      Else
         TCSetField("YMTP","DA1_PRCVEN","N",17,02)
         TCSetField("YMTP","PRCREAJ","N",17,02)
         TCSetField("YMTP","DA1REC","N",10,00)
      EndIf
   
   Next nMTP
   
   If !lTabUm
      
      For nMt := 1 To 2
      
         If nMt == 1
            cQryTMT := " SELECT COUNT(*) REGMAT "
         Else
            cQryTMT := " SELECT DA1.DA1_CODTAB,DA1.DA1_CODPRO,DA1_PRCVEN "
         EndIf
         
         cQryTMT += "   FROM "+RetSQLName("DA1")+" DA1 "
         cQryTMT += "  WHERE DA1.DA1_FILIAL = '"+xFilial("DA1")+"'"
         cQryTMT += "    AND DA1.DA1_CODTAB = '"+cTabMat+"'"
         cQryTMT += "    AND DA1.D_E_L_E_T_ <> '*'"
      
         MemoWrite("TabelaMatriz.SQL",cQryTMT)
      
         TCQuery cQryTMT NEW ALIAS "XMAT"
         
         If nMt == 1
            DbSelectArea("XMAT")
            nReg := XMAT->REGMAT
            DbCloseArea()
         EndIf
      
      Next nMt
      
      aTbMatriz := {}
      ProcRegua(nReg)
      DbSelectArea("XMAT")
      nCont := 1
      While !Eof()
         IncProc("Carregando Tabela Matriz ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nReg))) 
         aAdd(aTbMatriz,{XMAT->DA1_CODTAB,XMAT->DA1_CODPRO,XMAT->DA1_PRCVEN})
         DbSelectArea("XMAT")
         DbSkip()
         nCont ++
      EndDo
      DbSelectArea("XMAT")
      DbCloseArea()
      
   EndIf
   
   DbSelectArea("YMTP")
   ProcRegua(nRegYMTP)
   nCont := 1
   While YMTP->(!Eof())
      //"F","Produto","Descri??o","Tabela","Pre?o Atual","Pre?o Ajustado"
      IncProc("Calculando Reajuste ... "+Alltrim(Str(nCont))+" / "+Alltrim(Str(nRegYMTP)))
      nPrcMtz := 0
      If lTabUm
         nPrecoApl := nPrcApl
      Else
         If lJurSimp
            nPosProd := aScan(aTbMatriz, {|x| x[1]+x[2] == cTabMat+YMTP->DA1_CODPRO})
            nPrecoApl := 0
            If nPosProd > 0
               nJrApl := A410Arred((nReajMTP)) * (Val(Substr(YMTP->DA1_CODTAB,2,2)) - 1)
               nPrecoApl := A410Arred( ((aTbMatriz[nPosProd,3] * (nJrApl)) + aTbMatriz[nPosProd,3]) )
               nPrcMtz   := aTbMatriz[nPosProd,3]
            Else
               cDescPrd := Alltrim(Posicione("SB1",1,xFilial("SB1")+YMTP->DA1_CODPRO,"B1_DESC"))
               If MsgYesNo("O Produto "+YMTP->DA1_CODPRO+" - "+cDescPrd+" n?o foi localiado na Tabela de Pre?o Matriz "+cTabMat+" o pre?o deste produto n?o ser? atualizado na tabela "+YMTP->DA1_CODTAB+" deseja criar o produto na Tabela Matriz com pre?o zerado ?")
                  cQryDA1 := " SELECT (MAX(DA1_ITEM)+1) NUMIT FROM "+RetSQLName("DA1")+" WHERE DA1_CODTAB = '"+cTabMat+"' AND D_E_L_E_T_ <> '*'"
                  TCQuery cQryDA1 NEW ALIAS "XDA1"
                  DbSelectArea("XDA1")
                     cItemDA1 := StrZero(XDA1->NUMIT,4)
                  DbCloseArea()
                  DbSelectArea("DA1")
                  RecLock("DA1",.T.)
                     DA1->DA1_FILIAL := xFilial("DA1")
                     DA1->DA1_ITEM   := cItemDA1
                     DA1->DA1_CODTAB := cTabMat
                     DA1->DA1_CODPRO := YMTP->DA1_CODPRO
                     DA1->DA1_PRCVEN := 0
                     DA1->DA1_ATIVO  := "1"
                     DA1->DA1_TPOPER := "4"             
                     DA1->DA1_MOEDA  := 1
                     DA1->DA1_QTDLOT := 999
                  MsUnLock()
                  DbSelectArea("YMTP")
               Else
                  If MsgYesNo("Deseja EXCLUIR o Produto "+YMTP->DA1_CODPRO+" - "+cDescPrd+" da Tabela de Pre?o "+YMTP->DA1_CODTAB)
                     DbSelectArea("DA1")
                     nGoDA1 := YMTP->DA1REC
                     RecLock("DA1",.F.)
                        DELETE
                     MsUnLock()
                     DbSelectArea("YMTP")
                  EndIf
               EndIf
            EndIf
            //aAdd(aTbMatriz,{XMAT->DA1_CODTAB,XMAT->DA1_CODPRO,XMAT->DA1_PRCVEN})
            //aProdTabs[nRjt,5] := A410Arred( ((aProdTabs[nRjt,4] * (nReajMTP/100)) + aProdTabs[nRjt,4]) )
            //nPrecoApl := YMTP->PRCREAJ
         Else
            nPosProd := aScan(aTbMatriz, {|x| x[1]+x[2] == cTabMat+YMTP->DA1_CODPRO})
            nPrecoApl := 0
            If nPosProd > 0
               nJrApl := nReajMTP
               nPrecoApl := A410Arred( ((aTbMatriz[nPosProd,3] * (nJrApl)) + aTbMatriz[nPosProd,3]) )
               nPrcMtz   := aTbMatriz[nPosProd,3]
            Else
               cDescPrd := Alltrim(Posicione("SB1",1,xFilial("SB1")+YMTP->DA1_CODPRO,"B1_DESC"))
               If MsgYesNo("O Produto "+YMTP->DA1_CODPRO+" - "+cDescPrd+" n?o foi localiado na Tabela de Pre?o Matriz "+cTabMat+" o pre?o deste produto n?o ser? atualizado na tabela "+YMTP->DA1_CODTAB+" deseja criar o produto na Tabela Matriz com pre?o zerado ?")
                  cQryDA1 := " SELECT (MAX(DA1_ITEM)+1) NUMIT FROM "+RetSQLName("DA1")+" WHERE DA1_CODTAB = '"+cTabMat+"' AND D_E_L_E_T_ <> '*'"
                  TCQuery cQryDA1 NEW ALIAS "XDA1"
                  DbSelectArea("XDA1")
                     cItemDA1 := StrZero(XDA1->NUMIT,4)
                  DbCloseArea()
                  DbSelectArea("DA1")
                  RecLock("DA1",.T.)
                     DA1->DA1_FILIAL := xFilial("DA1")
                     DA1->DA1_ITEM   := cItemDA1
                     DA1->DA1_CODTAB := cTabMat
                     DA1->DA1_CODPRO := YMTP->DA1_CODPRO
                     DA1->DA1_PRCVEN := 0
                     DA1->DA1_ATIVO  := "1"
                     DA1->DA1_TPOPER := "4"             
                     DA1->DA1_MOEDA  := 1
                     DA1->DA1_QTDLOT := 999
                  MsUnLock()
                  DbSelectArea("YMTP")
               Else
                  If MsgYesNo("Deseja EXCLUIR o Produto "+YMTP->DA1_CODPRO+" - "+cDescPrd+" da Tabela de Pre?o "+YMTP->DA1_CODTAB)
                     DbSelectArea("DA1")
                     nGoDA1 := YMTP->DA1REC
                     RecLock("DA1",.F.)
                        DELETE
                     MsUnLock()
                     DbSelectArea("YMTP")
                  EndIf               
               EndIf
            EndIf
            nNumTb := Val(Substr(YMTP->DA1_CODTAB,2,2))
            nCJr := 2
            nPrecoTMP := nPrcMtz 
            While nCJr <= nNumTb
               nJrApl := nReajMTP
               //MsgInfo("nPrecoApl "+Transform(nPrecoApl,"@E 999,999,999.99"))
               nPrecoTMP := A410Arred( (nPrecoTMP * nJrApl) + nPrecoTMP )
               nPrecoApl := nPrecoTMP
               //MsgInfo("YMTP->DA1_CODTAB: "+YMTP->DA1_CODTAB+" YMTP->DA1_CODPRO "+YMTP->DA1_CODPRO+" nCJr "+Alltrim(Str(nCJr))+" nPrecoApl "+Transform(nPrecoApl,"@E 999,999,999.99")+" nJrApl "+Alltrim(Str(nJrApl)))
               nCJr ++
            EndDo
         EndIf
      EndIf
      aAdd(aProdTabs, {YMTP->DA1_CODPRO,YMTP->B1_DESC,YMTP->DA0_DESCRI,YMTP->DA1_PRCVEN,nPrecoApl,YMTP->DA1REC,nPrcMtz})
      DbSelectArea("YMTP")
      DbSkip()
      nCont ++
   EndDo
   DbSelectArea("YMTP")
   DbCloseArea()
   If Empty(aProdTabs)
      aAdd(aProdTabs, {"","","",0,0,0,0})
   EndIf
   
   aSort(aProdTabs, ,,{|x,y| x[1]<y[1]})
   
Return

Static Function MtPrgMTP

   Local aArea := GetArea()

  
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o grupo de produtos que        ')
   Aadd( aHelpPor, 'deseja atualizar pre?os das tabelas de ')
   Aadd( aHelpPor, 'vendas.                                ')


 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"01"  ,"Grupo"  ,"Grupo","Grupo","mv_ch1","C"  ,4       ,0       ,1      ,"G" ,""    ,"SBM" ,""     ,"","MV_PAR01","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Subgrupo Inicial de produtos que')
   Aadd( aHelpPor, 'deseja atualizar pre?os das tabelas de ')
   Aadd( aHelpPor, 'vendas.                                ')


 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"02"  ,"Do Subgrupo"  ,"Do Subgrupo","Do Subgrupo","mv_ch2","C"  ,5       ,0       ,1      ,"G" ,""    ,"SZY" ,""     ,"","MV_PAR02","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o Subgrupo Final de produtos que')
   Aadd( aHelpPor, 'deseja atualizar pre?os das tabelas de ')
   Aadd( aHelpPor, 'vendas.                                ')


 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"03"  ,"At? Subgrupo"  ,"At? Subgrupo","At? Subgrupo","mv_ch3","C"  ,5       ,0       ,1      ,"G" ,""    ,"SZY" ,""     ,"","MV_PAR03","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o fator para reajuste dos         ')
   Aadd( aHelpPor, 'pre?os do Grupo/Subgrupo.              ')


 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"04"  ,"Percentual de Reajuste"  ,"Percentual de Reajuste","Percentual de Reajuste","mv_ch4","N"  ,5       ,3       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR04","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Codigo do Produto inicial para faixa total')
   Aadd( aHelpPor, 'de produtos a ajustar.                 ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"05"  ,"Do Produto"  ,"Do Produto","Do Produto","mv_ch5","C"  ,15       ,0       ,0      ,"G" ,""    ,"SB1" ,""     ,"","MV_PAR05","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Codigo do Produto final   para faixa total')
   Aadd( aHelpPor, 'de produtos a ajustar.                 ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"06"  ,"At? Produto"  ,"At? Produto","At? Produto","mv_ch6","C"  ,15       ,0       ,0      ,"G" ,""    ,"SB1" ,""     ,"","MV_PAR06","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe os c?digos de produtos, separados ')
   Aadd( aHelpPor, 'por virgula, sobre os quais n?o ser?o  ')
   Aadd( aHelpPor, 'selecionados para este ajuste.         ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"07"  ,"Exceto Produtos"  ,"Exceto Produtos","Exceto Produtos","mv_ch7","C"  ,99       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR07","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Codigo da Tabela de Pre?o inicial para')
   Aadd( aHelpPor, ' faixa total de Tabelas a ajustar.    ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"08"  ,"Da Tabela"  ,"Da Tabela","Da Tabela","mv_ch8","C"  ,3       ,0       ,0      ,"G" ,""    ,"DA0" ,""     ,"","MV_PAR08","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Codigo da Tabela de Pre?o final   para')
   Aadd( aHelpPor, ' faixa total de Tabelas a ajustar.    ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"09"  ,"At? Tabela"  ,"At? Tabela","At? Tabela","mv_ch9","C"  ,3       ,0       ,0      ,"G" ,""    ,"DA0" ,""     ,"","MV_PAR09","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe os c?digos das tabelas, separados ')
   Aadd( aHelpPor, 'por virgula, sobre os quais n?o ser?o  ')
   Aadd( aHelpPor, 'selecionadas para este ajuste.         ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"10"  ,"Exceto Tabelas"  ,"Exceto Tabelas","Exceto Tabelas","mv_cha","C"  ,99       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR10","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
      
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a bitola dos produtos. ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"11"  ,"Bitola Produto"  ,"Bitola Produto","Bitola Produto","mv_chb","C"  ,3       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR11","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Codigo da Tabela de Pre?o Matriz')
   Aadd( aHelpPor, 'base para o reajuste dos pre?os ')
   Aadd( aHelpPor, 'das demais tabelas.             ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"12"  ,"Tabela Matriz"  ,"Tabela Matriz","Tabela Matriz","mv_chc","C"  ,3       ,0       ,0      ,"G" ,""    ,"DA0" ,""     ,"","MV_PAR12","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe se o tratamento ser?    ')
   Aadd( aHelpPor, 'realizado sobre a tabela Matriz ')

 //PutSx1(cGrupo,cOrdem,cPergunt             ,""                 ,""                 ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"13"  ,"? Primeira Tabela ?"  ,"? Primeira Tabela ?","? Primeira Tabela ?","mv_chd","N"  ,1       ,0       ,1      ,"C" ,""    ,"" ,""     ,"","MV_PAR13","Sim"  ,"","",""    ,"N?o"      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}
   
   Aadd( aHelpPor, 'Informe o pre?o para sele??o de produtos  ')
   Aadd( aHelpPor, 'definida nos par?metros.               ')


 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"14"  ,"Aplicar Pre?o"  ,"Aplicar Pre?o","Aplicar Pre?o","mv_che","N"  ,17       ,2       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR14","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Medida que ser? filtrada')
   Aadd( aHelpPor, 'na sele??o dos produtos. Se este  ')
   Aadd( aHelpPor, 'par?metro n?o for preenchido ser?o')
   Aadd( aHelpPor, 'consideradas todas as medidas.    ')

 //PutSx1(cGrupo,cOrdem,cPergunt ,""     ,""     ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"15"  ,"Filtrar Medida"  ,"Filtrar Medida","Filtrar Medida","mv_chf","C"  ,4       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR15","","","",""    ,""      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe o metodo do calculo de  ')
   Aadd( aHelpPor, 'juros.                          ')

 //PutSx1(cGrupo,cOrdem,cPergunt             ,""                 ,""                 ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"16"  ,"Calculo do Juros "  ,"Calculo do Juros ","Calculo do Juros ","mv_chg","N"  ,1       ,0       ,1      ,"C" ,""    ,"" ,""     ,"","MV_PAR16","Simples"  ,"","",""    ,"Composto"      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   /*
   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe S=Sim se par o sistema   ')
   Aadd( aHelpPor, 'filtrar apenas as tabelas padr?es')
   Aadd( aHelpPor, 'de vendas, Ex.: Se COSP ser?o    ')
   Aadd( aHelpPor, 'filtradas apenas tabelas com c?di-')
   Aadd( aHelpPor, 'gos iniciados pela letra S, se   ')
   Aadd( aHelpPor, 'CORS ser?o filtradas apenas tabe-')
   Aadd( aHelpPor, 'las iniciadas pela letra R e se  ')
   Aadd( aHelpPor, 'COPE ser?o filtradas apenas tabe-')
   Aadd( aHelpPor, 'las iniciadas pela letra P. Do   ')
   Aadd( aHelpPor, 'contr?rio, informando N=N?o o    ')
   Aadd( aHelpPor, 'sistema ir? ignorar este crit?rio')
   Aadd( aHelpPor, 'e apenas considera a faixa de    ')
   Aadd( aHelpPor, 'tabelas dos par?metros 7,8 E 9.  ')

 //PutSx1(cGrupo,cOrdem,cPergunt             ,""                 ,""                 ,cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,      "",     "",cHelp)
   PutSx1(xPerg ,"17"  ,"Tabelas Padr?o Vendas"  ,"Tabelas Padr?o Vendas","Tabelas Padr?o Vendas","mv_chh","N"  ,1       ,0       ,1      ,"C" ,""    ,"" ,""     ,"","MV_PAR17","Sim"  ,"","",""    ,"N?o"      ,"","",""    ,"","",""    ,"","",""    ,"","",aHelpPor,aHelpEng,aHelpSpa)
   */
   RestArea(aArea)

Return(.T.)

