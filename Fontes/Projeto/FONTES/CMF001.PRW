#INCLUDE "RWMAKE.CH"

//#Include "FiveWin.Ch"
//#Include "TcBrowse.Ch"
//#Include "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001   ³ Autor ³ Sabrina             ³ Data ³ 05.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ VLDUSER - C6_PRCVEN - Programa para a validacao do preco ³±±
±±³ 		 ³ de venda.  												³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CMF001(nRotina)

Local aAreaAtu  := GetArea()
Local lRet		:= .T.
Local nPrcMin	:= 0
Local nCusto	:= 0
Local nPrcVen 	:= &(ReadVar())
Local nPosProd  := 0
Local nPosValor := 0
Local nPosQtd   := 0
Local nPosPrcVd := 0
Local cCliente	:= ""
Local cLoja		:= "" 
Local cTipo		:= ""

If nRotina == 1
	nPosProd  := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_PRODUTO"})  
	nPosValor := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VLRITEM"})  
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_QUANT"})  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VRUNIT"})  
	cCliente	:= M->UA_CLIENTE  
	cLoja		:= M->UA_LOJA
	cTipo		:= "N"
Else
	nPosProd  := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRODUTO"})  
	nPosValor := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_VALOR"})  
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDVEN"})  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRCVEN" })  
	cCliente	:= M->C5_CLIENTE 
	cLoja		:= M->C5_LOJACLI
	cTipo		:= M->C5_TIPO
EndIf
DbSelectArea("SA1")
DbSetOrder(1)
DbSeek(xFilial("SA1")+cCliente+cLoja)

If xFilial("SF2") == "04" .and. cCliente $ "004502" .and. cTipo == "N" .and. AllTrim(Upper(SubStr(cUsuario,7,15))) <> "ADMINISTRADOR"
	If nPrcVen <> CABEDIII(aCols[n][nPosProd])
	
		nPrcVen := CABEDIII(aCols[n][nPosProd])
		aCols[n][nPosPrcVd] := nPrcVen
		aCols[n][nPosValor] := NoRound(nPrcVen * aCols[n][nPosQtd],2)	
		lRet := .F.
		
	EndIf
	
ElseIf SA1->A1_TIPO <> "X"
	If (cTIPO == "N" .OR. nRotina == 1) .And. !l410Auto .AND. !(cCliente $ "004502#006629" .AND. xFilial("SF2")=="01")
		If AllTrim(Upper( SubStr(cUsuario,7,15) ))	 <> "PEDIDO"
		// So verifica o Custo Para produto diferentes de SUCATAS
			If AllTrim(Posicione("SB1",1,xFilial("SB1")+aCols[n][nPosProd],"B1_GRUPO")) <> "80"
				nPrcMin	:= CMF001B(2,nRotina)  
			
				If nPrcVen < nPrcMin .and. nPrcMin > 0 
					MSGINFO("Preco digitado menor que o preco minimo ! Preco Minimo R$ "+ Transform(nPrcMin,"@E 999,999.99"))
					aCols[n][nPosValor] := NoRound(nPrcMin * aCols[n][nPosQtd],2)
					lRet := .F.
				ElseIf nPrcMin == 0
					lRet := .F.
					aCols[n][nPosValor] := NoRound(nPrcMin * aCols[n][nPosQtd],2)
				EndIf	
			EndIf
		Else
			nCusto := (Posicione("SB1",1,xFilial("SB1")+aCols[n][nPosProd],"B1_X_CST2"))*1.05
			If nPrcVen < nCusto
				lRet := .F.
			EndIf	
		Endif	
	EndIf
EndIf

RestArea(aAreaAtu)

Return(lRet)


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001A  ³ Autor ³ Zanardo             ³ Data ³ 18.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Gatilho - C6_PRODUTO - Calculo do preco de Venda. 	    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CMF001A()

Local aArea  		:= GetArea()
Local nPrcVen    	:= 	CMF001B(1)

RestArea(aArea)

Return(nPrcVen)
                      
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001B  ³ Autor ³ Zanardo             ³ Data ³ 18.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Calculo do preco de Venda. 	    						³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CMF001B(nOpcSB1,nRotina)

Local nPosProduto 	:= 0
Local cProduto 		:= ""
Local aUser 		:= U_GetModulo()
Local nPosAtivo 	:= aScan(aUser, {|x| x[3]==.T.})

If nRotina == 1
	nPosProduto 	:= aScan(aHeader,{|x| Alltrim(x[2])=="UB_PRODUTO"})
	cProduto 		:= aCols[N,nPosProduto]
Else
	nPosProduto 	:= aScan(aHeader,{|x| Alltrim(x[2])=="C6_PRODUTO"})
	cProduto 		:= aCols[N,nPosProduto]
EndIf	              

Return(CMF001G(nOpcSB1,cProduto,aUser,nPosAtivo,nRotina))





/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001C  ³ Autor ³ Zanardo             ³ Data ³ 07.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ VLDUSER - C6_TES,C5_TIPO, C6_CF - Autorizacao para       ³±±
±±			 ³ mundanca da TES e o Tipo do Pedido.						³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CMF001C()
Local aArea		:= GetArea()
Local lRet		:= .F.
Local aDepto 	:= PswRet(1)
Local lLibera 	:= .F.
Local aUser 	:= U_GetModulo()
Local nPosAtivo := aScan(aUser, {|x| x[3]==.T.})

If UPPER(Alltrim(aDepto[1,12])) == "FISCAL"  .or. aUser[nPosAtivo,2]>= "4" .or.;
	 (aUser[nPosAtivo,2] == "3" .and. UPPER(Alltrim(aDepto[1,12])) == "FINANCEIRO")
	lLibera 	:= .T.
EndIf 	

RestArea(aArea)
	
Return(lLibera)

//CABEDELO  CONTROLAR CLIENTES PARA COMPRA
User Function CABEDELO()
Local  lLibera :=  .F.
Local  cCliCab := GETMV("MV_CLICAB")
if xfilial("SC6") == '04' 
   if !(Alltrim(M->C5_CLIENTE) $ (cCliCab))
      Alert("Cliente nao autorizado ")      
   Else
      lLibera := .T.
   Endif
Else
   lLibera := .T.      
Endif
Return(lLibera)      

//CABEDIII
User Function CABEDIII(cProduto)
Local nValor := 0              
Local aArea  := GetArea()

DbSelectArea("SB1")
DbSetOrder(1)
If DbSeek(xFilial("SB1")+cProduto)

	IF     ("QUENTE" $  SB1->B1_DESC)
		nValor :=  2200
	ELSEIF ("FRIO" $ SB1->B1_DESC)  
		nValor := 2500
	ELSEIF ("GALVANIZADA"  $ SB1->B1_DESC)
		nValor := 3040
	ELSE  //XADREX 
		nValor := 2500
	ENDIF 
EndIf
RestArea(aArea)

RETURN(nvalor) 

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001D  ³ Autor ³ Larson Zordan       ³ Data ³16.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ VLDUSER - C6_LOTECTL - Valida se o lote esta empenhado e ³±±
±±³		     ³ altera a qtde do pedido conforme a qtde do lote			³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ COMAFAL                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function CMF001D(nRotina)
Local aArea		:= GetArea()
Local lRet		:= .T.
Local nPosProd  := 0
Local nPosQtd   := 0
Local nPosPrcVd := 0
Local nPosVlr   := 0
Local nPosQLb   := 0
Local nPos2UM   := 0
Local nPosLote	:= 0
Local nPosTES 	:= 0
Local cProduto  := ""
Local cNumOP    := ""
Local cNumOC    := ""
Local cCliente  := ""  
Local cTipo		:= ""

If nRotina == 1
	nPosProd  := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_PRODUTO"})  
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_QUANT"})  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VRUNIT"})  
	nPosVlr   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VLRITEM"})  
	nPosLote  := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_LOTE"})
	nPosTES   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_TES"})
	cCliente  := M->UA_CLIENTE
	cTipo		:= "N"
Else
	nPosProd  := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRODUTO"})  
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDVEN" })  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRCVEN" })  
	nPosVlr   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_VALOR"  })  
	nPosQLb   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDLIB" })  
	nPos2UM   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_UNSVEN" })  
	nPosLote  := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_LOTECTL"})
	nPosTES   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_TES"    })       
	cCliente  := M->C5_CLIENTE
	cTipo	  := M->C5_TIPO
EndIf

cProduto  := aCols[n,nPosProd]+"01"+aCols[n,nPosLote]

If (nRotina == 1 .or. cTipo == "N").And. !l410Auto .AND. !(cCliente $ "004502#006629" .AND. xFilial("SF2") == "01")
	If Left(aCols[n,nPosProd],2) == "10" .And. Posicione("SF4",1,xFilial("SF4")+aCols[n,nPosTES],"F4_ESTOQUE") == "S"
		SB8->(dbSetOrder(3))
		SB8->(dbSeek(xFilial("SB8")+cProduto))
		If SB8->B8_EMPENHO > 0
			cNumOP := Posicione("SD4",7,xFilial("SD4")+cProduto,"D4_OP"   )
			cNumOC := Posicione("SD4",7,xFilial("SD4")+cProduto,"D4_NUMOC")
			Aviso("ATENCAO !","Este LOTE " + aCols[n,nPosLote] + " JA Esta Empenhado Para a Ordem de Producao " + cNumOP +;
				If(!Empty(cNumOC)," e Ordem de Corte " + cNumOC,"") + ". Favor Selecionar Outro Lote.",{ " << Voltar "},2,"LOTE EMPENHADO")
			lRet := .F.
		EndIf
		
		If lRet .And. SB8->B8_EMPENHO == 0 .And. (SB8->B8_SALDO > 0 .And. Int(SB8->B8_SALDO) == 0)
			Aviso("ATENCAO !","Este LOTE " + aCols[n,nPosLote] + " NAO tem quantidade suficiente para Venda. ",{ " << Voltar "},2,"QUANTIDADE INSUFICIENTE")
			lRet := .F.
		EndIf
		
		If lRet
			aCols[n,nPosQtd] := SB8->B8_SALDO
			aCols[n,nPosVlr] := SB8->B8_SALDO * aCols[n,nPosPrcVd]
			If nRotina <> 1
				aCols[n,nPosQLb] := SB8->B8_SALDO
				aCols[n,nPos2UM] := SB8->B8_SALDO
			EndIf	
			Aviso("ATENCAO !","Favor Posicionar no Campo Preco Unitario e Confirmar Com <ENTER>. ",{ " << Voltar "},2,"PRECO UNITARIO")
		EndIf
	EndIf
EndIf	
RestArea(aArea)	
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001E  ³ Autor ³ Larson Zordan       ³ Data ³18.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ VLDUSER - C6_PRODUTO - Exibe tela com os lotes disponi-  ³±±
±±³		     ³ veis do produto. (Atualiza aCols conforme Reg Selecinado)³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ COMAFAL                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function CMF001E()
Local oDlg
Local oLbx
Local aLotes    := {}
Local cProd     := &(ReadVar("C6_PRODUTO"))
Local cLote, dDtVal, nQtd
Local lRet      := .T.

Local nPosQtVen := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDVEN" })  
Local nPosLote	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_LOTECTL"})
Local nPosDtVld := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_DTVALID"})

If M->C5_TIPO == "N" .And. !l410Auto

	If AllTrim(Posicione("SB1",1,xFilial("SB1")+cProd,"B1_GRUPO")) > "10"
	
		MsgRun("Aguarde... Verificando Lotes Disponiveis...","Lotes Disponiveis do Produto",{ || CMF001EProc(@aLotes,cProd) })
		
		If Len(aLotes) > 0
		
			Define MsDialog oDlg Title "Visualizar os Lotes do Produto" From 0,0 To 200,350 Of oMainWnd Pixel
			
			@ 0,0 ListBox oLbx Fields Header "Lote","Validade","Qt.Disp." Size 176, 82 Of oDlg Pixel
			
			oLbx:SetArray(aLotes)
			oLbx:bLine := {|| {aLotes[oLbx:nAt,1], aLotes[oLbx:nAt,2], Transform(aLotes[oLbx:nAt,3],"@E 99,999.999") } }
			
			Define SButton oBtn From  85,110 Type  1 Enable Of oDlg Action (aCols[n,nPosLote ]:=aLotes[oLbx:nAt,1],;
																			aCols[n,nPosDtVld]:=aLotes[oLbx:nAt,2],;
																			aCols[n,nPosQtVen]:=aLotes[oLbx:nAt,3],;
																			oDlg:End())
			Define SButton oBtn From  85,140 Type  2 Enable Of oDlg Action (lRet:=.F.,oDlg:End())
			
			Activate MsDialog oDlg Center
			
		Else
			Aviso("ATENCAO !","Nao Existem Lotes Disponiveis Para Venda Deste Produto.",{ " << Voltar " },2,"Lotes Nao Disponiveis")
			lRet := .F.
		EndIf	
	
	EndIf
EndIf 
U_CAF002C(cProd)	
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³CMF001EPro³ Autor ³ Larson Zordan       ³ Data ³18.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa os lotes disponiveis do produto para venda      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ COMAFAL                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function CMF001EProc(aLotes,cProd)
Local aStrucSB8	:= SB8->(dbStruct())
Local cAliasTRB := CriaTrab(Nil,.F.)
Local cQuery    := ""
Local nX

cQuery := "Select * "
cQuery += "From " + RetSqlName("SB8") + " "
cQuery += "Where B8_FILIAL  = '" + xFilial("SB8") + "'"
cQuery += "  And B8_PRODUTO = '" + cProd + "'"
cQuery += "  And B8_SALDO   > 0   " 
cQuery += "  And D_E_L_E_T_ = ' ' "
cQuery += "Order By B8_DTVALID,B8_LOTECTL"
cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTRB,.F.,.T.)
For nX := 1 To Len(aStrucSB8)
	If aStrucSB8[nX][2] <> "C" .And. FieldPos(aStrucSB8[nX][1]) <> 0
		TcSetField(cAliasTRB,aStrucSB8[nX][1],aStrucSB8[nX][2],aStrucSB8[nX][3],aStrucSB8[nX][4])
	EndIf
Next nX

dbSelectArea(cAliasTRB)
dbGoTop()
While !Eof()
	If SB8Saldo(NIL,NIL,NIL,NIL,cAliasTRB) > 0
		aAdd(aLotes,{B8_LOTECTL,B8_DTVALID,SB8Saldo(NIL,NIL,NIL,NIL,cAliasTRB),B8_NUMLOTE})
	EndIf	
	dbSkip()
EndDo

(cAliasTRB)->(dbCloseArea())

Return(.T.)


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001F  ³ Autor ³ Zanardo             ³ Data ³18.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Verifica a Integridade dos valores com a Cond. de Pagto. ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ COMAFAL                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function CMF001F(cCond,nRotina)

Local aArea		:= GetArea()
Local nX		:= 0
Local nPosProd 	:= 0
Local nPosQtd   := 0
Local nPosPrcVd := 0
Local nPosVlr   := 0
Local nPosUnt   := 0
Local aUser 	:= U_GetModulo()
Local nPosAtivo := aScan(aUser, {|x| x[3]==.T.})
Local lConf		:= .F.

If nRotina == 1
	nPosProd  := aScan(aHeader, {|x| Alltrim(x[2]) == "UB_PRODUTO"})
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_QUANT" })  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VRUNIT" })  
	nPosVlr   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VLRITEM"  })  
	nPosUnt   := aScan(aHeader, {|x| AllTrim(x[2]) == "UB_VLRITEM"  })  
Else
	nPosProd  := aScan(aHeader, {|x| Alltrim(x[2]) == "C6_PRODUTO"})
	nPosQtd   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDVEN" })  
	nPosPrcVd := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRCVEN" })  
	nPosVlr   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_VALOR"  })  
	nPosUnt   := aScan(aHeader, {|x| AllTrim(x[2]) == "C6_VALOR"  })  
EndIf

For nX := 1 to Len(aCols)
	If !Empty(aCols[nX,nPosProd])
		aCols[nX,nPosPrcVd] := CMF001G(2,aCols[nX,nPosProd],aUser,nPosAtivo,nRotina)
		aCols[nX,nPosUnt] := aCols[nX,nPosPrcVd]
		aCols[nX,nPosVlr] := aCols[nX,nPosQtd] * aCols[nX,nPosPrcVd]
		lConf		:= .T.
	EndIf
Next nX

If lConf
	Aviso("ATENCAO !","Favor Posicionar no Campo Preco Unitario e Confirmar Com <ENTER>. ",{ " << Voltar "},2,"PRECO UNITARIO")
EndIf	

RestArea(aArea)

Return(.T.)



/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡ao    ³ CMF001G  ³ Autor ³ Zanardo             ³ Data ³ 18.11.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Calculo do preco de Venda. 	    						³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Cliente   ³ Comafal                                                  ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CMF001G(nOpcSB1,cProduto,aUser,nPosAtivo,nRotina)

Local aArea		:= GetArea()
Local nMargem   := 0
Local nPreco	:= 0
Local nCusto 	:= 0
Local lCont 	:= .F.

dbSelectArea("SB1")
dbSetOrder(1)
If dbSeek(xFilial("SB1")+cProduto)
//	If Empty(SB1->B1_X_MGL) .or. Empty(SB1->B1_X_MGLMI) .Or. Empty(SB1->B1_X_MGLMD)
//		MSGINFO("Existe Margem de Lucro zerada, Verificar as margens de lucro minima, media e normal do produto " + cProduto + ". <SB1>")
//		lCont := .F.
//	Else
		//If nOpcSB1 == 2
		//	If aUser[nPosAtivo,2]>= "3"
		//		nMargem := SB1->B1_X_MGLMI
		//	Else
		//		nMargem := SB1->B1_X_MGLMD
		//	EndIf
		//	nCusto := SB1->B1_CUSTD
		//	If nCusto == 0  
		//		MSGINFO("Produto sem custo <SB1>")				
		//		lCont := .F.
		//	Else						
		//		lCont := .T.
		//	EndIf	
	//	Else
			nCusto := SB1->B1_PRV1
			lCont := .T.
			If nCusto == 0  
				MSGINFO("Produto sem Preco de Venda <SB1>")				
				lCont := .F.
			EndIf	
	//	Endif	
//	Endif	
Else
	MSGINFO("Produto nao encontrado <SB1>")
EndIf 
            
If lCont
	If M->C5_TIPO == "N" .or. nRotina == 1
//		If nOpcSB1 == 2
 //			nPreco 	:= nCusto + (nCusto*(nMargem/100))
  //	Else
			nPreco 	:= nCusto		
//		Endif	
		
		dbSelectArea("SA1")          
		dbSetOrder(1)
		If nRotina == 1
			dbSeek(xFilial("SA1")+M->UA_CLIENTE+M->UA_LOJA)
	    Else
		    dbSeek(xFilial("SA1")+M->C5_CLIENTE+M->C5_LOJACLI)
	    EndIf
		dbSelectArea("SE4")          
		dbSetOrder(1)  
		If nRotina == 1
			dbSeek(xFilial("SE4")+UA_CONDPG)
		Else	
			dbSeek(xFilial("SE4")+M->C5_CONDPAG)
		EndIf	
		If SE4->E4_X_PJURO > 0
			nJuros  := SE4->E4_X_PJURO / 100
			nPreco 	+= (nPreco * nJuros)
		EndIf
	EndIf
Endif

RestArea(aArea)
	
Return(nPreco)
