#Include "FiveWin.Ch"
#Include "TopConn.Ch"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Programa ³ SD3250I  ³ Autor ³ Larson Zordan         ³ Data ³10.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Ponto de Entrada Apos a Gravacao do SD3 (Mata250)          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±± 
±±³ ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ PROGRAMADOR  ³ DATA   ³HLPDSK³  MOTIVO DA ALTERACAO                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Larson       ³10/02/04³      ³ Desenvolvimento incial do programa.    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function SD3250I()      
Local oDlg
Local oLbx
Local aAreaAnt := GetArea()
Local aPerda   := {}
Local cLote    := ""
Local cBobina  := ""
Local cDescri  := ""
Local cEndere  := ""
Local cNumOP   := ""
Local cNumOC   := SC2->C2_NUMOC
Local dDtVal   := ""
Local nBobinas := 0
Local nQuant   := 0
Local nSaldo   := 0
Local nPeso    := 0
Local nX
Local lPerda   := .F.
Local lTudo    := .T.

lMsErroAuto    := .F.

If !l250Auto .And. !Empty(cNumOC) .And. AllTrim(Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_GRUPO")) == "20"
	SZ3->(dbSetOrder(1))
	If SZ3->(dbSeek(xFilial("SZ3")+cNumOC))
		If SZ2->(dbSeek(xFilial("SZ2")+cNumOC))
			While !SZ2->(Eof()) .And. xFilial("SZ2")+cNumOC == SZ2->(Z2_FILIAL+Z2_NUM)
				If SZ2->Z2_TIPO = "B"
					nBobinas ++
				EndIf
				SZ2->(dbSkip())
			EndDo
		EndIf
		RecLock("SZ3",.F.)
		Replace SZ3->Z3_OPSBXA With SZ3->Z3_OPSBXA + nBobinas
		MsUnLock()
                   
		nPeso  := SZ3->Z3_PESO
		lPerda := (SZ3->Z3_OPSBXA == SZ3->Z3_QTOPS)

		cLote  := M->D3_LOTECTL
		cNumOP := M->D3_OP
		dDtVal := M->D3_DTVALID
		nQuant := M->D3_QUANT
		

		//--> Grava o Empenho do Slit no Lote (SB8)
		dbSelectArea("SB8")
		dbSetOrder(3)
		If dbSeek(xFilial("SB8")+SC2->C2_PRODUTO+SC2->C2_LOCAL+cLote)
			RecLock("SB8",.F.)
			Replace B8_EMPENHO With nQuant	,;
					B8_EMPENH2 With nQuant
			MsUnLock()	
		EndIf
		
		//--> Grava o Empenho do Slit no Endereco (SBF)
		If Localiza(SC2->C2_PRODUTO)
			dbSelectArea("SBF")
			dbSetOrder(3)
			If dbSeek(xFilial("SBF")+SC2->C2_PRODUTO+SC2->C2_LOCAL+cLote)
				RecLock("SBF",.F.)
				Replace BF_EMPENHO With nQuant	,;
						BF_EMPEN2  With nQuant
				MsUnLock()
			EndIf
		EndIf
	
		//--> Grava o Lote do Slit no SD4
		dbSelectArea("SD4")
		dbSetOrder(1)
		If dbSeek(xFilial("SD4")+SC2->C2_PRODUTO+Left(cNumOP,8))
			RecLock("SD4",.F.)
			Replace D4_LOTECTL With cLote,;
					D4_DTVALID With dDtVal
			MsUnLock()
		EndIf
		
		//--> Apontar a Perda
		If lPerda
			dbSetOrder(2)
			If dbSeek(xFilial("SD4")+cNumOP)

				cBobina := D4_COD
				cDescri := Posicione("SB1",1,xFilial("SB1")+cBobina,"B1_DESC")
				cEndere := Posicione("SB1",1,xFilial("SB1")+cBobina,"B1_ENDEREC")

				While !Eof() .And. xFilial("SD4")+cNumOP == D4_FILIAL+D4_OP
					nSaldo := Posicione("SB8",3,xFilial("SB8")+cBobina+"01"+SD4->D4_LOTECTL,"B8_SALDO")
					If nSaldo > 0
						aAdd(aPerda, {	D4_LOTECTL					,;	// Lote
										cEndere						,;	// Endereco
										nSaldo						,;	// Qtde. de Perda
										CriaVar("B1_COD",.F.)		,;	// Produto Destino
										CriaVar("BC_MOTIVO",.F.)	,;	// Motivo
										D4_DTVALID              	})	// Data Validade
					EndIf
					dbSkip()
				EndDo
            EndIf
            If Len(aPerda) > 0
				Define MsDialog oDlg Title "Apontamento da Perda" From 0,0 To 278,601 Of oMainWnd Pixel
				
				@ 13, 1 To 38, 301 Of oDlg Pixel
				
				@ 22, 10 Say   RetTitle("Z3_PRODUTO")		Size  40,09 Of oDlg Pixel Color CLR_HBLUE
				@ 21, 33 MsGet cBobina 						Size  45,09 Of oDlg Pixel When .F. Center
				@ 21, 80 MsGet cDescri						Size  90,09 Of oDlg Pixel When .F.
				
				@ 40, 1 ListBox oLbx Fields Header "Lote","Endereco","Perda (Ton)","Prod. Destino","Motivo" ;
				Size 300,100 Of oDlg Pixel On DBLCLICK ( aPerda:=CtrPerda(oLbx:nAt,@oLbx,@aPerda,nPeso,@lTudo) )
				oLbx:SetArray(aPerda)
				oLbx:bLine := {|| {aPerda[oLbx:nAt,1], aPerda[oLbx:nAt,2], Transform(aPerda[oLbx:nAT,3],"@E 99,999.999"), aPerda[oLbx:nAT,4], aPerda[oLbx:nAT,5] } }
								
				Activate MsDialog oDlg Center On Init EnchoiceBar(oDlg, {|| oDlg:End() },{ ||oDlg:End() } )

				MsgRun("Aguarde... Apontando as Perdas...","",{ || ApontPerda(cNumOP,cBobina,aPerda) })
                
            EndIf
            
		EndIf
	EndIf
EndIf

RestArea(aAreaAnt)
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Programa ³CtrlPerda ³ Autor ³ Larson Zordan         ³ Data ³12.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Controla a Perda                                           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CtrlPerda(ExpN1,ExpO1,ExpA1,ExpN2)                         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha da ListaBox                                  ³±±
±±³          ³ ExpO1 = Objeto da ListaBox                                 ³±±
±±³          ³ ExpA1 = Array com os dados das perdas                      ³±±
±±³          ³ ExpN2 = Peso total da Ordem de Corte                       ³±±
±±³          ³ ExpL1 = Atualiza a variavel de controle de replicacao no   ³±±
±±³          ³         aPerda.   T = Replica 1x   F = Nao Ira Replicar    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpA1 = Array com os dados da Perda Preenchido             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±± 
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function CtrPerda(nX,oLbx,aPerda,nPeso,lTudo)
Local nOpc   := 0
Local cProd  := ""
Local cMoti  := ""
aPerda := U_CAI002PrdLot(nX,oLbx,aPerda,nPeso,0,4,5)
oLbx:Refresh()
cProd  := aPerda[nX,4]
cMoti  := aPerda[nX,5]
nOpc   := Aviso("ATENÇÃO !","Produto Destino : " + cProd + " " + Posicione("SB1",1,xFilial("SB1")+cProd,"B1_DESC") + Chr(10) + ;
							"Motivo da Perda : " + cMoti + Chr(10) + Chr(10) + ;
							"Deseja Replicar Estes Dados Para Todos Os Lotes ?" + Chr(10) ,;
							{ " Sim "," Não " },2,"Dados do Apontamento da Perda")
If nOpc == 1 .And. lTudo
	AEval(aPerda, {|x| x[4] := cProd } )
	AEval(aPerda, {|x| x[5] := cMoti } )
	lTudo := .F.
EndIf

oLbx:Refresh()

Return(aPerda)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Programa ³ApontPerda³ Autor ³ Larson Zordan         ³ Data ³12.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Apontamento da Perda (Rotina Automatica)                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ApontPerda(ExpC1,ExpC2,ExpA1)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Numero da Ordem de Producao                        ³±±
±±³          ³ ExpC2 = Produto (Bobina)                                   ³±±
±±³          ³ ExpA1 = Array com os dados das perdas                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±± 
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function ApontPerda(cNumOP,cBobina,aPerda)
Local aCabAP   := {}
Local aItemAP  := {}
Local aItem    := {}
Local nX

//Cabecalho
aCabAP:= {	{"BC_OP"		,cNumOP					,Nil}}

For nX := 1 To Len(aPerda)
	
	//Items
	aItem := {	{"BC_PRODUTO"	,cBobina    			,Nil},;
				{"BC_LOCORIG"	,"01"					,Nil},;
				{"BC_LOCALIZ"	,aPerda[nX,2]			,Nil},;
				{"BC_TIPO"  	,"R"					,Nil},;
				{"BC_MOTIVO"	,aPerda[nX,5]			,Nil},;
				{"BC_QUANT" 	,aPerda[nX,3]			,Nil},;
				{"BC_CODDEST"	,aPerda[nX,4]			,Nil},;
				{"BC_LOCAL" 	,"01"					,Nil},;
				{"BC_QTDDEST"	,aPerda[nX,3]			,Nil},;
				{"BC_OPERADOR"	,SubStr(cUsuario,7,15)	,Nil},;
				{"BC_LOTECTL"	,aPerda[nX,1]			,Nil},;
				{"BC_DTVALID"	,aPerda[nX,6]			,Nil} }
	
	aAdd(aItemAP,aItem)
	
	MSExecAuto({|x,y,z| Mata685(x,y,z)},aCabAP,aItemAP,3)
	
	If lMsErroAuto
		Aviso("ATENCAO !","Ocorreu um Erro ao Apontar a Perda da OP " + cNumOP,{" Sair >> "})
		MostraErro()
	EndIf	
	
	aItem   := {}
	aItemAP := {}

Next nX
Return(.T.)