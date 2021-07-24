#Include "fivewin.Ch"
//#Include "rdmake.Ch"
#Include "Protheus.Ch"
//#Include "Colors.Ch"
//#Include "Font.Ch"
//#Include "MsGraphi.Ch"
//#Include "TcBrowse.Ch"
//#Include "TopConn.Ch"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Programa ³  CAI002  ³ Autor ³ Larson Zordan         ³ Data ³17.12.2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Cadastro de Ordem de Corte (Antigo Aproveitamento)         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Observacao³ SX6 - MV_CONSEST = N (NAO Considera o Estoque na OP)       ³±±
±±³          ³ SX1 - MTA65008 = X1_PRESEL = 2 (NAO Sugere Lote/Endereco)  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±± 
±±³ ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ PROGRAMADOR  ³ DATA   ³HLPDSK³  MOTIVO DA ALTERACAO                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Larson       ³17/12/03³      ³ Desenvolvimento incial do programa.    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002()
Local aCores	   := {	{ "Z3_OPSBXA == 0"		 , "ENABLE" 	},;		// OC Em Aberto
						{ "Z3_OPSBXA == Z3_QTOPS", "DISABLE"	},;		// OC Baixada   
						{ "Z3_OPSBXA <  Z3_QTOPS", "BR_AMARELO"	} }		// OC Parcialmente Baixada
Private aButtons   := {	{ "PEDIDO" ,{ || CAI002View(1) },"Visualizar Empenhos"},;
					 	{ "PEDIDO2",{ || CAI002View(2) },"Visualizar Ordens de Produções"}} 
Private lViewRot   := .T. 
Private aRotina    := {	{ OemToAnsi("Pesquisar"       ) ,"AxPesqui"		, 0 , 1},;
                     	{ OemToAnsi("Visualizar"      ) ,"U_CAI002Mnt"		, 0 , 2},;
                     	{ OemToAnsi("Incluir"         ) ,"U_CAI002Mnt"		, 0 , 3},;
                     	{ OemToAnsi("Estornar"        ) ,"U_CAI002Mnt"		, 0 , 5},;
                     	{ OemToAnsi("Imprimir"        ) ,"U_CAI002Imp"		, 0 , 3},;
                     	{ OemToAnsi("&Reabre O.C."    ) ,"U_ReabreOC" 		, 0 , 5},;
                     	{ OemToAnsi("Apontar &Slit"   ) ,"U_CAI002Sli"		, 0 , 3},;
                     	{ OemToAnsi("Estornar &Slit"  ) ,"U_CAI002Mnt" 	, 0 , 5},;
                     	{ OemToAnsi("Apontar &Tubo"   ) ,"U_CAI002Tubo"	, 0 , 3},;
                     	{ OemToAnsi("Estornar &Tubo"  ) ,"U_CAI002EstT"	, 0 , 5},;
                     	{ OemToAnsi("Não Conforme"    ) ,"U_CAI002NCon"	, 0 , 3},;
                     	{ OemToAnsi("Produção/Dia"    ) ,"U_CAI002Agen"	, 0 , 3},;
                     	{ OemToAnsi("Legenda"         ) ,"U_CAI002Leg"		, 0 , 2} }  
Private cCadastro  := OemToAnsi("Cadastro de Ordens de Corte")
mBrowse(6,1,22,75,"SZ3",,,,,,aCores)
Return


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Mnt ³ Autor ³ Larson Zordan        ³ Data ³ 17.12.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Cadastro da Ordem de Corte                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002Mnt(ExpC1,ExpN1,ExpN2)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Mnt(cAlias,nReg,nOpc)
Local oDlg
Local oLbx1
Local oLbx2
Local oFld1
Local oFld2
Local oPerda
Local oPeso
Local oSobra
Local oTotal

Local aInfo	 	 := {}
Local aObjects	 := {}
Local aPosObj 	 := {}
Local aSize	 	 := {}

Local aBobinas   := {}
Local aSliter    := {}
Local aBobSel    := {}
Local aSliSel    := {}
Local aOps       := {}
Local aSD3       := {}
Local aLote      := Array(2)
Local aEmpPerda  := {}
Local aMedidas   := {}

Local cTitulo	 := cCadastro
Local cNumOC     := CriaVar("Z3_NUM"    ,.F.)
Local dData      := CriaVar("Z3_DATA"   ,.F.)
Local cProduto   := CriaVar("Z3_PRODUTO",.F.)
Local cDesc      := CriaVar("Z3_DESC"   ,.F.)
Local cLargRea   := CriaVar("Z3_LARGREA",.F.)
Local nPeso      := CriaVar("Z3_PESO"   ,.F.)
Local nPerda     := CriaVar("Z3_PERDA"  ,.F.)
Local nTotal     := CriaVar("Z3_TOTAL"  ,.F.)
Local nSobra     := CriaVar("Z3_SOBRA"  ,.F.)
Local dPrevIni   := CToD("  /  /  ")
Local dPrevEnt   := CToD("  /  /  ")
Local nBob       := 0
Local nBobi      := 0
Local nSli       := 0
Local nOpca      := 0
Local nX         := 0
Local nPosTraco := 0

Local lEd        := .T.

Private oNo	     := LoadBitmap( GetResources(), "LBNO"       )
Private oOk	     := LoadBitmap( GetResources(), "LBTIK"      )
Private oAb	     := LoadBitmap( GetResources(), "ENABLE"     )
Private oBx	     := LoadBitmap( GetResources(), "DISABLE"    )
Private oPc	     := LoadBitmap( GetResources(), "BR_AMARELO" )

Private lMsErroAuto := .F.

Define FONT oFnt NAME "Tahoma" BOLD

If     nOpc == 2
	cTitulo += " - Visualizar"
ElseIf nOpc == 3
	cTitulo += " - Incluir"
ElseIf nOpc == 4
	cTitulo += " - Estornar Ordem de Corte"
	If Z3_OPSBXA > 0
		Aviso("ATENCAO !","NÃO é possível Estornar esta Ordem de Corte."+Chr(10)+"Estorne as OPs para depois estornar esta OC.",{"<< Voltar"})
		Return(.T.)
	EndIf
ElseIf nOpc == 8
	cTitulo += " - Estornar Apontamento do SLIT"
	If Z3_OPSBXA == 0
		Aviso("ATENCAO !","NÃO é possível Estornar este Apontamento de SLIT"+Chr(10)+"Usar a rotina de Apontamento de Slit.",{"<< Voltar"})
		Return(.T.)
	EndIf
EndIf

If nOpc == 3             
	cNumOC := GetSxeNum("SZ3","Z3_NUM")
	dData  := dDataBase
	aAdd(aBobinas,{ .F., Space(10), 0, CtoD("  /  /  "), Space(6) })
	aAdd(aSliter ,{ Space(15), Space(30), 0, 0, Space(40) })
Else
	cNumOC     := Z3_NUM
	dData      := Z3_DATA
	cProduto   := Z3_PRODUTO
	cDesc      := Z3_DESC
	cLargRea   := Z3_LARGREA
	nPeso      := Z3_PESO
	nPerda     := Z3_PERDA
	nTotal     := Z3_TOTAL
	nSobra     := Z3_SOBRA
	dPrevIni   := Z3_PREVINI
	dPrevEnt   := Z3_PREVENT
	nBob       := 0
	nSli       := 0
	lEd        := .F.
	If SZ2->(dbSeek(xFilial("SZ2")+cNumOC))
		While !SZ2->(Eof()) .And. xFilial("SZ2")+cNumOC == SZ2->(Z2_FILIAL+Z2_NUM)
			If SZ2->Z2_TIPO == "B"
				If nOpc == 8
					aAdd(aBobinas,{	(CAI002SD4(SZ3->Z3_PRODUTO,SubStr(SZ2->Z2_DESC,1,10)) == 0),; // Verifica se a bobina ja foi apontada
									SubStr(SZ2->Z2_DESC, 1,10)	 		,;	// LoteCtl
						 		 	Val(SubStr(SZ2->Z2_DESC,19,12))	,;	// Saldo
						 			CtoD(SubStr(SZ2->Z2_DESC,32, 8))	,;	// Data de Validade
						 			SubStr(SZ2->Z2_DESC,12, 6)			})	// NumLote
				Else	 			 
					aAdd(aBobinas,{	SubStr(SZ2->Z2_DESC, 1,10)	 		,;	// LoteCtl
						 		 	Val(SubStr(SZ2->Z2_DESC,19,12))	,;	// Saldo
						 			CtoD(SubStr(SZ2->Z2_DESC,32, 8))	,;	// Data de Validade
						 			SubStr(SZ2->Z2_DESC,12, 6)			})	// NumLote
				EndIf		 			 
				nBob ++
			Else
				aAdd(aSliter ,{	SubStr(SZ2->Z2_DESC, 1,15)	 			,;	// Slit   
					 			SubStr(SZ2->Z2_DESC,17,30)	 			,;	// Descricao
					 		 	Val(SubStr(SZ2->Z2_DESC,48, 5))		,;	// Qt Rolos
					 		 	Val(SubStr(SZ2->Z2_DESC,54,12))		,;	// Peso Previsto
					 		    SubStr(SZ2->Z2_DESC,67,40) 			})	// Produto Acabado 
				nSli ++
			EndIf		
			SZ2->(dbSkip())
		EndDo
	EndIf
	If Len(aBobinas) == 0
		If nOpc == 8
			aAdd(aBobinas,{ .F., Space(10), 0, CtoD("  /  /  "), Space(6) })
		Else	
			aAdd(aBobinas,{ Space(10), 0, CtoD("  /  /  "), Space(6) })
		EndIf	
	EndIf
	If Len(aSliter) == 0	
		aAdd(aSliter ,{ Space(15), Space(30), 0, 0, Space(40) })
	EndIf	
EndIf

aSize := MsAdvSize()
aInfo := {aSize[1],aSize[2],aSize[3],aSize[4],3,3}

aAdd(aObjects,{100,050,.T.,.F.})
aAdd(aObjects,{100,100,.T.,.T.})
aPosObj := MsObjSize(aInfo, aObjects)

Define MsDialog oDlg Title cTitulo From aSize[7],0 To aSize[6],aSize[5] Of oMainWnd Pixel

@ aPosObj[1,1],aPosObj[1,2] To aPosObj[1,3],aPosObj[1,4] Of oDlg  Pixel

@ 21, 10 Say RetTitle("Z3_NUM")						Size  40,09 Of oDlg Pixel
@ 20, 45 MsGet cNumOC 									Size  40,09 Of oDlg Pixel When .F. Center

@ 21, 95 Say RetTitle("Z3_PRODUTO")					Size  40,09 Of oDlg Pixel Color CLR_HBLUE
@ 20,118 MsGet cProduto  F3 "BOB"						Size  45,09 Of oDlg Pixel When lEd Center Valid NaoVazio(cProduto) .And. CAI002ICpo(cProduto,@cDesc,@cLargRea,@aBobinas,@oLbx1,@aSliter,@oLbx2,@nPeso,@oFld1,@nBob,@oFld2,@nSli)
@ 20,163 MsGet cDesc									Size  90,09 Of oDlg Pixel When .F.

@ 21,275 Say RetTitle("Z3_LARGREA")					Size  40,09 Of oDlg Pixel Color CLR_HBLUE
@ 20,310 MsGet cLargRea Picture "9999"					Size  35,09 Of oDlg Pixel When lEd Center 
@ 21,346 Say "mm"										Size  10,09 Of oDlg Pixel

@ 36, 10 Say RetTitle("Z3_DATA")						Size  40,09 Of oDlg Pixel Color CLR_HBLUE
@ 35, 45 MsGet dData 									Size  40,09 Of oDlg Pixel When lEd

@ 36, 95 Say "Prev. de Inicio"							Size  60,09 Of oDlg Pixel Color CLR_HBLUE
@ 35,135 MsGet dPrevIni									Size  40,09 Of oDlg Pixel When lEd Valid NaoVazio(dPrevIni)

@ 36,190 Say "Prev. de Entrega"							Size  60,09 Of oDlg Pixel Color CLR_HBLUE
@ 35,235 MsGet dPrevEnt									Size  40,09 Of oDlg Pixel When lEd Valid NaoVazio(dPrevEnt) .And. (dPrevIni<=dPrevEnt) .And. (dPrevEnt>=dDataBase)

@ 51, 10 Say RetTitle("Z3_PERDA")						Size  40,09 Of oDlg Pixel
@ 50, 45 MsGet oPerda Var nPerda Picture "@E 999.99"	Size  40,09 Of oDlg Pixel Center When .F. FONT oFnt

@ 51, 95 Say RetTitle("Z3_TOTAL")						Size  40,09 Of oDlg Pixel
@ 50,128 MsGet oTotal Var nTotal Picture "@E 999.999"	Size  40,09 Of oDlg Pixel Center When .F. FONT oFnt
@ 51,168 Say "Ton"										Size  10,09 Of oDlg Pixel

@ 51,190 Say RetTitle("Z3_SOBRA")						Size  40,09 Of oDlg Pixel
@ 50,210 MsGet oSobra Var nSobra Picture "@E 9999"		Size  40,09 Of oDlg Pixel Center When .F. FONT oFnt
@ 51,252 Say "mm"										Size  10,09 Of oDlg Pixel

@ 51,280 Say RetTitle("Z3_PESO")						Size  40,09 Of oDlg Pixel
@ 50,310 MsGet oPeso Var nPeso Picture "@E 99,999.999" 	Size  30,09 Of oDlg Pixel Center When .F. FONT oFnt
@ 51,356 Say "Ton"										Size  10,09 Of oDlg Pixel

//--> ListBox das Bobinas
oFld1:=TFolder():New(aPosObj[2,1],2  ,{"BOBINAS : " + Str(nBob,3)},{},oDlg,,,, .T., .F.,101,10,)
If nOpc == 3 .Or. nOpc == 8
	@ aPosObj[2,1]+10,aPosObj[2,2] ListBox oLbx1 Fields Header "","Lote","Peso","Validade" ;
	ColSizes 2,35,20 Size 100,aPosObj[2,3]-80 Of oDlg Pixel On DBLCLICK ( aBobinas:=CAI002Marc(oLbx1:nAt,aBobinas,@nPeso,@oPeso,@oFld1,@nBob,dPrevIni,dPrevEnt,cProduto,nOpc),oLbx1:Refresh(),oDlg:Refresh() )
	oLbx1:SetArray(aBobinas)
	oLbx1:bLine := {|| {If(aBobinas[oLbx1:nAt,1],oOk,oNo), aBobinas[oLbx1:nAT,2], Transform(aBobinas[oLbx1:nAT,3],"@E 999.999"), DtoC(aBobinas[oLbx1:nAT,4])} }
Else
	@ aPosObj[2,1]+10,aPosObj[2,2] ListBox oLbx1 Fields Header "Lote","Peso","Validade" ;
	ColSizes 35,20 Size 100,aPosObj[2,3]-80 Of oDlg Pixel
	oLbx1:SetArray(aBobinas)
	oLbx1:bLine   := {|| { aBobinas[oLbx1:nAT,1], Transform(aBobinas[oLbx1:nAT,2],"@E 999.999"), DtoC(aBobinas[oLbx1:nAT,3])} }
EndIf

//--> ListBox dos Sliter
oFld2:=TFolder():New(aPosObj[2,1],109,{"SLITS : " + Str(nSli,3)},{},oDlg,,,, .T., .F., aPosObj[2,4]-110,10,)
If nOpc == 3
	@ aPosObj[2,1]+10,110 ListBox oLbx2 Fields Header "Produto","Descricao","Qt. Rolos","Peso Prev.","Prod. Acabado" ;
	ColSizes 30,90,30,30,200 Size aPosObj[2,4]-110,aPosObj[2,3]-80 Of oDlg Pixel On DBLCLICK (CAI002Item(oLbx2:nAt,@aSliter,@oLbx2,cLargRea,nPeso,@nPerda,@nTotal,@nSobra,@nSli,@oFld2,@oPerda,@oTotal,@oSobra),oLbx2:Refresh())
Else
	@ aPosObj[2,1]+10,110 ListBox oLbx2 Fields Header "Produto","Descricao","Qt. Rolos","Peso Prev.","Prod. Acabado" ;
	ColSizes 30,90,30,30,200 Size aPosObj[2,4]-110,aPosObj[2,3]-80 Of oDlg Pixel
EndIf 
oLbx2:SetArray(aSliter)
oLbx2:bLine := {|| {aSliter[oLbx2:nAt,1], aSliter[oLbx2:nAT,2], Transform(aSliter[oLbx2:nAT,3],"99999"), Transform(aSliter[oLbx2:nAT,4],"@E 999.999"), aSliter[oLbx2:nAT,5]} }

Activate MsDialog oDlg Center On Init EnchoiceBar(oDlg,{|| If(CAI002TudoOk(nOpc,aBobinas,aSliter), ( nOpca:=1, oDlg:End() ), .F.) },{||oDlg:End()},,If(lEd,"",aButtons))
	
If nOpca == 1 
	If nOpc == 3	// Incluir
		Begin Transaction
	
			//--> Gravando os dados da Ordem de Corte
			dbSelectArea("SZ3")
			RecLock("SZ3",.T.)
			Replace Z3_FILIAL	With xFilial("SZ3")	,;
					Z3_NUM   	With cNumOC			,;
					Z3_DATA	 	With dData			,;
					Z3_PRODUTO 	With cProduto		,;
					Z3_DESC		With cDesc			,;
					Z3_LARGREA	With cLargRea		,;
					Z3_PESO		With nPeso			,;
					Z3_PERDA	With nPerda			,;
					Z3_TOTAL	With nTotal			,;
					Z3_SOBRA	With nSobra			,;
					Z3_QTOPS	With nSli * nBob   	,;
					Z3_PREVINI	With dPrevIni		,;
					Z3_PREVENT	With dPrevEnt
			MsUnLock()
			
			//--> Grava as Bobinas  - Itens da Ordem de Corte
			dbSelectArea("SZ2")
			For nX := 1 To Len(aBobinas)
				If aBobinas[nX,1]
					aAdd(aBobSel,{cProduto,aBobinas[nX,2],aBobinas[nX,5],aBobinas[nX,3],aBobinas[nX,4]})
					RecLock("SZ2",.T.)
					Replace Z2_FILIAL	With xFilial("SZ2")					   ,;
							Z2_NUM		With cNumOC							   ,;
							Z2_TIPO		With "B"						   	   ,;
							Z2_DESC   	With aBobinas[nX,2] 			+ "|" + ;	// LoteCtl         
											 aBobinas[nX,5]				+ "|" + ;	// NumLote         
											 Str(aBobinas[nX,3],12,3) 	+ "|" + ;	// Saldo do Lote   
											 DToC(aBobinas[nX,4]) 		+ "|"		// Data Validade    
					MsUnLock()
				EndIf
			Next nX 
			
			//--> Grava os Sliters
			For nX := 1 To Len(aSliter)
				If aSliter[nX,3] > 0

					//--> Executa Funcao Para Criar Lote
					l250 := .T.
					U_CAI001(2,@aLote,aSliter[nX,1])
					l250 := .F.
				    nPosTraco := aT("-",aSliter[nX,5])  
				        				        				  				    
					aAdd(aSliSel,{aSliter[nX,1],aSliter[nX,4],alltrim(Left(aSliter[nX,5],nPosTraco-1)),aSliter[nX,3],Val(Right(AllTrim(aSliter[nX,2]),4)),aLote[1],aLote[2]})
					aAdd(aMedidas,{aSliter[nX,3], Val(Right(AllTrim(aSliter[nX,2]),4)) })
					RecLock("SZ2",.T.)
					Replace Z2_FILIAL	With xFilial("SZ2")					   ,;
							Z2_NUM		With cNumOC							   ,;
							Z2_TIPO		With "S"							   ,;
							Z2_DESC   	With aSliter[nX,1] 				+ "|" + ;
											 aSliter[nX,2] 				+ "|" + ;
											 Str(aSliter[nX,3],5) 		+ "|" + ;
											 Str(aSliter[nX,4],12,3) 	+ "|" + ;
											 aSliter[nX,5]                       
					Replace SZ2->Z2_DESC With Left(SZ2->Z2_DESC,109) + aLote[1]
					MsUnLock()
					      
				EndIf
			Next nX
			
			//--> Gravar OPs
			Processa( {|lEnd| CAI002Grava(@lEnd,aBobSel,aSliSel,cNumOC,cProduto,cLargRea,nPerda,nTotal,nSobra,nPeso,dPrevIni,dPrevEnt)},"Aguarde... Gerando Ordens de Producao...")
		
			EvalTrigger()
			If __lSX8
				ConfirmSX8()
			EndIf
		End Transaction
		
		If SZ3->(dbSeek(xFilial("SZ3")+cNumOC))
			//--> Grava o Peso Empenhado e Perda de Cada Lote
			dbSelectArea("SZ2")
			dbSetOrder(1)
			dbSeek(xFilial("SZ2")+cNumOC)
			While !Eof() .And. SZ2->Z2_FILIAL+SZ2->Z2_NUM == xFilial("SZ2")+cNumOC
				If Z2_TIPO = "B"
					aEmpPerda := CalcEmpOC(cNumOC,Val(cLargRea),Val(SubStr(Z2_DESC,19,12)),aMedidas)
					RecLock("SZ2",.F.)
					Replace Z2_DESC With SubStr(Z2_DESC,1,41) + Str(aEmpPerda[1],12,3) + "|" + Str(aEmpPerda[2],12,3) + "|"
					MsUnLock()
				EndIf
				dbSkip()
			EndDo	
		EndIf
	EndIf	

	If nOpc == 4	// Estornar a OC
		Begin Transaction

			//--> Excluindo os dados da Ordem de Corte
			dbSelectArea("SZ3")
			RecLock("SZ3",.F.)
			Delete
			MsUnLock()
	
			//--> Excluindo as Bobinas e Slits da Ordem de Corte
			dbSelectArea("SZ2")
			dbSeek(xFilial("SZ2")+cNumOC)
			While !Eof() .And. xFilial("SZ2")+cNumOC == Z2_FILIAL+Z2_NUM
				RecLock("SZ2",.F.)
				Delete
				MsUnLock()
				dbSkip()
			EndDo
	
			//--> Estornando as OPs
			Processa( {|lEnd| CAI002Estorna(@lEnd,cNumOC,(nSli*nBob))},"Aguarde... Estornando as Ordens de Producao...")
	
		End Transaction
	EndIf	

	If nOpc == 8	// Estornar o Apontamento do SLIT
		Begin Transaction
			//--> Estorna a Producao do Slit
			MsgRun("Aguarde... Estornando Apontamentos de SLITs...","", {|| ProcEstSlit(cNumOC,aBobinas,aSliter,nBobi)})
		End Transaction	
	EndIf
Else
	If ( __lSX8 )
		RollBackSX8()
	EndIf
EndIf                           

Return
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002TudoOk³ Autor ³ Larson Zordan       ³ Data ³ 23.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Valida Lotes e Sliters Selecionados                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002TudoOk(ExpA1,ExpA2)                                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Opcao do aRotina                                   ³±±
±±³          ³ ExpA1 = Array das Bobinas                                  ³±±
±±³          ³ ExpC2 = Array dos Sliters                                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpL1 = Retorno da Confirmacao de itens selecionados       ³±±
±±³          ³         .T. = Tem Itens Selecionados                       ³±±
±±³          ³         .F. = Nao Tem Itens Selecionados                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002TudoOk(nOpc,aBobinas,aSliter)
Local lRet := .F.
Local nX   
If nOpc == 3
	For nX := 1 To Len(aBobinas)
		If aBobinas[nX,1]
			lRet := .T.
			Exit	
		EndIf
	Next nX
    
	If lRet
		For nX := 1 To Len(aSliter)
			If aSliter[nX,3] > 0
				lRet := .T.
				Exit	
			EndIf
		Next nX
	EndIf	

	If !lRet
		Aviso("ATENCAO !","Favor Selecionar Uma BOBINA ou Informar a Qtd. Rolos Para Um SLIT.",{" << Voltar "})
	EndIf
Else
	lRet := .T.
EndIf

Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002ICpo³ Autor ³ Larson Zordan        ³ Data ³ 17.12.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Inicializa o campo Produto                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002Cpo(ExpC1,ExpC2,ExpC3,ExpA1)                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Codigo do Produto                                  ³±±
±±³          ³ ExpC2 = Descricao do Produto   (por referencia)            ³±±
±±³          ³ ExpC3 = Largura Real da Bobina (por referencia)            ³±±
±±³          ³ ExpA1 = Lotes da Bobina        (por referencia)            ³±±
±±³          ³ ExpO1 = Objeto da ListBox 1    (por referencia)            ³±±
±±³          ³ ExpA2 = Slit do Produto        (por referencia)            ³±±
±±³          ³ ExpO2 = Objeto da ListBox 2    (por referencia)            ³±±
±±³          ³ ExpN1 = Peso Total da OC       (por referencia)            ³±±
±±³          ³ ExpO3 = Objeto do Folder 1     (por referencia)            ³±±
±±³          ³ ExpN2 = Bobinas Selecionadas   (por referencia)            ³±±
±±³          ³ ExpO4 = Objeto do Folder 2     (por referencia)            ³±±
±±³          ³ ExpN3 = Sliters Selecionadas   (por referencia)            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002ICpo(cProduto,cDesc,cLargRea,aBobinas,oLbx1,aSliter,oLbx2,nPeso,oFld1,nBob,oFld2,nSli)
Local lRet  := .T.
Local lProc := .T.
SB1->(dbSetorder(1))
If SB1->(dbSeek(xFilial("SB1")+cProduto))
	aBobinas := {}
	aSliter  := {}
	nBob     := 0
	nSli     := 0
	cDesc    := SB1->B1_DESC
	cLargRea := Right(AllTrim(SB1->B1_DESC),4)
	nPeso    := 0
	
	//--> Processa os Lotes das Bobinas
	Processa( { |lEnd| CAI002Lote(@lEnd,@lProc,cProduto,@aBobinas) },"Aguarde...","Pesquisando as Bobinas deste Produto...")

	If !lProc
		Aviso("ATENCAO !!!","Este Produto NAO Tem Lotes Disponiveis Para Aplicacao em Ordem de Corte.",{"<< Voltar"})
		lRet := .F.
	EndIf	

	If lRet
		If lProc
			//--> Processa os Slits do Produto
			Processa( { |lEnd| CAI002Slit(@lEnd,@lProc,cProduto,@aSliter) },"Aguarde...","Pesquisando os Sliters do Produto...")
		EndIf
		If !lProc
			If Len(aSliter) == 0
				aAdd(aSliter,{ Space(9), Space(30), 0, 0, Space(40) } )
			EndIf	
			Aviso("ATENCAO !!!","Este Produto NAO Estruturas Disponiveis Para Aplicacao em Ordem de Corte.",{"<< Voltar"})
			lRet := .F.
		EndIf
	EndIf	

	oFld1:aDialogs[1]:cCaption := "BOBINAS : " + Str(nBob,3) 
	oLbx1:SetArray(aBobinas)
	oLbx1:bLine := {|| {If(aBobinas[oLbx1:nAt,1],oOk,oNo), aBobinas[oLbx1:nAT,2], Transform(aBobinas[oLbx1:nAT,3],"@E 999.999"), DtoC(aBobinas[oLbx1:nAT,4])} }
	oLbx1:Refresh()

	oFld2:aDialogs[1]:cCaption := "SLITS : " + Str(nSli,3) 
	oLbx2:SetArray(aSliter)
	oLbx2:bLine := {|| {aSliter[oLbx2:nAt,1], aSliter[oLbx2:nAT,2], Transform(aSliter[oLbx2:nAT,3],"99999"), Transform(aSliter[oLbx2:nAT,4],"@E 999.999"), aSliter[oLbx2:nAT,5]} }
	oLbx2:Refresh()

Else
	Aviso("ATENCAO !!!","Produto NAO Cadastrado !!!"+Chr(13)+Chr(13)+"Redigite novamente o codigo do produto ou use a tecla F3.",{"<< Voltar"})
	lRet := .F.
EndIf
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Lote³ Autor ³ Larson Zordan        ³ Data ³ 18.12.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa os Lotes da Bobina Selecionada                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002Lote(ExpL1,ExpC1,ExpA1)                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Controle da Funcao Processa                        ³±±
±±³          ³ ExpL2 = Flag de Controle do Array Vazio                    ³±±
±±³          ³ ExpC1 = Codigo do Produto                                  ³±±
±±³          ³ ExpA1 = Lotes da Bobina        (por referencia)            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Lote(lEnd,lProc,cProduto,aBobinas)
Local aStrucSB8	:= SB8->(dbStruct())
Local cAliasTRB := "CALOTE" + xFilial("SB8")
Local cQuery    := ""
Local nX

cQuery := "Select * "
cQuery += "From " + RetSqlName("SB8") + " "
cQuery += "Where B8_FILIAL  = '" + xFilial("SB8") + "'"
cQuery += "  And B8_PRODUTO = '" + cProduto + "'"
cQuery += "  And B8_SALDO   > 0   " 
cQuery += "  And B8_EMPENHO = 0   " 
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
ProcRegua(100)
If Eof()
	lProc := .F.
	aAdd(aBobinas,{ .F., Space(10), 0, CtoD("  /  /  "), Space(6) })
Else
	While !Eof()
	    IncProc()
		If SB8Saldo(NIL,NIL,NIL,NIL,cAliasTRB) > 0
			aAdd(aBobinas,{.F.,B8_LOTECTL,SB8Saldo(NIL,NIL,NIL,NIL,cAliasTRB),B8_DTVALID,B8_NUMLOTE})
		EndIf	
		dbSkip()
	EndDo
EndIf

(cAliasTRB)->(dbCloseArea())

Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Funcao   ³CAI002Marc³ Autor ³ Larson Zordan         ³ Data ³19.12.2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Seleciona Item na ListBox 1 - Bobinas                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Marc(nIt,aArray,nPeso,oPeso,oFld1,nBob,dPrevIni,dPrevEnt,cProduto,nOpc)
If Empty(dPrevIni) .Or. Empty(dPrevEnt)
	Aviso("ATENCAO !!!","Favor informar a Data de Previsao de Inicio e Fim da Producao.",{"<< Voltar"})
Else
	aArray[nIt,1] := !aArray[nIt,1]
	If nOpc == 3
		If aArray[nIt,1]
			If CAI002ValBob(cProduto,aArray[nIt,2],aArray[nIt,5],aArray[nIt,3])
				nPeso += aArray[nIt,3]
				nBob ++
			Else	
				aArray[nIt,1] := .F.
			EndIf	
		Else
			nPeso -= aArray[nIt,3]
			nBob --
		EndIf
	EndIf	
EndIf
oPeso:Refresh()
oFld1:aDialogs[1]:cCaption := "BOBINAS : " + Str(nBob,3) 
Return(aArray)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002ValBob³ Autor ³ Larson Zordan       ³ Data ³ 19.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Valida o Saldo da Bobina (SB8/SBF/SB2)                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002ValBob(ExpC1,ExpC2,ExpC3,ExpN1)                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Produto                                            ³±±
±±³          ³ ExpC2 = Lote do Produto                                    ³±±
±±³          ³ ExpC3 = Sub-Lote do Produto                                ³±±
±±³          ³ ExpN1 = Saldo do Lote                                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002ValBob(cProd,cLoteCtl,cLote,nSaldo)
Local aArea    := GetArea()
Local lRet     := .T.
Local nSldSBF  := 0
Local nSldSB2  := 0

dbSelectArea("SB2")
dbSeek(xFilial("SB2")+cProd+"01")
nSldSB2 := SaldoSB2(.F.,.T.,dDataBase,.T.,.T.,"SB2")

If Localiza(cProd)
	nSldSBF := Posicione("SBF",2,xFilial("SBF")+cProd+"01"+cLoteCtl,"BF_QUANT")
	If nSaldo <> nSldSBF
		Aviso("ATENCAO !",	"O Saldo do Endereco DIVERGE Com o Saldo do Lote."+Chr(10)+;
							"Saldo do Lote : "+Transform(nSaldo,"@E 999,999.999")+Chr(10)+;
							"Saldo do Endereco : "+Transform(nSldSBF,"@E 999,999.999"),{"<< Voltar"})
		lRet := .F.
	EndIf
EndIf
If (nSldSB2 - nSaldo) < 0
	Aviso("ATENCAO !",	"O Saldo Atual Disponivel DIVERGE Com o Saldo do Lote."+Chr(10)+;
						"Saldo do Lote : "+Transform(nSaldo,"@E 999,999.999")+Chr(10)+;
						"Saldo Disponivel : "+Transform(nSldSB2,"@E 999,999.999"),{"<< Voltar"})
	lRet := .F.
EndIf

RestArea(aArea)
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Slit³ Autor ³ Larson Zordan        ³ Data ³ 18.12.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa os Sliters do Produto                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002Slit(ExpL1,ExpC1,ExpA1)                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Controle da Funcao Processa                        ³±±
±±³          ³ ExpL2 = Flag de Controle do Array Vazio                    ³±±
±±³          ³ ExpC1 = Codigo do Produto                                  ³±±
±±³          ³ ExpA1 = Sliters do Produto     (por referencia)            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Slit(lEnd,lProc,cProduto,aSliter)

dbSelectArea("SG1")
dbSetOrder(2)
If dbSeek(xFilial("SG1")+cProduto)
	ProcRegua(100)
	While !Eof() .And. cProduto == G1_COMP
		IncProc()
		If AllTrim(Posicione("SB1",1,xFilial("SB1")+SG1->G1_COD,"B1_GRUPO")) $ "2030"
			aAdd(aSliter,{	G1_COD ,;
							Posicione("SB1",1,xFilial("SB1")+SG1->G1_COD,"B1_DESC") ,;
							0 ,;
							0 ,;
							Space(15) })
		EndIf					
		dbSkip()
	EndDo
EndIf

If Len(aSliter) == 0
	lProc := .F.
	aAdd(aSliter,{ Space(15), Space(30), 0, 0, Space(15) })
EndIf

Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002Item³ Autor ³ Larson Zordan        ³ Data ³ 21/12/2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Edita o campo Qt. Rolos e abre um parambox do Prod.Acabado ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Item(ExpN1,ExpA1,ExpO1)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha do Array a ser editada                       ³±±
±±³          ³ ExpA2 = Array com Dados do Sliter  (por referencia)        ³±±
±±³          ³ ExpO1 = Objeto do ListBox do Sliter                        ³±±
±±³          ³ ExpC1 = Largura Real da Bobina                             ³±±
±±³          ³ ExpN2 = Peso Total dos Lotes Selecionados                  ³±±
±±³          ³ ExpN3 = Perda das Bobinas          (por referencia)        ³±±
±±³          ³ ExpN4 = Total das Bobinas no Corte (por referencia)        ³±±
±±³          ³ ExpN5 = Sobra das Bobinas no Corte (por referencia)        ³±±
±±³          ³ ExpN6 = Qtde de Slits a Cortar     (por referencia)        ³±±
±±³          ³ ExpO2 = Objeto do Folder do Slits  (por referencia)        ³±±
±±³          ³ ExpO3 = Objeto do Get da Perda     (por referencia)        ³±±
±±³          ³ ExpO4 = Objeto do Get do Total     (por referencia)        ³±±
±±³          ³ ExpO5 = Objeto do Get da Sobra     (por referencia)        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Item(nX,aRegs,oLbx,cLargRea,nPeso,nPerda,nTotal,nSobra,nSli,oFld2,oPerda,oTotal,oSobra)
Local lRet    := .F.

If nPeso > 0
	While !lRet
		lEditCell( aRegs, oLbx , "99999", 3 )
		oLbx:SetFocus()
		If aRegs[nX,3] > 0
			//--> ParamBox do Produto Acabado
			lRet := CAI002PA(nX,@aRegs)
		EndIf
		If !lRet
			aRegs[nX,3] := 0
			aRegs[nX,4] := 0
			aRegs[nX,5] := Space(40)
		EndIf	
		//--> Calculo do Slit
		lRet := CAI002Calc(nX,@aRegs,cLargRea,nPeso,@nPerda,@nTotal,@nSobra,@oPerda,@oTotal,@oSobra,@nSli)

		oFld2:aDialogs[1]:cCaption := "SLITS : " + Str(nSli,3)
		oLbx:Refresh()
	EndDo		
Else
	Aviso("ATENCAO !!!","Para informar a Qt. de Rolos, sera necessario selecionar os Lotes ao Lado.",{"<<Voltar"})
	lRet := .F.
EndIf	
Return(aRegs)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002Calc³ Autor ³ Larson Zordan        ³ Data ³ 21/12/2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Seleciona o Produto Acabado aSer Produzido com o Slit      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Calc(ExpN1,ExpA1)                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha do Slit na ListBox                           ³±±
±±³          ³ ExpA1 = Dados do Slit (Array)                              ³±±
±±³          ³ ExpC1 = Largura Real da Bobina em milimetros               ³±±
±±³          ³ ExpN2 = Peso Total dos Lotes Selecionados (Bobinas)        ³±±
±±³          ³ ExpN3 = Percentual de Perda Por Lote (Bobina)              ³±±
±±³          ³ ExpN4 = Total de Tonelanda Prevista para Producao (C/Perda)³±±
±±³          ³ ExpN5 = Total de Sobra em milimetros                       ³±±
±±³          ³ ExpN6 = Qtde de Slits a produzir (Contador do Folder)      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpN1 = Peso Previsto Para Produzir                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Calc(nX,aRegs,cLargRea,nPeso,nPerda,nTotal,nSobra,oPerda,oTotal,oSobra,nSli)
Local lRet      := .T.
Local nTotSlit  := 0
Local nTotPeso  := 0
Local nTotPerd  := 0
Local nCalcPrv  := 0
Local nLargSlit := 0
Local nLargu    := 0
Local nQtRol    := 0
Local nZ        := 0

//--> Calcula os Slits selecionados
For nZ := 1 To Len(aRegs)
	If aRegs[nZ,3] > 0
		nTotSlit +=   Val(Right(AllTrim(aRegs[nZ,2]),4)) * aRegs[nZ,3]							//Somar a Larg do Slit * Qtd de Rolos
		nTotPeso += ((Val(Right(AllTrim(aRegs[nZ,2]),4)) * aRegs[nZ,3]) * nPeso) / Val(cLargRea)	//Somar o Peso Previsto do Slit
	EndIf
Next nZ	

If nPerda == 0
	//--> Calcula a Perda (%)    
	nTotPerd := 100 - ((nTotPeso * 100) / nPeso)
EndIf	

//--> Acrescenta a Perda no Peso da OC
nTotPeso += (nTotPeso * nTotPerd) / 100

//--> Valida a Largura Real da Bobina pela a Qtd de Rolos a Produzir
If nTotSlit > Val(cLargRea)
	Aviso("ATENCAO !!!","A Soma Largura dos Slits a Produzir NAO PODE SER Maior Que a Largura Real da Bobina.",{ "<< Voltar" })
	lRet := .F.
EndIf  

//--> Valida o Peso Total pelo Peso da OC Com a Perda
If nTotPeso > nPeso .Or. nTotPerd < 0
	Aviso("ATENCAO !!!","O Peso Total da OC NAO PODE SER Maior Que o Peso Total das Bobinas Selecionadas.",{ "<< Voltar" })
	lRet := .F.
EndIf  

If lRet
	//--> Qt. de Rolos a Produzir
	nQtRol      := aRegs[nX,3]
	
	//--> Lagura do Sliter (4 ultimos caracteres do campo descricao)
	nLargSlit   := Val(Right(AllTrim(aRegs[nX,2]),4))
	
	//--> Peso Previsto do Slit a ser Produzido
	nCalcPrv    := ((nLargSlit * nQtRol) * nPeso) / Val(cLargRea)
                 
    //--> Grava no Array o Peso Previsto 
    aRegs[nX,4] := nCalcPrv
               
	//--> Refaz os Calculos Acumulados                 
	nTotal := 0
	nSobra := 0
	nPerda := 0
	nSli   := 0
	nLargu := 0
	For nZ := 1 To Len(aRegs)
		//--> Calcula a Sobra da Bobina em Milimetros
		If aRegs[nZ,3] > 0
			//--> Contar do Folder do Slit
			nSli ++
			
			//--> Soma o Total de Peso Previsto
			nTotal += aRegs[nZ,4]
			
			//--> Soma a Largura do Slit Pela Qtd do Rolo
			nLargu += Val(Right(AllTrim(aRegs[nZ,2]),4)) * aRegs[nZ,3]

	    EndIf
	Next nZ

	//--> Calcula a Sobra da Bobina (mm)
	nSobra += Val(cLargRea) - nLargu

	//--> Calcula a Perda (%)    
	nPerda += 100 - ((nTotal * 100) / nPeso)

	//--> Atualiza os Obejtos dos Gets
	oPerda:Refresh()
	oTotal:Refresh()
	oSobra:Refresh()

EndIf

Return(lRet)
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002PA ³ Autor ³ Larson Zordan        ³ Data ³ 21/12/2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Seleciona o Produto Acabado aSer Produzido com o Slit      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002PA(ExpA1)                                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha atual do Slit                                ³±±
±±³          ³ ExpA1 = Array com os Dados do Slit (por referencia)        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpL1 = Retorno logico (T=Continua/F=Nao continua)         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002PA(nX,aRegs)
Local aRet       := {1}
Local aProdAcab  := Array(50)
Local cCompon    := ""
Local cSliter    := aRegs[nX,1]
Local lRet       := .T.
Local nZ         := 0
           
dbSelectArea("SG1")
dbSetOrder(2)
If dbSeek(xFilial("SG1")+cSliter)
	While !Eof() .And. xFilial("SG1")+cSliter == G1_FILIAL+G1_COMP
		nZ ++
		aProdAcab[nZ] := AllTrim(G1_COD) + " - " + AllTrim(Posicione("SB1",1,xFilial("SB1")+SG1->G1_COD,"B1_DESC"))
		dbSkip()
	EndDo
	
	If nZ == 1 
		aRegs[nX,5] := aProdAcab[1]
	Else	
		aSize(aProdAcab,nZ)
		lRet := ParamBox({{3,"Qual Prod. Acabado Ira Produzir",aRet[1],aProdAcab,120,"",.F.}}, "Selecione o Produto Acabado do Slit : " + cSliter, aRet)
		If lRet
			aRegs[nX,5] := aProdAcab[aRet[1]]
		Else
			aRegs[nX,5] := Space(40)
		EndIf
	EndIf
Else 
	Aviso("ATENCAO !!!","Este SLIT NAO Tem Produto Acabado." + Chr(13)+ "Devera Incluir um Produto Acabado pela rotina de Cadastro de Estruturas.",{"<< Voltar"})
	lRet := .F.
EndIf	
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Leg ³ Autor ³ Larson Zordan        ³ Data ³ 17.12.03 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Cria uma janela contendo a legenda da mBrowse              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Leg()
Local aLegenda := {	{"ENABLE"		,"Ordem Corte Em Aberto"		}   ,;
					{"DISABLE"		,"Ordem Corte Baixada"			}   ,; 
					{"BR_AMARELO" 	,"Ordem Corte Com Baixa Parcial"} }    

BrwLegenda(cCadastro,"Legenda",aLegenda)
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002Grava³ Autor ³ Larson Zordan       ³ Data ³ 21/12/2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Grava Ordem de Corte e Cria as OPs e Empenhos (Rot.Automat)³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Grava(ExpL1,ExpA1,ExpA2,ExpN1,ExpN2,ExpN3,ExpN4)     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Variavel logica da funcao Processa()               ³±±
±±³          ³ ExpA1 = Bobinas Selecionadas para a Ordem de Corte         ³±±
±±³          ³ ExpA2 = Slits Selecionadas para a Ordem de Corte           ³±±
±±³          ³ ExpC1 = Numero da Ordem de Corte                           ³±±
±±³          ³ ExpC2 = Produto da Ordem de Corte (Bobina)                 ³±±
±±³          ³ ExpC3 = Largura Real da Bobina                             ³±±
±±³          ³ ExpN1 = Percentual de Perda da Ordem de Corte              ³±±
±±³          ³ ExpN2 = Total de Peso Previsto na Ordem de Corte           ³±±
±±³          ³ ExpN3 = Total de Sobra em mm   na Ordem de Corte           ³±±
±±³          ³ ExpN4 = Peso Total de Bobinas  na Ordem de Corte           ³±±
±±³          ³ ExpD1 = Data da Previsao de Inicio  da OP                  ³±±
±±³          ³ ExpD2 = Data da Previsao de Entreda da OP                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Grava(lEnd,aBobSel,aSliSel,cNumOC,cProduto,cLargRea,nPerda,nTotal,nSobra,nPeso,dPrevIni,dPrevEnt)
Local nX
Local aOP       := {}
Local aProdAcab := ""
Local nQuant    := 0
Local cCC       := ""

//--> Declarado Para Utilizar no Ponto Entrada EMP650()
Private aEmpDad := { cNumOC, cProduto, cLargRea }
Private aEmpBob := aClone(aBobSel)
Private aEmpSli := aClone(aSliSel)
Private lOC		:= .T.

//--> Garante que o Parametro 08 do Pergunte MTA650 seja sempre 2=NAO
If SX1->(dbSeek("MTA65008"))
	RecLock("SX1",.F.)
	Replace X1_PRESEL With 2
	MsUnLock()
EndIf
     
PutMV("MV_EXPLOPU","S")
PutMV("MV_CONSEST","N")

lMsErroAuto     := .F.

ProcRegua(Len(aSliSel))

//--> Gerando as OPs Pela Rotina Automatica             
For nX := 1 To Len(aSliSel)
	IncProc()
	//--> Codigo do Produto Acabado
	cProdAcab := padr(Alltrim(aSliSel[nX,3]),15)  //aSliSel[nX,3]//
	
	//--> Quantidade a Produzir
	nQuant    := aSliSel[nX,2]
	
	//--> Centro de Custo
	cCC       := Posicione("SB1",1,xFilial("SB1")+cProdAcab,"B1_CC")
	
	//--> Monta Array para a Rotina Automatica
	aOP  := {	{"C2_FILIAL"	,xFilial("SC2")	,Nil},;
				{"C2_PRODUTO"	,cProdAcab		,Nil},;
				{"C2_CC"		,cCC			,Nil},;
				{"C2_QUANT"		,nQuant			,Nil},;
				{"C2_QTSEGUM"	,nQuant			,Nil},;
				{"C2_NUMOC"		,cNumOC			,Nil},;
				{"C2_DATPRI"	,dPrevIni 		,Nil},;
				{"C2_DATPRF"	,dPrevEnt 		,Nil},;
				{"AUTEXPLODE"	,"S"			,Nil} } //--> Incluir Este elemento para criar as OPs Intermediarias

	MsExecAuto({|x,y| Mata650(x,y)},aOP,3) //Inclusao
	If lMsErroAuto
		Aviso("ATENCAO !","Ocorreu um Erro ao Incluir uma OP.",{"Sair >>"})
		If lViewRot
			MostraErro()
		EndIf	
	EndIf	
Next nX

PutMV("MV_EXPLOPU","N") 
PutMV("MV_CONSEST","S")

If SX1->(dbSeek("MTA65008"))
	RecLock("SX1",.F.)
	Replace X1_PRESEL With 1
	MsUnLock()
EndIf

MsgRun("Aguarde... Atualizando Lotes Empenhados...",,{ || CAI002SD5() } )

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002SD5 ³ Autor ³ Larson Zordan       ³ Data ³ 26/01/2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Atualiza Dados dos Lotes Empenhados dos Slits              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002SD5()                                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002SD5()
Local cLote     := ""
Local dDtValid  := CtoD("  /  /  ")
Local nQuant    := 0
Local nX

//--> Grava os Numeros das OCs no SD4 (Ajuste de Empenho)
dbSelectArea("SC2")
dbSetOrder(10)
If dbSeek(xFilial("SC2")+SZ3->Z3_NUM)
	While !Eof() .And. C2_FILIAL+C2_NUMOC == xFilial("SC2")+SZ3->Z3_NUM
		dbSelectArea("SD4")
		dbSetOrder(2)
		If dbSeek(xFilial("SD4")+SC2->(C2_NUM+C2_ITEM))
			While !Eof() .And. D4_FILIAL+Left(D4_OP,8) == xFilial("SD4")+SC2->(C2_NUM+C2_ITEM)
				RecLock("SD4",.F.)
				Replace D4_NUMOC   With SC2->C2_NUMOC
				//--> Deixara apenas um Slit com a Qtde Total do Corte no Lote no SD4
				If Left(D4_COD,2) == "20" 
					If !Empty(D4_OPORIG)
						nPos := aScan(aEmpSli,{ |x| x[1] == D4_COD})
						If nPos > 0 
							If aEmpSli[nPos,2] <> D4_QUANT
								//--> Acerta o SC2
								dbSelectArea("SC2")
								dbSetOrder(1)
								If dbSeek(xFilial("SC2")+SD4->D4_OPORIG)
									RecLock("SC2",.F.)
									Replace C2_QUANT   With aEmpSli[nPos,2],;
											C2_QTSEGUM With aEmpSli[nPos,2]
									MsUnLock()
								EndIf
								
								//--> Acerta o SD4
								dbSelectArea("SD4")
								Replace D4_QUANT   With aEmpSli[nPos,2],;
										D4_QTDEORI With aEmpSli[nPos,2],;
										D4_QTSEGUM With aEmpSli[nPos,2]
							EndIf			
						EndIf
					Else
						Delete
					EndIf	
				EndIf	
				MsUnLock()
				dbSkip()
			EndDo
		EndIf
		dbSelectArea("SC2")
		dbSkip()
	EndDo
EndIf

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³  EMP650   ³ Autor ³ Larson Zordan       ³ Data ³ 29/12/2003³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Ponto de Entrada na montagem do aCols dos Empenhos das OPs ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ EMP650()                                                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function EMP650()
Local aAreaSB1 := SB1->(GetArea())
Local aAreaSBF := SBF->(GetArea())
Local cDoc     := ""
Local cLocaliz := ""
Local cNumSeri := ""
Local nPos
Local nQuant   := 0
Local nSlit    := 0
Local nX

/*============================================
Estrutura do Array aEmpDad
--------------------------
aEmpDad => Array Com os Dados da OC
		  aEmpDad[1]   => Numero da OC
		  aEmpDad[2]   => Codigo da Bobina
		  aEmpDad[3]   => Largura Real
==============================================
Estrutura do Array aEmpBob
--------------------------
aEmpBob => Array Com as Bobinas Selecionadas
		  aEmpBob[..1] => Bobina
		  aEmpBob[..2] => LoteCtl
		  aEmpBob[..3] => NumLote
		  aEmpBob[..4] => Saldo
		  aEmpBob[..5] => Data de Validade
==============================================
Estrutura do Array aEmpSli
--------------------------
aEmpSli => Array Com os Slits Selecionados
		  aEmpSli[..1] => Slit
		  aEmpSli[..2] => Peso do Slit
		  aEmpSli[..3] => Produto Acabado			  
		  aEmpSli[..4] => Qt Rolos
		  aEmpSli[..5] => Largura do Sliter
		  aEmpSli[..6] => Lote do Sliter
		  aEmpSli[..7] => Dt Validade do Sliter
=============================================*/

//--> Processo Apenas Para Ordens de Corte (Rotina Automatica)

If l650Auto

	SB1->(dbSetOrder(1))
	SBF->(dbSetOrder(2))
	
	//--> Grava o Num Ordem de Corte na OP
	If Empty(SC2->C2_NUMOC)
		RecLock("SC2",.F.)
		Replace SC2->C2_NUMOC With aEmpDad[1]
		MsUnLock()
    EndIf
    
    //--> Verifica se a Ordem do Empenho eh Slit
    nPos := aScan(aCols,{ |x| Left(x[nPosCod],2) == "20" })
	If nPos > 0
		//--> Varre o aCols Para Deletar os Slits diferentes
		For nX := 1 To Len(aCols)
			nPos := aScan(aEmpSli, { |x| x[1] == aCols[nX,nPosCod] })
			If nPos == 0
				aCols[nX,15] := .T.
			EndIf
	    Next nX
	EndIf
	    
	//--> Verifica se a Ordem do Empenho eh a Bobina
	nPos := aScan(aCols,{ |x| x[nPosCod] == aEmpDad[2] })
	If nPos > 0
		//--> Varre o aCols Para Deletar Bobinas diferentes
		For nX := 1 To Len(aCols)
			//--> So nao ira considerar as bobinas diferentes da Ordem de Corte
			If AllTrim(Posicione("SB1",1,xFilial("SB1")+aCols[nX,1]+"01","B1_GRUPO")) == "10"
				If aCols[nX,1] <> aEmpDad[2]
					aCols[nX,15] := .T.
				EndIf
			EndIf	
		Next nX

		For nX := 1 To Len(aEmpSli)
			//--> 
			If aEmpSli[nX,1] == SC2->C2_PRODUTO
				For nZ := 1 To Len(aEmpBob)
					
					//--> Formula: (( Qt Rolos * Largura do Slit) * Peso da Bobina) / Largura Real da Bobina
					nQuant := Round(( (aEmpSli[nX,4] * aEmpSli[nX,5]) * aEmpBob[nZ,4] ) / Val(aEmpDad[3]),3)
		
					If nPos > 0
						aCols[nPos,nPosQuant  ] := nQuant
						aCols[nPos,nPosLote   ] := aEmpBob[nZ,3]
						aCols[nPos,nPosLotCtl ] := aEmpBob[nZ,2]
						aCols[nPos,nPosDValid ] := aEmpBob[nZ,5]
						aCols[nPos,nPosQtSegum] := ConvUM(aEmpDad[2],nQuant,0,2)
						//--> Zera o nPos, pois para a proxima Bobina, devera criar um novo empenho
						nPos := 0
					Else
						SB1->(dbSeek(xFilial("SB1")+aEmpDad[2]))
						If Localiza(aEmpDad[2]) .And. SBF->(dbSeek(xFilial("SBF")+aEmpDad[2]+"01"+aEmpBob[nZ,2]))
							cLocaliz := SBF->BF_LOCALIZ
							cNumSeri := SBF->BF_NUMSERI
						Else
							cLocaliz := CriaVar("BF_LOCALIZ",.F.)
							cNumSeri := CriaVar("BF_NUMSERI",.F.)
						EndIf
						aAdd(aCols, { 	aEmpDad[2]					  ,;	// Produto (Bobina)
										nQuant						  ,;	// Quantidade Empenhada
										"01"						  ,;	// Armazem
										CriaVar("G1_TRT",.F.)		  ,;	// TRT
										aEmpBob[nZ,3]				  ,;	// NumLote
										aEmpBob[nZ,2]				  ,;	// LoteCtl
										aEmpBob[nZ,5]				  ,;	// Data de Validade
										CriaVar("D4_POTENCI",.F.)	  ,;	// Potencia do Lote
										cLocaliz					  ,;	// Endereco
										cNumSeri					  ,;	// Numero de Serie
										SB1->B1_UM				      ,;	// Unidade de Medida
										ConvUM(aEmpDad[2],nQuant,0,2),;	// Quant na 2a UM
										SB1->B1_SEGUM				  ,;	// 2a UM
										SB1->B1_DESC				  ,;	// Descricao da Bobina
										.F.							  })	// Flag Deletado
					EndIf
			    Next nZ
			EndIf
		Next nX
	EndIf
	
	RestArea(aAreaSBF)
	RestArea(aAreaSB1)
EndIf	

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002Estorna³ Autor ³ Larson Zordan     ³ Data ³ 08/01/2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Estorna as Ordens de Producao (Rot. Automat)               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Estorna(ExpL1,ExpC1)                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Variavel logica da funcao Processa()               ³±±
±±³          ³ ExpC1 = Numero da Ordem de Corte                           ³±±
±±³          ³ ExpN1 = Numero da OPs geradas                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Estorna(lEnd,cNumOC,nOP)
Local aArea := GetArea()
Local aOP   := {}
Local cOP

lMsErroAuto := .F.
        
dbSelectArea("SC2")
dbSetOrder(10)

If dbSeek(xFilial("SC2")+cNumOC)
	ProcRegua(nOP)
	While !Eof() .And. xFilial("SC2")+cNumOC == C2_FILIAL+C2_NUMOC
		IncProc()
		cOP := C2_NUM
		aOP := { {"C2_NUM", C2_NUM, Nil } }
		MSExecAuto({|x,y| Mata650(x,y)},aOP,5) //Exclusao
		If lMsErroAuto
			Aviso("ATENCAO !","Ocorreu um Erro ao Estornar a OP " + cOP + ".",{"Sair >>"})
			If lViewRot
				MostraErro()
			EndIf	
		EndIf	
		dbSkip()
	EndDo
EndIf	

RestArea(aArea)

Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Imp ³ Autor ³ Larson Zordan        ³ Data ³ 07.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Imprimir a Ordem de Corte Posicionada no Browse            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Imp(ExpC1,ExpN1,ExpN2)                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Imp(cAlias,nReg,nOpc)
Local CbTxt      := ""
Local Titulo     := OemToAnsi("Ordem de Corte " + SZ3->Z3_NUM)
Local cDesc1     := OemToAnsi("Esta rotina ira imprimir a Ordem de Corte, bem como as")	
Local cDesc2     := OemToAnsi("Ordens de Produções geradas por ela.                  ")	
Local cDesc3     := OemToAnsi("")
Local aEtiqSlit  := {}
Local wnrel		 := "CAI002"
Local Tamanho    := "M"
Local Limite     := 132
Local cString    := "SZ3"

Private aReturn  := { "Zebrado", 1,"Administracao", 1, 2, 1, "",1 }
Private nLastKey := 0
Private cPerg    := "      "

If Z3_OPSBXA > 0
	Aviso("ATENCAO !","A Ordem de Corte " + SZ3->Z3_NUM + " NAO poder ser impressa pois as OPs ja foram baixadas !",{"<< Voltar"})
	Return(.T.)
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Envia controle para a funcao SETPRINT                        ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
wnrel := SetPrint(cString,wnrel,cPerg,@Titulo,cDesc1,cDesc2,cDesc3,.F.,"",.F.,Tamanho)

If nLastKey == 27
	Set Filter to
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Set Filter to
	Return
Endif

RptStatus({|lEnd| CAI002PImp(@lEnd,wnrel,Titulo,Tamanho,Limite,@aEtiqSlit)},Titulo)

//--> Imprimir as Etiquetas dos Slits
If Aviso("Imprimir...","Deseja Imprimir as Etiquetas de Identificação dos Rolos dos Slits Produzidos Agora ?",{ " Sim "," Não " },2,"Etiqueta do Slit") == 1
	CAI002ImpEtq(aEtiqSlit) 
EndIf	

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002ImpEtq³ Autor ³ Larson Zordan       ³ Data ³01.04.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Imprimir as Etiquetas de Identificacao dos Rolos de Slits  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002ImpEtq(ExpA1)                                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Dados das Etiq de Identif do Slits                 ³±±
±±³          ³         [1] No. de Rastreabilidade (No. Lote)              ³±±
±±³          ³         [2] Peso Calculado                                 ³±±
±±³          ³         [3] Dimensoes                                      ³±±
±±³          ³         [4] OC - Ordem de Corte do Sliter                  ³±±
±±³          ³         [5] OF - Ordem de Fabricacao de Tubos              ³±±
±±³          ³         [6] Maquina                                        ³±±
±±³          ³         [7] Lote da Bobina                                 ³±±
±±³          ³         [8] Maquina Destino                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002ImpEtq(aEtiqSlit)
Local CbTxt      := ""
Local Titulo     := OemToAnsi("Etiquetas da OC " + SZ3->Z3_NUM)
Local cDesc1     := OemToAnsi("Esta rotina ira imprimir as Etiquetas de Identificação")	
Local cDesc2     := OemToAnsi("dos Rolos do Slits Produzidos nesta Ordem de Corte.   ")	
Local cDesc3     := OemToAnsi("")
Local wnrel		 := "CAI002E"
Local Tamanho    := "P"
Local Limite     := 80
Local cString    := "SZ3"

Private aReturn  := { "Zebrado", 1,"Administracao", 1, 2, 1, "",1 }
Private nLastKey := 0
Private cPerg    := "      "

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Envia controle para a funcao SETPRINT                        ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
wnrel := SetPrint(cString,wnrel,cPerg,@Titulo,cDesc1,cDesc2,cDesc3,.F.,"",.F.,Tamanho)

If nLastKey == 27
	Set Filter to
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
	Set Filter to
	Return
Endif

RptStatus({|lEnd| CAI002PIpEtq(@lEnd,wnrel,Titulo,Tamanho,Limite,aEtiqSlit)},Titulo)

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PIpEtq³ Autor ³ Larson Zordan       ³ Data ³01.04.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa a Impressao da Etiqueta dos Rolos dos Slits da OC ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002PIpEtq(lEnd,wnrel,Titulo,Tamanho,Limite,aEtiqSlit)
Local nX
Local Li := 0
SetPrc(0,0)
Li := PRow()
//@ Li,0 PSay Chr(20) + Chr(18)
Li ++ 

For nX := 1 To Len(aEtiqSlit)
	@ Li,01 PSay "Dimensao : " + aEtiqslit[nX,3]
	Li ++
	@ Li,01 PSay "      OC : " + aEtiqslit[nX,4] + "   OF : " + aEtiqslit[nX,5]
	Li ++
	@ Li,01 PSay "Lote Bob.: " + aEtiqSlit[nX,7] + " Lote : " + aEtiqSlit[nX,1]
	Li ++
	If !Empty(aEtiqslit[nX,6])
		@ Li,01 PSay " Maquina : " + Left(Posicione("SI3",1,xFilial("SI3")+aEtiqslit[nX,6],"I3_DESC"),15) + "  Peso :" + Transform(aEtiqslit[nX,2],"@E 99.999")
	Else	
		@ Li,01 PSay " Maquina : " + Space(15) + " Peso :" + Transform(aEtiqslit[nX,2],"@E 99.999")
	EndIf	
	Li ++
//	If xFilial("SF2") == "01"
//		If !Empty(aEtiqslit[nX,8])
//			@ Li,01 PSay "Maq.Dest.: " + Left(Posicione("SI3",1,xFilial("SI3")+aEtiqslit[nX,8],"I3_DESC"),15)
//		Else	
//			@ Li,01 PSay "Maq.Dest.: " + Space(15)
//		EndIf	
//	Else
		@ Li,01 PSay Substr(Alltrim(aEtiqslit[nX,9]),1,40)
//	EndIf	
	Li += 5
Next nX

If aReturn[5] = 1
	Set Printer TO
	dbCommitall()
	ourspool(wnrel)
Endif

MS_FLUSH()

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PImp³ Autor ³ Larson Zordan         ³ Data ³ 07.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa a Impressao da Ordem de Corte                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002PImp(lEnd,wnrel,Titulo,Tamanho,Limite,aEtiqSlit)
Local Cabec1   := " " + AllTrim(SM0->M0_NOME)+"/"+AllTrim(SM0->M0_FILIAL)+" - "+DToC(dDataBase) + "                                    O R D E M   D E   C O R T E                             Nro.: " + SZ3->Z3_NUM
Local Cabec2   := " " + AllTrim(SM0->M0_NOME)+"/"+AllTrim(SM0->M0_FILIAL)+" - "+DToC(dDataBase) + "                                          ORDEM DE CORTE  Nro.: " + SZ3->Z3_NUM + "                               LOTE : "
Local Cabec3   := " " + AllTrim(SM0->M0_NOME)+"/"+AllTrim(SM0->M0_FILIAL)+" - "+DToC(dDataBase) + "                                          ORDEM DE CORTE  Nro.: " + SZ3->Z3_NUM + "                  ORDEM PRODUCAO : "
Local aBobinas := {}
Local aSliter  := {}
Local aOps     := {}
Local aMps     := {}
Local lSalta   := .F.
Local nBob     := 0
Local nSli     := 0
Local nCopia   := 0
Local nX, nZ

If SZ2->(dbSeek(xFilial("SZ2")+SZ3->Z3_NUM))
	While !SZ2->(Eof()) .And. xFilial("SZ2")+SZ3->Z3_NUM == SZ2->(Z2_FILIAL+Z2_NUM)
		If SZ2->Z2_TIPO == "B"
			aAdd(aBobinas,{	 SubStr(SZ2->Z2_DESC, 1,10)	,;	// LoteCtl
				 		 Val(SubStr(SZ2->Z2_DESC,19,12))	,;	// Saldo
				 		CToD(SubStr(SZ2->Z2_DESC,32, 8))	,;	// Data de Validade
				 			 SubStr(SZ2->Z2_DESC,12, 6)	,;	// NumLote
				 		 Val(SubStr(SZ2->Z2_DESC,41,12))	,;	// Qtd. Empenhada
				 		 Val(SubStr(SZ2->Z2_DESC,54,12))})	// Perda
			nBob ++
		Else
			aAdd(aSliter ,{	 SubStr(SZ2->Z2_DESC, 1,15)	,;	// Slit   
				 			 SubStr(SZ2->Z2_DESC,17,30)	,; 	// Descricao
				 		 Val(SubStr(SZ2->Z2_DESC,48, 5))	,; 	// Qt Rolos
				 		 Val(SubStr(SZ2->Z2_DESC,54,12))	,; 	// Peso Previsto
				 		     SubStr(SZ2->Z2_DESC,67,40) })	// Produto Acabado 
			nSli ++
		EndIf		
		SZ2->(dbSkip())
	EndDo
EndIf

Cbtxt  := Space(10)
Cbcont := 0
m_pag  := 0
Li     := 80
nTipo  := If(aReturn[4]==1,15,18)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Impressao da Ordem de Corte                                  ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
For nCopia := 1 To 2
	Cabec("", "", "", "", Tamanho, nTipo, {Cabec1}, .F.)
	
	Li ++
	
	@ Li, 001 PSay "Produto...: " + SZ3->Z3_PRODUTO + SZ3->Z3_DESC
	@ Li, 060 PSay "Largura Real: " + SZ3->Z3_LARGREA + " mm"
	
	Li ++
	
	@ Li, 001 PSay "Emissao...: " + DToC(SZ3->Z3_DATA)
	@ Li, 027 PSay "Previsao de Inicio: " + DToC(SZ3->Z3_PREVINI)
	@ Li, 072 PSay "Previsao de Entrega: " + DToC(SZ3->Z3_PREVINI)
	
	Li ++
	
	@ Li, 001 PSay "Peso Total: " + Transform(SZ3->Z3_PESO,"@E 99,999.999") + " Ton"
	@ Li, 040 PSay "Perda: " + Str(SZ3->Z3_PERDA,6,2) + " %"
	@ Li, 060 PSay "Sobra: " + Str(SZ3->Z3_SOBRA,5,0) + " mm"
	@ Li, 080 PSay "Total da OC: " + Transform(SZ3->Z3_TOTAL,"@E 99,999.999") + " Ton"
	
	Li += 2
	
	@ Li, 000 PSay __PrtThinLine()
	
	Li ++
	
	@ Li, 000 PSay "<< " + Str(nBob,3) + " BOBINAS  >>"
	@ Li, 028 PSay "|"
	@ Li, 030 PSay "<< " + Str(nSli,3) + " SLITS  >>"
	
	@ Li, 000 PSay "<< " + Str(nBob,3) + " BOBINAS  >>"
	@ Li, 030 PSay "<< " + Str(nSli,3) + " SLITS  >>"
	
	Li ++ 
	
	@ Li,000 PSay "Lote         Peso  Validade | Produto   Descricao                     Qt Rolos  Peso Prev. Produto Acabado"
	@ Li,000 PSay "Lote         Peso  Validade | Produto   Descricao                     Qt Rolos  Peso Prev. Produto Acabado"
	//			   XXXXXXXXXX 999,999 99/99/99 | XXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 99999     999,999   XXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
	//			   0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012
	//			             1         2         3         4         5         6         7         8         9        10        11        12        13
	
	Li ++
	
	@ Li,000 PSay __PrtThinLine()
	
	Li ++
	
	For nX := 1 To Max(Len(aBobinas),Len(aSliter))
		If nX <= Len(aBobinas)
			@ Li, 000 PSay aBobinas[nX,1]
			@ Li, 011 PSay aBobinas[nX,2] Picture "@E 999.999"
			@ Li, 019 PSay aBobinas[nX,3]
		EndIf
		@ Li, 028 PSay "|"
		If nX <= Len(aSliter)
			@ Li, 030 PSay Left(aSliter[nX,1],9)
			@ Li, 040 PSay aSliter[nX,2]
			@ Li, 071 PSay aSliter[nX,3] Picture "99999"
			@ Li, 081 PSay aSliter[nX,4] Picture "999.999"
			@ Li, 091 PSay Left(aSliter[nX,5],40)
		EndIf
		Li ++
	Next nX
Next nCopia

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Impressao da Ordem de Corte Por Lote (BOBINA)                ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
For nX := 1 To Len(aBobinas)
	nSli := 0
	
	//--> Cria o Array aOps com os empenhos para cada sliter utilizado no lote
	CAI002ProcOP(aBobinas[nX,1],aSliter,@aOps)
	
	If Len(aOps) > 0
		Cabec("", "", "", "", Tamanho, nTipo, {Cabec2+aBobinas[nX,1]}, .F.)
		Li ++
		@ Li, 001 PSay "Bobina: " + SZ3->Z3_PRODUTO + SZ3->Z3_DESC
		Li += 2
		@ Li, 001 PSay "Largura Real: " + SZ3->Z3_LARGREA + " mm"
		@ Li, 030 PSay "Sobra: " + Str(SZ3->Z3_SOBRA,5,0) + " mm"
		@ Li, 050 PSay "Peso Lote: " + Transform(aBobinas[nX,2],"@E 999.999") + " Ton"
		@ Li, 080 PSay "Peso Verificado: _____________ Ton"
		Li += 2
		@ Li, 000 PSay __PrtThinLine()
		Li ++
		@ Li,000 PSay "Ordem Producao Lote Slit  Produto   Descricao                     Rolos Peso Emp. Peso Produzido Produto Destino"
		@ Li,000 PSay "Ordem Producao Lote Slit  Produto   Descricao                     Rolos Peso Emp. Peso Produzido Produto Destino"
		//			    XXXXXXXXXXXXX XXXXXXXXXX XXXXXXXXX XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX 999   999,999  [____________] xxxxxxxxx xxxxxxxxxxxxxxxxxxxxxxxxxx
		//			   0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012
		//			             1         2         3         4         5         6         7         8         9        10        11        12        13
		Li ++
		@ Li, 000 PSay __PrtThinLine()
		Li ++
		For nZ := 1 To Len(aOps)
			@ Li, 001 PSay aOps[nZ,8]
			@ Li, 015 PSay aOps[nZ,9]
			@ Li, 026 PSay Left(aOps[nZ,1],9)
			@ Li, 036 PSay aOps[nZ,2]
			@ Li, 067 PSay aOps[nZ,3] Picture "999"
			@ Li, 073 PSay aOps[nZ,4] Picture "@E 999.999"
			@ Li, 082 PSay "[____________]"
			@ Li, 097 PSay Left(aOps[nZ,5],36)
			Li ++
			nSli += aOps[nZ,4]
			For nCopia := 1 To aOps[nZ,3]
//				If xFilial("SF2") == '01'
//					aAdd(aEtiqSlit, {	aOps[nZ,9]												,;	// Lote do Slit   		- No. de Rastreabilidade (No. Lote)
//										(aOps[nZ,4]/aOps[nZ,3])								,;	// Peso Empenhado		- Peso Calculado
//										aOps[nZ,2]												,;	// Descricao do Slit	- Dimensoes
//										SZ3->Z3_NUM												,;	// Ordem de Corte		- OC - Ordem de Corte do Slitter
//										aOps[nZ,8]												,;	// Ordem de Producao	- OF Ordem de Fabricacao de Tubos (quando aplicavel)
//										Posicione("SB1",1,xFilial("SB1")+aOps[nZ,1],"B1_CC") ,;	// Centro de Custo		- Maquina (quando aplicavel)
//										aBobinas[nX,1]											,;	// Lote da Bobina   	- No. do Lote da Bobina do Slit
//										Posicione("SB1",1,xFilial("SB1")+Left(aOps[nZ,5],9),"B1_CC")})	// Centro de Custo		- Maquina Destino (CC do PA)
//				Else
					aAdd(aEtiqSlit, {	aOps[nZ,9]												,;	// Lote do Slit   		- No. de Rastreabilidade (No. Lote)
										(aOps[nZ,4]/aOps[nZ,3])								,;	// Peso Empenhado		- Peso Calculado
										aOps[nZ,2]												,;	// Descricao do Slit	- Dimensoes
										SZ3->Z3_NUM												,;	// Ordem de Corte		- OC - Ordem de Corte do Slitter
										aOps[nZ,8]												,;	// Ordem de Producao	- OF Ordem de Fabricacao de Tubos (quando aplicavel)
										Posicione("SB1",1,xFilial("SB1")+aOps[nZ,1],"B1_CC") ,;	// Centro de Custo		- Maquina (quando aplicavel)
										aBobinas[nX,1]											,;	// Lote da Bobina   	- No. do Lote da Bobina do Slit
										Posicione("SB1",1,xFilial("SB1")+Left(aOps[nZ,5],13),"B1_CC"),;
										Posicione("SB1",1,xFilial("SB1")+Left(aOps[nZ,5],13),"B1_DESC")})	// Centro de Custo		- Maquina Destino (CC do PA)
				
//				EndIf											
			Next nCopia						
		Next nZ
		@ Li, 000 PSay __PrtThinLine()
		Li ++
		@ Li, 054 PSay "Total Empenhado => " + Transform(nSli,"@E 999.999") + " Ton"
		Li ++
		@ Li, 054 PSay "Perda Prevista  => " + Transform(aBobinas[nX,2]-nSli,"@E 999.999") + " Ton"
		Li := 80
	EndIf
Next nX	


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Impressao da Ordem de Producao do Produto Acabado            ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
SC2->(dbSetOrder(1))
SD4->(dbSetOrder(2))
For nCopia := 1 To 2
	For nX := 1 To Len(aOps)
	
		//--> Cria o Array aMps do Produto Acabado com os Slis Cortados e Lotes
		If SD4->(dbSeek(xFilial("SD4")+Left(aOps[nX,8],8)))
			While !SD4->(Eof()) .And. xFilial("SD4")+Left(aOps[nX,8],8) == SD4->(D4_FILIAL+Left(D4_OP,8))
				If !Empty(SD4->D4_LOTECTL)
					aAdd(aMps, { SD4->D4_LOTECTL, SD4->D4_COD==aOps[nX,1], SD4->D4_QUANT } )
					nSli += SD4->D4_QUANT
				EndIf	
				SD4->(dbSkip())
			EndDo
		EndIf
		
		If Len(aMps) > 0                
			//--> Posiciona nas OPs
			SC2->(dbSeek(xFilial("SC2")+aOps[nX,8]))
			
			Cabec("", "", "", "", Tamanho, nTipo, {Cabec3+aOps[nX,8]}, .F.)
			Li ++
			@ Li, 001 PSay "Produto: " + aOps[nX,5]
			@ Li, 070 PSay "Emissao: " + DtoC(SC2->C2_EMISSAO)
			If nCopia == 1
				@ Li,114 PSay "<< VIA OPERADOR >>"
			Else
				@ Li,111 PSay "<< VIA ALIMENTADOR >>"
			EndIf	
			Li ++
			@ Li, 001 PSay "Quant..: " + Transform(nSli,"@E 99,999.999") + " " + SC2->C2_UM
			@ Li, 070 PSay "C.Custo: " + SC2->C2_CC
			Li ++
			@ Li, 000 PSay __PrtThinLine()
			Li ++
			@ Li, 020 PSay "               I N I C I O                         F I M"
			Li +=2
			@ Li, 020 PSay " Previsao : _____/_____/_____                _____/_____/_____"
			Li += 2
			@ Li, 020 PSay "   Ajuste : _____/_____/_____                _____/_____/_____"
			Li += 2
			@ Li, 020 PSay "     Real : _____/_____/_____                _____/_____/_____"
			Li ++
			@ Li, 000 PSay __PrtThinLine()
			Li ++
			@ Li,000 PSay "  Produto          Descricao                       Lote Slit   Lote Bobina   Quant.  UM  Rolos"
			@ Li,000 PSay "  Produto          Descricao                       Lote Slit   Lote Bobina   Quant.  UM  Rolos"
			//				 XXXXXXXXXXXXXXX  XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX  XXXXXXXXXX   XXXXXXXXXX  999,999  XX    999
			//			   0123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012
			//			             1         2         3         4         5         6         7         8         9        10        11        12        13
			Li ++
			@ Li, 000 PSay __PrtThinLine()
			Li ++
	
			For nZ := 1 To Len(aMps)
				@ Li,002 PSay aOps[nX,1]										   		// Slit
				@ Li,019 PSay aOps[nX,2]												// Descricao
				If aMps[nZ,2]
					@ Li,051 PSay aMps[nZ,1]
				Else
					@ Li,051 PSay aOps[nX,9]											// Lote do Slit
					@ Li,064 PSay aMps[nZ,1]											// Lote da Bobina Ocupada
				EndIf	
				@ Li,076 PSay aMps[nZ,3] Picture "@E 999.999"							// Quantidade
				@ Li,085 PSay Posicione("SB1",1,xFilial("SB1")+aOps[nX,1],"B1_UM")	// UM
				@ Li,090 PSay aOps[nX,3] Picture "9999"
				Li ++
			Next nZ
			
			aMps := {}
			nSli := 0
	
	    EndIf
	
	Next nX
Next nCopia

If aReturn[5] = 1
	Set Printer TO
	dbCommitall()
	ourspool(wnrel)
Endif

MS_FLUSH()

Return NIL

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002ProcOP³ Autor ³ Larson Zordan       ³ Data ³ 08.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa as OPs rerefentes a OC impressa                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002ProcOP(ExpC1,ExpA1,ExpA2)                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Lote da Bobina                                     ³±±
±±³          ³ ExpA1 = Dados dos Sliters                                  ³±±
±±³          ³ ExpA2 = Dados dos Sliters Empenhados (por referencia)      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002ProcOP(cLoteCtl,aSliter,aOps)
Local cAliasTRB := "CAI002OP"		
Local aStrucSD4	:= SD4->(dbStruct())
Local cSlit
Local cLote
Local cLoteBob
Local cOP
Local nPos
Local nX 

aOps   := {}

cQuery := "Select * From " + RetSqlName("SD4") + " "
cQuery += "Where D4_FILIAL  = '" + xFilial("SD4")+ "'"
cQuery += "  And D4_LOTECTL = '" + cLoteCtl + "'"
cQuery += "  And D4_QUANT   > 0  "
cQuery += "  And D_E_L_E_T_ = ' '"
cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTRB,.F.,.T.)
For nX := 1 To Len(aStrucSD4)
	If aStrucSD4[nX][2] <> "C" .And. FieldPos(aStrucSD4[nX][1]) <> 0
		TcSetField(cAliasTRB,aStrucSD4[nX][1],aStrucSD4[nX][2],aStrucSD4[nX][3],aStrucSD4[nX][4])
	EndIf
Next nX

dbSelectArea(cAliasTRB)
While !Eof()

	//--> Sempre a Bobina sera OP com sequencia 002 e Slit com sequencia 001
	cOP   := Left(D4_OP,8) + "001  "
	
	//--> Lote da Bobina
	cLoteBob := D4_LOTECTL
	
	//--> Pesquisa a OP do Slit
	cSlit := Posicione("SD4",2,xFilial("SD4")+cOP,"D4_COD")
	
	//--> Posiciona no array aSliter
	nPos := aScan(aSliter,{ |x| x[1] == cSlit }) 

	If nPos > 0	
		//--> Posiciona o Lote do Slit no Z2_DESC
		cLote := Space(10)	
		If SZ2->(dbSeek(xFilial("SZ2")+SZ3->Z3_NUM))
			While !SZ2->(Eof()) .And. xFilial("SZ2")+SZ3->Z3_NUM == SZ2->(Z2_FILIAL+Z2_NUM)
				If SZ2->Z2_TIPO == "S" .And. SubStr(SZ2->Z2_DESC,1,15) == cSlit
					cLote := SubStr(SZ2->Z2_DESC,110,10)
					Exit
				EndIf
				SZ2->(dbSkip())
			EndDo
		EndIf
		
		aAdd(aOps,{	aSliter[nPos,1]	,;	// Sliter
					aSliter[nPos,2]	,;	// Descricao
					aSliter[nPos,3]	,;	// Qt Rolos
					D4_QUANT  		,;	// Peso Empenhado do Lote (Bobina)
					aSliter[nPos,5]	,;	// Prod.Acabado
					0    			,;	// Qtde Produzida
					.F.				,;	// Flag para Marcar
					D4_OP			,;	// OP do Slit
					cLote			,;	// Lote do Slit
					cLoteBob   		})	// Lote da Bobina
    EndIf
	dbSkip()
EndDo
(cAliasTRB)->(dbCloseArea())

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Sli  ³ Autor ³ Larson Zordan       ³ Data ³ 08.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Apontamento de Sliters na Ordem de Corte                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Sli(ExpC1,ExpN1,ExpN2)                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Sli(cAlias,nReg,nOpc)
Local oDlg
Local oFld
Local oGet
Local oLbx1
Local oLbx2

Local aButtons   := {	{ "PEDIDO" ,{ || CAI002View(1) },"Visualizar Empenhos"},;
					 	{ "PEDIDO2",{ || CAI002View(2) },"Visualizar Ordens de Producçõs"}} 

Local aInfo	 	 := {}
Local aObjects	 := {}
Local aPosObj 	 := {}
Local aSize	 	 := {}

Local aBobinas   := {}
Local aBobs      
Local aMedidas   := {}
Local aOps       := {}
Local aSliter    := {}
Local aEtiqSlit  := {}
Local aPerda     := {}

Local nOpca      := 0

Local cLote      := CriaVar("D4_LOTECTL",.F.)
Local cNumOC     := CriaVar("Z3_NUM"    ,.F.)
Local cNumOP     := CriaVar("C2_NUM"    ,.F.)
Local cCC        := CriaVar("C2_CC"     ,.F.)
Local cCCTubo    := CriaVar("C2_CC"     ,.F.)
Local nPeso      := 0
Local nPVer      := 0
Local nX, nY

Local lEd        := .T.
Local lErro      := .F.

Local oCbx

aSize := MsAdvSize()
aInfo := {aSize[1],aSize[2],aSize[3],aSize[4],3,3}

aAdd(aObjects,{100,050,.T.,.F.})
aAdd(aObjects,{100,100,.T.,.T.})
aAdd(aObjects,{100,050,.T.,.F.})
aPosObj := MsObjSize(aInfo, aObjects)

cNumOC := SZ3->Z3_NUM
If SZ3->Z3_OPSBXA == SZ3->Z3_QTOPS
	Aviso("ATENCAO !","Ordem de Corte Com BAIXA TOTAL.",{"<< Voltar"})
	Return(.T.)
EndIf

cNumOP := Posicione("SC2",9,xFilial("SC2")+cNumOC,"C2_NUM")+"01002  "
cCC    := Posicione("SC2",1,xFilial("SC2")+cNumOP,"C2_CC" )

//--> Monta dados relacionados as Bobinas (Lotes) e Slits
If SZ2->(dbSeek(xFilial("SZ2")+cNumOC))
	While !SZ2->(Eof()) .And. xFilial("SZ2")+cNumOC == SZ2->(Z2_FILIAL+Z2_NUM)
		If SZ2->Z2_TIPO == "B"
			If CAI002SD4(SZ3->Z3_PRODUTO,SubStr(SZ2->Z2_DESC,1,10)) > 0  	// Verifique se a bobina tem saldo
				aAdd(aBobinas,{	 SubStr(SZ2->Z2_DESC, 1,10)	 		,;		// LoteCtl
						 		 Val(SubStr(SZ2->Z2_DESC,19,12))		,; 		// Saldo
								 CtoD(SubStr(SZ2->Z2_DESC,32, 8))		,;		// Data de Validade
						 		 SubStr(SZ2->Z2_DESC,12, 6)	 		,;		// NumLote
				 		 		 Val(SubStr(SZ2->Z2_DESC,41,12))		,;	 	// Qtd. Empenhada
				 		 		 Val(SubStr(SZ2->Z2_DESC,54,12))		,;		// Perda
						 		 SubStr(SZ2->Z2_DESC,67, 2) 			})		// Flag de Controle de Baixas "Bx" OP Baixada
			EndIf
		Else
			aAdd(aSliter ,{	 SubStr(SZ2->Z2_DESC, 1,15)	       		,;		// Slit   
				 			 SubStr(SZ2->Z2_DESC,17,30)	       		,; 		// Descricao
				 		     Val(SubStr(SZ2->Z2_DESC,48, 5))  		,; 		// Qt Rolos
				 		     Val(SubStr(SZ2->Z2_DESC,54,12))  		,; 		// Peso Previsto
				 		     SubStr(SZ2->Z2_DESC,67,40) })	   				// Produto Acabado 
		EndIf
		SZ2->(dbSkip())
	EndDo
EndIf

If Len(aBobinas) == 0
	Aviso("ATENCAO !","Ordem de Corte Foi Baixado Pelo Apontamento de Producao Padrao.",{" << Voltar "})
	Return(.T.)
EndIf
	
If Len(aSliter) == 0
	Aviso("ATENCAO !","Ordem de Corte Foi Baixado Pelo Apontamento de Producao Padrao.",{" << Voltar "})
	Return(.T.)
EndIf
	
If Len(aBobinas) > 0
	//--> Monta Array da Combo dos Lotes da Bobina
	aBobs := Array(Len(aBobinas))
	For nX := 1 To Len(aBobinas)
		aBobs[nX] := aBobinas[nX,1]
	Next nX
	nPeso := aBobinas[1,2]
		
	//--> Monta Array das Ops
	CAI002ProcOP(aBobs[1],aSliter,@aOps)
	AEval(aOps, {|x| x[6] := x[4] } )
		       
	//--> Monta Array Com as Medidas dos Slits
	For nX := 1 To Len(aSliter)
		aAdd(aMedidas,{aSliter[nX,1],aSliter[nX,4],Left(aSliter[nX,5],9),aSliter[nX,3],Val(Right(AllTrim(aSliter[nX,2]),4))})
	Next nX
		          
	//--> Monta Array Com os Dados Para Apontar a Perda
	CAI002Perda(@aPerda,aBobinas[1,1],aBobinas)
EndIf

Define MsDialog oDlg Title "Apontamento do SLIT da Ordem de Corte" From aSize[7],0 To aSize[6],aSize[5] Of oMainWnd Pixel

@ aPosObj[1,1],aPosObj[1,2] To aPosObj[1,3],aPosObj[1,4] Of oDlg  Pixel

@ 21, 10 Say RetTitle("Z3_NUM")						Size  40,09 Of oDlg Pixel Right
@ 20, 55 MsGet cNumOC Picture "@!" F3 "SZ3"				Size  30,09 Of oDlg Pixel When .F.

@ 21, 95 Say RetTitle("Z3_PRODUTO")					Size  40,09 Of oDlg Pixel Right 
@ 20,133 MsGet SZ3->Z3_PRODUTO   						Size  45,09 Of oDlg Pixel When .F.
@ 20,180 MsGet SZ3->Z3_DESC							Size  90,09 Of oDlg Pixel When .F.

@ 36, 10 Say RetTitle("Z3_LARGREA")					Size  40,09 Of oDlg Pixel Right
@ 35, 55 MsGet SZ3->Z3_LARGREA Picture "9999"			Size  25,09 Of oDlg Pixel When .F. Center
@ 36, 85 Say "mm"										Size  10,09 Of oDlg Pixel

@ 36, 95 Say RetTitle("Z3_SOBRA")						Size  40,09 Of oDlg Pixel Right
@ 35,133 MsGet SZ3->Z3_SOBRA   Picture "@E 999,999"	Size  30,09 Of oDlg Pixel When .F. Center
@ 36,170 Say "mm"										Size  10,09 Of oDlg Pixel

@ 36,200 Say "Centro de Custo"							Size  40,09 Of oDlg Pixel Right
@ 35,250 MsGet oGet Var cCC F3 "CTT"					Size  30,09 Of oDlg Pixel
        
@ 51, 10 Say "Lote"										Size  40,09 Of oDlg Pixel Right Color CLR_HBLUE
@ 50, 55 ComboBox oCbx Var cLote Items aBobs 			Size  45,40 Of oDlg Pixel When lEd Valid ( If(!Empty(cLote),(CAI002ProcOP(cLote,aSliter,@aOps), (nPeso:=aBobinas[aScan(aBobs,cLote),2]), CAI002Perda(@aPerda,aBobinas[aScan(aBobs,cLote),1],aBobinas), AEval(aOps, {|x| x[6] := x[4] } ), oLbx1:Refresh(), oLbx2:Refresh()), .F.) )
          
@ 51, 95 Say "Peso"										Size  30,09 Of oDlg Pixel Right
@ 50,133 MsGet nPeso Picture "@E 999.999"				Size  30,09 Of oDlg Pixel When .F. Center
@ 51,170 Say "Ton"										Size  10,09 Of oDlg Pixel

@ 51,182 Say "Peso Verificado"							Size  45,09 Of oDlg Pixel Right
@ 50,230 MsGet nPVer Picture "@E 999.999"				Size  30,09 Of oDlg Pixel Valid ( If(nPVer>0, AEval(aOps,{|x| x[6]:=0}), AEval(aOps,{|x| x[6]:=x[4]}) ), oLbx1:Refresh() ) 
@ 51,265 Say "Ton"										Size  10,09 Of oDlg Pixel

@ aPosObj[2,1]+1,3 ListBox oLbx1 Fields Header "Ord. Producao","Produto","Descricao","Qt. Rolos","Peso Emp.","Peso Prod.","Prod. Acabado" ;
ColSizes 40,40,90,30,30,30,230 Size aPosObj[2,4]-3,aPosObj[2,3]-78 Of oDlg Pixel On DBLCLICK ( aOps:=CAI002SliOP(oLbx1:nAt,aOps,cLote,nPeso,nPVer,@lEd,@oLbx1,aMedidas), oLbx1:Refresh(), oCbx:Refresh(), oDlg:Refresh() )
oLbx1:SetArray(aOps)
oLbx1:bLine := {|| {aOps[oLbx1:nAt,8], aOps[oLbx1:nAt,1], aOps[oLbx1:nAT,2], Transform(aOps[oLbx1:nAT,3],"99999"), Transform(aOps[oLbx1:nAT,4],"@E 999.999"), Transform(aOps[oLbx1:nAT,6],"@E 999.999"), aOps[oLbx1:nAT,5]} }

oFld:=TFolder():New(aPosObj[3,1]-8,2  ,{"Apontamento da Perda"},{},oDlg,,,, .T., .F.,,10,)
@ aPosObj[3,1]+1,3 ListBox oLbx2 Fields Header "Bobina","Lote","Endereco","Perda (mm)","Perda (Ton)","Prod. Destino","Motivo" ;
Size aPosObj[3,4]-3,45 Of oDlg Pixel On DBLCLICK ( aPerda:=U_CAI002PrdLot(oLbx2:nAt,oLbx2,aPerda,nPeso,nPVer,6,7), oLbx2:Refresh() )
oLbx2:SetArray(aPerda)
oLbx2:bLine := {|| {aPerda[oLbx2:nAt,1], aPerda[oLbx2:nAt,2], aPerda[oLbx2:nAT,3], Transform(aPerda[oLbx2:nAT,4],"99999"), Transform(aPerda[oLbx2:nAT,5],"@E 999.999"), aPerda[oLbx2:nAT,6], aPerda[oLbx2:nAT,7] } }

Activate MsDialog oDlg On Init EnchoiceBar(oDlg,{|| If(CAI002PrdVld(aOps,aPerda),(nOpca:=1,oDlg:End()),.T.)  },{||oDlg:End()},,aButtons)

If nOpca == 1
	Begin Transaction
	
		//--> Apontamento de Producao e Perda
		Processa( { |lEnd| CAI002ApntOP(@lEnd,aOps,aPerda,@lErro,cCC) },"Aguarde...","Processando Apontamento de Producao e Perda...")

		If !lErro
			//--> Baixa das Qtde de OPs no SZ3
			RecLock("SZ3",.F.)
			Replace SZ3->Z3_OPSBXA With (SZ3->Z3_OPSBXA+Len(aOPs))
			MsUnLock()
		EndIf
		
	End Transaction
              
	If !lErro
		//--> Monta Array com os dados Para Etiqueta
		For nX := 1 To Len(aOps)
			For nY := 1 To aOps[nX,3]
				cCCTubo := Posicione("SB1",1,xFilial("SB1")+Left(aOps[nX,5],9),"B1_CC")
				aAdd(aEtiqSlit, {	aOps[nX,9]					,;		// Lote do Slit   		- No. de Rastreabilidade (No. Lote)
									(aOps[nX,6]/aOps[nX,3])	,;		// Peso Produzido		- Peso Calculado
									aOps[nX,2]					,;		// Descricao do Slit	- Dimensoes
									cNumOC						,;		// Ordem de Corte		- OC - Ordem de Corte do Slitter
									aOps[nX,8]					,;		// Ordem de Producao	- OF Ordem de Fabricacao de Tubos (quando aplicavel)
									cCC 						,;		// Centro de Custo		- Maquina (quando aplicavel)
									aOps[nX,10]					,;		// Lote da Bobina   	- No. do Lote da Bobina do Slit
									cCCTubo						})		// Centro de Custo		- Maquina Destino (CC do PA)
									
			Next nY
		Next nX
			
		//--> Imprimir as Etiquetas dos Slits
		CAI002ImpEtq(aEtiqSlit)	
	EndIf	
EndIf

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002SD4  ³ Autor ³ Larson Zordan       ³ Data ³ 23.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Retorna a quantidade empenhada do lote no SD4              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002SD4(ExpC1,ExpC2)                                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Produto (Bobina)                                   ³±±
±±³          ³ ExpC2 = Lote da Bobina                                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpN1 = Quantidade Empenhada                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002SD4(cProduto,cLoteCtl)
Local cQuery := ""
Local nQuant := 0

cQuery := "Select D4_QUANT From " + RetSqlName("SD4") + " "
cQuery += "Where D4_FILIAL  = '" + xFilial("SD4") + "'"
cQuery += "  And D4_COD     = '" + cProduto + "'"
cQuery += "  And D4_LOTECTL = '" + cLoteCtl + "'"
cQuery += "  And D4_NUMOC   = '" + SZ3->Z3_NUM + "'"
cQuery += "  And D_E_L_E_T_ = ' '"
cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"TRBSD4",.F.,.T.)
nQuant := TRBSD4->D4_QUANT
TRBSD4->(dbCloseArea())

Return(nQuant)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002SliOP³ Autor ³ Larson Zordan       ³ Data ³ 08.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Digitar o Peso Produzido de Cada OP                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Sli(ExpN1,ExpA1,ExpC1,ExpN2,ExpN3,ExpL1,ExpO1,ExpA2) ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha do aOps                                      ³±±
±±³          ³ ExpA1 = Array com os dados das Ops dos Slit                ³±±
±±³          ³ ExpC1 = Lote Selecionado                                   ³±±
±±³          ³ ExpN2 = Peso da Bobina                                     ³±±
±±³          ³ ExpN3 = Peso Verificado da Bobina                          ³±±
±±³          ³ ExpL1 = Flag para edicao da ComboBox dos Lotes (p/ref)     ³±±
±±³          ³ ExpO1 = Objeto do ListBox das Ops              (p/ref)     ³±±
±±³          ³ ExpA2 = Array com as Medidas dos Slits                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002SliOP(nX,aOps,cLote,nPeso,nPVer,lEd,oLbx1,aMedidas)
Local lRet   := .T.
Local lPVer  := .F.
Local lZero  := .T.
Local nSobra := Round(( SZ3->Z3_SOBRA * If(nPVer>0, nPVer, nPeso) ) / Val(SZ3->Z3_LARGREA),3)
Local nTProd := 0
Local nZ 

While lRet
	lEditCell( aOps, oLbx1 , "@E 999.999", 6 )
	oLbx1:SetFocus()
	                    
	If aOps[nX,6] == 0 .And. lPVer 
		Exit
	EndIf
	
	//--> Se o Peso Prod For Diferente do Peso Emp., 
	//--> Recalcula sobre o nPVer (Peso Verificado)
	If aOps[nX,6] > 0 .And. aOps[nX,6] <> aOps[nX,4] .And. nPVer == 0
		If nPVer <= 0
			Aviso("ATENCAO !","O Peso Prod. esta diferente do Peso Emp. Corrija o peso ou digite PESO VERIFICADO (Peso Real da Bobina)."+Chr(10)+Chr(10)+"Deixe o valor com zero e altere o peso.",{"<< Voltar"})
			lPVer := .T.
			Loop
		EndIf
	EndIf

	//--> Formula: (( Qt Rolos * Largura do Slit) * Peso da Bobina) / Largura Real da Bobina
	aOps[nX,6] := Round(( (aMedidas[nX,4] * aMedidas[nX,5]) * If(nPVer>0, nPVer, nPeso) ) / Val(SZ3->Z3_LARGREA),3)
	
	lRet := .F.

EndDo

//--> Se for o ultimo item a apontar, verifica o total produzido com o peso verificado
If nPVer > 0
	For nZ := 1 To Len(aOps)
		nTProd += aOps[nZ,6]
		lZero  := (aOps[nZ,6] == 0)
		If lZero
			Exit
		EndIf
	Next nZ
	If !lZero
		aOps[nX,6] += ((nPVer-nSobra)-nTProd) 
	EndIf	
EndIf
       
//--> Verifica se existe alguma qtd digita para bloquear a ComboBox
For nZ := 1 To Len(aOps)
	If aOps[nZ,6] > 0
		lEd := .F.
		Exit
	Else
		lEd := .T.
	EndIf
Next nZ	

Return(aOps)
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Perda³ Autor ³ Larson Zordan       ³ Data ³ 09.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Monta array com a perda                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Perda(ExpA1,ExpC1,ExpA1)                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Array com os dados para a perda                    ³±±
±±³          ³ ExpC1 = Lote Selecionado                                   ³±±
±±³          ³ ExpA1 = Array com os Dados das Bobinas                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002Perda(aPerda,cLoteCtl,aBobinas)
Local aArea := GetArea()
Local cEndereco := Space(15)
Local nSaldo    := 0
Local nX

aPerda := {}

//--> Verifica o Endereco da Bobina
DbSelectArea("SZ3")
If Localiza(SZ3->Z3_PRODUTO)
	cEndereco := Posicione("SB1",1,xFilial("SB1")+SZ3->Z3_PRODUTO,"B1_ENDEREC")
EndIf	

//--> Monta Array Com os Dados Para Perda
nX := aScan(aBobinas,{ |x| x[1] == cLoteCtl })
nSaldo := aBobinas[nX,6]

If nSaldo > 0
	aAdd(aPerda,{ SZ3->Z3_PRODUTO, cLoteCtl, cEndereco, SZ3->Z3_SOBRA, nSaldo, Space(15), Space(6), aBobinas[nX,3] })
Else
	aAdd(aPerda,{ Space(15), Space(10), Space(15), 0, 0, Space(15), Space(6), CToD("  /  /  ") })
EndIf	

RestArea(aArea)

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PrdLot³ Autor ³ Larson Zordan       ³ Data ³ 12.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Digitar o Produto Destino da Perda do Lote                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002PrdLot(ExpN1,ExpO1,ExpA1,ExpN1,ExpN2)                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Linha do aPerda                                    ³±±
±±³          ³ ExpO1 = Objeto do ListBox da Perda                         ³±±
±±³          ³ ExpA1 = Array com os dados da Perda do Lote                ³±±
±±³          ³ ExpN1 = Peso do Lote                                       ³±±
±±³          ³ ExpN2 = Peso do Lote Verificado                            ³±±
±±³          ³ ExpN3 = Posicao do aPerda do Produto Destino               ³±±
±±³          ³ ExpN4 = Posicoa do aPerda do Motivo                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002PrdLot(nZ,oLbx2,aPerda,nPeso,nPVer,nPrdDest,nMotivo)
Local aTemps := {}
Local aRet   := {1}
Local aSucas 
Local lRet   := .F.
Local nX

If SZ3->Z3_SOBRA > 0
	//--> Monta Array das Sucatas
	SB1->(dbSetOrder(4))
	SB1->(dbSeek(xFilial("SB1")+"80  "))
	While !SB1->(Eof()) .And. SB1->(B1_FILIAL+B1_GRUPO) == (xFilial("SB1")+"80  ")
		aAdd(aTemps,{AllTrim(SB1->B1_COD) + " - " + SB1->B1_DESC})
		SB1->(dbSkip())
	EndDo
	aSucas := Array(Len(aTemps))
	For nX:= 1 To Len(aSucas)
		aSucas[nX] := aTemps[nX,1]
	Next nX
	While !lRet
		lRet := ParamBox({{3,"Qual o Produto Destino Para a Perda Deste Lote",aRet[1],aSucas,120,"",.F.}}, "Apontamento de Perda do Lote " + aPerda[nZ,2], aRet)
	EndDo
//	If xFilial("SF2") == "01"
//		aPerda[nZ,nPrdDest] := Left(aSucas[aRet[1]],9) 
//	Else
		aPerda[nZ,nPrdDest] := Left(aSucas[aRet[1]],13)
//	Endif	
	oLbx2:Refresh()
	
	aTemps := {}
	aRet   := {1}
	lRet   := .F.

	//--> Monta Array dos Motivos de Perda (Tabela 43)
	SX5->(dbSeek(xFilial("SX5")+"43"))
	While !SX5->(Eof()) .And. SX5->(X5_FILIAL+X5_TABELA) == (xFilial("SX5")+"43")
		aAdd(aTemps,{Left(SX5->X5_CHAVE,3) + " - " + SX5->X5_DESCRI})
		SX5->(dbSkip())
	EndDo
	aMotivo := Array(Len(aTemps))
	For nX:= 1 To Len(aMotivo)
		aMotivo[nX] := aTemps[nX,1]
	Next nX
	While !lRet
		lRet := ParamBox({{3,"Qual o Motivo da Perda Deste Lote",aRet[1],aMotivo,120,"",.F.}}, "Apontamento de Perda do Lote " + aPerda[nZ,2], aRet)
	EndDo
	aPerda[nZ,nMotivo] := aMotivo[aRet[1]]

EndIf	
Return(aPerda)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PrdVld³ Autor ³ Larson Zordan       ³ Data ³ 12.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Valida o Apontamento do Slit e da Perda                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002PrdLot(aOps,aPerda)                                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Dados dos Slits Apontados                          ³±±
±±³          ³ ExpA2 = Dados da  Perda Apontada                           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002PrdVld(aOps,aPerda)
Local lRet   := .T.
Local nX

//--> Verifica os Slits Produzidos
For nX := 1 To Len(aOps)
	If aOps[nX,6] == 0
		lRet := .F.
		Aviso("ATENCAO !","Existe Slit que NAO foi Apontado !",{"<< Voltar"})
	EndIf
Next nX

//--> Verifica a Perda Apontada
If lRet
	If aPerda[1,4] > 0
		If Empty(aPerda[1,6])
			lRet := .F.
			Aviso("ATENCAO !","Favor Informar o Produto Destino Para a Perda do Lote !",{"<< Voltar"})
		EndIf
		If lRet .And. Empty(aPerda[1,7])
			lRet := .F.
			Aviso("ATENCAO !","Favor Informa o Motivo da Perda do Lote !",{"<< Voltar"})
		EndIf
	EndIf
EndIf

Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002ApntOP³ Autor ³ Larson Zordan       ³ Data ³ 12.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Apontamento de Producao (Rotina Automatica)                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002ApntOP(ExpL1,ExpA1,ExpA2,ExpL2)                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Variavel logica da Funcao Processa                 ³±±
±±³          ³ ExpA1 = Dados das OPs a ser apontada                       ³±±
±±³          ³ ExpA2 = Dados das Perdas dos Lotes                         ³±±
±±³          ³ ExpL2 = .T.  se houver erro na rotina automatica           ³±±
±±³          ³ ExpC1 = Centro de Custo (Recurso)                          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002ApntOP(lEnd,aOps,aPerda,lErro,cCC)
Local aCabAP  := {}
Local aItemAP := {}
Local aItem   := {}
Local cLote
Local cNumOP
Local nCusto
Local nQuant
Local nX

lMsErroAuto   := .F.

SB1->(dbSetOrder(1))
SD4->(dbSetOrder(1))

ProcRegua(Len(aOps))

//--> Apontando as OPs Pela Rotina Automatica             
For nX := 1 To Len(aOps)
	IncProc()
	//--> OP do Sliter
	cNumOP := aOps[nX,8]
	
	//--> Sliter      
	SB1->(dbSeek(xFilial("SD3")+aOps[nX,1]+"01"))
	
	//--> Quantidade a Produzida
	nQuant := aOps[nX,6]
	
	//--> Lote do Sliter
	cLote  := aOps[nX,9]
	
	//--> Monta Array para a Rotina Automatica
	aOP    := {	{"D3_TM"     	,"300"    				,Nil},;
				{"D3_COD"    	,SB1->B1_COD			,Nil},;
				{"D3_UM"		,SB1->B1_UM				,Nil},;
				{"D3_QUANT"		,nQuant					,Nil},;
				{"D3_CONTA"  	,CriaVar("D3_CONTA")	,Nil},;
				{"D3_OP"     	,cNumOP   				,Nil},;
				{"D3_CC"     	,cCC      				,Nil},;
				{"D3_LOCAL"		,SB1->B1_LOCPAD			,Nil},;
				{"D3_SEGUM"		,SB1->B1_SEGUM			,Nil},;
				{"D3_QTSEGUM"	,nQuant					,Nil},;
				{"D3_PERDA"  	,0       				,Nil},;
				{"D3_LOTECTL"	,cLote 					,Nil},;
				{"D3_DTVALID"	,dDataBase+365			,Nil}} 

	MsExecAuto({|x,y| Mata250(x,y)},aOP,3) //Inclusao
	If lMsErroAuto
		lErro := .T.
		Aviso("ATENCAO !","Ocorreu um Erro ao Apontar a OP " + cNumOP,{"Sair >>"})
		If lViewRot
			MostraErro()
		EndIf	
	Else
		//--> Grava o Empenho do Slit no Lote (SB8)
		If AllTrim(SB1->B1_GRUPO) == "20"
			dbSelectArea("SB8")
			dbSetOrder(3)
			If dbSeek(xFilial("SB8")+SB1->B1_COD+SB1->B1_LOCPAD+cLote)
				RecLock("SB8",.F.)
				Replace B8_EMPENHO With nQuant	,;
						B8_EMPENH2 With nQuant
				MsUnLock()	
			EndIf
			
			//--> Grava o Empenho do Slit no Endereco (SBF)
			If Localiza(SB1->B1_COD)
				dbSelectArea("SBF")
				dbSetOrder(3)
				If dbSeek(xFilial("SBF")+SB1->B1_COD+SB1->B1_LOCPAD+cLote)
					RecLock("SBF",.F.)
					Replace BF_EMPENHO With nQuant	,;
							BF_EMPEN2  With nQuant
					MsUnLock()
				EndIf
			EndIf
		
			//--> Grava o Lote do Slit no SD4
			dbSelectArea("SD4")
			dbSetOrder(1)
			If dbSeek(xFilial("SD4")+SB1->B1_COD+Left(cNumOP,8))
				RecLock("SD4",.F.)
				Replace D4_LOTECTL With cLote,;
						D4_DTVALID With dDataBase+365
				MsUnLock()
			EndIf
			dbSetOrder(4)
		EndIf	
	EndIf
Next nX

//--> Apontando as Perdas das OPs Pela Rotina Automatica             
If !lErro .And. !Empty(aPerda[1,1])
	//Cabecalho
	aCabAP:= {	{"BC_OP"		,cNumOP					,Nil}}

	//Items
	aItem := {	{"BC_PRODUTO"	,aPerda[1,1]			,Nil},;
				{"BC_LOCORIG"	,"01"					,Nil},;
				{"BC_LOCALIZ"	,aPerda[1,3]			,Nil},;
				{"BC_TIPO"  	,"R"					,Nil},;
				{"BC_MOTIVO"	,aPerda[1,7]			,Nil},;
				{"BC_QUANT" 	,aPerda[1,5]			,Nil},;
				{"BC_CODDEST"	,aPerda[1,6]			,Nil},;
				{"BC_LOCAL" 	,"01"					,Nil},;
				{"BC_QTDDEST"	,aPerda[1,5]			,Nil},;
				{"BC_OPERADOR"	,SubStr(cUsuario,7,15)	,Nil},;
				{"BC_LOTECTL"	,aPerda[1,2]			,Nil},;
				{"BC_DTVALID"	,aPerda[1,8]			,Nil} }
	
	aAdd(aItemAP,aItem)
	
	MSExecAuto({|x,y,z| Mata685(x,y,z)},aCabAP,aItemAP,3)
	
	If lMsErroAuto
		Aviso("ATENCAO !","Ocorreu um Erro ao Apontar a Perda da OP " + cNumOP,{"Sair >>"})
		If lViewRot
			MostraErro()
		EndIf	
	EndIf	
EndIf	

//--> Atualizacao do Custo do Slit
For nX := 1 To Len(aOps)
	nCusto := 0
	If SB2->(dbSeek(xFilial("SD2")+aOps[nX,1]+"01"))
		nCusto := SB2->B2_CM1
	EndIf
	dbSelectArea("SB1")
	If dbSeek(xFilial("SB1")+aOps[nX,1]+"01")
		RecLock("SB1",.F.)
		Replace B1_X_CUSTO With nCusto
		MsUnLock()
	EndIf
Next nX

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002View ³ Autor ³ Larson Zordan       ³ Data ³27.01.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Visualizar os Empenhos e as Ordens de Producao da OC       ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002View(ExpN1)                                          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Opcao para visualizar                              ³±±
±±³          ³       1 = Visualizar os Empenhos                           ³±±
±±³          ³       2 = Visualizar as Ordens de Producao                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002View(nOpc)
Local oDlg
Local aAreaAnt  := GetArea()
Local aDados    := {}
Local cNumOP    := ""

dbSelectArea("SC2")
dbSetOrder(10)
If dbSeek(xFilial("SC2")+SZ3->Z3_NUM)
	While !Eof() .And. xFilial("SC2")+SZ3->Z3_NUM == C2_FILIAL+C2_NUMOC
		cNumOP := C2_NUM+C2_ITEM+C2_SEQUEN+C2_ITEMGRD
		If nOpc == 1
			dbSelectArea("SD4")
			dbSetOrder(2)
			If dbSeek(xFilial("SD4")+cNumOP)
				While !Eof() .And. xFilial("SD4")+cNumOP == D4_FILIAL+D4_OP
					aAdd(aDados,{D4_COD, D4_LOCAL, D4_QUANT, D4_LOTECTL, D4_DTVALID})
					If Empty(D4_NUMOC)
						RecLock("SD4",.F.)
						Replace D4_NUMOC With SZ3->Z3_NUM
						MsUnLock()
					EndIf
					dbSkip()
				EndDo
			EndIf
			dbSelectArea("SC2")
		Else
			aAdd(aDados,{cNumOP, C2_PRODUTO, C2_LOCAL, C2_QUANT, C2_UM})
		EndIf
		dbSkip()
	EndDo
EndIf

If Len(aDados) > 0
	Define MsDialog oDlg Title "Visualiza "+If(nOpc==1,"os Empenhos","as Ordens de Producao")+" Da Ordem de Corte" From 0,0 To 200,450 Of oMainWnd Pixel 
	
	If nOpc == 1
		@ 01,01 ListBox oLbx Fields Header "Produto","Amz","Quant.","Lote","Dt. Validade" ;
		Size 224, 88 Of oDlg Pixel
		oLbx:SetArray(aDados)
		oLbx:bLine := {|| {aDados[oLbx:nAt,1], aDados[oLbx:nAt,2], Transform(aDados[oLbx:nAt,3],"@E 999,999.999"), aDados[oLbx:nAt,4], DtoC(aDados[oLbx:nAt,5])} }
	Else
		@ 01,01 ListBox oLbx Fields Header "Ordem Producao","Produto","Amz","Quant.","UM" ;
		Size 224, 88 Of oDlg Pixel
		oLbx:SetArray(aDados)
		oLbx:bLine := {|| {aDados[oLbx:nAt,1], aDados[oLbx:nAt,2], aDados[oLbx:nAt,3], Transform(aDados[oLbx:nAt,4],"@E 999,999.999"), aDados[oLbx:nAt,5]} }
	EndIf
	@  90,170  Button "<< Voltar"  Size 45,11 Of oDlg Pixel Action (oDlg:End())
	
	Activate MsDialog oDlg Center
EndIf
RestArea(aAreaAnt)
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Tubo ³ Autor ³ Larson Zordan       ³ Data ³ 29.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Apontamento de Tubos da Ordem de Corte                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002Tubo(ExpC1,ExpN1,ExpN2)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Tubo(cAlias,nReg,nOpc)
Local oBmp
Local oBtn1
Local oBtn2
Local oBtn3
Local oDlg
Local oFnt1
Local oFnt2
Local oFnt3

Local aAreaAnt := GetArea()
Local aDados   := {Space(06),Space(15),Space(20),0,"  ",0,"  ",Space(10),dDataBase,Space(06),Space(30)}
Local aPerda   := {}
Local aOps     := {}
Local cTitulo  := "Apontamento do TUBO da Ordem de Corte"
Local cHoraIni := "07:00"
Local cHoraEnt := Left(Time(),5)
Local dDataIni := dDataBase
Local dDataEnt := dDataBase

Local lRet     := .T.
Local lErro    := .F.

Private cOP	   := CriaVar("D3_OP",.T.)
Private lTubo  := !Empty(cAlias)

If !lTubo
	cTitulo  := "Apontamento do Ordem de Produção COMAFAL"
EndIf

/*--------------------------------
| aDados[ 1] = Tipo              |
| aDados[ 2] = Produto           |
| aDados[ 3] = Descricao         |
| aDados[ 4] = Quant             |
| aDados[ 5] = UM                |
| aDados[ 6] = Quant 2a UM       |
| aDados[ 7] = 2a UM             |
| aDados[ 8] = Lote              |
| aDados[ 9] = Data de Validade  |
| aDados[10] = Centro de Custo   |
| aDados[11] = Desc. do CC       |
--------------------------------*/

Set Key VK_F4 To U_A002ShwOP(.T.)

dbSelectArea("SC2")
dbSetOrder(1)
      
Define FONT oFnt1 NAME "San Serif" 	Size 13,16 BOLD 
Define FONT oFnt2 NAME "San Serif" 	Size 10,10 BOLD 
Define FONT oFnt3 NAME "San Serif" 	Size  4,12     

Define MsDialog oDlg Title cTitulo From 0,0 To 400,300 Of oMainWnd Pixel

@   1,  5 BitMap oBmp File "COMAFAL.BMP" 		Size  60,45 Of oDlg Pixel NoBorder

@   1, 65 BitMap oBmp File "APONTAOP.BMP" 		Size  80,60 Of oDlg Pixel NoBorder

@  47, 5 To 181, 146 Of oDlg Pixel

@  53, 12 Say "Ordem Produção" 						Size  44,09 Of oDlg Pixel Color CLR_HBLUE
@  60, 12 MsGet cOP Picture "@!" 					Size  44,09 Of oDlg Pixel Valid( CAI002CpoOP(@aDados,@oDlg,.T.,.T.) ) 

/**********************************
* Five Solutions Consultoria
* Itacolomy Mariano
* 02/01/2008 - 13:44 hs
* Comentario: Comentado uso do campo Lote neste processo devido a descontinuidade da utilização do recurso de
*             rastreabilidade(Lote) para produtos acabados(TIPO = 'PA')
*  
* @  53, 58 Say "Lote"								Size  30,09 Of oDlg Pixel
* @  60, 58 MsGet aDados[8]							Size  30,09 Of oDlg Pixel When .F.
* @  53,105 Say "Dt. Validade"						Size  35,09 Of oDlg Pixel
* @  60,105 MsGet aDados[9]							Size  35,09 Of oDlg Pixel When .F.
*
*
***********************************/

@  73, 12 Say "Produto" 							Size  44,09 Of oDlg Pixel
@  80, 12 MsGet aDados[2] 							Size  44,09 Of oDlg Pixel When .F. Center 
@  80, 58 MsGet aDados[3] 							Size  82,09 Of oDlg Pixel When .F.

@  93, 12 Say "Quantidade" 							Size  44,09 Of oDlg Pixel Color CLR_HBLUE
@ 100, 12 MsGet aDados[4] Picture "@E 9999.999"	Size  44,09 Of oDlg Pixel Valid( If(aDados[4]==0,(Help(" ",1,"QUANTVAZIO"),.F.),.T.) )

@  93, 58 Say "UM" 									Size  10,09 Of oDlg Pixel
@ 100, 58 MsGet aDados[5] 							Size  10,09 Of oDlg Pixel When .F. Center 

@  93, 77 Say "Qtd. Embalagem"						Size  40,09 Of oDlg Pixel
@ 100, 77 MsGet aDados[6] Picture "@E 999,999,999"	Size  40,09 Of oDlg Pixel When .F.

@ 113, 12 Say "Data Início"							Size  33,09 Of oDlg Pixel
@ 120, 12 MsGet dDataIni	 						Size  33,09 Of oDlg Pixel Center Valid( A002Data(dDataIni,dDataEnt,cHoraIni,cHoraEnt) )

@ 113, 46 Say "Hora"       							Size  20,09 Of oDlg Pixel
@ 120, 46 MsGet cHoraIni Picture "99:99"			Size   5,09 Of oDlg Pixel Center Valid( A002Hora(dDataIni,dDataEnt,cHoraIni,cHoraEnt,"I") )

@ 113, 80 Say "Data Apont."							Size  33,09 Of oDlg Pixel
@ 120, 80 MsGet dDataEnt	 						Size  33,09 Of oDlg Pixel Center Valid( A002Data(dDataIni,dDataEnt,cHoraIni,cHoraEnt) )

@ 113,113 Say "Hora"       							Size  20,09 Of oDlg Pixel
@ 120,113 MsGet cHoraEnt Picture "99:99"			Size   5,09 Of oDlg Pixel Center Valid( A002Hora(dDataIni,dDataEnt,cHoraIni,cHoraEnt,"F") )

@ 132, 12 Say "C. Centro"							Size  30,09 Of oDlg Pixel
@ 139, 12 MsGet aDados[10] F3 "CTT"				Size  40,09 Of oDlg Pixel When .T.
@ 139, 55 MsGet SI3->I3_DESC        				Size  85,09 Of oDlg Pixel When .F.

/**********************************
* Five Solutions Consultoria
* Itacolomy Mariano
* 02/01/2008 - 13:44 hs
* Comentario: Comentado uso do campo Lote neste processo devido a descontinuidade da utilização do recurso de
*             rastreabilidade(Lote) para produtos acabados(TIPO = 'PA')
*  
* @ 172,25 Say AllTrim(aDados[2])+aDados[8] 		Size 100,09 Of oDlg Pixel Center Font oFnt2 
*
***********************************/
@ 172,25 Say AllTrim(aDados[2])         		Size 100,09 Of oDlg Pixel Center Font oFnt2 


@ 194,05 Say "[F4] Exibir OPs"						Size  80,09 Of oDlg Pixel Font oFnt3 Color CLR_HRED

Define SButton oBtn2 From 185, 85 Type  1 Enable Of oDlg OnStop "Confirma Apontamento" Action oDlg:End()
Define SButton oBtn3 From 185,115 Type  2 Enable Of oDlg OnStop "Cancela Apontamento"  Action (lRet:=.F.,oDlg:End())

Activate MsDialog oDlg Center

If lRet .And. !Empty(cOP)
	Begin Transaction
		//--> Apontamento de Producao do Tudo
		aAdd(aOps,{	aDados[2]						,;	// Tubo
					aDados[3]						,;	// Descricao
					0								,;	// Qt Rolos
					aDados[4]						,;	// Peso Empenhado
					""								,;	// Prod.Acabado
					aDados[4]						,;	// Qtde Produzida
					.F.								,;	// Flag para Marcar
					cOP  							,;	// OP do Tudo
					""						        })	// Lote do Tubo
/**********************************
* Five Solutions Consultoria
* Itacolomy Mariano
* 02/01/2008 - 13:44 hs
* Comentario: Comentado uso do campo Lote neste processo devido a descontinuidade da utilização do recurso de
*             rastreabilidade(Lote) para produtos acabados(TIPO = 'PA')
*  
*					aDados[8]						})	// Lote do Tubo
***********************************/

		aAdd(aPerda,{ Space(15), Space(10), Space(15), 0, 0, Space(15), Space(6), CToD("  /  /  ") })

		Processa( { |lEnd| CAI002ApntOP(@lEnd,aOps,aPerda,lErro,aDados[10]) },"Aguarde...","Processando Apontamento da Producao de Tubos...")
		
		//--> Grava Movimento da Producao (SH6) - Sera gravado apenas para Produto Acabado (OP Pai)
		RecLock("SH6",.T.)
		Replace SH6->H6_FILIAL		With xFilial("SH6")							,;
				SH6->H6_OP			With cOP									,;
				SH6->H6_PRODUTO		With aDados[2]								,;
				SH6->H6_RECURSO		With Right(AllTrim(aDados[10]),3)			,;
				SH6->H6_DATAINI		With dDataIni								,;
				SH6->H6_HORAINI		With cHoraIni								,;
				SH6->H6_DATAFIN		With dDataEnt								,;
				SH6->H6_HORAFIN		With cHoraEnt								,;
				SH6->H6_QTDPROD		With aDados[4]								,;
				SH6->H6_PT			With "T"									,;
				SH6->H6_DTAPONT		With dDataBase								,;
				SH6->H6_TEMPO		With ElapTime(cHoraIni+":00",cHoraEnt+":00"),;
				SH6->H6_OPERADO		With SubStr(cUsuario,7,15)					,;
				SH6->H6_TIPOTEM		With 1
				//SH6->H6_TEMPO		With ElapTime(cHoraIni+":00",cHoraEnt+":00"),;
/**********************************
* Five Solutions Consultoria
* Itacolomy Mariano
* 02/01/2008 - 13:44 hs
* Comentario: Comentado uso do campo Lote neste processo devido a descontinuidade da utilização do recurso de
*             rastreabilidade(Lote) para produtos acabados(TIPO = 'PA')
*  
*				SH6->H6_LOTECTL		With aDados[8]								,;
*				SH6->H6_DTVALID	With aDados[9]								,;
*
***********************************/
		MsUnLock()
		
	End Transaction
EndIf

Set Key VK_F4 To 

RestArea(aAreaAnt)

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002CpoOP³ Autor ³ Larson Zordan       ³ Data ³ 29.01.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Localiza a OP                                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002CpoOP(ExpA1,ExpO,ExpL1)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Dados da OP                                        ³±±
±±³          ³ ExpO1 = Objeto da Dialog                                   ³±±
±±³          ³ ExpL1 = Exibir o BitMap do Codigo de Barras                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002CpoOP(aDados,oDlg,lBmp,lApontaOP)
Local oBmp
Local aLote := Array(2)
Local cTipo := ""
Local lRet  := .T. 

If !dbSeek(xFilial("SC2")+cOP)
	Help(" ",1,"REGNOIS")	
	lRet := .F.
Else
	If lTubo .And. lApontaOP
		lRet := !Empty(SC2->C2_NUMOC)
	EndIf
EndIf
If lRet
	l250 := .T.
	U_CAI001(2,@aLote,SC2->C2_PRODUTO)
	l250 := .F.
	SI3->(dbSeek(xFilial("SI3")+SC2->C2_CC))
	cTipo		:= Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_GRUPO")
	aDados[ 1]	:= Left( Posicione("SBM",1,xFilial("SBM")+cTipo,"BM_DESC") ,6)
	aDados[ 2]	:= SC2->C2_PRODUTO
	aDados[ 3] 	:= Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_DESC")
	aDados[ 4] 	:= (SC2->C2_QUANT-SC2->C2_QUJE-SC2->C2_PERDA)
	aDados[ 5] 	:= SC2->C2_UM
	aDados[ 6] 	:= Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_QE")
	aDados[ 7] 	:= Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_SEGUM")
	aDados[ 8] 	:= aLote[1]
	aDados[ 9] 	:= aLote[2]
	aDados[10] 	:= SC2->C2_CC
	aDados[11] 	:= SI3->I3_DESC
	If lBmp
		@ 152,12 BitMap oBmp File "TBBARRAP.BMP" Size 150,80 Of oDlg Pixel NoBorder
	EndIf	
	oDlg:Refresh()
Else
	If lTubo
		Aviso("ATENCAO !","Esta Ordem de Produção NÃO Pertence a Nenhuma Ordem de Corte !",{ " << Voltar "})
	Else
		Aviso("ATENCAO !","Ordem de Produção NÃO Encontrada !",{ " << Voltar "})
	EndIf	
EndIf

Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ A002ShwOP  ³ Autor ³ Larson Zordan       ³ Data ³ 04.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Exibe e Pesquisa as OPs Geradas Pela Ordem de Corte        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ A002ShwOP(ExpL1)                                           ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 - Informa se ira Apontar a OP ou Estornar a OP       ³±±
±±³          ³         .T. - Apontar a OP                                 ³±±
±±³          ³         .F. - Estorna a OP                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function A002ShwOP(lApontaOP)
Local oDlg
Local oBtn
Local oLbx
Local aOPs    := {}
Local aCbx    := {"Ordem Producao","Produto","Ordem Corte"}
Local aPosLbx := { 1, 2, 5 }
Local cTitulo := "Ordens de Producao"

Local lRet  := .T.

MsgRun("Aguarde... Consultando as Ordens de Produção...",,{ || CAI002MntOP(@aOPs,lApontaOP) } )

If lTubo
	cTitulo += " Geradas Pelas Ordens de Corte"
EndIf                  

Define MsDialog oDlg Title cTitulo From 0,0 To 300,550 Of oMainWnd Pixel

@ 1,1 ListBox oLbx Fields Header "Ord. Producao","Produto","Descricao","Quant.","Ordem Corte" ;
	ColSizes 45,45,100,20 Size 275,130 Of oDlg Pixel

oLbx:SetArray(aOPs)
oLbx:bLine := {|| {aOPs[oLbx:nAt,1], aOPs[oLbx:nAT,2], aOPs[oLbx:nAt,3], Transform(aOPs[oLbx:nAT,4],"@E 99,999.999"), aOPs[oLbx:nAT,5] } }

@ 136,05 Button " Pesquisar " Action U_PesqLbx(@oLbx,aOPs,aCbx,aPosLbx,.T.) Size 44,11 Of oDlg Pixel

Define SButton oBtn From 135,180 Type 15 Enable Of oDlg OnStop "Outras Informações"
Define SButton oBtn From 135,210 Type  1 Enable Of oDlg Action (cOP :=aOPs[oLbx:nAt,1],oDlg:End())
Define SButton oBtn From 135,240 Type  2 Enable Of oDlg Action (lRet:=.F.,oDlg:End())

Activate MsDialog oDlg Center

Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002MntOP ³ Autor ³ Larson Zordan       ³ Data ³ 04.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Monta array com os dados para consulta das OPs             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002MntOP(ExpA1,ExpL1)                                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 - Array para atualizar os dados das OPs (por ref.)   ³±±
±±³          ³ ExpL1 - Informe se ira apontar ou estornar a OP            ³±±
±±³          ³         .T. - Apontar a OP                                 ³±±
±±³          ³         .F. - Estornar a OP                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002MntOP(aOPs,lApontaOP)
Local aAreaAnt  := GetArea()
Local cAliasTRB := CriaTrab(Nil,.F.)
Local aStrucSC2	:= SC2->(dbStruct())
Local cQry      := ""
Local nX

cQry := "Select * From " + RetSqlName("SC2") + " Where "
cQry += "    C2_FILIAL = '" + xFilial("SC2") + "' "
If lTubo
	cQry += "And C2_NUMOC <> '          ' "
	cQry += "And SUBSTRING(C2_PRODUTO,1,2) > '20' "
EndIf
/*If lApontaOP
	cQry += "And C2_QUANT  > (C2_QUJE+C2_PERDA) "
Else
	cQry += "And C2_QUJE   > 0 "
EndIf	*/

cQry += "And C2_DATRF = '          ' "
cQry += "Order By C2_NUM,C2_ITEM,C2_SEQUEN,C2_NUMOC"
cQry := ChangeQuery(cQry)
dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQry),cAliasTRB,.F.,.T.)
For nX := 1 To Len(aStrucSC2)
	If aStrucSC2[nX][2] <> "C" .And. FieldPos(aStrucSC2[nX][1]) <> 0
		TcSetField(cAliasTRB,aStrucSC2[nX][1],aStrucSC2[nX][2],aStrucSC2[nX][3],aStrucSC2[nX][4])
	EndIf
Next nX

dbSelectArea(cAliasTRB)
While !(cAliasTRB)->(Eof())
    aAdd( aOps, {	C2_NUM+C2_ITEM+C2_SEQUEN+C2_ITEMGRD										,;
    				C2_PRODUTO																,;
    				Posicione("SB1",1,xFilial("SB1")+(cAliasTRB)->C2_PRODUTO,"B1_DESC")	,;
    				C2_QUANT																,;
    				C2_NUMOC 																})
    				//If(lApontaOP,C2_QUANT,C2_QUJE)											,;
	(cAliasTRB)->(dbSkip())
EndDo
(cAliasTRB)->(dbCloseArea())
RestArea(aAreaAnt)
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002Agen³ Autor ³ Larson Zordan        ³ Data ³ 06.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Agendamento das Ordens de Corte                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002Agen(ExpC1,ExpN1,ExpN2)                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002Agen(cAlias,nReg,nOpc)
Local oDlg
Local oLbx
Local aDados 	:= {}
Local aDias     := {}
Local cPerg  	:= "CAI002"
Local cTitulo   := ""
Local nX

Local aInfo	 	:= {}
Local aObjects	:= {}
Local aPosObj 	:= {}
Local aSize	 	:= {}

If Empty(cAlias)
	Private aRotina    := {	{ "","",0,1},;
	                     	{ OemToAnsi("Produção/Dia" ) ,"U_CAI002Agen", 0 , 2}}
	nOpc    := 2
	cTitulo := OemToAnsi("Produção Por Dia - ")
EndIf

aSize := MsAdvSize()
aInfo := {aSize[1],aSize[2],aSize[3],aSize[4],3,3}

aAdd(aObjects,{100,100,.T.,.T.})
aPosObj := MsObjSize(aInfo, aObjects)

PutSX1(cPerg,"01","Selecione a Data a Considerar ","","","mv_ch1","N",01,0,0,"C","","","","","mv_par01","Emissão","","","","Prev. Início","","","Prev. Entrega","","","","","","","","",{"Selecionar a data a ser considerada na ","visualização da produção por dia.       "},{},{})
PutSX1(cPerg,"02","Da Data                       ","","","mv_ch2","D",08,0,0,"G","","","","","mv_par02","","","","","","","","","","","","","","","","",{"Data inicial a ser considerado na visua","lizacao da produção por dia.            "},{},{})
PutSX1(cPerg,"03","Até a Data                    ","","","mv_ch3","D",08,0,0,"G","","","","","mv_par03","","","","","","","","","","","","","","","","",{"Data final a ser considerado na visuali","zacao do produção por dia.              "},{},{})
PutSX1(cPerg,"04","Selecione a Qtde. a Considerar","","","mv_ch4","N",01,0,0,"C","","","","","mv_par04","Qtd. à Produzir","","","","Qtd. Produzida","","","","","","","","","","","",{"Selecionar a Qtde a ser considerada na ","visualização da produção por dia.       "},{},{})
If !Pergunte(cPerg,.T.,OemToAnsi("Produção Por Dia"))
	Return(.T.)
EndIf

Processa( {|lEnd| CAI002PrcOP(@lEnd,@aDados,@aDias)},"Aguarde... ",OemToAnsi("Processando as Produções Por Dia..."))

If (Len(aDados)-1) == 0
	Aviso(OemToAnsi("ATENÇÃO !"),OemToAnsi("Não Há Ordens de Produção no Período Informado."),{" Sair >> "})
	Return
EndIf

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Montagem de um aHeader fixo para inclusao dos saldos         ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
aHeader := {}
aCols   := {}

//--> Cria aHeader e aCols da GetDados.
dbSelectArea("SX3")
dbSetOrder(2)
For nX := 1 To (Len(aDias)-1)
	If     nX == 1
		dbSeek("C2_PRODUTO")
	ElseIf nX == (Len(aDias)-1)
		dbSeek("B1_DESC")
	ElseIf nX == (Len(aDias)-2)
		dbSeek("C2_PERDA")
	ElseIf nX == (Len(aDias)-3)
		dbSeek("C2_QUJE")
	Else
		dbSeek("C2_QUANT")
	EndIf
	aAdd(aHeader,{	aDias[nX]		,;
					SX3->X3_CAMPO	,;
					SX3->X3_PICTURE	,;
					If(nX==1,9,SX3->X3_TAMANHO	),;
					SX3->X3_DECIMAL	,;
					SX3->X3_VALID  	,;
					SX3->X3_USADO	,;
					SX3->X3_TIPO	,;
					SX3->X3_ARQUIVO	,;
					SX3->X3_CONTEXT	})
Next nX

//--> Cria aCols Com Base no aDados
aCols := aClone(aDados)

If     mv_par01 == 1
	cTitulo += OemToAnsi("Data da Emissão")
ElseIf mv_par01 == 2
	cTitulo += OemToAnsi("Data da Previsão de Início")
ElseIf mv_par01 == 3
	cTitulo += OemToAnsi("Data da Previsão de Entrega")
EndIf

If mv_par04 == 1
	cTitulo += OemToAnsi(" - Por Quantidade À Produzir")
Else
	cTitulo += OemToAnsi(" - Por Quantidade Produzida")
EndIf

Define MsDialog oDlg Title cTitulo From aSize[7],0 To aSize[6],aSize[5] Of oMainWnd Pixel

oGet := MSGetDados():New(aPosObj[1,1],aPosObj[1,2],aPosObj[1,3],aPosObj[1,4],2,"AllwaysTrue","AllwaysTrue","",.F.)
oGet:oBrowse:nFreeze := 1

Activate MsDialog oDlg Center On Init EnchoiceBar(oDlg,{|| oDlg:End() },{||oDlg:End()})

Return
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PrcOP³ Autor ³ Larson Zordan        ³ Data ³ 06.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa as Contagens das Ordens de Producao               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002PrcAg(ExpL1,ExpA1)                                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpL1 = Variavel da funcao Processa()                      ³±±
±±³          ³ ExpA1 = Dados do Agendamento                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function CAI002PrcOP(lEnd,aDados,aDias)

Local aStrucTRB	:= SC2->(dbStruct())
Local aTotais   := {}
Local nDias     := 1
Local nTot      := 0
Local nFalta    := 0
Local nPerda    := 0
Local nDescr    := 0
Local cAliasTRB := CriaTrab(Nil,.F.)
Local cQuery    := "" 
Local nX

If     mv_par01 == 1
	cQuery := "Select C2_EMISSAO As DDATA, "
ElseIf mv_par01 == 2
	cQuery := "Select C2_DATPRI  As DDATA, "
ElseIf mv_par01 == 3
	cQuery := "Select C2_DATPRF  As DDATA, "
EndIf
cQuery += "C2_PRODUTO As PRODUTO, "

If mv_par04 == 1
	cQuery += "Sum(C2_QUANT)-Sum(C2_QUJE) As PESO, Sum(C2_QUJE) As FALTA,"
Else
	cQuery += "Sum(C2_QUJE)  As PESO, (Sum(C2_QUANT)-Sum(C2_QUJE)) As FALTA,"
EndIf	

cQuery += " Sum(C2_PERDA) As PERDA "
cQuery += "From " + RetSqlName("SC2") + " Where "
cQuery += " D_E_L_E_T_ = ' ' And C2_FILIAL = '" + xFilial("SC2") + "'"
If     mv_par01 == 1
	cQuery += " And C2_EMISSAO >= '" + DtoS(mv_par02) + "' And C2_EMISSAO <= '" + DtoS(mv_par03) + "' "
    cQuery += " Group By C2_EMISSAO, C2_PRODUTO "
ElseIf mv_par01 == 2
	cQuery += " And C2_DATPRI >= '" + DtoS(mv_par02) + "' And C2_DATPRI <= '" + DtoS(mv_par03) + "' "
    cQuery += " Group By C2_DATPRI , C2_PRODUTO "
ElseIf mv_par01 == 3
	cQuery += " And C2_DATPRF >= '" + DtoS(mv_par02) + "' And C2_DATPRF <= '" + DtoS(mv_par03) + "' "
    cQuery += " Group By C2_DATPRF , C2_PRODUTO "
EndIf
cQuery += "Order By DDATA, PRODUTO"

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTRB,.F.,.T.)},"Aguarde...","Processando Dados das OPs...")
For nX := 1 To Len(aStrucTRB)
	If aStrucTRB[nX][2] <> "C" .And. FieldPos(aStrucTRB[nX][1]) <> 0
		TcSetField(cAliasTRB,aStrucTRB[nX][1],aStrucTRB[nX][2],aStrucTRB[nX][3],aStrucTRB[nX][4])
	EndIf
Next nX

nDias    	:= mv_par03 - mv_par02
aDias    	:= Array(nDias+7)
aDias[1]	:= "Produto"
aDias[2]	:= DtoC(mv_par02)

For nX := 3 To Len(aDias)
	aDias[nX] := DtoC(mv_par02+(nX-2))
Next nX

aDias[Len(aDias)-4] := "Total (Ton)"
aDias[Len(aDias)-3] := If(mv_par04==1,"Produzida","À Produzir") + " (Ton)"
aDias[Len(aDias)-2] := "Perda (Ton)"
aDias[Len(aDias)-1] := "Descrição do Produto"
aDias[Len(aDias)  ] := "FLAG"
nTot   := Len(aDias)-4
nFalta := Len(aDias)-3
nPerda := Len(aDias)-2
nDescr := Len(aDias)-1

//--> Cria o aDados Apenas com a Qtde referente ao mv_par04
dbSelectArea(cAliasTRB)
dbGoTop()
While !Eof()
	nPos := aScan(aDados,{ |x| x[1] == PRODUTO } )
	nX   := aScan(aDias,DtoC(StoD(DDATA)))
	If nPos == 0
		aAdd(aDados, Array(Len(aDias)) )
		nPos := Len(aDados)
		aDados[nPos, 1    ] := PRODUTO
		aDados[nPos,nDescr] := Posicione("SB1",1,xFilial("SB1")+(cAliasTRB)->PRODUTO,"B1_DESC")
		aDados[nPos,nX    ] := PESO
		aDados[nPos,nFalta] := FALTA
		aDados[nPos,nPerda] := CAI002PrdOP(PRODUTO,mv_par02,mv_par03)
	Else
		If aDados[nPos,nX] == Nil
		    aDados[nPos,nX    ] := PESO
			aDados[nPos,nFalta] := FALTA
			//aDados[nPOs,nPerda] := PERDA
		Else    
		    aDados[nPos,nX    ] += PESO
			aDados[nPos,nFalta] += FALTA
			//aDados[nPOs,nPerda] += PERDA
		EndIf    
	EndIf
	dbSkip()
EndDo
(cAliasTRB)->(dbCloseArea())

//--> Zerar os Elementos Nil
For nX := 1 To Len(aDados)
	For nZ := 2 To Len(aDados[nX])
		aDados[nX,Len(aDias)] := .F.
		If aDados[nX,nZ] == Nil
			aDados[nX,nZ] := 0
		EndIf	
	Next nZ
Next nX

//--> Totalizar Por Produto
For nX := 1 To Len(aDados)
	For nZ := 2 To (nTot-1)
		aDados[nX,nTot] += aDados[nX,nZ]
	Next nZ
Next nX

//--> Ordena o aDados Por Produto
aDados := aSort(aDados,,,{ |x,y| x[1] < y[1] })

//--> Totalizador aDados
aTotais         := Array(Len(aDias))
aTotais[1]      := "<< TOTAIS >>"
aTotais[nDescr] := "<< TOTAIS >>"
For nX := 2 To nPerda
	aTotais[nX] := 0
Next nX

//--> Totalizar Por Dia
For nX := 1 To Len(aDias)
	If nX > 1 .And. nX < nDescr
		For nZ := 1 To Len(aDados)
			aTotais[nX] += aDados[nZ,nX]
		Next nZ
	EndIf
Next nX
aAdd(aDados,aTotais)

Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002PrdOP³ Autor ³ Larson Zordan        ³ Data ³23.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa as Perdas das Ordens de Producao               	  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ACAI002PrdOP(ExpC1,ExpD1,ExpD2)                            ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Produto                                            ³±±
±±³          ³ ExpD1 = Data Inicial                                       ³±±
±±³          ³ ExpD2 = Data Final                                         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002PrdOP(cProd,dDataIni,dDataFim)
Local aAreaAnt  := GetArea()
Local aStrucTRB	:= SBC->(dbStruct())
Local cAliasSBC := CriaTrab(Nil,.F.)
Local cQuery    := "" 
Local nPerda    := 0
Local nX
cQuery := "Select Sum(BC_QUANT) As PERDA From " + RetSqlName("SBC") + " "
cQuery += "Where BC_FILIAL  = '" + xFilial("SBC") + "'"
cQuery += "  And D_E_L_E_T_ = ' '"
cQuery += "  And BC_PRODUTO = '" + cProd + "'"
cQuery += "  And BC_DATA   >= '" + DtoS(dDataIni) + "'"
cQuery += "  And BC_DATA   <= '" + DtoS(dDataFim) + "'"
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSBC,.F.,.T.)
For nX := 1 To Len(aStrucTRB)
	If aStrucTRB[nX][2] <> "C" .And. FieldPos(aStrucTRB[nX][1]) <> 0
		TcSetField(cAliasSBC,aStrucTRB[nX][1],aStrucTRB[nX][2],aStrucTRB[nX][3],aStrucTRB[nX][4])
	EndIf
Next nX
dbSelectArea(cAliasSBC)
nPerda := PERDA
(cAliasSBC)->(dbCloseArea())
RestArea(aAreaAnt)
Return(nPerda)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002NCon ³ Autor ³ Larson Zordan        ³ Data ³23.02.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Apontamento de Nao Conformidades                        	  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAII002NCon(ExpC1,ExpN1,ExpN2)                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002NCon(cAlias,nReg,nOpc)
Local oBmp
Local oBtn1
Local oBtn2
Local oBtn3
Local oDlg
Local oFnt1
Local oFnt2
Local oFnt3

Local aAreaAnt := GetArea()
Local aDados   := {Space(06),Space(15),Space(20),0,"  ",0,"  ",Space(10),dDataBase,Space(06),Space(30),Space(15),Space(30),Space(02),Space(15)}
Local aPerda   := {}
Local aOps     := {}
Local aCabAP   := {}
Local aItem    := {}
Local aItemAP  := {}
Local cTitulo  := "Apontamento do Não Conformidades COMAFAL"
Local cLocaliz := ""
Local dData    := dDataBase
Local nQtdAnt  := 0
Local nPerda   := 0

Local lRet     := .T.
Local lErro    := .F.

Private cOP	        := CriaVar("D3_OP",.T.)
Private lTubo       := !Empty(cAlias)
Private lMsErroAuto := .F.

/*--------------------------------
| aDados[ 1] = Tipo              |
| aDados[ 2] = Produto           |
| aDados[ 3] = Descricao         |
| aDados[ 4] = Quant             |
| aDados[ 5] = UM                |
| aDados[ 6] = Quant 2a UM       |
| aDados[ 7] = 2a UM             |
| aDados[ 8] = Lote              |
| aDados[ 9] = Data de Validade  |
| aDados[10] = Centro de Custo   |
| aDados[11] = Desc. do CC       |
| aDados[12] = Produto Destino   |
| aDados[13] = Desc. do Prod Dest|
| aDados[14] = Motivo da Nao Conf|
| aDados[15] = Descr. do Motivo  |
--------------------------------*/

Set Key VK_F4 To U_A002ShwOP(.T.)

dbSelectArea("SC2")
dbSetOrder(1)
      
Define FONT oFnt1 NAME "Sans Serif" 	Size 10,20 BOLD 
Define FONT oFnt2 NAME "Sans Serif" 	Size 10,10 BOLD 
Define FONT oFnt3 NAME "Sans Serif" 	Size  4,12     

Define MsDialog oDlg Title cTitulo From 0,0 To 400,300 Of oMainWnd Pixel

@   1,  5 BitMap oBmp File "COMAFAL.BMP" 			Size  60,45 Of oDlg Pixel NoBorder

@   1, 65 BitMap oBmp File "NAOCONF.BMP" 			Size  80,60 Of oDlg Pixel NoBorder

@  47, 5 To 181, 146 Of oDlg Pixel

@  53, 12 Say "Ordem Produção" 						Size  40,09 Of oDlg Pixel Color CLR_HBLUE
@  60, 12 MsGet cOP Picture "@!" 					Size  40,09 Of oDlg Pixel Valid( CAI002CpoPE(@aDados,@oDlg) ) 

@  53, 55 Say "Quantidade" 							Size  50,09 Of oDlg Pixel Color CLR_HBLUE
@  60, 55 MsGet aDados[4] Picture "@E 9999.999"		Size  15,09 Of oDlg Pixel Valid( CAI002VldQt(aDados[4]) ) 

@  53, 95 Say "Lote"								Size  25,09 Of oDlg Pixel
@  60, 95 MsGet aDados[8]							Size  40,09 Of oDlg Pixel When .F. Center

@  73, 12 Say "Produto" 							Size  40,09 Of oDlg Pixel
@  81, 12 MsGet aDados[2] 							Size  40,09 Of oDlg Pixel When .F. Center 
@  81, 55 MsGet aDados[3] 							Size  85,09 Of oDlg Pixel When .F.

@  93, 12 Say "C. Centro"							Size  30,09 Of oDlg Pixel
@ 100, 12 MsGet aDados[10] F3 "CTT"					Size  40,09 Of oDlg Pixel When .T.
@ 100, 55 MsGet SI3->I3_DESC        				Size  85,09 Of oDlg Pixel When .F.

@ 113, 12 Say "Produto Destino"						Size  80,09 Of oDlg Pixel Color CLR_HBLUE
@ 120, 12 MsGet aDados[12] F3 "SB1"					Size  40,09 Of oDlg Pixel Center Valid( aDados[13] := SB1->B1_DESC )
@ 120, 55 MsGet aDados[13] 							Size  85,09 Of oDlg Pixel When .F.

@ 133, 12 Say "Motivo"         						Size  80,09 Of oDlg Pixel Color CLR_HBLUE
@ 140, 12 MsGet aDados[14] F3 "43"					Size  40,09 Of oDlg Pixel Center Valid( aDados[15] := SX5->X5_DESCRI )
@ 140, 55 MsGet aDados[15]	    					Size  85,09 Of oDlg Pixel When .F.

@ 172,25 Say AllTrim(aDados[2])+aDados[8] 			Size 100,09 Of oDlg Pixel Center Font oFnt2 Color CLR_HRED

@ 194,05 Say "[F4] Exibir OPs"						Size  80,09 Of oDlg Pixel Font oFnt3 Color CLR_HRED

Define SButton oBtn2 From 185, 85 Type  1 Enable Of oDlg OnStop "Confirma Não Conformidade" Action oDlg:End()
Define SButton oBtn3 From 185,115 Type  2 Enable Of oDlg OnStop "Cancela Não Conformidade"  Action (lRet:=.F.,oDlg:End())

Activate MsDialog oDlg Center

If lRet
	// Altera o Empenho no Saldo do Lote (SB8)
	SB8->(dbSetOrder(3))
	SB8->(dbSeek(xFilial("SB8")+aDados[02]+"01"+aDados[8]))
	RecLock("SB8",.F.)
	Replace SB8->B8_EMPENHO With (SB8->B8_EMPENHO - aDados[4]),;
			SB8->B8_EMPENH2 With (SB8->B8_EMPENH2 - aDados[4])
	MsUnLock()

	// Altera o Empenho (SD4), pois NAO devera considerar a qtd da perda no empenho, apenas como saldo
	nQtdAnt  := SD4->D4_QUANT
	RecLock("SD4",.F.)
	Replace SD4->D4_QUANT 	With (SD4->D4_QUANT   - aDados[4]),;
			SD4->D4_QTSEGUM	With (SD4->D4_QTSEGUM - aDados[4])
	MsUnLock()
	
	// Funcao padrao da Alteracao do Ajuste de Empenho
	A380Grava(nQtdAnt, Space(6), aDados[8], "01")

	// Endereco
	cLocaliz := Posicione("SB1",1,xFilial("SB1")+aDados[2],"B1_ENDEREC")

	//Cabecalho
	aCabAP   := {	{"BC_OP"		,cOP   					,Nil}}

	//Items
	aItem    := {	{"BC_PRODUTO"	,aDados[2]    			,Nil},;
			    	{"BC_LOCORIG"	,"01"					,Nil},;
					{"BC_LOCALIZ"	,cLocaliz    			,Nil},;
					{"BC_TIPO"  	,"R"					,Nil},;
					{"BC_MOTIVO"	,aDados[14]				,Nil},;
					{"BC_QUANT" 	,aDados[4]				,Nil},;
					{"BC_CODDEST"	,aDados[12]				,Nil},;
					{"BC_LOCAL" 	,"01"					,Nil},;
					{"BC_QTDDEST"	,aDados[4]				,Nil},;
					{"BC_OPERADOR"	,SubStr(cUsuario,7,15)	,Nil},;
					{"BC_LOTECTL"	,aDados[8]				,Nil},;
					{"BC_DTVALID"	,aDados[9]   			,Nil} }
	
	aAdd(aItemAP,aItem)
	
	MSExecAuto({|x,y,z| Mata685(x,y,z)},aCabAP,aItemAP,3)
	
	If lMsErroAuto
		Aviso("ATENCAO !","Ocorreu um Erro ao Apontar a Perda da OP " + cOP,{" Sair >> "})
		MostraErro()
	EndIf	

EndIf

Set Key VK_F4 To 

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002CpoPE³ Autor ³ Larson Zordan       ³ Data ³ 25.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Localiza a OP para a Nao Conformidade                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002CpoPE(ExpA1,ExpO)                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Dados da OP                                        ³±±
±±³          ³ ExpO1 = Objeto da Dialog                                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002CpoPE(aDados,oDlg)
Local oBmp
Local cTipo := ""
Local lRet  := .T. 

//--> Posiciona o Empenho do PI
SD4->(dbSetOrder(2))
If !SD4->(dbSeek(xFilial("SD4")+cOP))
	Help(" ",1,"REGNOIS")	
	lRet := .F.
EndIf
If !Empty(SD4->D4_LOTECTL)
	SC2->(dbSeek(xFilial("SC2")+SD4->D4_OPORIG))
	SI3->(dbSeek(xFilial("SI3")+SC2->C2_CC))
	cTipo      := Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_GRUPO")
	aDados[ 1] := Left( Posicione("SBM",1,xFilial("SBM")+cTipo,"BM_DESC") ,6)
	aDados[ 2] := SC2->C2_PRODUTO
	aDados[ 3] := Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_DESC")
	aDados[ 4] := 0
	aDados[ 5] := SC2->C2_UM
	aDados[ 6] := 0
	aDados[ 7] := Posicione("SB1",1,xFilial("SB1")+SC2->C2_PRODUTO,"B1_SEGUM")
	aDados[ 8] := AllTrim(SD4->D4_LOTECTL)
	aDados[ 9] := SD4->D4_DTVALID
	aDados[10] := SC2->C2_CC
	aDados[11] := SI3->I3_DESC
	aDados[12] := Space(15)
	aDados[13] := Space(30)
	aDados[14] := Space(02)
	@ 152,12 BitMap oBmp File "TBBARRAV.BMP" Size 150,80 Of oDlg Pixel NoBorder
	oDlg:Refresh()
Else
	Aviso("ATENÇÃO !","O Produto Para NÃO Conformidade, NÃO Foi Produzido !",{ " << Voltar "})
	lRet := .F.
EndIf
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002VldQt³ Autor ³ Larson Zordan       ³ Data ³ 26.02.04 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Valida a Quantidade da Perda                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002VldQt(ExpA1,ExpO)                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpN1 = Quantidade da Perda                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL      		                                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002VldQt(nQuant)
Local lRet := .T.
If nQuant == 0
	Help(" ",1,"QUANTVAZIO")
	lRet := .F.
ElseIf nQuant > SD4->D4_QUANT
	Aviso("ATENÇÃO !","A Quantidade informada é maior que o saldo empenhado para esta Ordem de Produção !" + Chr(10) + "A Quantidade máxima para esta Não Conformidade é " + AllTrim(Transform(nQtdSB8,"@E 999.999")),{ " << Voltar "},2,"Quantidade Inválida")
	lRet := .F.
EndIf
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CalcEmpOC ³ Autor ³ Larson Zordan        ³ Data ³05.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Calcula o Empenho e Perda de Cada Bobina na Ordem de Corte ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CalcEmpOC(ExpC1,nLagura,ExpN1,ExpA1)                       ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Numero da Ordem de Corte                           ³±±
±±³          ³ ExpN1 = Peso da Bobina                                     ³±±
±±³          ³ ExpA1 = Medidas do Slits Selecionados                      ³±±
±±³          ³       [1] = Quantidade de Rolos                            ³±±
±±³          ³       [2] = Largura do Slit                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Retorno   ³ ExpA1 = Array Com As Quantidades de Empenho e Perda        ³±±
±±³          ³       [1] = Peso Empenhado da Bobina                       ³±±
±±³          ³       [2] = Perda da Bobina                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CalcEmpOC(cNumOC,nLargura,nPeso,aMedidas)
Local nQtEmp   := 0
Local nX

For nX := 1 To Len(aMedidas)
	//--> Formula: (( Qt Rolos * Largura do Slit) * Peso da Bobina) / Largura Real da Bobina
	nQtEmp += Round(( (aMedidas[nX,1] * aMedidas[nX,2]) * nPeso ) / nLargura,3)
Next nX

Return( {nQtEmp, (nPeso-nQtEmp) })

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ RefazEmpOC³ Autor ³ Larson Zordan        ³ Data ³05.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Grava os Empenhos e as Perdas na OCs (SZ2)                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ RefazEmpOC()                                               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function RefazEmpOC(cNumOC)
Local aAreaAnt  := GetArea()
Local aEmpPerda := {}
Local aMedidas  := {}
Local cCond     := "!Eof()"
Local nX

Default cNumOC  := ""

dbSelectArea("SZ3")
dbSetOrder(1)
dbGoTop()
If !Empty(cNumOC)
	dbSeek(xFilial("SZ3")+cNumOC)
	cCond += " .And. SZ3->Z3_NUM == '" + cNumOC + "'"
EndIf	
While &(cCond)
	aEmpPerda := {}
	aMedidas  := {}
	
	//--> Obtem o Peso Empenhado e Perda de Cada Lote
	dbSelectArea("SZ2")
	dbSetOrder(1)
	dbSeek(xFilial("SZ2")+SZ3->Z3_NUM)
	While !Eof() .And. SZ2->Z2_FILIAL+SZ2->Z2_NUM == xFilial("SZ2")+SZ3->Z3_NUM
		If Z2_TIPO == "S"
			aAdd(aMedidas,{Val(SubStr(SZ2->Z2_DESC,48, 5)), Val(Right(AllTrim(SubStr(SZ2->Z2_DESC,17,30)),4)) })
		EndIf
		dbSkip()
	EndDo	

	//--> Calcula o Empenho e a Perda e Grava no campo Z2_DESC
	dbSeek(xFilial("SZ2")+SZ3->Z3_NUM)
	While !Eof() .And. SZ2->Z2_FILIAL+SZ2->Z2_NUM == xFilial("SZ2")+SZ3->Z3_NUM
		If Z2_TIPO == "B"
			aEmpPerda := CalcEmpOC(SZ3->Z3_NUM,Val(SZ3->Z3_LARGREA),Val(SubStr(Z2_DESC,19,12)),aMedidas)
			RecLock("SZ2",.F.)
			Replace Z2_DESC With SubStr(Z2_DESC,1,39) + "|" + Str(aEmpPerda[1],12,3) + "|" + Str(aEmpPerda[2],12,3) + "|"
			MsUnLock()
		EndIf
		dbSkip()
	EndDo	

	dbSelectArea("SZ3")
	dbSkip()
EndDo 
RestArea(aAreaAnt)
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ ReabreOC  ³ Autor ³ Larson Zordan        ³ Data ³08.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Reabre a OC                                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ReabraOC()                                                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function ReabreOC(cAlias,nReg,nOpc)
Local aAreaAnt := GetArea()
Local cOC      := CriaVar("Z3_NUM",.F.)
If !Empty(cAlias)
	cOC := SZ3->Z3_NUM
EndIf
If U_MsgGet("Reabre a Ordem de Corte","Informe a Ordem de Corte :",@cOC,"@!")
	If Empty(cOC)
		NaoVazio(cOC)
		Return
	EndIf
	dbSelectArea("SZ3")
	dbSetOrder(1)
	If !dbSeek(xFilial("SZ3")+cOC)
		Aviso("ATENCAO","Nao existe Esta Ordem de Corte !",{" Sair >> "})
		Return
	EndIf
	Begin Transaction
		U_RefazEmpOC(cOC)
		MsgRun("Aguarde... Reabrindo Ordem de Corte " + cOC + "...","",{ || ProcReabre(cOC) })
	End Transaction
EndIf
RestArea(aAreaAnt)
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ ProcReabre³ Autor ³ Larson Zordan        ³ Data ³08.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa a Reabertura da Ordem de Corte                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ ProcReabre(cNumOC,nOPs)                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function ProcReabre(cNumOC)
Local cQuery   := ""
Local cProduto := SZ3->Z3_PRODUTO
Local cDoc     := ""
Local cLoteCtl := ""
Local cNumLote := ""
Local cOPIni   := ""
Local cOPFim   := ""
Local cTipo    := ""
Local nQtdOrig := 0
Local nQtdEmp  := 0
Local nQtdPerd := 0
Local nSaldo   := 0
Local nEmpenho := 0
Local nBobinas := 0
Local nSlits   := 0
Local lCont    := .T.

dbSelectArea("SZ2")
dbSetOrder(1)
dbSeek(xFilial("SZ2")+cNumOC)
While !Eof() .And. xFilial("SZ2")+cNumOC == Z2_FILIAL+Z2_NUM
                                
	cTipo := Z2_TIPO
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ ACERTA AS BOBINAS      ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	If cTipo == "B"
		nBobinas ++
		cLoteCtl := SubStr(SZ2->Z2_DESC, 1,10)
		nQtdOrig := Val(SubStr(SZ2->Z2_DESC,19,12))
		nQtdEmp  := Val(SubStr(SZ2->Z2_DESC,41,12))
		nQtdPerd := Val(SubStr(SZ2->Z2_DESC,54,12))
		
		//--> Refazer o Ajuste dos Empenhos (SD4)
		dbSelectArea("SD4")
		dbSetOrder(7)
		If dbSeek(xFilial("SD4")+cProduto+"01"+cLoteCtl)
			While !Eof() .And. xFilial("SD4")+cProduto+"01"+cLoteCtl == D4_FILIAL+D4_COD+D4_LOCAL+D4_LOTECTL
				//--> Ajusta o Empenhos (SD4)
				RecLock("SD4",.F.)
				Replace D4_QUANT   With D4_QTDEORI, D4_QTSEGUM With D4_QTDEORI
				MsUnLock()

				//--> Acerta a Composicao do Empenho (SDC)
				dbSelectArea("SDC")
				dbSetOrder(2)
				If dbSeek(xFilial("SDC")+cProduto+"01"+SD4->D4_OP+SD4->D4_TRT+cLoteCtl) 
					RecLock("SDC",.F.)
					Replace DC_QUANT   With SD4->D4_QUANT	,;
							DC_QTSEGUM With SD4->D4_QUANT	,;
							DC_QTDORIG With SD4->D4_QUANT
					MsUnLock()
				Else
					RecLock("SDC",.T.)
					Replace DC_FILIAL 	With xFilial("DC")	,;
							DC_ORIGEM 	With "SC2"			,;
							DC_PRODUTO 	With cProduto		,;
							DC_LOCAL	With "01"			,;
							DC_LOCALIZ	With Posicione("SB1",1,xFilial("SB1")+cProduto,"B1_ENDEREC"),;
							DC_LOTECTL	With cLoteCtl		,;
							DC_QUANT	With SD4->D4_QUANT	,;
							DC_OP		With SD4->D4_OP		,;
							DC_TRT		With SD4->D4_TRT	,;
							DC_QTDORIG	With SD4->D4_QUANT	,;
							DC_QTSEGUM	With SD4->D4_QUANT
					MsUnLock()
				EndIf	
				
				dbSelectArea("SD4")
				dbSkip()
			EndDo
		EndIf
		
		//--> Acerta o Saldo e Empenho do Lote (SB8)
		lCont := .T.
		dbSelectArea("SB8")
		dbSetOrder(5)
		If dbSeek(xFilial("SB8")+cProduto+cLoteCtl)
			While !Eof() .And. xFilial("SB8")+cProduto+cLoteCtl == SB8->(B8_FILIAL+B8_PRODUTO+B8_LOTECTL)
				RecLock("SB8",.F.)
				If lCont
					cDoc     := B8_DOC
					cNumLote := B8_NUMLOTE
					Replace B8_QTDORI  With nQtdOrig,;
							B8_QTDORI2 With nQtdOrig,;
							B8_SALDO   With nQtdOrig,;
							B8_SALDO2  With nQtdOrig,;
							B8_EMPENHO With nQtdEmp	,;
							B8_EMPENH2 With nQtdEmp
					lCont := .F.
				Else
					Delete
				EndIf
				MsUnLock()
				dbSkip()
			EndDo

			//--> Acerta o Movimentos do Lote (SD5)
			lCont := .T.
			dbSelectArea("SD5")
			dbSetOrder(2)
			If dbSeek(xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote)
				While !Eof() .And. xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote == SD5->(D5_FILIAL+D5_PRODUTO+D5_LOCAL+D5_LOTECTL+D5_NUMLOTE)
					RecLock("SD5",.F.)
					If lCont
						If cDoc == D5_DOC
							Replace D5_QUANT   With nQtdOrig,;
									D5_QTSEGUM With nQtdOrig
							lCont := .F.
						EndIf				
					Else
						Delete
					EndIf
					MsUnLock()
					dbSkip()
				EndDo
			EndIf

		Else
			//--> Verificar o Inventario
			dbSelectArea("SB7")
			dbSetOrder(2)
			If dbSeek(xFilial("SB7")+Space(6)+cLoteCtl+cProduto+"01")
				dbSelectArea("SB8")
				RecLock("SB8",.T.)
				Replace B8_FILIAL	With xFilial("SB8") ,;
						B8_QTDORI	With SB7->B7_QUANT	,;
						B8_PRODUTO	With SB7->B7_COD	,;
						B8_LOCAL	With "01"         	,;
						B8_DATA		With SB7->B7_DATA	,;
						B8_DTVALID	With SB7->B7_DTVALID,;
						B8_SALDO	With SB7->B7_QUANT	,;
						B8_EMPENHO	With nQtdEmp		,;
						B8_ORIGLAN	With "MAN"			,;
						B8_LOTECTL	With cLoteCtl    	,;
						B8_SALDO2	With SB7->B7_QTSEGUM,;
						B8_EMPENH2	With nQtdEmp		,;
						B8_DOC		With SB7->B7_DOC	,;
						B8_QTDORI2	With SB7->B7_QtSEGUM,;
						B8_POTENCI	With 0
				MsUnLock()

				dbSelectArea("SD5")
				RecLock("SD5",.T.)
				Replace D5_FILIAL	With xFilial("SD5")	,;
						D5_PRODUTO	With SB7->B7_COD	,;
						D5_LOCAL	With SB7->B7_LOCAL	,;
						D5_DOC		With SB7->B7_DOC	,;
						D5_DATA		With SB7->B7_DATA	,;
						D5_ORIGLAN	With "MAN"			,;
						D5_NUMSEQ	With ProxNum()		,;
						D5_QUANT	With SB7->B7_QUANT	,;
						D5_LOTECTL	With SB7->B7_LOTECTL,;
						D5_DTVALID	With SB7->B7_DTVALID,;
						D5_QTSEGUM	With SB7->B7_QTSEGUM,;
						D5_POTENCI	With 0
				MsUnLock()
			EndIf	

		EndIf

		//--> Acerta o Saldo e Empenho do Endereco (SBF)
		lCont := .T.
		dbSelectArea("SBF")
		dbSetOrder(2)
		If dbSeek(xFilial("SBF")+cProduto+"01"+cLoteCtl)
			While !Eof() .And. xFilial("SBF")+cProduto+"01"+cLoteCtl == SBF->(BF_FILIAL+BF_PRODUTO+BF_LOCAL+BF_LOTECTL)
				RecLock("SBF",.F.)
				If lCont
					Replace BF_QUANT   With nQtdOrig,;
							BF_QTSEGUM With nQtdOrig,;
							BF_EMPENHO With nQtdEmp	,;
							BF_EMPEN2  With nQtdEmp
					lCont := .F.
				Else
					Delete
				EndIf
				MsUnLock()
				dbSkip()
			EndDo
		Else	
			RecLock("SBF",.T.)
			Replace BF_FILIAL	With xFilial("SBF")	,;
					BF_PRODUTO  With cProduto   	,;
					BF_LOCAL	With "01"         	,;
					BF_PRIOR	With "ZZZ"			,;
					BF_LOCALIZ	With Posicione("SB1",1,xFilial("SB1")+cProduto+"01","B1_ENDEREC"),;
					BF_LOTECTL	With cLotectl		,;
					BF_QUANT	With nQtdOrig		,;
					BF_EMPENHO 	With nQtdEmp		,;
					BF_EMPEN2  	With nQtdEmp		,;
					BF_QTSEGUM	With nQtdOrig
			MsUnLock()
		EndIf

		//--> Acerta o Movimentos Enderecados do Lote (SDB)
		lCont := .T.
		dbSelectArea("SDB")
		dbSetOrder(2)
		If dbSeek(xFilial("SDB")+cProduto+"01"+cLoteCtl)
			While !Eof() .And. xFilial("SD5")+cProduto+"01"+cLoteCtl == SDB->(DB_FILIAL+DB_PRODUTO+DB_LOCAL+DB_LOTECTL)
				RecLock("SDB",.F.)
				If lCont
					If cDoc == DB_DOC
						Replace DB_QUANT   With nQtdOrig,;
								DB_QTSEGUM With nQtdOrig,;
								DB_EMPENHO With nQtdEmp	,;
								DB_EMP2    With nQtdEmp
						lCont := .F.
					EndIf				
				Else
					Delete
				EndIf
				MsUnLock()
				dbSkip()
			EndDo
		Else	
			RecLock("SDB",.T.)
			Replace DB_FILIAL	With xFilial("SDB")	,;
					DB_ITEM		With "0001"			,;
					DB_PRODUTO	With cProduto   	,;
					DB_LOCAL 	With "01"         	,;
					DB_LOCALIZ	With Posicione("SB1",1,xFilial("SB1")+cProduto+"01","B1_ENDEREC"),;
					DB_DOC		With cDoc       	,;
					DB_TM		With "499"			,;
					DB_ORIGEM	With "SD5"			,;
					DB_QUANT	With nQtdOrig		,;
					DB_DATA		With dDataBase   	,;
					DB_LOTECTL	With cLoteCtl       ,;
					DB_NUMSEQ	With ProxNum()		,;
					DB_TIPO		With "D"			,;
					DB_EMPENHO 	With nQtdEmp		,;
					DB_EMP2    	With nQtdEmp		,;
					DB_QTSEGUM	With nQtdOrig
			MsUnLock()		
		EndIf
		
	Else
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ ACERTA OS SLITS        ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

		nSlits ++
		cProduto := SubStr(Z2_DESC,1,15)
		nQtdEmp  := 0

		//--> Refazer o Ajuste dos Empenhos (SD4)
		dbSelectArea("SD4")
		dbSetOrder(6)
		If dbSeek(xFilial("SD4")+cNumOC)
			While !Eof() .And. xFilial("SD4")+cNumOC == D4_FILIAL+D4_NUMOC
				If cProduto == D4_COD
					cLoteCtl := D4_LOTECTL
					nQtdEmp  := D4_QTDEORI
					RecLock("SD4",.F.)
					Replace D4_QUANT   With nQtdEmp						,;
							D4_QTSEGUM With nQtdEmp						,;
							D4_LOTECTL With CriaVar("D4_LOTECTL",.F.)	,;
							D4_DTVALID With CriaVar("D4_DTVALID",.F.)
					MsUnLock()
					Exit
				EndIf
				dbSkip()
			EndDo	
		EndIf

		//--> Acerta o Saldo e Empenho do Lote (SB8)
		lCont := .T.
		dbSelectArea("SB8")
		dbSetOrder(5)
		If dbSeek(xFilial("SB8")+cProduto+cLoteCtl)
 			cDoc     := B8_DOC
			cNumLote := B8_NUMLOTE
			While !Eof() .And. xFilial("SB8")+cProduto+cLoteCtl == SB8->(B8_FILIAL+B8_PRODUTO+B8_LOTECTL)
				RecLock("SB8",.F.)
				Delete
				MsUnLock()
				dbSkip()
			EndDo
		EndIf

		//--> Acerta o Movimentos do Lote (SD5)
		dbSelectArea("SD5")
		dbSetOrder(2)
		If dbSeek(xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote)
			While !Eof() .And. xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote == SD5->(D5_FILIAL+D5_PRODUTO+D5_LOCAL+D5_LOTECTL+D5_NUMLOTE)
				If cDoc == D5_DOC
					RecLock("SD5",.F.)
					Delete
					MsUnLock()
				EndIf
				dbSkip()
			EndDo
		EndIf
	
	EndIf 

	dbSelectArea("SZ2")
	dbSkip()

	If cProduto <> If(cTipo == "B",SZ3->Z3_PRODUTO,SubStr(Z2_DESC,1,15))
		If cTipo == "B"
			//--> Ajusta o Saldo do Estoque (SB2)
			cQuery := "Select Sum(B8_SALDO) AS QATU, Sum(B8_EMPENHO) AS QEMP "
			cQuery += "From " + RetSqlName("SB8") + " Where "
			cQuery += "    B8_FILIAL  = '" + xFilial("SB8") + "' "
			cQuery += "And D_E_L_E_T_ = ' ' "
			cQuery += "And B8_PRODUTO = '" + cProduto + "' "
			ChangeQuery(cQuery)
			dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"TRB",.F.,.T.)
			dbSelectArea("TRB")
			nSaldo   := TRB->QATU
			nEmpenho := TRB->QEMP
			TRB->(dbCloseArea())
		Else
			nEmpenho := nQtdEmp
		EndIf
		dbSelectArea("SB2")
		dbSetOrder(1)
		dbSeek(xFilial("SB2")+cProduto)
		RecLock("SB2",.F.)
		If cTipo == "B"
			Replace B2_QATU		With nSaldo				,;
					B2_VATU1 	With (nSaldo * B2_CM1)	,;
					B2_QTSEGUM  With nSaldo				,;
					B2_LOCALIZ	With Posicione("SB1",1,xFilial("SB1")+cProduto,"B1_ENDEREC")
		EndIf
		Replace		B2_QEMP		With nEmpenho			,;
					B2_QEMP2	With nEmpenho			
		MsUnLock()
    EndIf

	If cNumOC <> Z2_NUM
		//--> Acerta as Ordens de Producao - Estorna (SC2)
		dbSelectArea("SC2")
		dbSetOrder(10)
		If dbSeek(xFilial("SC2")+cNumOC)
			While !Eof() .And. xFilial("SC2")+cNumOC == C2_FILIAL+C2_NUMOC

				//--> Ajusta Saldo do PA (SB2)
				If C2_SEQUEN == "001"
					dbSelectArea("SB2")
					If dbSeek(xFilial("SB2")+SC2->C2_PRODUTO+SC2->C2_LOCAL)
						RecLock("SB2",.F.)
						Replace B2_QATU 	With (B2_QATU - SC2->C2_QUJE)	,;
								B2_VATU1	With (B2_CM1  + B2_QATU)		,;
								B2_QTSEGUM	With B2_QATU
						MsUnLock()
					EndIf	
				EndIf

				//--> Ajusta Ordem de Producao (SC2)
				dbSelectArea("SC2")
				If Empty(cOPIni)
					cOPIni := C2_NUM+C2_ITEM+C2_SEQUEN+C2_ITEMGRD
				EndIf
				cOPFim   := C2_NUM+C2_ITEM+C2_SEQUEN+C2_ITEMGRD
				RecLock("SC2",.F.)
				Replace C2_QUJE 	With 0,;
						C2_PERDA 	With 0,;
						C2_APRATU1 	With 0,;
						C2_DATRF 	With CriaVar("C2_DATRF",.F.)
				MsUnLock()
				
				dbSkip()				
			EndDo
		EndIf

		//--> Acerta as Movimentacoes Internas - Estorna (SD3)
		dbSelectArea("SD3")
		dbSetOrder(1)
		If dbSeek(xFilial("SD3")+Left(cOPIni,8),.T.)
			While !Eof() .And. xFilial("SD3") == D3_FILIAL
				If D3_OP >= cOPIni .And. D3_OP <= cOPFim
					RecLock("SD3",.F.)
					Delete
					MsUnLock()
				EndIf
				dbSkip()
			EndDo
		EndIf
		
		//--> Acerta as Perdas Por OP - Estorna (SBC)
		dbSelectArea("SBC")
		dbSetOrder(1)
		If dbSeek(xFilial("SBC")+cOPFim)
			dbSelectArea("SB2")
			If dbSeek(xFilial("SB2")+SBC->BC_CODDEST+SBC->BC_LOCDEST)
				RecLock("SB2",.F.)
				Replace B2_QATU 	With (B2_QATU - SBC->BC_QTDDEST)	,;
						B2_VATU1	With (B2_CM1  + B2_QATU)			,;
						B2_QTSEGUM	With B2_QATU
				MsUnLock()
		    EndIf
				    
			dbSelectArea("SB8")
			dbSetOrder(3)
			If dbSeek(xFilial("SB8")+SBC->BC_CODDEST+SBC->BC_LOCAL+SBC->BC_LOTECTL)
				While !Eof() .And. xFilial("SB8")+SBC->BC_CODDEST+SBC->BC_LOCAL+SBC->BC_LOTECTL == B8_FILIAL+B8_PRODUTO+B8_LOCAL+B8_LOTECTL
					RecLock("SB8",.F.)
					cDoc     := B8_DOC
					cNumLote := B8_NUMLOTE
					If B8_SALDO > 0
						Replace B8_SALDO	With (B8_SALDO - SBC->BC_QTDDEST),	B8_SALDO2	With B8_SALDO
					Else
						Delete			
					EndIf
					MsUnLock()

					lCont := .T.
					dbSelectArea("SD5")
					dbSetOrder(2)
					If dbSeek(xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote)
						While !Eof() .And. xFilial("SD5")+cProduto+"01"+cLoteCtl+cNumLote == SD5->(D5_FILIAL+D5_PRODUTO+D5_LOCAL+D5_LOTECTL+D5_NUMLOTE)
							RecLock("SD5",.F.)
							If lCont
								If cDoc == D5_DOC
									Replace D5_QUANT   With (D5_QUANT - SBC->BC_QTDDEST),	D5_QTSEGUM With D5_QUANT
									lCont := .F.
								EndIf				
							Else
								Delete
							EndIf
							MsUnLock()
							dbSkip()
						EndDo
					EndIf
					
    				dbSelectArea("SB8")
					dbSkip()
				EndDo	
			EndIf
				
			dbSelectArea("SD3")
			dbSetOrder(4)
			If dbSeek(xFilial("SD3")+SBC->BC_SEQSD3+"E0",.T.)
				While !Eof() .And. xFilial("SD3")+SBC->BC_SEQSD3+"E0" == D3_FILIAL+D3_NUMSEQ+D3_CHAVE
					If SBC->BC_LOTECTL == D3_LOTECTL
						RecLock("SD3",.F.)
						Delete
						MsUnLock()
					EndIf
					dbSkip()
				EndDo
			EndIf

			dbSelectArea("SBC")
			RecLock("SBC",.F.)
			Delete
			MsUnLock()
		EndIf
		dbSelectArea("SZ2")
	EndIf
EndDo
dbSelectArea("SZ3")
RecLock("SZ3",.F.)
Replace Z3_QTOPS With (nBobinas * nSlits),	Z3_OPSBXA With 0
MsUnLock()
Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ProcEstSlit³ Autor ³ Larson Zordan        ³ Data ³16.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Processa o Estorno do Slit por Bobina                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ BuscaOPs(ExpC1,ExpA1,ExpA2)                                ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Numero da Ordem de Corte                           ³±±
±±³          ³ ExpA1 = Lotes das Bobinas                                  ³±±
±±³          ³         [1] .T. - Selecionado     .F. - Nao Selecionado    ³±±
±±³          ³         [2] LoteCtl                                        ³±±
±±³          ³         [3] Saldo                                          ³±±
±±³          ³         [4] Data de Validade                               ³±±
±±³          ³         [5] Numero do Lote                                 ³±±
±±³          ³ ExpA2 = Dados dos Sliters selecionados                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function ProcEstSlit(cNumOC,aBobinas,aSliter)
Local aAreaAnt  := GetArea()
Local aBobSel   := {}
Local aOPs      := {}
Local aStrucSD3	:= SD3->(dbStruct())
Local cAliasTRB := "ESTSD3" + xFilial("SB8")
Local cQuery    := ""
Local cSlitIni  := ""
Local cSlitFim  := ""
Local nX, nY
       
//--> Cria Array apenas com as Bobinas Selecionadas
For nX := 1 To Len(aBobinas)
	If aBobinas[nX,1]
		aAdd(aBobSel,aBobinas[nX,2])
	EndIf
Next nX

//--> Verifica Se ha bobina selecionada
If Len(aBobSel) == 0
	Aviso("ATENÇÃO !","Não Há Bobina Selecionada Para Estornar o Apontamento do Slit.",{ " << Voltar "},1,"Estornar Slit")
	Return(.T.)
EndIf	

//--> Sliter De/Ate Produzidos
For nX := 1 To Len(aSliter)
	If Empty(cSlitIni)
		cSlitIni := aSliter[nX,1]
	EndIf
	cSlitFim := aSliter[nX,1]
Next nX

dbSelectArea("SD4")
dbSetOrder(6)

//--> Pesquisa OPs 
For nY := 1 To Len(aBobSel)
	If dbSeek(xFilial("SD4")+cNumOC)
		While !Eof() .And. xFilial("SD4")+cNumOC == D4_FILIAL+D4_NUMOC
			If D4_LOTECTL == aBobSel[nY]
				cQuery := "Select * "
				cQuery += "From " + RetSqlName("SD3") + " "
				cQuery += "Where D3_FILIAL   = '" + xFilial("SD3")	+ "'"
				cQuery += "  And D3_TM       = '300'"
				cQuery += "  And D3_OP       = '" + SD4->D4_OP	+ "'"
				cQuery += "  And D3_COD     >= '" + cSlitIni	+ "'"
				cQuery += "  And D3_COD     <= '" + cSlitFim	+ "'"
				cQuery += "  And D_E_L_E_T_  = ' ' "
				cQuery += "Order By D3_OP"
				cQuery := ChangeQuery(cQuery)
				dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasTRB,.F.,.T.)
				For nX := 1 To Len(aStrucSD3)
					If aStrucSD3[nX][2] <> "C" .And. FieldPos(aStrucSD3[nX][1]) <> 0
						TcSetField(cAliasTRB,aStrucSD3[nX][1],aStrucSD3[nX][2],aStrucSD3[nX][3],aStrucSD3[nX][4])
					EndIf
				Next nX
				
				dbSelectArea(cAliasTRB)
				dbGoTop()
				
				While !Eof()
					aAdd(aOPs,{	D3_TM			,;	// Tipo do Movimento
								D3_COD			,;	// Produto
								D3_UM			,;	// 1a. UM
								D3_QUANT		,;	// Quantidade
								D3_CONTA		,;	// Conta Contabil
								D3_OP			,;	// Ordem de Producao
								D3_CC			,;	// Centro de Custo
								D3_LOCAL		,;	// Armazem
								D3_SEGUM		,;	// 2a UM
								D3_PERDA		,;	// Qtd Perda
								D3_LOTECTL		,;	// LoteCtl
								D3_DTVALID		,;	// Data da Validade
								SD4->D4_LOTECTL	,;	// Lote da Bobina
								SD4->D4_COD		})	// Codigo da Bobina
					dbSkip()
				EndDo
				(cAliasTRB)->(dbCloseArea())
	        EndIf
	        dbSelectArea("SD4")
			dbSkip()
		EndDo
	EndIf	
Next nY

SB8->(dbSetOrder(3))
//--> Iniciar o Estorno do Slit
For nX := 1 To Len(aOPs)
	//--> Subtrair o Empenho do Lote do SLIT
	If SB8->(dbSeek(xFilial("SB8")+aOPs[nX,2]+aOPs[nX,8]+aOPs[nX,11]))
		RecLock("SB8",.F.)
		Replace SB8->B8_EMPENHO With (SB8->B8_EMPENHO - aOPs[nX, 4]),;
				SB8->B8_EMPENH2 With (SB8->B8_EMPENH2 - aOPs[nX,10])
		MsUnLock()
	EndIf			

	//--> Monta Array para a Rotina Automatica
	aSD3   := {	{"D3_TM"     	,	aOPs[nX, 1]	,Nil}	,;
				{"D3_COD"    	,	aOPs[nX, 2]	,Nil}	,;
				{"D3_UM"		,	aOPs[nX, 3]	,Nil}	,;
				{"D3_QUANT"		,	aOPs[nX, 4]	,Nil}	,;
				{"D3_CONTA"  	,	aOPs[nX, 5]	,Nil}	,;
				{"D3_OP"     	,	aOPs[nX, 6]	,Nil}	,;
				{"D3_CC"     	,	aOPs[nX, 7]	,Nil}	,;
				{"D3_LOCAL"		,	aOPs[nX, 8]	,Nil}	,;
				{"D3_SEGUM"		,	aOPs[nX, 9]	,Nil}	,;
				{"D3_QTSEGUM"	,	aOPs[nX,10]	,Nil}	,;
				{"D3_PERDA"  	,	0          	,Nil}	,;
				{"D3_LOTECTL"	,	aOPs[nX,11]	,Nil}	,;
				{"D3_DTVALID"	,	aOPs[nX,12]	,Nil}	} 

	MsExecAuto({|x,y| Mata250(x,y)},aSD3,5) //Estornar
	If lMsErroAuto
		lErro := .T.
		Aviso("ATENCAO !","Ocorreu um Erro ao Estornar a OP " + aOPs[nX,6],{" Sair >> "},2,"Estorno da OP")
		If lViewRot
			MostraErro()
		EndIf
	Else	
		//--> Estornar o Apontamento da Perda
		If nX == Len(aOPs)
			dbSelectArea("SBC")
			dbSetOrder(1)
			If dbSeek(xFilial("SBC")+aOPs[nX,6])
				While !Eof() .And. xFilial("SBC")+aOPs[nX,6] == BC_FILIAL+BC_OP
					If aOPs[nX,14] == BC_PRODUTO .And. aOPs[nX,13] == BC_LOTECTL
						EstornaSBC(aOPs[nX,6],SBC->BC_NUMSEQ)
						Exit
					EndIf
					dbSkip()
				EndDo
			EndIf	
		EndIf
	EndIf
Next nX

//--> Atualiza a OC
dbSelectArea("SZ3")
RecLock("SZ3",.F.)
Replace Z3_OPSBXA With (Z3_OPSBXA - (Len(aBobSel)*Len(aSliter)))
MsUnLock()

RestArea(aAreaAnt)
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ CAI002EstT³ Autor ³ Larson Zordan        ³ Data ³15.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Estorno dos Tubos Apontados pela OC ou pelo MATA250        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002EstT(ExpC1,ExpN1,ExpN2)                              ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpC1 = Alias do arquivo                                   ³±±
±±³          ³ ExpN1 = Numero do registro                                 ³±±
±±³          ³ ExpN2 = Numero da opcao selecionada                        ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

User Function CAI002EstT(cAlias,nReg,nOpc)

Local oDlg
Local oFnt
Local oLbx

Local cTitulo := "Estorno do Apontamento de " + If(!Empty(cAlias),"Tubos","Ordem de Produção")

Local aAponta := {}
Local aDados  := {Space(06),Space(15),Space(20),0,"  ",0,"  ",Space(10),dDataBase,Space(06),Space(30)}
Local lGrava  := .F.

/*--------------------------------
| aDados[ 1] = Tipo              |
| aDados[ 2] = Produto           |
| aDados[ 3] = Descricao         |
| aDados[ 4] = Quant             |
| aDados[ 5] = UM                |
| aDados[ 6] = Quant 2a UM       |
| aDados[ 7] = 2a UM             |
| aDados[ 8] = Lote              |
| aDados[ 9] = Data de Validade  |
| aDados[10] = Centro de Custo   |
| aDados[11] = Desc. do CC       |
--------------------------------*/

Private cOP   := CriaVar("D3_OP",.T.)
Private lTubo := .T.
Private oNo	  := LoadBitmap( GetResources(), "LBNO"       )
Private oOk	  := LoadBitmap( GetResources(), "LBTIK"      )
      
Set Key VK_F4 To U_A002ShwOP(.F.)

Define FONT oFnt NAME "San Serif" Size  5,13

dbSelectArea("SC2")
dbSetOrder(1)

aAdd(aAponta,{ .F., Space(10), CtoD("  /  /  "), 0, CtoD("  /  /  "), Space(8), Space(15) })
      
Define MsDialog oDlg Title cTitulo From 0,0 To 400,600 Of oMainWnd Pixel
@ 15, 2 To 40,299 Of oDlg Pixel

@ 24, 10   Say "Ordem de Produção :" 		Size  55,09 Of oDlg Pixel
@ 23, 65 MsGet cOP Picture "@!" 			Size  45,09 Of oDlg Pixel Valid( CAI002CpoOP(@aDados,@oDlg,.F.,.F.) , CAI002ApoOP(@aAponta,@oLbx) ) 

@ 24,125   Say "Produto :"					Size  40,09 Of oDlg Pixel
@ 23,150 MsGet aDados[2]					Size  40,09 Of oDlg Pixel When .F.
@ 23,190 MsGet aDados[3]					Size 100,09 Of oDlg Pixel When .F.

@ 45,  2 ListBox oLbx Fields Header "","No. Lote","Validade","Quantidade","Apontado Em","Tipo Apontamento" 	;
		 Size 297,145 Of oDlg Pixel On DBLCLICK ( aAponta:=CAI002OPMarc(oLbx:nAt,aAponta),oLbx:Refresh(),oDlg:Refresh() )

oLbx:SetArray(aAponta)
oLbx:bLine := {|| {If(aAponta[oLbx:nAt,1],oOk,oNo), aAponta[oLbx:nAT,2], DtoC(aAponta[oLbx:nAT,3]) ,Transform(aAponta[oLbx:nAT,4],"@E 999.999"), DtoC(aAponta[oLbx:nAT,5]), aAponta[oLbx:nAT,6] } }

@ 193,10 Say "[F4] Exibir OPs"				Size  80,09 Of oDlg Pixel Font oFnt Color CLR_HRED

Activate MsDialog oDlg Center On Init EnchoiceBar(oDlg, {|| (lGrava:=.T.,oDlg:End()) },{|| oDlg:End() } )

If lGrava
	Begin Transaction
		MsgRun("Aguarde... Estornando os Apontamento da OP "+cOP,"",{ || CAI002GrvEst(aDados,aAponta)})
	End Transaction	
EndIf

Return
/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³CAI002ApoOP ³ Autor ³ Larson Zordan       ³ Data ³17.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Monata array com os Aponatmentos da OP                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ CAI002ApoOP(ExpA1)                                         ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpA1 = Array dos Apontamentos                             ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
Static Function CAI002ApoOP(aAponta,oLbx)
Local aAreaAnt := GetArea()
aAponta := {}
dbSelectArea("SD3")
dbSetOrder(1)
If dbSeek(xFilial("SD3")+cOP)
	While !Eof() .And. xFilial("SD3")+cOP == D3_FILIAL+D3_OP
		If D3_TM == "300" .And. Empty(D3_ESTORNO)
			aAdd(aAponta,{	.F., D3_LOTECTL, D3_DTVALID, D3_QUANT, D3_EMISSAO, If(D3_CF=="PR0","Produção","Mov.Interno"), D3_COD } )
		EndIf
		dbSkip()
	EndDo
EndIf

dbSelectArea("SBC")
dbSetOrder(1)
If dbSeek(xFilial("SBC")+cOP)
	While !Eof() .And. xFilial("SBC")+cOP == BC_FILIAL+BC_OP
		aAdd(aAponta,{	.T., BC_LOTECTL, BC_DTVALID, BC_QUANT, BC_DATA, "Perda", BC_PRODUTO } )
		dbSkip()
	EndDo
EndIf
If Len(aAponta) == 0
	aAdd(aAponta,{ .F., Space(10), CtoD("  /  /  "), 0, CtoD("  /  /  "), Space(8), Space(15) })
EndIf
oLbx:SetArray(aAponta)
oLbx:bLine := {|| {If(aAponta[oLbx:nAt,1],oOk,oNo), aAponta[oLbx:nAT,2], DtoC(aAponta[oLbx:nAT,3]) ,Transform(aAponta[oLbx:nAT,4],"@E 999.999"), DtoC(aAponta[oLbx:nAT,5]), aAponta[oLbx:nAT,6] } }
oLbx:Refresh()
RestArea(aAreaAnt)
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Funcao   ³CAI002OPMarc³ Autor ³ Larson Zordan       ³ Data ³17.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Seleciona Item na ListBox - Apontamentos                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002OPMarc(nX,aArray)
If aArray[nX,6] == "Perda"
	Aviso("ATENÇÃO !!!","Este Item NÃO Poderá Ser Desmarcado !",{ " << Voltar " })
Else
	aArray[nX,1] := !aArray[nX,1]
EndIf
Return(aArray)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Funcao   ³CAI002GrvEst³ Autor ³ Larson Zordan       ³ Data ³17.03.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Grava o estorno da OP                                      ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function CAI002GrvEst(aDados,aAponta)
Local aSD3 := {}
Local nX
Private lMsErroAuto := .F.

//--> Estornar Primeiro as Perdas
For nX := 1 To Len(aAponta)
	If aAponta[nX,6] == "Perda"
		dbSelectArea("SBC")
		dbSetOrder(1)
		If dbSeek(xFilial("SBC")+cOP)
			While !Eof() .And. xFilial("SBC")+cOP == BC_FILIAL+BC_OP
				If aAponta[nX,7] == BC_PRODUTO .And. aAponta[nX,2] == BC_LOTECTL
					EstornaSBC(cOP,SBC->BC_NUMSEQ)
					Exit
				EndIf
				dbSkip()
			EndDo
		EndIf
	EndIf
Next nX

//--> Estornar as OPs apontadas
For nX := 1 To Len(aAponta)
	If aAponta[nX,6] <> "Perda"
		//--> Subtrair o Empenho do Lote do SLIT
		If SB8->(dbSeek(xFilial("SB8")+aDados[2]+"01"+aAponta[nX,7]))
			RecLock("SB8",.F.)
			Replace SB8->B8_EMPENHO With (SB8->B8_EMPENHO - aAponta[nX,4]),;
					SB8->B8_EMPENH2 With (SB8->B8_EMPENH2 - aAponta[nX,4])
			MsUnLock()
		EndIf
		
		//--> Monta Array para a Rotina Automatica
		aSD3   := {	{"D3_TM"     	,	"300"      		,Nil}	,;
					{"D3_COD"    	,	aAponta[nX,7]	,Nil}	,;
					{"D3_UM"		,	aDados[5]  		,Nil}	,;
					{"D3_QUANT"		,	aAponta[nX,4]	,Nil}	,;
					{"D3_OP"     	,	cOP        		,Nil}	,;
					{"D3_LOCAL"		,	"01"       		,Nil}	,;
					{"D3_SEGUM"		,	aDados[7]  		,Nil}	,;
					{"D3_QTSEGUM"	,	aAponta[nX,4]	,Nil}	,;
					{"D3_PERDA"  	,	0          		,Nil}	,;
					{"D3_LOTECTL"	,	aAponta[nX,2]	,Nil}	,;
					{"D3_DTVALID"	,	aAponta[nX,3]	,Nil}	} 
	
		MsExecAuto({|x,y| Mata250(x,y)},aSD3,5) 	//Estornar
		If lMsErroAuto
			lErro := .T.
			Aviso("ATENÇÃO !","Ocorreu um Erro ao Estornar a OP " + cOP + " do Lote " + aAponta[nX,2],{" Sair >> "},2,"Estorno da OP")
			//If lViewRot
				MostraErro()
			//EndIf
		EndIf	
	EndIf	
Next nX	
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ A002Data ³ Autor ³ Larson Zordan         ³ Data ³28.06.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Faz a consistencia da data Final digitada                  ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ A002Data(ExpD1,ExpD2,ExpC1,ExpC2)                          ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpD1 - Data Inicial da Producao                           ³±±
±±³          ³ ExpD2 - Data Final da Producao (Apontamento)               ³±±
±±³          ³ ExpC1 - Hora Inicial da Producao                           ³±±
±±³          ³ ExpC2 - Hora Final da Producao (Apontamento)               ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function A002Data(dDtIni,dDtEnt,cHrIni,cHrEnt)
Local lRet := .T.

If dDtIni < dDtEnt
	Help(" ",1,"A680DATA")
	lRet := .F.
EndIf

If lRet .And. !Empty(dDtIni) .And. !Empty(dDtEnt) .And. !Empty(cHrIni) .And. !Empty(cHrEnt)
	If (DtoS(dDtIni) + cHrIni) >= (DtoS(dDtEnt) + cHrEnt)
		Help(" ",1,"A680Hora")
		Return(.F.)
	EndIf
EndIf

If ElapTime(cHrIni+":00",cHrEnt+":00") <= "00:00:00"
	Help(" ",1,"A680Hora")
	lRet := .F.
EndIf

Return( lRet )

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³Fun‡…o    ³ A002Hora ³ Autor ³ Larson Zordan         ³ Data ³28.06.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descri‡…o ³ Faz a consistencia dos Horarios digitados.                 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Sintaxe   ³ A002Hora(ExpD1,ExpD2,ExpC1,ExpC2,ExpC3)                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Parametros³ ExpD1 - Data Inicial da Producao                           ³±±
±±³          ³ ExpD2 - Data Final da Producao (Apontamento)               ³±±
±±³          ³ ExpC1 - Hora Inicial da Producao                           ³±±
±±³          ³ ExpC2 - Hora Final da Producao (Apontamento)               ³±±
±±³          ³ ExpC3 - Indica qual hora esta vindo a validacao            ³±±
±±³          ³         I - Hora Inicial                                   ³±±
±±³          ³         F - Hora Final                                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Uso      ³ COMAFAL                                                    ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/

Static Function A002Hora(dDtIni,dDtEnt,cHrIni,cHrEnt,cHora)
Local cCampo 	:=  If(cHora == "F", cHrEnt, cHrIni)
Local lRet   	:= .T.
Local nEndereco
Local nHora
Local nMinutos
Local nPos   	:= If(cHora == "F", At(":",cHrEnt), At(":",cHrIni))

If Val(Substr(cCampo,nPos+1,2)) >= 60
	nHora        	:= Val(Substr(cCampo,1,nPos-1))+1
	nMinutos		:= Val(Substr(cCampo,nPos+1,2))-60
EndIf

If !(Empty(Substr(cHrIni,1,nPos-1)) .Or. Empty(Substr(cHrIni,nPos)) .Or. Empty(Substr(cHrEnt,1,nPos-1)) .Or. Empty(Substr(cHrEnt,nPos)))
	If Val(Substr(cCampo,1,nPos-1)) > 24
		Help(" ",1,"A680HRINVL")
		lRet:=.F.
	ElseIf Val(Substr(cCampo,1,nPos-1)) == 24
		If Val(Substr(cCampo,nPos+1,2)) >= 60
			Help(" ",1,"A680HRINVL")
			lRet:=.F.
		EndIf
	EndIf
	If lRet .And. cHrEnt  <= cHrIni .And. dDtEnt == dDtIni
		Help(" ",1,"A680Hora")
		lRet := .F.
	EndIf
EndIf

If ElapTime(cHrIni+":00",cHrEnt+":00") <= "00:00:00"
	Help(" ",1,"A680Hora")
	lRet := .F.
EndIf

Return(lRet)
