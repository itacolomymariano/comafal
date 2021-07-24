#Include "RwMake.ch"     

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÚÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄ¿±±
±±³ Programa ³ M410LIOK ³ Autor ³ Eduardo Zanardo       ³ Data ³16.01.2004³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄ´±±
±±³Descricao ³ Ponto de entrada na função a410TudOk() do programa Mata410 ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³Uso       ³ COMAFAL                                                    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±± 
±±³ ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÂÄÄÄÄÄÄÂÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ PROGRAMADOR  ³ DATA   ³HLPDSK³  MOTIVO DA ALTERACAO                   ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ E. Zanardo   ³16/01/04³      ³ Desenvolvimento incial do programa.    ³±±
±±ÃÄÄÄÄÄÄÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÅÄÄÄÄÄÄÅÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ´±±
±±³ Larson       ³16/02/04³      ³ Nao permite vender bobina sem informar ³±±
±±³              ³        ³      ³ o lote da bobina.                      ³±±
±±ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÁÄÄÄÄÄÄÁÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß*/
User Function M410LIOK()

Local aArea		:= GetArea()
Local cCidade 	:= "" 
Local cUF 		:= ""
Local lRet 		:= .T.
Local nI 		:= 0
Local nPosQTD 	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_QTDVEN" })  
Local nPosProd	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_PRODUTO"})  
Local nPosLote	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_LOTECTL"})  
Local nPosTes	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_TES"}) 
Local nPosValor	:= aScan(aHeader, {|x| AllTrim(x[2]) == "C6_VALOR"})
Local nPosDel   := Len( aHeader ) + 1
Local nTotQtd 	:= 0
Local nVLFrete 	:= 0 
Local nValor	:= 0
Local cTes		:= Alltrim(GetMv("MV_TESSAI")) 
Local cTesEXT	:= Alltrim(GetMv("MV_TESEXT"))
Local nOpc		:= 0

If !Empty(M->C5_TPFRETE)
	For nI := 1 to Len(aCols)
		If !aCols[nI,nPosDel]
			nTotQtd += aCols[nI][nPosQTD]     
			nValor += aCols[nI][nPosValor]    
		EndIf		
	Next nI
	
	If M->C5_TPFRETE == "C"
		If nTotQtd >0 	    
			If xFilial("SC5") <> "04"
				If M->C5_TIPO $"B#D"   
					dbSelectArea("SA2") 
					dbSetOrder(1)  
					If MsSeek(xFilial("SA2")+M->C5_CLIENTE+M->C5_LOJACLI)
						cCidade := UPPER(SA2->A2_MUN)
						cUF := SA2->A2_EST
					Else
						MsgBox("Fornecedor nao cadastrado. <SA2>")
						lRet := .F.				
					EndIf
				Else
					dbSelectArea("SA1")
					dbSetOrder(1)
					If MsSeek(xFilial("SA1")+M->C5_CLIENTE+M->C5_LOJACLI)
						cCidade := UPPER(SA1->A1_MUN)
						cUF := SA1->A1_EST
					Else                  
						MsgBox("Cliente nao cadastrado. <SA1>")
						lRet := .F.					
					EndIf			
				EndIf
				If lRet
					dbSelectArea("SZ0") 
					dbSetOrder(1)  
				    If MsSeek(xFilial("SZ0")+cCidade+cUF)
						nVLFrete := NoRound(SZ0->Z0_FRETE * nTotQtd,2)
				    Else 
			    		MsgBox("Nao existe cadastro de frete para este municipio. <SZ0>")
						lRet := .F.							    
				    EndIf
				EndIf	
			Else 
				nVLFrete := NoRound(nValor * 0.13,2)			
			EndIf	
		EndIf	
	EndIf
	
	M->C5_FRETE := nVLFrete
	M->C5_PESOL := nTotQtd
Else
	MsgBox("Tipo do frete nao preenchido.")
	lRet := .F.							    
Endif

//--> Valida a venda para BOBINA
If lRet .And. AllTrim(Posicione("SB1",1,xFilial("SB1")+aCols[n,nPosProd],"B1_GRUPO")) == "10" .AND. Posicione("SF4",1,xFilial("SF4")+aCols[n,nPosTes],"F4_ESTOQUE") == "S" .and. nPosLote > 0
	If Empty(aCols[n,nPosLote])
		Aviso("ATENCAO !",	"Para Vender Uma Bobina, sera necessario informar o Lote da Bobina. "+;
							"Posicione no campo LOTE e tecle <F4> para selecionar o lote da bobina.",{ " << Voltar "},2,"LOTE DA BOBINA")
		lRet := .F.							
	EndIf
EndIf

If lRet 
	DbSelectArea("SE4")
	DbSetOrder(1)
	DbSeek(xFilial("SE4")+M->C5_CONDPAG)
 	If (SE4->E4_X_PJURO > 0 .And. M->C5_X_MENBO <> "S") .Or. Empty(M->C5_X_MENBO)
		Aviso("ATENCAO !","A informacao do Campo Boleto nao confere com a condicao de pagamento informada",{ " << Voltar "},2,"Cond.Pagto X Boleto")
		lRet := .F.									 	
 	EndIf
EndIf  

If (aCols[n,nPosTes] <> cTes .and. aCols[n,nPosTes] <> cTesEXT ).and. !U_CMF001C() 
	Aviso("ATENCAO !","Usuario nao tem permissao de utilizar a TES informada",{ " << Voltar "},2,"TES")
	lRet :=.F.
EndIf          

If M->C5_TIPO <> 'N' .and. !U_CMF001C()
	Aviso("ATENCAO !","Usuario nao tem permissao de utilizar o tipo de pedido informado",{ " << Voltar "},2,"Tipo do Pedido")
	lRet :=.F.
EndIf
/*
If lRet .And. M->C5_TIPO == "B"
	nOpc := Aviso("ATENCAO !","Este pedido faz parte de uma operacao triangular?",{"Sim","Nao"})
	If nOpc == 1 .And. Empty(Alltrim(M->C5_X_PREM)) .And. M->C5_TIPO == "B"
		Aviso("ATENCAO !","Favor informar o numero do pedido de Venda no campo Nr.Ped.Rems.",{ " << Voltar "},2,"Pedido de Remmessa")
		lRet :=.F.
	ElseIf nOpc == 1 .And. !Empty(Alltrim(M->C5_X_PREM)) .And. M->C5_TIPO == "B"
		DbSelectArea("SC5")
		DbSetOrder(1)
		If DbSeek(xFilial("SC5")+M->C5_X_PREM)
			If Empty(Alltrim(SC5->C5_X_PREM))
				RecLock("SC5",.F.)
				SC5->C5_X_PREM := M->C5_NUM
				MsUnlock()				
			ElseIf SC5->C5_X_PREM <> M->C5_NUM 
				Aviso("ATENCAO !","No Pedido de Venda ja existe um Pedido de remessa informado. Favor corrigir a informacao do campo Nr.Ped.Rems.",{ " << Voltar "},2,"Pedido de Remmessa")
				lRet :=.F.
			EndIf	
		Else      
			Aviso("ATENCAO !","Pedido de Venda nao encontrado. Favor corrigir a informacao do campo Nr.Ped.Rems.",{ " << Voltar "},2,"Pedido de Remmessa")
			lRet :=.F.
		EndIf
		
	EndIf	 
EndIf
*/
RestArea(aArea)

If !l410Auto
	ParamIXB:Refresh()
EndIf

Return(lRet)
