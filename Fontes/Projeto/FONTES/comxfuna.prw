#Include "PROTHEUS.CH"
#Include "TOPCONN.CH"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑฺฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฟฑฑ
ฑฑณ Programa ณ ConvLin  ณ Autor ณ Larson Zordan         ณ Data ณ21.08.2003ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณDescricao ณ Converte Linha em Centimetros e Vice-Versa                 ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe   ณ ConvLin(ExpN1,ExpC1)                                       ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณParametrosณ ExpN1 = Numero da linha (numero ou centimetros)            ณฑฑ
ฑฑณ          ณ ExpC1 = Tipo da Conversao (L=Linha e C=Centimetro)         ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณUso       ณ COMAFAL                                                    ณฑฑ
ฑฑภฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤูฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function ConvLin(nLin,cTipo)
Local nLinConv := 0
Local nX
              
//--> Converte em Linha em Centimetros
If cTipo == "C"
	nLinConv := 0.6
	For nX := 1 To nLin
		nLinConv += 0.4
	Next nx	
EndIf

Return(nLinConv)     

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณGetModulo บAutor  ณLarson Zordan       บ Data ณ  10/02/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณRetorna array contendo os modulos que o usuario tem utiliza,บฑฑ
ฑฑบ          ณo que esta ativo no momento e o nivel de acesso             บฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe   ณ GetModulo()  			                                  ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณRetorno	 ณ aArray					                                  ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑบUso       ณ EXPRESS                                                    บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function GetModulo()
Local nX 		:= 0
Local aUser   	:= PswRet(3)
Local aModulo 	:= RetModName()
Local aAtivos 	:= {}

For nX := 1 To Len(aUser[1])
	If Val(Left(aUser[1,nX],2)) == aModulo[nX,1] .And. SubStr(aUser[1,nX],3,1) <> "X"
		aAdd(aAtivos,{ aModulo[nX,2], SubStr(aUser[1,nX],3,1), (AllTrim(aModulo[nX,2])=="SIGA"+cModulo) } )
	EndIf
Next nX         

Return(aAtivos)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณForAtivo  บAutor  ณSabrina Moreira     บ Data ณ 23/10/03    บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณValida se o Fornecedor esta Ativo.					      บฑฑ
ฑฑบออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function ForAtivo(nRotina)
Local lRet     := .T.

If nRotina == 1
	If cTipo <> "D"
		If SA2->A2_STATUS <> "1"
			Aviso("Fornecedor Inativo !!!","Favor verificar o cadastro deste Fornecedor",{"<< Voltar"}) 	
			lRet := .F.
		EndIf
	EndIf
ElseIf nRotina == 2
	If SA2->A2_STATUS <> "1"
		Aviso("Fornecedor Inativo !!!","Favor verificar o cadastro deste Fornecedor",{"<< Voltar"}) 	
		lRet := .F.
	EndIf	
EndIf	    

Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑฺฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฟฑฑ
ฑฑณFun…o    ณ  PesqLbx   ณ Autor ณ Larson Zordan       ณ Data ณ 04.02.04 ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณDescri…o ณ Pesquisa na ListBox                                        ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe   ณ PesqLbx(ExpO1,ExpA1,ExpA2,ExpA3,ExpL1)                     ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณParametrosณ ExpO1 - Objeto da ListBox                                  ณฑฑ
ฑฑณ          ณ ExpA1 - Array da ListBox                                   ณฑฑ
ฑฑณ          ณ ExpA2 - Array das Opcoes da ComboBox da Pesquisa           ณฑฑ
ฑฑณ          ณ ExpA3 - Posicoes da ListBox conforme Opcoes da ComboBox    ณฑฑ
ฑฑณ          ณ ExpL1 - Ordena a ListBox conforme ordem da pesquisa        ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณ Uso      ณ COMAFAL      		                                      ณฑฑ
ฑฑภฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤูฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function PesqLbx(oLbx,aLbx,aCbx,aPosLbx,lIndexa)
Local oDlg
Local oCbx
Local cCbx   := ""
Local nCbx   := 1
Local cCampo := Space(15)
Local lRet   := .F.
Local nPos   := 0

Default lIndexa := .F.

Define MsDialog oDlg From 5,5 To 14,50 Title "Pesquisa"

@ 0.6,1.3 MsComboBox oCbx Var cCbx Items aCbx 	Size 165,44 Of oDlg
@ 2.1,1.3 MsGet cCampo Picture "@!"				Size 165,10 Of oDlg

Define SButton From 50,114 Type 1 Enable Of oDlg Action (lRet:=.T.,oDlg:End()) 
Define SButton From 50,146 Type 2 Enable Of oDlg Action oDlg:End() 			 

Activate MsDialog oDlg Centered

If lRet
	//--> Posicao no aCbx para busca na ListBox
	nCbx := aScan(aCbx, cCbx)

	//--> Ordena a ListBox conforme opcao da Pesquisa
	If lIndexa
		aLbx := aSort(aLbx,,,{ |x,y| x[ aPosLbx[nCbx] ] < y[ aPosLbx[nCbx] ] })
		oLbx:nAt := 1
	EndIf

	If !Empty(cCampo)
		//--> Pesquina na ListBox
		nPos := aScan(aLbx, { |x| AllTrim(x[ aPosLbx[nCbx] ]) == AllTrim(cCampo) } )

		//--> Posiciona a ListBox
		If nPos == 0
			Help(" ",1,"PESQ01")
			lRet := .F.
		Else
			oLbx:nAt := nPos
			oLbx:Refresh()
		EndIf	
	Else
		oLbx:Refresh()
	EndIf	
EndIf
Return(lRet)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑฺฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฟฑฑ
ฑฑณFun…o    ณ  MsgGet    ณ Autor ณ Larson Zordan       ณ Data ณ18.02.2004ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณDescri…o ณ Get em Linha Generica com Dialog                           ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe   ณ MsgGet(ExpC1,ExpC2,@ExpC3,ExpC4, cF3)                      ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณParametrosณ ExpC1 - Titulo da Dialog                                   ณฑฑ
ฑฑณ          ณ ExpC2 - Texto Sobre o Get                                  ณฑฑ
ฑฑณ          ณ ExpC3 - Variavel Passada Por Referencia                    ณฑฑ
ฑฑณ          ณ ExpC4 - Picture da Campo                                   ณฑฑ
ฑฑณ          ณ ExpC5 - Consulta SXB                                       ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณ Uso      ณ COMAFAL      		                                      ณฑฑ
ฑฑภฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤูฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function MsgGet( cTitle, cText, uVar, cPict, cF3 )
Local oBtn
Local oDlg
Local uTemp   := uVar
Local lOk     := .F.

Default cText := ""
Default cPict := "@!"
Default cF3   := ""

Define Dialog oDlg From 10, 20 To 18, 59.5 Title cTitle

@ 0.8, 4 Say   cText 								Size 250,10 Of oDlg 
If Right(cPict,1) <> "9"
	If Empty(cF3)
		@ 1.6, 4 MsGet uTemp Picture cPict 		  	Size 120,10 Of oDlg
	Else
		@ 1.6, 4 MsGet uTemp Picture cPict F3 cF3 	Size 120,10 Of oDlg
	EndIf	
Else
	@ 1.6, 4 MsGet uTemp Picture cPict 		  		Size  70,10 Of oDlg
EndIf
@ 0, 0 BitMap ResName "LOGIN" Size 25, 62 Of oDlg Adjust NoBorder When .F. Pixel

Define SButton oBtn From 40, 90 Type  1 Enable Of oDlg Action (lOk:=.T.,oDlg:End())
Define SButton oBtn From 40,120 Type  2 Enable Of oDlg Action (oDlg:End())

Activate Dialog oDlg Centered

If lOk
	uVar := uTemp
EndIf

Return(lOk)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณGravaSZ8  บAutor  ณEduardo Zanardo     บ Data ณ 18/02/04    บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณGrava as NF's excluidas									  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function GravaSZ8(cDoc,cSerie,cCliFor,cLoja,cTipo,cMotivo)
Local aArea 	:= GetArea()

DbSelectArea("SZ8")
RecLock("SZ8",.T.)
SZ8->Z8_FILIAL 	:= xFilial("SZ8")
SZ8->Z8_DOC 	:= cDoc
SZ8->Z8_SERIE 	:= cSerie
SZ8->Z8_CLIFOR 	:= cCliFor
SZ8->Z8_LOJA 	:= cLoja
SZ8->Z8_TIPO 	:= cTipo
SZ8->Z8_MOTIVO 	:= cMotivo
SZ8->Z8_USUARIO	:= Alltrim(Upper(Substr(cUsuario,7,15))) 
SZ8->Z8_DATA 	:= dDataBase
SZ8->Z8_HORA 	:= Time()
MsUnlock()

RestArea(aArea)

Return(.T.)



/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณRelatSZ8  บAutor  ณEduardo Zanardo     บ Data ณ  29/02/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณRelatorio das NF's excluidas								  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
     
User Function RelatSZ8()
Local aOrd           := {}
Local cDesc1         := ""
Local cDesc2         := ""
Local cDesc3         := ""
Local cPict          := ""
Local titulo         := "Relatorio de Exclusao de NF's. - COMAFAL/"+SM0->M0_ESTENT
Local Cabec1         := ""
Local Cabec2         := ""
Local Imprime        := .T.
Local nVeiImp		 := 0
Local cString		 := "SZ8" 
Local NomeProg		 := "RELSZ8" 
Local lAbortPrint 	 := .F.
Local limite      	 := 132
Local tamanho     	 := "M"      

Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey    := 0
Private cbcont     	:= 0
Private m_pag      	:= 01
Private wnrel      	:= "RELSZ8" // Coloque aqui o nome do arquivo usado para impressao em disco
	
wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.T.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

nTipo := If(aReturn[4]==1,15,18)

RptStatus({|| IMPSZ8(@Cabec1,@Cabec2,Titulo,tamanho,nTipo,NomeProg) },Titulo)

Return(.T.)                  

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณImpSZ8  บAutor  ณEduardo Zanardo     บ Data ณ 18/02/04    บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณRelatorio das NF's excluidas								  บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/

Static Function ImpSZ8(Cabec1,Cabec2,Titulo,tamanho,nTipo,NomeProg)

Local aArea 	:= GetArea()
Local cQuery 	:= ""
Local cAliasSZ8 := "cAliasSZ8" 
Local nSZ8		:= 0
Local aStruSZ8	:= SZ8->(dbStruct())   
Local aL 		:= Array(05)    
Local nLin		:= 80

	      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
	      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
aL[01]	:=	"|=========================================================================================================|"
aL[02]	:=	"|NF     |Serie|Tipo|Data       |Hora  |Usuario         |Motivo                                            |"
aL[03]	:=	"|---------------------------------------------------------------------------------------------------------|"	
aL[04]	:=	"|###### |###  |#   |########## |##### |############### |##################################################|"			
aL[05]	:=	"|=========================================================================================================|"

cQuery := "SELECT * "
cQuery += "FROM "
cQuery += RetSqlName('SZ8') 
cQuery += " WHERE "                                  
cQuery += "Z8_FILIAL = '" + xFilial("SZ8")+"'	AND "
cQuery += "Z8_DATA = '"  + DtoS(dDataBase) + "' 	AND "
cQuery += "D_E_L_E_T_ <> '*' " 
cQuery += "ORDER BY Z8_TIPO,Z8_DOCZ8_SERIE,Z8_USUARIO"
		
cQuery:=ChangeQuery(cQuery)
		
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSZ8,.T.,.T.)
		
For nSZ8 := 1 To len(aStruSZ8)
	If aStruSZ8[nSZ8][2] <> "C" .And. FieldPos(aStruSZ8[nSZ8][1])<>0
		TcSetField(cAliasSZ8,aStruSZ8[nSZ8][1],aStruSZ8[nSZ8][2],aStruSZ8[nSZ8][3],aStruSZ8[nSZ8][4])
	EndIf
Next nSZ8
		
DbSelectArea(cAliasSZ8)

Do While (cAliasSZ8)->(EOF())
	If nLin > 55
		Cabec1 := "Periodo: " + cMonth(dDatabase)+"/"+Str(Year(dDatabase))
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1
		FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)                                        
	EndIf          
	FmtLin({(cAliasSZ8)->Z8_DOC,(cAliasSZ8)->Z8_SERIE,(cAliasSZ8)->Z8_TIPO,(cAliasSZ8)->Z8_DATA,;
		(cAliasSZ8)->Z8_HORA,(cAliasSZ8)->Z8_USUARIO,;
		Substr((cAliasSZ8)->Z8_MOTIVO,1,50)},aL[4],"",,@nLin)                              	

	(cAliasSZ8)->(DbSkip())          
EndDo

FmtLin(,{aL[5]},,,@nLin)

(cAliasSZ8)->(DbCloseArea())          

SetPrc(0,0)

Set Device To Screen

If aReturn[5]==1
	dbCommitAll()
	Set Printer To
	OurSpool(wnrel)
Endif

MS_FLUSH()

RestArea(aArea)

Return(.T.)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑฺฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฟฑฑ
ฑฑณFun…o  ณRetX3Comb ณ Autoraณ Zanardo				    ณ Data ณ 03/03/04 ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณDescri…o ณMonta Array de acordo com Item selecionado Sx3 Opcoes       ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe   ณRetX3Comb                 				                  ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณUso       ณGenerico  				                                  ณฑฑ
ฑฑภฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤูฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
*/
User Function RetX3Comb(cCombo,nInitCBox,nTam,nTamCpo,cInitValue)

Local cDescItem		:= ""
Local cItem			:= "" 
Local aCombo		:= {}
Local nTamDesc		:= 0
Local cDesc			:= ""
Local ni			:= 0
Local lInitValue 	:= .F.
Local lHasInit		:= !Empty(cInitValue)

nTam:=0
If !Empty(Alltrim(cCombo))
	While .T.
		lInitValue:=.F.
		cDescItem := AllTrim(Substr(cCombo,1,At(";",cCombo)-1))
		cDesc := AllTrim(Subs(cDescItem,AT("=",cDescItem)+1))
		nTamDesc := Iif(Len(cDesc)>nTamDesc,Len(cDesc),nTamDesc)
		
		cItem := AllTrim(Substr(cDescItem,1,At("=",cDescItem)-1))
		
		If ( "&" $ cItem ) .And. ( ! lHasInit )
			lInitValue:=.T.
		EndIf
		
		cItem := StrTran(cItem,"&","",1)
		cItem := cItem+Space(nTamCpo-Len(cItem))
		cItem := Subs(cItem,1,nTamCpo)
		
		IF Len(cItem) == 0
			cDescItem := " "+cDescItem
		Endif
		
		If lInitValue
			cInitValue:= cItem
		EndIf
		
		If !Empty(cDescItem)
			AADD(aCombo,{cDescItem,cItem,cDesc})
		EndIf
		
		cCombo := Substr(cCombo,At(";",cCombo)+1)
		If At(";",cCombo) == 0
			cDescItem := AllTrim(cCombo)
			cDesc := AllTrim(Subs(cDescItem,AT("=",cDescItem)+1))
			nTamDesc := Iif(Len(cDesc)>nTamDesc,Len(cDesc),nTamDesc)
			cItem := AllTrim(Substr(cDescItem,1,At("=",cDescItem)-1))
			cItem := StrTran(cItem,"&","",1)
			IF Len(cItem) == 0
				cDescItem := " "+cDescItem
			Endif
			cItem := cItem+Space(nTamCpo-Len(cItem))
			cItem := Subs(cItem,1,nTamCpo)
			If !Empty(cDescItem)
				AADD(aCombo,{cDescItem,cItem,cDesc})
			EndIf
			Exit
		EndIf
		If Len(Substr(cDescItem,At("=",cDescItem)+1)) > nTam
			nTam := Len(AllTrim(cDescItem))
		EndIf
	End
	
	For ni:= 1 to Len(aCombo)
		aCombo[ni,3] := aCombo[ni,3]+Space(nTamDesc-Len(aCombo[ni,3]))
	Next
	
	If Len(Substr(cDescItem,At("=",cDescItem)+1)) > nTam
		nTam := Len(AllTrim(cDescItem))
	EndIf
	AADD(aCombo,{Space(nTam),Space(Len(aCombo[1,2])),Space(nTamDesc) } )
	nInitCBox := Ascan(aCombo,{|x| "&" $ x[1]})
	If nInitcBox > 0
		aCombo[nInitCbox] := {AllTrim(StrTran(aCombo[nInitCBox,1],"&","",1)),aCombo[nInitCBox,2],aCombo[nInitCBox,3]}
	Endif
	IF lHasInit
		nInitCBox := Ascan(aCombo,{|x| x[2] == cInitValue})
	Endif
	IF nInitcBox == 0
		nInitcBox := Len(aCombo)
	   cInitValue := aCombo[nInitcBox,2]
	EndIf
EndIf
Return aCombo
