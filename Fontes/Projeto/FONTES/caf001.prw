#Include "PROTHEUS.CH"
#Include "TopConn.Ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCAF001    บ Autor ณ Zanardo            บ Data ณ  16/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณControle de Balanca - Entrada e Saida. Antigo FAT131        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function CAF001()

Local aCampos     := {}

Private	aL		:= {}
Private oDlg
Private cTipos 	:= ""
Private nOpcMen := 0 
Private aHeader := {}
Private aCols   := {}
Private aRotina	:= {{"Pesquisar"	,"AxPesqui"  	,0,1},;
					{"Entrada"    	,"U_CAF001A"    ,0,2},;
					{"Saida"      	,"U_CAF001B"		,0,4},;
					{"Reimpressao"	,"U_CAF001F"		,0,2},;
					{"Estorno"    	,"U_CAF001D"		,0,5},;
					{"Relat.Diverg.","U_CAF001H"		,0,2}}

aAdd(aCampos, {"Placa do Veํculo      ","Z5_PLACA",'C',7,0,Pesqpict("SZ5","Z5_PLACA")})
aAdd(aCampos, {"Dt. Entrada ","Z5_DENTR",'D',8,0,nil})
aAdd(aCampos, {"Hr. Entrada ","Z5_HENTR",'D',8,0,nil})
aAdd(aCampos, {"Peso Entrada","Z5_PENTR",'N',8,2,Pesqpict("SZ5","Z5_PENTR")})
aAdd(aCampos, {"Observacao  ","Z5_OBSER",'C',72,0,nil})
aAdd(aCampos, {"Tipo        ","Z5_TIPOS",'D',08,0,Pesqpict("SZ5","Z5_TIPOS")})

dbSelectArea("SZ5")        

mBrowse(6,1,22,75,"SZ5",aCampos,"Z5_DSAID")

Return



/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001A   บ Autor ณ Zanardo            บ Data ณ  16/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Entrada das informacoes do Veiculo e suas notas de entrada บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function CAF001A(cAlias,nReg,nOpc)

Local aPosObj   := {}
Local aObjects  := {}
Local aSize     :=MsAdvSize()
Local aInfo     :={aSize[1],aSize[2],aSize[3],aSize[4],3,3}
Local cPlaca 	:= Space(7)      
Local dData 	:= dDatabase
Local cHora 	:= Time()
Local cObs 		:= Space(74)
Local nOpca  	:= 0    
Local nUsado	:= 0 
Local nX		:= 0
Local aCbx		:= {}
Local aCbo		:= {}
Local nCbx		:= 0
Local oGetD

Private lAutoriza := .F.
Private nDiverg	  := 0
Private nPeso 	  := 0

aHeader := {}
aCols	:= {}         
cTipos 	:= "" 
nOpcMen := nOpc
DbSelectArea("SX3")
SX3->(DbSetOrder(2))
If SX3->(MsSeek("Z5_MOTIV"))
	aCbx	:=	U_RetX3Comb(SX3->X3_CBOX,,,SX3->X3_TAMANHO)
	For nCbx := 1 To Len(aCbx)
		If !Empty(Alltrim(aCbx[nCbx][1]))
			aAdd(aCbo,aCbx[nCbx][1])
		EndIf	
	Next nCbx
Endif

SX3->(DbSetOrder(1))
MsSeek("SZ6")
While ( !Eof() .And. (SX3->X3_ARQUIVO == "SZ6") )
	IF X3USO(SX3->X3_USADO) .and. X3_BROWSE == "S"
		nUsado++
		Aadd(aHeader,{ AllTrim(X3Titulo()),;
			SX3->X3_CAMPO	,;
			SX3->X3_PICTURE,;
			SX3->X3_TAMANHO,;
			SX3->X3_DECIMAL,;
			SX3->X3_VALID	,;
			SX3->X3_USADO	,;
			SX3->X3_TIPO	,;
			SX3->X3_F3 ,;
			SX3->X3_CONTEXT,;
			SX3->X3_CBOX,;
			SX3->X3_RELACAO} )
	EndIf
	dbSelectArea("SX3")
	dbSkip()
EndDo

aadd(aCOLS,Array(nUsado+1))
For nX := 1 To nUsado
	aCols[1][nX] := CriaVar(aHeader[nX][2])
Next nX
aCOLS[1][Len(aHeader)+1] := .F.
aObjects := {}
AADD(aObjects,{100,045,.T.,.F.,.F.})
AADD(aObjects,{100,100,.T.,.T.,.F.})
aPosObj := MsObjSize( aInfo, aObjects )

dbSelectArea("SZ6")

DEFINE MSDIALOG oDlg TITLE "Entrada de Veiculo" From aSize[7],0 to aSize[6],aSize[5] of oMainWnd PIXEL

@ 15,010 Say "Placa: "  Of oDlg  Pixel
@ 15,030 GET cPlaca PICTURE "XXX9999" Valid NaoVazio(Alltrim(cPlaca)) Of oDlg  Pixel

@ 15,72 Say "Motivo: "  Of oDlg  Pixel
@ 15,92 COMBOBOX cTipos ITEMS aCbo SIZE 50,50 Valid NaoVazio(Alltrim(cTipos)) Of oDlg  Pixel

@ 15,150 Say "Peso Entrada: "  Of oDlg  Pixel
@ 15,190 Get nPeso Picture "@E 999,999.999" When .F. Of oDlg  Pixel

@ 15,240 Say "Data Entrada: "  Of oDlg  Pixel
@ 15,280 Get dData PICTURE "@D" SIZE 36,10 When .F.  Of oDlg  Pixel

@ 15,320 Say "Hora Entrada: "  Of oDlg  Pixel
@ 15,360 Get cHora PICTURE "99:99" When .F. Of oDlg  Pixel

@ 30,010 Say "Observacao "  Of oDlg  Pixel
@ 40,010 Get cObs SIZE 210,10  Of oDlg  Pixel

oGetD:= MSGetDados():New(aPosObj[2,1],aPosObj[2,2],aPosObj[2,3],aPosObj[2,4],nOpc,"U_CAF001LinOk(n)","AllwaysTrue","",.T.)

ACTIVATE MSDIALOG oDlg ON INIT EnchoiceBar(oDlg, {|| (If(U_CAF001TdOk(nOpc,0) .And. nPeso > 0,(nOpca := 1,oDlg:End()),)) },{||oDlg:End()},,{{"BMPINCLUIR",{|| Balanca()},"Balanca"}})

If nOpca == 1  
	If CAF001E(1,Padr(cPlaca,7),nPeso,dData,Padr(cHora,5),cObs)
		U_CAF001F(cAlias,nReg,nOpc,1,Padr(cPlaca,7),nPeso,dData,Padr(cHora,5),cObs)
	EndIf
EndIf

Return(.T.)


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001B   บ Autor ณ Zanardo            บ Data ณ  16/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Saida   das informacoes do Veiculo e suas notas            บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function CAF001B(cAlias,nReg,nOpc)

Local aPosObj   := {}
Local aObjects  := {}
Local aSize     := MsAdvSize()
Local aInfo     := {aSize[1],aSize[2],aSize[3],aSize[4],3,3}
Local cPlaca 	:= Space(7)      
Local dData 	:= dDatabase
Local cHora 	:= Time()
Local cObs 		:= Space(74)
Local nOpca  	:= 0    
Local nUsado	:= 0 
Local nEstrado 	:= 0
Local aCampos 	:= {}
Local nX		:= 0
Local oGetD
Local aCbx		:= {}
Local aCbo		:= {}
Local nCbx		:= 0
Local cMotivo 	:= 0

Private lAutoriza := .F.
Private nDiverg	  := 0
Private nPeso 	:= 0

aHeader := {}
aCols	:= {}         
cTipos 	:= SZ5->Z5_MOTIV 
nOpcMen := nOpc

If SZ5->Z5_MOTIV == "1"
	cMotivo := "Carregamento"
ElseIf	SZ5->Z5_MOTIV == "2" 
	cMotivo := "Descarregamento"
Else
	cMotivo := "Outros"
EndIf         

DbSelectArea("SX3")
SX3->(DbSetOrder(2))
If SX3->(MsSeek("Z5_MOTIV"))
	aCbx	:=	U_RetX3Comb(SX3->X3_CBOX,,,SX3->X3_TAMANHO)
	For nCbx := 1 To Len(aCbx)
		aAdd(aCbo,aCbx[nCbx][1])
	Next nCbx
Endif

// Carrega Campos da Get Dados
aAdd(aCampos,{"Z6_TIPON"})
aAdd(aCampos,{"Z6_PEDIDO"})
aAdd(aCampos,{"Z6_FORN"})
aAdd(aCampos,{"Z6_PESO"})

dbSelectArea("SX3") 
DbSetOrder(2)
For nX := 1 To Len(aCampos)
	If dbSeek(aCampos[nX,1])
		IF X3USO(SX3->X3_USADO)
			nUsado++
			Aadd(aHeader,{ AllTrim(X3Titulo()),;
				SX3->X3_CAMPO	,;
				SX3->X3_PICTURE,;
				SX3->X3_TAMANHO,;
				SX3->X3_DECIMAL,;
				SX3->X3_VALID	,;
				SX3->X3_USADO	,;
				SX3->X3_TIPO	,;
				SX3->X3_F3,;
				SX3->X3_CONTEXT } )
		EndIf
	EndIf	
Next nX

aadd(aCOLS,Array(nUsado+1))
For nX := 1 To nUsado
	aCols[1][nX] := CriaVar(aHeader[nX][2])
Next nX
aCOLS[1][Len(aHeader)+1] := .F.
aObjects := {}
AADD(aObjects,{100,065,.T.,.F.,.F.})
AADD(aObjects,{100,100,.T.,.T.,.F.})
aPosObj := MsObjSize( aInfo, aObjects )

DEFINE MSDIALOG oDlg TITLE "Saida de Veiculo" From aSize[7],0 to aSize[6],aSize[5] of oMainWnd PIXEL
@ 15,010 Say "Placa: "  Of oDlg  Pixel
@ 15,030 GET SZ5->Z5_PLACA PICTURE "AAA9999" When .F.  Of oDlg  Pixel

@ 15,075 Say "Peso Entrada: "  Of oDlg  Pixel
@ 15,120 Get SZ5->Z5_PENTR Picture "@E 999,999.99"  When .F.  Of oDlg  Pixel

@ 15,170 Say "Data Entrada: "  Of oDlg  Pixel
@ 15,210 Get SZ5->Z5_DENTR PICTURE "@D" SIZE 36,10 When .F.  Of oDlg  Pixel

@ 15,250 Say "Hora Entrada: "  Of oDlg  Pixel
@ 15,290 Get SZ5->Z5_HENTR PICTURE "99:99" When .F.  Of oDlg  Pixel

@ 30,010 Say "Motivo: "  Of oDlg  Pixel
@ 30,030 Get cMotivo  SIZE 50,10 When .F. Of oDlg  Pixel

@ 30,090 Say "Peso Saida: "  Of oDlg  Pixel
@ 30,120 Get nPeso PICTURE "@E 999,999.999" When .F. Of oDlg  Pixel

@ 30,170 Say "Peso Extra KG:"  Of oDlg  Pixel
@ 30,210 Get nEstrado PICTURE "@E 9,999.99"  Of oDlg  Pixel

@ 45,010 Say "Data Saida: " Of oDlg  Pixel
@ 45,050 Get dData PICTURE "@D" SIZE 36,10 When .F.  Of oDlg  Pixel

@ 45,090 Say "Hora Saida: "  Of oDlg  Pixel
@ 45,140 Get cHora PICTURE "99:99" When .F.  Of oDlg  Pixel

@ 60,010 Say "Observacao "  Of oDlg  Pixel
@ 70,010 Get cObs SIZE 210,10  Of oDlg  Pixel

oGetD:= MSGetDados():New(aPosObj[2,1],aPosObj[2,2],aPosObj[2,3],aPosObj[2,4],nOpc,"U_CAF001LinOk(n)","AllwaysTrue","",.T.)

ACTIVATE MSDIALOG oDlg ON INIT EnchoiceBar(oDlg, {|| (If(U_CAF001TdOk(nOpc,nEstrado) .And. nPeso > 0,(nOpca := 1,oDlg:End()),)) },{||oDlg:End()},,{{"BMPINCLUIR",{|| Balanca()},"Balanca"}})

If nOpca == 1  
	If CAF001E(2,Padr(SZ5->Z5_PLACA,7),nPeso,dData,Padr(cHora,5),cObs,nEstrado)
		U_CAF001F(cAlias,nReg,nOpc,2,Padr(SZ5->Z5_PLACA,7),nPeso,dData,Padr(cHora,5),cObs)
	EndIf
EndIf

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณCAF001D	บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณEstorno de Lancamento									      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function CAF001D(cALias,nReg,nOpc)

Local aArea 	:= GetArea()
Local aCampSZ5  := {}
Local nConfEst  := 0
Local cAliasSC9 := "cAliasSC9"
Local nSC9 		:= 0
Local aStruSC9 	:= SC9->(dbStruct())
Local lEstorna  := .T.
Local cMsg		:= ""

Private cCadastro := "Estorno"

aAdd(aCampSZ5,"Z5_PLACA")
aAdd(aCampSZ5,"Z5_DENTR")
aAdd(aCampSZ5,"Z5_HENTR")
aAdd(aCampSZ5,"Z5_PENTR")
aAdd(aCampSZ5,"Z5_DSAID")
aAdd(aCampSZ5,"Z5_HSAID")
aAdd(aCampSZ5,"Z5_PSAID")

dbSelectArea("SZ5")
dbGoTo(nReg)

nConfEst := AxVisual("SZ5",nReg,1,,,aCampSZ5)

If nConfEst == 1
	If !Empty(SZ5->Z5_HSAID)
		DbSelectArea("SZ6")    
		DbSetOrder(1)
		DbSeek(xFilial("SZ6")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DSAID)+SZ5->Z5_HSAID)
		While SZ6->(!EOF()) .and. ;
			 SZ5->Z5_PLACA+DTOS(SZ5->Z5_DSAID)+SZ5->Z5_HSAID == ;
			 SZ6->Z6_PLACA+DTOS(SZ6->Z6_DATA)+SZ6->Z6_HORA

			DbSelectArea("SC5")
			DbSetOrder(1)
			If MsSeek(xFilial("SC5")+SZ6->Z6_PEDIDO)
				If  !Empty(SC5->C5_NOTA )
					If lEstorna
						cMsg := "A carga possue o(s) seguinte(s) pedido(s) faturado(s): <SC5>" + Chr(13)
						cMsg += " - " + SC5->C5_NUM + Chr(13)
					Else	
						cMsg += " - " + SC5->C5_NUM + Chr(13)						
					EndIf
					lEstorna := .F.
				EndIf
			Else
				MsgAlert("O Pedido nao encontrado. <SC5>")
			EndIf            

			DbSelectArea("SZ6")    
			SZ6->(DbSkip())
			
		EndDo	
		
		If lEstorna 
		
			DbSelectArea("SZ6")    
			DbGotop()
			DbSetOrder(1)
			DbSeek(xFilial("SZ6")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DSAID)+SZ5->Z5_HSAID)

			BEGIN TRANSACTION	
				While SZ6->(!EOF()) .and. ;
					 SZ5->Z5_PLACA+DTOS(SZ5->Z5_DSAID)+SZ5->Z5_HSAID == ;
					 SZ6->Z6_PLACA+DTOS(SZ6->Z6_DATA)+SZ6->Z6_HORA
		
					cQuery := "SELECT * "
					cQuery += "FROM  " 
					cQuery += RetSqlName('SC9') + " SC9 "
					cQuery += "WHERE "
					cQuery += "SC9.C9_FILIAL = '" + xFilial("SC9") + "' AND "
					cQuery += "SC9.C9_PEDIDO = '" + SZ6->Z6_PEDIDO + "' AND " 
					cQuery += "SC9.D_E_L_E_T_ <> '*' "					
					cQuery	:= ChangeQuery(cQuery)
					dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC9,.T.,.T.)
					For nSC9 := 1 To len(aStruSC9)
						If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
							TcSetField(cAliasSC9,aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
						EndIf
					Next nSC9
	   				DbSelectArea(cAliasSC9)
	   				While (cAliasSC9)->(!EOF())
	   				    DbSelectArea("SC9")    
	   				    DbSetOrder(1)
	   				    DbSeek(xFilial("SC9")+(cAliasSC9)->C9_PEDIDO+(cAliasSC9)->C9_ITEM+(cAliasSC9)->C9_SEQUEN+(cAliasSC9)->C9_PRODUTO)
						RecLock("SC9",.F.)
						SC9->C9_X_FATUR := "N"
						SC9->(MsUnLock()) 
						DbSelectArea(cAliasSC9)
						(cAliasSC9)->(DbSkip())   				
	   				EndDo
	   				(cAliasSC9)->(DbCloseArea())
		
					RecLock("SZ6",.F.,.T.)
					dbDelete()
					MsUnlock()
						
					SZ6->(DbSkip())
				EndDo
	
				dbSelectArea("SZ5")
				dbGoTo(nReg)
				
				RecLock("SZ5",.F.)
				SZ5->Z5_DSAID := CTOD("  /  /  ")
				SZ5->Z5_HSAID :=  ""
				SZ5->Z5_PSAID := 0
				MsUnlock()
			END TRANSACTION	    		
		Else                                      
        	cMsg += "Favor Excluir as Notas referentes aos pedidos acima antes de estornar a carga " + Chr(13)           										
	    	Aviso("ATENCAO !!!",cMsg,{"<< Voltar"},3,"Ocorrencias: "+Chr(13))
		EndIf
	Else                    
		DbSelectArea("SZ6")    
		DbSetOrder(1)
		DbSeek(xFilial("SZ6")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DENTR)+SZ5->Z5_HENTR)
		BEGIN TRANSACTION	
			While SZ6->(!EOF()) .and. ;
				 SZ5->Z5_PLACA+DTOS(SZ5->Z5_DENTR)+SZ5->Z5_HENTR == ;
				 SZ6->Z6_PLACA+DTOS(SZ6->Z6_DATA)+SZ6->Z6_HORA
				RecLock("SZ6",.F.,.T.)
				dbDelete()
				MsUnlock()
					
				SZ6->(DbSkip())
			EndDo
		
			dbSelectArea("SZ5")
			dbGoTo(nReg)
			RecLock("SZ5",.F.,.T.)
			dbDelete()
			MsUnlock()	   
		END TRANSACTION	    							
	EndIf
EndIf	

RestArea(aArea)

Return(.T.)


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001E   บ Autor ณ Zanardo            บ Data ณ  16/01/04  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Grava as Informacoes								          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
Static Function CAF001E(nRotina,cPlaca,nPeso,dData,cHora,cObs,nEstrado)

Local cQuery 	:= ""
Local cAliasSZ5 := "cAliasSZ5"
Local aStruSZ5 	:= SZ5->(dbStruct())   
Local nSZ5 		:= 0
Local aArea		:= GetArea()
Local nX		:= 0 
Local lRet		:= .T.
Local cAliasSC9 := "cAliasSC9"
Local nSC9 		:= 0
Local aStruSC9 	:= SC9->(dbStruct())

If nPeso > 0  .AND. !Empty(cPlaca)
	If nRotina == 1
		cQuery := "SELECT * "
		cQuery += "FROM "
		cQuery += RetSqlName('SZ5') 
		cQuery += " WHERE "                                  
		cQuery += "Z5_FILIAL = '" + xFilial("SZ5")+"'	AND "
		cQuery += "Z5_PLACA =  '" + cPlaca + "' 		AND "
		cQuery += "Z5_DENTR <= '" + DtoS(dData) + "' 	AND "
		cQuery += "Z5_HENTR <> '" + cHora  + "' 		AND "
		cQuery += "Z5_PSAID = 0 						AND "
		cQuery += "D_E_L_E_T_ <> '*' "
		
		cQuery:=ChangeQuery(cQuery)
		
		dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSZ5,.T.,.T.)
		
		For nSZ5 := 1 To len(aStruSZ5)
			If aStruSZ5[nSZ5][2] <> "C" .And. FieldPos(aStruSZ5[nSZ5][1])<>0
				TcSetField(cAliasSZ5,aStruSZ5[nSZ5][1],aStruSZ5[nSZ5][2],aStruSZ5[nSZ5][3],aStruSZ5[nSZ5][4])
			EndIf
		Next nSZ5
		
		DbSelectArea(cAliasSZ5)
		
		If (cAliasSZ5)->(EOF())
			SZ5->(dbSetOrder(1))
			If !SZ5->(dbSeek(xFilial("SZ5")+cPLACA+DTOS(dData)+cHora))
				RecLock("SZ5",.T.)
				SZ5->Z5_FILIAL := xFilial("SZ5")
				SZ5->Z5_PLACA := cPlaca
				SZ5->Z5_DENTR := dData
				SZ5->Z5_HENTR := cHora
				SZ5->Z5_PENTR := nPeso 
				SZ5->Z5_OBSER := cObs
				SZ5->Z5_MOTIV := cTipos
				lRet := .T.
				SZ5->(MsUnLock())
				If Alltrim(cTipos) <> "1" 
					For nX := 1 To Len(aCols)
							If !aCols[nX,Len(aHeader)+1] .And. !Empty(Alltrim(aCols[nX,3]))
								RecLock("SZ6",.T.)
								SZ6->Z6_FILIAL := xFilial("SZ6")
								SZ6->Z6_PLACA := cPlaca
								SZ6->Z6_DATA  := dData
								SZ6->Z6_HORA  := cHora
								SZ6->Z6_NOTA  := aCols[nX,3]
								SZ6->Z6_SERIE := aCols[nX,4]
								SZ6->Z6_FORN  := aCols[nX,5]
								SZ6->Z6_PESO  := aCols[nX,6]
								SZ6->Z6_TIPON := aCols[nX,1] 
								SZ6->Z6_FORMUL:= aCols[nX,2]
								SZ6->(MsUnLock())
							EndIf	
					Next nX
				EndIf	
		    Else
				MsgAlert("Ja existe uma entrada com a mesma data e hora para este veiculo. <SZ5>")	    
				lRet := .F.
		    EndIf
		Else
			MsgAlert("Ja existe uma entrada anterior deste veํculo sem uma saํda. <SZ5>")
			lRet := .F.
		EndIf
		(cAliasSZ5)->(DbCloseArea())
	ElseIf nRotina == 2
		DbSelectArea("SZ5")
		DbSetOrder(1)
		If DbSeek(xFilial("SZ5")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DENTR)+SZ5->Z5_HENTR)
			RecLock("SZ5",.F.)		
			SZ5->Z5_DSAID := dData
			SZ5->Z5_HSAID := cHora
			SZ5->Z5_PSAID := nPeso
			SZ5->Z5_OBSSD := cObs
			SZ5->Z5_ESTR  := nEstrado
			If lAutoriza 
				SZ5->Z5_DIVERG := nDiverg
				SZ5->Z5_USEAUT := Alltrim(Upper(Substr(cUsuario,7,15))) 
			EndIf			
			SZ5->(MsUnLock())
			If cTipos == "1"	
				For nX := 1 To Len(aCols)  
					If !aCols[nX,Len(aHeader)+1]
						DbSelectArea("SZ6")
						RecLock("SZ6",.T.)
						SZ6->Z6_FILIAL := xFilial("SZ6")
						SZ6->Z6_PLACA := cPlaca
						SZ6->Z6_DATA  := dData
						SZ6->Z6_HORA  := cHora
						SZ6->Z6_TIPON := aCols[nX,01]
						SZ6->Z6_PEDIDO:= aCols[nX,02]
						SZ6->Z6_FORN  := aCols[nX,03]
						SZ6->Z6_PESO  := aCols[nX,04]
						SZ6->(MsUnLock())
						cQuery := "SELECT * "
						cQuery += "FROM  " 
						cQuery += RetSqlName('SC9') + " SC9 "
						cQuery += "WHERE "
						cQuery += "SC9.C9_FILIAL = '" + xFilial("SC9") + "' AND "
						cQuery += "SC9.C9_PEDIDO = '" + aCols[nX,02] + "' AND " 
						cQuery += "SC9.C9_BLEST IN ('  ','10') AND "
						cQuery += "SC9.C9_BLCRED IN ('  ','10') AND "
						cQuery += "SC9.D_E_L_E_T_ <> '*' "					
						cQuery	:= ChangeQuery(cQuery)
						dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC9,.T.,.T.)
						For nSC9 := 1 To len(aStruSC9)
							If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
								TcSetField(cAliasSC9,aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
							EndIf
						Next nSC9
		   				DbSelectArea(cAliasSC9)
		   				While (cAliasSC9)->(!EOF())
		   				    DbSelectArea("SC9")    
		   				    DbSetOrder(1)
		   				    DbSeek(xFilial("SC9")+(cAliasSC9)->C9_PEDIDO+(cAliasSC9)->C9_ITEM+(cAliasSC9)->C9_SEQUEN+(cAliasSC9)->C9_PRODUTO)
							RecLock("SC9",.F.)
							SC9->C9_X_FATUR := "S"
							SC9->(MsUnLock()) 
							DbSelectArea(cAliasSC9)
							(cAliasSC9)->(DbSkip())   				
		   				EndDo
		   				(cAliasSC9)->(DbCloseArea())
		   			EndIf	
				Next nX						
			EndIf
		Else
			MsgAlert("Nao existe entrada para este veiculo. <SZ5>")
			lRet := .F.	
		EndIf
	EndIf	
EndIf

RestArea(aArea)

Return(lRet)


/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001F   บ Autor ณ Zanardo            บ Data ณ  28/02/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Impressao do cupom de entrada                              บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function CAF001F(cAlias,nReg,nOpc,nRotina,cPlaca,nPeso,dData,cHora,cObs)

Local aOrd           := {}
Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Autorizacao de Entrada de Veiculo"
Local cPict          := ""
Local titulo         := "Controle de Veiculo - COMAFAL/"+SM0->M0_ESTENT
Local Cabec1         := ""
Local Cabec2         := ""
Local Imprime        := .T.
Local nVeiImp		 := 0
Local cString		 := "SZ5" 
Local NomeProg		 := "CAF001" 
Local lAbortPrint 	 := .F.
Local limite      	 := 80
Local tamanho     	 := "P"      

Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey    := 0
Private cbcont     	:= 0
Private m_pag      	:= 01
Private wnrel      	:= "CAF001" // Coloque aqui o nome do arquivo usado para impressao em disco
//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Monta a interface padrao com o usuario...                           ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
If nOpc == 4
	DbSelectArea(cAlias)
	DbGoto(nReg)
	If SZ5->Z5_MOTIV == "1"
		cTipos := "Carregamento"
	ElseIf	SZ5->Z5_MOTIV == "2" 
		cTipos := "Descarregamento"
	Else
		cTipos := "Outros"
	EndIf	
Else
	If cTipos == "1"
		cTipos := "Carregamento"
	ElseIf	cTipos == "2" 
		cTipos := "Descarregamento"
	Else
		cTipos := "Outros"
	EndIf		
EndIf
	
wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.T.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

nTipo := If(aReturn[4]==1,15,18)

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Processamento. RPTSTATUS monta janela com a regua de processamento. ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

RptStatus({|| RunReport(@Cabec1,@Cabec2,Titulo,nRotina,cPlaca,nPeso,dData,cHora,cObs,tamanho,nTipo,NomeProg) },Titulo)
Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณRUNREPORT บ Autor ณ AP5 IDE            บ Data ณ  20/05/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS บฑฑ
ฑฑบ          ณ monta a janela com a regua de processamento.               บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ Programa principal                                         บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
Static Function RunReport(Cabec1,Cabec2,Titulo,nRotina,cPlaca,nPeso,;
							dData,cHora,cObs,Tamanho,nTipo,NomeProg)
 
Local nVeiImp := 0
Local nLin    := 0
Local nItens  := 1
If  EMPTY(SZ5->Z5_DSAID)//Cupon de Entrada
	CAF001C(1)
	@ nLin++, 00 Psay "COMAFAL/"+SM0->M0_ESTENT+" - AUTORIZACAO DE ENTRADA DE VEICULO PLACA ==> " + SZ5->Z5_PLACA
	@ nLin++, 00 Psay "MOTIVO DA ENTRADA ==> " +  cTipos
	/*	      //>          1         2         3         4         5         6         7         8
		      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890
	aL[01]	:=	"|======================================================|"
	aL[02]	:=	"| DATA DA ENTRADA | HORA DA ENTRADA  | PESO DA ENTRADA |"
	aL[03]	:=	"|------------------------------------------------------|"
	aL[04]	:=	"|   ##########    |      #####       |    ########     |"		
	aL[05]	:=	"|======================================================|"
	aL[06]	:=	"|NOTA    |SERIE| COD. FORNEC. |         PESO           |"
	aL[07]	:=	"|------------------------------------------------------|"
	aL[08]	:=	"|######  | ### |    ######    |       ########         |"
	aL[09]	:=	"|======================================================|"	*/

	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({SZ5->Z5_DENTR,SZ5->Z5_HENTR,Transform(SZ5->Z5_PENTR,"@E 99,999.999")},aL[4],"",,@nLin)                              	
	DbSelectArea("SZ6")    
	DbSetOrder(1)	
	If DbSeek(xFilial("SZ6")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DENTR)+SZ5->Z5_HENTR)
		FmtLin(,{aL[5],aL[6],aL[7]},,,@nLin)
		While SZ6->(!Eof() .And. SZ6->Z6_PLACA = SZ5->Z5_PLACA .and. SZ6->Z6_DATA = SZ5->Z5_DENTR .and. SZ6->Z6_HORA = SZ5->Z5_HENTR)
			FmtLin({SZ6->Z6_NOTA,SZ6->Z6_SERIE,SZ6->Z6_FORN,Transform(SZ6->Z6_PESO,"@E 99,999.999")},aL[8],"",,@nLin)                              	
			SZ6->(dbSkip())
		Enddo
		FmtLin(,{aL[9]},,,@nLin)
	Else
		@ nLin++,03 Psay "Nao existem notas para este veiculo"
	Endif
	/*	      //>          1         2         3         4         5         6         7         8
		      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890
    */
	@ nLin++, 00 Psay "OBSERVACAO"
	@ nLin++, 00 Psay cObs
	nLin++
	@ nLin++, 00 Psay "_____________________________"
	@ nLin++, 00 Psay "          Faturista          "
	
Elseif !EMPTY(SZ5->Z5_DSAID) //Cupon de Saida
	CAF001C(2)	
	@ nLin++, 00 Psay "COMAFAL/"+SM0->M0_ESTENT+" - AUTORIZACAO DE SAIDA DE VEICULO PLACA ==> " + SZ5->Z5_PLACA
	@ nLin++, 00 Psay "MOTIVO DA SAIDA ==> " +  cTipos
	/*
		      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
		      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
	aL[01]	:=	"|===================================================================================================|"
	aL[02]	:=	"|DT.ENTRADA |HR.ENTRADA |PESO ENTRADA |DT.SAIDA   |HR.SAIDA |PESO SAIDA  |PESO LIQ.     |P.EXTRA(KG)|"
	aL[03]	:=	"|---------------------------------------------------------------------------------------------------|"	
	aL[04]	:=	"|########## |#####      |############ |########## |#####    |########### |############# |###########|"			
	aL[05]	:=	"|===================================================================================================|"	
	aL[06]	:=	"| PEDIDO | COD. CLIENTE | PESO NOTA FISCAL                                                          |"
	aL[07]	:=	"|---------------------------------------------------------------------------------------------------|"	
	aL[08]	:=	"| ###### |    ######    | ###########                                                               |"	
	aL[09]	:=	"|===================================================================================================|"	

	*/
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({SZ5->Z5_DENTR,SZ5->Z5_HENTR,Transform(SZ5->Z5_PENTR,"@E 99,999.999"),;
			SZ5->Z5_DSAID,SZ5->Z5_HSAID,Transform(SZ5->Z5_PSAID,"@E 99,999.999"),;
			Transform((SZ5->Z5_PSAID-SZ5->Z5_PENTR),"@E 99,999.999"),;
			Transform(SZ5->Z5_ESTR,"@E 99,999.999")},aL[4],"",,@nLin)                              
	FmtLin(,{aL[05]},,,@nLin)					
	DbSelectArea("SZ6")    
	DbSetOrder(1)	
	If DbSeek(xFilial("SZ6")+ SZ5->Z5_PLACA+DTOS(SZ5->Z5_DSAID)+SZ5->Z5_HSAID)
		FmtLin(,{aL[06]},,,@nLin)			
		While SZ6->(!Eof() .And. SZ6->Z6_PLACA == SZ5->Z5_PLACA .and. SZ6->Z6_DATA == SZ5->Z5_DSAID .and. SZ6->Z6_HORA = SZ5->Z5_HSAID)
			If nItens > 15  
				FmtLin(,{aL[05]},,,@nLin)
				nLin    := 0
				FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
				FmtLin({SZ5->Z5_DENTR,SZ5->Z5_HENTR,Transform(SZ5->Z5_PENTR,"@E 99,999.999"),;
						SZ5->Z5_DSAID,SZ5->Z5_HSAID,Transform(SZ5->Z5_PSAID,"@E 99,999.999"),;
						Transform((SZ5->Z5_PSAID-SZ5->Z5_PENTR),"@E 99,999.999"),;
						Transform(SZ5->Z5_ESTR,"@E 99,999.999")},aL[4],"",,@nLin)                              
				FmtLin(,{aL[05]},,,@nLin)					
				nItens := 1
			EndIf
			FmtLin({SZ6->Z6_PEDIDO,SZ6->Z6_FORN,Transform(SZ6->Z6_PESO,"@E 99,999.999")},aL[08],"",,@nLin)			
			SZ6->(dbSkip())
			nItens++
		Enddo
		FmtLin(,{aL[09]},,,@nLin)				
	Else
		@ nLin,03 Psay "Nao existe Pedido este veiculo"
	Endif
	nLin++
	@nLin, 00 Psay "OBSERVACAO ENTRADA"
	@nLin, 20 Psay Substr(Alltrim(SZ5->Z5_OBSER),1,50)
	nLin++
	@nLin, 00 Psay "OBSERVACAO SAIDA"
	@nLin, 55 Psay "_____________________________"
	nLin++
	@nLin, 00 Psay Substr(Alltrim(SZ5->Z5_OBSSD),1,50)
	@nLin, 55 Psay "          Faturista          "
EndIf

SetPrc(0,0)

Set Device To Screen

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Se impressao em disco, chama o gerenciador de impressao...          ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

If aReturn[5]==1
	dbCommitAll()
	Set Printer To
	OurSpool(wnrel)
Endif

MS_FLUSH()

Return

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001G   บ Autor ณ Zanardo            บ Data ณ  16/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ VLDUSER - Z6_PEDIDO									      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/
User Function CAF001G()

Local cPedido 	:= &(ReadVar()) 
Local aArea   	:= GetArea()
Local lRet    	:= .T.      
Local cAliasSC9 := "cAliasSC9"
Local nSC9 		:= 0
Local aStruSC9 	:= SC9->(dbStruct())

DbSelectArea("SZ6")
DbSetOrder(2)
If MsSeek(xFilial("SZ6")+cPedido)
	lRet:= .F.
	MsgAlert("Pedido ja foi utilizado na motagem da carga do veiculo placa :"+SZ6->Z6_PLACA+". <SZ6>")	            
Else
	DbSelectArea("SC5")
	DbSetOrder(1)
	If MsSeek(xFilial("SC5")+cPedido)
		If  Empty(SC5->C5_NOTA )
			DbSelectArea("SC9")
			DbSetOrder(1)  
			If MsSeek(xFilial("SC9")+cPedido)
				If SC9->C9_BLEST $"  #10" .And. SC9->C9_BLCRED $"  #10"
					cQuery := "SELECT SUM(SC9.C9_QTDLIB) AS QUANTVD "
					cQuery += "FROM  " 
					cQuery += RetSqlName('SC9') + " SC9 "
					cQuery += "WHERE "
					cQuery += "SC9.C9_FILIAL = '" + xFilial("SC9") + "' AND "
					cQuery += "SC9.C9_PEDIDO = '" + cPedido + "' AND " 
					cQuery += "SC9.C9_BLEST IN ('  ','10') AND "
					cQuery += "SC9.C9_BLCRED IN ('  ','10') AND "
					cQuery += "SC9.D_E_L_E_T_ <> '*' "					
					cQuery	:= ChangeQuery(cQuery)

					dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC9,.T.,.T.)

					For nSC9 := 1 To len(aStruSC9)
						If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
							TcSetField(cAliasSC9,aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
						EndIf
					Next nSC9
     				DbSelectArea(cAliasSC9)
     				If (cAliasSC9)->QUANTVD = SC5->C5_PESOL
						aCols[n,3]:= SC5->C5_CLIENTE
						aCols[n,4]:= SC5->C5_PESOL
					Else
						lRet:= .F.
						MsgAlert("A soma da quantidade de liberada dos itens nao esta igual ao peso total do pedido. <SC9>")	            					
					EndIf
						(cAliasSC9)->(DbCloseArea())
		        Else
					lRet:= .F.
					MsgAlert("O Pedido possue bloqueio. <SC9>")	            
				Endif
			Else
				lRet:= .F.
				MsgAlert("O Pedido nao possue itens liberados. <SC9>")	    
			EndIf
		Else
			lRet:= .F.
			MsgAlert("O Pedido ja foi faturado. <SC5>")	    
		EndIf	
	Else 
		lRet:= .F.
		MsgAlert("O Pedido nใo foi encontrado. <SC5>")	    
	Endif 
EndIf

RestArea(aArea)
Return(lRet)

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001LinOkบ Autor ณ Zanardo           บ Data ณ  27/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Valida a linha digitada no aCols  					      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/
User Function CAF001LinOk(nLin)   
Local lRet := .T. 
Local aArea := GetArea()
Local nX := 0

If !aCols[nLin,Len(aHeader)+1]
	If cTipos == "2" .And. nOpcMen == 2
		For nX := 1 to 5
			If Empty(Alltrim(aCols[nLin,nX])) .or. aCols[nLin,6] <= 0
				MsgAlert("Existe Campo sem informa็ใo")
			    lRet := .F.
			    nX := 5
			EndIf		    
		Next nX
		If lRet .And. (aCols[nLin,1]$"1#3")
			If aCols[nLin,1]== "1"
				DbSelectArea("SA2")
				DbSetOrder(1)
				If !MsSeek(xFilial("SA2")+aCols[nLin,5])
					MsgAlert("O Fornecedor informado nao estแ cadastrado.<SA2>")
			    	lRet := .F.			
				EndIf
			Else
				DbSelectArea("SA1")
				DbSetOrder(1)
				If !MsSeek(xFilial("SA1")+aCols[nLin,5])
					MsgAlert("O Cliente informado nao estแ cadastrado.<SA1>")
			    	lRet := .F.			
				ElseIf aCols[nLin,2]=="S" 
					DbSelectArea("SF2")
					DbSetOrder(1)
					If !MsSeek(xFilial("SF2")+aCols[nLin,3]+aCols[nLin,4])
						MsgAlert("NF de origem nao encontrado.<SF2>")
				    	lRet := .F.			
					EndIf			
				EndIf
			EndIf
		EndIf
	ElseIf cTipos == "1" .And. nOpcMen == 3
    	DbSelectArea("SC5")
    	DbSetOrder(1)
    	If MsSeek(xFilial("SC5")+aCols[nLin,2])
    		If SC5->C5_PESOL <>  aCols[nLin,4]
				MsgAlert("O peso informado estแ diferente do Pedido.<SC5>")
		   		lRet := .F.			    		
		   		aCols[nLin,4] := SC5->C5_PESOL
    		Endif
    	Else
    		MsgAlert("Pedido nao encontrado.<SC5>")
		   	lRet := .F.			
		EndIf
	EndIf
EndIf
RestArea(aArea)
	
Return(lRet)


/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณCAF001TdOk บ Autor ณ Zanardo           บ Data ณ  04/06/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Valida os itens digitados no aCols					      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

User Function CAF001TdOk(nOpc,nEstrado)

Local lRet 	:= .T.
Local nX 	:= 0
Local nCont := 0
Local aArea := GetArea()
Local nPesNFE := 0   
Local nPesPED := 0  
Local cMsg 	:= ""  

If cTipos == "2" 
	If nOpc == 2
		For nX:= 1 to Len(aCols)
			If !aCols[nX,Len(aHeader)+1] .and. U_CAF001LinOk(nX)
				nCont++	
			EndIf
		Next nX
		If  nCont == 0 
			MsgAlert("Notas Fiscais de Entrada nao informadas.")
			lRet := .F.
		Endif	
	ElseIf nOpc == 3
		If SZ6->(dbSeek(xFilial("SZ6")+SZ5->Z5_PLACA+DTOS(SZ5->Z5_DENTR)+SZ5->Z5_HENTR))

			While SZ6->(!Eof() .And. SZ6->Z6_FILIAL == xFilial("SZ6") .And.;
					 SZ6->Z6_PLACA == SZ5->Z5_PLACA .And.;
					 SZ6->Z6_DATA == SZ5->Z5_DENTR .And.;
					 SZ6->Z6_HORA == SZ5->Z5_HENTR)
					 
				nPesNFE+=SZ6->Z6_PESO

				SZ6->(dbSkip())

			Enddo
            If SZ5->Z5_PENTR < ((nPeso + nPesNFE)*0.99)      
	
	           	cMsg := "O pesagem da Entrada esta divergente em relacao a Pesagem da Saida mais as Notas Fiscais de Entrada." + Chr(13)
	       		cMsg += "Pesagem na Entrada 				(A): " + Str(SZ5->Z5_PENTR) + Chr(13)
				cMsg += "Pesagem na Saida   				(B): " + Str(nPeso) + Chr(13)           		
				cMsg += "Total das NFE      				(C): " + Str(nPesNFE) + Chr(13)
				cMsg += "Pesagem na Saida + Total das NFE (B+C): " + Str(nPeso + nPesNFE) + Chr(13)				           						
				cMsg += "Diferenca  			      (A-(B+C)): " + Str(SZ5->Z5_PENTR-(nPeso + nPesNFE)) + Chr(13)           										
	       		Aviso("ATENCAO !!!",cMsg,{"<< Voltar"},3,"Ocorrencias: "+Chr(13))

            	If !CAF001I()
	            	lRet := .F.
	            Else
	            	lAutoriza := .T.
	            	nDiverg	  := SZ5->Z5_PENTR - (nPeso + nPesNFE)
	            EndIf
	            	
            EndIf
		Else
			MsgAlert("Notas Fiscais de Entrada encontradas.<SZ6>")
			lRet := .F.
		Endif
	EndIf
ElseIf cTipos == "1"
	If nOpc == 3
		For nX:= 1 to Len(aCols)
			If !aCols[nX,Len(aHeader)+1]
				nPesPED += aCols[nX,4]
			EndIf
		Next nX 
		If (nEstrado/1000) <= (nPesPED * 0.01)
			If (nPeso - (nEstrado/1000)) > (nPesPED + SZ5->Z5_PENTR)
	           	cMsg := "O pesagem da Saida esta divergente em relacao a Pesagem da Entrada mais os Pedidos de Venda." + Chr(13)
				cMsg += "Pesagem na Saida   				(B): " + Str(nPeso) + Chr(13)           		
	       		cMsg += "Pesagem na Entrada 				(A): " + Str(SZ5->Z5_PENTR) + Chr(13)
				cMsg += "Total dos Pedidos     				(C): " + Str(nPesPED) + Chr(13)
				cMsg += "Pesagem na Entrada + Pedidos	  (A+C): " + Str(SZ5->Z5_PENTR + nPesPED) + Chr(13)				           						
				cMsg += "Diferenca  			      (B-(A+C)): " + Str(nPeso-(SZ5->Z5_PENTR + nPesPED)) + Chr(13)           										
	       		Aviso("ATENCAO !!!",cMsg,{"<< Voltar"},3,"Ocorrencias: "+Chr(13))
	           	lRet := .F.
			EndIf
		Else
       		Aviso("ATENCAO !!!","O Peso do Estrado esta a maior que 1% referente ao peso total dos pedidos",{"<< Voltar"},3,"Ocorrencias: "+Chr(13))
           	lRet := .F.
		EndIf	
	EndIf
EndIf

RestArea(aArea)
	
Return(lRet)

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFuno    ณBalanca   บ Autor ณ Zanardo            บ Data ณ  04/06/02   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescrio ณ Ativa balanca e busca informacoes no arquivo gerado        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/

Static Function Balanca()

Local cStrPeso 	:= ""
Local aOpcoes   := {OemtoAnsi("Balanca Nova"), OemtoAnsi("Balanca Antiga")}
Local nRadio    := 1
Local oDlgBal
Local oRadio
Local cBalPrg 	:= "" 
Local cBalDir 	:= ""

nPeso := 0
//If xFilial("SF2") $ "02#03#06"
	DEFINE MSDIALOG oDlgBal FROM 000,000 TO 70,100 TITLE "Selecione a Balanca" PIXEL
		
	@ 003,003 Radio oRadio Var nRadio Items OemtoAnsi("Balanca Nova"), OemtoAnsi("Balanca Antiga") 3D Size 50,10 Of oDlgBal PIXEL
	
	DEFINE SBUTTON FROM 23,10 TYPE 1 ACTION (oDlgBal:End()) ENABLE OF oDlgBal
	Activate Dialog oDlgBal Centered
	
	If nRadio == 1
		cBalPrg := "C:\PCLINK\PCLINK.EXE"        
		cBalDir := "C:\PCLINK\PESO.TXT" 
	Else
		cBalPrg := "C:\BALANCA\BALANCA.EXE"
		cBalDir := "C:\BALANCA\PESO.TXT" 
	EndIf
//Else
//	cBalPrg := "C:\BALANCA\BALANCA.EXE"
//	cBalDir := "C:\BALANCA\PESO.TXT" 
//EndIf	
	
WINEXEC(cBalPrg)
MsgInfo("Pressione para obter peso!")


StrPeso := MEMOREAD(cBalDir)
StrPeso := StrTran(StrPeso,Chr(10),"")
StrPeso := StrTran(StrPeso,Chr(13),"")
StrPeso := StrTran(StrPeso,Chr(26),"")
nPeso	:= Val(StrTran(AllTrim(StrPeso),",","."))
 
FErase(cBalDir)                       

ObjectMethod(oDlg,"Refresh()")

Return 


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณAJUSTASX1 บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณAjusta o SX1                                                บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ AFATR001                                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function AjustaSx1()

Local aArea := GetArea()

PutSx1("CAF001","01","Tipo do Cupom ?","","","mv_ch1","C",1,0,0,"C","","","","",;
				"MV_PAR01","Entrada","","","","Saida","","","","","","","","","","","")  

PutSx1("CAF001","02","Placa ?","","","mv_ch2","C",7,0,0,"G","","","","",;
				"MV_PAR02","","","","","","","","","","","","","","","","")

PutSx1("CAF001","03","Data ?","","","mv_ch3","D",8,0,0,"G","","","","",;
				"MV_PAR03","","","","","","","","","","","","","","","","")
				
PutSx1("CAF001","04","Hora ?","","","mv_ch4","C",5,0,0,"G","","","","",;
				"MV_PAR04","","","","","","","","","","","","","","","","")
				
RestArea(aArea)

Return(.T.)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณM460QRY   บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณPonto de Entrada para filtrar o SC9 MATA460A()              บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ MATA460A                                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function M460QRY()
Local aArea := GetArea()
Local cQuery := ParamIXB[1]

If xFilial("SF2") <> "04"
	cQuery += " AND SC9.C9_X_FATUR = 'S' "	
EndIf	

RestArea(aArea)

Return(cQuery)      

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณM460FIL   บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณPonto de Entrada para filtrar o SC9 MATA460A()              บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ MATA460A                                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function M460FIL()
Local aArea := GetArea()
Local cQuery := ""

If xFilial("SF2") <> "04"
	cQuery := " C9_X_FATUR == 'S' "
Else                             
	cQuery := "( C9_X_FATUR == 'S' .OR. C9_X_FATUR <> 'S') "	
EndIf	

RestArea(aArea)

Return(cQuery) 

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณCAF001C   บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณLayOut do Relatorio								          บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ MATA460A                                                   บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

Static Function CAF001C(nLayout)
Local aArea := GetArea()

Do Case
	Case nLayOut == 1
		aL :=	Array(09)
	 		      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
			      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
		aL[01]	:=	"|======================================================|"
		aL[02]	:=	"| DATA DA ENTRADA | HORA DA ENTRADA  | PESO DA ENTRADA |"
		aL[03]	:=	"|------------------------------------------------------|"
		aL[04]	:=	"|   ##########    |      #####       |   ###########   |"		
		aL[05]	:=	"|======================================================|"
		aL[06]	:=	"|NOTA    |SERIE| COD. FORNEC. |         PESO           |"
		aL[07]	:=	"|------------------------------------------------------|"
		aL[08]	:=	"|######  | ### |    ######    |      ###########       |"
		aL[09]	:=	"|======================================================|"	
		
	Case nLayout == 2
		aL :=	Array(09)
	 		      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
			      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890

		aL[01]	:=	"|================================================================================|"
		aL[02]	:=	"|DT.ENTR |HR.ET|PESO ENTRADA|DT.SAIDA|HR.SD|PESO SAIDA |PESO LIQ.    |P.EXTRA(KG)|"
		aL[03]	:=	"|--------------------------------------------------------------------------------|"	
		aL[04]	:=	"|########|#####|############|########|#####|###########|#############|###########|"			
		aL[05]	:=	"|================================================================================|"	
		aL[06]	:=	"| PEDIDO | COD. CLIENTE | PESO NOTA FISCAL                                       |"
		aL[07]	:=	"|--------------------------------------------------------------------------------|"	
		aL[08]	:=	"| ###### |    ######    | ###########                                            |"	
		aL[09]	:=	"|================================================================================|"	
		
	Case nLayout == 3
		
		aL :=	Array(05)
	 		      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
			      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
		aL[01]	:=	"|========================================================================================|"
		aL[02]	:=	"| Placa  |DT.Entrada|Hr.Entrada |DT.Saida  |Hr.Saida|Diveg.(T)| Estrado |    Usuario     |"
		aL[03]	:=	"|----------------------------------------------------------------------------------------|"	
		aL[04]	:=	"|####### |##########|  #####    |##########| #####  | ########| ########| ###############|"			
		aL[05]	:=	"|========================================================================================|"
		
EndCase
	
RestArea(aArea)

Return(.T.)


/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณCAF001I   บAutor  ณEduardo Zanardo     บ Data ณ  19/03/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณAutorizacao para divergencia de peso  			          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

Static Function CAF001I()

Local aArea := GetArea()
Local lRet  := .F.      
Local aDepto 	:= PswRet(1)
Local aUser 	:= U_GetModulo()
Local nPosAtivo := aScan(aUser, {|x| x[3]==.T.})
If (UPPER(Alltrim(aDepto[1,12])) == "FATURAMENTO" .and. aUser[nPosAtivo,2] == "2" )  .or. aUser[nPosAtivo,2]>= "4" .or.;
	 (aUser[nPosAtivo,2] == "3" .and. UPPER(Alltrim(aDepto[1,12])) == "FINANCEIRO")
	
	If APMSGNOYES("Autoriza divergencia de Peso?")
 		lRet := .T.                             
	EndIf 	
	
EndIf

RestArea(aArea)

Return(lRet)           

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณCAF001H   บAutor  ณEduardo Zanardo     บ Data ณ  29/02/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณRelatorio de Autorizacao para divergencia de peso	          บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
User Function CAF001H(cAlias,nReg,nOpc)
Local aOrd           := {}
Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Relatorio de Autorizacao para divergencia de peso"
Local cPict          := ""
Local titulo         := "Relatorio de Autorizacao para divergencia de peso - COMAFAL/"+SM0->M0_ESTENT
Local Cabec1         := ""
Local Cabec2         := ""
Local Imprime        := .T.
Local nVeiImp		 := 0
Local cString		 := "SZ5" 
Local NomeProg		 := "CAF001" 
Local lAbortPrint 	 := .F.
Local limite      	 := 80
Local tamanho     	 := "P"      

Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey    := 0
Private cbcont     	:= 0
Private m_pag      	:= 01
Private wnrel      	:= "CAF001" // Coloque aqui o nome do arquivo usado para impressao em disco
	
wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.T.,Tamanho,,.T.)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

nTipo := If(aReturn[4]==1,15,18)

RptStatus({|| CAF001J(@Cabec1,@Cabec2,Titulo,tamanho,nTipo,NomeProg) },Titulo)

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณCAF001J   บAutor  ณEduardo Zanardo     บ Data ณ  29/02/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณImp.Relatorio de Autorizacao para divergencia de peso	      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function CAF001J(Cabec1,Cabec2,Titulo,tamanho,nTipo,NomeProg)

Local cQuery 	:= ""
Local cAliasSZ5 := "cAliasSZ5"
Local aStruSZ5 	:= SZ5->(dbStruct())   
Local nSZ5 		:= 0 
Local nLin 		:= 80
cQuery := "SELECT * "
cQuery += "FROM "
cQuery += RetSqlName('SZ5') 
cQuery += " WHERE "                                  
cQuery += "Z5_FILIAL = '" + xFilial("SZ5")+"'	AND "
cQuery += "Z5_DENTR = '"  + DtoS(dDataBase) + "' 	AND "
cQuery += "(Z5_DIVERG > 0 OR Z5_ESTR > 0) AND "
cQuery += "D_E_L_E_T_ <> '*' "
		
cQuery:=ChangeQuery(cQuery)
		
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSZ5,.T.,.T.)
		
For nSZ5 := 1 To len(aStruSZ5)
	If aStruSZ5[nSZ5][2] <> "C" .And. FieldPos(aStruSZ5[nSZ5][1])<>0
		TcSetField(cAliasSZ5,aStruSZ5[nSZ5][1],aStruSZ5[nSZ5][2],aStruSZ5[nSZ5][3],aStruSZ5[nSZ5][4])
	EndIf
Next nSZ5

CAF001C(3)
		
DbSelectArea(cAliasSZ5)

Do While (cAliasSZ5)->(EOF())

/*	 		      //>          1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20
			      //>012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
		aL[01]	:=	"|========================================================================================|"
		aL[02]	:=	"| Placa  |DT.Entrada|Hr.Entrada |DT.Saida  |Hr.Saida|Diveg.(T)| Estrado |    Usuario     |"
		aL[03]	:=	"|----------------------------------------------------------------------------------------|"	
		aL[04]	:=	"|####### |##########|  #####    |##########| #####  | ########| ########| ###############|"			
		aL[05]	:=	"|========================================================================================|"
*/

	If nLin > 55
		Cabec1 := "Periodo: " + cMonth(dDatabase)+"/"+Str(Year(dDatabase))
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1
		FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)                                        
	EndIf	

	FmtLin({(cAliasSZ5)->Z5_PLACA,(cAliasSZ5)->Z5_DENTR,;
		(cAliasSZ5)->Z5_HENTR,(cAliasSZ5)->Z5_DSAID,;
		(cAliasSZ5)->Z5_HSAID,;
		Transform((cAliasSZ5)->Z5_DIVERG,"@E 999.999"),;
		Transform((cAliasSZ5)->Z5_ESTR,"@E 999.999"),;
		(cAliasSZ5)->Z5_USEAUT},aL[4],"",,@nLin)                              	
		
	(cAliasSZ5)->(DbSkip())          
EndDo

FmtLin(,{aL[5]},,,@nLin)

(cAliasSZ5)->(DbCloseArea())          

SetPrc(0,0)

Set Device To Screen

//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
//ณ Se impressao em disco, chama o gerenciador de impressao...          ณ
//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู

If aReturn[5]==1
	dbCommitAll()
	Set Printer To
	OurSpool(wnrel)
Endif

MS_FLUSH()

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณMTA440C9   บAutor  ณEduardo Zanardo     บ Data ณ  05/03/04  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณLibera para o Faturamento sem passar pela balan็a           บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
User Function MTA440C9()

Local aArea 	:= GetArea()

If SC5->C5_TIPO $"CPI" .OR.;
	 SC6->C6_CF$"5915#6915#5924#6924#5949#6949#5552#6552#5551#6551" .OR.;
	 Posicione("SF4",1,xFilial("SF4")+SC6->C6_TES,"F4_ESTOQUE") == "N" .OR.;
	 ((SC5->C5_CLIENTE $ "004502#006629" .AND. (xFilial("SF2")=="01" .OR. xFilial("SF2") == "04")))
	 
	DbSelectArea("SC9")
	DbSetOrder(1)
	If DbSeek(xFilial("SC9")+SC9->C9_PEDIDO+SC9->C9_ITEM+SC9->C9_SEQUEN+SC9->C9_PRODUTO)
		RecLock("SC9",.F.)
		C9_X_FATUR := "S"
		MsUnlock()
	EndIf
EndIf

RestArea(aArea)

Return(.T.)           

