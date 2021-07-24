#INCLUDE "PROTHEUS.CH"
#INCLUDE "topconn.ch"

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CRF010  ºAutor  ³Eduardo Zanardo     º Data ³  03/31/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ RESUMO GERENCIAL                   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP7                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/


User Function CRF010E(nLin,nMes,Cabec1,Cabec2,Titulo,aQtdVend)

Local aArea := GetArea()
Local aPRODUZIDO := {}
Local aProdmAQ := {}                       
Local aParadaMaq := {} 
Local aPerdaMaq := {}
Local aProdutividade := {} 
Local aMateriaP := {}
Local aMateriaA := {}
Local cQuery := ""
Local nI := 0
Local cAliasSD3 := "cAliasSD3"
Local nSD3 := 0
Local aStruSD3	:= SD3->(dbStruct())
Local cAliasSH1 := "cAliasSH1"
Local nSH1 := 0
Local aStruSH1	:= SH1->(dbStruct())
Local cAliasSH6 := "cAliasSH6"
Local nSH6 := 0
Local aStruSH6	:= SH6->(dbStruct())
Local cAliasSBC := "cAliasSBC"
Local nSBC := 0
Local aStruSBC	:= SBC->(dbStruct())
Local cAliasSB2 := "cAliasSB2"
Local nSB2 := 0
Local aStruSB2	:= SB2->(dbStruct())
Local cAliasSRD := "cAliasSRD"
Local nSRD := 0
Local aStruSRD	:= SRD->(dbStruct())
Local cAliasSE7 := "cAliasSE7"
Local nSE7 := 0
Local aStruSE7	:= SE7->(dbStruct())
Local cAliasSB9 := "cAliasSB9"
Local nSB9 := 0
Local aStruSB9	:= SB9->(dbStruct())
Local nReg := 0 
Local nTotal1 := 0
Local nTotal2 := 0
Local nTotal3 := 0

U_RFinLay3()
FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
FmtLin(,{aL[5],aL[7]},,,@nLin)
FmtLin({Transform(aQtdVend[1,2],"@R 999,999.999"),;
		Transform(aQtdVend[2,2],"@R 999,999.999"),;
		Transform(aQtdVend[3,2],"@R 999,999.999")},aL[08],"",,@nLin)
FmtLin(,{aL[9],aL[11]},,,@nLin)
cQuery := " SELECT SUBSTRING(D3_EMISSAO,1,6) AS EMISSAO, SUM(D3_QUANT) AS PRODUZIDO FROM "
cQuery += RetSqlName('SD3')
cQuery += " WHERE D3_FILIAL = '"+xFilial("SD3")+"'  "
cQuery += " AND SUBSTRING(D3_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(D3_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND D3_CF = 'PR0' AND D3_TM = '500' AND D3_LOCAL IN ('03','04') AND D_E_L_E_T_ = ' '  "         
cQuery += " GROUP BY SUBSTRING(D3_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D3_EMISSAO,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD3,.T.,.T.)},"Aguarde...","Qtd. Cts. Produzidos...")
aadd(aPRODUZIDO,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 
aadd(aPRODUZIDO,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aPRODUZIDO,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
For nSD3 := 1 To len(aStruSD3)
	If aStruSD3[nSD3][2] <> "C" .And. FieldPos(aStruSD3[nSD3][1])<>0
		TcSetField(cAliasSD3,aStruSD3[nSD3][1],aStruSD3[nSD3][2],aStruSD3[nSD3][3],aStruSD3[nSD3][4])
	EndIf
Next nSD3
Do While (cAliasSD3)->(!EOF())
		If (cAliasSD3)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aPRODUZIDO[1,2] := (cAliasSD3)->PRODUZIDO
      	ElseIf (cAliasSD3)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aPRODUZIDO[2,2] := (cAliasSD3)->PRODUZIDO
       	ElseIf (cAliasSD3)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
       		aPRODUZIDO[3,2] := (cAliasSD3)->PRODUZIDO
       	EndIf	
	(cAliasSD3)->(DbSkip())
EndDo
DbSelectArea(cAliasSD3)
dbCloseArea()
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
FmtLin({Transform(aPRODUZIDO[1,2],"@R 9,999,999.999"),;
	    Transform(aPRODUZIDO[2,2],"@R 9,999,999.999"),;
		Transform(aPRODUZIDO[3,2],"@R 9,999,999.999")},aL[10],"",,@nLin)
FmtLin(,{aL[13],aL[12]},,,@nLin) 
cQuery := " SELECT SUBSTRING(H6.H6_DTAPONT,1,6) AS EMISSAO,H1.H1_DESCRI AS MAQUINA, SUM(H6.H6_QTDPROD) AS PRODUZIDO "
cQuery += " FROM "
cQuery += RetSqlName('SH6') + " H6, "   
cQuery += RetSqlName('SH1') + " H1 " 
cQuery += " WHERE "
cQuery += " H6.H6_FILIAL = '"+xFilial("SH6")+"'  "
cQuery += " AND SUBSTRING(H6.H6_DTAPONT,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(H6.H6_DTAPONT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND H6.D_E_L_E_T_ = ' '  "  
cQuery += " AND H6.H6_QTDPROD > 0 "
cQuery += " AND H6.H6_RECURSO = H1.H1_CODIGO " 
cQuery += " AND H1.H1_FILIAL = '"+xFilial("SH1")+"'  " 
cQuery += " AND H1.D_E_L_E_T_ = ' '  "  
cQuery += " GROUP BY SUBSTRING(H6.H6_DTAPONT,1,6),H1.H1_DESCRI "
cQuery += " ORDER BY H1.H1_DESCRI ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD3,.T.,.T.)},"Aguarde...","Qtd. Cts. Prod. por Maquina...")
For nSH6 := 1 To len(aStruSH6)
	If aStruSH6[nSH6][2] <> "C" .And. FieldPos(aStruSH6[nSH6][1])<>0
		TcSetField(cAliasSD3,aStruSH6[nSH6][1],aStruSH6[nSH6][2],aStruSH6[nSH6][3],aStruSH6[nSH6][4])
	EndIf
Next nSH6 
For nSH1 := 1 To len(aStruSH1)
	If aStruSH1[nSH1][2] <> "C" .And. FieldPos(aStruSH1[nSH1][1])<>0
		TcSetField(cAliasSD3,aStruSH1[nSH1][1],aStruSH1[nSH1][2],aStruSH1[nSH1][3],aStruSH1[nSH1][4])
	EndIf
Next nSH1
Do While (cAliasSD3)->(!EOF())
	nReg := aScan(aProdmAQ,{|x| x[2]==(cAliasSD3)->MAQUINA})
	If nReg > 0                                      
		If (cAliasSD3)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aProdmAQ[nReg,3] += (cAliasSD3)->PRODUZIDO
      	ElseIf (cAliasSD3)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aProdmAQ[nReg,4] += (cAliasSD3)->PRODUZIDO
       	ElseIf (cAliasSD3)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
       		aProdmAQ[nReg,5] += (cAliasSD3)->PRODUZIDO
       	Endif	
    Else
		aadd(aProdmAQ,{(cAliasSD3)->EMISSAO,; //1
				   (cAliasSD3)->MAQUINA,;     //2
				   Iif((cAliasSD3)->EMISSAO = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),(cAliasSD3)->PRODUZIDO,0),; //3
				   Iif((cAliasSD3)->EMISSAO = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),(cAliasSD3)->PRODUZIDO,0),; //4
				   Iif((cAliasSD3)->EMISSAO = SUBSTR(DTOS(MV_PAR02),1,6),(cAliasSD3)->PRODUZIDO,0)})   //5
	EndIf				   
	(cAliasSD3)->(DbSkip())
EndDo
DbSelectArea(cAliasSD3)
dbCloseArea()
For nI:= 1 to len(aProdmAQ)
	If nLin >= 60
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5],aL[7]},,,@nLin)
	EndIf	
	FmtLin({Alltrim(aProdmAQ[nI,2]),;
		    Transform(aProdmAQ[nI,3],"99,999,999.999"),;
			Transform(aProdmAQ[nI,4],"99,999,999.999"),;
			Transform(aProdmAQ[nI,5],"99,999,999.999")},aL[14],"",,@nLin)
Next nI
FmtLin(,{aL[15],aL[16],aL[17]},,,@nLin)
//PARADA DE MAQUINA PEGAR PELO SH6 COM O CAMPO H6_MOTIVO <> ""  tabela 44 no sx5 para motivo de paradas
cQuery := " SELECT SUBSTRING(H6.H6_DTAPONT,5,2) AS EMISSAO,SUBSTRING(H1.H1_DESCRI,1,20) AS MAQUINA, H6.H6_TEMPO "
cQuery += " FROM "
cQuery += RetSqlName('SH6') + " H6, "   
cQuery += RetSqlName('SH1') + " H1 " 
cQuery += " WHERE "
cQuery += " H6.H6_FILIAL = '"+xFilial("SH6")+"'  " 
cQuery += " AND H6.H6_MOTIVO <> '  '  "
cQuery += " AND SUBSTRING(H6.H6_DTAPONT,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(H6.H6_DTAPONT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND H6.D_E_L_E_T_ = ' '  "             
cQuery += " AND H6.H6_RECURSO = H1.H1_CODIGO " 
cQuery += " AND H1.H1_FILIAL = '"+xFilial("SH1")+"'  " 
cQuery += " AND H1.D_E_L_E_T_ = ' '  "    
cQuery += " ORDER BY SUBSTRING(H6.H6_DTAPONT,5,2),H1.H1_DESCRI ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSH6,.T.,.T.)},"Aguarde...","Processando Parada por Maquina...")
For nSH6 := 1 To len(aStruSH6)
	If aStruSH6[nSH6][2] <> "C" .And. FieldPos(aStruSH6[nSH6][1])<>0
		TcSetField(cAliasSH6,aStruSH6[nSH6][1],aStruSH6[nSH6][2],aStruSH6[nSH6][3],aStruSH6[nSH6][4])
	EndIf
Next nSH6
For nSH1 := 1 To len(aStruSH1)
	If aStruSH1[nSH1][2] <> "C" .And. FieldPos(aStruSH1[nSH1][1])<>0
		TcSetField(cAliasSH6,aStruSH1[nSH1][1],aStruSH1[nSH1][2],aStruSH1[nSH1][3],aStruSH1[nSH1][4])
	EndIf
Next nSH1
DbSelectArea(cAliasSH6)
Do While (cAliasSH6)->(!EOF())
	nReg := aScan(aParadaMaq,{|x| x[2]==(cAliasSH6)->MAQUINA})
	If nReg > 0
		If Val((cAliasSH6)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))          
			If !Empty(Alltrim(aParadaMaq[nReg,3]))
				aadd(aParadaMaq,{(cAliasSH6)->EMISSAO,; //1
				   (cAliasSH6)->MAQUINA,;     //2
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),(cAliasSH6)->H6_TEMPO,0),; //3
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),(cAliasSH6)->H6_TEMPO,0),; //4
				   Iif(Val((cAliasSH6)->EMISSAO) = nMes,(cAliasSH6)->H6_TEMPO,0)})   //5
       		Else 
				aParadaMaq[nReg,3] := (cAliasSH6)->H6_TEMPO	       	
	       	Endif	
       	ElseIf Val((cAliasSH6)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1) 
			If !Empty(Alltrim(aParadaMaq[nReg,3]))
				aadd(aParadaMaq,{(cAliasSH6)->EMISSAO,; //1
				   (cAliasSH6)->MAQUINA,;     //2
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),(cAliasSH6)->H6_TEMPO,0),; //3
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),(cAliasSH6)->H6_TEMPO,0),; //4
				   Iif(Val((cAliasSH6)->EMISSAO) = nMes,(cAliasSH6)->H6_TEMPO,0)})   //5
       		Else 
	      		aParadaMaq[nReg,4] := (cAliasSH6)->H6_TEMPO
	      	EndIf	
       	Else    
			If !Empty(Alltrim(aParadaMaq[nReg,3]))
				aadd(aParadaMaq,{(cAliasSH6)->EMISSAO,; //1
				   (cAliasSH6)->MAQUINA,;     //2
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),(cAliasSH6)->H6_TEMPO,0),; //3
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),(cAliasSH6)->H6_TEMPO,0),; //4
				   Iif(Val((cAliasSH6)->EMISSAO) = nMes,(cAliasSH6)->H6_TEMPO,0)})   //5
       		Else 
				aParadaMaq[nReg,5] := (cAliasSH6)->H6_TEMPO
			EndIf	
       	EndIf	
    Else
		aadd(aParadaMaq,{(cAliasSH6)->EMISSAO,; //1
				   (cAliasSH6)->MAQUINA,;     //2
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),(cAliasSH6)->H6_TEMPO,0),; //3
				   Iif(Val((cAliasSH6)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),(cAliasSH6)->H6_TEMPO,0),; //4
				   Iif(Val((cAliasSH6)->EMISSAO) = nMes,(cAliasSH6)->H6_TEMPO,0)})   //5
	EndIf				   
	(cAliasSH6)->(DbSkip())
EndDo
DbSelectArea(cAliasSH6)
dbCloseArea()
For nI:= 1 to len(aParadaMaq)
	If nLin >= 60
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                                    
			FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin) 
			FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	EndIf	
	FmtLin({Alltrim(aParadaMaq[nI,2]),;
		    Transform(aParadaMaq[nI,3],"999,999.999"),;
			Transform(aParadaMaq[nI,4],"999,999.999"),;
			Transform(aParadaMaq[nI,5],"999,999.999")},aL[18],"",,@nLin)
Next nI
FmtLin(,{aL[19]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
FmtLin(,{aL[20],aL[21]},,,@nLin)
//SELECT BC_QTDDEST,* FROM SBC020 WHERE BC_FILIAL = '02'  MOTIVO DA PERDA TABELA SX5 43
//SELECT BC_QTDDEST,BC_TIPO,BC_MOTIVO,BC_CODDEST,BC_RECURSO,* FROM SBC020 WHERE BC_FILIAL = '02' AND BC_DATA >= '20060901' POR MAQUINA,MOTIVO,
//PRODUTO DESTINO.
cQuery := " SELECT SUBSTRING(BC.BC_DATA,5,2) AS EMISSAO ,SUBSTRING(H1.H1_DESCRI,1,20) AS MAQUINA, SUM(BC.BC_QTDDEST) AS PERDA " 
cQuery += " FROM " 
cQuery += RetSqlName('SBC') + " BC, " 
cQuery += RetSqlName('SH1') + " H1 " 
cQuery += " WHERE " 
cQuery += " BC.BC_FILIAL = '"+xFilial("SBC")+"'  " 
cQuery += " AND SUBSTRING(BC.BC_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(BC.BC_DATA,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND BC.D_E_L_E_T_ = ' '  "             
cQuery += " AND BC.BC_RECURSO = H1.H1_CODIGO " 
cQuery += " AND H1.H1_FILIAL = '"+xFilial("SH1")+"'  " 
cQuery += " AND H1.D_E_L_E_T_ = ' '  "                           
cQuery += " GROUP BY SUBSTRING(BC.BC_DATA,5,2),H1.H1_DESCRI "
cQuery += " ORDER BY SUBSTRING(BC.BC_DATA,5,2),H1.H1_DESCRI ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSBC,.T.,.T.)},"Aguarde...","Processando Perda por Maquina...")
For nSBC := 1 To len(aStruSBC)
	If aStruSBC[nSBC][2] <> "C" .And. FieldPos(aStruSBC[nSBC][1])<>0
		TcSetField(cAliasSBC,aStruSBC[nSBC][1],aStruSBC[nSBC][2],aStruSBC[nSBC][3],aStruSBC[nSBC][4])
	EndIf
Next nSBC
For nSH1 := 1 To len(aStruSH1)
	If aStruSH1[nSH1][2] <> "C" .And. FieldPos(aStruSH1[nSH1][1])<>0
		TcSetField(cAliasSBC,aStruSH1[nSH1][1],aStruSH1[nSH1][2],aStruSH1[nSH1][3],aStruSH1[nSH1][4])
	EndIf
Next nSH1
DbSelectArea(cAliasSBC)
Do While (cAliasSBC)->(!EOF())
	nReg := aScan(aPerdaMaq,{|x| x[2]==(cAliasSBC)->MAQUINA})
	If nReg > 0
		If Val((cAliasSBC)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))          
				aPerdaMaq[nReg,3] := (cAliasSBC)->PERDA
       	ElseIf Val((cAliasSBC)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1) 
	      		aPerdaMaq[nReg,4] := (cAliasSBC)->PERDA
       	Else    
				aPerdaMaq[nReg,5] := (cAliasSBC)->PERDA
      	EndIf	
    Else
		aadd(aPerdaMaq,{(cAliasSBC)->EMISSAO,; //1
				   (cAliasSBC)->MAQUINA,;     //2
				   Iif(Val((cAliasSBC)->EMISSAO) = Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),(cAliasSBC)->PERDA,0),; //3
				   Iif(Val((cAliasSBC)->EMISSAO) = Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),(cAliasSBC)->PERDA,0),; //4
				   Iif(Val((cAliasSBC)->EMISSAO) = nMes  ,(cAliasSBC)->PERDA,0)})   //5
	EndIf				   
	(cAliasSBC)->(DbSkip())
EndDo
DbSelectArea(cAliasSBC)
dbCloseArea()
For nI:= 1 to len(aPerdaMaq)
	If nLin >= 60
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                                    
			FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin) 
			FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	EndIf	
	FmtLin({Alltrim(aPerdaMaq[nI,2]),;
		    Transform(aPerdaMaq[nI,3],"999,999.999"),;
			Transform(aPerdaMaq[nI,4],"999,999.999"),;
			Transform(aPerdaMaq[nI,5],"999,999.999")},aL[22],"",,@nLin)
Next nI
FmtLin(,{aL[23]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUBSTRING(SB9.B9_DATA,1,6) AS MES,SUM(SB9.B9_QINI) as QUANT, SUM(SB9.B9_QINI*SB9.B9_CUSTD) AS CUSTO "
cQuery += " FROM "
cQuery += RetSqlName('SB9') + " SB9, " 
cQuery += RetSqlName('SB1') + " SB1 "
cQuery += " WHERE SB9.D_E_L_E_T_ <> '*' AND "
cQuery += " SB9.B9_FILIAL = '"+xFilial("SB9")+"' AND "               
cQuery += " SUBSTRING(SB9.B9_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(SB9.B9_DATA,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " SUBSTRING(SB9.B9_COD,1,1) <> 'R' AND "
cQuery += " SB9.B9_LOCAL IN ('01') AND "
cQuery += " SB1.B1_FILIAL = '"+xFilial("SB1")+"' AND "               
cQuery += " SB1.D_E_L_E_T_ <> '*' AND "
cQuery += " SB1.B1_COD = SB9.B9_COD AND " 
cQuery += " SB1.B1_TIPO = 'MP' "
cQuery += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
cQuery += " ORDER BY SUBSTRING(SB9.B9_DATA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSB9,.T.,.T.)},"Aguarde...","Processando Materia Prima...")
For nSB9 := 1 To len(aStruSB9)
	If aStruSB9[nSB9][2] <> "C" .And. FieldPos(aStruSB9[nSB9][1])<>0
		TcSetField(cAliasSB9,aStruSB9[nSB9][1],aStruSB9[nSB9][2],aStruSB9[nSB9][3],aStruSB9[nSB9][4])
	EndIf
Next nSB9
DbSelectArea(cAliasSB9)
aadd(aMateriaP,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 
aadd(aMateriaP,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aMateriaP,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
Do While (cAliasSB9)->(!EOF())
		If (cAliasSB9)->MES == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aMateriaP[1,2] := (cAliasSB9)->QUANT 
       		aMateriaP[1,3] := (cAliasSB9)->CUSTO
       	ElseIf (cAliasSB9)->MES == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aMateriaP[2,2] := (cAliasSB9)->QUANT 
       		aMateriaP[2,3] := (cAliasSB9)->CUSTO
       	ElseIf (cAliasSB9)->MES == SUBSTR(DTOS(MV_PAR02),1,6)
       		aMateriaP[3,2] := (cAliasSB9)->QUANT 
       		aMateriaP[3,3] := (cAliasSB9)->CUSTO
       	EndIf	
	(cAliasSB9)->(DbSkip())
EndDo
DbSelectArea(cAliasSB9)
dbCloseArea() 
FmtLin({Transform(aMateriaP[1,2],"99,999,999.999"),;
		Transform(aMateriaP[2,2],"99,999,999.999"),;
		Transform(aMateriaP[3,2],"99,999,999.999")},aL[24],"",,@nLin)
FmtLin(,{aL[25]},,,@nLin)
FmtLin({Transform(aMateriaP[1,3],"99,999,999.99"),;
		Transform(aMateriaP[2,3],"99,999,999.99"),;
		Transform(aMateriaP[3,3],"99,999,999.99")},aL[26],"",,@nLin)
FmtLin(,{aL[27]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUBSTRING(SB9.B9_DATA,1,6) AS MES,SUM(SB9.B9_QINI) as QUANT, SUM(SB9.B9_QINI*SB9.B9_CUSTD) AS CUSTO "
cQuery += " FROM "
cQuery += RetSqlName('SB9') + " SB9 "
cQuery += " WHERE SB9.D_E_L_E_T_ <> '*'  "
cQuery += " AND SB9.B9_FILIAL = '"+xFilial("SB9")+"'  "               
cQuery += " AND SUBSTRING(SB9.B9_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(SB9.B9_DATA,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND SB9.B9_LOCAL IN ('03','04') "
cQuery += " GROUP BY SUBSTRING(SB9.B9_DATA,1,6) "
cQuery += " ORDER BY SUBSTRING(SB9.B9_DATA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSB9,.T.,.T.)},"Aguarde...","Processando Produto Acabado...")
For nSB9 := 1 To len(aStruSB9)
	If aStruSB9[nSB9][2] <> "C" .And. FieldPos(aStruSB9[nSB9][1])<>0
		TcSetField(cAliasSB9,aStruSB9[nSB9][1],aStruSB9[nSB9][2],aStruSB9[nSB9][3],aStruSB9[nSB9][4])
	EndIf
Next nSB9
aadd(aMateriaA,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 
aadd(aMateriaA,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aMateriaA,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
Do While (cAliasSB9)->(!EOF())
		If (cAliasSB9)->MES == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aMateriaA[1,2] := (cAliasSB9)->QUANT 
       		aMateriaA[1,3] := (cAliasSB9)->CUSTO
       	ElseIf (cAliasSB9)->MES == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aMateriaA[2,2] := (cAliasSB9)->QUANT 
       		aMateriaA[2,3] := (cAliasSB9)->CUSTO
       	ElseIf (cAliasSB9)->MES == SUBSTR(DTOS(MV_PAR02),1,6)
       		aMateriaA[3,2] := (cAliasSB9)->QUANT 
       		aMateriaA[3,3] := (cAliasSB9)->CUSTO
       	EndIf	
	(cAliasSB9)->(DbSkip())
EndDo
DbSelectArea(cAliasSB9)
dbCloseArea() 
FmtLin({Transform(aMateriaA[1,2],"99,999,999.999"),;
		Transform(aMateriaA[2,2],"99,999,999.999"),;
		Transform(aMateriaA[3,2],"99,999,999.999")},aL[28],"",,@nLin)
FmtLin(,{aL[29]},,,@nLin)
FmtLin({Transform(aMateriaA[1,3],"99,999,999.99"),;
		Transform(aMateriaA[2,3],"99,999,999.99"),;
		Transform(aMateriaA[3,3],"99,999,999.99")},aL[30],"",,@nLin)
FmtLin(,{aL[31]},,,@nLin)
cQuery := " SELECT RD_DATARQ , SUM(RD_HORAS) AS TOTAL FROM " 
cQuery += RetSqlName('SRD') 
cQuery += " WHERE RD_DATARQ >=  '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND RD_DATARQ <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND RD_CC >= '12301' AND RD_CC <= '12322' " 
cQuery += " AND RD_PD IN ('002','031','033','034','036') " 
cQuery += " AND RD_FILIAL = '"+xFilial("SRD")+"' "
cQuery += " GROUP BY RD_DATARQ " 
cQuery += " ORDER BY RD_DATARQ ASC" 
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSRD,.T.,.T.)},"Aguarde...","Processando Horas Trabalhadas...")
For nSRD := 1 To len(aStruSRD)
	If aStruSRD[nSRD][2] <> "C" .And. FieldPos(aStruSRD[nSRD][1])<>0
		TcSetField(cAliasSRD,aStruSRD[nSRD][1],aStruSRD[nSRD][2],aStruSRD[nSRD][3],aStruSRD[nSRD][4])
	EndIf
Next nSRD
aadd(aProdutividade,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 
aadd(aProdutividade,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aProdutividade,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
DbSelectArea(cAliasSRD)
Do While (cAliasSRD)->(!EOF())
	If (cAliasSRD)->RD_DATARQ == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
   		aProdutividade[1,2] := (cAliasSRD)->TOTAL
   	ElseIf (cAliasSRD)->RD_DATARQ == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
   		aProdutividade[2,2] := (cAliasSRD)->TOTAL
   	ElseIf (cAliasSRD)->RD_DATARQ == SUBSTR(DTOS(MV_PAR02),1,6)
   		aProdutividade[3,2] := (cAliasSRD)->TOTAL
   	EndIf	
	(cAliasSRD)->(DbSkip())
EndDo
DbSelectArea(cAliasSRD)
dbCloseArea()
FmtLin({Transform(aPRODUZIDO[1,2]/(aProdutividade[1,2]),"9,999,999.999"),;
        Transform(aPRODUZIDO[2,2]/(aProdutividade[2,2]),"9,999,999.999"),;
		Transform(aPRODUZIDO[3,2]/(aProdutividade[3,2]),"9,999,999.999")},aL[32],"",,@nLin)
FmtLin(,{aL[33]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '700303' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Materia Prima Nacional...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[40],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[41]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '700306' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Materia Prima Importada...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[42],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[43]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '700307' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Tributos de Importacao...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[44],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[45]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '700308' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Despesas de Importacao...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[46],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[47]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600201' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete de Importacao...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[48],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[49]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200308','200408','200312','200324','200325','200326','200327','200328') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Insumos...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[50],"",,@nLin)
	nTotal1 += (cAliasSE7)->VALOR1
	nTotal2 += (cAliasSE7)->VALOR2
	nTotal3 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[51]},,,@nLin)	
FmtLin({Transform(nTotal1,"@R 99,999,999.99"),;
	    Transform(nTotal2,"@R 99,999,999.99"),;
		Transform(nTotal3,"@R 99,999,999.99")},aL[52],"",,@nLin)
FmtLin(,{aL[53]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100301','100302','100303','100304','100305','100309','100310','100313','100306','100307','100311',"
cQuery += " '100401','100402','100403','100404','100405','100406','100407','100409','100410','100411') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Mao de Obra...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[54],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[55]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200311','200411') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Energia...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[56],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[57]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200310','200410') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Agua...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[58],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[59]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " ((E7_NATUREZ >= '100301' AND E7_NATUREZ <= '100317') OR "
cQuery += " (E7_NATUREZ >= '100401' AND E7_NATUREZ <= '100417') OR " 
cQuery += " (E7_NATUREZ >= '200301' AND E7_NATUREZ <= '200377') OR "
cQuery += " (E7_NATUREZ >= '200401' AND E7_NATUREZ <= '200421') OR " 
cQuery += " (E7_NATUREZ >= '700303' AND E7_NATUREZ <= '700309') OR "
cQuery += " E7_NATUREZ IN ('600201','600301','600501','700301')) AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Gastos Gerais Coml...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[34],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[35]},,,@nLin)	
cQuery := " SELECT SUM(E7_VAL"+aMes[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))]+") AS VALOR1 , SUM(E7_VAL"+aMes[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)]+")  AS VALOR2 ,SUM(E7_VAL"+aMes[Month(MV_PAR02)]+ ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " ((E7_NATUREZ >= '100301' AND E7_NATUREZ <= '100317') OR "
cQuery += " (E7_NATUREZ >= '100401' AND E7_NATUREZ <= '100417') OR " 
cQuery += " (E7_NATUREZ >= '200301' AND E7_NATUREZ <= '200377') OR "
cQuery += " (E7_NATUREZ >= '200401' AND E7_NATUREZ <= '200421') OR " 
cQuery += " (E7_NATUREZ >= '700303' AND E7_NATUREZ <= '700309') OR "
cQuery += " E7_NATUREZ IN ('600201','600301','600501','700301')) AND " 
cQuery += " E7_NATUREZ NOT IN ('100301','100309','100310','100313','100317','100401','100409','100410','100417','200313') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Meta Orcamento Ind....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	If nLin >= 60
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5],aL[7]},,,@nLin)
	EndIf	
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
		    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
			Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[36],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[37]},,,@nLin)	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " ((E7_NATUREZ >= '100301' AND E7_NATUREZ <= '100317') OR "
cQuery += " (E7_NATUREZ >= '100401' AND E7_NATUREZ <= '100417') OR " 
cQuery += " (E7_NATUREZ >= '200301' AND E7_NATUREZ <= '200377') OR "
cQuery += " (E7_NATUREZ >= '200401' AND E7_NATUREZ <= '200421') OR " 
cQuery += " (E7_NATUREZ >= '700303' AND E7_NATUREZ <= '700309') OR "
cQuery += " E7_NATUREZ IN ('600201','600301','600501','700301')) AND " 
cQuery += " E7_NATUREZ NOT IN ('100301','100309','100310','100313','100317','100401','100409','100410','100417','200313') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Realizado Orcamento Ind....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[7]},,,@nLin)
EndIf	
FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[38],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[39]},,,@nLin)
RestArea(aArea)

RETURN(.T.)
