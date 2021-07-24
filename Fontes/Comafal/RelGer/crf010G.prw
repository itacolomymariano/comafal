#INCLUDE "RWMAKE.CH"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³RunReport ºAutor  ³Eduardo Zanardo     º Data ³  03/31/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP7                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CRF010G(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,nLin,nMes)

Local aArea := GetArea()
Local aStruSD2 	:= SD2->(dbStruct()) 
Local nSD2 		:= 0
Local aStruSD1 	:= SD1->(dbStruct()) 
Local nSD1 		:= 0
Local cQuery    := "" 
Local nTotalvlr := 0
Local nSF3		:= 0
Local aStruSF3  := SF3->(dbStruct())
U_RFinLay6()
nlin:= 61
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[39],aL[2],aL[1],aL[42],aL[4]},,,@nLin)
EndIf
cQuery := " SELECT D2.D2_CF, SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS TOTAL FROM "
cQuery +=   RetSqlName('SD2') + " D2 " 
cQuery += " WHERE  D2.D2_FILIAL = '"+xFilial("SD2")+"' AND  SUBSTRING(D2.D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) +"' AND " 
cQuery += " D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND "
cQuery += " D2.D2_CF NOT IN (" + Alltrim(cCFOPVenda) + ") AND  SUBSTRING(D2_CF,1,1) >= '5' AND   D2.D_E_L_E_T_ = ' ' "    
cQuery += " GROUP BY D2.D2_CF "
cQuery += " ORDER BY D2.D2_CF ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("c010GSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2 
DbSelectArea("c010GSD2")
c010GSD2->(dbGotop())
Do While c010GSD2->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[39],aL[2],aL[1],aL[42],aL[4]},,,@nLin)
	EndIf
	FmtLin({c010GSD2->D2_CF,;
			fDesc("SX5","13"+c010GSD2->D2_CF,"X5_DESCRI"),;
			Transform(c010GSD2->TOTAL,"@R 99,999,999.99")},aL[06],"",,@nLin)
	nTotalvlr += c010GSD2->TOTAL
	c010GSD2->(DbSkip())			
EndDo 
nTotalvlr := 0
FmtLin(,{aL[9],aL[40],aL[12],aL[13]},,,@nLin)                                
DbSelectArea("c010GSD2")                     
dbCloseArea()
cQuery := " SELECT F3_EMISSAO,F3_NFISCAL,F3_SERIE,SUM(F3_VALCONT) AS VALOR FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) + "' "
cQuery += " AND F3_ESPECIE = 'NF' "
cQuery += " AND F3_CFO IN (" + Alltrim(cCFOPVenda) + ") " 
cQuery += " AND F3_OBSERV <> 'NF CANCELADA' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery += "	GROUP BY F3_EMISSAO,F3_NFISCAL,F3_SERIE "
cQuery += "	ORDER BY F3_EMISSAO,F3_NFISCAL,F3_SERIE ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField("c010GSF3",aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("c010GSF3")
If c010GSF3->(!EOF()) 	                                                    
	Do While c010GSF3->(!EOF()) 
		If nLin >= 60
			nLin := 0
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                          
			FmtLin(,{aL[40],aL[12],aL[13]},,,@nLin) 
		EndIf
		DbSelectArea("SF2")
		DbSetOrder(1)
		If DbSeek(xFilial("SF2")+c010GSF3->F3_NFISCAL+c010GSF3->F3_SERIE)
			If SF2->F2_VALFAT <> c010GSF3->VALOR
				FmtLin({c010GSF3->F3_NFISCAL,;
						c010GSF3->F3_SERIE,;         
						"Valor Divergente",;
						Transform(c010GSF3->VALOR,"@R 99,999,999.99")},aL[14],"",,@nLin)
			EndIf
		Else				
			FmtLin({c010GSF3->F3_NFISCAL,;
					c010GSF3->F3_SERIE,;         
					"Nao tem Venda",;
					Transform(c010GSF3->VALOR,"@R 99,999,999.99")},aL[14],"",,@nLin)
		EndIf
		DbSelectArea("c010GSF3")
		c010GSF3->(DbSkip())
	EndDo
EndIf	
FmtLin(,{aL[11],aL[41]},,,@nLin) 
DbSelectArea("c010GSF3")
dbCloseArea()
cQuery := " SELECT SUM(F3_VALCONT) AS TOTAL FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND F3_OBSERV <> 'NF CANCELADA' "
cQuery += " AND F3_CFO NOT IN (" + Alltrim(cCFOPVenda) + ") " 
cQuery += " AND SUBSTRING(F3_CFO,1,1) >= '5' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField("c010GSF3",aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("c010GSF3")
Do While c010GSF3->(!EOF())
	nTotalvlr += c010GSF3->TOTAL
	c010GSF3->(DbSkip())
EndDo
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[1]},,,@nLin)
EndIf
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[10],"",,@nLin)
FmtLin(,{aL[19],aL[20],aL[21],aL[22],aL[43],aL[23],aL[24]},,,@nLin)                           
DbSelectArea("c010GSF3")
dbCloseArea()
cQuery := "SELECT D1.D1_CF, "
cQuery += "SUM(D1_TOTAL+D1_VALIPI+D1_ICMSRET+D1_VALFRE+D1_SEGURO+D1_DESPESA) AS TOTAL "
cQuery += "FROM "
cQuery +=  RetSqlName('SD1') + " D1 "
cQuery += "WHERE "
cQuery += "D1.D1_FILIAL = '"+xFilial("SD1")+"' AND "
cQuery += "SUBSTRING(D1.D1_DTDIGIT,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) +"' AND " 
cQuery += "SUBSTRING(D1.D1_CF,1,1) <= '3' AND "
cQuery += "D1.D_E_L_E_T_ = ' ' "
cQuery += "GROUP BY D1.D1_CF "
cQuery += "ORDER BY D1.D1_CF ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GNFS",.T.,.T.)
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("c010GNFS",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea("c010GNFS")
nTotalvlr:=0
Do While c010GNFS->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[20],aL[21],aL[22],aL[43],aL[23],aL[24]},,,@nLin)                           
	EndIf
	FmtLin({c010GNFS->D1_CF,;
			fDesc("SX5","13"+c010GNFS->D1_CF,"X5_DESCRI"),;
			Transform(c010GNFS->TOTAL,"@R 99,999,999.99")},aL[25],"",,@nLin)
	nTotalvlr += c010GNFS->TOTAL
	c010GNFS->(DbSkip())			
EndDo                   
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[44]},,,@nLin)
EndIf
FmtLin(,{aL[44],aL[45],aL[35],aL[36]},,,@nLin)
DbSelectArea("c010GNFS")
dbCloseArea()
cQuery := " SELECT D1_DOC,D1_SERIE,D1_DTDIGIT,SUM(D1_TOTAL+D1_VALIPI+D1_ICMSRET+D1_VALFRE+D1_SEGURO+D1_DESPESA) AS TOTAL FROM "
cQuery += RetSqlName('SD1')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " D1_FILIAL = '"+xFilial("SD1")+"' AND "
cQuery += " SUBSTRING(D1_DTDIGIT,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' AND "
cQuery += " D1_FILIAL = '"+xFilial("SD1")+"' AND "
cQuery += " SUBSTRING(D1_CF,1,1) <= '3' AND "
cQuery += " D1_DOC+D1_SERIE NOT IN (SELECT F3_NFISCAL+F3_SERIE FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND SUBSTRING(F3_CFO,1,1) <= '3' "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' )"
cQuery += "GROUP BY D1_DOC,D1_SERIE,D1_DTDIGIT "
cQuery += "ORDER BY D1_DOC,D1_SERIE,D1_DTDIGIT ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSD1",.T.,.T.)
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("c010GSD1",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea("c010GSD1")
Do While c010GSD1->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[44],aL[45],aL[35],aL[36]},,,@nLin)
	EndIf
	FmtLin({c010GSD1->D1_DOC,;
			c010GSD1->D1_SERIE,;
			Transform(c010GSD1->TOTAL,"@R 99,999,999.99")},aL[37],"",,@nLin)
	c010GSD1->(DbSkip())
EndDo                   
FmtLin(,{aL[38],aL[26]},,,@nLin)
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[27],"",,@nLin)
FmtLin(,{aL[28],aL[46]},,,@nLin)
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[46]},,,@nLin)
EndIf
FmtLin(,{aL[31],aL[32]},,,@nLin)
nTotalvlr := 0
DbSelectArea("c010GSD1")
dbCloseArea()
cQuery := " SELECT F3_NFISCAL,F3_SERIE,F3_EMISSAO,F3_VALCONT,F3_OBSERV FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) + "' "
cQuery += " AND SUBSTRING(F3_CFO,1,1) <= '3' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery += " AND F3_NFISCAL+F3_SERIE NOT IN (SELECT D1_DOC+D1_SERIE FROM "
cQuery += RetSqlName('SD1')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND SUBSTRING(D1_CF,1,1) <= '3'  "
cQuery += " AND SUBSTRING(D1_DTDIGIT,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND D1_FILIAL = '"+xFilial("SD1")+"' )"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField("c010GSF3",aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("c010GSF3")
If c010GSF3->(!EOF()) 	                                                    
	Do While c010GSF3->(!EOF()) 
		If nLin >= 60
			nLin := 0
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                          
			FmtLin(,{aL[46],aL[31],aL[32]},,,@nLin)
		EndIf
		FmtLin({c010GSF3->F3_NFISCAL,;
			c010GSF3->F3_SERIE,;
			c010GSF3->F3_OBSERV,;		         
			Transform(c010GSF3->F3_VALCONT,"@R 99,999,999.99")},aL[33],"",,@nLin)
		c010GSF3->(DbSkip())
	EndDo
EndIf	                
FmtLin(,{aL[30],aL[47]},,,@nLin)
nTotalvlr :=0
DbSelectArea("c010GSF3")
dbCloseArea()
cQuery := "SELECT SUM(F3_VALCONT) AS TOTAL FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_CFO,1,1) <= '3' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010GSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField("c010GSF3",aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("c010GSF3")
Do While c010GSF3->(!EOF())
	nTotalvlr += c010GSF3->TOTAL
	c010GSF3->(DbSkip())
EndDo                   
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[1]},,,@nLin)
EndIf
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[29],"",,@nLin)
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[1]},,,@nLin)
EndIf
FmtLin(,{aL[34]},,,@nLin)

DbSelectArea("c010GSF3")
dbCloseArea()

RestArea(aArea)

Return(.T.)
