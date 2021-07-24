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

User Function CRF010J(nLin,nMes,Cabec1,Cabec2,Titulo,aQtdVend) 

Local aArea 	:= GetArea()
Local cQuery 	:= ""
Local nI 		:= 0
Local cAliasSD2 := "cAliasSD2"
Local cAliasSF2 := "cAliasSF2"
Local nSD2	 	:= 0
Local aStruSD2 	:= SD2->(dbStruct())
Local aStruSF2 	:= SF2->(dbStruct())
Local nSF2	 	:= 0
Local aGastosCom:= {}
Local aFatLFDIF := {}  
Local aFatFTDIF := {}  
Local aFatOutros:= {}  
Local nX := 0
// GERANDO SUCATA
cQuery := " SELECT SUBSTRING(D2.D2_EMISSAO,1,6) AS EMISSAO, SUM(D2_CUSTO1) AS REPOSICAO  FROM " + RetSqlName('SD2') + " D2 WHERE "
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " D2.D2_TIPO = 'N' AND D2.D2_LOCAL IN ('03','04') AND D2.D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " D2.D2_COD  IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += "'057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += "'038675','044926','057558','037503') AND "
cQuery += " D2.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D2.D2_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.T.,.T.)},"Aguarde...","Processando Sucata ...")
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cAliasSD2,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
aadd(aSucata,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aSucata,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aSucata,{SUBSTR(DTOS(MV_PAR02),1,6),0,0}) 
DbSelectArea(cAliasSD2)
If 	(cAliasSD2)->(!EOF())
	Do While (cAliasSD2)->(!EOF())
			If (cAliasSD2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
				aSucata[1,3]:=(cAliasSD2)->REPOSICAO
			elseIf (cAliasSD2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
				aSucata[2,3]:=(cAliasSD2)->REPOSICAO
			elseIf (cAliasSD2)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
				aSucata[3,3]:=(cAliasSD2)->REPOSICAO
			EndIf	
		(cAliasSD2)->(DbSkip())
	EndDo
EndIf	
DbSelectArea(cAliasSD2)
dbCloseArea()

cQuery := " SELECT SUBSTRING(F2.F2_EMISSAO,1,6) AS EMISSAO, SUM(F2_VALFAT) AS TOTAL FROM "
cQuery +=   RetSqlName('SF2') + " F2 WHERE F2.F2_FILIAL  = '"+xFilial("SF2")+"' AND F2.D_E_L_E_T_ = ' ' AND "
cQuery += " F2.F2_EMISSAO >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"01' AND " 
cQuery += " F2.F2_EMISSAO <= '" + DTOS(MV_PAR02) +"' AND F2.F2_TIPO    = 'N' AND "          
cQuery += " F2.F2_DOC+F2.F2_SERIE+F2.F2_EMISSAO IN (SELECT D2.D2_DOC+D2.D2_SERIE+D2.D2_EMISSAO FROM "
cQuery +=   RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND "
cQuery += " D2.D2_TIPO = 'N' AND  D2.D2_LOCAL IN ('03','04') AND "
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " D2.D2_COD  IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += "'057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += "'038675','044926','057558','037503') AND "
cQuery += " D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.T.,.T.)},"Aguarde...","Processando Sucata ...")
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cAliasSD2,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea(cAliasSD2)
If 	(cAliasSD2)->(!EOF())
	Do While (cAliasSD2)->(!EOF())
			If (cAliasSD2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
				aSucata[1,2]:=(cAliasSD2)->TOTAL
			elseIf (cAliasSD2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
				aSucata[2,2]:=(cAliasSD2)->TOTAL
			elseIf (cAliasSD2)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
				aSucata[3,2]:=(cAliasSD2)->TOTAL
			EndIf	
		(cAliasSD2)->(DbSkip())
	EndDo
EndIf	
DbSelectArea(cAliasSD2)
dbCloseArea()
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6],aL[7]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[9]},,,@nLin)
EndIf	
FmtLin(,{aL[34]},,,@nLin)
FmtLin({Transform(aSucata[1,2],"@R 99,999,999.99"),;
	    Transform(aSucata[2,2],"@R 99,999,999.99"),;
		Transform(aSucata[3,2],"@R 99,999,999.99")},aL[35],"",,@nLin)
FmtLin(,{aL[36]},,,@nLin)	 
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6],aL[7]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[9]},,,@nLin)
EndIf	
FmtLin(,{aL[42],aL[41]},,,@nLin)	 
cQuery := " SELECT SUBSTRING(F2_EMISSAO,1,6) AS EMISSAO,F2_COND,SUM(F2_VALFAT) AS VALOR  FROM "                  
cQuery += RetSqlName('SF2') + " F2 "
cQuery += " WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"' AND SUBSTRING(F2.F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND F2.D_E_L_E_T_ = ' ' "
cQuery += " AND F2.F2_TIPO = 'N' AND F2.F2_DOC+F2.F2_SERIE IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery += RetSqlName('SD2') + " D2 WHERE D2.D2_LOCAL IN ('03','04') AND D2.D2_COD  IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += "'057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += "'038675','044926','057558','037503') AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"'  AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(F2_EMISSAO,1,6),F2_COND "
cQuery += " ORDER BY SUBSTRING(F2_EMISSAO,1,6),F2_COND ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSF2,.T.,.T.)},"Aguarde...","Processando Condicao de Faturamento Sucata...")
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cAliasSF2,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea(cAliasSF2)
aCondFat1  := {}  
Do While (cAliasSF2)->(!EOF())
	nReg := aScan(aCondFat1,{|x| x[2]==Posicione("SE4",1,xFilial("SE4")+(cAliasSF2)->F2_COND,"E4_DESCRI")})
	If nReg >0 
		If (cAliasSF2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aCondFat1[nReg,3] := (cAliasSF2)->VALOR
		ElseIf (cAliasSF2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aCondFat1[nReg,4] := (cAliasSF2)->VALOR
		ElseIf (cAliasSF2)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aCondFat1[nReg,5] := (cAliasSF2)->VALOR
		EndIf	
	Else
		If (cAliasSF2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aadd(aCondFat1,{(cAliasSF2)->EMISSAO,Posicione("SE4",1,xFilial("SE4")+(cAliasSF2)->F2_COND,"E4_DESCRI"),(cAliasSF2)->VALOR,0,0}) 
		ElseIf (cAliasSF2)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aadd(aCondFat1,{(cAliasSF2)->EMISSAO,Posicione("SE4",1,xFilial("SE4")+(cAliasSF2)->F2_COND,"E4_DESCRI"),0,(cAliasSF2)->VALOR,0}) 
		ElseIf (cAliasSF2)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aadd(aCondFat1,{(cAliasSF2)->EMISSAO,Posicione("SE4",1,xFilial("SE4")+(cAliasSF2)->F2_COND,"E4_DESCRI"),0,0,(cAliasSF2)->VALOR}) 		
		EndIf	
	Endif
	(cAliasSF2)->(DbSkip())
EndDo
DbSelectArea(cAliasSF2)
dbCloseArea()
For nI := 1 To len(aCondFat1)
	If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[6],aL[7]},,,@nLin)
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5],aL[9]},,,@nLin)
	EndIf	
	FmtLin({aCondFat1[nI,2],Transform(aCondFat1[nI,3],"@R 99,999,999.99"),Transform(aCondFat1[nI,4],"@R 99,999,999.99"),;
			Transform(aCondFat1[nI,5],"@R 99,999,999.99")},aL[43],"",,@nLin)
Next nI
FmtLin(,{aL[44]},,,@nLin)
FmtLin({Iif(Val(aSucata[1,1]) == Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),Transform(aSucata[1,3],"@R 99,999,999.99"),0),;
	    Iif(Val(aSucata[2,1]) == Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),Transform(aSucata[2,3],"@R 99,999,999.99"),0),;
		Iif(Val(aSucata[3,1]) == nMes,Transform(aSucata[3,3],"@R 99,999,999.99"),0)},aL[37],"",,@nLin)
FmtLin(,{aL[38]},,,@nLin)	 
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[6],aL[7]},,,@nLin)
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5],aL[9]},,,@nLin)
EndIf	
FmtLin({Transform(NoRound((((aSucata[1,2])/(aSucata[1,3]))-1)*100,2),"@R 99,999,999.99"),;
	    Transform(NoRound((((aSucata[2,2])/(aSucata[2,3]))-1)*100,2),"@R 99,999,999.99"),;
		Transform(NoRound((((aSucata[3,2])/(aSucata[3,3]))-1)*100,2),"@R 99,999,999.99")},aL[39],"",,@nLin)
FmtLin(,{aL[40]},,,@nLin)	 		          
nLin := 0
Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
nLin:=PROW()+1
U_CRF010E(@nLin,nMes,Cabec1,Cabec2,Titulo,aQtdVend)  
nLin := 0
nLin++
U_CRF010D(@nLin,nMes,Cabec1,Cabec2,Titulo) 
nLin := 0
nLin++
U_CRF010C(@nLin,nMes,Cabec1,Cabec2,Titulo)  
nLin := 0
nLin++
U_CRF010I(@nLin,nMes,Cabec1,Cabec2,Titulo) 
RestArea(aArea)
Return(.T.)
