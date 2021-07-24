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

Local aStruSD2 	:= SD2->(dbStruct()) 
Local nSD2 		:= 0
Local aStruSF2 	:= SF2->(dbStruct()) 
Local nSF2 		:= 0
Local aStruSD1 	:= SD1->(dbStruct()) 
Local nSD1 		:= 0
Local cAliasNFS := "cAliasNFS"
Local cAliasSD1 := "cAliasSD1"
Local cAliasSD2 := "cAliasSD2"
Local cQuery    := "" 
Local nTotalvlr := 0
Local cALiasSF3 := "cALiasSF3"
Local nSF3		:= 0
Local aStruSF3  := SF3->(dbStruct())
nLin := 0
Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
U_RFinLay6()
nLin:=PROW()+1
FmtLin(,{aL[39],aL[2],aL[1],aL[42],aL[4]},,,@nLin)
cQuery := " SELECT D2.D2_CF, SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS TOTAL FROM "
cQuery +=   RetSqlName('SD2') + " D2 " 
cQuery += " WHERE  D2.D2_FILIAL = '"+xFilial("SD2")+"' AND  SUBSTRING(D2.D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) +"' AND " 
cQuery += " D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND "
cQuery += " D2.D2_CF NOT IN (" + Alltrim(cCFOPVenda) + ") AND  SUBSTRING(D2_CF,1,1) >= '5' AND   D2.D_E_L_E_T_ = ' ' "    
cQuery += " GROUP BY D2.D2_CF "
cQuery += " ORDER BY D2.D2_CF ASC"
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasNFS,.T.,.T.)},"Aguarde...","Notas de Saida do Tipo(Devolucao, Beneficiamento, Comp. de Preco, Comp. de Impostos e outros")
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cAliasNFS,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2 
DbSelectArea(cAliasNFS)
Do While (cAliasNFS)->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[39],aL[2],aL[1],aL[42],aL[4]},,,@nLin)
	EndIf
	FmtLin({(cAliasNFS)->D2_CF,;
			fDesc("SX5","13"+(cAliasNFS)->D2_CF,"X5_DESCRI"),;
			Transform((cAliasNFS)->TOTAL,"@R 99,999,999.99")},aL[06],"",,@nLin)
	nTotalvlr += (cAliasNFS)->TOTAL
	(cAliasNFS)->(DbSkip())			
EndDo
DbSelectArea(cAliasNFS)
dbCloseArea()
FmtLin(,{aL[3],aL[5],aL[16],aL[17]},,,@nLin)
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[39],aL[16],aL[17]},,,@nLin)
EndIf
cQuery := " SELECT F2_DOC,F2_SERIE,F2_EMISSAO,F2_VALFAT AS TOTAL FROM "
cQuery += RetSqlName('SF2')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " F2_FILIAL = '"+xFilial("SF2")+"' AND "
cQuery += " SUBSTRING(F2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' AND "
cQuery += " F2_DOC+F2_SERIE IN (SELECT D2_DOC+D2_SERIE FROM "
cQuery += RetSqlName('SD2')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' AND "
cQuery += " D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " SUBSTRING(D2_CF,1,1) >= '5' AND "
cQuery += " D2_CF NOT IN (" + Alltrim(cCFOPVenda) + ")) AND "
cQuery += " F2_DOC+F2_SERIE NOT IN (SELECT F3_NFISCAL+F3_SERIE FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND F3_CFO NOT IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND SUBSTRING(F3_CFO,1,1) >= '5'  "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' )"
cQuery += "ORDER BY F2_DOC,F2_SERIE,F2_EMISSAO ASC"
	
cQuery:=ChangeQuery(cQuery)
	
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD2,.T.,.T.)},"Aguarde...","Processando Diferenca Faturamento (NF´s que constam apenas no Faturamento)...")

For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cAliasSD2,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
	
DbSelectArea(cAliasSD2)
Do While (cAliasSD2)->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[39],aL[16],aL[17]},,,@nLin)
	EndIf
	FmtLin({(cAliasSD2)->F2_DOC,;
			(cAliasSD2)->F2_SERIE,;         
			Transform((cAliasSD2)->TOTAL,"@R 99,999,999.99")},aL[18],"",,@nLin)
	(cAliasSD2)->(DbSkip())
EndDo
DbSelectArea(cAliasSD2)
dbCloseArea()
FmtLin(,{aL[7],aL[15]},,,@nLin)
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[8],"",,@nLin)
nTotalvlr := 0
FmtLin(,{aL[9],aL[40],aL[12],aL[13]},,,@nLin)                                
/*
cQuery := " SELECT F3_NFISCAL,F3_SERIE,F3_EMISSAO,F3_VALCONT,F3_OBSERV FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(nMes,2) + "' "
cQuery += " AND F3_CFO NOT IN (" + Alltrim(cCFOPVenda) + ") " 
cQuery += " AND SUBSTRING(F3_CFO,1,1) >= '5' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery += " AND F3_NFISCAL+F3_SERIE NOT IN (SELECT D2_DOC+D2_SERIE FROM "
cQuery += RetSqlName('SD2')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND D2_CF NOT IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND SUBSTRING(D2_CF,1,1) >= '5' "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND D2_FILIAL = '"+xFilial("SD2")+"' )"
cQuery += "	ORDER BY F3_NFISCAL,F3_SERIE,F3_EMISSAO ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSF3,.T.,.T.)},"Aguarde...","Processando Diferenca Fiscal (NF´s que constam apenas no Livro Fiscal)...")
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField(cAliasSF3,aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea(cAliasSF3)
If (cAliasSF3)->(!EOF()) 	                                                    
	Do While (cAliasSF3)->(!EOF()) 
		If nLin >= 60
			nLin := 0
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                          
			FmtLin(,{aL[40],aL[12],aL[13]},,,@nLin) 
		EndIf
		FmtLin({(cAliasSF3)->F3_NFISCAL,;
				(cAliasSF3)->F3_SERIE,;         
				(cAliasSF3)->F3_OBSERV,;
				Transform((cAliasSF3)->F3_VALCONT,"@R 99,999,999.99")},aL[14],"",,@nLin)
		(cAliasSF3)->(DbSkip())
	EndDo
EndIf	
DbSelectArea(cAliasSF3)
dbCloseArea()
*/
FmtLin(,{aL[11],aL[41]},,,@nLin) 
cQuery := " SELECT SUM(F3_VALCONT) AS TOTAL FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND F3_OBSERV <> 'NF CANCELADA' "
cQuery += " AND F3_CFO NOT IN (" + Alltrim(cCFOPVenda) + ") " 
cQuery += " AND SUBSTRING(F3_CFO,1,1) >= '5' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSF3,.T.,.T.)},"Aguarde...","Total no Livro Fiscal")
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField(cAliasSF3,aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea(cAliasSF3)
Do While (cAliasSF3)->(!EOF())
	nTotalvlr += (cAliasSF3)->TOTAL
	(cAliasSF3)->(DbSkip())
EndDo
DbSelectArea(cAliasSF3)
dbCloseArea()
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[1]},,,@nLin)
EndIf
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[10],"",,@nLin)
FmtLin(,{aL[19],aL[20],aL[21],aL[22],aL[43],aL[23],aL[24]},,,@nLin)                           
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
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasNFS,.T.,.T.)},"Aguarde...","Notas de Entrada do Tipo(Compras,Devolucao, Beneficiamento, Comp. de Preco, Comp. de Impostos e outros")
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField(cAliasNFS,aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea(cAliasNFS)
nTotalvlr:=0
Do While (cAliasNFS)->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[20],aL[21],aL[22],aL[43],aL[23],aL[24]},,,@nLin)                           
	EndIf
	FmtLin({(cAliasNFS)->D1_CF,;
			fDesc("SX5","13"+(cAliasNFS)->D1_CF,"X5_DESCRI"),;
			Transform((cAliasNFS)->TOTAL,"@R 99,999,999.99")},aL[25],"",,@nLin)
	nTotalvlr += (cAliasNFS)->TOTAL
	(cAliasNFS)->(DbSkip())			
EndDo
DbSelectArea(cAliasNFS)
dbCloseArea()      
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[44]},,,@nLin)
EndIf
FmtLin(,{aL[44],aL[45],aL[35],aL[36]},,,@nLin)
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
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSD1,.T.,.T.)},"Aguarde...","Diferenca Compras (Nf´s que  constam apenas no Compras)")
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField(cAliasSD1,aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea(cAliasSD1)
Do While (cAliasSD1)->(!EOF())
	If nLin >= 60
		nLin := 0
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                          
		FmtLin(,{aL[44],aL[45],aL[35],aL[36]},,,@nLin)
	EndIf
	FmtLin({(cAliasSD1)->D1_DOC,;
			(cAliasSD1)->D1_SERIE,;
			Transform((cAliasSD1)->TOTAL,"@R 99,999,999.99")},aL[37],"",,@nLin)
	(cAliasSD1)->(DbSkip())
EndDo
DbSelectArea(cAliasSD1)
dbCloseArea()
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
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSF3,.T.,.T.)},"Aguarde...","Processando Diferenca Fiscal (NF´s que constam apenas no Livro Fiscal)...")
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField(cAliasSF3,aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea(cAliasSF3)
If (cAliasSF3)->(!EOF()) 	                                                    
	Do While (cAliasSF3)->(!EOF()) 
		If nLin >= 60
			nLin := 0
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                          
			FmtLin(,{aL[46],aL[31],aL[32]},,,@nLin)
		EndIf
		FmtLin({(cAliasSF3)->F3_NFISCAL,;
			(cAliasSF3)->F3_SERIE,;
			(cAliasSF3)->F3_OBSERV,;		         
			Transform((cAliasSF3)->F3_VALCONT,"@R 99,999,999.99")},aL[33],"",,@nLin)
		(cAliasSF3)->(DbSkip())
	EndDo
EndIf	
DbSelectArea(cAliasSF3)
dbCloseArea()
FmtLin(,{aL[30],aL[47]},,,@nLin)
nTotalvlr :=0
cQuery := "SELECT SUM(F3_VALCONT) AS TOTAL FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_CFO,1,1) <= '3' "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSF3,.T.,.T.)},"Aguarde...","Compras no Livro Fiscal")
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField(cAliasSF3,aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea(cAliasSF3)
Do While (cAliasSF3)->(!EOF())
	nTotalvlr += (cAliasSF3)->TOTAL
	(cAliasSF3)->(DbSkip())
EndDo
DbSelectArea(cAliasSF3)
dbCloseArea()
If nLin >= 60
	nLin := 0
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                          
	FmtLin(,{aL[1]},,,@nLin)
EndIf
FmtLin({Transform(nTotalvlr,"@R 99,999,999.99")},aL[29],"",,@nLin)
FmtLin(,{aL[34]},,,@nLin)                           
Return(.T.)
