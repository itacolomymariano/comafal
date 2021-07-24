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
User Function CRF010H(Cabec1,Cabec2,Titulo,nLin,nMes)
Local aPedMesTrava := {}
Local aPedAntTrava := {}
Local aTitMesTrava := {}
Local aTitAntTrava := {}
Local aPagamentos := {}
LoCAL aTitRec :={}    
LoCAL aInadiplencia := {}     
Local cAliasSC7 := "cAliasSC7"
Local nSC7	 	:= 0
Local aStruSC7 	:= SC7->(dbStruct())
Local cALiasSE2 := "cALiasSE2"
Local nSE2		:= 0
Local aStruSE2  := SE2->(dbStruct())
Local cALiasSE5 := "cALiasSE5"
Local nSE5		:= 0
Local aStruSE5  := SE5->(dbStruct())
Local cALiasSE1 := "cALiasSE2"
Local nSE1		:= 0
Local aStruSE1  := SE2->(dbStruct())
Local cALiasSE7 := "cALiasSE7"
Local aStruSE7  := SE7->(dbStruct())
Local nSE7		:= 0
Local dDataIni := ctod(Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+"01")
Local dDataFim := MV_PAR02 
Local nVal190 := 0
Local nVal290 := 0
Local nVal390 := 0
Local nI := 0

cQuery := " SELECT SUM(E7_VAL"+aMes[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))]+") AS VALOR1 , SUM(E7_VAL"+aMes[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)]+")  AS VALOR2 ,SUM(E7_VAL"+aMes[Month(MV_PAR02)]+ ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATADMMT) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Meta Orcamento Adm....")
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
		FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
		aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],;
		aMeses[nMes],;
		aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
		aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],;
		aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5]},,,@nLin)
	EndIf	
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
		    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
			Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[52],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[53]},,,@nLin)	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATADMMT) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Realizado Orcamento Adm....")
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
			FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
			FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	EndIf	
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	        Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		    Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[54],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[58],aL[59],aL[60]},,,@nLin)	 
nVal190 := 0
nVal290 := 0
nVal390 := 0
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '200135' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Comafal - RS ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[61],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[62]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '200360' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Cabomar ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[63],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[64]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200361','200362','200363','200364','200369','200370','200375') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Maquinas ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[65],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[66]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200313') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Finame...")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[67],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[68]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200132') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Leasing ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[69],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3

	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[70]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200365','200366','200368','200372','200373','200374','200376','200377','200321') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Benfeitorias ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[71],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[72]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200367') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","ISO 9001 - 2000 ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[73],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[74]},,,@nLin)	
FmtLin({Transform(nVal190,"@R 99,999,999.99"),;
	    Transform(nVal290,"@R 99,999,999.99"),;
		Transform(nVal390,"@R 99,999,999.99")},aL[75],"",,@nLin)
FmtLin(,{aL[76],aL[77],aL[78]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200111','200306','200406') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Maquinas - Motores ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[79],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[80]},,,@nLin)	
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('700309') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Importado ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[81],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[82]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200109') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Informatica ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[83],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[84]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200110') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Moveis Utensilios ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[85],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[86],aL[87],aL[88]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100115','500103') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","INSS ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[89],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[90]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100114','100122','100214','100222','100312','100317','100412','100417') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","FGTS ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[91],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[92]},,,@nLin)      
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200126') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","CPMF ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[93],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[94]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200127') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","IPTU ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[95],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[96]},,,@nLin) 
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200208') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","IPI ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[97],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[98]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('200211') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","ICMS ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[99],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[100]},,,@nLin)        
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('700206') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Emprestimos entre Empresas ....")
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea(cAliasSE7)
Do While (cAliasSE7)->(!EOF())
	FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[101],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()		
FmtLin(,{aL[102]},,,@nLin)


Return(.T.)
