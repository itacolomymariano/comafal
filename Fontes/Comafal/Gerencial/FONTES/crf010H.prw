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
User Function CRF010H(Cabec1,Cabec2,Titulo,nLin,nMes)
Local aPedMesTrava := {}
Local aPedAntTrava := {}
Local aTitMesTrava := {}
Local aTitAntTrava := {}
Local aPagamentos := {}
LoCAL aTitRec :={}    
LoCAL aInadiplencia := {}     
Local nSC7	 	:= 0
Local aStruSC7 	:= SC7->(dbStruct())
Local nSE2		:= 0
Local aStruSE2  := SE2->(dbStruct())
Local nSE5		:= 0
Local aStruSE5  := SE5->(dbStruct())
Local nSE1		:= 0
Local aStruSE1  := SE2->(dbStruct())
Local nSE7		:= 0
Local aStruSE7  := SE7->(dbStruct())
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
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
Do While cAliasSE7->(!EOF())
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
	FmtLin({Transform(cAliasSE7->VALOR1,"@R 99,999,999.99"),;
		    Transform(cAliasSE7->VALOR2,"@R 99,999,999.99"),;
			Transform(cAliasSE7->VALOR3,"@R 99,999,999.99")},aL[52],"",,@nLin)
	cAliasSE7->(DbSkip())
EndDo
DbSelectArea("cAliasSE7")
dbCloseArea()
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
FmtLin(,{aL[53]},,,@nLin)	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATADMMT) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
Do While cAliasSE7->(!EOF())
	If nLin >= 60
			Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
			nLin:=PROW()+1                                    
			FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
			FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	EndIf	
	FmtLin({Transform(cAliasSE7->VALOR1,"@R 99,999,999.99"),;
	        Transform(cAliasSE7->VALOR2,"@R 99,999,999.99"),;
		    Transform(cAliasSE7->VALOR3,"@R 99,999,999.99")},aL[54],"",,@nLin)
	cAliasSE7->(DbSkip())
EndDo
DbSelectArea("cAliasSE7")
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
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
Do While cAliasSE7->(!EOF())
	FmtLin({Transform(cAliasSE7->VALOR1,"@R 99,999,999.99"),;
	    	Transform(cAliasSE7->VALOR2,"@R 99,999,999.99"),;
			Transform(cAliasSE7->VALOR3,"@R 99,999,999.99")},aL[61],"",,@nLin)
	nVal190 += (cAliasSE7)->VALOR1
	nVal290 += (cAliasSE7)->VALOR2
	nVal390 += (cAliasSE7)->VALOR3
	cAliasSE7->(DbSkip())
EndDo
DbSelectArea("cAliasSE7")
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
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
Do While cAliasSE7->(!EOF())
	FmtLin({Transform(cAliasSE7->VALOR1,"@R 99,999,999.99"),;
	    	Transform(cAliasSE7->VALOR2,"@R 99,999,999.99"),;
			Transform(cAliasSE7->VALOR3,"@R 99,999,999.99")},aL[63],"",,@nLin)
	nVal190 += cAliasSE7->VALOR1
	nVal290 += cAliasSE7->VALOR2
	nVal390 += cAliasSE7->VALOR3
	cAliasSE7->(DbSkip())
EndDo
DbSelectArea("cAliasSE7")
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
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
Do While cAliasSE7->(!EOF())
	FmtLin({Transform(cAliasSE7->VALOR1,"@R 99,999,999.99"),;
	    	Transform(cAliasSE7->VALOR2,"@R 99,999,999.99"),;
			Transform(cAliasSE7->VALOR3,"@R 99,999,999.99")},aL[65],"",,@nLin)
	nVal190 += cAliasSE7->VALOR1
	nVal290 += cAliasSE7->VALOR2
	nVal390 += cAliasSE7->VALOR3
	cAliasSE7->(DbSkip())
EndDo
DbSelectArea("cAliasSE7")
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

If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf


FmtLin(,{aL[102]},,,@nLin)

If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf


Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³Pedidos   ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Ajusta o SX1                                                º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AFATR001                                                   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Pedidos(aPedido,nQtdAtivo,nMes)

Local aArea := GetArea()
Local cQuery 	:= ""
Local nSC5	 	:= 0 
Local nSC6	 	:= 0 
Local nSC9	 	:= 0
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSC6 	:= SC6->(dbStruct())  
Local aStruSC9 	:= SC9->(dbStruct()) 
Local nSA1	 	:= 0
Local aStruSA1 	:= SA1->(dbStruct()) 
cQuery := "SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES,"
cQuery += "SUM(C6.C6_VALOR) as PedBrt, "
cQuery += "COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += "WHERE "
cQuery += "C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND "
cQuery += "C5.D_E_L_E_T_ = ' ' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +  "' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C5.C5_FILIAL = C6.C6_FILIAL AND "
cQuery += "C6.C6_NUM = C5.C5_NUM AND "
cQuery += "C6.C6_LOCAL IN ('03','04','05') AND "
cQuery += "C6.D_E_L_E_T_ = ' ' "          
cQuery += "GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += "ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
dbSelectArea("cAliasPED")
aPedido := {}
aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0,0,0,0,0,0,0})
aadd(aPedido,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0,0,0,0,0,0,0})
aadd(aPedido,{SUBSTR(DTOS(MV_PAR02),1,6),0,0,0,0,0,0,0,0})
//                                     1 2 3 4 5 6,7 8 9
while cAliasPED->(!eof())
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aPedido[1,2] += cAliasPED->PedBrt
		aPedido[1,3] += cAliasPED->QtdPED
	EndIf
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aPedido[2,2] += cAliasPED->PedBrt
		aPedido[2,3] += cAliasPED->QtdPED
	EndIf    
	If cAliasPED->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aPedido[3,2] += cAliasPED->PedBrt
		aPedido[3,3] += cAliasPED->QtdPED
	EndIf    
	cAliasPED->(dbskip())
enddo
DbSelectArea("cAliasPED")
DbCloseArea()
cQuery := "SELECT SUBSTRING(C9.C9_DATALIB,1,6) AS MES, SUM(C9.C9_PRCVEN*C9.C9_QTDLIB) as PedBrt, COUNT(DISTINCT C9.C9_PEDIDO) as QtdPED FROM "
cQuery += RetSqlName('SC9') + " C9 WHERE C9.C9_FILIAL = '"+xFilial("SC9")+"'  AND "
cQuery += "C9.D_E_L_E_T_ = ' ' AND " 
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) = '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND C9.C9_NFISCAL <> '      ' AND " 
cQuery += "C9.C9_PEDIDO IN (SELECT C5.C5_NUM FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) = '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C5.C5_TIPO = 'N' AND C5.D_E_L_E_T_ = ' ' AND C6.C6_FILIAL = '"+xFilial("SC6")+"' AND C6.C6_NUM = C5.C5_NUM AND "
cQuery += "C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND C6.C6_LOCAL IN ('03','04','05') AND C6.D_E_L_E_T_ = ' ' ) "
cQuery += "GROUP BY SUBSTRING(C9.C9_DATALIB,1,6) "
cQuery += "ORDER BY SUBSTRING(C9.C9_DATALIB,1,6) ASC"
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC9 := 1 To len(aStruSC9)
	If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
		TcSetField("cAliasPED",aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
	EndIf
Next nSC9
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==cAliasPED->MES})
	If nReg > 0
		aPedido[nReg,6] :=cAliasPED->QtdPED
		aPedido[nReg,7] :=cAliasPED->PedBrt
	EndIf
	cAliasPED->(dbskip())
enddo
DbSelectArea("cAliasPED")
DbCloseArea()
cQuery := "SELECT SUBSTRING(C9.C9_DATALIB,1,6) AS MES, SUM(C9.C9_PRCVEN*C9.C9_QTDLIB) as PedBrt, COUNT(DISTINCT C9.C9_PEDIDO) as QtdPED FROM "
cQuery += RetSqlName('SC9') + " C9 WHERE C9.C9_FILIAL = '"+xFilial("SC9")+"' AND "
cQuery += "C9.D_E_L_E_T_ = ' ' AND " 
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' AND C9.C9_NFISCAL <> '      ' AND " 
cQuery += "C9.C9_PEDIDO IN (SELECT C5.C5_NUM FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) = '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' AND "
cQuery += "C5.C5_TIPO = 'N' AND C5.D_E_L_E_T_ = ' ' AND C6.C6_FILIAL = '"+xFilial("SC6")+"' AND C6.C6_NUM = C5.C5_NUM AND "
cQuery += "C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND C6.C6_LOCAL IN ('03','04','05') AND C6.D_E_L_E_T_ = ' ' ) "
cQuery += "GROUP BY SUBSTRING(C9.C9_DATALIB,1,6) "
cQuery += "ORDER BY SUBSTRING(C9.C9_DATALIB,1,6) ASC"
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC9 := 1 To len(aStruSC9)
	If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
		TcSetField("cAliasPED",aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
	EndIf
Next nSC9
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==cAliasPED->MES})
	If nReg > 0
		aPedido[nReg,6] :=cAliasPED->QtdPED
		aPedido[nReg,7] :=cAliasPED->PedBrt
	EndIf
	cAliasPED->(dbskip())
enddo
DbSelectArea("cAliasPED")
DbCloseArea() 
cQuery := "SELECT SUBSTRING(C9.C9_DATALIB,1,6) AS MES, SUM(C9.C9_PRCVEN*C9.C9_QTDLIB) as PedBrt, COUNT(DISTINCT C9.C9_PEDIDO) as QtdPED FROM "
cQuery += RetSqlName('SC9') + " C9 WHERE C9.C9_FILIAL = '"+xFilial("SC9")+"' AND "
cQuery += "C9.D_E_L_E_T_ = ' ' AND " 
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' AND C9.C9_NFISCAL <> '      ' AND " 
cQuery += "C9.C9_PEDIDO IN (SELECT C5.C5_NUM FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) = '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' AND "
cQuery += "C5.C5_TIPO = 'N' AND C5.D_E_L_E_T_ = ' ' AND C6.C6_FILIAL = '"+xFilial("SC6")+"' AND C6.C6_NUM = C5.C5_NUM AND "
cQuery += "C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND C6.C6_LOCAL IN ('03','04','05') AND C6.D_E_L_E_T_ = ' ' ) "
cQuery += "GROUP BY SUBSTRING(C9.C9_DATALIB,1,6) "
cQuery += "ORDER BY SUBSTRING(C9.C9_DATALIB,1,6) ASC"
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC9 := 1 To len(aStruSC9)
	If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
		TcSetField("cAliasPED",aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
	EndIf
Next nSC9
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==cAliasPED->MES})
	If nReg > 0
		aPedido[nReg,6] :=cAliasPED->QtdPED
		aPedido[nReg,7] :=cAliasPED->PedBrt
	EndIf
	cAliasPED->(dbskip())
enddo
DbSelectArea("cAliasPED")
DbCloseArea()
cQuery := "SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES,"
cQuery += "COUNT(DISTINCT C5.C5_CLIENTE) as CLIENTE "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += "WHERE "
cQuery += "C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND "
cQuery += "C5.D_E_L_E_T_ = ' ' AND "
cQuery += "C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += "C6.C6_NUM = C5.C5_NUM AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C6.C6_LOCAL IN ('03','04','05') AND "
cQuery += "C6.D_E_L_E_T_ = ' ' "          
cQuery += "GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += "ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==cAliasPED->MES})
	If nReg > 0
		aPedido[nReg,04] := cAliasPED->CLIENTE
	EndIf	
	cAliasPED->(dbskip())
enddo
DbSelectArea("cAliasPED")
DbCloseArea()
cQuery := " SELECT SUBSTRING(A1_DTCAD,1,6) AS NOVOS,COUNT(*) AS QTD "
cQuery += " FROM "
cQuery += RetSqlName('SA1')
cQuery += " WHERE "
cQuery += " A1_FILIAL = '"+xFilial("SA1")+"' AND "
cQuery += " D_E_L_E_T_ = ' ' AND "                             
cQuery += "SUBSTRING(A1_DTCAD,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND "
cQuery += "SUBSTRING(A1_DTCAD,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += "GROUP BY SUBSTRING(A1_DTCAD,1,6) "
cQuery += "ORDER BY SUBSTRING(A1_DTCAD,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSA1",.T.,.T.)
For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasSA1",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1
dbSelectArea("cAliasSA1")
while cAliasSA1->(!eof())
	nReg := aScan(aPedido,{|x| x[1]==cAliasSA1->NOVOS})
	If nReg > 0
		aPedido[nReg,05] := cAliasSA1->QTD
	EndIf	
	cAliasSA1->(dbskip())
enddo

DbSelectArea("cAliasSA1")
DbCloseArea()
Restarea(aArea)
Return(.t.)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatFin    ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatFin                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function FatFin(aFatFin,nMes)   

Local aArea		:= GetArea()
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local cQuery 	:= ""	
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local nSB1		:= 0
Local aStruSB1  := SB1->(dbStruct())

//aL[23]	:=		"|Faturamento que nao entraram no Financeiro                                                                                                                                                            |"
//aL[24]	:=		"|Cliente| Loja | Nome                                   |Nota Fiscal | Serie | Produto       | Descricao                                         | QTD           | Armazem | Valor                     |"		
//aL[25]	:=		"|###### | ##   | ###################################### | ######     |   ### | ##############| ################################################# |###############|   ##    | ###############           |"

cQuery := " SELECT D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL,SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR "
cQuery += " FROM "         
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SA1') + " A1, "
cQuery += RetSqlName('SB1') + " B1 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04','05') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4) + StrZero(nMes,2) + "'  "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_DOC+D2.D2_SERIE NOT IN (SELECT E1_NUM+E1_PREFIXO "
cQuery += " FROM "
cQuery += RetSqlName('SE1')
cQuery += " WHERE "
cQuery += " D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460') "
cQuery += " AND D2.D2_CLIENTE+D2.D2_LOJA = A1.A1_COD+A1.A1_LOJA "
cQuery += " AND A1.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_COD = B1.B1_COD " 
cQuery += " AND B1.B1_FILIAL = '"+xFilial("SB1")+"' "
cQuery += " AND B1.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery += " ORDER BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("cAliasSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasSD2",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

For nSB1 := 1 To len(aStruSB1)
	If aStruSB1[nSB1][2] <> "C" .And. FieldPos(aStruSB1[nSB1][1])<>0
		TcSetField("cAliasSD2",aStruSB1[nSB1][1],aStruSB1[nSB1][2],aStruSB1[nSB1][3],aStruSB1[nSB1][4])
	EndIf
Next nSB1
	
DbSelectArea("cAliasSD2")
Do While cAliasSD2->(!EOF())
	aadd(aFatFin,{cAliasSD2->D2_CLIENTE,cAliasSD2->D2_LOJA,cAliasSD2->A1_NREDUZ,cAliasSD2->D2_DOC,cAliasSD2->D2_SERIE,;
	cAliasSD2->D2_COD,cAliasSD2->B1_DESC,cAliasSD2->D2_QUANT,cAliasSD2->D2_LOCAL,cAliasSD2->VALOR})
	cAliasSD2->(DbSkip())
EndDo
DbSelectArea("cAliasSD2")
dbCloseArea()

RestArea(aArea)


Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatEst    ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatEst                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function FatEst(aFatEst,nMes)   

Local aArea		:= GetArea()
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local cQuery 	:= ""	
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local nSB1		:= 0
Local aStruSB1  := SB1->(dbStruct())

cQuery := " SELECT D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL,SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR "
cQuery += " FROM "         
cQuery += RetSqlName('SD2') + " D2, "
cQuery += RetSqlName('SA1') + " A1, "
cQuery += RetSqlName('SB1') + " B1 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04','05') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "'  "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_TES IN (SELECT F4_CODIGO FROM "
cQuery +=  RetSqlName('SF4') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND F4_TIPO = 'S' AND F4_ESTOQUE = 'N') "
cQuery += " AND D2.D2_CLIENTE+D2.D2_LOJA = A1.A1_COD+A1.A1_LOJA "
cQuery += " AND A1.D_E_L_E_T_ = ' ' "
cQuery += " AND D2.D2_COD = B1.B1_COD " 
cQuery += " AND B1.B1_FILIAL = '"+xFilial("SB1")+"' "
cQuery += " AND B1.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery += " ORDER BY D2_CLIENTE,D2_LOJA,A1_NREDUZ,D2_DOC,D2_SERIE,D2_COD,B1_DESC,D2_QUANT,D2_LOCAL "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("cAliasSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2

For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasSD2",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1

For nSB1 := 1 To len(aStruSB1)
	If aStruSB1[nSB1][2] <> "C" .And. FieldPos(aStruSB1[nSB1][1])<>0
		TcSetField("cAliasSD2",aStruSB1[nSB1][1],aStruSB1[nSB1][2],aStruSB1[nSB1][3],aStruSB1[nSB1][4])
	EndIf
Next nSB1
DbSelectArea("cAliasSD2")
Do While cAliasSD2->(!EOF())
	aadd(aFatEst,{cAliasSD2->D2_CLIENTE,cAliasSD2->D2_LOJA,cAliasSD2->A1_NREDUZ,cAliasSD2->D2_DOC,cAliasSD2->D2_SERIE,;
	cAliasSD2->D2_COD,cAliasSD2->B1_DESC,cAliasSD2->D2_QUANT,cAliasSD2->D2_LOCAL,cAliasSD2->VALOR})
	cAliasSD2->(DbSkip())
EndDo        
DbSelectArea("cAliasSD2")
dbCloseArea()
RestArea(aArea)
Return(.T.)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³Inativos   ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07  º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³FatEst                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Inativos(cData)   
Local aArea := GetArea()
Local nQtdDest  := 0
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local cQuery    := "" 
cQuery := " SELECT COUNT(*) AS GERAL "
cQuery += " FROM "
cQuery +=   RetSqlName('SA1') + " A1  "
cQuery += " WHERE "
cQuery += " A1.A1_FILIAL = '"+xFilial("SA1")+"' AND "  
cQuery += " A1_ULTCOM < '"+ cData +"' AND "
cQuery += " A1.D_E_L_E_T_ = ' '  "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSA1",.T.,.T.)
For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasSA1",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1
dbSelectArea("cAliasSA1")
while cAliasSA1->(!eof())
	nQtdDest := cAliasSA1->GERAL
	cAliasSA1->(dbskip())
enddo
DbSelectArea("cAliasSA1")
DbCloseArea()
RestArea(aArea)
Return(nQtdDest)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³Ativos    ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Ativos                                                      º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function Ativos(cDataIni,cDtFim)   

Local aArea := GetArea()
Local nQtdAtivo := 0
Local nSA1		:= 0
Local aStruSA1  := SA1->(dbStruct())
Local cQuery    := "" 
cQuery := " SELECT COUNT(*) AS GERAL "
cQuery += " FROM "
cQuery +=   RetSqlName('SA1') + " A1  "
cQuery += " WHERE "                                    
cQuery += " A1.A1_FILIAL = '"+xFilial("SA1")+"' AND "  
cQuery += " SUBSTRING(A1_ULTCOM,1,6) >= '"+ cDataIni +"' AND " 
//cQuery += " SUBSTRING(A1_ULTCOM,1,6) <= '" + cDtFim + "' AND "
//cQuery += " SUBSTRING(A1_DTCAD,1,6) < '"+ cDtFim +"' AND "
cQuery += " A1.D_E_L_E_T_ = ' '  "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSA1",.T.,.T.)
For nSA1 := 1 To len(aStruSA1)
	If aStruSA1[nSA1][2] <> "C" .And. FieldPos(aStruSA1[nSA1][1])<>0
		TcSetField("cAliasSA1",aStruSA1[nSA1][1],aStruSA1[nSA1][2],aStruSA1[nSA1][3],aStruSA1[nSA1][4])
	EndIf
Next nSA1
dbSelectArea("cAliasSA1")
while cAliasSA1->(!eof())
	nQtdAtivo := cAliasSA1->GERAL
	cAliasSA1->(dbskip())
enddo
DbSelectArea("cAliasSA1")
DbCloseArea()

RestArea(aArea)

Return(nQtdAtivo)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatLF     ºAutor  ³Eduardo Zanardo     º Data ³  19/03/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Faturamento de vendas livro fiscal                          º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function FatLF(nMes)

Local aArea		:= GetArea()
Local nSF3		:= 0
Local aStruSF3  := SF3->(dbStruct())
Local cQuery 	:= ""	
Local nValor	:= 0 

cQuery := "SELECT SUBSTRING(F3_ENTRADA,1,6) AS EMISSAO ,SUM(F3_VALCONT) AS VALOR FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE SUBSTRING(F3_ENTRADA,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND F3_CFO IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND F3_OBSERV <> 'NF CANCELADA' "
cQuery += " AND D_E_L_E_T_ <> '*' " 
cQuery += " GROUP BY SUBSTRING(F3_ENTRADA,1,6) "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField(cAliasSF3,aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("cAliasSF3")
Do While cAliasSF3->(!EOF())
	nValor:= cAliasSF3->VALOR
	cAliasSF3->(DbSkip())
EndDo
DbSelectArea("cAliasSF3")
dbCloseArea()

RestArea(aArea)

Return(nValor)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatLFDIF  ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Faturamento de vendas livro fiscal                          º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function FatLFDIF(aFatLFDIF,nMes)
Local aArea		:= GetArea()
Local nSF3		:= 0
Local aStruSF3  := SF3->(dbStruct())
Local cQuery 	:= ""	
cQuery := " SELECT F3_NFISCAL,F3_SERIE,F3_ENTRADA,F3_VALCONT,F3_OBSERV FROM "
cQuery += RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " F3_FILIAL = '"+xFilial("SF3")+"' "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4)+ StrZero(nMes,2) + "' "
cQuery += " AND F3_ESPECIE = 'NF' "
cQuery += " AND F3_CFO IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND D_E_L_E_T_ <> '*' "
cQuery += " AND F3_NFISCAL+F3_SERIE NOT IN (SELECT D2_DOC+D2_SERIE FROM "
cQuery += RetSqlName('SD2')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND D2_TIPO = 'N' "
cQuery += " AND D2_CF IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND D2_LOCAL IN ('03','04','05') "
cQuery += " AND SUBSTRING(D2_EMISSAO,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4) + StrZero(nMes,2) + "' "
cQuery += " AND D2_FILIAL = '"+xFilial("SD2")+"' )"
cQuery += " ORDER BY F3_NFISCAL,F3_SERIE,F3_ENTRADA ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSF3",.T.,.T.)
For nSF3 := 1 To len(aStruSF3)
	If aStruSF3[nSF3][2] <> "C" .And. FieldPos(aStruSF3[nSF3][1])<>0
		TcSetField("cAliasSF3",aStruSF3[nSF3][1],aStruSF3[nSF3][2],aStruSF3[nSF3][3],aStruSF3[nSF3][4])
	EndIf
Next nSF3
DbSelectArea("cAliasSF3")
If cAliasSF3->(!EOF()) 	                                                    
	Do While cAliasSF3->(!EOF()) 
		aadd(aFatLFDIF,{cAliasSF3->F3_NFISCAL,cAliasSF3->F3_SERIE,cAliasSF3->F3_OBSERV,cAliasSF3->F3_VALCONT})
		cAliasSF3->(DbSkip())
	EndDo
EndIf	         
DbSelectArea("cAliasSF3")
dbCloseArea()

RestArea(aArea)

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatFTDIF  ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Faturamento de vendas livro fiscal                          º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ crf010                                                     º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function FatFTDIF(aFatFTDIF,nMes)
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local nSF2		:= 0
Local aStruSF2  := SF2->(dbStruct())
Local cQuery 	:= ""	
Local aArea		:= GetArea()

cQuery := " SELECT F2_DOC,F2_SERIE,F2_EMISSAO,F2_VALFAT AS TOTAL FROM "
cQuery += RetSqlName('SF2') + " F2 "
cQuery += " WHERE F2.D_E_L_E_T_ <> '*' AND "
cQuery += " F2_FILIAL = '"+xFilial("SF2")+"' AND "
cQuery += " SUBSTRING(F2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' AND "
cQuery += " F2_TIPO = 'N' AND "
cQuery += " F2.F2_DOC+F2.F2_SERIE+F2.F2_EMISSAO IN (SELECT D2.D2_DOC+D2.D2_SERIE+D2.D2_EMISSAO FROM "
cQuery +=   RetSqlName('SD2')  + " D2 WHERE "
cQuery += " D2.D_E_L_E_T_ <> '*' AND "
cQuery += " D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " SUBSTRING(D2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' AND "
cQuery += " D2_TIPO = 'N' AND "
cQuery += " D2_LOCAL IN ('03','04','05') AND "
cQuery += " D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " D2_CF IN (" + Alltrim(cCFOPVenda) + ")) AND "
cQuery += " F2_DOC+F2_SERIE NOT IN (SELECT F3_NFISCAL+F3_SERIE FROM "
cQuery +=   RetSqlName('SF3')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND F3_CFO IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND SUBSTRING(F3_ENTRADA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND F3_FILIAL = '"+xFilial("SF3")+"' )" 
cQuery += " AND F2_DOC+F2_SERIE NOT IN (SELECT E1_NUM+E1_PREFIXO FROM "
cQuery +=   RetSqlName('SE1')
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_TIPO = 'NF' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(nMes,2) + "' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' ) "
cQuery += " ORDER BY F2_DOC,F2_SERIE,F2_EMISSAO ASC "
	
cQuery:=ChangeQuery(cQuery)
	
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSD2",.T.,.T.)

For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("cAliasSD2",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea("cAliasSD2")
Do While cAliasSD2->(!EOF())
	aadd(aFatFTDIF,{cAliasSD2->F2_DOC,cAliasSD2->F2_SERIE,cAliasSD2->TOTAL})
	cAliasSD2->(DbSkip())
EndDo
DbSelectArea("cAliasSD2")
dbCloseArea()

RestArea(aArea)

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºFunction  ³FatOutros ºAutor  ³Eduardo Zanardo     º Data ³  24/03/07   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Layout do relatorio                                         º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ RFin003                                                    º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function FatOutros(aFatOutros,nMes)             

Local aArea		:= GetArea()
Local nSD2		:= 0
Local aStruSD2  := SD2->(dbStruct())
Local cQuery 	:= ""	

cQuery := " SELECT D2.D2_EMISSAO,D2.D2_DOC,D2.D2_SERIE,D2.D2_COD,D2.D2_QUANT,D2.D2_LOCAL, " 
cQuery += " (D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS TOTAL "
cQuery += " FROM "
cQuery +=   RetSqlName('SD2') + " D2 "
cQuery += " WHERE "
cQuery += " D2.D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " SUBSTRING(D2_EMISSAO,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4) + StrZero(nMes,2) + "' AND "
cQuery += " D2.D2_TIPO = 'N' AND "
cQuery += " D2.D2_LOCAL NOT IN ('03','04','05') AND "
cQuery += " D2.D2_FILIAL = '"+xFilial("SD2")+"' AND "
cQuery += " D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " D2.D_E_L_E_T_ = ' ' "
cQuery += " ORDER BY D2.D2_EMISSAO,D2.D2_DOC,D2.D2_SERIE,D2.D2_COD,D2.D2_QUANT,D2.D2_LOCAL ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("cAliasSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
DbSelectArea("cAliasSD2")
Do While cAliasSD2->(!EOF())
	aadd(aFatOutros,{cAliasSD2->D2_DOC,cAliasSD2->D2_SERIE,cAliasSD2->D2_COD,Substr(Alltrim(Posicione("SB1",1,xFilial("SB1")+cAliasSD2->D2_COD,"B1_DESC")),1,20),;
					 cAliasSD2->D2_QUANT,cAliasSD2->D2_LOCAL,cAliasSD2->TOTAL})
	cAliasSD2->(DbSkip())
EndDo                    
DbSelectArea("cAliasSD2")
dbCloseArea()

RestArea(aArea)

Return(.T.)
