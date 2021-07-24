#INCLUDE "PROTHEUS.CH"
#INCLUDE "topconn.ch"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CRF010  ºAutor  ³Eduardo Zanardo     º Data ³  03/31/03   º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ RESUMO GERENCIAL                   º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function CRF010D(nLin,nMes,Cabec1,Cabec2,Titulo)

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
Local aMes1 := {}
Local aMes2 := {}
Local aMes3 := {}

U_RFinLay4()
Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
nLin:=PROW()+1                                    
FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
FmtLin(,{aL[5]},,,@nLin)
cQuery := " SELECT SUBSTRING(C7_EMISSAO,5,2) AS EMISSAO, SUM(C7_TOTAL) AS VALOR FROM "
cQuery += RetSqlName("SC7")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND C7_FILIAL = '"+xFilial("SC7")+"' "
cQuery += " AND C7_X_STAT = '1' "
cQuery += " AND C7_EMISSAO >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"01' " 
cQuery += " AND C7_EMISSAO <= '" + DTOS(MV_PAR02) + "' "
cQuery += " GROUP BY SUBSTRING(C7_EMISSAO,5,2) " 
cQuery += " ORDER BY SUBSTRING(C7_EMISSAO,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC7,.T.,.T.)
For nSC7 := 1 To len(aStruSC7)
	If aStruSC7[nSC7][2] <> "C" .And. FieldPos(aStruSC7[nSC7][1])<>0
		TcSetField(cAliasSC7,aStruSC7[nSC7][1],aStruSC7[nSC7][2],aStruSC7[nSC7][3],aStruSC7[nSC7][4])
	EndIf
Next nSC7
While (cAliasSC7)->(!Eof())
	aadd(aPedMesTrava,{(cAliasSC7)->EMISSAO,(cAliasSC7)->VALOR})
	(cAliasSC7)->(DbSkip())
EndDo
DbSelectArea(cAliasSC7)
(cAliasSC7)->(DbCloseArea())
If Len(aPedMesTrava) < 3
	For nI := 1 to 3-Len(aPedMesTrava)
		aadd(aPedMesTrava,{"00",0})	
	Next nI
EndIf	
FmtLin({Transform(aPedMesTrava[1,2],"@R 99,999,999.99"),;
	    Transform(aPedMesTrava[2,2],"@R 99,999,999.99"),;
		Transform(aPedMesTrava[3,2],"@R 99,999,999.99")},aL[08],"",,@nLin)
FmtLin(,{aL[9]},,,@nLin)
cQuery := " SELECT SUBSTRING(C7_EMISSAO,5,2) AS EMISSAO, SUM(C7_TOTAL) AS VALOR FROM "
cQuery += RetSqlName("SC7")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND C7_FILIAL = '"+xFilial("SC7")+"' "
cQuery += " AND C7_X_STAT = '1' "
cQuery += " AND SUBSTRING(C7_EMISSAO,1,6) <= '"  + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' "
cQuery += " GROUP BY SUBSTRING(C7_EMISSAO,5,2) " 
cQuery += " ORDER BY SUBSTRING(C7_EMISSAO,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC7,.T.,.T.)
For nSC7 := 1 To len(aStruSC7)
	If aStruSC7[nSC7][2] <> "C" .And. FieldPos(aStruSC7[nSC7][1])<>0
		TcSetField(cAliasSC7,aStruSC7[nSC7][1],aStruSC7[nSC7][2],aStruSC7[nSC7][3],aStruSC7[nSC7][4])
	EndIf
Next nSC7
If (cAliasSC7)->(!Eof())
	While (cAliasSC7)->(!Eof())
		aadd(aPedAntTrava,{(cAliasSC7)->EMISSAO,(cAliasSC7)->VALOR})
		(cAliasSC7)->(DbSkip())
	EndDo
Else
	aadd(aPedAntTrava,{"00",0})
EndIf
DbSelectArea(cAliasSC7)
(cAliasSC7)->(DbCloseArea())

cQuery := " SELECT SUBSTRING(C7_EMISSAO,5,2) AS EMISSAO, SUM(C7_TOTAL) AS VALOR FROM "
cQuery += RetSqlName("SC7")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND C7_FILIAL = '"+xFilial("SC7")+"' "
cQuery += " AND C7_X_STAT = '1' "
cQuery += " AND SUBSTRING(C7_EMISSAO,1,6) <= '"  + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' "
cQuery += " GROUP BY SUBSTRING(C7_EMISSAO,5,2) " 
cQuery += " ORDER BY SUBSTRING(C7_EMISSAO,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC7,.T.,.T.)
For nSC7 := 1 To len(aStruSC7)
	If aStruSC7[nSC7][2] <> "C" .And. FieldPos(aStruSC7[nSC7][1])<>0
		TcSetField(cAliasSC7,aStruSC7[nSC7][1],aStruSC7[nSC7][2],aStruSC7[nSC7][3],aStruSC7[nSC7][4])
	EndIf
Next nSC7
If (cAliasSC7)->(!Eof())
	While (cAliasSC7)->(!Eof())
		aadd(aPedAntTrava,{(cAliasSC7)->EMISSAO,(cAliasSC7)->VALOR})
		(cAliasSC7)->(DbSkip())
	EndDo
Else
	aadd(aPedAntTrava,{"00",0})
EndIf
DbSelectArea(cAliasSC7)
(cAliasSC7)->(DbCloseArea())
cQuery := " SELECT SUBSTRING(C7_EMISSAO,1,6) AS EMISSAO, SUM(C7_TOTAL) AS VALOR FROM "
cQuery += RetSqlName("SC7")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND C7_FILIAL = '"+xFilial("SC7")+"' "
cQuery += " AND C7_X_STAT = '1' "
cQuery += " AND SUBSTRING(C7_EMISSAO,1,6) <= '"  + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2) + "' "
cQuery += " GROUP BY SUBSTRING(C7_EMISSAO,1,6) " 
cQuery += " ORDER BY SUBSTRING(C7_EMISSAO,1,6) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC7,.T.,.T.)
For nSC7 := 1 To len(aStruSC7)
	If aStruSC7[nSC7][2] <> "C" .And. FieldPos(aStruSC7[nSC7][1])<>0
		TcSetField(cAliasSC7,aStruSC7[nSC7][1],aStruSC7[nSC7][2],aStruSC7[nSC7][3],aStruSC7[nSC7][4])
	EndIf
Next nSC7                          
aadd(aPedAntTrava,{Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),0})
aadd(aPedAntTrava,{Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),0})
aadd(aPedAntTrava,{nMes,0})
While (cAliasSC7)->(!Eof())
	If Val(Substr((cAliasSC7)->EMISSAO,5,2)) <= Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
		aPedAntTrava[1,2]+=(cAliasSC7)->VALOR
	EndIf	
	If Val(Substr((cAliasSC7)->EMISSAO,5,2)) <= Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)
		aPedAntTrava[2,2]+=(cAliasSC7)->VALOR
	EndIf	
	If Val(Substr((cAliasSC7)->EMISSAO,5,2)) <= nMes
		aPedAntTrava[3,2]+=(cAliasSC7)->VALOR
	EndIf
	(cAliasSC7)->(DbSkip())
EndDo
DbSelectArea(cAliasSC7)
(cAliasSC7)->(DbCloseArea())
FmtLin({Transform(aPedAntTrava[1,2],"@R 99,999,999.99"),;
	    Transform(aPedAntTrava[2,2],"@R 99,999,999.99"),;
		Transform(aPedAntTrava[3,2],"@R 99,999,999.99")},aL[10],"",,@nLin)
FmtLin(,{aL[11]},,,@nLin)
cQuery := " SELECT SUBSTRING(E2_VENCREA,5,2) AS EMISSAO,SUM(E2_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE2")
cQuery += " WHERE D_E_L_E_T_ <> '*' "                                            
cQuery += " AND SUBSTRING(E2_VENCREA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E2_VENCREA,1,6) <= '" + substr(dtos(MV_PAR02),1,6)+ "' "
cQuery += " AND E2_FILIAL = '"+xFilial("SE2")+"' "
cQuery += " AND E2_X_STAT = '1' "
cQuery += " GROUP BY SUBSTRING(E2_VENCREA,5,2) " 
cQuery += " ORDER BY SUBSTRING(E2_VENCREA,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE2,.T.,.T.)
For nSE2 := 1 To len(aStruSE2)
	If aStruSE2[nSE2][2] <> "C" .And. FieldPos(aStruSE2[nSE2][1])<>0
		TcSetField(cAliasSE2,aStruSE2[nSE2][1],aStruSE2[nSE2][2],aStruSE2[nSE2][3],aStruSE2[nSE2][4])
	EndIf
Next nSE2
While (cAliasSE2)->(!Eof())  
	aadd(aTitMesTrava,{(cAliasSE2)->EMISSAO,(cAliasSE2)->VALOR})
	(cAliasSE2)->(DbSkip())
EndDo
DbSelectArea(cAliasSE2)
(cAliasSE2)->(DbCloseArea())
If Len(aTitMesTrava) < 3
	For nI := 1 to 3-Len(aTitMesTrava)
		aadd(aTitMesTrava,{"00",0})	
	Next nI
EndIf	             
FmtLin({Transform(aTitMesTrava[1,2],"@R 99,999,999.99"),;
	    Transform(aTitMesTrava[2,2],"@R 99,999,999.99"),;
		Transform(aTitMesTrava[3,2],"@R 99,999,999.99")},aL[12],"",,@nLin)
FmtLin(,{aL[13]},,,@nLin)
cQuery := " SELECT SUBSTRING(E2_VENCREA,5,2) AS EMISSAO,SUM(E2_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE2")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND SUBSTRING(E2_VENCREA,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' "
cQuery += " AND E2_FILIAL = '"+xFilial("SE2")+"' "
cQuery += " AND E2_X_STAT = '1' "
cQuery += " GROUP BY SUBSTRING(E2_VENCREA,5,2) " 
cQuery += " ORDER BY SUBSTRING(E2_VENCREA,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE2,.T.,.T.)
For nSE2 := 1 To len(aStruSE2)
	If aStruSE2[nSE2][2] <> "C" .And. FieldPos(aStruSE2[nSE2][1])<>0
		TcSetField(cAliasSE2,aStruSE2[nSE2][1],aStruSE2[nSE2][2],aStruSE2[nSE2][3],aStruSE2[nSE2][4])
	EndIf
Next nSE2
If (cAliasSE2)->(!Eof())
	While (cAliasSE2)->(!Eof())
			aadd(aTitAntTrava,{(cAliasSE2)->EMISSAO,(cAliasSE2)->VALOR})
		(cAliasSE2)->(DbSkip())
	EndDo
Else
	aadd(aTitAntTrava,{"00",0})
Endif
DbSelectArea(cAliasSE2)
(cAliasSE2)->(DbCloseArea())
cQuery := " SELECT SUBSTRING(E2_VENCREA,5,2) AS EMISSAO,SUM(E2_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE2")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND SUBSTRING(E2_VENCREA,1,6) <= '"  + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' "
cQuery += " AND E2_FILIAL = '"+xFilial("SE2")+"' "
cQuery += " AND E2_X_STAT = '1' "
cQuery += " GROUP BY SUBSTRING(E2_VENCREA,5,2) " 
cQuery += " ORDER BY SUBSTRING(E2_VENCREA,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE2,.T.,.T.)
For nSE2 := 1 To len(aStruSE2)
	If aStruSE2[nSE2][2] <> "C" .And. FieldPos(aStruSE2[nSE2][1])<>0
		TcSetField(cAliasSE2,aStruSE2[nSE2][1],aStruSE2[nSE2][2],aStruSE2[nSE2][3],aStruSE2[nSE2][4])
	EndIf
Next nSE2
If (cAliasSE2)->(!Eof())
	While (cAliasSE2)->(!Eof())
			aadd(aTitAntTrava,{(cAliasSE2)->EMISSAO,(cAliasSE2)->VALOR})
		(cAliasSE2)->(DbSkip())
	EndDo
Else
	aadd(aTitAntTrava,{"00",0})
Endif
DbSelectArea(cAliasSE2)
(cAliasSE2)->(DbCloseArea())
cQuery := " SELECT SUBSTRING(E2_VENCREA,5,2) AS EMISSAO,SUM(E2_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE2")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND SUBSTRING(E2_VENCREA,1,6) <= '"  + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2) + "' "
cQuery += " AND E2_FILIAL = '"+xFilial("SE2")+"' "
cQuery += " AND E2_X_STAT = '1' "
cQuery += " GROUP BY SUBSTRING(E2_VENCREA,5,2) " 
cQuery += " ORDER BY SUBSTRING(E2_VENCREA,5,2) ASC" 
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE2,.T.,.T.)
For nSE2 := 1 To len(aStruSE2)
	If aStruSE2[nSE2][2] <> "C" .And. FieldPos(aStruSE2[nSE2][1])<>0
		TcSetField(cAliasSE2,aStruSE2[nSE2][1],aStruSE2[nSE2][2],aStruSE2[nSE2][3],aStruSE2[nSE2][4])
	EndIf
Next nSE2
If (cAliasSE2)->(!Eof())
	While (cAliasSE2)->(!Eof())
			aadd(aTitAntTrava,{(cAliasSE2)->EMISSAO,(cAliasSE2)->VALOR})
		(cAliasSE2)->(DbSkip())
	EndDo
Else
	aadd(aTitAntTrava,{"00",0})
Endif
DbSelectArea(cAliasSE2)
(cAliasSE2)->(DbCloseArea())
FmtLin({Transform(aTitAntTrava[1,2],"@R 99,999,999.99"),;
	    Transform(aTitAntTrava[2,2],"@R 99,999,999.99"),;
		Transform(aTitAntTrava[3,2],"@R 99,999,999.99")},aL[14],"",,@nLin)
FmtLin(,{aL[15],aL[16],aL[17]},,,@nLin) 
cQuery := "SELECT E5_MOTBX,SUBSTRING(E5_DATA,5,2) AS EMISSAO ,SUM(E5_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE5")
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E5_FILIAL = '"+xFilial("SE5")+"'  "                                
cQuery += " AND SUBSTRING(E5_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E5_DATA,1,6) <= '" + substr(dtos(MV_PAR02),1,6)+ "' "
cQuery += " AND E5_RECPAG = 'P' "
cQuery += " GROUP BY E5_MOTBX,SUBSTRING(E5_DATA,5,2) "
cQuery += " ORDER BY E5_MOTBX,SUBSTRING(E5_DATA,5,2) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE5,.T.,.T.)},"Aguarde...","Processando Pagamentos...")
For nSE5 := 1 To len(aStruSE5)
	If aStruSE5[nSE5][2] <> "C" .And. FieldPos(aStruSE5[nSE5][1])<>0
		TcSetField(cAliasSE5,aStruSE5[nSE5][1],aStruSE5[nSE5][2],aStruSE5[nSE5][3],aStruSE5[nSE5][4])
	EndIf
Next nSE5
dbSelectArea(cAliasSE5)
while (cAliasSE5)->(!eof())
//IIf((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO")))))
//Iif(,,)
//Iif((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO"))))))
	nReg := aScan(aPagamentos,{|x| x[2]==Iif((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO")))))) })
	If nReg >0 
		If Val((cAliasSE5)->EMISSAO) == Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
			aPagamentos[nReg,3] := (cAliasSE5)->VALOR
		ElseIf Val((cAliasSE5)->EMISSAO) == Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)			
			aPagamentos[nReg,4] := (cAliasSE5)->VALOR
		ElseIf Val((cAliasSE5)->EMISSAO) == nMes
			aPagamentos[nReg,5] := (cAliasSE5)->VALOR
		EndIf	
	Else
		If Val((cAliasSE5)->EMISSAO) == Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,Iif((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO")))))),(cAliasSE5)->VALOR,0,0}) 
		ElseIf Val((cAliasSE5)->EMISSAO) == Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)			
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,Iif((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO")))))),0,(cAliasSE5)->VALOR,0}) 
		ElseIf Val((cAliasSE5)->EMISSAO) == nMes
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,Iif((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="VEN","VENDOR",Iif((cAliasSE5)->E5_MOTBX=="DEB","DEBITO CC",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO","COMPENSACAO")))))),0,0,(cAliasSE5)->VALOR}) 		
		EndIf	
	Endif
	(cAliasSE5)->(DbSkip())
EndDo
DbSelectArea(cAliasSE5)
(cAliasSE5)->(DbCloseArea())
For nI := 1 To len(aPagamentos)
	If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
		FmtLin(,{aL[5]},,,@nLin)
	EndIf	
	FmtLin({aPagamentos[nI,2],;    
			Transform(aPagamentos[nI,3],"@R 99,999,999.99"),;
			Transform(aPagamentos[nI,4],"@R 99,999,999.99"),;
			Transform(aPagamentos[nI,5],"@R 99,999,999.99")},aL[18],"",,@nLin)
Next nI
FmtLin(,{aL[19]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf	
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery +=   RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' AND SUBSTRING(E1_VENCREA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=   RetSqlName('SD2') + " D2 WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') AND D2.D2_FILIAL = '"+xFilial("SD2")+"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
DbSelectArea(cAliasSE1)
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 	 
aadd(aInadiplencia,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0}) 	 
While (cAliasSE1)->(!Eof())
	If (cAliasSE1)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aTitRec[1,2]  := (cAliasSE1)->VALOR
		aInadiplencia[1,2]  := (cAliasSE1)->VALOR
	Endif	
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea()) 
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery +=   RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' AND SUBSTRING(E1_VENCREA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+ "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2  WHERE D2.D2_TIPO = 'N' " 
cQuery += " AND D2.D2_LOCAL IN ('03','04') AND D2.D2_FILIAL = '"+xFilial("SD2")+"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
DbSelectArea(cAliasSE1)
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0}) 	
aadd(aInadiplencia,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 	
While (cAliasSE1)->(!Eof())
	If (cAliasSE1)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aTitRec[2,2]  := (cAliasSE1)->VALOR
		aInadiplencia[2,2]  := (cAliasSE1)->VALOR
	EndIf	
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery +=   RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' AND SUBSTRING(E1_VENCREA,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2)+ "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2  WHERE D2.D2_TIPO = 'N' " 
cQuery += " AND D2.D2_LOCAL IN ('03','04') AND D2.D2_FILIAL = '"+xFilial("SD2")+"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
DbSelectArea(cAliasSE1)
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2),0}) 	
aadd(aInadiplencia,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2),0,0}) 	
While (cAliasSE1)->(!Eof())
	If (cAliasSE1)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2)
		aTitRec[3,2]  		:= (cAliasSE1)->VALOR
		aInadiplencia[3,2]  := (cAliasSE1)->VALOR
	EndIf	
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
FmtLin({    Transform(aTitRec[1,2],"@R 99,999,999.99"),;
			Transform(aTitRec[2,2],"@R 99,999,999.99"),;
			Transform(aTitRec[3,2],"@R 99,999,999.99")},aL[20],"",,@nLin)
FmtLin(,{aL[21],aL[55],aL[56],aL[57]},,,@nLin)
dDataIni := ctod("01/"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+"/"+Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)))         
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '" + SUBSTR(DTOS(dDataIni),1,6)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
DbSelectArea(cAliasSE1)    
If (cAliasSE1)->(!Eof())
	While (cAliasSE1)->(!Eof())
		aadd(aMes1,{Substr((cAliasSE1)->EMISSAO,5,2),(cAliasSE1)->VALOR})
		(cAliasSE1)->(DbSkip())
	EndDo
Else
	For nI := 1 to 3
		aadd(aMes1,{strzero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
	Next nI	
EndIf	   
If Len(aMes1) < 3
	For nI := 1 to 3
		aadd(aMes1,{strzero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
	Next nI	
EndIf	
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea()) 

dDataIni := ctod("01/"+StrZero(Month(MV_PAR02),2)+"/"+Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)))         
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '" + SUBSTR(DTOS(dDataIni),1,6)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1            
DbSelectArea(cAliasSE1)    
If (cAliasSE1)->(!Eof())
	While (cAliasSE1)->(!Eof())
		aadd(aMes2,{Substr((cAliasSE1)->EMISSAO,5,2),(cAliasSE1)->VALOR})
		(cAliasSE1)->(DbSkip())
	EndDo
Else
	For nI := 1 to 3
		aadd(aMes2,{strzero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0}) 
	Next nI	
EndIf	   
If Len(aMes2) < 3
	For nI := 1 to 3
		aadd(aMes2,{strzero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0}) 
	Next nI	
EndIf		
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())

dDataIni := ctod("01/"+StrZero(Month(MV_PAR02)+1,2)+"/"+Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)))         
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '" + SUBSTR(DTOS(dDataIni),1,6)+ "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Month(MV_PAR02),2) + "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
DbSelectArea(cAliasSE1)
If (cAliasSE1)->(!Eof())
	While (cAliasSE1)->(!Eof())
		aadd(aMes3,{(cAliasSE1)->EMISSAO,(cAliasSE1)->VALOR})
		(cAliasSE1)->(DbSkip())
	EndDo
Else
	For nI := 1 to 3
		aadd(aMes3,{strzero(nMes,2),0}) 
	Next nI	
EndIf	   
If Len(aMes3) < 3
	For nI := 1 to 3
		aadd(aMes3,{strzero(nMes,2),0}) 
	Next nI	
EndIf	
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
FmtLin({    Transform(aMes1[1,2],"@R 99,999,999.99"),;
			Transform(aMes2[1,2],"@R 99,999,999.99"),;
			Transform(aMes3[1,2],"@R 99,999,999.99")},aL[22],"",,@nLin)
FmtLin(,{aL[23]},,,@nLin)     
FmtLin({    Transform(aMes1[2,2],"@R 99,999,999.99"),;
			Transform(aMes2[2,2],"@R 99,999,999.99"),;
			Transform(aMes3[2,2],"@R 99,999,999.99")},aL[24],"",,@nLin)
FmtLin(,{aL[25]},,,@nLin)     
FmtLin({    Transform(aMes1[3,2],"@R 99,999,999.99"),;
			Transform(aMes2[3,2],"@R 99,999,999.99"),;
			Transform(aMes3[3,2],"@R 99,999,999.99")},aL[26],"",,@nLin)
FmtLin(,{aL[27]},,,@nLin)     
cQuery := " SELECT SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >  '" + SUBSTR(DTOS(MV_PAR02+90),1,6) + "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
aMes1 := {}
DbSelectArea(cAliasSE1)
While (cAliasSE1)->(!Eof())
	aadd(aMes1,{"",(cAliasSE1)->VALOR})
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
cQuery := " SELECT SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >  '" + SUBSTR(DTOS(MV_PAR02+90),1,6) + "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+ "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
aMes2 := {}
DbSelectArea(cAliasSE1)
While (cAliasSE1)->(!Eof())
	aadd(aMes2,{"",(cAliasSE1)->VALOR})
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
cQuery := " SELECT SUM(E1_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE1")    
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND E1_ORIGEM = 'MATA460' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >  '" + SUBSTR(DTOS(MV_PAR02+90),1,6) + "' "
cQuery += " AND SUBSTRING(E1_EMISSAO,1,6) <= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Month(MV_PAR02),2) + "' "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
aMes3 := {}
DbSelectArea(cAliasSE1)
While (cAliasSE1)->(!Eof())
	aadd(aMes3,{"",(cAliasSE1)->VALOR})
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
FmtLin({    Transform(aMes1[1,2],"@R 999,999,999.99"),;
			Transform(aMes2[1,2],"@R 999,999,999.99"),;
			Transform(aMes3[1,2],"@R 999,999,999.99")},aL[28],"",,@nLin)
FmtLin(,{aL[29]},,,@nLin)     
aTitRec := {}
cQuery := " SELECT SUBSTRING(E5_DATA,1,6) AS EMISSAO , SUM(E5_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE5")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E5_FILIAL = '"+xFilial("SE5")+"' "
cQuery += " AND E5_RECPAG = 'R' " 
cQuery += " AND E5_TIPO <> 'NCC' " 
cQuery += " AND (E5_NATUREZ = '800401' OR E5_NATUREZ = '      ') " 
cQuery += " AND SUBSTRING(E5_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E5_DATA,1,6) <= '" + substr(dtos(MV_PAR02),1,6)+ "' "
cQuery += " GROUP BY SUBSTRING(E5_DATA,1,6) "
cQuery += " ORDER BY SUBSTRING(E5_DATA,1,6) ASC "
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE5,.T.,.T.)},"Aguarde...","Processando Liquidados...")
For nSE5 := 1 To len(aStruSE5)
	If aStruSE5[nSE5][2] <> "C" .And. FieldPos(aStruSE5[nSE5][1])<>0
		TcSetField(cAliasSE5,aStruSE5[nSE5][1],aStruSE5[nSE5][2],aStruSE5[nSE5][3],aStruSE5[nSE5][4])
	EndIf
Next nSE5
dbSelectArea(cAliasSE5)
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 	 
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0}) 	
aadd(aTitRec,{substr(dtos(MV_PAR02),1,6),0}) 	
while (cAliasSE5)->(!eof())
	If (cAliasSE5)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aTitRec[1,2] := (cAliasSE5)->VALOR
	ElseIf (cAliasSE5)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aTitRec[2,2] := (cAliasSE5)->VALOR
	ElseIf (cAliasSE5)->EMISSAO == substr(dtos(MV_PAR02),1,6)
		aTitRec[3,2] := (cAliasSE5)->VALOR
	EndIf
	(cAliasSE5)->(DbSkip())
EndDo
DbSelectArea(cAliasSE5)
(cAliasSE5)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
FmtLin({    Transform(aTitRec[1,2],"@R 99,999,999.99"),;
			Transform(aTitRec[2,2],"@R 99,999,999.99"),;
			Transform(aTitRec[3,2],"@R 99,999,999.99")},aL[30],"",,@nLin)
FmtLin(,{aL[31]},,,@nLin)          
If nLin >= 60                                               
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
FmtLin(,{aL[32]},,,@nLin)     
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
EndIf
FmtLin(,{aL[33]},,,@nLin)    
aPagamentos:= {}
cQuery := "SELECT E5_MOTBX,SUBSTRING(E5_DATA,5,2) AS EMISSAO ,SUM(E5_VALOR) AS VALOR FROM "
cQuery += RetSqlName("SE5")
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E5_FILIAL = '"+xFilial("SE5")+"'  "                                
cQuery += " AND SUBSTRING(E5_DATA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E5_DATA,1,6) <= '" + substr(dtos(MV_PAR02),1,6)+ "' "
cQuery += " AND E5_RECPAG = 'R' " 
cQuery += " AND E5_TIPO <> 'NCC' " 
cQuery += " AND (E5_NATUREZ = '800401' OR E5_NATUREZ = '      ') " 
cQuery += " GROUP BY E5_MOTBX,SUBSTRING(E5_DATA,5,2) "
cQuery += " ORDER BY E5_MOTBX,SUBSTRING(E5_DATA,5,2) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE5,.T.,.T.)},"Aguarde...","Processando Recebimentos...")
For nSE5 := 1 To len(aStruSE5)
	If aStruSE5[nSE5][2] <> "C" .And. FieldPos(aStruSE5[nSE5][1])<>0
		TcSetField(cAliasSE5,aStruSE5[nSE5][1],aStruSE5[nSE5][2],aStruSE5[nSE5][3],aStruSE5[nSE5][4])
	EndIf
Next nSE5
dbSelectArea(cAliasSE5)
while (cAliasSE5)->(!eof())
	nReg := aScan(aPagamentos,{|x| x[2]==IIf((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA","COMPENSACAO"))))})
	If nReg >0 
		If Val((cAliasSE5)->EMISSAO) == Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
			aPagamentos[nReg,3] := (cAliasSE5)->VALOR
		ElseIf Val((cAliasSE5)->EMISSAO) == Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)			
			aPagamentos[nReg,4] := (cAliasSE5)->VALOR
		ElseIf Val((cAliasSE5)->EMISSAO) == nMes
			aPagamentos[nReg,5] := (cAliasSE5)->VALOR
		EndIf	
	Else
		If Val((cAliasSE5)->EMISSAO) == Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,IIf((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA","COMPENSACAO")))),(cAliasSE5)->VALOR,0,0}) 
		ElseIf Val((cAliasSE5)->EMISSAO) == Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)			
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,IIf((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA","COMPENSACAO")))),0,(cAliasSE5)->VALOR,0}) 
		ElseIf Val((cAliasSE5)->EMISSAO) == nMes
			aadd(aPagamentos,{(cAliasSE5)->EMISSAO,IIf((cAliasSE5)->E5_MOTBX=="DAC","DACAO",Iif((cAliasSE5)->E5_MOTBX=="NOR","NORMAL",Iif((cAliasSE5)->E5_MOTBX=="DEV","DEVOLUCAO",Iif((cAliasSE5)->E5_MOTBX=="LOJ","OUTRA LOJA","COMPENSACAO")))),0,0,(cAliasSE5)->VALOR}) 		
		EndIf	
	Endif
	(cAliasSE5)->(DbSkip())
EndDo
DbSelectArea(cAliasSE5)
(cAliasSE5)->(DbCloseArea())
For nI := 1 To len(aPagamentos)
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
	FmtLin({aPagamentos[nI,2],;    
			Transform(aPagamentos[nI,3],"@R 99,999,999.99"),;
			Transform(aPagamentos[nI,4],"@R 99,999,999.99"),;
			Transform(aPagamentos[nI,5],"@R 99,999,999.99")},aL[34],"",,@nLin)
Next nI
FmtLin(,{aL[35]},,,@nLin)
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
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,6) AS EMISSAO , SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) <= '" + substr(dtos(MV_PAR02),1,6)+ "' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,6) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
aTitRec := {}        
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 	 
aadd(aTitRec,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0}) 	
aadd(aTitRec,{substr(dtos(MV_PAR02),1,6),0}) 	
While (cAliasSE1)->(!Eof())                 
	nReg := aScan(aInadiplencia,{|x| x[1]==(cAliasSE1)->EMISSAO})
	If nReg > 0
		aInadiplencia[nReg,03] := (cAliasSE1)->VALOR
	EndIf
	If (cAliasSE1)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aTitRec[1,2]  := (cAliasSE1)->VALOR
	ElseIf (cAliasSE1)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aTitRec[2,2]  := (cAliasSE1)->VALOR
	Elseif (cAliasSE1)->EMISSAO == substr(dtos(MV_PAR02),1,6)
		aTitRec[3,2]  := (cAliasSE1)->VALOR
	EndIf	
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[2,2],"@R 99,999,999.99"),;
		Transform(aTitRec[3,2],"@R 99,999,999.99")},aL[36],"",,@nLin)
FmtLin(,{aL[37]},,,@nLin)    
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
EndIf
FmtLin({Transform((aInadiplencia[1,3]/aInadiplencia[1,2])*100,"@R 9,999.99"),;
	    Transform((aInadiplencia[2,3]/aInadiplencia[2,2])*100,"@R 9,999.99"),;
		Transform((aInadiplencia[3,3]/aInadiplencia[3,2])*100,"@R 9,999.99")},aL[38],"",,@nLin)
FmtLin(,{aL[39]},,,@nLin)    
cQuery := " SELECT SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '200601' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) <= '" +"2006"+STRZERO(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+ "' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04')  "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1           
While (cAliasSE1)->(!Eof())
		aTitRec[1,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
cQuery := " SELECT SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '200601' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) <= '" +"2006"+STRZERO(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+ "' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04')  "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1           
While (cAliasSE1)->(!Eof())
		aTitRec[2,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
cQuery := " SELECT SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) >= '200601' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,6) <= '200612' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04','05')  "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1           
While (cAliasSE1)->(!Eof())
		aTitRec[3,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf    
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[2,2],"@R 99,999,999.99"),;
		Transform(aTitRec[3,2],"@R 99,999,999.99")},aL[40],"",,@nLin)
FmtLin(,{aL[41]},,,@nLin)    
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,4) AS EMISSAO , SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,4) = '2005' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04')  "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,4) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,4) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1                    
If (cAliasSE1)->(!Eof())
	While (cAliasSE1)->(!Eof())
		aTitRec[1,1]  := (cAliasSE1)->EMISSAO
		aTitRec[1,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
	EndDo
EndIf
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())

If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
EndIf    
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99")},aL[42],"",,@nLin)
FmtLin(,{aL[43]},,,@nLin)  
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,4) AS EMISSAO , SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,4) = '2004' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,4) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,4) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
If (cAliasSE1)->(!Eof())
	While (cAliasSE1)->(!Eof())
			aTitRec[1,1]  := (cAliasSE1)->EMISSAO
			aTitRec[1,2]  := (cAliasSE1)->VALOR
		(cAliasSE1)->(DbSkip())
	EndDo
EndIf	
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf    
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99")},aL[44],"",,@nLin)
FmtLin(,{aL[45]},,,@nLin) 
cQuery := " SELECT SUBSTRING(E1_VENCREA,1,4) AS EMISSAO , SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,4) = '2003' "
cQuery += " AND E1_SALDO > 0 "     
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM  "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(E1_VENCREA,1,4) "
cQuery += " ORDER BY SUBSTRING(E1_VENCREA,1,4) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1
While (cAliasSE1)->(!Eof())
	aTitRec[1,1]  := (cAliasSE1)->EMISSAO
	aTitRec[1,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf    
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99"),;
		Transform(aTitRec[1,2],"@R 99,999,999.99")},aL[46],"",,@nLin)
FmtLin(,{aL[47]},,,@nLin)
cQuery := " SELECT SUM(E1_SALDO) AS VALOR FROM "
cQuery += RetSqlName("SE1")
cQuery += " WHERE D_E_L_E_T_ <> '*' "
cQuery += " AND E1_FILIAL = '"+xFilial("SE1")+"' "
cQuery += " AND SUBSTRING(E1_VENCREA,1,4) <= '2002' "
cQuery += " AND E1_SALDO > 0 "
cQuery += " AND E1_NUM+E1_PREFIXO IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=  RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_TIPO = 'N' "
cQuery += " AND D2.D2_LOCAL IN ('03','04') "
cQuery += " AND D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ")  "   
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE1,.T.,.T.)
For nSE1 := 1 To len(aStruSE1)
	If aStruSE1[nSE1][2] <> "C" .And. FieldPos(aStruSE1[nSE1][1])<>0
		TcSetField(cAliasSE1,aStruSE1[nSE1][1],aStruSE1[nSE1][2],aStruSE1[nSE1][3],aStruSE1[nSE1][4])
	EndIf
Next nSE1           
While (cAliasSE1)->(!Eof())
	aTitRec[1,1]  := ""
	aTitRec[1,2]  := (cAliasSE1)->VALOR
	(cAliasSE1)->(DbSkip())
EndDo
DbSelectArea(cAliasSE1)
(cAliasSE1)->(DbCloseArea())
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[7],aL[6]},,,@nLin) 
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)
EndIf
FmtLin({Transform(aTitRec[1,2],"@R 99,999,999.99")},aL[48],"",,@nLin)
FmtLin(,{aL[49]},,,@nLin)
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATADMGE) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Gastos Gerais Adm....")
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
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5]},,,@nLin)

EndIf	
FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),;
	    Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[50],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
dbCloseArea()
FmtLin(,{aL[51]},,,@nLin)	

U_CRF010H(Cabec1,Cabec2,Titulo,@nLin,nMes)

RETURN(.T.)
