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

User Function CRF010C(nLin,nMes,Cabec1,Cabec2,Titulo)

Local aArea := GetArea()
Local cQuery := ""
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSD2 	:= SD2->(dbStruct())
Local aStruSF2 	:= SF2->(dbStruct())
Local nSC5	 	:= 0
Local nSD2	 	:= 0
Local nSF2	 	:= 0
Local nSD1	 	:= 0
Local nSF1	 	:= 0
Local nX        := 0
Local cAliasReg := "cALiasReg"
Local aFreteFob := {}
Local aFreteCif := {}
Local aValCif   := {}  
Local aFreteDif := {}
Local cALiasSE7 := "cALiasSE7"
Local aStruSE7  := SE7->(dbStruct())
Local nSE7		:= 0
Local nFATFrete := 0 
LoCAL aPerFrete := {}
Local aVLFrete  := {}
//Logistica
U_RFinLay5()
Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
nLin:=PROW()+1                                    
FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin) 
FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
FmtLin(,{aL[5],aL[6],aL[7]},,,@nLin)
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteDif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteDif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteFob,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteFob,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteFob,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteFob,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteFob,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteFob,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteCif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteCif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aValCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aValCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aValCif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aValCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aValCif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aValCif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteDif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aFreteDif,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aFreteDif,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aPerFrete,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aPerFrete,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aPerFrete,{SUBSTR(DTOS(MV_PAR02),1,6),0})
aadd(aVLFrete,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0}) 
aadd(aVLFrete,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aVLFrete,{SUBSTR(DTOS(MV_PAR02),1,6),0})

cQuery := " SELECT SUBSTRING(F2.F2_EMISSAO,1,6) AS EMISSAO ,SUM(F2.F2_VALFAT) AS TOTAL FROM "                    
cQuery +=   RetSqlName('SF2') + " F2 WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"' "
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND F2.F2_TIPO = 'N' AND F2.F2_DOC+F2.F2_SERIE IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=   RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND D2.D2_TIPO = 'N' AND D2.D2_LOCAL IN ('03','04','05') AND D2.D_E_L_E_T_ <> '*' "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVdInt) + ") AND D2.D2_PEDIDO IN (SELECT C5_NUM FROM " 
cQuery +=   RetSqlName('SC5') + " C5  WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_TPFRETE = 'F'  AND C5.D_E_L_E_T_ <> '*')) "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)},"Aguarde...","Processando Frete Fob...")
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cALiasREG,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2

dbSelectArea(cAliasREG)
If (cAliasREG)->(!Eof())
	While (cAliasREG)->(!EOF())
			If (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
				aFreteFob[1,2]:= (cALiasREG)->TOTAL
			ElseIf (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
				aFreteFob[2,2]:= (cALiasREG)->TOTAL
			ElseIf (cAliasREG)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
				aFreteFob[3,2]:= (cALiasREG)->TOTAL
			Endif	
		(cAliasREG)->(DbSkip())
	EndDo
EndIf
DbSelectArea(cAliasREG)
(cAliasREG)->(DbCloseArea())

FmtLin(,{aL[9]},,,@nLin)
cQuery := " SELECT SUBSTRING(F2.F2_EMISSAO,1,6) AS EMISSAO ,SUM(F2.F2_VALFAT) AS TOTAL "
cQuery += " FROM "                    
cQuery +=   RetSqlName('SF2') + " F2 "
cQuery += " WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"' "
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND F2.F2_TIPO = 'N' AND F2.F2_DOC+F2.F2_SERIE IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=   RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND D2.D2_TIPO = 'N'  AND D2.D2_LOCAL IN ('03','04','05') AND D2.D_E_L_E_T_ <> '*'  AND D2.D2_CF IN (" + Alltrim(cCFOPVdExt) + ") "
cQuery += " AND D2.D2_PEDIDO IN (SELECT C5_NUM FROM " 
cQuery +=   RetSqlName('SC5') + " C5 WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_TPFRETE = 'F' AND C5.D_E_L_E_T_ <> '*')) "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)},"Aguarde...","Processando Frete Fob...")
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cALiasREG,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
dbSelectArea(cAliasREG)
If (cAliasREG)->(!Eof())
	While (cAliasREG)->(!EOF())
			If (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
				aFreteFob[4,2]:= (cALiasREG)->TOTAL
			ElseIf (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
				aFreteFob[5,2]:= (cALiasREG)->TOTAL
			ElseIf (cAliasREG)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
				aFreteFob[6,2]:= (cALiasREG)->TOTAL
			Endif	
		(cAliasREG)->(DbSkip())
	EndDo
EndIf
DbSelectArea(cAliasREG)
(cAliasREG)->(DbCloseArea())
FmtLin({Transform(aFreteFob[1,2],"@R 99,999,999.99"),;
		Transform(aFreteFob[2,2],"@R 99,999,999.99"),;
		Transform(aFreteFob[3,2],"@R 99,999,999.99"),;
		Transform(aFreteFob[4,2],"@R 99,999,999.99"),;
		Transform(aFreteFob[5,2],"@R 99,999,999.99"),;
		Transform(aFreteFob[6,2],"@R 99,999,999.99")},aL[08],"",,@nLin)
FmtLin(,{aL[9]},,,@nLin)
cQuery := " SELECT SUBSTRING(F2.F2_EMISSAO,1,6) AS EMISSAO ,SUM(F2_VALFAT) AS TOTAL FROM "                    
cQuery +=   RetSqlName('SF2') + " F2  WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"' "
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND F2.F2_TIPO = 'N' AND F2.F2_DOC+F2.F2_SERIE IN (SELECT D2.D2_DOC+D2.D2_SERIE FROM "
cQuery +=   RetSqlName('SD2') + " D2  WHERE D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND D2.D2_TIPO = 'N'  AND D2.D2_LOCAL IN ('03','04','05')  AND D2.D_E_L_E_T_ <> '*'  AND D2.D2_CF IN (" + Alltrim(cCFOPVdInt) + ") "
cQuery += " AND D2.D2_PEDIDO IN (SELECT C5_NUM FROM " 
cQuery +=   RetSqlName('SC5') + " C5 "
cQuery += " WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_TPFRETE = 'C' AND C5.D_E_L_E_T_ <> '*')) "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)},"Aguarde...","Processando Frete Fob...")
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField(cALiasREG,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
dbSelectArea(cAliasREG)
If (cAliasREG)->(!EOF())                                  
	While (cAliasREG)->(!EOF())                                  
		If (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aFreteCif[1,2]:= (cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aFreteCif[2,2]:= (cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aFreteCif[3,2]:= (cALiasREG)->TOTAL
		EndIf	
		(cAliasREG)->(DbSkip())
	EndDo
EndIf

DbSelectArea(cAliasREG)
(cAliasREG)->(DbCloseArea())

FmtLin(,{aL[11]},,,@nLin)
cQuery := " SELECT SUBSTRING(D2.D2_EMISSAO,1,6) AS EMISSAO ,SUM(D2_VALFRE) AS TOTAL FROM "
cQuery += RetSqlName('SD2') + " D2  WHERE  D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_TIPO = 'N' AND D2.D_E_L_E_T_ <> '*' "             
cQuery += " AND D2.D2_LOCAL IN ('03','04','05') AND D2.D2_CF IN (" + Alltrim(cCFOPVdExt) + ") AND D2.D2_PEDIDO IN (SELECT C5_NUM FROM " 
cQuery += RetSqlName('SC5') + " C5  WHERE  C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_TPFRETE = 'C' AND C5.D_E_L_E_T_ <> '*') "
cQuery += " GROUP BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D2.D2_EMISSAO,1,6) ASC "

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)},"Aguarde...","Processando Frete CIF...")

For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cALiasREG,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
DbSelectArea(cAliasREG)
If (cAliasREG)->(!EOF())                                  
	While (cAliasREG)->(!EOF())                                  
		If (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aFreteCif[4,2]:= (cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aFreteCif[5,2]:= (cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aFreteCif[6,2]:= (cALiasREG)->TOTAL
		EndIf	
		(cAliasREG)->(DbSkip())
	EndDo
EndIf
DbSelectArea(cAliasREG)
(cAliasREG)->(DbCloseArea())
FmtLin({Transform(aFreteCif[1,2],"@R 99,999,999.99"),;
		Transform(aFreteCif[2,2],"@R 99,999,999.99"),;
		Transform(aFreteCif[3,2],"@R 99,999,999.99"),; 
		Transform(aFreteCif[4,2],"@R 99,999,999.99"),;
		Transform(aFreteCif[5,2],"@R 99,999,999.99"),;
		Transform(aFreteCif[6,2],"@R 99,999,999.99")},aL[10],"",,@nLin)
FmtLin(,{aL[11]},,,@nLin)

For nX := nMes-2 to nMes
	cQuery := " SELECT F2_DOC,F2_PLIQUI, SUM(D2_QUANT) as QtdVen, SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS TOTAL FROM "
	cQuery += RetSqlName('SF2') + " F2, "
	cQuery += RetSqlName('SD2') + " D2, "
	cQuery += RetSqlName('SC5') + " C5 WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"' "                                        
	If nMes == nX
		cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,6) + "' AND "			
	ElseIf nX == nMes-2 
		cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "' AND "	
	ElseIf nX == nMes-1 		
		cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2) + "' AND "		
	EndIf
	cQuery += " F2.F2_TIPO = 'N' AND F2.D_E_L_E_T_ = ' ' AND D2.D2_FILIAL = '"+xFilial("SD2")+"' AND D2.D2_CF IN (" + Alltrim(cCFOPVdInt) + ") "
	cQuery += "AND D2.D2_DOC = F2.F2_DOC AND D2.D2_SERIE = F2.F2_SERIE AND D2.D2_LOCAL IN ('03','04','05') AND D2.D_E_L_E_T_ = ' ' " 
	cQuery += "AND C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_NUM = D2_PEDIDO AND C5_TPFRETE = 'C' AND C5.D_E_L_E_T_ <> '*' "
	cQuery += "GROUP BY F2_DOC,F2_PLIQUI "
	cQuery += "ORDER BY F2_DOC "
	cQuery:=ChangeQuery(cQuery)
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)
	For nSD2 := 1 To len(aStruSD2)
		If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
			TcSetField(cALiasREG,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
		EndIf
	Next nSD2
	For nSF2 := 1 To len(aStruSF2)
		If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
			TcSetField(cALiasREG,aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
		EndIf
	Next nSF2
	dbSelectArea(cAliasREG)
	nFatTotal := 0
	nDevTotal := 0
	nFreTotal := 0
	nQtdTotal := 0
	nFatTrans := 0
	nDevTrans := 0
	nFreTrans := 0
	nQtdTrans := 0
	nCusTrans := 0
	nVLFrt  := 0
	nPesoNF	  := 0   
	while (cAliasREG)->(!EOF())
			nPesoNF	  := 0
			SF2->(DbSetOrder(1))
			SF2->(DbSeek(xFilial('SF2')+(cAliasREG)->F2_DOC))
			SC6->(DbSetOrder(4))
			SC6->(DbSeek(xFilial('SC6')+(cAliasREG)->F2_DOC))
			SD2->(DbSetOrder(3))
			SD2->(DbSeek(xFilial('SD2')+(cAliasREG)->F2_DOC))
			Do While SF2->F2_DOC == SD2->D2_DOC .AND. SF2->F2_FILIAL == SD2->D2_FILIAL
				DbSelectArea("SB1")
				DbSetOrder(1)
				DbSeek(xFilial("SB1")+SD2->D2_COD)
				nPesoNF += SB1->B1_PESO * SD2->D2_QUANT
				SD2->(DbSkip())
			EndDo
			If nPesoNF <> (cAliasREG)->F2_PLIQUI
				RecLock("SF2",.F.)
				SF2->F2_PLIQUI := nPesoNF
				MsUnLock()
			EndIf
			cNumped:=SC6->C6_num
			cCodi:=space(6)
			cLoja:=space(2)
			cNome:=space(40)
			cUf:=space(2)
			SC5->(DbSetOrder(1))
			SC5->(dbseek(xFilial('SC5')+cNumped))
			
			SA1->(DbSetOrder(1))
			SA1->(dbseek(xFilial('SA1')+SF2->(F2_CLIENTE+F2_LOJA)))
			cCodi:=SA1->A1_COD
			cLoja:=SA1->A1_LOJA
			cNome:=SA1->A1_NREDUZ
			cMuni:=SA1->A1_MUN
			cUf  :=SA1->A1_EST
			dbSelectArea("SZ0") 
			dbSetOrder(1)  
			MsSeek(xFilial("SZ0")+cMuni+cUF)
			If (SF2->F2_PLIQUI/1000) > SZ0->Z0_QTDMIN 
				nVLFrt := NoRound(SZ0->Z0_FRETE * ((SF2->F2_PLIQUI/1000)/1000),2)
			Else 
				nVLFrt := SZ0->Z0_VLMIN
			EndIf	
			If nX == nMes
				aVLFrete[3,2] += Iif(SF2->F2_FRETE > 0 ,SF2->F2_FRETE,nVLFrt)
			ElseIf nX == nMes-1
				aVLFrete[2,2] += Iif(SF2->F2_FRETE > 0 ,SF2->F2_FRETE,nVLFrt)
			ElseIf nX == nMes-2
				aVLFrete[1,2] += Iif(SF2->F2_FRETE > 0 ,SF2->F2_FRETE,nVLFrt)
			EndIf
			nFATFrete += (cAliasREG)->TOTAL
			(cAliasREG)->(dbskip())
	EndDo
	DbSelectArea(cAliasREG)              
	(cAliasREG)->(DbCloseArea())
	If nX == nMes
		aPerFrete[3,2] := (aVLFrete[3,2]/nFATFrete)*100
	ElseIf nX == nMes-1
		aPerFrete[2,2] := (aVLFrete[2,2]/nFATFrete)*100	
	ElseIf nX == nMes-2
		aPerFrete[1,2] := (aVLFrete[1,2]/nFATFrete)*100
	EndIf
	nFATFrete := 0
Next nX
FmtLin({Transform(aVLFrete[1,2],"@R 99,999,999.99"),;
		Transform(aVLFrete[2,2],"@R 99,999,999.99"),;
		Transform(aVLFrete[3,2],"@R 99,999,999.99")},aL[12],"",,@nLin)
FmtLin(,{aL[13]},,,@nLin)
FmtLin({Transform(aPerFrete[1,2],"@R 999.99"),;
		Transform(aPerFrete[2,2],"@R 999.99"),;
		Transform(aPerFrete[3,2],"@R 999.99")},aL[14],"",,@nLin)
FmtLin(,{aL[15]},,,@nLin) 
cQuery := " SELECT SUBSTRING(D2.D2_EMISSAO,1,6) AS EMISSAO ,SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS TOTAL "
cQuery += " FROM "
cQuery += RetSqlName('SD2') + " D2 WHERE  D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"'  " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"'  "
cQuery += " AND D2.D2_TIPO = 'N' AND D2.D_E_L_E_T_ <> '*' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND D2.D2_PEDIDO IN (SELECT C5_NUM FROM " 
cQuery += RetSqlName('SC5') + " C5  WHERE  C5.C5_FILIAL = '"+xFilial("SC5")+"' AND C5_TPFRETE NOT IN ('C','F') AND C5.D_E_L_E_T_ <> '*') "
cQuery += " GROUP BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D2.D2_EMISSAO,1,6) ASC "

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cALiasREG,.T.,.T.)},"Aguarde...","Processando Frete Diversos...")

For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField(cALiasREG,aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
dbSelectArea(cAliasREG)
If (cAliasREG)->(!EOF())
	While (cAliasREG)->(!EOF())
		If (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aFreteDif[1,2]:=(cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aFreteDif[2,2]:=(cALiasREG)->TOTAL
		ElseIf (cAliasREG)->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aFreteDif[3,2]:=(cALiasREG)->TOTAL
		EndIf	
		(cAliasREG)->(DbSkip())
	EndDo
Endif
DbSelectArea(cAliasREG)
(cAliasREG)->(DbCloseArea())

FmtLin({Transform(aFreteDif[1,2],"@R 99,999,999.99"),;
		Transform(aFreteDif[2,2],"@R 99,999,999.99"),;
		Transform(aFreteDif[3,2],"@R 99,999,999.99")},aL[16],"",,@nLin)

FmtLin(,{aL[17],aL[18],aL[19]},,,@nLin)    
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600201' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Importacao (600201)")
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
FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[20],"",,@nLin)
	(cAliasSE7)->(DbSkip())
EndDo
DbSelectArea(cAliasSE7)
DbCloseArea()
FmtLin(,{aL[21]},,,@nLin)	
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND E7_FILIAL = '"+xFilial("SE7")+"' AND E7_NATUREZ = '600401' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Exportacao (600401)")
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
FmtLin({Transform((cAliasSE7)->VALOR1,"@R 99,999,999.99"),Transform((cAliasSE7)->VALOR2,"@R 99,999,999.99"),;
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[22],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[23]},,,@nLin)	

cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600601' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Benefic.   (600601)")
	
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
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[24],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[25]},,,@nLin)

cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600501' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Transf.    (600501)")
	
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
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[26],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[27]},,,@nLin)

cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600303' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Compras COM(600303)")
	
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
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[28],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[29]},,,@nLin)

cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600301' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Compras IND(600301)")
	
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
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[30],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[31]},,,@nLin)


cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '600302' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)},"Aguarde...","Frete Compras ADM(600302)")
	
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
		Transform((cAliasSE7)->VALOR3,"@R 99,999,999.99")},aL[32],"",,@nLin)

	(cAliasSE7)->(DbSkip())
EndDo
	
DbSelectArea(cAliasSE7)
dbCloseArea()

FmtLin(,{aL[33]},,,@nLin)



RestArea(aArea)

RETURN(.T.)
