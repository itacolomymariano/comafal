#INCLUDE "RWMAKE.CH"
/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?RunReport ?Autor  ?Eduardo Zanardo     ? Data ?  03/31/03   ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? AP7                                                        ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/
User Function CRF010B(Cabec1,Cabec2,Titulo,nLin)

Local aArea 	:= GetArea()
Local aVetor 	:= {}
Local cQuery 	:= ""
Local nSD2	 	:= 0
Local nSF2	 	:= 0
Local nSD1	 	:= 0                       
Local nSC9      := 0 
Local aStruSC9 	:= SC9->(dbStruct())  
Local aStruSD2 	:= SD2->(dbStruct())
Local aStruSF2 	:= SF2->(dbStruct())
Local aStruSD1 	:= SD1->(dbStruct())
Local aStruSE7  := SE7->(dbStruct())
Local nI 		:= 0
Local nReg 		:= 0
Local nCustoAcu := 0
Local nVlPedLiq := 0
Local nVlPedDec := 0
Local nVlPedBrt := 0
Local nValFat	:= 0
Local nValDev   := 0
Local nVdNormal := 0
Local nAcuBrt 	:= 0
Local nAcuLiq 	:= 0
Local nQtdPed 	:= 0
Local nMes := Month(MV_PAR02)      
Local aCondFat1  := {}  
Local aPrcMedio := {}
Local aQtdVend := {} 
Local aQtdPed := {0,0,0}
Local aVlPedLiq  := {0,0,0}
Local aComissao := {}
Local aPComis := {}
Local aGastosCom:= {}
Local aFatLFDIF := {}  
Local aFatFTDIF := {}  
Local aFatOutros:= {}  
Local nX := 0
Local aCustoDev := {}
Local nSD3 := 0
Local aStruSD3	:= SD3->(dbStruct())
//??????????????????????????????????????????????????????????????????Ŀ
//?aVetor[n,1] = Dia                                                 ?
//?aVetor[n,2] = Pedidos Bruto                                       ?
//?aVetor[n,3] = Pedidos Desconto                                    ?
//?aVetor[n,4] = Pedidos Liquido    	                             ?
//?aVetor[n,5] = Faturados Bruto			                         ?
//?aVetor[n,6] = Devolucoes     	                                 ?
//?aVetor[n,7] = Bonificacoes						                 ?
//?aVetor[n,8] = Faturados Liq.				                         ?
//?aVetor[n,9] = Acumulado Bruto                 					 ?
//?aVetor[n,10] = Acumulado Liq.				    				 ?
//?aVetor[n,11] = Custo		    									 ?
//?aVetor[n,12] = Margem Liq.%	DIA									 ?
//?aVetor[n,13] = Margem Liq.%	ACUM								 ?
//?aVetor[n,15] = Qtd de Pedidos         							 ?
//????????????????????????????????????????????????????????????????????
aadd(aFatBruto,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0,0,0})  
aadd(aFatBruto,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0,0,0}) 
aadd(aFatBruto,{SUBSTR(DTOS(MV_PAR02),1,6),0,0,0,0}) 
aadd(aDevInt,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aDevInt,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aDevInt,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
aadd(aDevExt,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aDevExt,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aDevExt,{SUBSTR(DTOS(MV_PAR02),1,6),0,0}) 
aadd(aCstBruto,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aCstBruto,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aCstBruto,{SUBSTR(DTOS(MV_PAR02),1,6),0,0}) 
aadd(aPrcMedio,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aPrcMedio,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aPrcMedio,{SUBSTR(DTOS(MV_PAR02),1,6),0,0}) 
aadd(aQtdVend,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aQtdVend,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aQtdVend,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})       
aadd(aComissao,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aComissao,{0,0,0})
aadd(aGastosCom,{0,0,0}) 
aadd(aGastosCom,{0,0,0})
aadd(aCustoDev,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})  
aadd(aCustoDev,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0}) 
aadd(aCustoDev,{SUBSTR(DTOS(MV_PAR02),1,6),0,0}) 

cQuery := "SELECT C9.C9_DATALIB, SUM(C9.C9_PRCVEN*C9.C9_QTDLIB) as PedBrt, COUNT(DISTINCT C9.C9_PEDIDO) as QtdPED FROM "
cQuery += RetSqlName('SC9') + " C9 WHERE C9.C9_FILIAL = '"+xFilial("SC9")+"' AND "
cQuery += "C9.D_E_L_E_T_ = ' ' AND SUBSTRING(C9.C9_DATALIB,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND C9.C9_NFISCAL <> '      ' AND " 
cQuery += "C9.C9_PEDIDO IN (SELECT C5.C5_NUM FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 WHERE C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND C5.D_E_L_E_T_ = ' ' AND C6.C6_FILIAL = '"+xFilial("SC6")+"' AND C6.C6_NUM = C5.C5_NUM AND "
cQuery += "C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND C6.C6_LOCAL IN ('03','04','05') AND C6.D_E_L_E_T_ = ' ' ) "
cQuery += "GROUP BY C9.C9_DATALIB "
cQuery += "ORDER BY C9.C9_DATALIB ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BPED",.T.,.T.)
For nSC9 := 1 To len(aStruSC9)
	If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
		TcSetField("c010BPED",aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
	EndIf
Next nSC9
dbSelectArea("c010BPED")
c010BPED->(dbGoTop())
while c010BPED->(!eof())
		//                                                                              1                                  2
		//                              1,               2,3,               4,5,6,7,8,9,0,1,2,3,4,               5,6,7,8,9,0,1
		aadd(aVetor,{c010BPED->C9_DATALIB,c010BPED->PedBrt,0,c010BPED->PedBrt,0,0,0,0,0,0,0,0,0,0,c010BPED->QtdPED,0,0,0,0,0,0})
	c010BPED->(dbskip())
EndDo                     
DbSelectArea("c010BPED")
DbCloseArea()
// Devolu??es pegas de D1
cQuery := " SELECT D1_DTDIGIT,SUM(D1_TOTAL)AS DEVOLUCAO FROM "+RetSqlName('SD1')
cQuery += " WHERE  D1_FILIAL = '"+xFilial("SD1")+"' AND D1_DTDIGIT >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) + "01' AND "
cQuery += " D1_DTDIGIT <= '" + DTOS(MV_PAR02) + "' AND  D1_TIPO = 'D' AND D1_CF IN (" +cCFOPDvVd + ") AND  D1_NFORI <> '      ' AND "
cQuery += " D_E_L_E_T_ <> '*' " 
cQuery += " GROUP BY D1_DTDIGIT "
cQuery += " ORDER BY D1_DTDIGIT ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BDev",.T.,.T.)
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("c010BDev",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea("c010BDev")    
While c010BDev->(!eof())
	nReg := aScan(aVetor,{|x| x[1]==c010BDev->D1_DTDIGIT})
	If nReg > 0
		aVetor[nReg,06] := c010BDev->DEVOLUCAO
	Else 
		//                                                                        1                   2
		//                                 1,2,3,4,5,                     6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1
		aadd(aVetor,{c010BDev->D1_DTDIGIT,0,0,0,0,c010BDev->DEVOLUCAO,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0})
	EndIf
	c010BDev->(dbskip())
EndDo
DbSelectArea("c010BDev")    
DbCloseArea()                                                       
    
cQuery := "SELECT D2.D2_EMISSAO , SUM(D2_CUSTO1) AS REPOSICAO FROM "
cQuery +=  RetSqlName('SD2') + " D2 " 
cQuery += "WHERE D2.D2_TIPO    = 'N' AND D2.D2_FILIAL  = '"+xFilial("SD2")+"' "
cQuery += "AND D2.D2_LOCAL IN ('03','04','05') "
cQuery += "AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += "AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND D2.D_E_L_E_T_ = ' ' "
cQuery += "GROUP BY D2.D2_EMISSAO "
cQuery += "ORDER BY D2.D2_EMISSAO ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BNFS",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("c010BNFS",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2 
dbSelectArea("c010BNFS")
while c010BNFS->(!eof())
	nReg := aScan(aVetor,{|x| x[1]==c010BNFS->D2_EMISSAO})
	If nReg > 0
		aVetor[nReg,11] := c010BNFS->REPOSICAO 
	Else
		//                                                   1                                        2
		//                                 1,2,3,4,5,6,7,8,9,0,                     1,2,3,4,5,6,7,8,9,0,1
		aadd(aVetor,{c010BNFS->D2_EMISSAO,0,0,0,0,0,0,0,0,0,c010BNFS->REPOSICAO,0,0,0,0,0,0,0,0,0,0})
	EndIf
	c010BNFS->(dbskip())
EndDo                   
dbSelectArea("c010BNFS")
DbCloseArea()

cQuery := "SELECT F2.F2_EMISSAO, SUM(F2_VALFAT) AS TOTAL FROM "
cQuery +=  RetSqlName('SF2') + " F2 WHERE F2.F2_FILIAL  = '"+xFilial("SF2")+"' AND F2.D_E_L_E_T_ = ' ' AND "
cQuery += "F2.F2_EMISSAO >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"01' AND " 
cQuery += "F2.F2_EMISSAO <= '" + DTOS(MV_PAR02) +"' AND F2.F2_TIPO    = 'N' AND  "          
cQuery += "F2.F2_DOC+F2.F2_SERIE+F2.F2_EMISSAO IN (SELECT D2.D2_DOC+D2.D2_SERIE+D2.D2_EMISSAO FROM "
cQuery +=  RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND D2.D2_LOCAL IN ('03','04','05') AND "
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND D2.D_E_L_E_T_ = ' ') "
cQuery += "GROUP BY F2.F2_EMISSAO "
cQuery += "ORDER BY F2.F2_EMISSAO ASC"
cQuery:=ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BNFS",.T.,.T.)

For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("c010BNFS",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2    
dbSelectArea("c010BNFS")
while c010BNFS->(!eof())
	nReg := aScan(aVetor,{|x| x[1]==c010BNFS->F2_EMISSAO})
	If nReg > 0
		aVetor[nReg,05] := c010BNFS->TOTAL
	Else
		//                                                                    1                   2
		//                                 1,2,3,4,                 5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1
		aadd(aVetor,{c010BNFS->F2_EMISSAO,0,0,0,c010BNFS->TOTAL,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0})
	EndIf
	c010BNFS->(dbskip())
EndDo       
dbSelectArea("c010BNFS")
DbCloseArea()

SetRegua(len(aVetor))
//Acumulados Brt. e Liq. // Margem Liq.% DIA, Margem Liq.% ACUM
For nI := 1 To Len(aVetor)
	INCREGUA()
	aVetor[nI,8] := aVetor[nI,5]-(aVetor[nI,6]+aVetor[nI,7])
	If nI-1 > 0
		aVetor[nI,9]  := aVetor[nI-1,9]  + aVetor[nI,5]
		aVetor[nI,21] := aVetor[nI-1,21] + aVetor[nI,11]
		aVetor[nI,10] := aVetor[nI-1,10]+aVetor[nI,8]
		aVetor[nI,19] := aVetor[nI-1,19]
	Else
		aVetor[nI,9]  := aVetor[nI,5]
		aVetor[nI,10] := aVetor[nI,8]
		aVetor[nI,21] := aVetor[nI,11]
	Endif
	aVetor[nI,12] += Iif(aVetor[nI,5]/aVetor[nI,11] ==0,0,NoRound(((aVetor[nI,5]/aVetor[nI,11])-1)*100,2))
	aVetor[nI,13] += Iif(aVetor[nI,9]/aVetor[nI,21] ==0,0,NoRound(((aVetor[nI,9]/aVetor[nI,21])-1)*100,2))
Next nI
aVetor := asort( aVetor,,, { |x,y| x[1] < y[1] } )
SetRegua(len(aVetor))
U_RFinLayOut()
nMes :=Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))
For nI := 1 to Len(aVetor)
	INCREGUA()   
	If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                
		FmtLin(,{aL[1],aL[2],aL[3],aL[4]},,,@nLin)	
	EndIf	
	If Month(aVetor[nI,1]) <> nMes 
		FmtLin(,{aL[6]},,,@nLin)
		FmtLin({Transform(nQtdPed,"@R 9999"),Transform(nVlPedBrt,"@R 9999,999,999.99"),Transform(nVlPedDec,"@R 9999,999,999.99"),Transform(nVlPedLiq,"@R 9999,999,999.99"),;
				Transform(nValFat,"@R 9999,999,999.99"),Transform(nValDev,"@R 99,999,999.99"),Transform(0,"@R 99,999,999.99"),;
				Transform(nVdNormal,"@R 9999,999,999.99"),Transform(nAcuBrt,"@R 9999,999,999.99"),Transform(nAcuLiq,"@R 9999,999,999.99"),;
				Transform(nCustoAcu,"@R 9999,999,999.99")},aL[7],"",,@nLin)
		If MV_PAR04 == 1
			If nI-1 > 0
				U_ERF017(aVetor[nI-1,1],@nLin)
			Endif	
		EndIf	
		FmtLin(,{aL[8],aL[31],aL[15],aL[16]},,,@nLin)  
		aFatFTDIF := {}
		U_FatFTDIF(@aFatFTDIF,nMes)              
		If Len(aFatFTDIF) > 0
			For nX := 1 to len(aFatFTDIF)
				FmtLin({aFatFTDIF[nI,1],aFatFTDIF[nI,2],Transform(aFatFTDIF[nI,3],"@R 9999,999,999.99")},aL[17],"",,@nLin)
			Next nX
		EndIf
		FmtLin(,{aL[18],aL[32]},,,@nLin)
		aFatFTDIF := {}
        U_FatLFDIF(@aFatLFDIF,nMes) 
		FmtLin(,{aL[11],aL[12]},,,@nLin)  
		If Len(aFatLFDIF) > 0
			For nX := 1 to Len(aFatLFDIF)
				FmtLin({aFatLFDIF[nX,1],;
						aFatLFDIF[nX,2],;
						aFatLFDIF[nX,3],;
						aFatLFDIF[nX,4],;
						Transform(aFatLFDIF[nX,4],"@R 9999,999,999.99")},aL[13],"",,@nLin)
			Next nX
		EndIf                         
		FmtLin(,{aL[33],aL[34]},,,@nLin)
		FmtLin({Transform(U_FatLF(nMes),"@R 9999,999,999.99")},aL[09],"",,@nLin)
		FmtLin(,{aL[10],aL[14],aL[19],aL[20]},,,@nLin)   
		aFatOutros:= {}
		U_FatOutros(@aFatOutros,nMes)             
		If Len(aFatOutros) > 0
			For nX := 1 to Len(aFatOutros)
				FmtLin({Alltrim(aFatOutros[nX,1]),;
						Alltrim(aFatOutros[nX,2]),;
						Alltrim(aFatOutros[nX,3]),;
						Alltrim(aFatOutros[nX,4]),;
						Transform(aFatOutros[nX,5],"@R 999,999.999"),;
						Alltrim(aFatOutros[nX,6]),;
						Transform(aFatOutros[nX,7],"@R 99,999,999.99")},aL[21],"",,@nLin)
			Next nX
		EndIf
		FmtLin(,{aL[35],aL[22],aL[36],aL[23],aL[24]},,,@nLin)  
		aFatFin := {}        
		U_FatFin(@aFatFin,nMes)             
		If Len(aFatFin) > 0
			For nX := 1 to Len(aFatFin)
				FmtLin({Alltrim(aFatFin[nX,1]),;
						Alltrim(aFatFin[nX,2]),;
						Alltrim(aFatFin[nX,3]),;
						Alltrim(aFatFin[nX,4]),;
						Alltrim(aFatFin[nX,5]),;
						Alltrim(aFatFin[nX,6]),;
						Alltrim(aFatFin[nX,7]),;
						Transform(aFatFin[nX,8],"@R 999,999.999"),;
						Alltrim(aFatFin[nX,9]),;			
						Transform(aFatFin[nX,10],"@R 9999,999,999.99")},aL[25],"",,@nLin)
			Next nX
		EndIf
		FmtLin(,{aL[26],aL[37],aL[38],aL[27],aL[28]},,,@nLin)  
		aFatEst := {}
		U_FatEst(@aFatEst,nMes)             
		If Len(aFatEst) > 0
			For nX := 1 to len(aFatEst)
				FmtLin({Alltrim(aFatEst[nX,1]),;
						Alltrim(aFatEst[nX,2]),;
						Alltrim(aFatEst[nX,3]),;
						Alltrim(aFatEst[nX,4]),;
						Alltrim(aFatEst[nX,5]),;
						Alltrim(aFatEst[nX,6]),;
						Alltrim(aFatEst[nX,7]),;
						Transform(aFatEst[nX,8],"@R 999,999.999"),;
						Alltrim(aFatEst[nX,9]),;			
						Transform(aFatEst[nX,10],"@R 9999,999,999.99")},aL[29],"",,@nLin)
			Next nX
		EndIf	
		FmtLin(,{aL[39],aL[30]},,,@nLin)  
		U_CRF010G(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,@nLin,nMes)         
		U_RFinLayOut()
		nMes := Month(aVetor[nI,1])
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1              
		FmtLin(,{aL[1],aL[2],aL[3],aL[4]},,,@nLin)	
		nQtdPed   := 0
		nVlPedBrt := 0
		nVlPedDec := 0
		nVlPedLiq := 0
		nValFat   := 0
		nValDev   := 0
		nVdNormal := 0
		nAcuBrt   := 0
		nAcuLiq   := 0
		nCustoAcu := 0
	EndIf
	If aVetor[nI,2]+aVetor[nI,3]+aVetor[nI,4]+aVetor[nI,5]+aVetor[nI,6]+aVetor[nI,7] > 0
		FmtLin({aVetor[nI,1],; 					//Dia
		Transform(aVetor[nI,15],"@R 999"),;				//QTD
		Transform(aVetor[nI,2],"@R 9999,999,999.99"),; 	//Pedidos Bruto
		Transform(aVetor[nI,3],"@R 9999,999,999.99"),;		//Pedidos Desconto
		Transform(aVetor[nI,4],"@R 9999,999,999.99"),;		//Pedidos Liquido
		Transform(aVetor[nI,5],"@R 99999,999,999.99"),;	//Faturados Bruto
		Transform(aVetor[nI,6],"@R 99,999,999.99"),;		//Devolucoes
		Transform(aVetor[nI,7],"@R 99,999,999.99"),;		//Bonificacoes
		Transform(aVetor[nI,8],"@R 9999,999,999.99"),;		//Faturados Liq.
		Transform(aVetor[nI,9],"@R 9999,999,999.99"),;		//Acumulado Bruto
		Transform(aVetor[nI,10],"@R 9999,999,999.99"),;	//Acumulado Liq.
		Transform(aVetor[nI,11],"@R 9999,999,999.99"),;	//Custo
		Alltrim(Transform(aVetor[nI,12],"@E 999.99")),;
		Alltrim(Transform(aVetor[nI,13],"@E 999.99"))},;	//Margem Liq.%
		aL[05],"",,@nLin)
		nQtdPed   += aVetor[nI,15]    
		If nMes == (val(SUBSTR(DTOS(MV_PAR02),5,2))-2)
			aQtdPed[1] += aVetor[nI,15]    
			aVlPedLiq[1] += aVetor[nI,4]    
		ElseIf 	nMes == (val(SUBSTR(DTOS(MV_PAR02),5,2))-1)
			aQtdPed[2] += aVetor[nI,15]     
			aVlPedLiq[2] += aVetor[nI,4]    
		ElseIf 	nMes == val(SUBSTR(DTOS(MV_PAR02),5,2))
			aQtdPed[3] += aVetor[nI,15]    
			aVlPedLiq[3] += aVetor[nI,4]    
		EndIf	
		nVlPedBrt += aVetor[nI,2]  
		nVlPedDec += aVetor[nI,3]
		nVlPedLiq += aVetor[nI,4]
		nValFat   += aVetor[nI,5]
		nValDev   += aVetor[nI,6]
		nVdNormal += aVetor[nI,8]
		nAcuBrt   += aVetor[nI,5]
		nAcuLiq   += aVetor[nI,5]-(aVetor[nI,6]+aVetor[nI,7])
		nCustoAcu += aVetor[nI,11]
	EndIf
Next nI
DbSelectArea("SF2")
FmtLin(,{aL[6]},,,@nLin)  
FmtLin({Transform(nQtdPed,"@R 9999"),Transform(nVlPedBrt,"@R 9999,999,999.99"),Transform(nVlPedDec,"@R 9999,999,999.99"),;
				Transform(nVlPedLiq,"@R 9999,999,999.99"),Transform(nValFat,"@R 9999,999,999.99"),Transform(nValDev,"@R 99,999,999.99"),;
				Transform(0,"@R 99,999,999.99"),Transform(nVdNormal,"@R 9999,999,999.99"),Transform(nAcuBrt,"@R 9999,999,999.99"),;
				Transform(nAcuLiq,"@R 9999,999,999.99"),Transform(nCustoAcu,"@R 9999,999,999.99")},aL[7],"",,@nLin)
If MV_PAR04 == 1
	U_ERF017(MV_PAR02,@nLin)
EndIf   
FmtLin(,{aL[8],aL[31],aL[15],aL[16]},,,@nLin)  
aFatFTDIF := {}
U_FatFTDIF(@aFatFTDIF,nMes)              
If Len(aFatFTDIF) > 0
	For nI := 1 to len(aFatFTDIF)
		FmtLin({aFatFTDIF[nI,1],aFatFTDIF[nI,2],Transform(aFatFTDIF[nI,3],"@R 9999,999,999.99")},aL[17],"",,@nLin)
	Next nI
EndIf
FmtLin(,{aL[18],aL[32]},,,@nLin)
aFatFTDIF := {}
U_FatLFDIF(@aFatLFDIF,nMes) 
FmtLin(,{aL[11],aL[12]},,,@nLin)  
If Len(aFatLFDIF) > 0
	For nI := 1 to Len(aFatLFDIF)
		FmtLin({aFatLFDIF[nI,1],;
				aFatLFDIF[nI,2],;
				aFatLFDIF[nI,3],;
				aFatLFDIF[nI,4],;
				Transform(aFatLFDIF[nI,4],"@R 9999,999,999.99")},aL[13],"",,@nLin)
	Next nI
EndIf                         
FmtLin(,{aL[33],aL[34]},,,@nLin)
FmtLin({Transform(U_FatLF(nMes),"@R 9999,999,999.99")},aL[09],"",,@nLin)
FmtLin(,{aL[10],aL[14],aL[19],aL[20]},,,@nLin)   
aFatLFDIF := {}
aFatOutros:= {}
U_FatOutros(@aFatOutros,nMes)             
If Len(aFatOutros) > 0
	For nI := 1 to Len(aFatOutros)
		FmtLin({Alltrim(aFatOutros[nI,1]),;
				Alltrim(aFatOutros[nI,2]),;
				Alltrim(aFatOutros[nI,3]),;
				Alltrim(aFatOutros[nI,4]),;
				Transform(aFatOutros[nI,5],"@R 999,999.999"),;
				Alltrim(aFatOutros[nI,6]),;
				Transform(aFatOutros[nI,7],"@R 99,999,999.99")},aL[21],"",,@nLin)
	Next nI
EndIf
FmtLin(,{aL[35],aL[22],aL[36],aL[23],aL[24]},,,@nLin)  
aFatFin := {}
U_FatFin(@aFatFin,nMes)             
If Len(aFatFin) > 0
	For nI := 1 to Len(aFatFin)
		FmtLin({Alltrim(aFatFin[nI,1]),;
				Alltrim(aFatFin[nI,2]),;
				Alltrim(aFatFin[nI,3]),;
				Alltrim(aFatFin[nI,4]),;
				Alltrim(aFatFin[nI,5]),;
				Alltrim(aFatFin[nI,6]),;
				Alltrim(aFatFin[nI,7]),;
				Transform(aFatFin[nI,8],"@R 999,999.999"),;
				Alltrim(aFatFin[nI,9]),;			
				Transform(aFatFin[nI,10],"@R 9999,999,999.99")},aL[25],"",,@nLin)
	Next nI
EndIf
FmtLin(,{aL[26],aL[37],aL[38],aL[27],aL[28]},,,@nLin)  
aFatEst := {}
U_FatEst(@aFatEst,nMes)             
If Len(aFatEst) > 0
	For nI := 1 to len(aFatEst)
		FmtLin({aFatEst[nI,1],;
				aFatEst[nI,2],;
				aFatEst[nI,3],;
				aFatEst[nI,4],;
				aFatEst[nI,5],;
				aFatEst[nI,6],;
				aFatEst[nI,7],;
				Transform(aFatEst[nI,8],"@R 999,999.999"),;
				aFatEst[nI,9],;			
				Transform(aFatEst[nI,10],"@R 9999,999,999.99")},aL[29],"",,@nLin)
	Next nI
EndIf	
FmtLin(,{aL[39],aL[30]},,,@nLin) 
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                
	FmtLin(,{aL[1],aL[2],aL[3],aL[4]},,,@nLin)	
EndIf
U_CRF010G(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,@nLin,nMes)         
nMes :=Month(MV_PAR02)
U_CRF010F(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,@nLin,nMes,aQtdPed,aVlPedLiq)
U_RFinLay2()
nLin := 61
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf
cQuery := " SELECT SUBSTRING(D2_EMISSAO,1,6) AS EMISSAO, "
cQuery += " SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR, SUM(D2_CUSTO1) AS CUSTO, " 
cQuery += " SUM(D2_QUANT) AS QUANT FROM "
cQuery += RetSqlName('SD2') + " D2 WHERE "
cQuery += " D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND D2.D2_TIPO    = 'N' AND D2.D2_LOCAL IN ('03','04','05') AND " 
cQuery += " D2.D2_COD NOT IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += " '057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += " '038675','044926','057558','037503') AND D2.D2_CF IN (" + Alltrim(cCFOPVdInt) + ") AND D2.D_E_L_E_T_ = ' ' AND "
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " GROUP BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("c010BSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2 
DbSelectArea("c010BSD2")
Do While c010BSD2->(!EOF())
	If c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
   		aCstBruto[1,2] := c010BSD2->CUSTO
   		aPrcMedio[1,2] := c010BSD2->VALOR/c010BSD2->QUANT
   		aQtdVend[1,2]  := c010BSD2->QUANT
   	ElseIf c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
   		aCstBruto[2,2] := c010BSD2->CUSTO 
   		aPrcMedio[2,2] := c010BSD2->VALOR/c010BSD2->QUANT
   		aQtdVend[2,2] := c010BSD2->QUANT
   	ElseIf c010BSD2->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
   		aCstBruto[3,2] := c010BSD2->CUSTO 
  		aPrcMedio[3,2] := c010BSD2->VALOR/c010BSD2->QUANT
  		aQtdVend[3,2]  := c010BSD2->QUANT
   	EndIf	
	c010BSD2->(DbSkip())
EndDo
DbSelectArea("c010BSD2")
dbCloseArea() 
cQuery := " SELECT SUBSTRING(F2_EMISSAO,1,6) AS EMISSAO, SUM(F2_VALFAT) AS VALOR, SUM(F2_VALICM) AS ICMS,SUM(F2_VALIPI ) AS IPI FROM "
cQuery += RetSqlName('SF2') + " F2  WHERE F2.F2_FILIAL  = '"+xFilial("SF2")+"' AND "
cQuery += " F2.D_E_L_E_T_ = ' ' AND  SUBSTRING(F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND F2.F2_TIPO    = 'N' AND "          
cQuery += " F2.F2_DOC+F2.F2_SERIE+F2.F2_EMISSAO IN (SELECT  D2.D2_DOC+D2.D2_SERIE+D2.D2_EMISSAO FROM "
cQuery +=   RetSqlName('SD2') + " D2 WHERE  D2.D2_FILIAL = '"+xFilial("SD2")+"' AND D2.D2_LOCAL IN ('03','04','05') AND "
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_CF IN (" + Alltrim(cCFOPVdInt) + ") AND "
cQuery += " D2.D2_COD NOT IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += " '057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += " '038675','044926','057558','037503') AND D2.D_E_L_E_T_ <> '*') "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSD2",.T.,.T.)
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("c010BSD2",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea("c010BSD2")
Do While c010BSD2->(!EOF())
		If c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aFatBruto[1,2] := c010BSD2->VALOR 
       		aFatBruto[1,4] := c010BSD2->ICMS 
      		aFatBruto[1,5] := c010BSD2->IPI
      	ElseIf c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aFatBruto[2,2] := c010BSD2->VALOR   
       		aFatBruto[2,4] := c010BSD2->ICMS
     		aFatBruto[2,5] := c010BSD2->IPI
       	ElseIf c010BSD2->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
       		aFatBruto[3,2] := c010BSD2->VALOR
       		aFatBruto[3,4] := c010BSD2->ICMS       		
     		aFatBruto[3,5] := c010BSD2->IPI
       	EndIf	
	c010BSD2->(DbSkip())
EndDo
DbSelectArea("c010BSD2")
dbCloseArea()
cQuery := " SELECT SUBSTRING(D2_EMISSAO,1,6) AS EMISSAO, SUM(D2_TOTAL+D2_VALIPI+D2_ICMSRET+D2_VALFRE+D2_SEGURO+D2_DESPESA) AS VALOR, SUM(D2_CUSTO1) AS CUSTO, " 
cQuery += " SUM(D2_QUANT) AS QUANT FROM "
cQuery += RetSqlName('SD2') + " D2 "
cQuery += " WHERE D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND  D2.D_E_L_E_T_ = ' ' AND D2.D2_TIPO    = 'N' AND D2.D2_LOCAL IN ('03','04','05') AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " D2.D2_COD NOT IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += " '057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += " '038675','044926','057558','037503') AND "
cQuery += " D2.D2_CF IN (" + Alltrim(cCFOPVdExt) + ") AND D2.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D2.D2_EMISSAO,1,6) "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSD2",.T.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("c010BSD2",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2
DbSelectArea("c010BSD2")
Do While c010BSD2->(!EOF())
	nReg := aScan(aCstBruto,{|x| x[1]==c010BSD2->EMISSAO})
	If nReg > 0
		aCstBruto[nReg,03] := c010BSD2->CUSTO
	Endif	
	nReg := aScan(aPrcMedio,{|x| x[1]==c010BSD2->EMISSAO})
	If nReg > 0
		aPrcMedio[nReg,03] := c010BSD2->VALOR/c010BSD2->QUANT
	EndIf	
	nReg := aScan(aQtdVend,{|x| x[1]==c010BSD2->EMISSAO})
	If nReg > 0
		aQtdVend[nReg,03] := c010BSD2->QUANT
	Endif	
	c010BSD2->(DbSkip())
EndDo                 
DbSelectArea("c010BSD2")
dbCloseArea()
cQuery := " SELECT SUBSTRING(F2_EMISSAO,1,6) AS EMISSAO, SUM(F2_VALFAT) AS VALOR FROM "
cQuery +=   RetSqlName('SF2') + " F2 " 
cQuery += " WHERE F2.F2_FILIAL  = '"+xFilial("SF2")+"' AND F2.D_E_L_E_T_ = ' ' AND "
cQuery += " SUBSTRING(F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " F2.F2_TIPO    = 'N' AND F2.F2_DOC+F2.F2_SERIE+F2.F2_EMISSAO IN (SELECT D2.D2_DOC+D2.D2_SERIE+D2.D2_EMISSAO FROM " 
cQuery +=   RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL = '"+xFilial("SD2")+"' AND  D2.D2_LOCAL IN ('03','04','05') AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += " D2.D2_COD NOT IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += " '057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += " '038675','044926','057558','037503') AND D2.D2_CF IN (" + Alltrim(cCFOPVdExt) + ") AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(F2.F2_EMISSAO,1,6) "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSD2",.T.,.T.)
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("c010BSD2",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea("c010BSD2")
Do While c010BSD2->(!EOF())
		If c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aFatBruto[1,3] := c010BSD2->VALOR
      	ElseIf c010BSD2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aFatBruto[2,3] := c010BSD2->VALOR
       	ElseIf c010BSD2->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
       		aFatBruto[3,3] := c010BSD2->VALOR
       	EndIf
       	c010BSD2->(DbSkip())
EndDo
FmtLin({Transform(aFatBruto[1,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[1,3],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,3],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,3],"@R 99,999,999.99")},aL[09],"",,@nLin)
FmtLin(,{aL[10],aL[11]},,,@nLin)
DbSelectArea("c010BSD2")
dbCloseArea()
cQuery := " SELECT SUBSTRING(F2_EMISSAO,1,6) AS EMISSAO,F2_COND,SUM(F2_VALFAT) AS VALOR FROM "                  
cQuery += RetSqlName('SF2') + " F2 "
cQuery += " WHERE F2.F2_FILIAL = '"+xFilial("SF2")+"'  AND SUBSTRING(F2.F2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(F2.F2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND F2.D_E_L_E_T_ = ' ' "
cQuery += " AND F2.F2_TIPO = 'N'  AND F2.F2_DOC+F2.F2_SERIE IN (SELECT  D2.D2_DOC+D2.D2_SERIE FROM "
cQuery += RetSqlName('SD2') + " D2 WHERE D2.D2_FILIAL = '"+xFilial("SD2")+"' "
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND D2.D2_LOCAL IN ('03','04','05') "
cQuery += " AND D2.D2_COD NOT IN ('038787','057524','041643','038760','038580','037501','051356','038747','038702'," 
cQuery += "'057557','056077','056019','059750','038774','060587','038617','038721','038734','038532','052235',"
cQuery += "'038675','044926','057558','037503') "
cQuery += " AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") "
cQuery += " AND D2.D_E_L_E_T_ = ' ') "
cQuery += " GROUP BY SUBSTRING(F2_EMISSAO,1,6),F2_COND "
cQuery += " ORDER BY SUBSTRING(F2_EMISSAO,1,6),F2_COND ASC "
cQuery :=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSF2",.T.,.T.)
For nSF2 := 1 To len(aStruSF2)
	If aStruSF2[nSF2][2] <> "C" .And. FieldPos(aStruSF2[nSF2][1])<>0
		TcSetField("c010BSF2",aStruSF2[nSF2][1],aStruSF2[nSF2][2],aStruSF2[nSF2][3],aStruSF2[nSF2][4])
	EndIf
Next nSF2
DbSelectArea("c010BSF2")
Do While c010BSF2->(!EOF())
	nReg := aScan(aCondFat1,{|x| x[2]==Posicione("SE4",1,xFilial("SE4")+c010BSF2->F2_COND,"E4_DESCRI")})
	If nReg >0 
		If c010BSF2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aCondFat1[nReg,3] := c010BSF2->VALOR
		ElseIf c010BSF2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aCondFat1[nReg,4] := c010BSF2->VALOR
		ElseIf c010BSF2->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aCondFat1[nReg,5] := c010BSF2->VALOR
		EndIf	
	Else
		If c010BSF2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
			aadd(aCondFat1,{c010BSF2->EMISSAO,Posicione("SE4",1,xFilial("SE4")+c010BSF2->F2_COND,"E4_DESCRI"),c010BSF2->VALOR,0,0}) 
		ElseIf c010BSF2->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
			aadd(aCondFat1,{c010BSF2->EMISSAO,Posicione("SE4",1,xFilial("SE4")+c010BSF2->F2_COND,"E4_DESCRI"),0,c010BSF2->VALOR,0}) 
		ElseIf c010BSF2->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
			aadd(aCondFat1,{c010BSF2->EMISSAO,Posicione("SE4",1,xFilial("SE4")+c010BSF2->F2_COND,"E4_DESCRI"),0,0,c010BSF2->VALOR}) 		
		EndIf	
	Endif
	c010BSF2->(DbSkip())
EndDo
For nI := 1 To len(aCondFat1)
	If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
		FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
		FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
		FmtLin(,{aL[7],aL[8]},,,@nLin)
	EndIf	
	FmtLin({aCondFat1[nI,2],;    //
			Transform(aCondFat1[nI,3],"@R 99,999,999.99"),;
			Transform(aCondFat1[nI,4],"@R 99,999,999.99"),;
			Transform(aCondFat1[nI,5],"@R 99,999,999.99")},aL[12],"",,@nLin)
Next nI         
FmtLin(,{aL[13],aL[14],aL[15]},,,@nLin)
DbSelectArea("c010BSF2")
dbCloseArea()
// Devolu??es pegas de D1
cQuery := "SELECT SUBSTRING(D1_DTDIGIT,1,6) AS EMISSAO, SUM(D1_TOTAL)AS DEVOLUCAO "
cQuery += "FROM "+RetSqlName('SD1')
cQuery += " WHERE "
cQuery += "D1_FILIAL = '"+xFilial("SD1")+"'  "
cQuery += " AND SUBSTRING(D1_DTDIGIT,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(D1_DTDIGIT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND D1_TIPO = 'D' AND "
cQuery += " D1_CF IN ('1201','2201','1202','2202') AND "
cQuery += " D1_NFORI <> ' ' AND "
cQuery += "D_E_L_E_T_ <> '*' "
cQuery += " GROUP BY SUBSTRING(D1_DTDIGIT,1,6) "
cQuery += " ORDER BY SUBSTRING(D1_DTDIGIT,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BDev",.T.,.T.)
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("c010BDev",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea("c010BDev")
Do While c010BDev->(!EOF())                  
	If c010BDev->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
  		aDevInt[1,2] := c010BDev->DEVOLUCAO
   	ElseIf c010BDev->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
   		aDevInt[2,2] := c010BDev->DEVOLUCAO
   	ElseIf c010BDev->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
   		aDevInt[3,2] := c010BDev->DEVOLUCAO
    EndIf
	c010BDev->(DbSkip())
EndDo
DbSelectArea("c010BDev")
DbCloseArea()
cQuery := " SELECT SUBSTRING(D1_DTDIGIT,1,6) AS EMISSAO, SUM(D1_TOTAL)AS DEVOLUCAO "
cQuery += " FROM "+RetSqlName('SD1')
cQuery += " WHERE "
cQuery += " D1_FILIAL = '"+xFilial("SD1")+"'  "
cQuery += " AND SUBSTRING(D1_DTDIGIT,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(D1_DTDIGIT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " and D1_TIPO = 'D' AND "
cQuery += " D1_CF IN ('3201','3202') AND "
cQuery += " D1_NFORI <> ' ' AND "
cQuery += " D_E_L_E_T_ <> '*' "
cQuery += " GROUP BY SUBSTRING(D1_DTDIGIT,1,6) "
cQuery += " ORDER BY SUBSTRING(D1_DTDIGIT,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BDev",.T.,.T.)
For nSD1 := 1 To len(aStruSD1)
	If aStruSD1[nSD1][2] <> "C" .And. FieldPos(aStruSD1[nSD1][1])<>0
		TcSetField("c010BDev",aStruSD1[nSD1][1],aStruSD1[nSD1][2],aStruSD1[nSD1][3],aStruSD1[nSD1][4])
	EndIf
Next nSD1
DbSelectArea("c010BDev")
Do While c010BDev->(!EOF())
	If c010BDev->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
  		aDevExt[1,2] := c010BDev->DEVOLUCAO
   	ElseIf c010BDev->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
   		aDevExt[2,2] := c010BDev->DEVOLUCAO
   	ElseIf c010BDev->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
   		aDevExt[3,2] := c010BDev->DEVOLUCAO
    EndIf
	c010BDev->(DbSkip())
EndDo
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf	                  
//		aL[17]	:=		"|DEVOLUCAO DE VENDAS        :|################|################|################|                          |################|################|################|                                        |" 
FmtLin(,{aL[15]},,,@nLin) 
FmtLin({Transform(aDevInt[1,2],"@R 99,999,999.99"),;
		Transform(aDevInt[2,2],"@R 99,999,999.99"),;
		Transform(aDevInt[3,2],"@R 99,999,999.99"),;
		Transform(aDevExt[1,2],"@R 99,999,999.99"),;
		Transform(aDevExt[2,2],"@R 99,999,999.99"),;
		Transform(aDevExt[3,2],"@R 99,999,999.99")},aL[17],"",,@nLin)
DbSelectArea("c010BDev")
DbCloseArea()  
FmtLin(,{aL[93],aL[18],aL[19]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf	
//		aL[20]	:=		"|FATURAMENTO LIQUIDO        :|################|################|################|                          |################|################|################|                                        |"
FmtLin({Transform(aFatBruto[1,2]-aDevInt[1,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,2]-aDevInt[2,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,2]-aDevInt[3,2],"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99")},aL[20],"",,@nLin)
FmtLin(,{aL[21]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf	          
//aL[22]	:=		"|CUSTO M.P.                 :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aCstBruto[1,2]-aCustoDev[1,2],"@R 99,999,999.99"),;
		Transform(aCstBruto[2,2]-aCustoDev[2,2],"@R 99,999,999.99"),;
		Transform(aCstBruto[3,2]-aCustoDev[3,2],"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99")},aL[22],"",,@nLin)
FmtLin(,{aL[23]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf	
//aL[24]	:=		"|MARGEM                     :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(Iif(((((aFatBruto[1,2]-aDevInt[1,2])/(aCstBruto[1,2]-aCustoDev[1,2]))-1)*100)>0,NoRound((((aFatBruto[1,2]-aDevInt[1,2])/(aCstBruto[1,2]-aCustoDev[1,2]))-1)*100,2),0),"@R 99,999,999.99"),;
		Transform(Iif(((((aFatBruto[2,2]-aDevInt[2,2])/(aCstBruto[2,2]-aCustoDev[2,2]))-1)*100)>0,NoRound((((aFatBruto[2,2]-aDevInt[2,2])/(aCstBruto[2,2]-aCustoDev[2,2]))-1)*100,2),0),"@R 99,999,999.99"),;
		Transform(Iif(((((aFatBruto[3,2]-aDevInt[3,2])/(aCstBruto[3,2]-aCustoDev[3,2]))-1)*100)>0,NoRound((((aFatBruto[3,2]-aDevInt[3,2])/(aCstBruto[3,2]-aCustoDev[3,2]))-1)*100,2),0),"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99")},aL[24],"",,@nLin)
FmtLin(,{aL[25],aL[26],aL[82]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf          
cQuery := " SELECT SUBSTRING(D3_EMISSAO,1,6) AS EMISSAO, SUM(D3_QUANT) AS PRODUZIDO FROM "
cQuery += RetSqlName('SD3')
cQuery += " WHERE D3_FILIAL = '"+xFilial("SD3")+"'  "
cQuery += " AND SUBSTRING(D3_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' " 
cQuery += " AND SUBSTRING(D3_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' "
cQuery += " AND D3_CF = 'PR0' AND D3_TM = '300'  AND D3_LOCAL = '01' AND D3_GRUPO IN ('30','40','50','52','60','70') "
cQuery += " AND D_E_L_E_T_ = ' '  "
cQuery += " GROUP BY SUBSTRING(D3_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(D3_EMISSAO,1,6) ASC "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSD3",.T.,.T.)
For nSD3 := 1 To len(aStruSD3)
	If aStruSD3[nSD3][2] <> "C" .And. FieldPos(aStruSD3[nSD3][1])<>0
		TcSetField("cAliasSD3",aStruSD3[nSD3][1],aStruSD3[nSD3][2],aStruSD3[nSD3][3],aStruSD3[nSD3][4])
	EndIf
Next nSD3
Do While cAliasSD3->(!EOF())
		If cAliasSD3->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
       		aPRODUZIDO[1,2] := cAliasSD3->PRODUZIDO
      	ElseIf cAliasSD3->EMISSAO == Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
       		aPRODUZIDO[2,2] := cAliasSD3->PRODUZIDO
       	ElseIf cAliasSD3->EMISSAO == SUBSTR(DTOS(MV_PAR02),1,6)
       		aPRODUZIDO[3,2] := cAliasSD3->PRODUZIDO
       	EndIf	
	cAliasSD3->(DbSkip())
EndDo
DbSelectArea("cAliasSD3")
dbCloseArea()
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100301','100302','100303','100304','100305','100306','100307','100309','100310','100311','100312','100313','100317','100401','100402','100403','100404','100405','100406','100407','100409','100410',"
cQuery += " '100411','100412','100417','200314','600601','200317','200318','200319','200417','200418','200419','200320','200420','200308','200408','200312','200412','200311','200411','200310','200410',"
cQuery += " '200309','200303','200403') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("cAliasSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("cAliasSE7")
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[05],"",,@nLin)
	FmtLin(,{aL[6],aL[7]},,,@nLin)
EndIf	                                             
If cAliasSE7->(!EOF())
//	  	aL[83]	:=		"|CUSTO MEDIO (Ton)          :|################|################|################|                                                                                                                      |"
//	   	aL[84]	:=		"|  * Producao (Ton)         :|################|################|################|                                                                                                                      |"
	FmtLin({Transform((cAliasSE7->VALOR1/aPRODUZIDO[1,2])+(aCstBruto[1,2]/aQtdVend[1,2]),"@R 99,999,999.99"),;
			Transform((cAliasSE7->VALOR2/aPRODUZIDO[2,2])+(aCstBruto[2,2]/aQtdVend[2,2]),"@R 99,999,999.99"),;
			Transform((cAliasSE7->VALOR3/aPRODUZIDO[3,2])+(aCstBruto[3,2]/aQtdVend[3,2]),"@R 99,999,999.99")},aL[83],"",,@nLin)
	FmtLin({Transform(cAliasSE7->VALOR1/aPRODUZIDO[1,2],"@R 99,999,999.99"),;
		    Transform(cAliasSE7->VALOR2/aPRODUZIDO[2,2],"@R 99,999,999.99"),;
			Transform(cAliasSE7->VALOR3/aPRODUZIDO[3,2],"@R 99,999,999.99")},aL[84],"",,@nLin)
Else
	FmtLin({Transform(aCstBruto[1,2]/aQtdVend[1,2],"@R 99,999,999.99"),;
			Transform(aCstBruto[2,2]/aQtdVend[2,2],"@R 99,999,999.99"),;
			Transform(aCstBruto[3,2]/aQtdVend[3,2],"@R 99,999,999.99")},aL[83],"",,@nLin)
	FmtLin({Transform(0,"@R 99,999,999.99"),;
		    Transform(0,"@R 99,999,999.99"),;
			Transform(0,"@R 99,999,999.99")},aL[84],"",,@nLin)
EndIf
//	   	aL[85]	:=		"|  * Materia Prima (Ton)    :|################|################|################|                                                                                                                      |"
FmtLin({Transform(aCstBruto[1,2]/aQtdVend[1,2],"@R 99,999,999.99"),;
		Transform(aCstBruto[2,2]/aQtdVend[2,2],"@R 99,999,999.99"),;
		Transform(aCstBruto[3,2]/aQtdVend[3,2],"@R 99,999,999.99")},aL[85],"",,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf          
FmtLin(,{aL[86],aL[87],aL[27]},,,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf                  
//		aL[28]	:=		"|PRECO MEDIO (Venda)        :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aPrcMedio[1,2],"@R 99,999,999.99"),;
		Transform(aPrcMedio[2,2],"@R 99,999,999.99"),;
		Transform(aPrcMedio[3,2],"@R 99,999,999.99"),;
		Transform(aPrcMedio[1,3],"@R 99,999,999.99"),;
		Transform(aPrcMedio[2,3],"@R 99,999,999.99"),;
		Transform(aPrcMedio[3,3],"@R 99,999,999.99")},aL[28],"",,@nLin)
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf              
FmtLin(,{aL[29],aL[30],aL[88]},,,@nLin)		
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf              
//		aL[89]	:=		"|MARGEM MEDIA               :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(((aPrcMedio[1,2]/((cAliasSE7->VALOR1/aPRODUZIDO[1,2])+(aCstBruto[1,2]/aQtdVend[1,2])))-1)*100,"@R 99,999,999.99"),;
		Transform(((aPrcMedio[2,2]/((cAliasSE7->VALOR2/aPRODUZIDO[2,2])+(aCstBruto[2,2]/aQtdVend[2,2])))-1)*100,"@R 99,999,999.99"),;
		Transform(((aPrcMedio[3,2]/((cAliasSE7->VALOR3/aPRODUZIDO[3,2])+(aCstBruto[3,2]/aQtdVend[3,2])))-1)*100,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99"),;
		Transform(0,"@R 99,999,999.99")},aL[89],"",,@nLin)
FmtLin(,{aL[90],aL[91],aL[31]},,,@nLin)				
DbSelectArea("cAliasSE7")
DbCloseArea()                          
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ = '100203' AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("c010BSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("c010BSE7")
Do While c010BSE7->(!EOF())       
	aComissao[1,1] := c010BSE7->VALOR1
	aComissao[1,2] := c010BSE7->VALOR2
	aComissao[1,3] := c010BSE7->VALOR3
	c010BSE7->(DbSkip())
EndDo
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)	
EndIf	
//		aL[32]	:=		"|COMISSAO                                                                                                                                                                                              |"		
FmtLin(,{aL[32]},,,@nLin)
//		aL[33]	:=		"|  * Valor R$               :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aComissao[1,1],"@R 99,999,999.99"),;
	    Transform(aComissao[1,2],"@R 99,999,999.99"),;
		Transform(aComissao[1,3],"@R 99,999,999.99")},aL[33],"",,@nLin)
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[6],aL[7]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[9]},,,@nLin)
EndIf	
//		aL[34]	:=		"|  * Percentual (%)         :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform((aComissao[1,1]/(aFatBruto[1,2]-aDevInt[1,2]))*100,"@R 99,999,999.99"),;
	    Transform((aComissao[1,2]/(aFatBruto[2,2]-aDevInt[2,2]))*100,"@R 99,999,999.99"),;
		Transform((aComissao[1,3]/(aFatBruto[3,2]-aDevInt[3,2]))*100,"@R 99,999,999.99")},aL[34],"",,@nLin)
FmtLin(,{aL[35],aL[36],aL[37]},,,@nLin)	
DbSelectArea("c010BSE7")
dbCloseArea()
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATCOMGE) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("c010BSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("c010BSE7")
Do While c010BSE7->(!EOF())
  		aGastosCom[1,1] := c010BSE7->VALOR1
   		aGastosCom[1,2] := c010BSE7->VALOR2
   		aGastosCom[1,3] := c010BSE7->VALOR3
	c010BSE7->(DbSkip())
EndDo
DbSelectArea("c010BSE7")
dbCloseArea()
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN (" + Alltrim(cNATGEEX) + ") AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("c010BSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("c010BSE7")
Do While c010BSE7->(!EOF())
  		aGastosCom[2,1] := c010BSE7->VALOR1
   		aGastosCom[2,2] := c010BSE7->VALOR2
   		aGastosCom[2,3] := c010BSE7->VALOR3
	c010BSE7->(DbSkip())
EndDo
If nLin >= 60
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf	
//		aL[38]	:=		"|GASTOS GERAIS COML - MI    :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aGastosCom[1,1],"@R 99,999,999.99"),;
	    Transform(aGastosCom[1,2],"@R 99,999,999.99"),; 
	    Transform(aGastosCom[1,3],"@R 99,999,999.99"),;
	    Transform(aGastosCom[2,1],"@R 99,999,999.99"),;
	    Transform(aGastosCom[2,2],"@R 99,999,999.99"),;
	    Transform(aGastosCom[2,3],"@R 99,999,999.99")},aL[38],"",,@nLin)
FmtLin(,{aL[39]},,,@nLin)	    
aGastosCom := {}         
aadd(aGastosCom,{0,0,0}) 
DbSelectArea("c010BSE7")
dbCloseArea()
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100201','100202','100204','100206','100207','100208','100209','100210','100211','100213','100214',"
cQuery += "'100215','100216','100222') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("c010BSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("c010BSE7")
Do While c010BSE7->(!EOF())
  		aGastosCom[1,1] := c010BSE7->VALOR1
   		aGastosCom[1,2] := c010BSE7->VALOR2
   		aGastosCom[1,3] := c010BSE7->VALOR3
	c010BSE7->(DbSkip())
EndDo
//		aL[40]	:=		"|  * MAO DE OBRA            :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aGastosCom[1,1],"@R 99,999,999.99"),;
	    Transform(aGastosCom[1,2],"@R 99,999,999.99"),; 
	    Transform(aGastosCom[1,3],"@R 99,999,999.99"),},aL[40],"",,@nLin)
FmtLin(,{aL[41]},,,@nLin)	    
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3],aL[4],aL[5]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[06],"",,@nLin)
	FmtLin(,{aL[7],aL[8]},,,@nLin)
EndIf	    
aGastosCom := {}
aadd(aGastosCom,{0,0,0})  
DbSelectArea("c010BSE7")
dbCloseArea()
cQuery := " SELECT SUM(E7_X_VLU"+StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)+") AS VALOR1 , SUM(E7_X_VLU"+StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)+")  AS VALOR2 ,SUM(E7_X_VLU"+StrZero(Month(MV_PAR02),2) + ")  AS VALOR3 " 
cQuery += " FROM "
cQuery += RetSqlName('SE7') 
cQuery += " WHERE D_E_L_E_T_ <> '*' AND "
cQuery += " E7_FILIAL = '"+xFilial("SE7")+"' AND "        
cQuery += " E7_NATUREZ IN ('100203') AND "
cQuery += " E7_ANO = '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +"' "
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"c010BSE7",.T.,.T.)
For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField("c010BSE7",aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7
DbSelectArea("c010BSE7")
Do While c010BSE7->(!EOF())
  		aGastosCom[1,1] := c010BSE7->VALOR1
   		aGastosCom[1,2] := c010BSE7->VALOR2
   		aGastosCom[1,3] := c010BSE7->VALOR3
	c010BSE7->(DbSkip())
EndDo
//		aL[42]	:=		"|  * COMISSAO               :|################|################|################|                          |################|################|################|                                        |"		
FmtLin({Transform(aGastosCom[1,1],"@R 99,999,999.99"),;
	    Transform(aGastosCom[1,2],"@R 99,999,999.99"),; 
	    Transform(aGastosCom[1,3],"@R 99,999,999.99")},aL[42],"",,@nLin)
FmtLin(,{aL[43]},,,@nLin)	    
DbSelectArea("c010BSE7")
dbCloseArea()
/*
U_CRF010J(@nLin,nMes,Cabec1,Cabec2,Titulo,aQtdVend) 
*/
RestArea(aArea)
Return(.T.)
