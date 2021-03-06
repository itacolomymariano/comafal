#INCLUDE "RWMAKE.CH"
/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?CRF010  ?Autor  ?Eduardo Zanardo     ? Data ?  15/11/06     ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? RESUMO GERENCIAL                                           ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? AP7                                                        ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function CRF010F(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,nLin,nMes,aQtdPed,aVlPedLiq)

Local aArea := GetArea()
Local aPedido   := {}          
Local nQtdAtivo := 0
Local nQtdDest  := 0  
Local cQuery    := ""
Local nSC6      := 0 
Local nSC5      := 0 
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSC6 	:= SC6->(dbStruct())  
Local aAliasPED := {}
Local nReg      := 0
Local aFatMes   := {}
Local aFatMesAnt := {}
Local aPedPendMes := {}
U_Pedidos(@aPedido,@nQtdAtivo,nMes)
U_RFinLay9()              
nLin := 61
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
			aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[6]},,,@nLin)
Endif	
//	  	aL[07]	:=		"|CLIENTES                         :|################|################|################|                                                                                                                |"
//FmtLin({Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-240),1,6),Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2))+(U_Inativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01"))+aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),5],"@R 999,999,999"),; 
//		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-210),1,6),Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)))+(U_Inativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01"))+aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),5],"@R 999,999,999"),;
//		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-180),1,6),SUBSTR(DTOS(MV_PAR02),1,6))+aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),5],"@R 999,999,999")},aL[7],"",,@nLin)

//		aL[08]	:=		"|  * Novos                        :|################|################|################|                                                                                                                |"
FmtLin({    Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),5],"@R 999,999,999"),;
			Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),5],"@R 999,999,999"),;
			Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),5],"@R 999,999,999")},aL[8],"",,@nLin)
//		aL[09]	:=		"|  * Ativos                       :|################|################|################|                                                                                                                |"
FmtLin({Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-240),1,6),"200707"),"@R 999,999,999"),; 
		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-210),1,6),"200708"),"@R 999,999,999"),;
		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-180),1,6),"200709"),"@R 999,999,999")},aL[9],"",,@nLin)
FmtLin({Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01"),"@R 999,999,999"),;
		Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01"),"@R 999,999,999"),;
		Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01"),"@R 999,999,999")},aL[10],"",,@nLin)
FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),4],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),4],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),4],"@R 999,999,999")},aL[11],"",,@nLin)
FmtLin({Transform((aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),4]/(U_Ativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01")+U_Inativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01")))*100,"@R 999.99"),;
		Transform((aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),4]/(U_Ativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01")+U_Inativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01")))*100,"@R 999.99"),;
		Transform((aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),4]/(U_Ativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01")+U_Inativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01")))*100,"@R 999.99")},aL[12],"",,@nLin)
FmtLin(,{aL[13],aL[14],aL[15],aL[16]},,,@nLin)
FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),3],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),3],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),3],"@R 999,999,999")},aL[17],"",,@nLin)
FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),2],"@R 99,999,999.99"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),2],"@R 99,999,999.99"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),2],"@R 99,999,999.99")},aL[18],"",,@nLin)
FmtLin(,{aL[19],aL[20],aL[21]},,,@nLin)		
cQuery := " SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt, "
cQuery += " COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "            
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C5.C5_FILIAL = C6.C6_FILIAL AND "
cQuery += " C5.C5_NUM IN (SELECT D2.D2_PEDIDO FROM "
cQuery +=  RetSqlName('SD2') + " D2 " 
cQuery += "WHERE D2.D2_TIPO    = 'N' AND D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND "
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND D2.D_E_L_E_T_ = ' ') AND "
cQuery += " C5.C5_NUM = C6.C6_NUM AND "    
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) = SUBSTRING(C6.C6_DATFAT,1,6) AND "
cQuery += " SUBSTRING(C6.C6_DATFAT,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(C6.C6_DATFAT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " C6.C6_LOCAL IN ('03','04','05') AND " 
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
aAliasPED := {}
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
aPedPendMes := {}
aadd(aPedPendMes,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aPedPendMes,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aPedPendMes,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
DbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,2] += cAliasPED->PedBrt
		aAliasPED[1,3] += cAliasPED->QtdPED
	EndIf
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,2] += cAliasPED->PedBrt
		aAliasPED[2,3] += cAliasPED->QtdPED
	EndIf    
	If cAliasPED->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,2] += cAliasPED->PedBrt
		aAliasPED[3,3] += cAliasPED->QtdPED
	EndIf    
	cAliasPED->(dbskip())
EndDo                    
//        aL[21]	:=		"|PEDIDOS FATURADOS                                                                                                                                                                                     |"
//        aL[22]	:=		"|  * Quantidade (Mes)             :|################|################|################|                                                                                                                |"
//        aL[23]	:=		"|  * Valor R$ (Mes)               :|################|################|################|                                                                                                                |"
FmtLin({Transform(aAliasPED[1,3],"@R 999,999,999"),;
		Transform(aAliasPED[2,3],"@R 999,999,999"),;
		Transform(aAliasPED[3,3],"@R 999,999,999")},aL[22],"",,@nLin)
FmtLin({Transform(aAliasPED[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2],"@R 99,999,999.99")},aL[23],"",,@nLin)
aFatMes	:= aAliasPED
FmtLin(,{aL[24],aL[25]},,,@nLin)		                                                                                            
dbSelectArea("cAliasPED")
DbCloseArea()
aAliasPED := {}
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
cQuery := " SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt, "
cQuery += " COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "            
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) +StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C5.C5_NUM NOT IN (SELECT D2.D2_PEDIDO FROM "
cQuery +=  RetSqlName('SD2') + " D2 " 
cQuery += "WHERE D2.D2_TIPO    = 'N' AND D2.D2_FILIAL  = '"+xFilial("SD2")+"' AND "
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += "SUBSTRING(D2.D2_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND D2.D_E_L_E_T_ = ' ') AND "
cQuery += " C5.C5_FILIAL = C6.C6_FILIAL AND "
cQuery += " C5.C5_NUM = C6.C6_NUM AND "    
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND " 
cQuery += " C6.C6_NOTA = '         ' AND " 
cQuery += " SUBSTRING(C6.C6_DATFAT,1,6) = '        ' AND " 
cQuery += " C6.C6_LOCAL IN ('03','04','05') AND " 
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,3] += cAliasPED->PedBrt
		aAliasPED[1,2] += cAliasPED->QtdPED
	EndIf
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,3] += cAliasPED->PedBrt
		aAliasPED[2,2] += cAliasPED->QtdPED
	EndIf    
	If cAliasPED->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,3] += cAliasPED->PedBrt
		aAliasPED[3,2] += cAliasPED->QtdPED
	EndIf 
	cAliasPED->(dbskip())   
EndDo
//   	aL[26]	:=		"|PEDIDOS PENDENTES                                                                                                                                                                                     |"
//      aL[27]	:=		"|  * Quantidade (Mes)             :|################|################|################|                                                                                                                |"
//	 	aL[28]	:=		"|  * Valor R$ (Mes)               :|################|################|################|                                                                                                                |"
FmtLin(,{aL[26]},,,@nLin)		                                                                                            
FmtLin({Transform(aAliasPED[1,2],"@R 999,999,999"),;
		Transform(aAliasPED[2,2],"@R 999,999,999"),;
		Transform(aAliasPED[3,2],"@R 999,999,999")},aL[27],"",,@nLin)
FmtLin({Transform(aAliasPED[1,3],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,3],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,3],"@R 99,999,999.99")},aL[28],"",,@nLin)
		aPedPendMes[1,3] := aAliasPED[1,2]
		aPedPendMes[1,2] := aAliasPED[1,3]
		aPedPendMes[2,3] := aAliasPED[2,2]
		aPedPendMes[2,2] := aAliasPED[2,3]
		aPedPendMes[3,3] := aAliasPED[3,2]
		aPedPendMes[3,2] := aAliasPED[3,3]
FmtLin(,{aL[29],aL[30],aL[31],aL[32]},,,@nLin)		
dbSelectArea("cAliasPED")
DbCloseArea()
aAliasPED := {}
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
cQuery := " SELECT SUBSTRING(C6.C6_DATFAT,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt, "
cQuery += " COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "            
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) < '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C5.C5_FILIAL = C6.C6_FILIAL AND "
cQuery += " C5.C5_NUM = C6.C6_NUM AND "    
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) < SUBSTRING(C6.C6_DATFAT,1,6) AND "
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND " 
cQuery += " C6.C6_NOTA <> '         ' AND " 
cQuery += " SUBSTRING(C6.C6_DATFAT,1,6) <> '      ' AND " 
cQuery += " SUBSTRING(C6.C6_DATFAT,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C6.C6_LOCAL IN ('03','04','05') AND " 
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C6.C6_DATFAT,1,6) "
cQuery += " ORDER BY SUBSTRING(C6.C6_DATFAT,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
dbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,3] += cAliasPED->PedBrt
		aAliasPED[1,2] += cAliasPED->QtdPED
	EndIf
	If cAliasPED->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,3] += cAliasPED->PedBrt
		aAliasPED[2,2] += cAliasPED->QtdPED
	EndIf    
	If cAliasPED->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,3] += cAliasPED->PedBrt
		aAliasPED[3,2] += cAliasPED->QtdPED
	EndIf 
	cAliasPED->(dbskip())   
EndDo
//	   	aL[32]	:=		"|PEDIDOS FATURADOS                                                                                                                                                                                     |"
//      aL[33]	:=		"|  * Quantidade (Meses Ant.)      :|################|################|################|                                                                                                                |"
//		aL[34]	:=		"|  * Valor R$ (Meses Ant.)        :|################|################|################|                                                                                                                |"
FmtLin({Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),2],"@R 999,999,999"),;
		Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),2],"@R 999,999,999"),;
		Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),2],"@R 999,999,999")},aL[33],"",,@nLin)
FmtLin({Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),3],"@R 99,999,999.99"),;
		Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),3],"@R 99,999,999.99"),;
		Transform(aAliasPED[aScan(aAliasPED,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),3],"@R 99,999,999.99")},aL[34],"",,@nLin)
aFatMesAnt := aAliasPED

FmtLin(,{aL[35],aL[36],aL[37]},,,@nLin)		
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
			aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[6]},,,@nLin)
Endif
dbSelectArea("cAliasPED")
DbCloseArea()
cQuery := " SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt, "
cQuery += " COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "            
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) < '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += " C6.C6_NUM = C5.C5_NUM AND "    
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " C6.C6_LOCAL IN ('03','04','05') AND " 
cQuery += " C6.C6_NOTA = '         ' AND  "
cQuery += " C6.C6_DATFAT = '        ' AND  "
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"
cQuery:=ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"cAliasPED",.T.,.T.)
For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField("cAliasPED",aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6
For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField("cAliasPED",aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5
aAliasPED := {}
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})
DbSelectArea("cAliasPED")
while cAliasPED->(!eof())
	If cAliasPED->MES < Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,2] += cAliasPED->PedBrt
		aAliasPED[1,3] += cAliasPED->QtdPED
	EndIf
	If cAliasPED->MES < Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,2] += cAliasPED->PedBrt
		aAliasPED[2,3] += cAliasPED->QtdPED
	EndIf    
	If cAliasPED->MES < SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,2] += cAliasPED->PedBrt
		aAliasPED[3,3] += cAliasPED->QtdPED
	EndIf    
	cAliasPED->(dbskip())
EndDo
//     	aL[37]	:=		"|PEDIDOS PENDENTES                                                                                                                                                                                     |"
//      aL[38]	:=		"|  * Quantidade (Meses Ant.)      :|################|################|################|                                                                                                                |"
//		aL[39]	:=		"|  * Valor R$ (Meses Ant.)        :|################|################|################|                                                                                                                |"
FmtLin({Transform(aAliasPED[1,3],"@R 999,999,999"),;
		Transform(aAliasPED[2,3],"@R 999,999,999"),;
		Transform(aAliasPED[3,3],"@R 999,999,999")},aL[38],"",,@nLin)
FmtLin({Transform(aAliasPED[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2],"@R 99,999,999.99")},aL[39],"",,@nLin)
FmtLin(,{aL[29]},,,@nLin)                   
FmtLin(,{aL[40],aL[41],aL[42]},,,@nLin)		                           
If nLin >= 60
	Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
	nLin:=PROW()+1                                    
	FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
	FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
			aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
	FmtLin(,{aL[5],aL[6]},,,@nLin)
Endif
//	    aL[48]	:=		"|QTD. TOTAL PEDIDOS FATURADOS     :|################|################|################|                                                                                                                |"
//		aL[43]	:=		"|VALOR TOTAL PEDIDOS FATURADOS    :|################|################|################|                                                                                                                |"
FmtLin({Transform(aFatMes[1,3]+aFatMesAnt[1,2],"@R 999,999,999"),;
		Transform(aFatMes[2,3]+aFatMesAnt[2,2],"@R 999,999,999"),;
		Transform(aFatMes[3,3]+aFatMesAnt[3,2],"@R 999,999,999")},aL[48],"",,@nLin)
FmtLin({Transform(aFatMes[1,2]+aFatMesAnt[1,3],"@R 99,999,999.99"),;
		Transform(aFatMes[2,2]+aFatMesAnt[2,3],"@R 99,999,999.99"),;
		Transform(aFatMes[3,2]+aFatMesAnt[3,3],"@R 99,999,999.99")},aL[43],"",,@nLin)
FmtLin(,{aL[44]},,,@nLin)                                              
//	    aL[49]	:=		"|QTD. TOTAL PEDIDOS PENDENTES     :|################|################|################|                                                                                                                |"
//		aL[45]	:=		"|VALOR TOTAL DE PEDIDOS PENDENTES :|################|################|################|                                                                                                                |"
FmtLin({Transform(aAliasPED[1,3]+aPedPendMes[1,3],"@R 999,999,999"),;
		Transform(aAliasPED[2,3]+aPedPendMes[2,3],"@R 999,999,999"),;
		Transform(aAliasPED[3,3]+aPedPendMes[3,3],"@R 999,999,999")},aL[49],"",,@nLin)
FmtLin({Transform(aAliasPED[1,2]+aPedPendMes[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2]+aPedPendMes[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2]+aPedPendMes[3,2],"@R 99,999,999.99")},aL[45],"",,@nLin)
FmtLin(,{aL[46],aL[47]},,,@nLin)
DbSelectArea("cAliasPED")
DbCloseArea()

RestArea(aArea)

nlin:= 61

Return(.T.)                          
