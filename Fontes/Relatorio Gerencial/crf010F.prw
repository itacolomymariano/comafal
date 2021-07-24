/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³CRF010  ºAutor  ³Eduardo Zanardo     º Data ³  15/11/06     º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³ RESUMO GERENCIAL                                           º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ AP7                                                        º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/

User Function CRF010F(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo,nLin,nMes,aQtdPed,aVlPedLiq)

Local aPedido   := {}          
Local nQtdAtivo := 0
Local nQtdDest  := 0  
Local cQuery    := ""
Local nSC6      := 0 
Local nSC5      := 0 
Local nSC9      := 0
Local cAliasPED := "cAliasPED"
Local aStruSC5 	:= SC5->(dbStruct())
Local aStruSC6 	:= SC6->(dbStruct())  
Local aStruSC9 	:= SC9->(dbStruct()) 
Local aAliasPED := {}
Local nReg      := 0

U_Pedidos(@aPedido,@nQtdAtivo,nMes)

U_RFinLay9()              

FmtLin(,{aL[1],aL[2],aL[3]},,,@nLin)
FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],;
		aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[04],"",,@nLin)
FmtLin(,{aL[5]},,,@nLin)
FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),5],"@R 999,999,999"),;
			Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),5],"@R 999,999,999"),;
			Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),5],"@R 999,999,999")},aL[6],"",,@nLin)
FmtLin(,{aL[7]},,,@nLin)

FmtLin({Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01"),"@R 999,999,999"),; 
		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01"),"@R 999,999,999"),;
		Transform(U_Ativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01"),"@R 999,999,999")},aL[8],"",,@nLin)

FmtLin(,{aL[9]},,,@nLin)

FmtLin({Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01"),"@R 999,999,999"),;
		Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01"),"@R 999,999,999"),;
		Transform(U_Inativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01"),"@R 999,999,999")},aL[10],"",,@nLin)
FmtLin(,{aL[11]},,,@nLin)
FmtLin(,{aL[12]},,,@nLin)

FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),4],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),4],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),4],"@R 999,999,999")},aL[13],"",,@nLin)
FmtLin(,{aL[14]},,,@nLin)

FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),2],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),2],"@R 999,999,999"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),2],"@R 999,999,999")},aL[15],"",,@nLin)
FmtLin(,{aL[16]},,,@nLin)

FmtLin({Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),3],"@R 99,999,999.99"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),3],"@R 99,999,999.99"),;
		Transform(aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),3],"@R 99,999,999.99")},aL[17],"",,@nLin)
FmtLin(,{aL[18]},,,@nLin)

FmtLin({Transform((aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)}),4]/U_Ativos(SUBSTR(DTOS(MV_PAR02-240),1,6)+"01"))*100,"@R 999.99"),;
		Transform((aPedido[aScan(aPedido,{|x| x[1]==Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)}),4]/U_Ativos(SUBSTR(DTOS(MV_PAR02-210),1,6)+"01"))*100,"@R 999.99"),;
		Transform((aPedido[aScan(aPedido,{|x| x[1]==SUBSTR(DTOS(MV_PAR02),1,6)}),4]/U_Ativos(SUBSTR(DTOS(MV_PAR02-180),1,6)+"01"))*100,"@R 999.99")},aL[19],"",,@nLin)
FmtLin(,{aL[20]},,,@nLin)
FmtLin(,{aL[21]},,,@nLin)

FmtLin({Transform(aQtdPed[1],"@R 999,999,999"),;
		Transform(aQtdPed[2],"@R 999,999,999"),;
		Transform(aQtdPed[3],"@R 999,999,999")},aL[22],"",,@nLin)
FmtLin(,{aL[23]},,,@nLin)

FmtLin({Transform(aVlPedLiq[1],"@R 99,999,999.99"),;
		Transform(aVlPedLiq[2],"@R 99,999,999.99"),;
		Transform(aVlPedLiq[3],"@R 99,999,999.99")},aL[24],"",,@nLin)
FmtLin(,{aL[25]},,,@nLin)

cQuery := " SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt, "
cQuery += " COUNT(DISTINCT C5.C5_NUM) as QtdPED "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += " C6.C6_NUM = C5.C5_NUM AND "    
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " C6.C6_LOCAL IN ('03','04') AND " 
cQuery += " C6.C6_DATFAT = '        ' AND  "
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasPED,.T.,.T.)},"Aguarde...","Processando Pedidos Nao Faturados(mes)...")

For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField(cAliasPED,aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6

For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField(cAliasPED,aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0,0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0,0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0,0})

dbSelectArea(cAliasPED)

while (cAliasPED)->(!eof())
	If (cAliasPED)->MES <= Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,2] := (cAliasPED)->PedBrt
		aAliasPED[1,3] := (cAliasPED)->QtdPED
	EndIf
	If (cAliasPED)->MES <= Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,2] := (cAliasPED)->PedBrt
		aAliasPED[2,3] := (cAliasPED)->QtdPED
	EndIf    
	If (cAliasPED)->MES <= SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,2] := (cAliasPED)->PedBrt
		aAliasPED[3,3] := (cAliasPED)->QtdPED
	EndIf    
	(cAliasPED)->(dbskip())
EndDo                     

DbSelectArea(cAliasPED)
(cAliasPED)->(DbCloseArea())

FmtLin({Transform(aAliasPED[1,3],"@R 999,999,999"),;
		Transform(aAliasPED[2,3],"@R 999,999,999"),;
		Transform(aAliasPED[3,3],"@R 999,999,999")},aL[26],"",,@nLin)
FmtLin(,{aL[27]},,,@nLin)

FmtLin({Transform(aAliasPED[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2],"@R 99,999,999.99")},aL[28],"",,@nLin)
FmtLin(,{aL[29]},,,@nLin)

cQuery := "SELECT SUBSTRING(C9.C9_DATALIB,1,6) as MES, "
cQuery += "SUM(C9.C9_PRCVEN*C9.C9_QTDLIB) as PedBrt "
cQuery += " FROM "
cQuery += RetSqlName('SC9') + " C9 "
cQuery += "WHERE "
cQuery += "C9.C9_FILIAL = '"+xFilial("SC9")+"' AND "
cQuery += "C9.D_E_L_E_T_ = ' ' AND "
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += "SUBSTRING(C9.C9_DATALIB,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C9.C9_NFISCAL <> '      ' AND " 
cQuery += "C9.C9_PEDIDO IN (SELECT C5.C5_NUM "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += "WHERE "
cQuery += "C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += "C5.C5_TIPO = 'N' AND "
cQuery += "C5.D_E_L_E_T_ = ' ' AND "
cQuery += "C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += "C6.C6_NUM = C5.C5_NUM AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) >= '"+ Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4))+ StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND "
cQuery += "SUBSTRING(C5.C5_EMISSAO,1,6) <= '"+ SUBSTR(DTOS(MV_PAR02),1,6) +"' AND "
cQuery += "C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += "C6.C6_LOCAL IN ('03','04') AND "
cQuery += "C6.D_E_L_E_T_ = ' ' "
cQuery += ") "
cQuery += "GROUP BY SUBSTRING(C9.C9_DATALIB,1,6) "
cQuery += "ORDER BY SUBSTRING(C9.C9_DATALIB,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasPED,.T.,.T.)},"Aguarde...","Processando Pedidos Faturados...")

For nSC9 := 1 To len(aStruSC9)
	If aStruSC9[nSC9][2] <> "C" .And. FieldPos(aStruSC9[nSC9][1])<>0
		TcSetField(cAliasPED,aStruSC9[nSC9][1],aStruSC9[nSC9][2],aStruSC9[nSC9][3],aStruSC9[nSC9][4])
	EndIf
Next nSC9

aAliasPED :={}

aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0})

dbSelectArea(cAliasPED)

while (cAliasPED)->(!eof())
	If (cAliasPED)->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,2] := (cAliasPED)->PedBrt
	EndIf
	If (cAliasPED)->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,2] := (cAliasPED)->PedBrt
	EndIf    
	If (cAliasPED)->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,2] := (cAliasPED)->PedBrt
	EndIf    
	(cAliasPED)->(dbskip())
EndDo                     

DbSelectArea(cAliasPED)
(cAliasPED)->(DbCloseArea())

FmtLin({Transform(aAliasPED[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2],"@R 99,999,999.99")},aL[30],"",,@nLin)
FmtLin(,{aL[31]},,,@nLin)

cQuery := " SELECT SUBSTRING(C5.C5_EMISSAO,1,6) as MES, "
cQuery += " SUM(C6.C6_VALOR) as PedBrt "
cQuery += " FROM "
cQuery += RetSqlName('SC5') + " C5 , " + RetSqlName('SC6') + " C6 "
cQuery += " WHERE "
cQuery += " C5.C5_FILIAL = '"+xFilial("SC5")+"' AND "
cQuery += " C5.C5_TIPO = 'N' AND "
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) >= '" + Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2) +"' AND " 
cQuery += " SUBSTRING(C5.C5_EMISSAO,1,6) <= '" + SUBSTR(DTOS(MV_PAR02),1,6) +"' AND " 
cQuery += " C5.D_E_L_E_T_ = ' ' AND "
cQuery += " C6.C6_FILIAL = '"+xFilial("SC6")+"' AND "
cQuery += " C6.C6_NUM = C5.C5_NUM AND "    
cQuery += " C6_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " C6.C6_LOCAL IN ('03','04') AND " 
cQuery += " C6.C6_DATFAT = '        ' AND  "
cQuery += " C6.D_E_L_E_T_ = ' ' "
cQuery += " GROUP BY SUBSTRING(C5.C5_EMISSAO,1,6) "
cQuery += " ORDER BY SUBSTRING(C5.C5_EMISSAO,1,6) ASC"

cQuery:=ChangeQuery(cQuery)

MsAguarde({|| dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasPED,.T.,.T.)},"Aguarde...","Processando Pedidos Nao Faturados(mes)...")

For nSC6 := 1 To len(aStruSC6)
	If aStruSC6[nSC6][2] <> "C" .And. FieldPos(aStruSC6[nSC6][1])<>0
		TcSetField(cAliasPED,aStruSC6[nSC6][1],aStruSC6[nSC6][2],aStruSC6[nSC6][3],aStruSC6[nSC6][4])
	EndIf
Next nSC6

For nSC5 := 1 To len(aStruSC5)
	If aStruSC5[nSC5][2] <> "C" .And. FieldPos(aStruSC5[nSC5][1])<>0
		TcSetField(cAliasPED,aStruSC5[nSC5][1],aStruSC5[nSC5][2],aStruSC5[nSC5][3],aStruSC5[nSC5][4])
	EndIf
Next nSC5

aAliasPED :={}

aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2),0})
aadd(aAliasPED,{Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2),0})
aadd(aAliasPED,{SUBSTR(DTOS(MV_PAR02),1,6),0})

dbSelectArea(cAliasPED)

while (cAliasPED)->(!eof())
	If (cAliasPED)->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2)),2)
		aAliasPED[1,2] := (cAliasPED)->PedBrt
	EndIf
	If (cAliasPED)->MES = Iif((nMes-2 == -1 .and. nMes == 1) .or. (nMes-2 == 0 .and. nMes == 2) ,StrZero(Year(MV_PAR02)-1,4),SUBSTR(DTOS(MV_PAR02),1,4)) + StrZero(Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1),2)
		aAliasPED[2,2] := (cAliasPED)->PedBrt
	EndIf    
	If (cAliasPED)->MES = SUBSTR(DTOS(MV_PAR02),1,6)
		aAliasPED[3,2] := (cAliasPED)->PedBrt
	EndIf    
	(cAliasPED)->(dbskip())
EndDo                     

DbSelectArea(cAliasPED)
(cAliasPED)->(DbCloseArea())

FmtLin({Transform(aAliasPED[1,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[2,2],"@R 99,999,999.99"),;
		Transform(aAliasPED[3,2],"@R 99,999,999.99")},aL[32],"",,@nLin)
FmtLin(,{aL[33]},,,@nLin)


Return(.T.)                          

