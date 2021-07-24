#Include "PROTHEUS.CH"
#Include "TopConn.Ch"

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCMB007    บ Autor ณ Zanardo            บ Data ณ  22/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณAjuste dos Orcamentos - SE7                                 บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/

User Function CMB007(nMes)

Local cPerg := "CMB007"

If nMes == Nil .or. nMes <= 0   
	AjustaSX1(cPerg)
	Pergunte(cPerg,.T.)
Endif	
MsgRun("Aguarde...Ajustando Orcamentos...",,{|| CMB007A(nMes)})

Return(.T.)

/*ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณCMB007A    บ Autor ณ Zanardo            บ Data ณ  22/01/04  บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDescricao ณProcessamento do Dados                                      บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿*/

Static Function CMB007A(nMes)

Local cQuery    	:= "" 
Local cAliasSE7 	:= "cAliasSE7"
Local nSE7      	:= 0
Local aStruSE7  	:= SE7->(dbStruct())
Local cAliasSC7 	:= "cAliasSC7"
Local nSC7      	:= 0
Local aStruSC7  	:= SC7->(dbStruct())
Local cAliasSE2 	:= "cAliasSE2"
Local nSE2      	:= 0
Local aStruSE2		:= SE2->(dbStruct())
Local nSE5      	:= 0
Local aStruSE5		:= SE5->(dbStruct())
Local cAliasSE5 	:= "cAliasSE5"
Local nReservado 	:= 0
Local nUtilizado    := 0
Local nOutros		:= 0
Local cMes			:= ""
Local nValPac		:= 0 
Local cMesSC7 		:= ""
Local cAnoSC7		:= ""


If nMes <> Nil .and. nMes > 0   
	cMes			:= StrZero(nMes,2)
	cQuery := "Select E7_ANO, E7_NATUREZ "
	cQuery += "From "
	cQuery += RetSqlName('SE7')
	cQuery += " Where "
	cQuery += "E7_FILIAL = '" + xFilial("SE7")+"'	AND "
	cQuery += "E7_ANO = '"+ SUBSTR(DTOS(MV_PAR02),1,4) + "' AND "
	cQuery += "D_E_L_E_T_ <> '*' "             
Else 
	cMes			:= StrZero(MV_PAR02,2) 
	cQuery := "Select E7_ANO, E7_NATUREZ "
	cQuery += "From "
	cQuery += RetSqlName('SE7')
	cQuery += " Where "
	cQuery += "E7_FILIAL = '" + xFilial("SE7")+"'	AND "
	cQuery += "E7_ANO = '"+ MV_PAR01 + "' AND "
	cQuery += "E7_NATUREZ >= '"+ MV_PAR03 + "' AND "
	cQuery += "E7_NATUREZ <= '"+ MV_PAR04 + "' AND "
	cQuery += "D_E_L_E_T_ <> '*' "
EndIf
cQuery:=ChangeQuery(cQuery)
	
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE7,.T.,.T.)

For nSE7 := 1 To len(aStruSE7)
	If aStruSE7[nSE7][2] <> "C" .And. FieldPos(aStruSE7[nSE7][1])<>0
		TcSetField(cAliasSE7,aStruSE7[nSE7][1],aStruSE7[nSE7][2],aStruSE7[nSE7][3],aStruSE7[nSE7][4])
	EndIf
Next nSE7

DbSelectArea(cAliasSE7)   

Do While (cAliasSE7)->(!EOF())

	nReservado 	:= 0
	nUtilizado  := 0
    nOutros		:= 0
    If nMes == Nil .or. nMes <= 0   	 
		// Valor de Outros Meses - titulos vencidos em outros meses mas compensados no mes em questao
		cQuery := "SELECT E5.E5_NATUREZ, SUBSTRING(E5.E5_DATA,1,6) AS MESANO, SUM(E5.E5_VALOR) AS OUTROS_MESES "
		cQuery += "FROM " + RetSqlName('SE5') + " E5, " + RetSqlName('SE2') + " E2 "
		cQuery += "WHERE "
		cQuery += "E5.E5_FILIAL = '" + xFilial('SE5') + "' "	// Filial do SE5 - Movimentacoes bancarias
		cQuery += "AND E5.E5_PREFIXO = E2.E2_PREFIXO "			// Prefixos iguais
		cQuery += "AND E5.E5_NUMERO = E2.E2_NUM "					// Numeros iguais
		cQuery += "AND E5.E5_PARCELA = E2.E2_PARCELA "				// Parcelas iguais
		cQuery += "AND E5.E5_CLIFOR = E2.E2_FORNECE "				// Fornecedores iguais
		cQuery += "AND E5.E5_LOJA = E2.E2_LOJA "					// Lojas iguais
		cQuery += "AND E5.E5_NATUREZ = '" +  (cAliasSE7)->E7_NATUREZ + "' "		// Natureza
		cQuery += "AND SUBSTRING(E5.E5_DATA,1,6) = '" + Alltrim(MV_PAR01)+cMes + "' "	
		cQuery += "AND E5.D_E_L_E_T_ <> '*' "						// Registros Ativos do SE5
		cQuery += "AND SUBSTRING(E2.E2_VENCREA,5,2) < SUBSTRING(E5.E5_DATA,5,2) "					// Mes anterior
		cQuery += "AND SUBSTRING(E2.E2_VENCREA,1,4) <= SUBSTRING(E5.E5_DATA,1,4) "					// Mesmo ano ou anterior
		cQuery += "AND E2.E2_FILIAL = '" + xFilial('SE2') + "' "	// Filial do SE2 - Contas a Pagar
		cQuery += "AND E2.E2_X_STAT = ' ' "
		cQuery += "AND E2.D_E_L_E_T_ <> '*' "						// Registros Ativos do SE2
		cQuery += "GROUP BY E5.E5_NATUREZ, SUBSTRING(E5.E5_DATA,1,6)"
			
		cQuery := ChangeQuery(cQuery)
		
		dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE5,.T.,.T.)
		
		For nSE5 := 1 To Len(aStruSE5)
			If aStruSE5[nSE5][2] <> "C" .AND. FieldPos(aStruSE5[nSE5][1]) <> 0
				TcSetField(cAliasSE5,aStruSE5[nSE5][1],aStruSE5[nSE5][2],aStruSE5[nSE5][3],aStruSE5[nSE5][4])
			EndIf
		EndFor
		
		nOutros := (cAliasSE5)->OUTROS_MESES		// Valor de outros Meses
			
		(cAliasSE5)->(dbCloseArea())
	
	
		// Valor a Realizar - Valores que nao foram compensados de titulos que vencem no mes em questao
		cQuery := "SELECT E2.E2_NATUREZ, SUBSTRING(E2.E2_VENCREA,1,6) AS MESANO, "
		cQuery += "SUM(E2.E2_SALDO) AS A_REALIZAR FROM " + RetSqlName('SE2') + " E2 "
		cQuery += "WHERE E2.E2_FILIAL = '" + xFilial('SE2') + "' " 
		cQuery += "AND E2.E2_TIPO <> 'AB-' "					// Abatimento
		cQuery += "AND E2.E2_TIPO <> 'NDF' "					// Nota Debito Fornecedor
		cQuery += "AND E2.E2_NATUREZ = '" + (cAliasSE7)->E7_NATUREZ + "' "	// Natureza
		cQuery += "AND SUBSTRING(E2.E2_VENCREA,1,6) = '" + Alltrim(MV_PAR01)+cMes + "' "	
		cQuery += "AND E2.E2_SALDO > 0 "						// Valor ainda a compensar
		cQuery += "AND E2.E2_X_STAT = ' ' "
		cQuery += "AND E2.D_E_L_E_T_ <> '*' "					// Registros ativos
		cQuery += "GROUP BY E2.E2_NATUREZ, SUBSTRING(E2.E2_VENCREA,1,6)"
		cQuery := ChangeQuery(cQuery)
		dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE2,.T.,.T.)
		
		For nSE2 := 1 To len(aStruSE2)
			If aStruSE2[nSE2][2] <> "C" .And. FieldPos(aStruSE2[nSE2][1])<>0
				TcSetField(cAliasSE2,aStruSE2[nSE2][1],aStruSE2[nSE2][2],aStruSE2[nSE2][3],aStruSE2[nSE2][4])
			EndIf
		Next nSE2
		
		DbSelectArea(cAliasSE2)   	
		nReservado := (cAliasSE2)->A_REALIZAR		// Valor a Realizar
		(cAliasSE2)->(DbCloseArea())	
	
		cQuery := "SELECT C7.C7_TOTAL,C7.C7_VALIPI,C7.C7_VALFRE,C7.C7_SEGURO, "
		cQuery += "C7.C7_DESPESA,C7.C7_COND,C7.C7_DATPRF "
		cQuery += "FROM "
		cQuery += RetSqlName('SC7') + " C7 "  
		cQuery += "WHERE "
		cQuery += "C7.C7_FILIAL = '" + xFilial('SC7') + "' "
		cQuery += "AND C7.C7_X_STAT = ' ' "
		cQuery += "AND C7.C7_X_NATUR = '" + (cAliasSE7)->E7_NATUREZ + "' "
		cQuery += "AND C7.C7_TES IN "
		cQuery += "(SELECT F4.F4_CODIGO FROM "
		cQuery += RetSqlName('SF4') + " F4 " 
		cQuery += "WHERE "
		cQuery += "F4_FILIAL = '" + xFilial("SF4")+"' "
		cQuery += "AND F4_DUPLIC = 'S' "	 
		cQuery += "AND D_E_L_E_T_ <> '*' ) "	
		cQuery += "AND C7.C7_NUM NOT IN "
		cQuery += "(SELECT D1_PEDIDO FROM "
		cQuery += RetSqlName('SD1') + " D1 " 
		cQuery += "WHERE D1_FILIAL = '" + xFilial('SD1') + "' "
		cQuery += "AND D_E_L_E_T_ <> '*' ) "
		cQuery += "AND C7.D_E_L_E_T_ <> '*' "
	
		cQuery := ChangeQuery(cQuery)
	
		dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSC7,.T.,.T.)
		
		For nSC7 := 1 To len(aStruSC7)
			If aStruSC7[nSC7][2] <> "C" .And. FieldPos(aStruSC7[nSC7][1])<>0
				TcSetField(cAliasSC7,aStruSC7[nSC7][1],aStruSC7[nSC7][2],aStruSC7[nSC7][3],aStruSC7[nSC7][4])
			EndIf
		Next nSC7
		
		DbSelectArea(cAliasSC7)
		Do While (cAliasSC7)->(!EOF())
		    aParcelas 	:= Condicao((cAliasSC7)->C7_TOTAL+(cAliasSC7)->C7_VALIPI+;
		    						(cAliasSC7)->C7_VALFRE+(cAliasSC7)->C7_SEGURO+;
		    						(cAliasSC7)->C7_DESPESA,(cAliasSC7)->C7_COND,,(cAliasSC7)->C7_DATPRF)	// Parcelas        
		    						
			For nI := 1 To Len(aParcelas)	
				dData	:=	aParcelas[nI,1]				// Data projetada da parcela
				nValPac	:=	aParcelas[nI,2]				// valor projetado da parcela
				cMesSC7 :=	StrZero(Month(dData),2)		// Mes projetado da parcela
				cAnoSC7	:=	StrZero(Year(dData),4)		// Ano da(s) parcela(s)
				If cMesSC7 == StrZero(MV_PAR02,2) .And. cAnoSC7 == MV_PAR01 
					nReservado += nValPac
				EndIf
			Next nI       
			DbSelectArea(cAliasSC7)
			(cAliasSC7)->(DbSkip()) 
		
		EndDo
		(cAliasSC7)->(DbCloseArea()) 	
		DbSelectArea("SE7")
	    DbSetOrder(1)
	    DbSeek(xFilial("SE7")+(cAliasSE7)->E7_NATUREZ+(cAliasSE7)->E7_ANO)
	    RecLock("SE7",.F.) 
		&("SE7->E7_X_RES" + cMes) := nReservado
		&("SE7->E7_X_OUT" + cMes) := nOutros
    	MsUnlock()
    EndIf
    
	nUtilizado := U_RetOrcUtz((cAliasSE7)->E7_NATUREZ,cMes,nMes)

   	DbSelectArea("SE7")
    DbSetOrder(1)
    DbSeek(xFilial("SE7")+(cAliasSE7)->E7_NATUREZ+(cAliasSE7)->E7_ANO)
    RecLock("SE7",.F.) 
	&("SE7->E7_X_VLU" + cMes) := nUtilizado		    
	MsUnlock()
    
    DbSelectArea(cAliasSE7)
	(cAliasSE7)->(DbSkip())

EndDo                       

(cAliasSE7)->(DbCloseArea())	

Return(.T.)
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณAJUSTASX1 บAutor  ณEduardo Zanardo     บ Data ณ  22/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณAjusta o SX1                                                บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ CMB007                                                     บฑฑ              
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/
Static Function AjustaSx1(cPerg)

Local aArea := GetArea()

cPerg := Padr(cPerg,6)

//PutSx1(cGrupo,cOrdem,cPergunt       ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3 ,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03,"","",cDef04,"","",cDef05,"","",aHelpPor,"","",cHelp)
  PutSx1(cPerg ,"01"  ,"Ano ?"        ,"","","mv_ch1","C"  ,04      ,0       ,0      ,"G"  ,""  ,""   ,"","","MV_PAR01","","","","","","","","","","","","","","","","")  
  PutSx1(cPerg ,"02"  ,"Mes ?"        ,"","","mv_ch2","N"  ,02      ,0       ,0      ,"G"  ,""  ,""   ,"","","MV_PAR02","","","","","","","","","","","","","","","","")    
  PutSx1(cPerg ,"03" ,"Natureza de ?" ,"","","mv_ch3","C"  ,10      ,0       ,0      ,"G"  ,""  ,"SED","","","MV_PAR03","","","","","","","","","","","","","","","","")
  PutSx1(cPerg ,"04" ,"Natureza ate ?","","","mv_ch4","C"  ,10      ,0       ,0      ,"G"  ,""  ,"SED","","","MV_PAR04","","","","","","","","","","","","","","","","")
				
RestArea(aArea)

Return(.T.)

/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบFunction  ณRetOrcUtz บAutor  ณEduardo Zanardo     บ Data ณ  22/01/04   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบDesc.     ณRetorna o valor Utilizado do Orcamento                      บฑฑ
ฑฑฬออออออออออุออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออนฑฑ
ฑฑบUso       ณ CMB007                                                     บฑฑ              
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/                            

User Function RetOrcUtz(cNat,cMes,nMes)

Local aArea			:= GetArea()
Local cQuery 		:= ""
Local nSE5      	:= 0
Local aStruSE5		:= SE5->(dbStruct())
Local cAliasSE5 	:= "cAliasSE5"
Local lManual 		:= .F.
Local cCarteira 	:= ""
Local nValor 		:= 0
Local nVlMovFin 	:= 0
Local cTipodoc   	:= ""
Local cTipo190		:= ""
Local nRecno 		:= 0
Local nAbat 		:= 0
Local nTotBaixado	:= 0
Local nTotAbat   	:= 0
Local nTotValor  	:= 0
Local nTotMovFin 	:= 0
Local nTotComp	  	:= 0

Private cCond3		:= ""
Private cNumero    	:= ""
Private cPrefixo   	:= ""
Private cParcela   	:= ""
Private dBaixa     	:= CTOD("  /  /    ")
Private cBanco     	:= ""
Private cLoja      	:= ""
Private cSeq       	:= ""
Private cNumCheq   	:= ""
Private cRecPag		:= ""
Private cMotBaixa	:= ""
Private cCheque    	:= ""
Private cTipo      	:= ""
Private cFornece   	:= ""
Private dDigit     	:= ""
Private lBxTit	  	:= .F.
Private cFilorig 		:= ""



cQuery := "SELECT E5_FILIAL,E5_DATA,E5_TIPO,E5_MOEDA,E5_VALOR,E5_NATUREZ,E5_BANCO,E5_AGENCIA,"
cQuery += "E5_CONTA,E5_NUMCHEQ,E5_VENCTO,E5_RECPAG,E5_BENEF,E5_HISTOR,"
cQuery += "E5_TIPODOC,E5_SITUACA,E5_PREFIXO,E5_NUMERO,"
cQuery += "E5_PARCELA,E5_CLIFOR,E5_LOJA,E5_DTDIGIT,"
cQuery += "E5_MOTBX,E5_SEQ,"
cQuery += "E5_FILORIG,E5_VLJUROS,"
cQuery += "E5_VLMULTA,E5_SITUA,E5_ITEMD,E5_ITEMC,"
cQuery += "R_E_C_N_O_ SE5RECNO "
cQuery += "FROM " + RetSqlName('SE5') + " E5 "
cQuery += "WHERE E5_FILIAL = '" + xFilial('SE5') + "' "	
cQuery += "AND D_E_L_E_T_ <> '*' "						
cQuery += "AND E5_NATUREZ = '" + cNat + "' "		
cQuery += "AND E5_RECPAG = 'P' "					
cQuery += "AND E5_TIPO <> 'PA' "					
cQuery += "AND E5_SITUACA NOT IN ('C','E','X') " 
cQuery += "AND E5_TIPODOC NOT IN ('DC','D2','JR','J2','TL','MT','M2','CM','C2','TR','TE') "
If nMes <> Nil .and. nMes > 0   
	cQuery += "AND SUBSTRING(E5.E5_DATA,1,6) = '" + SUBSTR(DTOS(MV_PAR02),1,4)+cMes + "' "		
Else	
	cQuery += "AND SUBSTRING(E5.E5_DATA,1,6) = '" + Alltrim(MV_PAR01)+cMes + "' "		
Endif
cQuery += "ORDER BY E5_FILIAL,E5_NATUREZ,E5_PREFIXO,E5_NUMERO,E5_PARCELA,E5_TIPO,E5_DTDIGIT,E5_RECPAG,E5_CLIFOR,E5_LOJA"
		
cQuery := ChangeQuery(cQuery)

dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasSE5,.T.,.T.)

For nSE5 := 1 To Len(aStruSE5)
	If aStruSE5[nSE5][2] <> "C" .AND. FieldPos(aStruSE5[nSE5][1]) <> 0
		TcSetField(cAliasSE5,aStruSE5[nSE5][1],aStruSE5[nSE5][2],aStruSE5[nSE5][3],aStruSE5[nSE5][4])
	EndIf
Next nSE5
	
DbSelectArea(cAliasSE5)

While (cAliasSE5)->(!EOF())

	lManual := .f.
	If (Empty((cAliasSE5)->E5_TIPODOC) ) .Or.(Empty((cAliasSE5)->E5_NUMERO))
		lManual := .t.
	EndIf
	
	If !CMB007Filt(cAliasSE5)
		DbSelectArea(cAliasSE5)
		(cAliasSE5)->(dbSkip())
		Loop
	Endif	                   
	
	cNumero    	:= (cAliasSE5)->E5_NUMERO
	cPrefixo   	:= (cAliasSE5)->E5_PREFIXO
	cParcela   	:= (cAliasSE5)->E5_PARCELA
	dBaixa     	:= (cAliasSE5)->E5_DATA
	cBanco     	:= (cAliasSE5)->E5_BANCO
	cLoja      	:= (cAliasSE5)->E5_LOJA
	cSeq       	:= (cAliasSE5)->E5_SEQ
	cNumCheq   	:= (cAliasSE5)->E5_NUMCHEQ
	cRecPag 	:= (cAliasSE5)->E5_RECPAG
	cMotBaixa	:= (cAliasSE5)->E5_MOTBX
	cCheque    	:= (cAliasSE5)->E5_NUMCHEQ
	cTipo      	:= (cAliasSE5)->E5_TIPO
	cFornece   	:= (cAliasSE5)->E5_CLIFOR
	cLoja      	:= (cAliasSE5)->E5_LOJA
	dDigit     	:= (cAliasSE5)->E5_DTDIGIT
	lBxTit	  	:= .F.
	cFilorig    := (cAliasSE5)->E5_FILORIG
	
	If ((cAliasSE5)->E5_RECPAG == "R" .and. ! ((cAliasSE5)->E5_TIPO $ "PA /"+MV_CPNEG )) .or. ;	//Titulo normal
		((cAliasSE5)->E5_RECPAG == "P" .and.   ((cAliasSE5)->E5_TIPO $ "RA /"+MV_CRNEG )) 	//Adiantamento
		dbSelectArea("SE1")
		dbSetOrder(1)
		lBxTit := MsSeek(cFilial+cPrefixo+cNumero+cParcela+cTipo)
		If !lBxTit
			lBxTit := dbSeek((cAliasSE5)->E5_FILORIG+cPrefixo+cNumero+cParcela+cTipo)
		Endif				
		cCarteira := "R"
		While SE1->(!Eof()) .and. SE1->E1_PREFIXO+SE1->E1_NUM+SE1->E1_PARCELA+SE1->E1_TIPO==cPrefixo+cNumero+cParcela+cTipo
			If SE1->E1_CLIENTE == cFornece .And. SE1->E1_LOJA == cLoja	// Cliente igual, Ok
				Exit
			Endif
			SE1->( dbSkip() )
		EndDo
		If !SE1->(EOF()) .and. !lManual .and.  ; //.And. mv_par11 == 1
			((cAliasSE5)->E5_RECPAG == "R" .and. !((cAliasSE5)->E5_TIPO $ MVPAGANT+"/"+MV_CPNEG))
			If !Empty(SE1->E1_SITUACA)
				dbSelectArea(cAliasSE5)
				(cAliasSE5)->(dbSkip())		      // filtro de registros desnecessarios
				Loop
			Endif
		Endif
		cCond3:="E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+DtoS(E5_DATA)+E5_SEQ+E5_NUMCHEQ==cPrefixo+cNumero+cParcela+cTipo+DtoS(dBaixa)+cSeq+cNumCheq"
		nValor := nVlMovFin := 0
	Else
		dbSelectArea("SE2")
		DbSetOrder(1)
		cCarteira := "P"
		lBxTit := MsSeek(cFilial+cPrefixo+cNumero+cParcela+cTipo+cFornece+cLoja)
		If !lBxTit
			lBxTit := dbSeek((cAliasSE5)->E5_FILORIG+cPrefixo+cNumero+cParcela+cTipo+cFornece+cLoja)
		Endif				
		cCond3:="E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+DtoS(E5_DATA)+E5_SEQ+E5_NUMCHEQ==cPrefixo+cNumero+cParcela+cTipo+cFornece+DtoS(dBaixa)+cSeq+cNumCheq"
		nValor := nVlMovFin := 0
		cCheque    := Iif(Empty((cAliasSE5)->E5_NUMCHEQ),SE2->E2_NUMBCO,(cAliasSE5)->E5_NUMCHEQ)
	Endif
	dbSelectArea(cAliasSE5)
	While (cAliasSE5)->(!Eof()) .And. &(cCond3)
		cTipodoc   := (cAliasSE5)->E5_TIPODOC
		If (cAliasSE5)->E5_SITUACA $ "C/E/X" 
			dbSelectArea(cAliasSE5)
			(cAliasSE5)->( dbSkip() )
			Loop
		EndIF
		
		If (cAliasSE5)->E5_LOJA != cLoja
			Exit
		Endif
		
		//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
		//ณ Verifica o vencto do Titulo ณ
		//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู                           
		If nMes <> Nil .and. nMes > 0   
			If SE2->(!Eof()) .And.	Substr(Dtos(SE2->E2_VENCREA),1,6) <> SUBSTR(DTOS(MV_PAR02),1,4)+cMes
				dbSelectArea(cAliasSE5)
				(cAliasSE5)->(dbSkip())
				Loop
			Endif
        Else
			If SE2->(!Eof()) .And.	Substr(Dtos(SE2->E2_VENCREA),1,6) <> Alltrim(MV_PAR01)+cMes
				dbSelectArea(cAliasSE5)
				(cAliasSE5)->(dbSkip())
				Loop
			Endif        
        EndIf  
		dBaixa     	:= (cAliasSE5)->E5_DATA
		cBanco     	:= (cAliasSE5)->E5_BANCO
		cSeq       	:= (cAliasSE5)->E5_SEQ
		cNumCheq   	:= (cAliasSE5)->E5_NUMCHEQ
		cRecPag		:= (cAliasSE5)->E5_RECPAG
		cMotBaixa	:= (cAliasSE5)->E5_MOTBX
		cTipo190	:= (cAliasSE5)->E5_TIPO
		dbSelectArea("SM2")
		dbSetOrder(1)
		dbSeek((cAliasSE5)->E5_DATA)
		dbSelectArea(cAliasSE5)
		nRecSe5:=(cAliasSE5)->SE5RECNO
		If (cAliasSE5)->E5_TIPODOC $ "VL/V2/BA/RA/PA/CP"
			nValor		+=(cAliasSE5)->E5_VALOR
		Else
			nVlMovFin	+=(cAliasSE5)->E5_VALOR
		Endif
		DbSelectArea(cAliasSE5)	
		(cAliasSE5)->(dbSkip())
		If lManual		// forca a saida do looping se for mov manual
			Exit
		Endif
	EndDo
	If (nValor+nVlMovFin) > 0
		//ฺฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฟ
		//ณ C lculo do Abatimento        ณ
		//ภฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤู
		If cCarteira == "R" .and. !lManual
			dbSelectArea("SE1")
			nRecno := Recno()
			nAbat := SomaAbat(cPrefixo,cNumero,cParcela,"R",1,,cFornece,cLoja)
			dbSelectArea("SE1")
			dbGoTo(nRecno)
		Elseif !lManual
			dbSelectArea("SE2")
			nRecno := Recno()
			nAbat :=	SomaAbat(cPrefixo,cNumero,cParcela,"P",1,,cFornece,cLoja)
			dbSelectArea("SE2")
			dbGoTo(nRecno)
		EndIF
		If !lManual
			lOriginal := .T.
		Else
			dbSelectArea("SE5")
			dbgoto(nRecSe5)
			nAbat:= 0
			lOriginal := .t.
			nRecSe5:=(cAliasSE5)->SE5RECNO
		Endif
		If empty(cMotBaixa)
			cMotBaixa := "NOR"  //NORMAL
		Endif
		nTotBaixado	+= Iif(cTipodoc == "CP",0,nValor)		// no soma, j  somou no principal
		nTotAbat   	+= nAbat
		nTotValor  	+= IIF( nVlMovFin <> 0, nVlMovFin , Iif(MovBcoBx(cMotBaixa),nValor,0))
		nTotMovFin 	+= nVlMovFin	
		nTotComp	+= Iif(cTipodoc == "CP",nValor,0)
		nValor := nAbat := nVlMovFin := 0
	Endif
	dbSelectArea(cAliasSE5)
Enddo
(cAliasSE5)->(DbCloseArea())	

RestArea(aArea)

Return(nTotValor)

/*/
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑฺฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤยฤฤฤฤฤฤยฤฤฤฤฤฤฤฤฤฤฟฑฑ
ฑฑณFuncao    ณCMB007Filtณ Autor ณ Eduardo Zanardo       ณ Data ณ 20.02.04 ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณDescrio ณ Testa as condicoes do registro do SE5 para permitir a impr.ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณSintaxe e ณ CMB007Filt()      			    						  ณฑฑ
ฑฑรฤฤฤฤฤฤฤฤฤฤลฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤดฑฑ
ฑฑณ Uso      ณ CMB007								    				  ณฑฑ
ฑฑภฤฤฤฤฤฤฤฤฤฤมฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤฤูฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
/*/
Static Function CMB007Filt(cAliasSE5)
Local lRet 		:= .T.
Local lManual 	:= .F.

If Empty((cAliasSE5)->E5_TIPODOC) .Or. Empty((cAliasSE5)->E5_NUMERO)
	lManual := .t.
EndIf

Do Case
Case (cAliasSE5)->E5_TIPODOC $ "DC/D2/JR/J2/TL/MT/M2/CM/C2" 
	lRet := .F.
Case (cAliasSE5)->E5_SITUACA $ "C/E/X" .or. (cAliasSE5)->E5_TIPODOC $ "TR#TE" .or.;
	((cAliasSE5)->E5_TIPODOC == "CD" .and. (cAliasSE5)->E5_VENCTO > (cAliasSE5)->E5_DATA)
	lRet := .F.
Case (cAliasSE5)->E5_TIPODOC == "E2"
	lRet := .F.
Case (cAliasSE5)->E5_TIPODOC == "TR"
	lRet := .F.
Case !MovBcoBx((cAliasSE5)->E5_MOTBX) .and. !lManual	
	lRet := .F.
Case TemBxCanc((cAliasSE5)->(E5_PREFIXO+E5_NUMERO+E5_PARCELA+E5_TIPO+E5_CLIFOR+E5_LOJA+E5_SEQ))	
	lRet := .F.
EndCase

Return lRet                                    
