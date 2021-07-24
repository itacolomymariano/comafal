#Include "RwMake.Ch"
#Include "TopConn.Ch"
User Function ERF017(dDataini,nlin)

Local aArea := GetArea()
Local Cabec1 := "Num NF       Cliente                                                   EMISSAO"
Local Cabec2 := "Produto                           Preco            Quant       IPI     Frete        Total  Custo Un    Tot Custo      Lucro     (%)"
Local Titulo := "  Faturamento: "+DtoC(mv_par01)+" A "+DtoC(mv_par02)+' - Analitico'
Local aStruSD2 	:= SD2->(dbStruct())
Local nSD2	 	:= 0
Local _SItem  := 0
Local _SICusto:= 0
Local _SLucro := 0
Local _SDesc  := 0
Local _TItem  := 0
Local _TDesc  := 0
Local _TICusto:= 0
Local _TLucro := 0
Local _MIPI   := 0
Local _MFrete := 0
Local _MItem  := 0
Local _MCusto := 0
Local _MLucro := 0

nLin := 60

cQuery := " SELECT D2_DOC, D2_SERIE, D2_CLIENTE, D2_LOJA, "
cQuery += " D2_EMISSAO, B1.B1_COD, B1.B1_DESC, D2_QUANT, "
cQuery += " D2_PRCVEN, D2_TOTAL, D2_CUSTO1, D2_DESC, D2_VALIPI, "
cQuery += " D2_VALFRE, D2_PEDIDO,D2_ICMSRET,D2_SEGURO,D2_DESPESA "
cQuery += " FROM " 
cQuery += RetSqlName('SD2') + " D2 , " 
cQuery += RetSqlName('SB1') + " B1  " 
cQuery += " WHERE D2.D2_FILIAL  = '" + xFilial("SD2") + "' AND "
cQuery += " SUBSTRING(D2.D2_EMISSAO,1,6)  = '" + Substr(Dtos(dDataini),1,6) + "'  AND " 
cQuery += " D2.D2_TIPO    = 'N' AND "
cQuery += " D2.D2_LOCAL IN ('03','04','05') AND "
cQuery += " D2.D2_CF IN (" + Alltrim(cCFOPVenda) + ") AND "
cQuery += " D2.D_E_L_E_T_ <> '*' AND "    
cQuery += " D2.D2_COD = B1.B1_COD AND " 
cQuery += " B1.B1_FILIAL = '"+xFilial("SB1")+"' AND "
cQuery += " B1.D_E_L_E_T_ <> '*' "                
cQuery += " ORDER BY D2.D2_DOC, B1.B1_COD ASC "
cQuery := ChangeQuery(cQuery)
dbUseArea(.T.,"TOPCONN",TCGenQry(,,cQuery),"TMP",.F.,.T.)
For nSD2 := 1 To len(aStruSD2)
	If aStruSD2[nSD2][2] <> "C" .And. FieldPos(aStruSD2[nSD2][1])<>0
		TcSetField("TMP",aStruSD2[nSD2][1],aStruSD2[nSD2][2],aStruSD2[nSD2][3],aStruSD2[nSD2][4])
	EndIf
Next nSD2 
dbSelectArea("TMP")
TMP->(dbGoTop())
While TMP->(!Eof())
	If nLin >= 55
		Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
		nLin:=9
	EndIf	
	_nf:= TMP->D2_DOC
	_EMISSAO := TMP->D2_EMISSAO
	While TMP->(!Eof()) .and. TMP->D2_DOC == _nf .and. nLin<55 .and. !lAbortPrint
		@ nLin,000 Psay TMP->D2_DOC
		DbSelectArea("SA1")
		DbSetOrder(1)
		DbSeek(xFilial("SA1")+TMP->D2_CLIENTE+TMP->D2_LOJA)
		@ nLin,013 Psay TMP->(D2_cliente +'/'+ D2_loja) +' '+ SA1->A1_NREDUZ
		@ nLin,071 psay TMP->D2_EMISSAO picture '@D'
		nLin++
		nLin++
		While TMP->(!Eof()) .and. TMP->D2_DOC == _nf .and. nLin<55 .and. !lAbortPrint
			nQuant :=  TMP->D2_QUANT
			@ nLin,000 Psay Alltrim(TMP->B1_COD) + '-' + Alltrim(Substr(TMP->b1_desc,1,20))
			@ nLin,033 Psay TMP->D2_prcven picture "@E 9,999.99"
			@ nLin,051 Psay nQuant picture "@E 99,999.999"
			@ nLin,060 Psay TMP->D2_VALIPI picture "@E 9,999.99"
			@ nLin,069 Psay TMP->D2_VALFRE picture "@E 9,999.99"
			@ nLin,081 Psay TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA picture "@E 99,999.99" //TMP->((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant) picture "@E 99,999.99"
			@ nLin,092 Psay Iif(Alltrim(TMP->B1_COD) <> "038787",TMP->D2_CUSTO1/nQuant,0) picture "@E 9,999.99"
			@ nLin,104 Psay Iif(Alltrim(TMP->B1_COD) <> "038787",TMP->D2_CUSTO1,0) picture "@E 99,999.99"
			@ nLin,116 Psay ERF017B() picture "@E 9,999.99"
			@ nLin,130 Psay ERF017C() picture "@E 999,999.99"
			_SItem  += TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA  //((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant)
			_SDesc  += TMP->D2_desc
			_SICusto+= Iif(Alltrim(TMP->B1_COD) <> "038787",TMP->D2_CUSTO1,0)
			_SLucro += TMP->((TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA)- Iif(Alltrim(TMP->B1_COD) <> "038787",D2_CUSTO1,0)) //((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant-D2_CUSTO1)
			_TItem  += TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA   //TMP->((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant)
			_TDesc  += TMP->D2_desc
			_TICusto+= Iif(Alltrim(TMP->B1_COD) <> "038787",TMP->D2_CUSTO1,0)
			_TLucro += TMP->((TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA)-Iif(Alltrim(TMP->B1_COD) <> "038787",D2_CUSTO1,0)) //TMP->((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant-D2_CUSTO1)
			
			TMP->( dbSkip() )
			IncRegua()
			nLin++
			If _nf <> TMP->D2_DOC
				@ nLin,001 psay repl('-',132)
				nLin++
				@ nLin,014 Psay 'T o t a l da NF -> '
				@ nLin,078 Psay _SItem 		Picture "@E 9,999,999.99"
				@ nLin,101 Psay _SICusto 	Picture "@E 9,999,999.99"
				@ nLin,114 Psay _SLucro 	Picture "@E 9,999,999.99"
				_TVenda := _SItem
				_TCusNF := _SICusto
				@ nLin,130 Psay ((_TVenda/_TCusNF)-1)*100 Picture "@E 99,999.99"
				_MItem  += _SItem
				_MCusto += _SICusto
				_MLucro += _SLucro
				_SItem  := 0
				_SDesc  := 0
				_SICusto:= 0
				_SDifer := 0
				_SLucro := 0
				nLin++
				nLin++
			EndIf
		EndDo
		_nf := TMP->D2_DOC
		If Month(_EMISSAO) <> Month(TMP->D2_EMISSAO)
			@ nLin,001 psay repl('-',132)
			nLin++
			@ nLin,001 Psay 'Total Mes -> '
			@ nLin,065 Psay _MIpi 	Picture "@E 999,999.99"
			@ nLin,076 Psay _MFrete	Picture "@E 999,999.99"
			@ nLin,088 Psay _MItem 	Picture "@E 9,999,999.99"
			@ nLin,102 Psay _MCusto	Picture "@E 9,999,999.99"
			@ nLin,115 Psay _MLucro	Picture "@E 9,999,999.99"
			@ nLin,130 Psay ((_MItem/_MCusto)-1)*100  Picture "@E 99,999.99"
			_MIPI   := 0
			_MFrete := 0
			_MItem  := 0
			_MCusto := 0
			_MLucro := 0
			nLin += 3
		EndIf
	EndDo
Enddo
dbSelectArea("TMP")
dbCloseArea()  

@ nLin++,000 psay repl('*', 132)
@ nLin  ,014 Psay 'Total Geral   ---> '
@ nLin  ,078 Psay _TItem 	Picture "@E 9,999,999.99"
@ nLin  ,101 Psay _TICusto 	Picture "@E 9,999,999.99"
@ nLin  ,114 Psay _TLucro 	Picture "@E 999,999.99"
@ nLin,126 Psay ((_TItem/_TICusto)-1)*100 Picture "@E 99,999.99"

RestArea(aArea)

Return

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
��� Programa � ERF017B  � Autor � Larson Zordan         � Data �   12.2001���
�������������������������������������������������������������������������Ĵ��
���Descricao � Chamada do Relatorio                                       ���
�������������������������������������������������������������������������Ĵ��
���Uso       � EXPRESS                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
����������������������������������������������������������������������������*/
Static Function ERF017B()

nQuant := TMP->D2_QUANT
_lucro := TMP->((TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA)-Iif(Alltrim(TMP->B1_COD) <> "038787",D2_CUSTO1,0))   //TMP->((D2_prcven*nQuant)+(D2_valipi/D2_QUANT)*nQuant+(D2_VALFRE/D2_QUANT)*nQuant-D2_CUSTO1)

Return ( _lucro )

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
��� Programa � ERF017C  � Autor � Larson Zordan         � Data �   12.2001���
�������������������������������������������������������������������������Ĵ��
���Descricao � Chamada do Relatorio                                       ���
�������������������������������������������������������������������������Ĵ��
���Uso       � EXPRESS                                                    ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
����������������������������������������������������������������������������*/
Static Function ERF017C()

nQuant   := TMP->D2_QUANT
_venda   := TMP->D2_TOTAL+TMP->D2_VALIPI+TMP->D2_ICMSRET+TMP->D2_SEGURO+TMP->D2_VALFRE+TMP->D2_DESPESA
_custo   := Iif(Alltrim(TMP->B1_COD) <> "038787",TMP->D2_CUSTO1,0)
_percent := ( (_venda/_custo) -1 ) * 100

Return (_percent)         