#Include "Protheus.CH"
#Include "TOPCONN.CH"

/*
Nomenclatura e definicao do Lote do Produto (PI e PA)

���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ���Ŀ
� X �� X �� 9 �� 9 �� 9 �� 9 �� 9 �� 9 �� 9 �� 9 �
��������������������������������������������������    
  \     |    \    |     \  |     \    \    |    |
   \    |     \   |      \ |      \____\___|____|____________ Sequencia Numerica do Lote
    \   |      \  |       \|      
     \  |       \ |        \_________________________________ Ano da Producao do Poduto
      \ |        \|        
       \|         \__________________________________________ Mes da Producao do Produto
        \
         \___________________________________________________ Letra do Lote (BM_LOTE)


No paramtero MV_LOTECTL, sera armazenado apenas o Mes, Ano e Sequencia.

�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
��� Programa � CAi001   � Autor � Larson Zordan         � Data �13.08.2003���
�������������������������������������������������������������������������Ĵ��
���Descricao � Programa de composicao do Lote dos PIs e MPs               ���
�������������������������������������������������������������������������Ĵ��
���Sintaxe   � CAI001(ExpN1,ExpA1,ExpC1)                                  ���
�������������������������������������������������������������������������Ĵ��
���Parametros� ExpN1 = Opcao para a criacao do Lote do produto            ���
���          �         1 = Chama a funcao pelo D3_PRODUTO (X3_VLDUSER)    ���
���          �         2 = Chama a funcao pelo PE EMP650 (CAI002)         ���
���          � ExpA1 = Array com os dados do lote criado                  ���
���          �         [..1] = Lote do Produto                            ���
���          �         [..2] = Data da Validade do Lote                   ���
���          � ExpC1 = Produto                                            ���
�������������������������������������������������������������������������Ĵ��
���Uso       � COMAFAL                                                    ���
�������������������������������������������������������������������������Ĵ��
��� ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ���
�������������������������������������������������������������������������Ĵ��
��� PROGRAMADOR  � DATA     �HLPDSK�  MOTIVO DA ALTERACAO                 ���
�������������������������������������������������������������������������Ĵ��
��� Larson       �13/08/2003�      � Desenvolvimento incial do programa.  ���
�������������������������������������������������������������������������Ĵ��
��� Larson       �14/01/2004�      � Inclusao de parametros para criar o  ���
���              �          �      � lote a partir de outros programas.   ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������*/
User Function CAi001(nOpc,aLote,cProduto)
Local aAreaSBM   := SBM->(GetArea())
Local aAreaSB1   := SB1->(GetArea())
Local cLote      := GetMv("MV_LOTECTL")  //--> Criar o Parametro
Local cLetra     := "XX"
Local cNumOC     := CriaVar("C2_NUMOC",.F.)

Default aLote    := Array(2)
Default cProduto := Posicione("SC2",1,xFilial("SC2")+M->D3_OP,"C2_PRODUTO")

SB1->(dbSetOrder(1))
SB1->(dbSeek(xFilial("SB1")+cProduto))

SBM->(dbSetOrder(1))
//�������������������������������������������������Ŀ
//� Garante que a composicao do Lote sera apenas    �
//� para os produtos do tipo PI e PA. (MATA250)     �
//���������������������������������������������������

/**********************************
* Five Solutions Consultoria
* Itacolomy Mariano
* 03/01/2008 - 11:43 hs
* Comentario: Comentado uso do campo Lote neste processo devido a descontinuidade da utiliza��o do recurso de
*             rastreabilidade(Lote) para produtos acabados(TIPO = 'PA') e Slit(TIPO = 'PI')
*  
* If l250 .And. SB1->B1_TIPO $ "PI|PA"
*
***********************************/
If l250 .And. Substr(SB1->B1_TIPO,1,2) <> "PI" .And. Substr(SB1->B1_TIPO,1,2) <> "PA"

	If AllTrim(SB1->B1_GRUPO) == "20" .And. nOpc == 1
		cNumOC := Posicione("SC2",1,xFilial("SC2")+M->D3_OP,"C2_NUMOC")
	EndIf
	
	//--> Nao Tem OC e nao eh SLIT, portanto devera gerar o Lote
	If Empty(cNumOC)
		//�������������������������������������������������Ŀ
		//� Caso o BM_LOTE estiver em branco, padronizar    �
		//� sempre em XX.                                   �
		//���������������������������������������������������
		SBM->(dbSeek(xFilial("SBM")+SB1->B1_GRUPO))
		cLetra := If(Empty(SBM->BM_LOTE),cLetra,SBM->BM_LOTE)
		
		//�������������������������������������������������Ŀ
		//� Gera o proximo numero do Lote com 4 caracteres. �
		//� Sera Zerado quando mudar o Mes/Ano.             �
		//���������������������������������������������������
		If Left(cLote,4) == (StrZero(Month(dDataBase),2) + Right(Str(Year(dDataBase),4),2))
			cLote := Left(cLote,4) + Soma1(Right(cLote,4))
		Else
			cLote := (StrZero(Month(dDataBase),2) + Right(Str(Year(dDataBase),4),2)) + "0001"
		EndIf
	
		//�������������������������������������������������Ŀ
		//� Grava o ultimo numero do lote gerado no SX6.    �
		//���������������������������������������������������
		PutMv("MV_LOTECTL",cLote)
	
		aLote[1] := cLetra+cLote
		aLote[2] := dDataBase+365
	
		If nOpc == 1
			M->D3_LOTECTL := aLote[1]
			M->D3_DTVALID := aLote[2]
		EndIf	
	Else
		SZ2->(dbSetOrder(1))
		SZ2->(dbSeek(xFilial("SZ2")+cNumOC))
		While !SZ2->(Eof()) .And. cNumOC == SZ2->Z2_NUM
			If Left(SZ2->Z2_DESC,15) == cProduto
				M->D3_LOTECTL := SubStr(SZ2->Z2_DESC,110,10)
				M->D3_DTVALID := Posicione("SZ3",1,xFilial("SZ3")+cNumOC,"Z3_DATA")+365
				Exit
			EndIf
			SZ2->(dbSkip())
		EndDo
	EndIf	
EndIf
RestArea(aAreaSBM)

Return(.T.)

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
��� Programa � CAi001A  � Autor � Larson Zordan         � Data �14.08.2003���
�������������������������������������������������������������������������Ĵ��
���Descricao � Programa para compor o lote da bobina (D1_LOTEFOR)         ���
�������������������������������������������������������������������������Ĵ��
���Uso       � COMAFAL                                                    ���
�������������������������������������������������������������������������Ĵ��
��� ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL.                     ���
�������������������������������������������������������������������������Ĵ��
��� PROGRAMADOR  � DATA     �HLPDSK�  MOTIVO DA ALTERACAO                 ���
�������������������������������������������������������������������������Ĵ��
��� Larson       �14/08/2003�      � Desenvolvimento incial do programa.  ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������*/
User Function CAi001A()
Local lRet   := .T.
Local nPosPR := aScan(aHeader,{ |x| Upper(AllTrim(x[2])) == "D1_COD"    })
Local nPosLT := aScan(aHeader,{ |x| Upper(AllTrim(x[2])) == "D1_LOTECTL"})
Local nPosVD := aScan(aHeader,{ |x| Upper(AllTrim(x[2])) == "D1_DTVALID"})
Local cLoteF := M->D1_LOTEFOR

//�������������������������������������������������Ŀ
//� Verifica o campo Lote do Fornecedor para compor �
//� o lote do produto (D1_LOTECTL e D1_DTVALID).    �
//���������������������������������������������������
If Empty(cLoteF) .And. Rastro(aCols[n,nPosPR]) .And. !Empty(SA2->A2_LOTE)
	Aviso("LOTE EM BRANCO","O Lote do Fornecedor devera ser preenchido conforme descrito no produto.",{"<< Voltar"})
	lRet := .T.
EndIf

//�������������������������������������������������Ŀ
//� Grava o lote do produto ja composto com a letra �
//� definida no cadastro do fornecedor.             �
//���������������������������������������������������
aCols[n,nPosLT] := SA2->A2_LOTE + cLoteF
aCols[n,nPosVD] := dDataBase + 365

Return(lRet)
