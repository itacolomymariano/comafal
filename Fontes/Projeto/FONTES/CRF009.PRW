#INCLUDE "rwmake.ch"
#INCLUDE "topconn.ch"

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  � CRF009   � Autor � AP5 IDE            � Data �  13/11/03   ���
�������������������������������������������������������������������������͹��
���Descricao � Codigo gerado pelo AP5 IDE.                                ���
�������������������������������������������������������������������������͹��
���Uso       � AP5 IDE                                                    ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/

User Function CRF009()

Local aOrd := {}
Local cDesc1        := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2        := "de acordo com os parametros informados pelo usuario."
Local cDesc3        := "Pedido"
Local cPict         := ""
Local titulo        := "PEDIDO"
Local nLin          := 80
Local Cabec1   := ""
Local Cabec2   := ""
Local imprime  := .T.
Local cPerg       := "CRF009"

Private CbTxt       := ""
Private lEnd        := .F.
Private lAbortPrint := .F.
Private limite      := 132
Private tamanho     := "M"
Private nomeprog    := "CRF009" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo       := 18
Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey    := 0
Private cbtxt  := Space(10)
Private cbcont := 00
Private m_pag  := 01
Private wnrel  := "CRF009" // Coloque aqui o nome do arquivo usado para impressao em disco
Private cString := "SC5"

DbSelectArea("SC5")
DbSetOrder(1)

DbSelectArea("SC6")
DbSetOrder(1)

DbSelectArea("SA1")
DbSetOrder(1)

DbSelectArea("SB1")
DbSetOrder(1)

AjustaSX1(cPerg)

pergunte(cPerg,.F.)

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

If nLastKey == 27
	Return
EndIf

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
EndIf

nTipo := If(aReturn[4]==1,15,18)

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Fun��o    �RUNREPORT � Autor � AP5 IDE            � Data �  13/11/03   ���
�������������������������������������������������������������������������͹��
���Descri��o � Funcao auxiliar chamada pela RPTSTATUS. A funcao RPTSTATUS ���
���          � monta a janela com a regua de processamento.               ���
�������������������������������������������������������������������������͹��
���Uso       � Programa principal                                         ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/

Static Function RunReport(Cabec1,Cabec2,Titulo,nLin)

Local nOrdem 	:= 0
Local cTIPO		:=""
Local nLINHAS	:=0
Local nQTDTOT	:=0
Local nVALTOT	:=0
Local nVALDESC	:=0
Local aLINHAS 	:={}                              
Local cObsUser 	:= ""   
Local cPrazo	:= "" 

SC9->(dbSetOrder(1))

DbSelectArea(cString)
DbSetOrder(1)

SetRegua(val(MV_PAR02)-val(MV_PAR01)+1)

SC5->( MsSeek( xFilial('SC5') + MV_PAR01, .t. ) )
while SC5->(!EOF() .and. C5_FILIAL == xFilial('SC5') .and. C5_NUM <= MV_PAR02 )
   If lAbortPrint
      @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
      Exit
   EndIf
/*   
   //--> Permite a impressao apenas dos pedidos ja liberados
   If SC9->(dbSeek(xFilial("SC9")+SC5->C5_NUM))
        If !Empty(SC9->C9_BLEST) .Or. !Empty(SC9->C9_BLCRED)
	   		SC5->(dbSkip())
   			Loop
   		EndIf	
   Else
   		SC5->(dbSkip())
   		Loop
   EndIf
*/
   Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)

   cTIPO	:=""
   nLINHAS	:=0
   nQTDTOT	:=0
   nVALTOT	:=0
   nVALDESC	:=0
   aLINHAS 	:={}                              
   cObsUser := ""

   If SC5->C5_TIPO=="N"
      cTIPO:="Vendas"
   ELSEIf SC5->C5_TIPO=="D"
      cTIPO:="Dev.(Compras)"
   ELSEIf SC5->C5_TIPO=="C"
      cTIPO:="Compl. valores"
   ELSEIf SC5->C5_TIPO=="I"
      cTIPO:="Compl. ICMS"
   ELSEIF SC5->C5_TIPO=="P"
      cTIPO:="Compl. IPI"
   ENDIF
   @ prow()+1,0 PSAY If(SC5->C5_TPFRETE=="F","FOB","CIf") //chr(18)+chr(14)+If(SC5->C5_TPFRETE=="F","FOB","CIf")+chr(20)
   @ prow()  ,PCOL() PSAY '   DEPARTAMENTO DE VENDAS'//chr(18)+chr(14)+'   DEPARTAMENTO DE VENDAS' + chr (20)
   cCAB := '(Cotacao) '+alltrim(SM0->M0_CIDCOB)
   @ prow()+1,0 PSAY space(39-len(cCAB)) + cCAB//chr(14) + space(39-len(cCAB)) + cCAB + chr (20)
   @ prow(),pcol()+2 PSAY iIf( SC5->C5_LIBEROK <> 'S', '(BLOQUEADO)', ' ' )
   @ prow(),pcol()+2 PSAY ""//chr(15)
   @ prow()+1,0 PSAY replic("=",132)
   @ prow()+1,0 PSAY '*----------------------------------------------------------------------------------------------------------------------------------*'
   @ prow()+1,0 PSAY '|Vendedor                 |Cliente                 |Deposito                |Caixa                   |Expedicao                    |'
   @ prow()+1,0 PSAY '|-------------------------|------------------------|------------------------|------------------------|-----------------------------|'
   @ prow()+1,0 PSAY '|                         |                        |                        |                        |                             |'
   @ prow()+1,0 PSAY '|-------------------------|------------------------|------------------------|------------------------|-----------------------------|'
   @ prow()+1,0 PSAY '|' + replic("=",130) + '|'//CHR(15)+'|' + replic("=",130) + '|'
   @ prow()+1,0 PSAY '|' + Alltrim(SC5->C5_NUM)+ '|' + ' ('+cTIPO+')'
   @ prow()+1,0 PSAY '|'+replic("-",130)+'|'//CHR(15)+'|'+replic("-",130)+'|'

   DbSelectArea("SA3")
   MsSeek(xFILIAL("SA3")+SC5->C5_VEND1)
   DbSelectArea("SC5")
   @ prow()+1,000 PSAY '| Emitente : '+SUBSTR(cUSUARIO,7,10)
   @ prow()  ,024 PSAY 'Vendedor : '+C5_VEND1+'-'+SA3->A3_NOME
   @ prow()  ,131 PSAY '|'
   @ prow()+1,000 PSAY '|'+replic("=",130)+'|'
   @ prow()+1,000 PSAY '|'//chr(18)+'|'
   @ prow()  ,032 PSAY 'DADOS DO CLIENTE'
   @ prow()  ,077 PSAY '|' //+ chr(15)
   @ prow()+1,000 PSAY '|'+replic("-",130)+'|'

   If SC5->C5_TIPO $ "N#C#I#P"
      DbSelectArea("SA1")
      MsSeek(xFILIAL("SA1")+SC5->C5_CLIENTE+SC5->C5_LOJACLI)

      SX5->( DbSetOrder(1) )
      SX5->( MsSeek(xFILIAL("SX5")+'T3' + SA1->A1_SATIV1) )

      DbSelectArea("SC5")
      @ prow()+1,0 PSAY '| Nome : '+SA1->(A1_COD+" "+A1_LOJA+'-'+alltrim(A1_NOME)+' ('+ALLTRIM(A1_SATIV1))+')'
   else
      DbSelectArea("SA2")
      MsSeek(xFILIAL("SA2")+SC5->C5_CLIENTE+SC5->C5_LOJACLI)
      DbSelectArea("SC5")
      @ prow()+1,0 PSAY '| Nome : '+SA2->(A2_COD+'-'+alltrim(A2_NOME))
   endIf

   @ prow(),131 PSAY '|'
   @ prow()+1,0 PSAY '| End. : '
   @ prow(),pcol() PSAY If(C5_TIPO $ "N#C#I#P",SA1->A1_END,SA2->A2_END)
   @ prow(),pcol()+1 PSAY 'Bairro:'+SUBSTR(If(C5_TIPO $ "N#C#I#P",SA1->A1_BAIRRO,SA2->A2_BAIRRO),1,20)
   @ prow(),pcol()+1 PSAY 'Cidade:'+If(C5_TIPO $ "N#C#I#P",SA1->A1_MUN,SA2->A2_MUN)
   @ prow(),pcol()+1 PSAY 'Estado:'+If(C5_TIPO $ "N#C#I#P",SA1->A1_EST,SA2->A2_EST)
   @ prow(),131 PSAY '|'
   @ prow()+1,0 PSAY '| C.E.P.:'
   @ prow(),pcol() PSAY If(C5_TIPO $ "N#C#I#P",TransForm(SA1->A1_CEP,'@R 99999-999'),TransForm(SA2->A2_CEP,'@R 99999-999'))
   If C5_TIPO $ "N#C#I#P"
      @ prow(),pCOL()+1 PSAY '| Fones:'+TransForm(SA1->A1_TEL,'@R (999)999.99999') + ' / '
      @ prow(),pcol()+1 PSAY TransForm(SA1->A1_FAX,'@R (999)999.99999')
      @ prow(),pcol()+1 PSAY 'CGC/CPF:'+TransForm(SA1->A1_CGC,'@R 99.999.999/9999-99')
      @ prow(),pcol()+1 PSAY 'I.E.:' + SA1->A1_INSCR
   else
      @ prow(),pCOL()+1 PSAY '| Fones:'+TransForm(SA2->A2_TEL,'@R (999)999.99999') + ' / '
      @ prow(),pcol()+1 PSAY TransForm(SA2->A2_FAX,'@R (999)999.99999')
      @ prow(),pcol()+1 PSAY 'CGC/CPF:'+TransForm(SA2->A2_CGC,'@R 99.999.999/9999-99')
      @ prow(),pcol()+1 PSAY 'I.E.:' + SA2->A2_INSCR
   endIf
   @ prow(),131 PSAY '|'
   If C5_TIPO  $ "N#C#I#P"
      @ prow()+1,0 PSAY '| Cob. : '
      @ prow()  ,pcol() PSAY SA1->A1_ENDCOB
      @ prow()  ,pcol()+1 PSAY 'Bairro:'+SA1->A1_BAIRROC
      @ prow()  ,pcol()+1 PSAY 'Cidade:'+SA1->A1_MUNC
      @ prow()  ,pcol()+1 PSAY 'Estado:'+SA1->A1_ESTC
      @ prow()  ,131 PSAY '|'
      @ prow()+1,0 PSAY '| Segmento : '+SX5->X5_DESCRI
      @ prow()  ,pcol()+3 PSAY 'Nome Fantasia : '+SA1->A1_NREDUZ
   endIf
   @ prow(),131 PSAY '|'
   @ prow()+1,0 PSAY '*---------------------------------------------------------------------------------------------++-----------------------------------*'
   @ prow()+1,0 PSAY '|                                                                                             || ===== CONTROLE DE SEPARACAO ======|'
   @ prow()+1,0 PSAY '*------------------------------------------------------------------------------------------++--------------------------------------*'
   @ prow()+1,0 PSAY '|Item| Codigo     | Nomenclatura                   | TES |  Quant. | Prc-Unit.|  Prc-Total || Quant. Carreg |     Data Entrega     |'
   @ prow()+1,0 PSAY '|------------------------------------------------------------------------------------------++--------------------------------------|'
   *                  | 99 | 99.999.999 | 123456789012345678901234567890 | 999 | 999.999 | 9.999,99 | 999.999,99 ||               |                      |
   *                            1         2         3         4         5         6         7         8         9        10        11        12        13        14
   *                  012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890

   DbSelectArea("SC6")
   DbSetOrder(1)
   MsSeek(xFILIAL("SC6")+SC5->C5_NUM)

   nQTDGER:=0
   while SC6->(!EOF()) .and. SC6->C6_FILIAL+SC6->C6_NUM == SC5->C5_FILIAL+SC5->C5_NUM 
      SB1->(DbSetOrder(1) )
      SB1->(MsSeek(xFILIAL("SB1")+SC6->C6_PRODUTO) )
      SB2->(MsSeek( xFILIAL("SB2") + SB1->( B1_COD + B1_LOCPAD ) ) )

      //                       1            2              3              4             5              6               7              8               9          10            11
      AADD(aLINHAS,{SB1->B1_COD,SB1->B1_DESC,SC6->C6_QTDVEN,SC6->C6_PRCVEN,SC6->C6_VALOR,SC6->C6_QTDLIB,SC5->C5_TPFRETE,SB1->B1_LOCPAD,SB2->B2_LOCALIZ,SC6->C6_TES,SC6->C6_ENTREG})

      nQTDTOT  += SC6->C6_QTDVEN
      nVALTOT  += SC6->C6_VALOR
      nVALDESC += SC6->C6_VALDESC

      DbSelectArea("SC6")
      SC6->( dbskip() )
      @ prow(),131 PSAY'|'
   ENDDO

   aSORT(aLINHAS,,,{|X,Y| X[1] < Y[1]})

   FOR I:=1 TO LEN(aLINHAS)
       @ prow()+1,000 PSAY '| ' + strzero(I,2) + ' | '

       // Produto
       @ prow()  ,007 PSAY TransForm(aLINHAS[I][1],'@R 999999999')+'  | '+aLINHAS[I][2]
       @ prow()  ,051 PSAY '| ' + aLinhas[I,10] + ' |'

       // Quantidade
       @ prow()  ,059 PSAY TransForm(aLINHAS[I][3],'@E 999.999') + ' |'

       // Preco unitario
       @ prow()  ,068 PSAY TransForm(aLINHAS[I][4],'@E 99,999.99') + ' |'

       // Preco total
       @ prow(),080 PSAY TransForm(aLINHAS[I][5],'@E 999,999.99') + ' || '
       @ prow(),108 PSAY '|   '+dtoc(aLINHAS[I][11])//+chr(14)+dtoc(aLINHAS[I][11])+chr(20)
       @ prow(),pcol()+2 PSAY ' |'//chr(15)+' |'
   NEXT

   nICMS := 0
   DbSelectArea("SC5")
   @ prow()+1,000 PSAY replic('=',132)

   @ prow()+1,000 PSAY ' Totais => Peso:' + TransForm(C5_PBRUTO,'@E 9,999,999')+;
                       ' / Qtd.:' + TransForm(nQTDTOT,'@E 999.999') + ;
                       ' / Valor(R$):' + TransForm(nVALTOT,'@E 9999,999.99') + ;
                       ' / ICMS-F(R$):' + TransForm(nICMS,'@E 9,999.99')

   @ prow(),pcol()+1 PSAY '/ Frete:'+TransForm(SC5->C5_FRETE,'@R 9999.99')
   
   @ prow(),pcol()+1 PSAY '/ Total Geral(R$):'
   @ prow(),pcol() PSAY TransForm(nVALTOT+nICMS,'@E 9999,999.99')

   SE4->( DbSetOrder(1) )
   If SE4->( MsSeek(xFILIAL("SE4")+SC5->C5_CONDPAG) )
        cPRAZO := ALLTRIM(SE4->E4_DESCRI)
   ENDIf

   @ prow()+1,000 PSAY " Condicao de Pagamento: "+SC5->C5_CONDPAG+" - "+ cPRAZO
   @ prow()+1,000 PSAY replic('=',132)

   DbSelectArea("SC5")
   @ prow()+1,000 PSAY '| Obs.:'

   If nVALDESC > 0
      @ prow()+1,000 PSAY '| Desconto:'
      @ prow()  ,012 PSAY nVALDESC picture '@E 9,999.99'
   endIf

   @ prow()+1,000 PSAY '| Obs. Usuario.:' + SC5->C5_X_OBUSE    // RICHARD - OBSERVACAO DO USUARIO
   @ prow()+1,000 PSAY replic('=',132)

   @ prow()+1,000 PSAY 'Obs.: COTACAO SUJEITA A CONFIRMACAO DE PRECO E PRAZO                                                                1. VIA EXPEDICAO'

   @ prow()+1,000 PSAY '      SEM VALOR FISCAL                                                                                              2. VIA CAIXA'

   @ prow()+2,000 PSAY 'Este pedido sera excluido em ' + DtoC(SC5->C5_EMISSAO+GetMv("MV_DEXCPV")) //CHR(14) + 'Este pedido sera excluido em ' + DtoC(SC5->C5_EMISSAO+GetMv("MV_DEXCPV")) + CHR(20) + CHR(15)

	If SC5->C5_NUM <> MV_PAR02
	   @ 0,0 PSAY ""//chr(18)
	EndIf	   

   SC5->( dbSkip() ) // Avanca o ponteiro do registro no arquivo
EndDo

SET DEVICE TO SCREEN

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
EndIf

MS_FLUSH()

Return

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Function  �AJUSTASX1 �Autor  �Eduardo Zanardo     � Data �  13/11/03   ���
�������������������������������������������������������������������������͹��
���Desc.     �Ajusta o SX1                                                ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/
Static Function AjustaSx1(cPerg)

Local aArea := GetArea()

//+--------------------------------------------------------------+
//� Variaveis utilizadas para parametros                         �
//� MV_PAR01             // De Pedido				             �
//� MV_PAR02             // Ate o Pedido		                 �
//+--------------------------------------------------------------+

PutSx1(cPerg,"01","De Pedido ?","","","mv_ch1","C",6,0,0,"G","","","","",;
		"MV_PAR01","","","","","","","","","","","","","","","","")  
PutSx1(cPerg,"02","Ate o Pedido ?","","","mv_ch2","C",6,0,0,"G","","","","",;
		"MV_PAR02","","","","","","","","","","","","","","","","")				

RestArea(aArea)

Return(.T.)
