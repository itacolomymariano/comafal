#INCLUDE "rwmake.ch"
#include "TopConn.Ch"

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �PrcMdVda  � Autor � Five Solutions     � Data �  15/10/2010 ���
�������������������������������������������������������������������������͹��
���Descricao � Rela��o de Pre�o M�dio do Grupo de Materiais por Ano/M�s   ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL - PE, RS e SP - Faturamento                        ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
/*/

User Function PrcMdVda


//���������������������������������������������������������������������Ŀ
//� Declaracao de Variaveis                                             �
//�����������������������������������������������������������������������

Local cDesc1         := "Este programa tem como objetivo imprimir relatorio "
Local cDesc2         := "de acordo com os parametros informados pelo usuario."
Local cDesc3         := "Faturamento M�dio por Grupo de Materiais"
Local cPict          := ""
Local titulo       := "Faturamento M�dio por Grupo de Materiais"
Local nLin         := 80

Local Cabec1       := "Grupo                                      Ano/Mes"
Local Cabec2       := "Filial   Preco Medio de Vendas"

Local imprime      := .T.
Local aOrd := {}
Private lEnd         := .F.
Private lAbortPrint  := .F.
Private CbTxt        := ""
Private limite           := 80
Private tamanho          := "P"
Private nomeprog         := "PrcMdVda" // Coloque aqui o nome do programa para impressao no cabecalho
Private nTipo            := 15
Private aReturn          := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
Private nLastKey        := 0
Private cPerg       := "RVMPRD"
Private cbtxt      := Space(10)
Private cbcont     := 00
Private CONTFL     := 01
Private m_pag      := 01
Private wnrel      := "PrcMdVda" // Coloque aqui o nome do arquivo usado para impressao em disco

Private cCFOPVenda  := "5101,6101,5102,6102,5103,6103,5104,6104,5105,6105,"
        cCFOPVenda  += "5106,6106,6107,6108,5109,6109,5110,6110,5111,6111,5112,6112,"
        cCFOPVenda  += "5113,6113,5114,6114,5115,6115,5116,6116,5117,6117,5118,6118,5119,6119,"
        cCFOPVenda  += "5120,6120,5122,6122,5123,6123"//,7101,7102,7105,7106,7127"
           
Private cString := "SD2"

dbSelectArea("SD2")
dbSetOrder(1)

MontPerg()
Pergunte(cPerg,.F.)

//���������������������������������������������������������������������Ŀ
//� Monta a interface padrao com o usuario...                           �
//�����������������������������������������������������������������������

wnrel := SetPrint(cString,NomeProg,cPerg,@titulo,cDesc1,cDesc2,cDesc3,.F.,aOrd,.F.,Tamanho,,.F.)

dPeriodo := MV_PAR01
cGrpPrd  := MV_PAR02
cFiliais := Alltrim(MV_PAR03)

If nLastKey == 27
	Return
Endif

SetDefault(aReturn,cString)

If nLastKey == 27
   Return
Endif

nTipo := If(aReturn[4]==1,15,18)

//���������������������������������������������������������������������Ŀ
//� Processamento. RPTSTATUS monta janela com a regua de processamento. �
//�����������������������������������������������������������������������

RptStatus({|| RunReport(Cabec1,Cabec2,Titulo,nLin) },Titulo)
Return

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Fun��o    �RUNREPORT � Autor � AP6 IDE            � Data �  15/10/10   ���
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

Local nOrdem

dbSelectArea(cString)
dbSetOrder(1)

//�����������������������������������������������������������������������
//� QUERY - Sele��o dos Registros do Relat�rio.                         �
//�����������������������������������������������������������������������
/*
dPeriodo := MV_PAR01
cGrpPrd  := MV_PAR02
cFiliais := Alltrim(MV_PAR03)
  */
cQrySD2 := " SELECT CASE WHEN SD2.D2_FILIAL = '02' THEN 'COSP' "
cQrySD2 += "             WHEN SD2.D2_FILIAL = '03' THEN 'CORS' "
cQrySD2 += "             WHEN SD2.D2_FILIAL = '04' THEN 'COPB' "
cQrySD2 += "             WHEN SD2.D2_FILIAL = '06' THEN 'COPE' END AS FILIAL, "
cQrySD2 += "        SBM.BM_DESC GRUPO, "
cQrySD2 += "        SUBSTRING(SD2.D2_EMISSAO,1,6) MES, "
cQrySD2 += "        AVG(SD2.D2_PRCVEN+(SD2.D2_VALIPI/SD2.D2_QUANT)+(SD2.D2_VALFRE/SD2.D2_QUANT)) PRMDVDA "
cQrySD2 += "   FROM "+RetSQLName("SD2")+" SD2,"+RetSQLName("SBM")+" SBM "
cQrySD2 += "  WHERE SD2.D2_FILIAL IN "+FormatIn(cFiliais,",")
cQrySD2 += "    AND SBM.BM_FILIAL = '"+xFilial("SBM")+"' "
cQrySD2 += "    AND SD2.D_E_L_E_T_ <> '*' "
cQrySD2 += "    AND SBM.D_E_L_E_T_ <> '*' "
cQrySD2 += "    AND SUBSTRING(SD2.D2_COD,1,2) = SUBSTRING(SBM.BM_GRUPO,1,2) "
cQrySD2 += "    AND SD2.D2_CLIENTE <> '006629' "
cQrySD2 += "    AND SD2.D2_GRUPO = '"+cGrpPrd+"'"
cQrySD2 += "    AND SD2.D2_CF IN "+FormatIn(cCFOPVenda,",")
cQrySD2 += "    AND (SD2.D2_TIPO <> 'I') "
cQrySD2 += "    AND  SUBSTRING(SD2.D2_EMISSAO,1,6) = '"+Substr(DTOS(dPeriodo),1,6)+"'" 
cQrySD2 += "  GROUP BY SD2.D2_FILIAL,SBM.BM_DESC,SUBSTRING(SD2.D2_EMISSAO,1,6) "
cQrySD2 += "  ORDER BY SD2.D2_FILIAL,SBM.BM_DESC,SUBSTRING(SD2.D2_EMISSAO,1,6) ASC "

MemoWrite("PrecoMedioVendaGrupo.SQL",cQrySD2)
MemoWrite("C:\TEMP\PrecoMedioVendaGrupo.SQL",cQrySD2)

TCQuery cQrySD2 NEW ALIAS "TMVDA" 
TCSetField("TMVDA","PRMDVDA","N",17,02)

//�����������������������������������������������������������������������
//� SETREGUA -> Indica quantos registros serao processados para a regua �
//�����������������������������������������������������������������������

//SetRegua(RecCount())

//dbGoTop()
DbSelectArea("TMVDA")
While TMVDA->(!EOF())

   //���������������������������������������������������������������������Ŀ
   //� Verifica o cancelamento pelo usuario...                             �
   //�����������������������������������������������������������������������

   If lAbortPrint
      @nLin,00 PSAY "*** CANCELADO PELO OPERADOR ***"
      Exit
   Endif

   //���������������������������������������������������������������������Ŀ
   //� Impressao do cabecalho do relatorio. . .                            �
   //�����������������������������������������������������������������������

   If nLin > 55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
      Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
      nLin := 9
   Endif

   //          10        20        30        40        50        60        70        80
   // 012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
   //"Grupo                                      Ano/M�s"
   //"Filial   Preco Medio de Vendas"
   // xxxxxxxxxXxxxxxxxxxXxxxxxxxxxXxxxxxxxxxX   2010/10
   // COPE             99,999,999.99
   //
   
   cGrpo := TMVDA->GRUPO
   @nLin,000 PSAY Substr(TMVDA->GRUPO,1,40)
   @nLin,043 PSAY Substr(TMVDA->MES,1,4)+"/"+Substr(TMVDA->MES,5,2)
   
   nLin ++
   nLin ++

   While TMVDA->GRUPO == cGrpo .And. TMVDA->(!EOF())
   
      If nLin > 55 // Salto de P�gina. Neste caso o formulario tem 55 linhas...
         Cabec(Titulo,Cabec1,Cabec2,NomeProg,Tamanho,nTipo)
         nLin := 9
      Endif   

      @nLin,000 PSAY Alltrim(TMVDA->FILIAL)
      @nLin,017 PSAY PRMDVDA Picture "@E 99,999,999.99"

      nLin := nLin + 1 // Avanca a linha de impressao
   
   DbSelectArea("TMVDA") 
   DbSkip() // Avanca o ponteiro do registro no arquivo
   EndDo
   
EndDo


//���������������������������������������������������������������������Ŀ
//� Finaliza a execucao do relatorio...                                 �
//�����������������������������������������������������������������������
DbSelectArea("TMVDA")
DbCloseArea()

SET DEVICE TO SCREEN


//���������������������������������������������������������������������Ŀ
//� Se impressao em disco, chama o gerenciador de impressao...          �
//�����������������������������������������������������������������������

If aReturn[5]==1
   dbCommitAll()
   SET PRINTER TO
   OurSpool(wnrel)
Endif

MS_FLUSH()

Return

Static Function MontPerg

   Local aArea := GetArea()

   //PutSx1(cGrupo,cOrdem,cPergunt  ,"","",cVar    ,cTipo,nTamanho,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
 
   Local aHelpPor := {}
   Local aHelpEng := {}
   Local aHelpSpa := {}

   Aadd( aHelpPor, 'Informe a Data de Referencia')
   Aadd( aHelpPor, 'Ser�o considerados Ano/M�s')
                                                                                                           //F3
   PutSx1(cPerg ,"01","Data de Refer�ncia"  ,"Data de Refer�ncia","Data de Refer�ncia","mv_ch1","D"  ,08       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR01","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)
   
   Aadd( aHelpPor, 'Informe o Grupo de Produtos')

 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"02"  ,"Grupo"  ,"Grupo","Grupo","mv_ch2","C"  ,4       ,0       ,0      ,"G" ,""    ,"SBM" ,""     ,"","MV_PAR02","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   aHelpPor := {}
   aHelpEng := {}
   aHelpSpa := {}

   Aadd( aHelpPor, 'Informe as Fiiais separando por virgula')
 //PutSx1(cGrupo,cOrdem,cPergunt       ,""           ,""           ,cVar    ,cTipo,nTamanho ,nDecimal,nPresel,cGSC,cValid,cF3,cGrpSxg,"",cVar01    ,cDef01 ,"","",cCnt01,cDef02       ,"","",cDef03     ,"","",cDef04      ,"","",cDef05,"","",aHelpPor,"","",cHelp)
   PutSx1(cPerg ,"03"  ,"Filiais"  ,"Filiais","Filiais","mv_ch3","C"  ,50       ,0       ,0      ,"G" ,""    ,"" ,""     ,"","MV_PAR03","","","",""    ,"","","","","","","","","","","","",aHelpPor,aHelpEng,aHelpSpa)

   RestArea(aArea)

Return(.T.)