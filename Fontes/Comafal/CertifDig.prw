#include "TopConn.Ch"
#include "Color.Ch"
#include "Colors.Ch"

/*
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????ͻ??
???Programa  ?CERTIFDIG ?Autor  ?Five Solutions      ? Data ?  12/09/2011 ???
?????????????????????????????????????????????????????????????????????????͹??
???Desc.     ? Relatorio Grafico para Impressao do Certificado Digital de ???
???          ? Qualidade.                                                 ???
?????????????????????????????????????????????????????????????????????????͹??
???Uso       ? COMAFAL - PE, SP e RS.                                    ???
?????????????????????????????????????????????????????????????????????????ͼ??
?????????????????????????????????????????????????????????????????????????????
?????????????????????????????????????????????????????????????????????????????
*/

User Function CertifDig() //nOpcImp)
   Local aCores	   := {	{ " EMPTY(ZJ_PEDIDO) ", "BR_VERDE" 	},;		
						{ "!EMPTY(ZJ_PEDIDO)" , "BR_VERMELHO" }}
   

   Private cCadastro  := OemToAnsi("Certificado Digital de Qualidade")
   Private aRotina    := {	{ OemToAnsi("Pesquisar"       ) ,"AxPesqui"		, 0 , 1},;
                            { OemToAnsi("Visualizar"      ) ,"AxVisual"	    , 0 , 2},;
                            { OemToAnsi("Alterar"         ) ,"AxAltera"	    , 0 , 4},;
                            { OemToAnsi("Emitir"          ) ,"U_BrowsCD"	, 0 , 2},;
                    	    { OemToAnsi("Legenda"         ) ,"U_LegSNOE"	, 0 , 2}}


   mBrowse(6,1,22,75,"SZJ",,,,,,aCores)

Return

User Function BrowsCD
   
   Local cAreaKi := GetArea()
   
   Private oPrint
   Private cArqLogo := "CO_LOGO.BMP" 
   Private cPedCert := SZJ->ZJ_PEDIDO 
   Private cCodItCrt:= SZJ->ZJ_PRODUTO
   Private cOrdProd := SZJ->ZJ_OP
   
   cCliCrtf := Posicione("SC5",1,xFilial("SC5")+cPedCert,"C5_CLIENTE")
   cLjCrtf  := Posicione("SC5",1,xFilial("SC5")+cPedCert,"C5_LOJACLI")
   
   Private cNmeCli  := Alltrim(Posicione("SA1",1,xFilial("SA1")+cCliCrtf+cLjCrtf,"A1_NOME"))
   Private cNumCert := SZJ->ZJ_CODIGO
   
   
   
   cQrySD2 := " SELECT D2_EMISSAO,D2_DOC,D2_SERIE,D2_COD,SUM(D2_QUANT) D2_QUANT "
   cQrySD2 += "   FROM "+RetSQLName("SD2")
   cQrySD2 += "  WHERE D2_FILIAL = '"+xFilial("SD2")+"'"
   cQrySD2 += "    AND D2_PEDIDO = '"+cPedCert+"'"
   cQrySD2 += "    AND D2_COD = '"+cCodItCrt+"'"
   cQrySD2 += "    AND D2_CLIENTE = '"+cCliCrtf+"'"
   cQrySD2 += "    AND D2_LOJA = '"+cLjCrtf+"'"
   cQrySD2 += "    AND D_E_L_E_T_ <> '*'"
   cQrySD2 += "   GROUP BY D2_EMISSAO,D2_DOC,D2_SERIE,D2_COD "
   MemoWrite("NotaFiscalCertificado.SQL",cQrySD2)
   MemoWrite("C:\TEMP\NotaFiscalCertificado.SQL",cQrySD2)
   
   TCQuery cQrySD2 NEW ALIAS "CSD2"
   
   TCSetField("CSD2","D2_EMISSAO","D",08,00)
   TCSetField("CSD2","D2_QUANT","N",17,03)
   
   DbSelectArea("CSD2")
   Private dDataCrt := CSD2->D2_EMISSAO
   Private nQtdCert := CSD2->D2_QUANT
   Private cNFCerti := CSD2->D2_DOC+"/"+CSD2->D2_SERIE 
   DbSelectArea("CSD2")
   DbCloseArea()
   
   Private nPsoTMt  := 0
   Private nPsoUnT  := 0
   Private nPsoNTMt := 0
   Private nPsoNUnT := 0
   Private nQAmarr  := 0
   
   DbSelectArea("SB1")
   DbSetOrder(1)
   cDscItem := ""
   If DbSeek(xFilial("SB1")+cCodItCrt)
      Private cDscItem := Alltrim(SB1->B1_DESC)
      Private nPsoTMt  := SB1->B1_PESTEMT
      Private nPsoUnT  := SB1->B1_PESTEOP
      Private nPsoNTMt := SB1->B1_PETEMMT
      Private nPsoNUnT := SB1->B1_PETEMPC
      Private nQAmarr  := SB1->B1_PECAAMA
   EndIf


   //Acabamentos
   Private lSupOlea := .F.
   Private lExtSerr := .F.
   Private lExtFaca := .F.
   //Ensaios Mecanicos
   Private lAcht90 := .F.
   Private lAcht0 := .F.
   Private lAlrga := .F.
   Private lCrva := .F.
   //Analise Visual
   Private lEmpCnf := .F.
   Private lTorCnf := .F.
               
   DbSelectArea("SBZ")
   DbSetOrder(1)
   If DbSeek(xFilial("SBZ")+cCodItCrt)

      //Acabamentos
      lSupOlea  := IF(SBZ->BZ_SUPOLEA == "S",.T.,.F.)
      lExtSerr  := IF(SBZ->BZ_EXTSERR == "S",.T.,.F.)
      lExtFaca  := IF(SBZ->BZ_EXTFACA == "S",.T.,.F.)
      //Ensaios Mecanicos
      lAcht90 := IF(SBZ->BZ_ACHAT90 == "S",.T.,.F.)
      lAcht0  := IF(SBZ->BZ_ACHAT0 == "S",.T.,.F.)
      lAlrga  := IF(SBZ->BZ_ALARG == "S",.T.,.F.)
      lCrva   := IF(SBZ->BZ_CURVAM == "S",.T.,.F.)
      //Analise Visual
      lEmpCnf  := IF(SBZ->BZ_EMPCONF == "S",.T.,.F.)
      lTorCnf  := IF(SBZ->BZ_TORCCON == "S",.T.,.F.)
      
   EndIf   
   //        10        20        30
   //1234567890123456789012345678901234567890
   //TB INDL FF QUAD 40X40 1.06 6000                                
   Private cDimenso := Substr(cDscItem,16,5) //Dimens?o do Material
   Private cEspessu := Substr(cDscItem,22,4) //Espessura do Material
   Private cComprim := Substr(cDscItem,27,4) //Comprimento do Material
   Private cLteForn := ""
   
   Private nCarbono := 0
   Private nManganes:= 0
   Private nFosforo := 0
   Private nEnxofre := 0
   Private nSilicio := 0
   Private nCromo   := 0
   Private nNiquel  := 0
   Private nCobre   := 0
   Private nAlumini := 0
   
   /*
   DbSelectArea("SZI")
   DbSetOrder(1)
   If DbSeek(xFilial("SZI")+cLteForn)
      
      nCarbono := SZI->ZI_CARBONO
      nManganes:= SZI->ZI_MANGANE
      nFosforo := SZI->ZI_FOSFORO
      nEnxofre := SZI->ZI_ENXOFRE
      nSilicio := SZI->ZI_SILICIO
      nCromo   := SZI->ZI_CROMO
      nNiquel  := SZI->ZI_NIQUEL
      nCobre   := SZI->ZI_COBRE
      nAlumini := SZI->ZI_ALUMINI
   
   EndIf
   */
      
   //Relacionamento para pegar Composi??o Qu?mica dos Lotes de Fornecedores
   // ZZ2 - Leitura de Etiquetas
   // SC2 - Ordem de Produ??o
   // SZ2 - Itens da Ordem de Corte
   // SB8 - Saldos por Lote
   // SZI - Composi??o Qu?mica dos Lotes de Fornecedores
   
   cQryZZ2 := " SELECT ZZ2.ZZ2_NUM AS PEDIDO, SC2.C2_NUM+SC2.C2_ITEM+SC2.C2_SEQUEN AS OP, SC2.C2_NUMOC ORDEM_CORTE, "
   cQryZZ2 += "        SB8.B8_LOTECTL AUTO_BOBINA, SB8.B8_LOTEFOR LOTE_FORNECEDOR, "
   cQryZZ2 += "        SZI.ZI_LOTEFOR,SZI.ZI_CARBONO,SZI.ZI_MANGANE,SZI.ZI_FOSFORO,SZI.ZI_ENXOFRE,SZI.ZI_SILICIO,"
   cQryZZ2 += "        SZI.ZI_CROMO,SZI.ZI_NIQUEL,SZI.ZI_COBRE,SZI.ZI_ALUMINI"
   
   cQryZZ2 += "  FROM "+RetSQLName("ZZ2")+" ZZ2,"+RetSQLName("SC2")+" SC2,"+RetSQLName("SZ2")+" SZ2, "+RetSQLName("SB8")+" SB8,"+RetSQLName("SZI")+" SZI"
   cQryZZ2 += " WHERE ZZ2.ZZ2_FILIAL = '"+xFilial("ZZ2")+"'"
   cQryZZ2 += "   AND SC2.C2_FILIAL = '"+xFilial("SC2")+"'" 
   cQryZZ2 += "   AND SZ2.Z2_FILIAL = '"+xFilial("SZ2")+"'"
   cQryZZ2 += "   AND SB8.B8_FILIAL = '"+xFilial("SB8")+"'"
   cQryZZ2 += "   AND SZI.ZI_FILIAL = '"+xFilial("SZI")+"'"
   
   
   cQryZZ2 += "   AND ZZ2.ZZ2_NUM = '"+cPedCert+"'"
   cQryZZ2 += "   AND ZZ2.ZZ2_OP = SC2.C2_NUM+SC2.C2_ITEM+SC2.C2_SEQUEN" //cProducao  := ALLTRIM(SC2->C2_NUM) + ALLTRIM(SC2->C2_ITEM) + ALLTRIM(SC2->C2_SEQUEN)
   
   cQryZZ2 += "   AND SC2.C2_NUMOC = SZ2.Z2_NUM "
    
   cQryZZ2 += "   AND SZ2.Z2_TIPO = 'B'"
   cQryZZ2 += "   AND SUBSTRING(SZ2.Z2_DESC,1,10) = SB8.B8_LOTECTL "
   
   cQryZZ2 += "   AND SB8.B8_LOTEFOR = SZI.ZI_LOTEFOR "
   
   cQryZZ2 += "   AND ZZ2.D_E_L_E_T_ <> '*'"
   cQryZZ2 += "   AND SZ2.D_E_L_E_T_ <> '*'"
   cQryZZ2 += "   AND SC2.D_E_L_E_T_ <> '*'"
   cQryZZ2 += "   AND SB8.D_E_L_E_T_ <> '*'"
   cQryZZ2 += "   AND SZI.D_E_L_E_T_ <> '*'"
   
   MemoWrite("Certi_Digital_LotesFornecedores.SQL",cQryZZ2)
   MemoWrite("C:\TEMP\Certi_Digital_LotesFornecedores.SQL",cQryZZ2)
   
   TCQuery cQryZZ2 NEW ALIAS "TLTF"
   
   TCSetField("TLTF","ZI_CARBONO","N",5,3)
   TCSetField("TLTF","ZI_MANGANE","N",5,3)
   TCSetField("TLTF","ZI_FOSFORO","N",5,3)
   TCSetField("TLTF","ZI_ENXOFRE","N",5,3)
   TCSetField("TLTF","ZI_SILICIO","N",5,3)
   TCSetField("TLTF","ZI_CROMO","N",5,3)
   TCSetField("TLTF","ZI_NIQUEL","N",5,3)
   TCSetField("TLTF","ZI_COBRE","N",5,3)
   TCSetField("TLTF","ZI_ALUMINI","N",5,3)
   
   Private aCompQuimi := {}
   DbSelectArea("TLTF")
   While TLTF->(!Eof())
      If aScan(aCompQuimi, { |x| x[1] == TLTF->ZI_LOTEFOR }) == 0
         MsgInfo("Acrescentando Lote Fornecedor: ["+TLTF->ZI_LOTEFOR+"]")
         aAdd(aCompQuimi, { TLTF->ZI_LOTEFOR,TLTF->ZI_CARBONO,TLTF->ZI_MANGANE,TLTF->ZI_FOSFORO,TLTF->ZI_ENXOFRE,TLTF->ZI_SILICIO,TLTF->ZI_CROMO,TLTF->ZI_NIQUEL,TLTF->ZI_COBRE,TLTF->ZI_ALUMINI })
      EndIf
      DbSelectArea("TLTF")
      DbSkip()
   EndDo
   DbSelectArea("TLTF")
   DbCloseArea()
   
   oPrint:= TMSPrinter():New( "Certificado Digital de Qualidade - COMAFAL" )
   oPrint:Setup()
   oPrint:SetLandscape() //ou SetPortrait()

   //oPrint:Setup()
   oPrint:StartPage()   // Inicia uma nova p?gina
   //Fontes
   oFont8  := TFont():New("Arial",9,8 ,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont8n := TFont():New("Arial",9,8 ,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont10 := TFont():New("Arial",9,10,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont10n:= TFont():New("Arial",9,10,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont12 := TFont():New("Arial",9,12,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont12n := TFont():New("Arial",9,12,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont14 := TFont():New("Arial",9,14,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont16 := TFont():New("Arial",9,16,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont16n:= TFont():New("Arial",9,16,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont18 := TFont():New("Arial",9,18,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont18n:= TFont():New("Arial Black",9,18,.T.,.F.,5,.T.,5,.T.,.F.)
   oFont24 := TFont():New("Arial",9,24,.T.,.T.,5,.T.,5,.T.,.F.)
   oFont36 := TFont():New("Arial",9,36,.T.,.T.,5,.T.,5,.T.,.F.)
   //oBrush := TBrush():New("",4)//Fundo Cinza
   oBrush := TBrush():New("",CLR_HGRAY)//Fundo Cinza
   nTitPag := 0
   nLinha  := 80
   
   // Filial por Helton:
   Do case
      case xFilial("SZB")="02"
           cFilialH:="SP"
      case xFilial("SZB")="03"
           cFilialH:="RS"
      case xFilial("SZB")="06"
           cFilialH:="PE"      
   EndCase
   //----fim-----------------
      
/*
oPrint:Say  (0050+nLinha,1200,"|",oFont8)
oPrint:Say  (0100+nLinha,1200,"|",oFont8)
oPrint:Say  (0150+nLinha,1200,"|",oFont8)
oPrint:Say  (0200+nLinha,1200,"|",oFont8)
oPrint:Say  (0250+nLinha,1200,"|",oFont8)
oPrint:Say  (0300+nLinha,1200,"|",oFont8)
oPrint:Say  (0350+nLinha,1200,"|",oFont8)
oPrint:Say  (0400+nLinha,1200,"|",oFont8)
oPrint:Say  (0450+nLinha,1200,"|",oFont8)
oPrint:Say  (0500+nLinha,1200,"|",oFont8)
oPrint:Say  (0550+nLinha,1200,"|",oFont8)

oPrint:Line (0050+nLinha,1700,0150+nLinha, 1700)//coluna 1
oPrint:Line (0050+nLinha,1850,0235+nLinha, 1850)//coluna 2
oPrint:Line (0150+nLinha,1450,0235+nLinha, 1450)//coluna 3
oPrint:Line (0150+nLinha,2100,0235+nLinha, 2100)//coluna 4
oPrint:Line (0235+nLinha,1600,0320+nLinha, 1600)//coluna 5
oPrint:Line (0405+nLinha,1950,0490+nLinha, 1950)//coluna 6

oReport:SkipLine()
oReport:SkipLine()
oReport:SkipLine()
oReport:FatLine()
oReport:SkipLine()
oReport:SayBitmap (oReport:Row(),005,cLogo,291,057)
oReport:SkipLine()
oReport:SkipLine()
oReport:SkipLine()
oReport:FatLine()

*/
      //oPrint:Say  (0050+nLinha,1200,"|",oFont8)
     //oReport:SayBitmap (oReport:Row(),005,cLogo,291,057)  
     //oPrint:SayBitmap(080,050, "\Bitmaps\LogoSiga.bmp",300,150) // Tem que estar abaixo do RootPath
/*
SayBitmap(nRow,nCol,cBitmap,nWidth,nHeight,nRaster)
Imprime uma imagem no relat?rio.

nRow		Linha para impress?o da imagem
nCol		Coluna para impress?o da imagem
cBitmap	Nome da imagem, podendo ser path de um arquivo ou resource compilado no 
reposit?rio
nWidth		Largura da imagem
nHeight	Altura da imagem
nRaster	Compatibilidade - N?o utilizado

*/
      //oPrint:SayBitmap(0100,0140, cArqLogo,400,250) // A imagem tem que estar abaixo do RootPath
      // Say(nRow,nCol,cText,oFont,nWidth,nClrText,nBkMode,nPad)
      //Primeiro Quadro - Logo Tipo - CERTIFICADO DE QUALIDADE - Numero do Certificado
      oPrint:Line (0050,0100,0275, 0100) //coluna 1
      oPrint:Line (0050,0100,0050, 3000) //Linha 1
      oPrint:Line (0050,3000,0275, 3000) //coluna 2
      oPrint:Line (0275,0100,0275, 3000) //Linha 2
	  oPrint:Say  (100,0920,"CERTIFICADO DE QUALIDADE",oFont24)//,,CLR_HBLUE)
      oPrint:SayBitmap(0080,0140, cArqLogo,280,190) // A imagem tem que estar abaixo do RootPath
      oPrint:Line (0050,0620,0275, 0620) //coluna apos a logo da empresa
      oPrint:Line (0050,2300,0275, 2300) //coluna antes do numero do certificado
      oPrint:Line (0175,2300,0175, 3000) //Linha que divide o numero do certificado
      oPrint:Say  (100,2420,"CERTIFICADO "+cFilialH+" N?",oFont16)//,,CLR_HBLUE) - Texto do numero do certificado
      oPrint:Say  (185,2500,cNumCert,oFont18,,CLR_HBLUE) //- numero do certificado
      
      //Segundo Quadro - Cliente
      oPrint:Line (0300,0100,0400, 0100) //coluna 1
      oPrint:Line (0300,0100,0300, 3000) //Linha 1
      oPrint:Line (0300,3000,0400, 3000) //coluna 2
      oPrint:Line (0400,0100,0400, 3000) //Linha 2
      oPrint:Say  (0320,0240,"CLIENTE:",oFont10)//,,CLR_HBLUE) - Texto CLIENTE:
      oPrint:Say  (0320,0840,cNmeCli,oFont12n)//,,CLR_HBLUE) - Imprime nome do Cliente:


      //Terceiro Quadro - Dados do Sistema - Pedido, Nota Fiscal, Data de Emissao e Quantidade.
      oPrint:Line (0450,0100,0600, 0100) //coluna 1
      oPrint:Line (0450,0100,0450, 3000) //Linha 1
      oPrint:Line (0450,3000,0600, 3000) //coluna 2
      oPrint:Line (0600,0100,0600, 3000) //Linha 2
      
      oPrint:Say  (0480,0140,"PEDIDO N?:",oFont10n)//,,CLR_HBLUE) - Texto PEDIDO N?:
      oPrint:Line (0450,0700,0600, 0700) //coluna apos o numero do pedido
      oPrint:Say  (0480,0750,"NOTA FISCAL:",oFont10n)//,,CLR_HBLUE) - Texto NOTA FISCAL:
      oPrint:Line (0450,1300,0600, 1300) //coluna apos o numero da Nota Fiscal
      oPrint:Say  (0480,1350,"DATA DE EMISSAO:",oFont10n)//,,CLR_HBLUE) - Texto DATA DE EMISSAO:
      oPrint:Line (0450,1900,0600, 1900) //coluna apos a Data de Emissao
      oPrint:Say  (0480,1950,"QUANTIDADE (Kg):",oFont10n)//,,CLR_HBLUE) - Texto QUANTIDADE (Kg):
      
      oPrint:Say  (0525,0140,cPedCert,oFont12)//,,CLR_HBLUE) - Numero do Pedido de Vendas.
      oPrint:Say  (0525,0750,cNFCerti,oFont12)//,,CLR_HBLUE) - Numero da Nota Fiscal de Venda.
      oPrint:Say  (0525,1550,DTOC(dDataCrt),oFont12)//,,CLR_HBLUE) - Data da Emissao
      oPrint:Say  (0525,2450,Transform(nQtdCert,"@E 999,999,999.999"),oFont12)//,,CLR_HBLUE) - Texto QUANTIDADE (Kg):
       

      //Quarto Quadro - ESPECIFICACAO DO PRODUTO
      oPrint:Line (0650,0100,0750, 0100) //coluna 1
      oPrint:Line (0650,0100,0650, 3000) //Linha 1
      oPrint:Line (0650,3000,0750, 3000) //coluna 2
      oPrint:Line (0750,0100,0750, 3000) //Linha 2
      
	  aCoords := {0655,0105,0745,2999}
      oPrint:FillRect(aCoords,oBrush)    //Tarja cinza
	  oPrint:Say  (0670,1200,"ESPECIFICA??O DO PRODUTO",oFont18)//,,CLR_HBLUE) - ESPECIFICA??O DO PRODUTO
	  
      //Quinto Quadro - Dados dos Produtos

      oPrint:Line (0750,0100,0950, 0100) //coluna 1
      oPrint:Line (0750,3000,0950, 3000) //coluna 2
      oPrint:Line (0950,0100,0950, 3000) //Linha 2
      
      oPrint:Say  (0800,0140,"C?DIGO DO ITEM:",oFont10)//,,CLR_HBLUE) - Texto C?DITO DO ITEM:
      oPrint:Line (0750,0500,0950, 0500) //coluna apos o codigo do item
      oPrint:Say  (0800,0550,"DESCRI??O DO ITEM:",oFont10)//,,CLR_HBLUE) - Texto DESCRI??O DO ITEM:
      oPrint:Line (0750,1500,0950, 1500) //coluna apos a descricao do item
      oPrint:Line (0825,1500,0825, 2550) //Linha 1 que divide Pesos Te?ricos por Metro/Unidade
      oPrint:Line (0875,1500,0875, 2550) //Linha 2 que divide Pesos Te?ricos por Metro/Unidade
      
      oPrint:Say  (0830,1600,"Minimo",oFont10)//,,CLR_HBLUE) - Texto Minimo - PESO TE?RICO (Kg/m)
      oPrint:Say  (0830,1850,"Maximo",oFont10)//,,CLR_HBLUE) - Texto Maximo - PESO TE?RICO (Kg/m)
      
      oPrint:Say  (0830,2100,"Minimo",oFont10)//,,CLR_HBLUE) - Texto Minimo - PESO TE?RICO (Kg/Un)
      oPrint:Say  (0830,2400,"Maximo",oFont10)//,,CLR_HBLUE) - Texto Maximo - PESO TE?RICO (Kg/Un)
      
      oPrint:Line (0825,1775,0950, 1775) //coluna que divide Peso Te?ricos por Metro M?nimo/M?ximo
      oPrint:Line (0825,2300,0950, 2300) //coluna que divide Peso Te?ricos por Unidade M?nimo/M?ximo
      
      oPrint:Say  (0770,1550,"PESO TE?RICO (Kg/m):",oFont10)//,,CLR_HBLUE) - Texto PESO TE?RICO (Kg/m):
      oPrint:Line (0750,2050,0950, 2050) //coluna apos o peso te?rico (Kg/m)
      oPrint:Say  (0770,2100,"PESO TE?RICO (Kg/Un):",oFont10)//,,CLR_HBLUE) - Texto PESO TE?RICO (Kg/Un):
      oPrint:Line (0750,2550,0950, 2550) //coluna apos o peso te?rico (Kg/Un)
      oPrint:Say  (0770,2600,"PE?AS/AMARRADO:",oFont10)//,,CLR_HBLUE) - Texto PE?AS/AMARRADO:

      
      oPrint:Say  (0875,0120,cCodItCrt,oFont12)//,,CLR_HBLUE) - Codigo do Item do Certificado
      oPrint:Say  (0875,0550,cDscItem,oFont12)//,,CLR_HBLUE) - Descri??o do Item do Certificado

      oPrint:Say  (0895,1525,Transform(nPsoTMt,"@E 999,999,999.999"),oFont12)//,,CLR_HBLUE) - Peso Te?rico M?nimo por Metro (Kg/m)
      oPrint:Say  (0895,2050,Transform(nPsoUnT,"@E 999,999,999.999"),oFont12)//,,CLR_HBLUE) - Peso Te?rico M?nimo por Unidade (Kg/Un)
            
      oPrint:Say  (0895,1800,Transform(nPsoNTMt,"@E 999,999,999.999"),oFont12)//,,CLR_HBLUE) - Peso Te?rico M?ximo por Metro (Kg/m)
      oPrint:Say  (0895,2300,Transform(nPsoNUnT,"@E 999,999,999.999"),oFont12)//,,CLR_HBLUE) - Peso Te?rico M?ximo por Unidade (Kg/Un)
       
      oPrint:Say  (0875,2750,Transform(nQAmarr,"@E 999,999,999,999"),oFont12)//,,CLR_HBLUE) - Quantidade de Pe?as do Amarrado

      //Sexto Quadro - Acabamentos - Ensaios Mec?nicos - Anaise Visual - Observa??es
      //               Origem - Dimensoes - Espessura e Comprimento

      oPrint:Line (0950,0100,1550, 0100) //coluna 1
      oPrint:Line (0950,0825,1550, 0825) //coluna 2
      oPrint:Line (0950,1187,1150, 1187) //coluna da Ordem de Producao
      oPrint:Line (0950,1550,1550, 1550) //coluna 3
      oPrint:Line (0950,1912,1150, 1912) //coluna da Espessura Parede
      oPrint:Line (1080,0825,1080, 2275) //Linha 1 da Ordem, Dimensoes,Espessura e Comprimento
      oPrint:Line (1150,0825,1150, 2275) //Linha 2 da Ordem, Dimensoes,Espessura e Comprimento
      oPrint:Line (0950,2275,1550, 2275) //coluna 4
      oPrint:Line (0950,3000,1550, 3000) //coluna 5
      oPrint:Line (1550,0100,1550, 3000) //Linha 2
      
      oPrint:Say  (1080,0340,"ACABAMENTOS:",oFont10)//,,CLR_HBLUE) - Texto ACABAMENTOS:
      oPrint:Say  (0970,0940,"ORDEM",oFont10)//,,CLR_HBLUE) - Texto ORDEM
      oPrint:Say  (1000,0910,"PRODU??O",oFont10)//,,CLR_HBLUE) - Texto PRODUCAO
      
      oPrint:Say  (0980,1230,"DIMENS?ES (mm):",oFont10)//,,CLR_HBLUE) - Texto DIMENS?ES (mm)
      oPrint:Say  (0970,1620,"ESPESSURA",oFont10)//,,CLR_HBLUE) - Texto ESPESSURA
      oPrint:Say  (1000,1620,"PAREDE(mm):",oFont10)//,,CLR_HBLUE) - Texto PAREDE(mm):
      
      oPrint:Say  (0970,1970,"COMPRIMENTO",oFont10)//,,CLR_HBLUE) - Texto COMPRIMENTO
      oPrint:Say  (1000,2035,"(mm)",oFont10)//,,CLR_HBLUE) - Texto (mm)
      
      oPrint:Say  (0970,2285,"OBSERVA??ES:",oFont10)//,,CLR_HBLUE) - Texto COMPRIMENTO
      
      oPrint:Say  (1105,0890,cOrdProd,oFont10)//,,CLR_HBLUE) - Numero da Ordem de Producao
      
      //oPrint:Say  (1105,1350,Transform(nDimenso,"@E 999,999,999.999"),oFont10)//,,CLR_HBLUE) - Dimens?es
      oPrint:Say  (1105,1350,cDimenso,oFont10)//,,CLR_HBLUE) - Dimens?es
      
      //oPrint:Say  (1105,1700,Transform(nEspessu,"@E 999,999,999.999"),oFont10)//,,CLR_HBLUE) - Espessura
      oPrint:Say  (1105,1700,cEspessu,oFont10)//,,CLR_HBLUE) - Espessura
      
      //oPrint:Say  (1105,2055,Transform(nComprim,"@E 999,999,999.999"),oFont10)//,,CLR_HBLUE) - Comprimento
      oPrint:Say  (1105,2055,cComprim,oFont10)//,,CLR_HBLUE) - Comprimento
      
      oPrint:Say  (1170,1070,"ENSAIOS MEC?NICOS",oFont10)//,,CLR_HBLUE) - ENSAIOS MEC?NICOS
      
      oPrint:Say  (1170,1812,"ANALISE VISUAL",oFont10)//,,CLR_HBLUE) - ANALISE VISUAL
      
      fCaixa(lSupOlea,1230,0170,1) 
      oPrint:Say  (1230,0170,"SUPERFICIE OLEADA",oFont10)//,,CLR_HBLUE) - Texto SUPERFICIE OLEADA
      
      fCaixa(lAcht90,1230,0890,4)
      oPrint:Say  (1230,0890,"ACHATAMENTO 90'",oFont10)//,,CLR_HBLUE) - Texto ACHATAMENTO 90'
      
      fCaixa(lEmpCnf,1230,1615,8)
      oPrint:Say  (1230,1615,"EMPENAMENTO CONFORME'",oFont10)//,,CLR_HBLUE) - Texto EMPENAMENTO CONFORME
      
      fCaixa(lExtSerr,1310,0170,2) 
      oPrint:Say  (1310,0170,"EXTREMIDADE COM CORTE SERRA",oFont10)//,,CLR_HBLUE) - Texto EXTREMIDADE COM CORTE SERRA
      
      fCaixa(lAcht0,1310,0890,5)
      oPrint:Say  (1310,0890,"ACHATAMENTO 0'",oFont10)//,,CLR_HBLUE) - Texto ACHATAMENTO 0'
      
      fCaixa(lTorCnf,1310,1615,9)
      oPrint:Say  (1310,1615,"TOR??O CONFORME'",oFont10)//,,CLR_HBLUE) - Texto TOR??O CONFORME
      
      fCaixa(lExtFaca,1390,0170,3) 
      oPrint:Say  (1390,0170,"EXTREMIDADE COM CORTE FACA",oFont10)//,,CLR_HBLUE) - Texto EXTREMIDADE COM CORTE FACA
      
      fCaixa(lAlrga,1390,0890,6)
      oPrint:Say  (1390,0890,"ALARGAMENTO'",oFont10)//,,CLR_HBLUE) - Texto ALARGAMENTO
      
      fCaixa(lCrva,1470,0890,7)
      oPrint:Say  (1470,0890,"CURVAMENTO'",oFont10)//,,CLR_HBLUE) - Texto CURVAMENTO
            
      //S?timo Quadro - COMPOSI??O QU?MICA

      oPrint:Line (1600,0100,1680, 0100) //coluna 1
      oPrint:Line (1600,0100,1600, 3000) //Linha 1
      oPrint:Line (1600,3000,1680, 3000) //coluna 2
      oPrint:Line (1680,0100,1680, 3000) //Linha 2
                  
	  //aCoords := {2105,0830,2199,2274}
	  //oPrint:Say  (2105,1200,"COMPOSI??O QU?MICA",oFont24)//,,CLR_HBLUE) - COMPOSI??O QU?MICA

	  aCoords := {1605,0103,1679,2999}
      oPrint:FillRect(aCoords,oBrush)    //Tarja cinza
	  oPrint:Say  (1605,1200,"COMPOSI??O QU?MICA",oFont18)//,,CLR_HBLUE) - COMPOSI??O QU?MICA

      oPrint:Line (1680,0100,1780, 0100) //coluna 1
      oPrint:Line (1680,0310,1780, 0310) //coluna Lote Fornecedor
      
      oPrint:Line (1680,0450,1780, 0450) //coluna carbono
      oPrint:Line (1680,0590,1780, 0590) //coluna mangan?s
      oPrint:Line (1680,0730,1780, 0730) //coluna Fosforo
      oPrint:Line (1680,0870,1780, 0870) //coluna Enxofre
      
      oPrint:Line (1680,1010,1780, 1010) //coluna Silicio
      oPrint:Line (1680,1150,1780, 1150) //coluna Cromo
      oPrint:Line (1680,1290,1780, 1290) //coluna Niquel
      oPrint:Line (1680,1430,1780, 1430) //coluna Cobre
      oPrint:Line (1680,1570,1780, 1570) //coluna Aluminio

      oPrint:Line (1680,1775,1780, 1775) //coluna 2 Lote Fornecedor
      
      oPrint:Line (1680,1911,1780, 1911) //coluna 2 carbono
      oPrint:Line (1680,2047,1780, 2047) //coluna 2 mangan?s
      oPrint:Line (1680,2183,1780, 2183) //coluna 2 Fosforo
      oPrint:Line (1680,2319,1780, 2319) //coluna 2 Enxofre
      
      oPrint:Line (1680,2455,1780, 2455) //coluna 2 Silicio
      oPrint:Line (1680,2591,1780, 2591) //coluna 2 Cromo
      oPrint:Line (1680,2727,1780, 2727) //coluna 2 Niquel
      oPrint:Line (1680,2863,1780, 2863) //coluna 2 Cobre
      oPrint:Line (1680,3000,1780, 3000) //coluna 2 Aluminio

      oPrint:Line (1780,0100,1780, 3000) //Linha 2
            
      //oPrint:Say  (2230,0850,"AUTO DA BOBINA",oFont10)//,,CLR_HBLUE) - Texto AUTO DA BOBINA

	  aCoords := {1685,0103,1779,309}
      oPrint:FillRect(aCoords,oBrush)    //Tarja cinza
      oPrint:Say  (1700,0175,"LOTE",oFont8)//,,CLR_HBLUE) - Texto LOTE //AUTO DA BOBINA
      oPrint:Say  (1730,0110,"FORNECEDOR",oFont8)//,,CLR_HBLUE) - Texto FORNECEDOR
      
      oPrint:Say  (1700,0380,"C",oFont10)//,,CLR_HBLUE) - Texto C
      //oPrint:Say  (2250,1210,"X-100",oFont10)//,,CLR_HBLUE) - Texto X-100
      oPrint:Say  (1700,0520,"Mn",oFont10)//,,CLR_HBLUE) - Texto Mn
      //oPrint:Say  (2250,1500,"X-100",oFont10)//,,CLR_HBLUE) - Texto X-100
      oPrint:Say  (1700,0660,"P",oFont10)//,,CLR_HBLUE) - Texto P
      //oPrint:Say  (2250,1780,"X-1000",oFont10)//,,CLR_HBLUE) - Texto X-1000
      oPrint:Say  (1700,0800,"S",oFont10)//,,CLR_HBLUE) - Texto S
      //oPrint:Say  (2250,2050,"X-1000",oFont10)//,,CLR_HBLUE) - Texto X-1000
        
      oPrint:Say  (1700,0940,"Si",oFont10)//,,CLR_HBLUE) - Texto Ni
      oPrint:Say  (1700,1080,"Cr",oFont10)//,,CLR_HBLUE) - Texto Cr
      oPrint:Say  (1700,1220,"Ni",oFont10)//,,CLR_HBLUE) - Texto Ni
      oPrint:Say  (1700,1360,"Cu",oFont10)//,,CLR_HBLUE) - Texto Cu
      oPrint:Say  (1700,1500,"Al",oFont10)//,,CLR_HBLUE) - Texto Al


	  aCoords := {1685,1573,1779,1774}
      oPrint:FillRect(aCoords,oBrush)    //Tarja cinza
      
      If Len(aCompQuimi) > 1
      
         oPrint:Say  (1700,1650,"LOTE",oFont8)//,,CLR_HBLUE) - Texto LOTE //AUTO DA BOBINA
         oPrint:Say  (1730,1580,"FORNECEDOR",oFont8)//,,CLR_HBLUE) - Texto FORNECEDOR

         oPrint:Say  (1700,1843,"C",oFont10)//,,CLR_HBLUE) - Texto C
         //oPrint:Say  (2250,1210,"X-100",oFont10)//,,CLR_HBLUE) - Texto X-100
         oPrint:Say  (1700,1979,"Mn",oFont10)//,,CLR_HBLUE) - Texto Mn
         //oPrint:Say  (2250,1500,"X-100",oFont10)//,,CLR_HBLUE) - Texto X-100
         oPrint:Say  (1700,2115,"P",oFont10)//,,CLR_HBLUE) - Texto P
         //oPrint:Say  (2250,1780,"X-1000",oFont10)//,,CLR_HBLUE) - Texto X-1000
         oPrint:Say  (1700,2251,"S",oFont10)//,,CLR_HBLUE) - Texto S
         //oPrint:Say  (2250,2050,"X-1000",oFont10)//,,CLR_HBLUE) - Texto X-1000
        
         oPrint:Say  (1700,2387,"Si",oFont10)//,,CLR_HBLUE) - Texto Ni
         oPrint:Say  (1700,2523,"Cr",oFont10)//,,CLR_HBLUE) - Texto Cr
         oPrint:Say  (1700,2659,"Ni",oFont10)//,,CLR_HBLUE) - Texto Ni
         oPrint:Say  (1700,2795,"Cu",oFont10)//,,CLR_HBLUE) - Texto Cu
         oPrint:Say  (1700,2931,"Al",oFont10)//,,CLR_HBLUE) - Texto Al
      
      EndIf
            
      //Imprime as Composi??es Qu?micas dos Lotes de Fornecedores composto neste Produto
      nBoxLin:= 1780
      nLinha := 1790 //2350
      lImp2Col := .T.
      For nCQ := 1 To Len(aCompQuimi)
          
         lImp2Col := .T.
                
         cLteForn  := aCompQuimi[nCQ,1]
         nCarbono  := aCompQuimi[nCQ,2]
         nManganes := aCompQuimi[nCQ,3]
         nFosforo  := aCompQuimi[nCQ,4]
         nEnxofre  := aCompQuimi[nCQ,5]

         nSilicio := aCompQuimi[nCQ,6]
         nCromo   := aCompQuimi[nCQ,7]
         nNiquel  := aCompQuimi[nCQ,8]
         nCobre   := aCompQuimi[nCQ,9]
         nAlumini := aCompQuimi[nCQ,10]

         nColCrt  := nBoxLin + 50
         n2ColCrt := nBoxLin + 50 //1780
         
	     aCoords := {(nBoxLin + 5),0103,(nColCrt - 1),309}
         oPrint:FillRect(aCoords,oBrush)    //Tarja cinza
         oPrint:Line (nBoxLin,0100,nColCrt, 0100) //coluna 1
         oPrint:Line (nBoxLin,0310,nColCrt, 0310) //coluna Lote Fornecedor
         
         oPrint:Line (nBoxLin,0450,nColCrt, 0450) //coluna carbono
         oPrint:Line (nBoxLin,0590,nColCrt, 0590) //coluna mangan?s
         oPrint:Line (nBoxLin,0730,nColCrt, 0730) //coluna Fosforo
         oPrint:Line (nBoxLin,0870,nColCrt, 0870) //coluna Enxofre
      
         oPrint:Line (nBoxLin,1010,nColCrt, 1010) //coluna Silicio
         oPrint:Line (nBoxLin,1150,nColCrt, 1150) //coluna Cromo
         oPrint:Line (nBoxLin,1290,nColCrt, 1290) //coluna Niquel
         oPrint:Line (nBoxLin,1430,nColCrt, 1430) //coluna Cobre
         oPrint:Line (nBoxLin,1570,nColCrt, 1570) //coluna Aluminio

         oPrint:Line (nBoxLin,1775,n2ColCrt, 1775) //coluna 2 Lote Fornecedor
      
         oPrint:Line (nBoxLin,1911,n2ColCrt, 1911) //coluna 2 carbono
         oPrint:Line (nBoxLin,2047,n2ColCrt, 2047) //coluna 2 mangan?s
         oPrint:Line (nBoxLin,2183,n2ColCrt, 2183) //coluna 2 Fosforo
         oPrint:Line (nBoxLin,2319,n2ColCrt, 2319) //coluna 2 Enxofre
      
         oPrint:Line (nBoxLin,2455,n2ColCrt, 2455) //coluna 2 Silicio
         oPrint:Line (nBoxLin,2591,n2ColCrt, 2591) //coluna 2 Cromo
         oPrint:Line (nBoxLin,2727,n2ColCrt, 2727) //coluna 2 Niquel
         oPrint:Line (nBoxLin,2863,n2ColCrt, 2863) //coluna 2 Cobre
         oPrint:Line (nBoxLin,3000,n2ColCrt, 3000) //coluna 2 Aluminio
      
         oPrint:Line (nBoxLin,3000,nColCrt, 3000) //coluna 6
         If Mod(nCQ,2) > 0
            
            oPrint:Say  (nLinha,0145,cLteForn,oFont10)//,,CLR_HBLUE) - Lote do Fornecedor
            oPrint:Say  (nLinha,0320,Transform(nCarbono,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Carbono
            oPrint:Say  (nLinha,0460,Transform(nManganes,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Manganes
            oPrint:Say  (nLinha,0600,Transform(nFosforo,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Fosforo
            oPrint:Say  (nLinha,0740,Transform(nEnxofre,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Enxofre

            oPrint:Say  (nLinha,0880,Transform(nSilicio,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Silicio
            oPrint:Say  (nLinha,1020,Transform(nCromo,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cromo
            oPrint:Say  (nLinha,1160,Transform(nNiquel,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Niquel
            oPrint:Say  (nLinha,1300,Transform(nCobre,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cobre
            oPrint:Say  (nLinha,1440,Transform(nAlumini,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Aluminio
            
            cLteForn := " "
            nCarbono  := 0
            nManganes := 0
            nFosforo  := 0
            nEnxofre  := 0

            nSilicio := 0
            nCromo   := 0
            nNiquel  := 0
            nCobre   := 0
            nAlumini := 0
	     
	     Else
	        
	        If ( nCarbono + nManganes + nFosforo + nSilicio + nCromo + nNiquel + nCobre + nAlumini ) > 0
	           
	           lImp2Col := .F.
	     
               oPrint:Say  (nLinha,1580,cLteForn,oFont10)//,,CLR_HBLUE) - Lote do Fornecedor
               oPrint:Say  (nLinha,1785,Transform(nCarbono,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Carbono
               oPrint:Say  (nLinha,1921,Transform(nManganes,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Manganes
               oPrint:Say  (nLinha,2057,Transform(nFosforo,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Fosforo
               oPrint:Say  (nLinha,2193,Transform(nEnxofre,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Enxofre

               oPrint:Say  (nLinha,2329,Transform(nSilicio,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Silicio
               oPrint:Say  (nLinha,2465,Transform(nCromo,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cromo
               oPrint:Say  (nLinha,2601,Transform(nNiquel,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Niquel
               oPrint:Say  (nLinha,2737,Transform(nCobre,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cobre
               oPrint:Say  (nLinha,2873,Transform(nAlumini,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Aluminio

               cLteForn := " "
               nCarbono  := 0
               nManganes := 0
               nFosforo  := 0
               nEnxofre  := 0

               nSilicio := 0
               nCromo   := 0
               nNiquel  := 0
               nCobre   := 0
               nAlumini := 0
            
	           nLinha := nLinha + 50
	           nBoxLin := nBoxLin + 50
               oPrint:Line (nBoxLin,0100,nBoxLin, 3000) //Linha 2
            
            EndIf
         	        
	     EndIf
      
      Next nCQ
      
	  If Len(aCompQuimi) > 0 .And. lImp2Col
         /*
         oPrint:Say  (nLinha,1580,cLteForn,oFont10)//,,CLR_HBLUE) - Lote do Fornecedor
         oPrint:Say  (nLinha,1785,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Carbono
         oPrint:Say  (nLinha,1921,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Manganes
         oPrint:Say  (nLinha,2057,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Fosforo
         oPrint:Say  (nLinha,2193,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Enxofre

         oPrint:Say  (nLinha,2329,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Silicio
         oPrint:Say  (nLinha,2465,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cromo
         oPrint:Say  (nLinha,2601,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Niquel
         oPrint:Say  (nLinha,2737,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Cobre
         oPrint:Say  (nLinha,2873,Transform(0,"@E 999.999"),oFont10)//,,CLR_HBLUE) - Quantidade de Aluminio	           
         */

	     nLinha := nLinha + 50
	     nBoxLin := nBoxLin + 50
         oPrint:Line (nBoxLin,0100,nBoxLin, 3000) //Linha 2

      EndIf
      
      
	  If nLinha > 2300 //2500
	     oPrint:EndPage()     // Finaliza a p?gina
	     oPrint:StartPage()   // Inicia uma nova p?gina
	     nLinha  := 80
      EndIf

   Ms_Flush()
   oPrint:EndPage()     // Finaliza a p?gina
   oPrint:Preview()     // Visualiza antes de imprimir
               
   RestArea(cAreaKi)
   
Return

User Function LegSNOE()
   Local aLegenda := {	{"BR_VERDE"		,"N?o Emitido" }   ,;
					    {"BR_VERMELHO"  ,"J? Emitido"	  }   }
   BrwLegenda(cCadastro,"Legenda",aLegenda)
Return(.T.)

//Static Function fCaixa(lMarca)
Static Function fCaixa(lMarca,nLnha,nCl,nOpCxCr) 
/*
      oPrint:Line (1685,0115,1685, 0155) //Linha 1
      oPrint:Line (1685,0115,1715, 0115) //coluna 1
      oPrint:Line (1685,0155,1715, 0155) //coluna 2
      oPrint:Line (1715,0115,1715, 0155) //Linha 2
      If lMarca
         oPrint:Line (1685,0115,1715, 0155) //Linha 1 Vertical para Baixo
         oPrint:Line (1685,0155,1715, 0115) //Linha 2 Vertical para Baixo
      EndIf
*/

   //oPrint:Say  (1680,0170,"SUPERFICIE OLEADA",oFont10)//,,CLR_HBLUE) - Texto SUPERFICIE OLEADA
   //Caixa Marcada
   nCx1Linha := nLnha + 5
   nCx2Linha := nCx1Linha + 30
   nCx1Col   := nCl - 55  //0115
   nCx2Col   := nCx1Col + 40 //0155
     
   oPrint:Line (nCx1Linha,nCx1Col,nCx1Linha, nCx2Col) //Linha 1
   oPrint:Line (nCx1Linha,nCx1Col,nCx2Linha, nCx1Col) //coluna 1
   oPrint:Line (nCx1Linha,nCx2Col,nCx2Linha, nCx2Col) //coluna 2
   oPrint:Line (nCx2Linha,nCx1Col,nCx2Linha, nCx2Col) //Linha 2
   If lMarca
      oPrint:Line (nCx1Linha,nCx1Col,nCx2Linha, nCx2Col) //Linha 1 Vertical para Baixo
      oPrint:Line (nCx1Linha,nCx2Col,nCx2Linha, nCx1Col) //Linha 2 Vertical para Baixo
   EndIf
   
Return