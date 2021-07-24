#INCLUDE "FONT.CH"
#INCLUDE "COLORS.CH"
#INCLUDE "PROTHEUS.CH"

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �VLDCGCCTB �Autor  �Five Solutions      � Data �  02/07/2008 ���
�������������������������������������������������������������������������͹��
���Desc.     � Verifica��o do CNPJ/CPF Digitado                           ���
���          �                                                            ���
�������������������������������������������������������������������������͹��
���Uso       � COMAFAL PE/RS/SP                                          ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function VldCGCCTB
	
	lRetVld := .T.
	cNomeFunc := Alltrim(FunName())
	
	If cNomeFunc == "MATA030"
	   cCNPJ   := Alltrim(M->A1_CGC)
	   //A Regra Abaixo foi extra�da do Gatilho A1_CGC Sequ�ncia 011.
	   cCtaCtb := IIF(M->A1_PESSOA=="J","11201001"+SUBS(M->A1_CGC,6,3)+SUBS(M->A1_CGC,13,2),"11201001"+SUBS(M->A1_CGC,7,5))
	Else
	   cCNPJ   := Alltrim(M->A2_CGC) 
	   //A Regra Abaixo foi extra�da do Gatilho A2_CGC Sequ�ncia 009.
	   cCtaCtb := IIF(M->A2_TIPO=="J","21102001"+SUBS(M->A2_CGC,6,3)+SUBS(M->A2_CGC,13,2),"21102001"+SUBS(M->A2_CGC,7,5))
	EndIf
	
	cCtaExit:= ""
	cAreaAntes := Alias()
	DbSelectArea("CT1")
	DbSetOrder(1)
	If DbSeek(xFilial("CT1")+cCtaCtb)
	   cCtaExit := CT1_CONTA 
	EndIf
	
	
	DbSelectArea(cAreaAntes)
   	DEFINE MSDIALOG oDlgCGC FROM 0,0 TO 250,700 TITLE "Valida��o do CNPJ/Cta. Cont�bil" Of oMainWnd PIXEL
	DEFINE FONT oFnt01 NAME "Arial" SIZE 0, -13 BOLD
	DEFINE FONT oBold NAME "Arial" SIZE 0, -16 BOLD
	DEFINE FONT oFtTmp NAME "Arial" SIZE 0, -58 BOLD
	
	@ 000,000 BITMAP oBmp RESNAME "LOGIN" oF oDlgCGC SIZE 60,150 NOBORDER WHEN .F. PIXEL

	@ 011,040 TO 013,300 LABEL '' OF oDlgCGC PIXEL
	@ 003,050 SAY "Verifique o CNPJ/CPF Digitado" Of oDlgCGC PIXEL SIZE 130 ,9 FONT oFnt01

	//@ 012,130 SAY "Inicio "+DTOC(DtIniExp)+" as "+HrIniExp+" hs / T�rmino "+DTOC(Date())+" as "+Time()+" hs" Of oDlgCGC PIXEL SIZE 200,88 FONT oFtTmp

	@ 023,55 SAY TRANSFORM(cCNPJ,IIf(Len(cCNPJ)>=14,"@R 99.999.999/9999-99","@R 999.999.999-99")) FONT oFtTmp COLOR CLR_HBLUE Of oDlgCGC PIXEL SIZE 300,88 //50,88
	If !Empty(cCtaExit)
	   @ 052,50 SAY "A Conta Cont�bil "+Alltrim(cCtaExit)+" j� existe no Plano de Contas !!!" FONT oBold COLOR CLR_HRED Of oDlgCGC PIXEL SIZE 200,88
	Else
	   @ 052,50 SAY "Conta "+Alltrim(cCtaCtb)+" Dispon�vel para inclus�o no Plano de Contas !!!" FONT oBold Of oDlgCGC PIXEL SIZE 270,88
	EndIf
	
    TButton():New(110,270,"Confirma", oDlgCGC,{||oDlgCGC:End()},32,10,,oDlgCGC:oFont,.F.,.T.,.F.,,.F.,,,.F.)
    TButton():New(110,315,"Cancela ", oDlgCGC,{||RunDel()},32,10,,oDlgCGC:oFont,.F.,.T.,.F.,,.F.,,,.F.)

	ACTIVATE MSDIALOG oDlgCGC CENTERED	
Return(lRetVld)
Static Function RunDel
   lRetVld   := .F.   
   oDlgCGC:End()
Return(lRetVld)