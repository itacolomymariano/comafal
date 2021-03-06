#include "rwmake.ch"

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  � CMb001   � Autor � Valmir Belchior       � Data � 29/05/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � ExecBlock para Gravar Agencia e Conta do Sispag Itau       ���
���          � pois ha dif.de lay-out quando bancos forem diferentes      ���
�������������������������������������������������������������������������Ĵ��
��� Adaptado por Jeniffer Larrosa para a versao 7.10                      ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������*/
User Function CMb001()


// Variaveis Locais
LOCAL cConta	:=	""
LOCAL nConta	:=	0
LOCAL cDiv		:=	""

// Variaveis Privadas
PRIVATE cReturn := ""

cConta	:=	AllTrim(SA2->A2_NUMCON)		// Numero da Conta
cDiV	:=	Right(cConta,1)				// Digito da conta
nConta	:=	Val(Subs(cConta,1,Len(cConta)-1))
cReturn :=	StrZero(Val(AllTrim(SA2->A2_AGENCIA)),5) + " " + StrZero(nConta,12) + " " + cDiV

Return cReturn                     

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  � CMb001A  � Autor � Zanardo               � Data � 07/05/04 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Retorna o saldo do titulo considerando os abatimentos      ���
���������������������������������������������������������������������i����Ĵ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������*/
User Function CMb001A(cPrefixo,cNumero,cParcela,cCarteira,cCliFor,cLoja,cTipo,cFilTit,cNatureza)

Local nSaldo 	:= 0
Local aArea 	:= GetArea()
Local nAbat		:= SomaAbat(cPrefixo,cNumero,cParcela,cCarteira,1,,cCliFor,cLoja)
Local nVlTit	:= SaldoTit(cPrefixo,cNumero,cParcela,cTipo,cNatureza,cCarteira,cCliFor,1,,,cLoja,cFilTit)

RestArea(aArea)

Return(nVlTit-nAbat)

