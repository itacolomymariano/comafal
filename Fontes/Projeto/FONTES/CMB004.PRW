#include "rwmake.ch"        

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  � CMb004   � Autor � Valmir Belchior       � Data � 14/05/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Este Rdmake compoe as Rotinas de geracao do SISPAG, arquivo���
���          � 341REM.PAG e 341RET.PAG                                    ���
���          � Calcula o layout para o DV do Codigo de Barras             ���
�������������������������������������������������������������������������Ĵ��
��� Adaptado por Jeniffer Larrosa para a versao 7.10                      ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������*/
User Function CMb004()        

// Variaveis Privadas
PRIVATE cDigCamp := 0

If	Len(Alltrim(SE2->E2_X_CODBA)) == 47
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,33,1)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) == 44		  // Alterado de 44 p/ 47. Comtempla o Fator de Vencimento.
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,5,1)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) >= 36 .and. Len(Alltrim(SE2->E2_X_CODBA)) <= 43
	cDigCamp	:=	Substr(SE2->E2_X_CODBA,33,1)
Else
	cDigCamp	:=	"0"
Endif

Return cDigCamp