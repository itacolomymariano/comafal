#include "rwmake.ch"        

/*���������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  � CMB005   � Autor � Valmir Belchior       � Data � 14/05/01 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � Este Rdmake compoe as Rotinas de geracao do SISPAG, arquivo���
���          � 341REM.PAG e 341RET.PAG                                    ���
���          � Calcula o layout para o Campo Livre de Dados do Codigo de  ���
���          � Barras.                                                    ���
�������������������������������������������������������������������������Ĵ��
��� Adaptado por Jeniffer Larrosa para a versao 7.10                      ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������*/
User Function CMB005()     

// Variaveis Private
PRIVATE cCampFree

If	Len(Alltrim(SE2->E2_X_CODBA)) == 44
	cCampFree := Substr(SE2->E2_X_CODBA,20,25)
ElseIF Len(Alltrim(SE2->E2_X_CODBA)) == 47
	cCampFree := Substr(SE2->E2_X_CODBA,5,5) + Substr(SE2->E2_X_CODBA,11,10) + Substr(SE2->E2_X_CODBA,22,10)
ElseIf Len(Alltrim(SE2->E2_X_CODBA)) >= 36 .AND. Len(Alltrim(SE2->E2_X_CODBA)) <= 40
	cCampFree := Substr(SE2->E2_X_CODBA,5,5) + Substr(SE2->E2_X_CODBA,11,10) + Substr(SE2->E2_X_CODBA,22,10)
Else
	cCampFree := Replicate("0",25)
EndIf

Return cCampFree