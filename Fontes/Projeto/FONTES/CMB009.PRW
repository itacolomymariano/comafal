#include "rwmake.ch"        // incluido pelo assistente de conversao do AP5 IDE

User Function 	CMB009()        // incluido pelo assistente de conversao do AP5 IDE

SetPrvt("_cValAbat,	_cRetLiq")

/*/
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������Ŀ��
���Programa  � CMB009   � Autor � Alexandre Miguel      � Data � 02/10/03 ���
�������������������������������������������������������������������������Ĵ��
���Descri��o � ExecBlock para Gravar o Valor do Titulo sem abatimento p/  ���
���          � titulos das Concessionarias - SISPAG - 115/129- Descontos  ���
�������������������������������������������������������������������������Ĵ��
��� Uso      � Adaptado por Sabrina Moreira para a Versao 7.10            ���
��������������������������������������������������������������������������ٱ�
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/


_cValAbat	:= 0
_cRetLiq	:= 0

_cValAbat	:= SOMAABAT(E2_PREFIXO,E2_NUM,E2_PARCELA,"P",1,DDATABASE,E2_FORNECE,E2_LOJA)
_cRetLiq	:= SE2->E2_VALOR - SE2->E2_DECRESC 
_cRetLiq	:= Strzero(_cRetLiq * 100,15)


Return(_cRetLiq)        
