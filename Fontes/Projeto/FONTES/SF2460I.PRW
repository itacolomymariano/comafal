#include "rwmake.ch"     

/*
���������������������������������������������������������������������������
���������������������������������������������������������������������������
�����������������������������������������������������������������������Ŀ��
���Fun�ao    � SF2460I  � Autor � Zanardo             � Data � 19.11.03 ���
�����������������������������������������������������������������������Ĵ��
���Descri��o � Ponto de entrada na gravacao do SF2.                     ���
�����������������������������������������������������������������������Ĵ��
���Cliente   � Comafal                                                  ���
������������������������������������������������������������������������ٱ�
���������������������������������������������������������������������������
���������������������������������������������������������������������������
*/

User Function SF2460I()  

Local aArea := GetArea()
DbSelectArea("SD2")
DbSetOrder(3)
If DbSeek(xFILIAL("SD2")+SF2->F2_DOC+SF2->F2_SERIE+SF2->F2_CLIENTE+SF2->F2_LOJA)

	While SD2->D2_DOC+SD2->D2_SERIE+SD2->D2_CLIENTE+SD2->D2_LOJA == ;
			SF2->F2_DOC+SF2->F2_SERIE+SF2->F2_CLIENTE+SF2->F2_LOJA .AND. ;
			SD2->(!EOF())
			
		DbSelectArea("SB1")
		DbSetOrder(1)
		If MsSeek(xFILIAL("SB1")+SD2->D2_COD)
			RecLock("SD2",.F.)
			SD2->D2_X_CST2 := NoRound(SB1->B1_X_CST2 * SD2->D2_QUANT,2)
			MsUnlock()
		EndIf
		SD2->(DbSkip())
	EndDo
EndIf
//���������������������������������������������������������������������������������������Ŀ
//� Grava F2_X_USERI no momento da confirmacao da Nota Fiscal - Sabrina Moreira -16/09/03 �
//�����������������������������������������������������������������������������������������
RecLock ("SF2",.F.)
SF2->F2_X_USERI :=  Upper(SubStr(cUsuario,7,15))
MsUnLock()

RestArea(aArea)

Return(NIL)      
