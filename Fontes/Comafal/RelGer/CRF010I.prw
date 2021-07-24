/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
ฑฑษออออออออออัออออออออออหอออออออัออออออออออออออออออออหออออออัอออออออออออออปฑฑ
ฑฑบPrograma  ณRunReport บAutor  ณEduardo Zanardo     บ Data ณ  03/31/03   บฑฑ
ฑฑฬออออออออออุออออออออออสอออออออฯออออออออออออออออออออสออออออฯอออออออออออออนฑฑ
ฑฑบUso       ณ AP7                                                        บฑฑ
ฑฑศออออออออออฯออออออออออออออออออออออออออออออออออออออออออออออออออออออออออออผฑฑ
ฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑฑ
฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿฿
*/

User Function CRF010I(nLin,nMes,Cabec1,Cabec2,Titulo) 
Local aArea := GetArea() 

nLin := 0
nLin++

U_LayoutDRE()                                 

nLin := 0
nLin++

FmtLin(,{aL[01]},,,@nLin)	 		          
FmtLin({aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes],aMeses[Iif(nMes-2 == -1 .and. nMes == 1 ,11,Iif(nMes-2 == 0 .and. nMes == 2 ,12,nMes-2))],aMeses[Iif(nMes-1 == 0 .and. nMes == 1,12,nMes-1)],aMeses[nMes]},aL[02],"",,@nLin)
FmtLin(,{aL[03]},,,@nLin)	 		            
FmtLin({Transform(aFatBruto[1,2]+aFatBruto[1,3]+aSucata[1,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,2]+aFatBruto[2,3]+aSucata[2,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,2]+aFatBruto[3,3]+aSucata[3,2],"@R 99,999,999.99")},aL[04],"",,@nLin)
FmtLin({Transform(aFatBruto[1,2]+aSucata[1,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,2]+aSucata[2,2],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,2]+aSucata[3,2],"@R 99,999,999.99")},aL[05],"",,@nLin)
FmtLin({Transform(aFatBruto[1,3],"@R 99,999,999.99"),;
		Transform(aFatBruto[2,3],"@R 99,999,999.99"),;
		Transform(aFatBruto[3,3],"@R 99,999,999.99")},aL[06],"",,@nLin)
FmtLin(,{aL[07]},,,@nLin)	 		          
FmtLin(,{aL[08]},,,@nLin)	 		          
FmtLin(,{aL[09]},,,@nLin)	 		          
FmtLin(,{aL[10]},,,@nLin)	 		          
FmtLin(,{aL[11]},,,@nLin)	 		          
FmtLin({Transform(((aFatBruto[1,2]+aSucata[1,2])*9.25)/100,"@R 99,999,999.99"),;
		Transform(((aFatBruto[2,2]+aSucata[2,2])*9.25)/100,"@R 99,999,999.99"),;
		Transform(((aFatBruto[3,2]+aSucata[3,2])*9.25)/100,"@R 99,999,999.99")},aL[12],"",,@nLin)
FmtLin({Transform(aDevInt[1,2]+aDevExt[1,2],"@R 99,999,999.99"),;
		Transform(aDevInt[2,2]+aDevExt[2,2],"@R 99,999,999.99"),;
		Transform(aDevInt[3,2]+aDevExt[3,2],"@R 99,999,999.99")},aL[13],"",,@nLin) 		          
FmtLin({Transform(aDevInt[1,2],"@R 99,999,999.99"),;
		Transform(aDevInt[2,2],"@R 99,999,999.99"),;
		Transform(aDevInt[3,2],"@R 99,999,999.99")},aL[14],"",,@nLin) 		          
FmtLin({Transform(aDevExt[1,2],"@R 99,999,999.99"),;
		Transform(aDevExt[2,2],"@R 99,999,999.99"),;
		Transform(aDevExt[3,2],"@R 99,999,999.99")},aL[15],"",,@nLin) 		          

FmtLin(,{aL[16]},,,@nLin)	 		          
FmtLin(,{aL[17]},,,@nLin)	 		          
FmtLin({Transform((aFatBruto[1,2]+aFatBruto[1,3]+aSucata[1,2])-(aDevInt[1,2]+aDevExt[1,2])-(((aFatBruto[1,2]+aSucata[1,2])*9,25)/100),"@R 99,999,999.99"),;
		Transform((aFatBruto[2,2]+aFatBruto[2,3]+aSucata[2,2])-(aDevInt[2,2]+aDevExt[2,2])-(((aFatBruto[2,2]+aSucata[2,2])*9,25)/100),"@R 99,999,999.99"),;
		Transform((aFatBruto[3,2]+aFatBruto[3,3]+aSucata[3,2])-(aDevInt[3,2]+aDevExt[3,2])-(((aFatBruto[3,2]+aSucata[3,2])*9,25)/100),"@R 99,999,999.99")},aL[18],"",,@nLin) 		          
FmtLin(,{aL[19]},,,@nLin)	 		          
FmtLin(,{aL[20]},,,@nLin)	 		          
FmtLin(,{aL[21]},,,@nLin)	 		          
FmtLin(,{aL[22]},,,@nLin)	 		          
FmtLin(,{aL[23]},,,@nLin)	 		          
FmtLin(,{aL[24]},,,@nLin)	 		          
FmtLin(,{aL[25]},,,@nLin)	 		          
FmtLin(,{aL[26]},,,@nLin)	 		          
FmtLin(,{aL[27]},,,@nLin)	 		          
FmtLin(,{aL[28]},,,@nLin)	 		          
FmtLin(,{aL[29]},,,@nLin)	 		          
FmtLin(,{aL[30]},,,@nLin)	 		          
FmtLin(,{aL[31]},,,@nLin)	 		          
FmtLin(,{aL[32]},,,@nLin)	 		          
FmtLin(,{aL[33]},,,@nLin)	 		          
FmtLin(,{aL[34]},,,@nLin)	 		          
FmtLin(,{aL[35]},,,@nLin)	 		          
FmtLin(,{aL[36]},,,@nLin)	 		          
FmtLin(,{aL[37]},,,@nLin)	 		          
FmtLin(,{aL[38]},,,@nLin)	 		          
FmtLin(,{aL[39]},,,@nLin)
FmtLin(,{aL[40]},,,@nLin)	 		          
FmtLin(,{aL[41]},,,@nLin)	 		           
FmtLin(,{aL[42]},,,@nLin)	 		          
FmtLin(,{aL[43]},,,@nLin)	 		          
FmtLin(,{aL[44]},,,@nLin)	 		          
FmtLin(,{aL[45]},,,@nLin)	 		          
FmtLin(,{aL[46]},,,@nLin)	 		          
FmtLin(,{aL[47]},,,@nLin)	 		          
FmtLin(,{aL[48]},,,@nLin)	 		          
FmtLin(,{aL[49]},,,@nLin)	 		          
FmtLin(,{aL[50]},,,@nLin)	 		          
FmtLin(,{aL[51]},,,@nLin)	 		          

RestArea(aArea)
Return(.T.)
