/*
!short: FiveWin main Header File */

#ifndef _FIVEWIN_CH
#define _FIVEWIN_CH

#define FWVERSION    "FiveWin 1.9.2 - November 1996"
#define FWCOPYRIGHT  "(c) FiveTech, 1993-6"

#include "Dialog.ch"
#include "Font.ch"
#include "Ini.ch"
#include "Menu.ch"
#include "Print.ch"

#ifndef CLIPPER501
   #include "Colors.ch"
   #include "DLL.ch"
   #include "Folder.ch"
   #include "Objects.ch"
   #include "ODBC.ch"
   #include "DDE.ch"
   #include "Video.ch"
   #include "VKey.ch"
   #include "Tree.ch"
//   #include "WinApi.ch"
#endif

#define CRLF Chr(13)+Chr(10)

/*----------------------------------------------------------------------------//
!short: Running multiple instances of a FiveWin EXE */

#xcommand SET MULTIPLE <on:ON,OFF> => SetMultiple( Upper(<(on)>) == "ON" )

/*----------------------------------------------------------------------------//
!short: ACCESSING / SETTING Variables */

#xtranslate bSETGET(<uVar>) => ;
            { | u | If( PCount() == 0, <uVar>, <uVar> := u ) }

/*----------------------------------------------------------------------------//
!short: Default parameters management */

#xcommand DEFAULT <uVar1> := <uVal1> ;
               [, <uVarN> := <uValN> ] => ;
                  <uVar1> := If( <uVar1> == nil, <uVal1>, <uVar1> ) ;;
                [ <uVarN> := If( <uVarN> == nil, <uValN>, <uVarN> ); ]

/*----------------------------------------------------------------------------//
!short: DO ... UNTIL support */

#xcommand DO            => while .t.
#xcommand UNTIL <uExpr> => if <uExpr>; exit; end; end

/*----------------------------------------------------------------------------//
!short: Idle periods management */

#xcommand SET IDLEACTION TO <uIdleAction> => SetIdleAction( <{uIdleAction}> )

/*----------------------------------------------------------------------------//
!short: DataBase Objects */

#xcommand DATABASE <oDbf> => <oDbf> := TDataBase():New()

/*----------------------------------------------------------------------------//
!short: General release command */

#xcommand RELEASE <ClassName> <oObj1> [,<oObjN>] ;
       => ;
          <oObj1>:End() ; <oObj1> := nil ;
        [ ; <oObjN>:End() ; <oObjN> := nil ]

/*----------------------------------------------------------------------------//
!short: Brushes */

#xcommand DEFINE BRUSH [ <oBrush> ] ;
             [ STYLE <cStyle> ] ;
             [ COLOR <nRGBColor> ] ;
             [ <file:FILE,FILENAME,DISK> <cBmpFile> ] ;
             [ <resource:RESOURCE,NAME,RESNAME> <cBmpRes> ] ;
       => ;
          [ <oBrush> := ] TBrush():New( [ Upper(<(cStyle)>) ], <nRGBColor>,;
             <cBmpFile>, <cBmpRes> )

#xcommand SET BRUSH ;
             [ OF <oWnd> ] ;
             [ TO <oBrush> ] ;
       => ;
          <oWnd>:SetBrush( <oBrush> )

/*----------------------------------------------------------------------------//
!short: Pens */

#xcommand DEFINE PEN <oPen> ;
             [ STYLE <nStyle> ] ;
             [ WIDTH <nWidth> ] ;
             [ COLOR <nRGBColor> ] ;
             [ <of: OF, WINDOW, DIALOG> <oWnd> ] ;
       => ;
          <oPen> := TPen():New( <nStyle>, <nWidth>, <nRGBColor>, <oWnd> )

#xcommand ACTIVATE PEN <oPen> => <oPen>:Activate()

/*----------------------------------------------------------------------------//
!short: ButtonBar Commands */

#xcommand DEFINE BUTTONBAR [ <oBar> ] ;
             [ <size: SIZE, BUTTONSIZE, SIZEBUTTON > <nWidth>, <nHeight> ] ;
             [ <_3d: 3D, 3DLOOK> ] ;
             [ <mode: TOP, LEFT, RIGHT, BOTTOM, FLOAT> ] ;
             [ <wnd: OF, WINDOW, DIALOG> <oWnd> ] ;
             [ CURSOR <oCursor> ] ;
      => ;
         [ <oBar> := ] TBar():New( <oWnd>, <nWidth>, <nHeight>, <._3d.>,;
             [ Upper(<(mode)>) ], <oCursor> )

#xcommand @ <nRow>, <nCol> BUTTONBAR [ <oBar> ] ;
             [ SIZE <nWidth>, <nHeight> ] ;
             [ BUTTONSIZE <nBtnWidth>, <nBtnHeight> ] ;
             [ <_3d: 3D, 3DLOOK> ] ;
             [ <mode: TOP, LEFT, RIGHT, BOTTOM, FLOAT> ] ;
             [ <wnd: OF, WINDOW, DIALOG> <oWnd> ] ;
             [ CURSOR <oCursor> ] ;
      => ;
         [ <oBar> := ] TBar():NewAt( <nRow>, <nCol>, <nWidth>, <nHeight>,;
             <nBtnWidth>, <nBtnHeight>, <oWnd>, <._3d.>, [ Upper(<(mode)>) ],;
             <oCursor> )

