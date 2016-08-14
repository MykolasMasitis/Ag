PROCEDURE ag
#INCLUDE INCLUDE\MAIN.H
_SCREEN.Visible = .f.
ON SHUTDOWN CLEAR EVENTS 

PUBLIC gcOldDir, gcOldPath, gcOldClassLib, gcOldEscape, gcOldTalk
PUBLIC oApp, IsFvpDll

PUBLIC fso AS Scripting.FileSystemObject 
fso = CREATEOBJECT('Scripting.FileSystemObject')

PUSH KEY
CLEAR
CLEAR PROGRAM
CLOSE ALL

#IF DEBUGMODE
	SET RESOURCE ON
#ELSE
	SET RESOURCE OFF
	SET VIEW OFF
#ENDIF

IF SET('TALK') = 'ON'
	SET TALK OFF
	gcOldTalk = 'ON'
ELSE
	gcOldTalk = 'OFF'
ENDIF

gcOldEscape     = SET('ESCAPE')
gcOldDir        = FULLPATH(CURDIR())
gcOldPath       = SET('PATH')
gcOldClassLib   = SET('CLASSLIB')
plOldFlagEnd    = .T.
plOldFirstStart = .T.

SET SAFETY OFF 
SET STATUS BAR OFF 

_TOOLTIPTIMEOUT = 0
_INCSEEK = 5
_DBLCLICK = 1.5

LOCAL i, lcSys16, lcProgram

IF TYPE("_VFP.Projects.Count") = "N"
	FOR m.i = 1 TO _VFP.Projects.Count
		IF TYPE("_VFP.Projects(m.i)") = "O"
			_VFP.Projects(m.i).Close()
		ENDIF
	ENDFOR
ENDIF


m.lcSys16 = SYS(16)
m.lcProgram = SUBSTR(m.lcSys16, AT(":", m.lcSys16) - 1)
CD LEFT(m.lcProgram, RAT("\", m.lcProgram))
*-- Åñëè ñòàðòîâàëè èç START.PRG, òî ïåðåõîäèì â ãîëîâíóþ äèðåêòîðèþ
IF (RIGHT(m.lcProgram, 3) == "FXP") OR (RIGHT(m.lcProgram, 3) == "PRG")
	CD ..
ENDIF

SET DEFAULT TO (SYS(5)+SYS(2003) + '\')
SET PROCEDURE TO Utils

PUBLIC tyear, tmonth, tdat1, tdat2, d_beg, d_end

PUBLIC tc_defdir, tc_basedir, tc_prgdir, tc_outdir, tc_tmpldir, tc_expimp, tc_aisdir,;
 tc_lpudir, tc_mcod, tc_lpuname, tc_lpuguid, tc_isfirst, tc_qcomp, tc_outlpudir, tc_comdir,;
 tc_post, tc_fam, tc_im, tc_ot, gcUser, mcod

tc_defdir = SUBSTR(SYS(5)+SYS(2003),1,RAT('\',SYS(5)+SYS(2003))-1)

IF !fso.FileExists(tc_defdir+'\Bin\Ag.cfg')
 m.tyear = 2012
 m.tmonth = 10
 m.tdat1 = CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0'))
 m.tdat2 = GOMONTH(CTOD('01.'+PADL(tMonth,2,'0')+'.'+PADL(tYear,4,'0')),1)-1

 tc_basedir = tc_defdir + '\Base'
 tc_prgdir  = tc_defdir + '\Bin'
 tc_comdir  = tc_defdir + '\Common'
 tc_outdir  = tc_defdir + '\Out'
 tc_tmpldir = tc_defdir + '\Templates'
 tc_expimp  = tc_defdir + '\ExpImp'
 tc_aisdir  = 'c:\aisoms\aisoms'

 CREATE TABLE &tc_defdir\Bin\Ag (tyear n(4), tmonth n(2), tdat1 d, tdat2 d,;
 basedir c(100), bindir c(100), comdir c(100), outdir c(100), tmpldir c(100), expdir c(100), aisdir c(100))
 USE 
 =OpenFile(tc_defdir+'\Bin\Ag', 'Ag', 'shar')
 INSERT INTO Ag (tyear, tmonth, tdat1, tdat2, basedir, bindir, comdir, outdir, tmpldir, expdir, aisdir) VALUES ;
  (m.tyear, m.tmonth, m.tdat1, m.tdat2, m.tc_basedir, m.tc_prgdir, m.tc_comdir, m.tc_outdir, ;
  m.tc_tmpldir, m.tc_expimp, m.tc_aisdir)  
 USE IN Ag
 fso.CopyFile(tc_defdir+'\Bin\Ag.dbf', tc_defdir+'\Bin\Ag.cfg')
 fso.DeleteFile(tc_defdir+'\Bin\Ag.dbf')
ENDIF 

=OpenFile(tc_defdir+'\Bin\Ag.cfg', 'Ag', 'shar')
SELECT Ag
m.tyear      = tyear
m.tmonth     = PADL(tmonth,2,'0')
m.tdat1      = tdat1
m.tdat2      = tdat2
m.tc_basedir = ALLTRIM(basedir)
m.tc_prgdir  = ALLTRIM(bindir)
m.tc_comdir  = ALLTRIM(comdir)
m.tc_outdir  = ALLTRIM(outdir)
m.tc_tmpldir = ALLTRIM(tmpldir)
m.tc_expimp  = ALLTRIM(expdir)
m.tc_aisdir  = ALLTRIM(aisdir)
USE 

m.gcUser = 'USR007'

m.mcod = '0101189'

oApp = NEWOBJECT("goapp", "CLASSES\MYCLASS.VCX")

SET PATH TO BITMAPS, CLASSES, FORMS, INCLUDE, PROGS, REPORTS
SET CLASSLIB TO MYCLASS

SET LIBRARY TO &tc_prgdir\vfpzip

=ComReind() 

LOCAL lhFile

*gcMainWindCaption = NAME_MAIN_WINDOW
gcOldWindCaption = _SCREEN.Caption
goEnvironment = CREATEOBJECT("Environment",)
goEnvironment.Set(.t.)

DIMENSION gaToolBars[11,2]
gaToolBars[1,1]  = TB_STANDARD_LOC
gaToolBars[2,1]  = TB_LAYOUT_LOC 
gaToolBars[3,1]  = TB_QUERY_LOC
gaToolBars[4,1]  = TB_VIEWDESIGNER_LOC
gaToolBars[5,1]  = TB_COLORPALETTE_LOC  
gaToolBars[6,1]  = TB_FORMCONTROLS_LOC
gaToolBars[7,1]  = TB_DATADESIGNER_LOC
gaToolBars[8,1]  = TB_REPODESIGNER_LOC
gaToolBars[9,1]  = TB_REPOCONTROLS_LOC
gaToolBars[10,1] = TB_PRINTPREVIEW_LOC
gaToolBars[11,1] = TB_FORMDESIGNER_LOC

FOR m.i = 1 TO ALEN(gaToolBars, 1)
 IF WEXIST(gaToolBars[m.i,1]) AND WVISIBLE(gaToolBars[m.i,1])
  gaToolBars[m.i,2] = .T.
  #IF DEBUGMODE
   HIDE WINDOW (gaToolBars[m.i,1])
  #ELSE
   RELEASE WINDOWS (gaToolBars[m.i,1])
  #ENDIF		
 ENDIF	
ENDFOR
RELEASE m.i

m.LocalPath = SYS(5)+SYS(2003) + '\'
* Íå çàïóùåíà ëè óæå ïðîãðàììà?
#IF !DEBUGMODE
	IF FILE(m.LocalPath + 'GRIDS.DBF')
		m.lhFile = FOPEN(m.LocalPath + 'GRIDS.DBF', 12)
		IF m.lhFile < 0
			*Ïðîãðàììà óæå îäèí ðàç çàïóùåíà
			=MESSAGEBOX("Ïîâòîðíûé çàïóñê ïðîãðàììû!", ;
				MB_ICONEXCLAMATION, ;
				NAME_MAIN_WINDOW)
			RETURN .F.
		ELSE
			=FCLOSE(m.lhFile)
		ENDIF
	ENDIF
#ENDIF


 SET SYSMENU TO
 SET SYSMENU ON
* SET MESSAGE TO gcMainWindCaption
 SET STATUS BAR ON 
 
_SCREEN.Picture = '&tc_prgdir\jpg.jpg'
 
 WITH _SCREEN
  .Width      = 640
  .Height     = 480-100
  .BackColor  = RGB(192,192,192)
  .Icon = 'cross.ico'
  .AutoCenter = .t.
  .Visible = .t.
 ENDWITH 


 DEFINE PAD mmenu_1 OF _MSYSMENU PROMPT '\<Ñ×ÅÒ' COLOR SCHEME 3 KEY ALT+A , ""
 DEFINE PAD mmenu_2 OF _MSYSMENU PROMPT '\<ÐÅÃÈÑÒÐ' COLOR SCHEME 3 KEY ALT+B , ""
 DEFINE PAD mmenu_3 OF _MSYSMENU PROMPT '\<ÏÐÈÊÐÅÏËÅÍÈÅ' COLOR SCHEME 3 KEY ALT+C , ""
 DEFINE PAD mmenu_4 OF _MSYSMENU PROMPT '\<ÑÅÐÂÈÑ' COLOR SCHEME 3 KEY ALT+D , ""

 ON PAD mmenu_1 OF _MSYSMENU ACTIVATE POPUP popAccExpImp
 ON PAD mmenu_2 OF _MSYSMENU ACTIVATE POPUP popRegExpImp
 ON PAD mmenu_3 OF _MSYSMENU ACTIVATE POPUP popPriklNas
 ON PAD mmenu_4 OF _MSYSMENU ACTIVATE POPUP popServ

 DEFINE POPUP popPriklNas MARGIN RELATIVE SHADOW COLOR SCHEME 4
 DEFINE BAR 01 OF popPriklNas PROMPT 'ÂÛÃÐÓÇÊÀ ÏÐÈÊÐÅÏËÅÍÈß' PICTURE 'DATABASE.ICO'
 DEFINE BAR 02 OF popPriklNas PROMPT '\-' PICTURE 'DATABASE.ICO'
 DEFINE BAR 03 OF popPriklNas PROMPT 'ÓÏÀÊÎÂÊÀ ÔÀÉËÀ T_ATTACHEDPATS_27258' PICTURE 'DATABASE.ICO'
 ON SELECTION BAR 01 OF popPriklNas DO MakePrikl
 ON SELECTION BAR 03 OF popPriklNas DO PackAttPats

 DEFINE POPUP popAccExpImp MARGIN RELATIVE SHADOW COLOR SCHEME 4
 DEFINE BAR 01 OF popAccExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÔÀÉËÀ TALON' PICTURE 'DATABASE.ICO'
 DEFINE BAR 02 OF popAccExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÔÀÉËÎÂ MEN È WOMEN' PICTURE 'DATABASE.ICO' SKIP FOR !fso.FileExists(tc_basedir+'\Talon.dbf')
 DEFINE BAR 03 OF popAccExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÍÎÂÎÃÎ Ñ×ÅÒÀ' PICTURE 'DATABASE.ICO' ;
  SKIP FOR !fso.FileExists(tc_basedir+'\Talon.dbf') OR !fso.FileExists(tc_basedir+'\Men.dbf') OR !fso.FileExists(tc_basedir+'\Women.dbf')
 DEFINE BAR 04 OF popAccExpImp PROMPT 'ÇÀÃÐÓÇÊÀ ÒÀËÎÍÎÂ' PICTURE 'DATABASE.ICO' ;
   SKIP FOR !fso.FileExists(tc_basedir+'\Talon.dbf') OR !fso.FileExists(tc_basedir+'\Men.dbf') OR !fso.FileExists(tc_basedir+'\Women.dbf')
 DEFINE BAR 05 OF popAccExpImp PROMPT 'ÓÄÀËÅÍÈÅ ÏÎÑËÅÄÍÅÉ ÇÀÃÐÓÇÊÈ' PICTURE 'DATABASE.ICO'
 DEFINE BAR 06 OF popAccExpImp PROMPT '\-'
 DEFINE BAR 07 OF popAccExpImp PROMPT 'ÂÛÕÎÄ' PICTURE 'CLOSE.BMP' KEY ALT+F4, "ALT+F4"
 ON SELECTION BAR 01 OF popAccExpImp DO MakeTalon
 ON SELECTION BAR 02 OF popAccExpImp DO MakeMenWomen
 ON SELECTION BAR 03 OF popAccExpImp DO MakeNewAcc
 ON SELECTION BAR 04 OF popAccExpImp DO LoadTalon
 ON SELECTION BAR 05 OF popAccExpImp DO DelLastLoad
 ON SELECTION BAR 07 OF popAccExpImp CLEAR EVENTS

 DEFINE POPUP popRegExpImp MARGIN RELATIVE SHADOW COLOR SCHEME 4
 DEFINE BAR 01 OF popRegExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÐÅÃÈÑÒÐÀ' PICTURE 'DATABASE.ICO'
 DEFINE BAR 02 OF popRegExpImp PROMPT 'ÔÎÐÌÈÐÎÂÀÍÈÅ ÇÀÏÐÎÑÀ Ê ÅÐÇ' PICTURE 'DATABASE.ICO'
 DEFINE BAR 03 OF popRegExpImp PROMPT 'ÏÎÈÑÊ ÎÒÂÅÒÀ ÎÒ ÅÐÇ' PICTURE 'DATABASE.ICO'
 DEFINE BAR 04 OF popRegExpImp PROMPT 'ÎÁÐÀÁÎÒÊÀ ÎÒÂÅÒÀ ÎÒ ÅÐÇ' PICTURE 'DATABASE.ICO'
 DEFINE BAR 05 OF popRegExpImp PROMPT '\-'
 DEFINE BAR 06 OF popRegExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÔÀÉËÀ DOCTOR'
 DEFINE BAR 07 OF popRegExpImp PROMPT 'ÏÎÑÒÐÎÅÍÈÅ ÔÀÉËÀ OTDEL'

 ON SELECTION BAR 01 OF popRegExpImp DO MakePeople
 ON SELECTION BAR 02 OF popRegExpImp DO MakeRequest
 ON SELECTION BAR 03 OF popRegExpImp DO FindAnswer
 ON SELECTION BAR 04 OF popRegExpImp DO ProcAnswer
 ON SELECTION BAR 06 OF popRegExpImp DO MakeDoctor
 ON SELECTION BAR 07 OF popRegExpImp DO MakeOtdel

 DEFINE POPUP popServ MARGIN RELATIVE SHADOW COLOR SCHEME 4
 DEFINE BAR 01 OF popServ PROMPT 'ÍÀÑÒÐÎÉÊÀ ÎÒ×ÅÒÍÎÃÎ ÏÅÐÈÎÄÀ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 02 OF popServ PROMPT 'Ñ×ÅÒ-ÔÀÊÒÓÐÀ Â ÁÄ ÏÏÎ ÎÌÑ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 03 OF popServ PROMPT 'Ñ×ÅÒ-ÔÀÊÒÓÐÀ Â ÂÛÃÐÓÇÊÅ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 04 OF popServ PROMPT 'ÄÎÑÒÓÏÍÀß ÑÓÌÌÀ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 05 OF popServ PROMPT '\-'
 DEFINE BAR 06 OF popServ PROMPT 'ÓÄÀËÈÒÜ ÑËÅÄÛ ÈÌÏÎÐÒÀ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 07 OF popServ PROMPT 'ÓÄÀËÈÒÜ ÊÀÐÒÛ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 08 OF popServ PROMPT 'ÎÁÍÓËÈÒÜ ÔÀÉË PATIENT_45426 (D_TYPE)' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 09 OF popServ PROMPT '\-'
 DEFINE BAR 10 OF popServ PROMPT 'ÎÁÍÎÂÈÒÜ ÂÅÊÒÎÐÀ (T_POLICY_53495)' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 11 OF popServ PROMPT '\-'
 DEFINE BAR 12 OF popServ PROMPT 'ÎÁÍÓËÈÒÜ ÁÄ PDPAccStorage' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 13 OF popServ PROMPT 'ÓÏÀÊÎÂÀÒÜ ÒÀËÎÍÛ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 14 OF popServ PROMPT 'ÓÏÀÊÎÂÀÒÜ ÏÀÖÈÅÍÒÎÂ' PICTURE 'HELPTXT.BMP'
 DEFINE BAR 15 OF popServ PROMPT '\-'
 DEFINE BAR 16 OF popServ PROMPT 'ÑÎÕÐÀÍÈÒÜ!'
 DEFINE BAR 17 OF popServ PROMPT 'ÏÅÐÅÑÒÐÎÈÒÜ UID Â T_TAPDIAGNOSIS_31571'
 DEFINE BAR 18 OF popServ PROMPT 'ÂÎÑÑÒÀÍÎÂÈÒÜ!'
 DEFINE BAR 19 OF popServ PROMPT '\-'
 DEFINE BAR 20 OF popServ PROMPT 'ÏÅÐÅÑÒÐÎÈÒÜ UID Â T_TAP_10874'

 ON SELECTION BAR 01 OF popServ DO FORM SetPeriod
 ON SELECTION BAR 02 OF popServ DO oms2sql
 ON SELECTION BAR 03 OF popServ DO oms2talon
 ON SELECTION BAR 04 OF popServ DO oms2good
 ON SELECTION BAR 06 OF popServ DO KillImportTrace
 ON SELECTION BAR 07 OF popServ DO KillCards
 ON SELECTION BAR 08 OF popServ DO UpdatePat45426
 ON SELECTION BAR 10 OF popServ DO UpdateVector
 ON SELECTION BAR 12 OF popServ DO PackAccStorage
 ON SELECTION BAR 13 OF popServ DO PackTalon
 ON SELECTION BAR 14 OF popServ DO PackPeople
 ON SELECTION BAR 16 OF popServ DO SaveUIDInTapDiags
 ON SELECTION BAR 17 OF popServ DO CorUIDInTapDiags
 ON SELECTION BAR 18 OF popServ DO RestoreUIDInTapDiags
 ON SELECTION BAR 20 OF popServ DO CorUIDInTap10874

 SET SYSMENU AUTOMATIC
 SET SYSMENU ON
 
 READ EVENT
 
 =Finish()
 
 FUNCTION Finish
 #IF DEBUGMODE
	*-- Âûâîä âñåõ VFP toolbars, êîòîðûå áûëè ñêðûòû
	FOR m.i = 1 TO ALEN(gaToolBars, 1)
		IF WEXIST(gaToolBars[m.i,1]) AND gaToolBars[m.i,2]
			SHOW WINDOW (gaToolBars[m.i,1])
		ENDIF
	ENDFOR
 #ENDIF


 #IF DEBUGMODE
  goEnvironment.ReSet
  SET SYSMENU TO DEFAULT
  WITH _SCREEN
   .TitleBar = 1
   .WindowState = 2
   .LockScreen = .F.
   .Picture = ''
   .BackColor = RGB(255,255,255)
  ENDWITH 
  SET SYSMENU ON
 #ELSE
 _SCREEN.Caption = ""
 #ENDIF

 #IF !DEBUGMODE
  ON SHUTDOWN
  QUIT
 #ELSE
  ON SHUTDOWN
  _SCREEN.Icon =""
*  _SCREEN.FirstStart = .T.
  *SET HELP TO
  CLEAR ALL
  CLOSE ALL
  CLEAR PROGRAM
  SET SYSMENU NOSAVE
  SET SYSMENU TO DEFAULT
  SET SYSMENU ON
 #ENDIF
RETURN 