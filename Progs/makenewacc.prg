PROCEDURE MakeNewAcc
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре онярпнхрэ мнбши явер?'+;
  CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_basedir+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик рюкнмнб!'+CHR(13)+CHR(10),0+16,'talon.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(tc_basedir+'\men.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик MEN!'+CHR(13)+CHR(10),0+16,'men.dbf')
  RETURN 
 ENDIF 
 IF !fso.FileExists(tc_basedir+'\women.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик WOMEN!'+CHR(13)+CHR(10),0+16,'women.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(tc_basedir+'\talon', 'talon', 'shar', 'patientuid')>0
  RETURN 
 ENDIF 
 IF OpenFile(tc_basedir+'\men', 'men', 'shar', 'dr')>0
  USE IN talon 
  RETURN 
 ENDIF 
 IF OpenFile(tc_basedir+'\women', 'women', 'shar', 'dr')>0
  USE IN talon 
  USE IN men 
  RETURN 
 ENDIF 
 
 oldsetnear = SET("Near")
 SET NEAR ON 
 
 WAIT "напюанрйю..." WINDOW NOWAIT 
 
 SELECT men 
 oord = ORDER()
 SET ORDER TO 
 COUNT FOR used=.t. TO m.UsedInMen
 SET ORDER TO (oord)
 IF m.UsedInMen>0
  WAIT CLEAR 
  MESSAGEBOX(CHR(13)+CHR(10)+'тюик MEN сфе хяонкэгнбюкяъ!'+CHR(13)+CHR(10)+;
  'янгдюире ецн гюмнбн!',0+16,TRANSFORM(m.UsedInMen,'999999'))
  USE IN talon 
  USE IN men 
  USE IN women
  RETURN 
 ENDIF 
* REPLACE ALL used WITH .f.

 SELECT women
 oord = ORDER()
 SET ORDER TO 
 COUNT FOR used=.t. TO m.UsedInWomen
 SET ORDER TO (oord)
 IF m.UsedInWomen>0
  WAIT CLEAR 
  MESSAGEBOX(CHR(13)+CHR(10)+'тюик WOMEN сфе хяонкэгнбюкяъ!'+CHR(13)+CHR(10)+'янгдюире ецн гюмнбн!',0+16,TRANSFORM(m.UsedInWoMen,'999999'))
  USE IN talon 
  USE IN men 
  USE IN women
  RETURN 
 ENDIF 
* REPLACE ALL used WITH .f.
 
 m.totrecs = 0
 m.replrecs = 0

 SELECT talon 
 m.pazid = patientuid
 m.w = w
 m.BornYear = YEAR(dr)
 IF m.w==1
  m.newpazid  = IIF(SEEK(m.BornYear, 'men'), men.id, 0)
  m.newdr     = IIF(SEEK(m.BornYear, 'men'), men.dr, {})
  m.neww      = IIF(SEEK(m.BornYear, 'men'), men.w, 0)
  m.isreplace = IIF(SEEK(m.BornYear, 'men'), .t., .f.)
  REPLACE men.used WITH .t. IN men
 ELSE
  m.newpazid  = IIF(SEEK(m.BornYear, 'women'), women.id, 0)
  m.newdr     = IIF(SEEK(m.BornYear, 'women'), women.dr, {})
  m.neww      = IIF(SEEK(m.BornYear, 'women'), women.w, 0)
  m.isreplace = IIF(SEEK(m.BornYear, 'women'), .t., .f.)
  REPLACE women.used WITH .t. IN women
 ENDIF 

 SCAN 
  m.totrecs = m.totrecs + 1
  IF patientuid == m.pazid
*   REPLACE npatuid WITH m.newpazid, newdr WITH m.newdr, neww WITH m.neww, isreplace WITH m.isreplace
  ELSE 
   IF w==1
    m.newpazid  = IIF(SEEK(YEAR(dr), 'men'), men.id, 0)
    m.newdr     = IIF(SEEK(YEAR(dr), 'men'), men.dr, {})
    m.neww      = IIF(SEEK(YEAR(dr), 'men'), men.w, 0)
    m.isreplace = IIF(SEEK(YEAR(dr), 'men'), .t., .f.)
    REPLACE men.used WITH .T. IN men
   ELSE
    m.newpazid  = IIF(SEEK(YEAR(dr), 'women'), women.id, 0)
    m.newdr     = IIF(SEEK(YEAR(dr), 'women'), women.dr, {})
    m.neww      = IIF(SEEK(YEAR(dr), 'women'), women.w, 0)
    m.isreplace = IIF(SEEK(YEAR(dr), 'women'), .t., .f.)
    REPLACE women.used WITH .T. IN women
   ENDIF 
*   REPLACE npatuid WITH m.newpazid, newdr WITH m.newdr, neww WITH m.neww, isreplace WITH m.isreplace
  ENDIF 

  REPLACE ntapuid WITH 0, npatuid WITH m.newpazid, newdr WITH m.newdr, neww WITH m.neww, isreplace WITH m.isreplace

  m.pazid = patientuid
  m.w = w
  m.BornYear = YEAR(dr)
  REPLACE ok WITH IIF(!IsFound OR !IsReplace OR w!=neww OR YEAR(dr)!=YEAR(newdr), .f., ok)
 ENDSCAN 
 
* =OpenFile(tc_basedir+'\talon', 'talon2', 'shar', 'tapuid', 'agai')

* SELECT talon
* SET ORDER TO tapuid
* m.tapuid = tapuid
* m.newtapuid = tapuid+1
* DO WHILE SEEK(m.newtapuid, 'talon2')
*  m.newtapuid = m.newtapuid + 1
* ENDDO 

* SCAN 
*  IF tapuid == m.tapuid
*   REPLACE ntapuid WITH m.newtapuid
*  ELSE 
*   m.newtapuid = MAX(tapuid,m.newtapuid)+1
*   DO WHILE SEEK(m.newtapuid, 'talon2')
*    m.newtapuid = m.newtapuid + 1
*   ENDDO 
*   REPLACE ntapuid WITH m.newtapuid
*  ENDIF 
*  m.tapuid = tapuid
* ENDSCAN 

 WAIT "гюосяй лндскъ яюлндхюцмнярхйх..." WINDOW NOWAIT 
  SELECT npatuid DISTINCT FROM Talon WHERE npatuid IN ;
   (SELECT patientuid DISTINCT FROM talon) INTO CURSOR across
  SELECT across
  nAcross = RECCOUNT('across')
  USE IN across 
 WAIT CLEAR 
 
 IF nAcross==0
  MESSAGEBOX(CHR(13)+CHR(10)+'ньханй тнплхпнбюмхъ яверю ме намюпсфемн!'+;
   CHR(13)+CHR(10),0+64,'')
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'намюпсфемш опнакелш: '+TRANSFORM(nAcross,'99999')+' онбрнпнб!'+;
   CHR(13)+CHR(10), 0+16, '')
 ENDIF 
 
 USE IN talon
* USE IN talon2
 USE IN men 
 USE IN women 

* IF OpenFile(tc_basedir+'\talon', 'talon', 'excl')<=0
*  SELECT talon 
*  WAIT "ондявер окнуху рюкнмнб..." WINDOW  NOWAIT 
*  COUNT FOR !ok TO m.nHowMuchNotOK
*  WAIT CLEAR 
*  MESSAGEBOX('асдер сдюкемн '+ALLTRIM(STR(m.nHowMuchNotOK))+' гюохяеи!',0+64,'')
*  WAIT "сдюкемхе окнуху рюкнмнб..." WINDOW  NOWAIT 
*  DELETE FOR !ok
*  PACK 
*  USE 
*  WAIT CLEAR 
* ENDIF 
 
 SET NEAR &oldsetnear
 
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'напюанрйю гюйнмвемю!'+CHR(13)+CHR(10), 0+64, '')

RETURN 