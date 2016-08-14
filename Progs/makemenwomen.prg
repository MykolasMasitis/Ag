PROCEDURE MakeMenWomen
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œŒ—“–Œ»“‹ ‘¿…À€ MEN » WOMEN?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_comdir+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À –≈√»—“–¿!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_basedir+'\talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À “¿ÀŒÕŒ¬!'+CHR(13)+CHR(10),0+16,'talon.dbf')
  RETURN 
 ENDIF 

 IF OpenFile(tc_comdir+'\people', 'people', 'shar')>0
  RETURN 
 ENDIF 

 IF OpenFile(tc_basedir+'\talon', 'talon', 'shar')>0
  USE IN people
  RETURN 
 ENDIF 

 IF fso.FileExists(tc_basedir+'\men.dbf')
  fso.DeleteFile(tc_basedir+'\men.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\men.cdx')
  fso.DeleteFile(tc_basedir+'\men.cdx')
 ENDIF 

 IF fso.FileExists(tc_basedir+'\women.dbf')
  fso.DeleteFile(tc_basedir+'\women.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\women.cdx')
  fso.DeleteFile(tc_basedir+'\women.cdx')
 ENDIF 
 
 WAIT "Œ¡–¿¡Œ“ ¿..." WINDOW NOWAIT 
 
 m.totrecs = RECCOUNT('people')
 m.men = 0
 m.women = 0 
* SELECT * FROM people WHERE w==1 AND ok AND id NOT in (SELECT patientuid DISTINCT FROM talon) ;
  INTO dbf tc_basedir+'\men'
 SELECT * FROM people WHERE w=1 AND ok AND !used INTO dbf tc_basedir+'\men'
 m.men = _tally
 INDEX ON YEAR(dr) FOR !used TAG dr
 USE 

* SELECT * FROM people WHERE w==2 AND ok AND id NOT in (SELECT patientuid DISTINCT FROM talon) ;
  INTO dbf tc_basedir+'\women'
 SELECT * FROM people WHERE w=2 AND ok AND !used INTO dbf tc_basedir+'\women'
 m.women = _tally
 INDEX ON YEAR(dr) FOR !used TAG dr
 USE 
 
 USE IN people 
 USE IN talon 
 
 WAIT CLEAR 

 MESSAGEBOX(CHR(13)+CHR(10)+'ƒÀﬂ ƒ¿À‹Õ≈…ÿ≈… Œ¡–¿¡Œ“ »'+CHR(13)+CHR(10)+;
   'Œ“Œ¡–¿ÕŒ '+STR(m.men+m.women,5)+' «¿œ»—≈… –≈√»—“–¿ »« '+STR(m.TotRecs,5)+CHR(13)+CHR(10), 0+64, 'freeppl')

RETURN 