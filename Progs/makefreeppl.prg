PROCEDURE MakeFreePpl
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре онярпнхрэ тюик днярсомнцн пецхярпю?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_comdir+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик ябндмнцн пецхярпю!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(tc_comdir+'\people', 'people', 'shar')>0
  RETURN 
 ENDIF 
 
 WAIT "напюанрйю..." WINDOW NOWAIT 
 
 m.TotRecs = 0
 m.GoodRecs = 0 
 SELECT people
 COPY STRUCTURE TO tc_comdir+'\freeppl' 
 =OpenFile(tc_comdir+'\freeppl', 'freeppl', 'shar')
 SELECT people 

 SCAN 
  m.TotRecs = m.TotRecs + 1
  IF ok AND sn_pol==sn_polerz
   m.GoodRecs = m.GoodRecs + 1
   SCATTER MEMVAR 
   INSERT INTO freeppl FROM MEMVAR 
  ENDIF 
 ENDSCAN 

 USE 
 USE IN freeppl
 
 WAIT CLEAR 
 
 MESSAGEBOX(CHR(13)+CHR(10)+'дкъ дюкэмеиьеи напюанрйх'+CHR(13)+CHR(10)+;
   'днярсомн '+STR(m.GoodRecs,5)+' гюохяеи пецхярпю хг '+STR(m.TotRecs,5)+CHR(13)+CHR(10), 0+64, 'freeppl')

RETURN 