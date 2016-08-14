PROCEDURE MakeTalonExpImp
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œŒ—“–Œ»“‹ ‘¿…À TALON'+;
  CHR(13)+CHR(10)+'»« ‘¿…À¿ »ÃœŒ–“¿?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(tc_expimp)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ »ÃœŒ–“¿!'+CHR(13)+CHR(10),0+16,tc_expimp)
  RETURN 
 ENDIF 
 
 olddir = SYS(5)+SYS(2003)
 SET DEFAULT TO (tc_expimp)
 IF !fso.FileExists(tc_expimp+'\expimp.dat')
  ImpFile=GETFILE('dat')
 ELSE 
  ImpFile = tc_expimp+'\expimp.dat'
 ENDIF 
 SET DEFAULT TO (olddir)
 
 IF EMPTY(ImpFile)
  olddir = SYS(5)+SYS(2003)
  MESSAGEBOX(CHR(13)+CHR(10)+'¬€ Õ»◊≈√Œ Õ≈ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
  
 ffile = fso.GetFile(ImpFile)
 IF ffile.size >= 2
  fhandl = ffile.OpenAsTextStream
  lcHead = fhandl.Read(2)
  fhandl.Close
 ELSE 
  lcHead = ''
 ENDIF 

 llIsOneZip = .f.
 IF lcHead == 'PK' && ›ÚÓ zip-Ù‡ÈÎ!
  UnzipOpen(ImpFile)
  DataPack ='datapack.xml'
  IF UnzipGotoFileByName(DataPack)
   llIsOneZip = .t.
   UnzipClose()
  ENDIF 
  UnzipClose()
 ELSE 
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ “Œ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 

 IF llIsOneZip = .F.
  MESSAGEBOX(CHR(13)+CHR(10)+'Õ≈ “Œ ¬€¡–¿À»!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 

 WAIT "–¿—œ¿ Œ¬ ¿ STAT_Tap.csv..." WINDOW NOWAIT 
 UnzipOpen(ImpFile)

 IF !UnzipGotoFileByName('STAT_Tap.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À STAT_Tap.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_expimp+'\STAT_Tap.csv') 
  fso.DeleteFile(tc_expimp+'\STAT_Tap.csv')
 ENDIF 
 ppp=UnzipFile(tc_expimp)

 WAIT "–¿—œ¿ Œ¬ ¿ Stat_TAPdiagnosis.csv..." WINDOW NOWAIT 
 IF !UnzipGotoFileByName('Stat_TAPdiagnosis.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À Stat_TAPdiagnosis.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_expimp+'\Stat_TAPdiagnosis.csv') 
  fso.DeleteFile(tc_expimp+'\Stat_TAPdiagnosis.csv')
 ENDIF 
 ppp=UnzipFile(tc_expimp)

 WAIT "–¿—œ¿ Œ¬ ¿ Stat_TAPservice.csv..." WINDOW NOWAIT 
 IF !UnzipGotoFileByName('Stat_TAPservice.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À Stat_TAPservice.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_expimp+'\Stat_TAPservice.csv') 
  fso.DeleteFile(tc_expimp+'\Stat_TAPservice.csv')
 ENDIF 
 ppp=UnzipFile(tc_expimp)

 UnzipClose()
 WAIT CLEAR 
 
 IF OpenFile(tc_comdir+'\people', 'people', 'shar', 'id')>0
  RETURN 
 ENDIF 
 
 fso.CopyFile(tc_tmpldir+'\Tap.dbf', tc_basedir+'\Tap.dbf')
 =OpenFile(tc_basedir+'\Tap', 'Tap', 'excl') 
 SELECT Tap 
 INDEX ON TapUID TAG TapUID
 SET ORDER TO TapUID

 fhandl = fso.OpenTextFile(tc_expimp+'\STAT_Tap.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
* MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAP..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.tapuid     = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.lpuuid     = INT(VAL(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1)))
  m.patientuid = INT(VAL(SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1)))
  m.addressuid = INT(VAL(SUBSTR(CurString, AT(',',CurString,4)+1, AT(',',CurString,5)-AT(',',CurString,4)-1)))
  m.policyuid  = INT(VAL(SUBSTR(CurString, AT(',',CurString,5)+1, AT(',',CurString,6)-AT(',',CurString,5)-1)))
  m.bookuid    = INT(VAL(SUBSTR(CurString, AT(',',CurString,6)+1, AT(',',CurString,7)-AT(',',CurString,6)-1)))
  m.d_u        = SUBSTR(CurString, AT(',',CurString,7)+1, AT(',',CurString,8)-AT(',',CurString,7)-1)
  m.d_u        = CTOD(SUBSTR(m.d_u,9,2)+'.'+SUBSTR(m.d_u,6,2)+'.'+SUBSTR(m.d_u,1,4))
  m.doctoruid  = INT(VAL(SUBSTR(CurString, AT(',',CurString,8)+1, AT(',',CurString,9)-AT(',',CurString,8)-1)))
  m.nurseuid   = INT(VAL(SUBSTR(CurString, AT(',',CurString,9)+1, AT(',',CurString,10)-AT(',',CurString,9)-1)))
  
  INSERT INTO tap FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')

 fso.CopyFile(tc_tmpldir+'\TapDiagnosis.dbf', tc_basedir+'\TapDiagnosis.dbf')
 
 =OpenFile(tc_basedir+'\TapDiagnosis', 'TapDiagnosis', 'EXCL')
 SELECT TapDiagnosis
 INDEX ON diaguid TAG diaguid
 SET ORDER TO diaguid

 fhandl = fso.OpenTextFile(tc_expimp+'\Stat_TAPdiagnosis.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
* MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAPDIAGNOSIS..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.diaguid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.tapuid  = INT(VAL((SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))))
  m.ds      = ALLTRIM((SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1)))
  
  INSERT INTO tapDiagnosis FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')

 fso.CopyFile(tc_tmpldir+'\TapService.dbf', tc_basedir+'\TapService.dbf')
 
 OpenFile(tc_basedir+'\TapService',  'TapService', 'SHARED')

 fhandl = fso.OpenTextFile(tc_expimp+'\Stat_TAPservice.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
* MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAPSERVICE..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.diaguid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.coduid  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))
  m.k_u    = VAL(SUBSTR(CurString, AT(',',CurString,2)+1))
  
  INSERT INTO tapservice FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')
 
 WAIT "—¡Œ– ¿ ‘¿…À¿ TALON..." WINDOW NOWAIT 
 SELECT TapService
 SET RELATION TO DiagUID INTO TapDiagnosis
 SELECT TapDiagnosis
 SET RELATION TO tapuid INTO tap 
 
 CREATE TABLE &tc_basedir\Talon ;
  (ok l, tapuid i, ntapuid i, patientuid i, npatuid i, dr d, newdr d, w n(1), neww n(1), isfound l, isreplace l, ;
   diaguid i, ds c(6), d_u d, coduid c(8), cod n(6), k_u n(3), doctoruid i, nurseuid i, s_all n(11,2))
 USE 
 USE &tc_basedir\Talon IN 0 ALIAS talon EXCLUSIVE 
 SELECT talon 
 INDEX ON tapuid TAG tapuid 
 INDEX ON patientuid TAG patientuid
 SET ORDER TO tapuid 
 
 USE &tc_comdir\TarifN IN 0 ALIAS tarif SHARED ORDER cod 
 
 m.SumOfAcc = 0
 
 SELECT TapService
 SCAN 
  SCATTER MEMVAR && m.diaguid, m.cod, m.k_u
  
  m.ds = TapDiagnosis.ds
  m.tapuid = tap.tapuid
  m.patientuid = tap.patientuid
  m.addressuid = tap.addressuid
  m.policyuid = tap.policyuid
  m.bookuid = tap.bookuid
  m.d_u = tap.d_u
  m.doctoruid = tap.doctoruid
  m.nurseuid = tap.nurseuid
  
  m.cod = INT(VAL(SUBSTR(m.coduid,3)))
  
  m.isfound = IIF(SEEK(m.patientuid, 'people'), .t., .f.)
  m.dr = IIF(SEEK(m.patientuid, 'people'), people.dr, {})
  m.w = IIF(SEEK(m.patientuid, 'people'), people.w, 0)
  
  m.s_all = IIF(SEEK(m.cod, 'tarif'), tarif.tarif*m.k_u, 0)

  m.SumOfAcc = m.SumOfAcc + m.s_all
  
  IF !m.IsFound OR INLIST(cod,1561,1601) OR ;
     BETWEEN(ds,'Z32.0 ','Z39.2 ') OR ;
     BETWEEN(ds,'O20.0 ','O99.8 ')
   m.ok = .F.
  ELSE
   m.ok = .T.
  ENDIF 
  
  INSERT INTO Talon FROM MEMVAR 
  
 ENDSCAN 
 WAIT CLEAR 
 
 SELECT talon 
 COPY TO &tc_basedir\tTalon
 ZAP
 APPEND FROM &tc_basedir\tTalon
 SET ORDER TO patientuid
 
 SELECT people
 
 SCAN 
  m.id = id
  IF SEEK(m.id, 'talon')
   REPLACE used WITH .t.
  ELSE 
   REPLACE used WITH .f.
  ENDIF 
 ENDSCAN 
  
 USE IN People
 USE IN Talon 

 fso.DeleteFile(tc_basedir+'\tTalon.dbf')
 
 SELECT TapDiagnosis
 SET RELATION OFF INTO tap 

 SELECT Tap
 SET ORDER TO 
 DELETE TAG ALL 
 USE 

 SELECT TapService
 SET RELATION OFF INTO TapDiagnosis
 USE 

 SELECT TapDiagnosis
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 
 USE IN tarif 

 IF fso.FileExists(tc_expimp+'\STAT_Tap.csv') 
  fso.DeleteFile(tc_expimp+'\STAT_Tap.csv')
 ENDIF 
 IF fso.FileExists(tc_expimp+'\Stat_TAPdiagnosis.csv') 
  fso.DeleteFile(tc_expimp+'\Stat_TAPdiagnosis.csv')
 ENDIF 
 IF fso.FileExists(tc_expimp+'\Stat_TAPservice.csv') 
  fso.DeleteFile(tc_expimp+'\Stat_TAPservice.csv')
 ENDIF 
 
 IF fso.FileExists(tc_basedir+'\TAP.dbf') 
  fso.DeleteFile(tc_basedir+'\TAP.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\TAPDiagnosis.dbf') 
  fso.DeleteFile(tc_basedir+'\TAPDiagnosis.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\TAPService.dbf') 
  fso.DeleteFile(tc_basedir+'\TAPService.dbf')
 ENDIF 

 MESSAGEBOX('—”ÃÃ¿ —◊≈“¿: '+TRANSFORM(m.SumOfAcc, '9 999 999.99'),0+64,'')
 
RETURN 
