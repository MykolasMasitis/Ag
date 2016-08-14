PROCEDURE Start
 SET CENTURY ON 
 SET DATE GERMAN 
 SET SAFETY OFF 

 DefDir = 'd:\planum189'
 SET DEFAULT TO &DefDir
 PUBLIC fso as Scripting.FileSystemObject 
 fso = CREATEOBJECT('Scripting.FileSystemObject')
 
 && ƒÂÎ‡ÂÏ Ò‚Ó‰Ì˚È people
 FileName = DefDir+'\People\Patient_Patient.csv'
 IF !fso.FileExists(FileName)
  MESSAGEBOX('ÕÂ Û‰‡ÎÓÒ¸ Ì‡ÈÚË Ù‡ÈÎ'+CHR(10)+CHR(13)+;
  FileNmae, 0+16, '')
  RETURN 
 ENDIF 
 
 fso.CopyFile(DefDir+'\Templates\PatientPatient.dbf', DefDir+'\People\PatientPatient.dbf')
 
 USE DefDir+'\People\PatientPatient' IN 0 ALIAS People EXCLUSIVE 
 SELECT People
 ALTER TABLE People ADD COLUMN IsUsed l
 INDEX ON PatientUID TAG PatientUID
 SET ORDER TO PatientUID

 fhandl = fso.OpenTextFile(FileName)
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
 MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 DO WHILE !fhandl.AtEndOfStream
  WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PatientPatient..." WINDOW NOWAIT 
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.patientuid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.fam        = PROPER(SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1))
  m.im         = PROPER(SUBSTR(CurString, AT(',',CurString,3)+1, AT(',',CurString,4)-AT(',',CurString,3)-1))
  m.ot         = PROPER(SUBSTR(CurString, AT(',',CurString,4)+1, AT(',',CurString,5)-AT(',',CurString,4)-1))
  m.dr         = SUBSTR(CurString, AT(',',CurString,5)+1, AT(',',CurString,6)-AT(',',CurString,5)-1)
  m.dr         = CTOD(SUBSTR(m.dr,9,2)+'.'+SUBSTR(m.dr,6,2)+'.'+SUBSTR(m.dr,1,4))
  m.w          = INT(VAL(SUBSTR(CurString, AT(',',CurString,7)+1, AT(',',CurString,8)-AT(',',CurString,7)-1)))
  
  INSERT INTO People FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')
 
 && ƒÂÎ‡ÂÏ Ò‚Ó‰Ì˚È people
 
 FileName = DefDir+'\Base\Patient_Patient.csv'
 IF !fso.FileExists(FileName)
  MESSAGEBOX('ÕÂ Û‰‡ÎÓÒ¸ Ì‡ÈÚË Ù‡ÈÎ'+CHR(10)+CHR(13)+;
  FileNmae, 0+16, '')
  RETURN 
 ENDIF 
 
 fso.CopyFile(DefDir+'\Templates\PatientPatient.dbf', DefDir+'\Base\PatientPatient.dbf')
 
 USE DefDir+'\Base\PatientPatient' IN 0 ALIAS Patient EXCLUSIVE 
 SELECT Patient
 INDEX ON PatientUID TAG PatientUID
 SET ORDER TO PatientUID

 fhandl = fso.OpenTextFile(FileName)
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
 MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 DO WHILE !fhandl.AtEndOfStream
  WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PatientPatient..." WINDOW NOWAIT 
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.patientuid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.fam        = PROPER(SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1))
  m.im         = PROPER(SUBSTR(CurString, AT(',',CurString,3)+1, AT(',',CurString,4)-AT(',',CurString,3)-1))
  m.ot         = PROPER(SUBSTR(CurString, AT(',',CurString,4)+1, AT(',',CurString,5)-AT(',',CurString,4)-1))
  m.dr         = SUBSTR(CurString, AT(',',CurString,5)+1, AT(',',CurString,6)-AT(',',CurString,5)-1)
  m.dr         = CTOD(SUBSTR(m.dr,9,2)+'.'+SUBSTR(m.dr,6,2)+'.'+SUBSTR(m.dr,1,4))
  m.w          = INT(VAL(SUBSTR(CurString, AT(',',CurString,7)+1, AT(',',CurString,8)-AT(',',CurString,7)-1)))
  
  INSERT INTO Patient FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 USE IN Patient
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')

 FileName = DefDir+'\Base\Stat_TAP.csv'
 IF !fso.FileExists(FileName)
  MESSAGEBOX('ÕÂ Û‰‡ÎÓÒ¸ Ì‡ÈÚË Ù‡ÈÎ'+CHR(10)+CHR(13)+;
  FileNmae, 0+16, '')
  RETURN 
 ENDIF 
 
 fso.CopyFile(DefDir+'\Templates\Tap.dbf', DefDir+'\Base\Tap.dbf')
 
 USE DefDir+'\Base\Tap' IN 0 ALIAS Tap EXCLUSIVE 
 SELECT Tap 
 INDEX ON TapUID TAG TapUID
 SET ORDER TO TapUID

 fhandl = fso.OpenTextFile(FileName)
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
 MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 DO WHILE !fhandl.AtEndOfStream
  WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAP..." WINDOW NOWAIT 
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.tapuid     = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.lpuuid = INT(VAL(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1)))
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
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')

 FileName = DefDir+'\Base\Stat_TAPdiagnosis.csv'
 IF !fso.FileExists(FileName)
  MESSAGEBOX('ÕÂ Û‰‡ÎÓÒ¸ Ì‡ÈÚË Ù‡ÈÎ'+CHR(10)+CHR(13)+;
  FileNmae, 0+16, '')
  RETURN 
 ENDIF 
 
 fso.CopyFile(DefDir+'\Templates\TapDiagnosis.dbf', DefDir+'\Base\TapDiagnosis.dbf')
 
 USE DefDir+'\Base\TapDiagnosis' IN 0 ALIAS TapDiagnosis EXCLUSIVE 
 SELECT TapDiagnosis
 INDEX ON diaguid TAG diaguid
 SET ORDER TO diaguid

 fhandl = fso.OpenTextFile(FileName)
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
 MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 DO WHILE !fhandl.AtEndOfStream
  WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAPDIAGNOSIS..." WINDOW NOWAIT 
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
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')

 FileName = DefDir+'\Base\Stat_TAPservice.csv'
 IF !fso.FileExists(FileName)
  MESSAGEBOX('ÕÂ Û‰‡ÎÓÒ¸ Ì‡ÈÚË Ù‡ÈÎ'+CHR(10)+CHR(13)+;
  FileNmae, 0+16, '')
  RETURN 
 ENDIF 
 
 fso.CopyFile(DefDir+'\Templates\TapService.dbf', DefDir+'\Base\TapService.dbf')
 
 USE DefDir+'\Base\TapService' IN 0 ALIAS TapService SHARED 

 fhandl = fso.OpenTextFile(FileName)
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
 MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 DO WHILE !fhandl.AtEndOfStream
  WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ TAPSERVICE..." WINDOW NOWAIT 
  CurString = fhandl.ReadLine
  HowManyCommas = OCCURS(',', CurString)
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.diaguid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.cod    = ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))
  m.k_u    = VAL(SUBSTR(CurString, AT(',',CurString,2)+1))
  
  INSERT INTO tapservice FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')
 
 WAIT "—¡Œ– ¿ ‘¿…À¿ TALON..." WINDOW NOWAIT 
 SELECT TapService
 SET RELATION TO DiagUID INTO TapDiagnosis
 SELECT TapDiagnosis
 SET RELATION TO tapuid INTO tap 
 
 CREATE TABLE &DefDir\Base\Talon ;
  (tapuid i, ntapuid i, patientuid i, npatuid i, dr d, newdr d, w n(1), neww n(1), isfound l, isreplace l, ;
   diaguid i, d_u d, cod c(8), k_u n(3), doctoruid i, nurseuid i, s_all n(11,2))
 USE 
 USE &DefDir\Base\Talon IN 0 ALIAS talon EXCLUSIVE 
 SELECT talon 
 INDEX ON tapuid TAG tapuid 
 SET ORDER TO tapuid 
 
 USE &DefDir\Common\TarifN IN 0 ALIAS tarif SHARED ORDER cod 
 
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
  
  m.usl = INT(VAL(SUBSTR(m.cod,3)))
  
  m.isfound = IIF(SEEK(m.patientuid, 'people'), .t., .f.)
  m.dr = IIF(SEEK(m.patientuid, 'people'), people.dr, {})
  m.w = IIF(SEEK(m.patientuid, 'people'), people.w, 0)
  
  m.s_all = IIF(SEEK(m.usl, 'tarif'), tarif.tarif*m.k_u, 0)

  m.SumOfAcc = m.SumOfAcc + m.s_all
  
  INSERT INTO Talon FROM MEMVAR 
  
 ENDSCAN 
 WAIT CLEAR 
 
 SELECT talon 
 COPY TO &DefDir\Base\tTalon
 ZAP
 APPEND FROM &DefDir\Base\tTalon
 SET ORDER TO 
 DELETE TAG ALL 
 
 SELECT * FROM people WHERE w=1 AND patientuid NOT IN (SELECT patientuid FROM talon) ;
  INTO TABLE &DefDir\Base\Men 
 USE 

 SELECT * FROM people WHERE w=2 AND patientuid NOT IN (SELECT patientuid FROM talon) ;
  INTO TABLE &DefDir\Base\WoMen 
 USE 
 
 USE IN People
 USE IN Talon 

 fso.DeleteFile(DefDir+'\Base\tTalon.dbf')
 
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
 
 MESSAGEBOX('—”ÃÃ¿ —◊≈“¿: '+TRANSFORM(m.SumOfAcc, '9 999 999.99'),0+64,'')
 
RETURN 