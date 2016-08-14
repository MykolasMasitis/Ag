PROCEDURE MakePeopleExpImp

 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œŒ—“–Œ»“‹ ‘¿…À –≈√»—“–¿'+;
  CHR(13)+CHR(10)+'»« ‘¿…À¿ »ÃœŒ–“¿?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FolderExists(tc_comdir)
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ƒ»–≈ “Œ–»ﬂ COMMON!'+CHR(13)+CHR(10),0+16,tc_comdir)
  RETURN 
 ENDIF 
 
 olddir = SYS(5)+SYS(2003)
 SET DEFAULT TO (tc_comdir)
 IF !fso.FileExists(tc_comdir+'\people.dat')
  ImpFile=GETFILE('dat')
 ELSE 
  ImpFile = tc_comdir+'\people.dat'
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
 RELEASE m.ffile 

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
 
 ffile = fso.GetFile(ImpFile)
 CreateDate = ffile.DateLastModified
 RELEASE m.ffile 
 IF DATETIME()-CreateDate>12*60*60
  IF MESSAGEBOX(CHR(13)+CHR(10)+; 
  '‘¿…À »ÃŒ–“¿ '+ALLTRIM(m.ImpFile)+CHR(13)+CHR(10)+; 
  '¡€À —Œ«ƒ¿Õ '+TTOC(CreateDate) +'.'+CHR(13)+CHR(10)+; 
  '—≈…◊¿— '+TTOC(DATE())+'.'+CHR(13)+CHR(10)+; 
  CHR(13)+CHR(10)+'Œ¡–¿¡¿“€¬¿“‹ ¬—≈ –¿¬ÕŒ?',4+32, '¬Õ»Ã¿Õ»≈') == 7
  RETURN 
 ENDIF 
  
 ENDIF 

 WAIT "–¿—œ¿ Œ¬ ¿ Patient_Patient.csv..." WINDOW NOWAIT 
 UnzipOpen(ImpFile)
 IF !UnzipGotoFileByName('Patient_Patient.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À Patient_Patient.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_comdir+'\Patient_Patient.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Patient.csv')
 ENDIF 
 ppp=UnzipFile(tc_comdir)

 WAIT "–¿—œ¿ Œ¬ ¿ Patient_Policy.csv..." WINDOW NOWAIT 
 IF !UnzipGotoFileByName('Patient_Policy.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À Patient_Policy.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_comdir+'\Patient_Policy.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Policy.csv')
 ENDIF 
 ppp=UnzipFile(tc_comdir)

 WAIT "–¿—œ¿ Œ¬ ¿ Patient_Address.csv..." WINDOW NOWAIT 
 IF !UnzipGotoFileByName('Patient_Address.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À Patient_Address.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_comdir+'\Patient_Address.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Address.csv')
 ENDIF 
 ppp=UnzipFile(tc_comdir)

 WAIT "–¿—œ¿ Œ¬ ¿ EHR_Book.csv..." WINDOW NOWAIT 
 IF !UnzipGotoFileByName('EHR_Book.csv')
  MESSAGEBOX(CHR(13)+CHR(10)+'¿–’»¬ Õ≈ —Œƒ≈–∆»“ ‘¿…À EHR_Book.csv!'+CHR(13)+CHR(10),0+16,ImpFile)
  RETURN 
 ENDIF 
 IF fso.FileExists(tc_comdir+'\EHR_Book.csv') 
  fso.DeleteFile(tc_comdir+'\EHR_Book.csv')
 ENDIF 
 ppp=UnzipFile(tc_comdir)

 UnzipClose()
 WAIT CLEAR 
 
 IF fso.FileExists(tc_comdir+'\people.dbf')
  fso.DeleteFile(tc_comdir+'\people.dbf')
 ENDIF 
 
 CREATE TABLE &tc_comdir\people ;
  (ok l, used l, id i, c_t n(2), c_i c(25), bookid i, policyuid i, s_pol c(25), n_pol c(25),;
   tip c(1), sn_pol c(25), sn_polerz c(25), ;
   qerz c(2), v_erz2 c(3), d_erz2 d, fam c(25), im c(20), ot c(20), w n(1), dr d, ;
   ul n(5), kladr c(4), dom c(7), kor c(5), str c(5), kv c(5), adrid i, regdate c(19), upddate c(19)) 
 USE 
 =OpenFile(tc_comdir+'\people', 'people', 'excl')
 SELECT people
 INDEX ON id TAG id 
 SET ORDER TO id

 =OpenFile(tc_comdir+'\territxx', 'territ', 'shared', 'c_t')
 =OpenFile(tc_comdir+'\street', 'street', 'shared', 'ul ')
* SELECT street 

* fhandl = fso.OpenTextFile(tc_comdir+'\MskOMS_Street.csv')
* CurString = fhandl.ReadLine
* HowManyCommas = OCCURS(',', CurString)+1

* m.HowManyRecs = 0
* WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ STREET (MskOMS_Street.csv)..." WINDOW NOWAIT 
* DO WHILE !fhandl.AtEndOfStream
*  m.HowManyRecs = m.HowManyRecs + 1
*  CurString = fhandl.ReadLine
  
*  m.ul    = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
*  m.kladr = ALLTRIM(SUBSTR(CurString, AT(',',CurString,8)+1, AT(',',CurString,9)-AT(',',CurString,8)-1))
    
*  IF SEEK(m.ul, 'street')
*   REPLACE kladr WITH m.kladr
*  ENDIF 
  
* ENDDO 
* WAIT CLEAR 
* fhandl.Close
  
 fhandl = fso.OpenTextFile(tc_comdir+'\Patient_Patient.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
* MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
 m.HowManyRecs = 0
 
 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PEOPLE (Patient_Patient.csv)..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  
  m.HowManyRecs = m.HowManyRecs + 1

  m.id  = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.fam = SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1)
  m.im  = SUBSTR(CurString, AT(',',CurString,3)+1, AT(',',CurString,4)-AT(',',CurString,3)-1)
  m.ot  = SUBSTR(CurString, AT(',',CurString,4)+1, AT(',',CurString,5)-AT(',',CurString,4)-1)
  m.dr  = SUBSTR(CurString, AT(',',CurString,5)+1, AT(',',CurString,6)-AT(',',CurString,5)-1)
  m.dr  = CTOD(SUBSTR(m.dr,9,2)+'.'+SUBSTR(m.dr,6,2)+'.'+SUBSTR(m.dr,1,4))
  m.w   = INT(VAL(SUBSTR(CurString, AT(',',CurString,7)+1, AT(',',CurString,8)-AT(',',CurString,7)-1)))

  m.regdate = SUBSTR(CurString, AT(',',CurString,24)+1, AT(',',CurString,25)-AT(',',CurString,24)-1)
  m.upddate = SUBSTR(CurString, AT(',',CurString,25)+1, AT(',',CurString,26)-AT(',',CurString,25)-1)
  
  IF !IsCorrectDate(m.regdate)
   m.regdate = '1995-01-01 03:50:00'
  ENDIF 

  IF !IsCorrectDate(m.upddate)
   m.upddate = '1995-01-01 03:50:00'
  ENDIF 

  INSERT INTO People FROM MEMVAR 
 
 ENDDO 
 WAIT CLEAR 
 fhandl.Close

* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ '+TRANSFORM(m.HowManyRecs,'999999'), 0+64, 'Patient_Patient.csv')
 
* fhandl = fso.OpenTextFile(tc_comdir+'\NSI_SMO.csv')
* CurString = fhandl.ReadLine
* HowManyCommas = OCCURS(',', CurString)+1
 
 =OpenFile(tc_comdir+'\smo', 'smo', 'shar', 'q_ogrn')
 
* WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ SMO (NSI_SMO.csv)..." WINDOW NOWAIT 
* DO WHILE !fhandl.AtEndOfStream
*  CurString = fhandl.ReadLine
  
*  m.c_t = INT(VAL(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1)))
*  IF m.c_t!=77
*   LOOP 
*  ENDIF 

*  m.code   = SUBSTR(CurString, 1, AT(',',CurString,1)-1)
*  m.q_ogrn =SUBSTR(CurString, AT(',',CurString,13)+1, AT(',',CurString,14)-AT(',',CurString,13)-1)
  
*  IF SEEK(m.q_ogrn, 'smo')
*   REPLACE smo.code WITH m.code IN smo 
*  ENDIF 
  
* ENDDO 
* WAIT CLEAR 
* fhandl.Close
 
 fhandl = fso.OpenTextFile(tc_comdir+'\Patient_Policy.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1
 
* MESSAGEBOX(' ŒÀ-¬Œ —“ŒÀ¡÷Œ¬: '+STR(HowManyCommas,2), 0+64, '')
* m.HowManyRecs = 0
* m.HowManyRecs2 = 0
 
 SELECT people
 
 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PEOPLE (Patient_Policy.csv)..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  CurString = STRTRAN(CurString,'"','')
  
  m.id     = INT(VAL(ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))))
  m.policyuid =  INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.s_pol  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,3)+1, AT(',',CurString,4)-AT(',',CurString,3)-1))
  m.n_pol  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,4)+1, AT(',',CurString,5)-AT(',',CurString,4)-1))
  m.sn_pol = IIF(!EMPTY(m.s_pol), m.s_pol + ' ' + m.n_pol, m.n_pol)
  m.policystate = ALLTRIM(SUBSTR(CurString, AT(',',CurString,6)+1, AT(',',CurString,7)-AT(',',CurString,6)-1))
*  m.smoid  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,2)-1))
*  m.q = IIF(SEEK(m.smoid, 'smo', 'code'), smo.q, '')
   
  IF SEEK(m.id, 'people')
*   m.HowManyRecs = m.HowManyRecs + 1
*   REPLACE sn_pol WITH m.sn_pol, q WITH m.q, policyuid WITH m.policyuid, s_pol WITH m.s_pol, n_pol WITH m.n_pol,;
    smoid WITH INT(VAL(m.smoid))
   REPLACE sn_pol WITH m.sn_pol, policyuid WITH m.policyuid, s_pol WITH m.s_pol,n_pol WITH m.n_pol
   IF (EMPTY(Tip) AND !EMPTY(m.policystate)) OR (!EMPTY(Tip) AND (tip > m.policystate))
    REPLACE tip WITH m.policystate
   ENDIF 
  ELSE 
*   m.HowManyRecs2 = m.HowManyRecs2 + 1
  ENDIF 
  
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ '+TRANSFORM(m.HowManyRecs,'999999'), 0+64, 'Patient_Policy.csv')

* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ '+TRANSFORM(m.HowManyRecs2,'999999'), 0+64, 'Patient_Policy.csv')

 fhandl = fso.OpenTextFile(tc_comdir+'\EHR_book.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1

* m.HowManyRecs = 0

 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PEOPLE (EHR_book.csv)..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  
*  m.HowManyRecs = m.HowManyRecs + 1

  m.bookid = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.id     = ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1)) && PatientID
  m.c_i    = ALLTRIM(SUBSTR(CurString, AT(',',CurString,3)+1, AT(',',CurString,4)-AT(',',CurString,3)-1))
    
  IF SEEK(m.id, 'people')
   REPLACE c_i WITH m.c_i, bookid WITH m.bookid
  ENDIF 
  
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ '+TRANSFORM(m.HowManyRecs,'999999'), 0+64, 'EHR_book.csv')

 fhandl = fso.OpenTextFile(tc_comdir+'\Patient_Address.csv')
 CurString = fhandl.ReadLine
 HowManyCommas = OCCURS(',', CurString)+1

* m.HowManyRecs = 0

 WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PEOPLE (Patient_Address.csv)..." WINDOW NOWAIT 
 DO WHILE !fhandl.AtEndOfStream
  CurString = fhandl.ReadLine
  CurString = STRTRAN(CurString,'"','')
  
*  m.HowManyRecs = m.HowManyRecs + 1

  m.TipOfAddress = ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))
  
  IF m.TipOfAddress!='1'
   LOOP 
  ENDIF 
  
  m.id   = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
  m.c_t  = INT(VAL(ALLTRIM(SUBSTR(CurString, AT(',',CurString,5)+1, AT(',',CurString,6)-AT(',',CurString,5)-1))))
  m.kladr   = ALLTRIM(SUBSTR(CurString, AT(',',CurString,8)+1, AT(',',CurString,9)-AT(',',CurString,8)-1))
  m.ul   = IIF(SEEK(m.kladr, 'street', 'kladr'), street.ul, 0)
  m.dom  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,13)+1, AT(',',CurString,14)-AT(',',CurString,13)-1))
  m.kor  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,14)+1, AT(',',CurString,15)-AT(',',CurString,14)-1))
  m.str  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,15)+1, AT(',',CurString,16)-AT(',',CurString,15)-1))
  m.kv   = ALLTRIM(SUBSTR(CurString, AT(',',CurString,16)+1, AT(',',CurString,17)-AT(',',CurString,16)-1))
  m.adrid = INT(VAL(ALLTRIM(SUBSTR(CurString, AT(',',CurString,17)+1, AT(',',CurString,18)-AT(',',CurString,17)-1))))
*  m.moddate = ALLTRIM(SUBSTR(CurString, AT(',',CurString,20)+1, 19))
    
  IF SEEK(m.id, 'people')
   REPLACE c_t WITH m.c_t, ul WITH m.ul, dom WITH m.dom, kor WITH m.kor, str WITH m.str, kv WITH m.kv, ;
    adrid WITH m.adrid, kladr WITH m.kladr
  ENDIF 
  
 ENDDO 
 WAIT CLEAR 
 fhandl.Close
* MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ '+TRANSFORM(m.HowManyRecs,'999999'), 0+64, 'Patient_Address.csv')

* fhandl = fso.OpenTextFile(tc_comdir+'\MskOMS_Policy.csv')
* CurString = fhandl.ReadLine
* HowManyCommas = OCCURS(',', CurString)+1

* WAIT "«¿œŒÀÕ≈Õ»≈ ‘¿…À¿ PEOPLE (MskOMS_Policy.csv)..." WINDOW NOWAIT 
* DO WHILE !fhandl.AtEndOfStream
*  CurString = fhandl.ReadLine
  
*  m.id   = INT(VAL(SUBSTR(CurString, 1, AT(',',CurString,1)-1)))
*  m.v_erz  = ALLTRIM(SUBSTR(CurString, AT(',',CurString,1)+1, AT(',',CurString,2)-AT(',',CurString,1)-1))
*  m.d_erz  = SUBSTR(CurString, AT(',',CurString,2)+1, AT(',',CurString,3)-AT(',',CurString,1)-1)
*  m.d_erz  = CTOD(SUBSTR(m.d_erz,9,2)+'.'+SUBSTR(m.d_erz,6,2)+'.'+SUBSTR(m.d_erz,1,4))
    
*  IF SEEK(m.id, 'people')
*   REPLACE v_erz WITH m.v_erz, d_erz WITH m.d_erz
*  ENDIF 
  
* ENDDO 
* WAIT CLEAR 
* fhandl.Close

 SELECT people
 SCAN 
  SCATTER MEMVAR 
  m.ok = .t.
  m.ok = IIF(!SEEK(m.c_t, 'territ'), .f., m.ok)
  m.ok = IIF(EMPTY(policyuid), .f., m.ok)
  m.ok = IIF(!INLIST(m.w,1,2), .f., m.ok)
  m.ok = IIF(!SEEK(m.ul, 'street'), .f., m.ok)
  m.ok = IIF(c_t=77 and EMPTY(m.kv), .f., m.ok)
  m.ok = IIF(MONTH(m.dr)=1 and DAY(m.dr)=1, .f., m.ok)
  m.ok = IIF(m.c_t=77 and (!IsKms(m.sn_pol) and !IsVs(m.sn_pol) and !IsEnp(m.sn_pol)), .f., m.ok)
  
  REPLACE ok WITH m.ok
 ENDSCAN 

 USE IN people
 USE IN street 
 USE IN territ
 USE IN smo 

 IF fso.FileExists(tc_comdir+'\Patient_Patient.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Patient.csv')
 ENDIF 
 IF fso.FileExists(tc_comdir+'\Patient_Policy.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Policy.csv')
 ENDIF 
 IF fso.FileExists(tc_comdir+'\Patient_Address.csv') 
  fso.DeleteFile(tc_comdir+'\Patient_Address.csv')
 ENDIF 
 IF fso.FileExists(tc_comdir+'\EHR_Book.csv') 
  fso.DeleteFile(tc_comdir+'\EHR_Book.csv')
 ENDIF 
* IF fso.FileExists(tc_comdir+'\MskOMS_policy.csv') 
*  fso.DeleteFile(tc_comdir+'\MskOMS_policy.csv')
* ENDIF 
* IF fso.FileExists(tc_comdir+'\MskOMS_Street.csv') 
*  fso.DeleteFile(tc_comdir+'\MskOMS_Street.csv')
* ENDIF 
* IF fso.FileExists(tc_comdir+'\NSI_SMO.csv') 
*  fso.DeleteFile(tc_comdir+'\NSI_SMO.csv')
* ENDIF 
 
 MESSAGEBOX('Œ¡–¿¡Œ“¿ÕŒ: '+STR(m.HowManyRecs,6)+' «¿œ»—≈…!', 0+64, '')
