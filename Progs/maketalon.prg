PROCEDURE MakeTalon
 IF MESSAGEBOX(CHR(13)+CHR(10)+'¬€ ’Œ“»“≈ œŒ—“–Œ»“‹ ‘¿…À TALON'+;
  CHR(13)+CHR(10)+'»« ¡ƒ MS SQL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF !fso.FileExists(tc_comdir+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PEOPLE.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\people', 'people', 'shar', 'id')>0
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\osoerzxx', 'osoerz', 'shar', 'ans_r')>0
  IF USED('osoerz')
   USE IN osoerz
  ENDIF 
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 
 SELECT people
 REPLACE ALL Used WITH .f. 
 USE IN people

 IF fso.FileExists(tc_basedir+'\Talon.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'‘¿…À “¿ÀŒÕ ”∆≈ —‘Œ–Ã»–Œ¬¿Õ!'+CHR(13)+CHR(10)+;
   '¬€ ’Œ“»“≈ —‘Œ–Ã»–Œ¬¿“‹ ≈√Œ «¿ÕŒ¬Œ?',4+32,'') = 7
    RETURN 
  ELSE 
   IF OpenFile(tc_basedir+'\talon', 'talon', 'excl')>0
    IF USED('Talon')
     USE IN Talon
    ENDIF 
    RETURN 
   ENDIF 
   SELECT talon 
   DELETE TAG ALL 
   USE 

   iii = 1
   TalonBak = 'Talon.'+PADL(iii,3,'0')
   DO WHILE fso.FileExists(tc_basedir+'\'+TalonBak)
    iii = iii + 1
    TalonBak = 'Talon.'+PADL(iii,3,'0')
   ENDDO 
   fso.CopyFile(tc_basedir+'\Talon.dbf', tc_basedir+'\'+TalonBak)
   fso.DeleteFile(tc_basedir+'\Talon.dbf')
  ENDIF 
 ENDIF 
 
 IF fso.FileExists(tc_basedir+'\men.dbf')
  fso.DeleteFile(tc_basedir+'\men.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\women.dbf')
  fso.DeleteFile(tc_basedir+'\women.dbf')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\men.cdx')
  fso.DeleteFile(tc_basedir+'\men.cdx')
 ENDIF 
 IF fso.FileExists(tc_basedir+'\women.cdx')
  fso.DeleteFile(tc_basedir+'\women.cdx')
 ENDIF 
 
 m.d_beg = m.tdat1
 m.d_end = m.tdat2
 m.choice = .f.
 DO FORM SelectPeriod
 IF m.choice!=.t.
  RETURN 
 ENDIF 
 m.cd_beg = STR(YEAR(m.d_beg),4)+PADL(MONTH(m.d_beg),2,'0')+PADL(DAY(m.d_beg),2,'0')
 m.cd_end = STR(YEAR(m.d_end),4)+PADL(MONTH(m.d_end),2,'0')+PADL(DAY(m.d_end),2,'0')

 SET NULL OFF 
 
 WAIT "œŒƒ Àﬁ◊≈Õ»≈   —≈–¬≈–” ¡¿«€ ƒ¿ÕÕ€’...." WINDOW NOWAIT 
 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 
* lrslt = SQLEXEC(lnHndl, "USE PDPAccStorage")
* IF lrslt<=0
*  = MESSAGEBOX('Cannot select database!', 16, 'PDPAccStorage')
*  RETURN 
* ELSE 
**  = MESSAGEBOX('Database selected!', 48, 'PDPAccStorage')
* ENDIF 

* WAIT "T_INVSERVICE_07640..." WINDOW NOWAIT 
* CREATE TABLE &tc_comdir\invservice (uid i, tapuid i)
* INDEX on tapuid TAG tapuid
* SET ORDER TO tapuid 
* SelString01 = "select MAX(uid) as uid, tapuid_52112 as tapuid from t_invservice_07640 ;
  where d_u_50426 between ?m.cd_beg and ?m.cd_end group by tapuid_52112"
* SelString01 = "select uid as uid, tapuid_52112 as tapuid from t_invservice_07640 ;
  where d_u_50426 between ?m.cd_beg and ?m.cd_end"
* SelString02 = ""
* SelString = SelString01 + SelString02
* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_invservice_07640')
* ELSE 
*  SCAN 
*   SCATTER MEMVAR 
*   INSERT INTO invservice FROM MEMVAR 
*  ENDSCAN 
*  USE 
*  USE IN invservice
* ENDIF 
* WAIT CLEAR 

* WAIT "T_INVSERVAUDIT_56858..." WINDOW NOWAIT 
* CREATE TABLE &tc_comdir\invservaudit (invservuid i, code c(2))
* INDEX ON invservuid TAG invservuid
* SET ORDER TO invservuid
* SelString01 = "select invserviceuid_22043 as invservuid, code_43906 as code from t_invservaudit_56858;
  where fatal_02118=1"
* SelString01 = "select invserviceuid_22043 as invservuid, code_43906 as code from t_invservaudit_56858"
* SelString02 = ""
* SelString = SelString01 + SelString02
* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_invservaudit_56858')
* ELSE 
*  SCAN 
*   SCATTER MEMVAR 
*   INSERT INTO invservaudit FROM MEMVAR 
*  ENDSCAN 
*  USE 
*  USE IN invservaudit
* ENDIF 
* WAIT CLEAR 
 
* SELECT invservice
* SET RELATION TO uid INTO invservaudit

 lrslt = SQLEXEC(lnHndl, "USE PDPStdStorage")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPStdStorage')
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPStdStorage')
 ENDIF 

 CREATE TABLE &tc_basedir\Talon ;
  (ok l, tapuid i, ntapuid i, patientuid i, npatuid i, dr d, newdr d, w n(1), neww n(1), isfound l, ;
   isreplace l, diaguid i, ndiaguid i, ds c(6), d_u d, coduid c(8), cod n(6), k_u n(3), doctoruid i, nurseuid i, s_all n(11,2),;
   d_type c(1), v_beg61560 c(20), v_beg31571 c(20), v_beg10874 c(20))
 USE 
 USE &tc_basedir\Talon IN 0 ALIAS talon EXCLUSIVE 
 SELECT talon 
 INDEX ON tapuid TAG tapuid 
 INDEX ON patientuid TAG patientuid
 INDEX ON diaguid TAG diaguid
 
 =OpenFile(tc_comdir+'\TarifN', 'tarif', 'shar', 'cod')
 =OpenFile(tc_comdir+'\people', 'people', 'shar', 'id')
 SELECT people
 SET RELATION TO v_erz2 INTO osoerz

 WAIT "T_TAPSERVICE_61560..." WINDOW NOWAIT 

 SelString01 = "select version_begin as v_beg61560, version_end as v_end61560, diagnosisuid_34765 as diaguid,; 
  servicecode_20924 as coduid, CAST(count_23546 as numeric(3)) as k_u "
 SelString02 = "from t_tapservice_61560 where version_end=9223372036854775807 order by version_begin" 
 SelString = SelString01 + SelString02
 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_tapservice_61560')
 ELSE 
  SCAN 
   m.diaguid = diaguid
   m.coduid  = coduid
   m.cod     = INT(VAL(ALLTRIM(SUBSTR(m.coduid,3))))
   m.k_u     = k_u
   m.s_all   = IIF(SEEK(m.cod, 'tarif'), tarif.tarif*m.k_u, 0)
   m.v_beg61560 = ALLTRIM(v_beg61560)
   m.v_end61560 = ALLTRIM(v_end61560)
   
   INSERT INTO talon FROM MEMVAR 
   
  ENDSCAN 
  USE 
 ENDIF 

 WAIT CLEAR 

 WAIT "T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 SelString = "select uid, version_begin as v_beg31571, tapuid_30432 as tapuid,;
  icdcode_39884 as icd from t_tapdiagnosis_31571 where version_end=9223372036854775807"

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_tapdiagnosis_31571')
 ELSE 
*  = MESSAGEBOX('Records selected!', 48, 't_tapdiagnosis_31571')
*  COPY TO &tc_basedir\tapdiagnosis_31571
  SET ORDER to diaguid IN talon 
  SCAN 
   m.diaguid = uid
   m.tapuid = tapuid
   m.ds = icd
   m.v_beg31571 = v_beg31571
   UPDATE talon SET tapuid=m.tapuid, ds=m.ds, v_beg31571=m.v_beg31571 WHERE diaguid=m.diaguid
  ENDSCAN 
  USE 
 ENDIF 

 WAIT CLEAR 

 WAIT "TAP_10874..." WINDOW NOWAIT 
 SelString = "select uid, version_begin as v_beg10874, patientuid_40511 as patuid, ;
  fillingdate_36966 as d_u, doctoruid_47963 as doc1, nurseuid_48406 as doc2 ;
   from t_tap_10874 where version_end=9223372036854775807 and medicalprogramm_28647='ŒÃ—Ã '"

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 'tap_10874')
 ELSE 
*  = MESSAGEBOX('Records selected!', 48, 'tap_10874')
*  COPY TO &tc_basedir\tap_10874
  SET ORDER to tapuid IN talon 
  SCAN 
   m.tapuid = uid
   m.patientuid = patuid
   m.d_u = TTOD(d_u)
   m.doctoruid = IIF(ISNULL(doc1), 0, doc1)
   m.nurseuid = IIF(ISNULL(doc2), 0, doc2)
   m.v_beg10874 = v_beg10874
   UPDATE talon SET patientuid=m.patientuid, d_u=m.d_u, doctoruid=m.doctoruid, nurseuid=m.nurseuid, ;
    v_beg10874=m.v_beg10874 WHERE tapuid=m.tapuid
  ENDSCAN 
  USE 
 ENDIF 

 WAIT CLEAR 


 WAIT "Œ“ Àﬁ◊≈Õ»≈ Œ“ ¡¿«€ ƒ¿ÕÕ€’..." WINDOW NOWAIT 

 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && —ÓÂ‰ËÌÂÌËÂ ÛÒÔÂ¯ÌÓ Á‡Í˚ÚÓ
*   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && Œ¯Ë·Í‡ ÛÓ‚Ìˇ ÒÓÂ‰ËÌÂÌËˇ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && Œ¯Ë·Í‡ ÛÓ‚Ìˇ ÓÍÛÊÂÌËˇ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && “‡ÍÓ„Ó ·˚Ú¸ ÌÂ ‰ÓÎÊÌÓ, ÌÓ ‚ÒÂ-Ú‡ÍË...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 

RELEASE lnhndl, lresult, ln_disconn

USE IN tarif
USE IN talon 

WAIT "—Œ–“»–Œ¬ ¿ TALON..." WINDOW NOWAIT 
*=OpenFile(tc_basedir+'\talon', 'talon', 'excl', 'tapuid')
=OpenFile(tc_basedir+'\talon', 'talon', 'excl')
SELECT * FROM talon WHERE BETWEEN(d_u,m.d_beg,m.d_end) ORDER BY patientuid, d_u asc, cod ASC INTO CURSOR curbnm
COPY STRUCTURE TO tc_basedir+'\ttemp'
=OpenFile(tc_basedir+'\ttemp', 'ttemp', 'excl')
SELECT curbnm

SCAN 
 SCATTER MEMVAR 

 IF INLIST(m.cod,1561,1601) OR ;
   BETWEEN(m.ds,'Z32.0 ','Z39.2 ') OR ;
   BETWEEN(m.ds,'O20.0 ','O99.8 ')
*  OR BETWEEN(m.cod,1211,1227) OR BETWEEN(m.cod,9001,9362)
  LOOP 
 ENDIF
 
 INSERT INTO ttemp FROM MEMVAR 
ENDSCAN 
*COPY TO tc_basedir+'\ttemp' FOR BETWEEN(d_u,m.d_beg,m.d_end)
USE 
USE IN ttemp
SELECT talon 
ZAP
APPEND FROM tc_basedir+'\ttemp'
m.SumOfAcc = 0

m.ptuid=0
SCAN 
 m.patientuid = patientuid
 IF m.patientuid!=m.ptuid
  m.ptuid = m.patientuid
  m.cod = cod
  m.ok = .t.
 ENDIF 
 m.isfound = .f.
 m.dr      = {}
 m.w       = 0
 IF SEEK(m.patientuid, 'people')
  m.isfound = .t.
  m.dr      = people.dr
  m.w       = people.w
  REPLACE used WITH .t. IN people 
 ENDIF 

 REPLACE isfound WITH m.isfound, dr WITH m.dr, w WITH m.w, ok WITH .f.
 
 IF !m.ok
  LOOP 
 ENDIF 
 
 m.ok = .f.
 IF people.c_t=77 AND SEEK(people.v_erz2, 'osoerz')
  IF osoerz.kl='y'
   m.ok = .t.
  ENDIF 
 ENDIF 
 
 REPLACE IsFound WITH m.IsFound, dr WITH m.dr, w WITH m.w, OK WITH m.ok
 m.SumOfAcc = m.SumOfAcc + s_all

ENDSCAN 
USE 
USE IN people
USE IN osoerz

fso.DeleteFile(tc_basedir+'\ttemp.dbf')
WAIT CLEAR 

MESSAGEBOX('—”ÃÃ¿ —◊≈“¿: '+TRANSFORM(m.SumOfAcc, '99 999 999.99'),0+64,'')

 
RETURN 