PROCEDURE MakeTalonSQL
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ВЫ ХОТИТЕ ПОСТРОИТЬ ФАЙЛ TALON'+;
  CHR(13)+CHR(10)+'ИЗ БД MS SQL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF fso.FileExists(tc_basedir+'\Talon.dbf')
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ ТАЛОН УЖЕ СФОРМИРОВАН!'+CHR(13)+CHR(10)+;
   'ВЫ ХОТИТЕ СФОРМИРОВАТЬ ЕГО ЗАНОВО?',4+32,'') = 7
    RETURN 
  ELSE 
   =OpenFile(tc_basedir+'\talon', 'talon', 'excl')
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
 
 WAIT "ПОДКЛЮЧЕНИЕ К СЕРВЕРУ БАЗЫ ДАННЫХ...." WINDOW NOWAIT 
 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 
 lrslt = SQLEXEC(lnHndl, "USE PDPAccStorage")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPAccStorage')
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPAccStorage')
 ENDIF 

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

* SelString = "select uid, CAST(version_begin as numeric(19)) as v_beg, tapuid_30432 as tapuid,;
  icdcode_39884 as icd from t_tapdiagnosis_31571 where is_top=1 and deleted_by is null"
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

* SelString = "select uid, CAST(version_begin as numeric(19)) as v_beg, patientuid_40511 as patuid, ;
  fillingdate_36966 as d_u, doctoruid_47963 as doc1, nurseuid_48406 as doc2 ;
   from t_tap_10874 where is_top=1 and deleted_by is null and medicalprogramm_28647='ОМСМК'"
 SelString = "select uid, version_begin as v_beg10874, patientuid_40511 as patuid, ;
  fillingdate_36966 as d_u, doctoruid_47963 as doc1, nurseuid_48406 as doc2 ;
   from t_tap_10874 where version_end=9223372036854775807 and medicalprogramm_28647='ОМСМК'"

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

* WAIT "T_TAPSERVICE_63336..." WINDOW NOWAIT 

* SelString = "select version_begin as v_beg, diagnosisuid_59030 as diaguid, servicecode_18926 as codusl, ;
   speccase_39152 as speccase from t_tapservice_63336 where version_end=9223372036854775807"

* lrslt = SQLEXEC(lnHndl, selstring)
* IF lrslt<=0
*  = MESSAGEBOX('Cannot process select!', 16, 't_tapservice_63336')
* ELSE 
*  SCAN 
*   m.diaguid = diaguid
*   m.codusl = codusl
*   m.v_beg = v_beg
*   m.d_type = IIF(ISNULL(speccase), '0', ALLTRIM(speccase))
*   UPDATE talon SET v_beg63336=m.v_beg, d_type=m.d_type WHERE diaguid=m.diaguid AND coduid=PADR(ALLTRIM(m.codusl),8)
*  ENDSCAN 
* ENDIF 

* WAIT CLEAR 

 WAIT "ОТКЛЮЧЕНИЕ ОТ БАЗЫ ДАННЫХ..." WINDOW NOWAIT 

 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && Соединение успешно закрыто
*   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && Ошибка уровня соединения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && Ошибка уровня окружения
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && Такого быть не должно, но все-таки...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 

RELEASE lnhndl, lresult, ln_disconn

USE IN tarif
USE IN talon 

WAIT "СОРТИРОВКА TALON..." WINDOW NOWAIT 
=OpenFile(tc_basedir+'\talon', 'talon', 'excl', 'tapuid')
*=OpenFile(tc_comdir+'\people', 'people', 'shar', 'id')
SELECT talon 
*SET RELATION TO tapuid INTO invservice ADDITIVE 
*COPY TO tc_basedir+'\ttemp' FOR BETWEEN(d_u,m.d_beg,m.d_end) AND EMPTY(invservaudit.code)
COPY TO tc_basedir+'\ttemp' FOR BETWEEN(d_u,m.d_beg,m.d_end)
ZAP
APPEND FROM tc_basedir+'\ttemp'
m.SumOfAcc = 0
SCAN 
 m.patientuid = patientuid
 m.isfound = IIF(SEEK(m.patientuid, 'people'), .t., .f.)
 m.dr = IIF(SEEK(m.patientuid, 'people'), people.dr, {})
 m.w = IIF(SEEK(m.patientuid, 'people'), people.w, 0)
  
 IF !m.IsFound OR INLIST(cod,1561,1601) OR ;
    BETWEEN(ds,'Z32.0 ','Z39.2 ') OR ;
    BETWEEN(ds,'O20.0 ','O99.8 ') OR ;
    BETWEEN(cod,1211,1227) OR BETWEEN(cod,9001,9362)
  m.ok = .F.
 ELSE
  m.ok = .T.
 ENDIF 
 REPLACE isfound WITH m.isfound, dr WITH m.dr, w WITH m.w, ok WITH m.ok
 m.SumOfAcc = m.SumOfAcc + s_all

ENDSCAN 
*SET RELATION OFF INTO invservice
USE 
USE IN people
*USE IN invservice
*USE IN invservaudit

fso.DeleteFile(tc_basedir+'\ttemp.dbf')
WAIT CLEAR 

MESSAGEBOX('СУММА СЧЕТА: '+TRANSFORM(m.SumOfAcc, '99 999 999.99'),0+64,'')

 
RETURN 