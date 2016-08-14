PROCEDURE KillImportTrace
 SET NULL OFF 
 
 WAIT "ондйкчвемхе й яепбепс аюгш дюммшу...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 WAIT CLEAR 
 
 m.DBName = "PDPStdStorage"
 lrslt = SQLEXEC(lnHndl, "USE "+m.DBName)
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, m.DBName)
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, m.DBName)
 ENDIF 

 SelString = "select distinct sourceuid_32508 as suid from t_tap_08603"
 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process select!', 16, 't_tap_08603')
 ELSE 
*  = MESSAGEBOX('Records selected!', 48, 't_tap_08603')
 ENDIF 

 WAIT "намскемхе рюакхжш DBO.T_TAP_08603..." WINDOW NOWAIT 
* DelString = "delete t_tap_08603"
 DelString = "truncate table t_tap_08603"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_tap_08603')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_tap_08603')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_TAPDIAGNOSIS_49193..." WINDOW NOWAIT 
* DelString = "delete t_tap_08603"
 DelString = "truncate table t_tapdiagnosis_49193"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_tapdiagnosis_49193')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_tapdiagnosis_49193')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_POLICY_55481..." WINDOW NOWAIT 
 DelString = "truncate table t_policy_55481"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_policy_55481')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_policy_55481')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_PATIENT_43594..." WINDOW NOWAIT 
 DelString = "truncate table t_patient_43594"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_patient_43594')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_patient_43594')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_ADDRESS_23832..." WINDOW NOWAIT 
 DelString = "truncate table t_address_23832"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_address_23832')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_address_23832')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_BOOK_54341..." WINDOW NOWAIT 
 DelString = "truncate table t_book_54341"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_book_54341')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_book_54341')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_EMPLOYEE_18047..." WINDOW NOWAIT 
 DelString = "truncate table t_employee_18047"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_employee_18047')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_employee_18047')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_SERVICEPROVIDER_04118.." WINDOW NOWAIT 
 DelString = "truncate table t_serviceprovider_04118"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_serviceprovider_04118')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_serviceprovider_04118')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_ORGUNIT_02795.." WINDOW NOWAIT 
 DelString = "truncate table t_orgunit_02795"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_orgunit_02795')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_orgunit_02795')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_COUNTCONSTRAINT_18495.." WINDOW NOWAIT 
 DelString = "truncate table t_countconstraint_18495"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_countconstraint_18495')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_countconstraint_18495')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_DETACHEDSERVICE_25609.." WINDOW NOWAIT 
 DelString = "truncate table t_detachedservice_25609"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_detachedservice_25609')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_detachedservice_25609')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш DBO.T_DUL_41245..." WINDOW NOWAIT 
 DelString = "truncate table t_dul_41245"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_dul_41245')
 ELSE 
  = MESSAGEBOX('Records deleted!', 48, 't_dul_41245')
 ENDIF 
 WAIT CLEAR 

 SELECT sqlresult
 IF RECCOUNT()>0
  m.DBName = "PDPDirectoryNet"
  lrslt = SQLEXEC(lnHndl, "USE "+m.DBName)
  IF lrslt<=0
   = MESSAGEBOX('Cannot select database!', 16, m.DBName)
  ELSE 
   SCAN 
    m.suid = suid
    SelString = "delete from TNode where uid=?m.suid"
    lrslt = SQLEXEC(lnHndl, selstring)
    IF lrslt<=0
     MESSAGEBOX('Cannot delete uid='+ALLTRIM(STR(m.suid))+'!', 16, 'TNode')
    ELSE 
     MESSAGEBOX('uid '+ALLTRIM(STR(m.suid))+' deleted!', 48, 'TNode')
    ENDIF 
   ENDSCAN 
  ENDIF 
 ENDIF 
 USE 

 WAIT "нрйкчвемхе нр яепбепю..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && яНЕДХМЕМХЕ СЯОЕЬМН ГЮЙПШРН
*  = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && нЬХАЙЮ СПНБМЪ ЯНЕДХМЕМХЪ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && нЬХАЙЮ СПНБМЪ НЙПСФЕМХЪ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && рЮЙНЦН АШРЭ МЕ ДНКФМН, МН БЯЕ-РЮЙХ...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 
 WAIT CLEAR 

 RELEASE lnhndl, lresult, ln_disconn
 
RETURN 