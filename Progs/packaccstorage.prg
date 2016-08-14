PROCEDURE PackAccStorage
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
 
 m.DBName = "PDPAccStorage"
 lrslt = SQLEXEC(lnHndl, "USE "+m.DBName)
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, m.DBName)
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, m.DBName)
 ENDIF 

 WAIT "намскемхе рюакхжш T_INVSERVICE_07640..." WINDOW NOWAIT 
 DelString = "truncate table t_invservice_07640"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invservice_07640')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invservice_07640')
 ENDIF 
 WAIT CLEAR 

* WAIT "намнбкемхе LAST_UID..." WINDOW NOWAIT 
* UpdString = "update table last_uid set uid=1 where entity='t_invservice_07640'"

* lrslt = SQLEXEC(lnHndl, updstring)
* IF lrslt<=0
*  MESSAGEBOX('Cannot process update!', 16, 'last_uid')
* ELSE 
**  = MESSAGEBOX('Records updated!', 48, 'last_uid')
* ENDIF 
* WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVSERVAUDIT_56858..." WINDOW NOWAIT 
 DelString = "truncate table t_invservaudit_56858"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invservaudit_56858')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invservaudit_56858')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVPATIENT_21972..." WINDOW NOWAIT 
 DelString = "truncate table t_invpatient_21972"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invpatient_21972')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invpatient_21972')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVPATIENTAUDIT_54011..." WINDOW NOWAIT 
 DelString = "truncate table t_invpatientaudit_54011"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invpatientaudit_54011')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invpatientaudit_54011')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVSPECIAL_21849..." WINDOW NOWAIT 
 DelString = "truncate table t_invspecial_21849"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invspecial_21849')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invspecial_21849')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVTARIFF_29106..." WINDOW NOWAIT 
 DelString = "truncate table t_invtariff_29106"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invtariff_29106')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invtariff_29106')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVUNIT_07314..." WINDOW NOWAIT 
 DelString = "truncate table t_invunit_07314"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invunit_07314')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invunit_07314')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVFINALSERVICE_61755..." WINDOW NOWAIT 
 DelString = "truncate table t_invfinalservice_61755"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invfinalservice_61755')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invfinalservice_61755')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_INVFINTOTAL_41962..." WINDOW NOWAIT 
 DelString = "truncate table t_invfintotal_41962"

 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process delete!', 16, 't_invfintotal_41962')
 ELSE 
*  = MESSAGEBOX('Records deleted!', 48, 't_invfintotal_41962')
 ENDIF 
 WAIT CLEAR 

* WAIT "намскемхе рюакхжш LAST_UID..." WINDOW NOWAIT 
* DelString = "truncate table last_uid"

* lrslt = SQLEXEC(lnHndl, delstring)
* IF lrslt<=0
*  MESSAGEBOX('Cannot process delete!', 16, 'last_uid')
* ELSE 
**  = MESSAGEBOX('Records deleted!', 48, 'last_uid')
* ENDIF 
* WAIT CLEAR 

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