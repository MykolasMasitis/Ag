PROCEDURE UpdateVector
 
 IF !fso.FileExists(tc_comdir+'\people.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик PEOPLE.DBF!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 AMFile = 'Am_'+PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)
 IF !fso.FileExists(tc_comdir+'\'+AMFile+'.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрярбсер тюик '+tc_comdir+'\'+AMFile+'!'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 
 
 ffile = fso.GetFile(tc_comdir+'\'+AMFile+'.dbf')
 CreateDate = ffile.DateLastModified
 RELEASE m.ffile 
 IF DATETIME()-CreateDate>12*60*60
  IF MESSAGEBOX(CHR(13)+CHR(10)+; 
   'тюик нрберю '+ALLTRIM(tc_comdir+'\'+AmFile)+CHR(13)+CHR(10)+; 
   'ашк янгдюм '+TTOC(CreateDate) +'.'+CHR(13)+CHR(10)+; 
   'яеивюя '+TTOC(DATE())+'.'+CHR(13)+CHR(10)+; 
   CHR(13)+CHR(10)+'напюаюршбюрэ бяе пюбмн?',4+32, 'бмхлюмхе') == 7
   RETURN 
  ENDIF 
 ENDIF 

 IF OpenFile(tc_comdir+'\people', 'people', 'shar')>0
  IF USED('people')
   USE IN people
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\'+AmFile, 'answer', 'excl')>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('answer')
   USE IN answer
  ENDIF 
  RETURN 
 ENDIF 
 IF OpenFile(tc_comdir+'\smo', 'smo', 'shar', 'q')>0
  IF USED('people')
   USE IN people
  ENDIF 
  IF USED('answer')
   USE IN answer
  ENDIF 
  IF USED('smo')
   USE IN smo
  ENDIF 
  RETURN 
 ENDIF 

 WAIT "напюанрйю..." WINDOW NOWAIT 
 SELECT answer
 INDEX on recid TAG recid 
 SET ORDER TO recid
 SELECT people
 SET RELATION TO PADL(RECNO(),6,'0') INTO answer
 
 COUNT FOR !EMPTY(answer.recid) TO m.nRelated
 COUNT FOR c_t=77 TO m.RecsInPeople

 WAIT CLEAR 
 
 IF m.RecsInPeople != m.nRelated
  IF MESSAGEBOX(CHR(13)+CHR(10)+'йнк-бн гюохяеи б пецхярпе ('+STR(m.RecsInPeople,5)+')'+CHR(13)+CHR(10)+;
   'ме янбоюдюер я йнк-бнл гюохяеи б нрбере ('+STR(m.nRelated,5)+')'+CHR(13)+CHR(10)+;
   'опнднкфхрэ?'+CHR(13)+CHR(10), 4+32, '') == 7
   SET RELATION OFF INTO answer
   USE 
   SELECT answer 
   SET ORDER TO 
   DELETE TAG ALL 
   USE 
   RETURN 
  ENDIF 
 ENDIF 
 
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

 WAIT "намскемхе рюакхжш T_POLICY_53495..." WINDOW NOWAIT 
 DelString = "truncate table t_policy_53495"
 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process truncate!', 16, 't_policy_53495')
 ELSE 
*  = MESSAGEBOX('Table truncated!', 48, 't_policy_53495')
 ENDIF 
 WAIT CLEAR 

 WAIT "намскемхе рюакхжш T_PATIENT_45426..." WINDOW NOWAIT 
 DelString = "truncate table t_patient_45426"
 lrslt = SQLEXEC(lnHndl, delstring)
 IF lrslt<=0
  MESSAGEBOX('Cannot process truncate!', 16, 't_patient_45426')
 ELSE 
*  = MESSAGEBOX('Table truncated!', 48, 't_patient_45426')
 ENDIF 
 WAIT CLEAR 

 m.v_beg      = '34932767000000'
 m.v_end      = '9223372036854775807'
 m.is_top     = 1
 m.created_by = 99
 
 WAIT "напюанрйю..." WINDOW NOWAIT 
 SCAN FOR policyuid!=0
  IF !EMPTY(answer.recid)
*   m.qerz = answer.q
   m.v_erz = answer.ans_r
   m.policyuid = policyuid
   m.patientuid = id
   m.checkdate = d_erz2
   m.s_pol = answer.s_pol
   m.n_pol = answer.n_pol
   m.quid = IIF(SEEK(answer.q, 'smo'), smo.code, 0)

   InsCom1 = "insert into t_policy_53495 "
   InsCom2 = "(version_begin,version_end,is_top,created_by,policyuid_26734,checkvector_17389,;
    checkdate_08921,polseries_05608,polnumber_37881,smouid_65250) "
   InsCom3 = "values "
   InsCom4 = "(?m.v_beg,?m.v_end,?m.is_top,?m.created_by,?m.policyuid,?m.v_erz,?m.checkdate,;
    ?m.s_pol,?m.n_pol,?m.quid)"
   InsCom5 = ""
   InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
   lrslt = SQLEXEC(lnHndl, InsCom)
   IF lrslt<=0
    =AERROR(errarr)
    ErrMsg = ALLTRIM(errarr(2))
*    = MESSAGEBOX(ErrMsg, 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX('m.policyuid='+STR(m.policyuid), 16, '')
*    = MESSAGEBOX('Cannot process insert!', 16, 't_policy_53495')
    EXIT 
   ELSE 
*    m.ndsadded = m.ndsadded + 1
   ENDIF 

   InsCom1 = "insert into t_patient_45426 "
   InsCom2 = "(version_begin,version_end,is_top,created_by,patientuid_29790,checkvector_05009,checkdate_58618,speccase_10311) "
   InsCom3 = "values "
   InsCom4 = "(?m.v_beg,?m.v_end,?m.is_top,?m.created_by,?m.patientuid,?m.v_erz,?m.checkdate,'0')"
   InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4
   lrslt = SQLEXEC(lnHndl, InsCom)
   IF lrslt<=0
    =AERROR(errarr)
    ErrMsg = ALLTRIM(errarr(2))
    = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX('m.policyuid='+STR(m.policyuid), 16, '')
    EXIT 
   ELSE 
   ENDIF 

  ENDIF 
 ENDSCAN 
 WAIT CLEAR 

 SET RELATION OFF INTO answer
 SELECT answer 
 SET ORDER TO 
 DELETE TAG ALL 
 USE 
 USE IN people
 USE IN smo

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