PROCEDURE LoadTalon
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре гюцпсгхрэ дюммше'+;
  CHR(13)+CHR(10)+'б ад MS SQL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF !fso.FileExists(tc_basedir+'\Talon.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'нрясрябсер тюик TALON.DBF'+CHR(13)+CHR(10),0+16,'')
  RETURN 
 ENDIF 

 =OpenFile(tc_basedir+'\Talon', 'talon', 'shar', 'tapuid')
 =OpenFile(tc_comdir+'\People', 'People', 'shar', 'id')
 
 SELECT talon 
 PUBLIC m.SetChoice, m.GoodSum, m.ChoosedSum
 m.ChoosedSum = 0 
* SUM s_all TO m.IsSum
 SUM FOR ok s_all TO m.GoodSum 

 m.SetChoice = 0
 DO FORM SetSum
 IF m.SetChoice=2
  USE IN people
  USE IN talon
  RETURN 
 ENDIF 
 
 IF m.ChoosedSum<=0
  USE IN people
  USE IN talon
  RETURN 
 ENDIF 
 
 IF m.ChoosedSum > m.GoodSum
  MESSAGEBOX('люйяхлюкэмюъ ясллю: '+TRANSFORM(m.GoodSum,'99 999 999.99')+' пса.',0+16,'')
  USE IN people
  USE IN talon
  RETURN 
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
 
 lrslt = SQLEXEC(lnHndl, "USE PDPStdStorage")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPStdStorage')
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPStdStorage')
 ENDIF 
 
 WAIT "бшвхякемхе люйяхлюкэмнцн UID T_TAP_10874..." WINDOW NOWAIT 
 m.maxtapuid = 0
 SelCom = "select MAX(uid) as uid from t_tap_10874"
 lrslt = SQLEXEC(lnHndl, SelCom, 'sqlres')
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 'tap_10874')
 ELSE 
  m.maxtapuid = IIF(RECCOUNT('sqlres')==1, IIF(!ISNULL(sqlres.uid),sqlres.uid,0) , 0)
  USE IN sqlres
 ENDIF 
 WAIT CLEAR 
 
 WAIT "бшвхякемхе люйяхлюкэмнцн UID T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 m.maxdsuid = 0
 SelCom = "select MAX(uid) as uid from t_tapdiagnosis_31571"
 lrslt = SQLEXEC(lnHndl, SelCom, 'sqlres')
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 't_tapdiagnosis_31571')
 ELSE 
  m.maxdsuid = IIF(RECCOUNT('sqlres')==1, IIF(!ISNULL(sqlres.uid),sqlres.uid,0) , 0)
  USE IN sqlres
 ENDIF 
 WAIT CLEAR 

 WAIT "днаюбкемхе гюохях б T_TAP_10874..." WINDOW NOWAIT 

 m.ntapuid       = m.maxtapuid + 1
 m.dsuid         = m.maxdsuid + 1
 m.version_end   = '9223372036854775807'
 m.is_top        = 1
 m.created_by    = 99
 m.lpuuid        = 55

 SELECT talon 

 m.ntapsadded = 0
 m.ndsadded   = 0
 m.nusladded  = 0
 m.ndtadded   = 0
 
 m.tapuid = 0
 
 m.nNewRecs = 0
 
 m.SumAdded = 0

 m.nDone=0
 SCAN 
  IF EMPTY(npatuid)
   LOOP 
  ENDIF 
  IF !EMPTY(ntapuid)
   LOOP 
  ENDIF 
  IF !OK
   LOOP 
  ENDIF 

  m.nNewRecs  = m.nNewRecs + 1
  
  m.vbeg10874 = INT(VAL(v_beg10874)) + 1000000
  m.vbeg31571 = INT(VAL(v_beg31571)) + 1000000
  m.vbeg61560 = INT(VAL(v_beg61560)) + 1000000
  m.patuid    = npatuid
  m.adruid    = IIF(SEEK(m.patuid, 'people', 'id'), people.adrid, 0)
  m.policyuid = IIF(SEEK(m.patuid, 'people', 'id'), people.policyuid, 0)
  m.bookuid   = IIF(SEEK(m.patuid, 'people', 'id'), people.bookid, 0)
  m.d_u       = d_u
  m.doc1      = doctoruid
  m.doc2      = nurseuid
  m.mprog     = 'нлялй'
  m.splace    = '1'
  m.ds        = ds
  m.ismain    = 1
  m.coduid    = ALLTRIM(coduid)
  m.k_u       = k_u

  IF tapuid=m.tapuid
   InsCom1 = "insert into t_tapservice_61560 "
   InsCom2 = "(version_begin,version_end,is_top,created_by,diagnosisuid_34765,servicecode_20924,count_23546) "
   InsCom3 = "values "
   InsCom4 = "(?m.vbeg61560,?m.version_end,?m.is_top,?m.created_by,?m.dsuid,?m.coduid,?m.k_u) "
   InsCom5 = ""
   InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
   lrslt = SQLEXEC(lnHndl, InsCom)
   IF lrslt<=0
    =AERROR(errarr)
    = MESSAGEBOX(ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
    EXIT 
   ELSE 
    m.nusladded = m.nusladded + 1
   ENDIF 

   InsCom1 = "insert into t_tapservice_63336 "
   InsCom2 = "(version_begin,version_end,is_top,created_by,diagnosisuid_59030,servicecode_18926,speccase_39152) "
   InsCom3 = "values "
   InsCom4 = "(?m.vbeg61560+2,?m.version_end,?m.is_top,?m.created_by,?m.dsuid,?m.coduid,'0') "
   InsCom5 = ""
   InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
   lrslt = SQLEXEC(lnHndl, InsCom)
   IF lrslt<=0
     =AERROR(errarr)
     ErrMsg = ALLTRIM(errarr(2))
     = MESSAGEBOX(ErrMsg, 16, ALLTRIM(STR(errarr(1))))
     = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
    EXIT 
   ELSE 
    m.ndtadded = m.ndtadded + 1
   ENDIF 

   REPLACE ntapuid WITH m.ntapuid, ndiaguid WITH m.dsuid

   m.SumAdded = m.SumAdded + s_all

   IF m.SumAdded>=m.ChoosedSum
    EXIT 
   ENDIF 

   LOOP 
  ENDIF 
  
  m.ntapuid = m.ntapuid+1
  m.dsuid = m.dsuid + 1

  m.tapuid = tapuid
  InsCom1 = "insert into t_tap_10874 "
  InsCom2 = "(uid,version_begin,version_end,is_top,created_by,lpuuid_25856,patientuid_40511,addressuid_30547,;
   policyuid_53853,bookuid_60769,fillingdate_36966,doctoruid_47963,nurseuid_48406,medicalprogramm_28647,;
   serviceplace_59680) "
  InsCom3 = "values "
  DO CASE 
   CASE m.doc1>0 AND m.doc2=0
    InsCom4 = "(?m.ntapuid,?m.vbeg10874,?m.version_end,?m.is_top,?m.created_by,?m.lpuuid,?m.patuid,;
     ?m.adruid,?m.policyuid,?m.bookuid,?m.d_u,?m.doc1,NULL,?m.mprog,?m.splace) "
   CASE m.doc1=0 AND m.doc2>0
    InsCom4 = "(?m.ntapuid,?m.vbeg10874,?m.version_end,?m.is_top,?m.created_by,?m.lpuuid,?m.patuid,;
     ?m.adruid,?m.policyuid,?m.bookuid,?m.d_u,NULL,?m.doc2,?m.mprog,?m.splace) "
   CASE m.doc1=0 AND m.doc2=0
    InsCom4 = "(?m.ntapuid,?m.vbeg10874,?m.version_end,?m.is_top,?m.created_by,?m.lpuuid,?m.patuid,;
     ?m.adruid,?m.policyuid,?m.bookuid,?m.d_u,NULL,NULL,?m.mprog,?m.splace) "
   OTHERWISE 
    InsCom4 = "(?m.ntapuid,?m.vbeg10874,?m.version_end,?m.is_top,?m.created_by,?m.lpuuid,?m.patuid,;
     ?m.adruid,?m.policyuid,?m.bookuid,?m.d_u,?m.doc1,?m.doc2,?m.mprog,?m.splace) "
  ENDCASE 
  InsCom5 = ""
  InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
  lrslt = SQLEXEC(lnHndl, InsCom)
  IF lrslt<=0
   EXIT 
  ELSE 
   m.ntapsadded = m.ntapsadded + 1
  ENDIF 
  
  InsCom1 = "insert into t_tapdiagnosis_31571 "
  InsCom2 = "(uid,version_begin,version_end,is_top,created_by,tapuid_30432,icdcode_39884,ismain_36277) "
  InsCom3 = "values "
  InsCom4 = "(?m.dsuid,?m.vbeg31571,?m.version_end,?m.is_top,?m.created_by,?m.ntapuid,?m.ds,?m.ismain) "
  InsCom5 = ""
  InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
  lrslt = SQLEXEC(lnHndl, InsCom)
  IF lrslt<=0
   EXIT 
  ELSE 
   m.ndsadded = m.ndsadded + 1
  ENDIF 

  InsCom1 = "insert into t_tapservice_61560 "
  InsCom2 = "(version_begin,version_end,is_top,created_by,diagnosisuid_34765,servicecode_20924,count_23546) "
  InsCom3 = "values "
  InsCom4 = "(?m.vbeg61560,?m.version_end,?m.is_top,?m.created_by,?m.dsuid,?m.coduid,?m.k_u) "
  InsCom5 = ""
  InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
  lrslt = SQLEXEC(lnHndl, InsCom)
  IF lrslt<=0
    =AERROR(errarr)
    ErrMsg = ALLTRIM(errarr(2))
    = MESSAGEBOX(ErrMsg, 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
   EXIT 
  ELSE 
   m.nusladded = m.nusladded + 1
  ENDIF 

  InsCom1 = "insert into t_tapservice_63336 "
  InsCom2 = "(version_begin,version_end,is_top,created_by,diagnosisuid_59030,servicecode_18926,speccase_39152) "
  InsCom3 = "values "
  InsCom4 = "(?m.vbeg61560+2,?m.version_end,?m.is_top,?m.created_by,?m.dsuid,?m.coduid,'0') "
  InsCom5 = ""
  InsCom = InsCom1 + InsCom2 + InsCom3 + InsCom4 + InsCom5
  lrslt = SQLEXEC(lnHndl, InsCom)
  IF lrslt<=0
    =AERROR(errarr)
    ErrMsg = ALLTRIM(errarr(2))
    = MESSAGEBOX(ErrMsg, 16, ALLTRIM(STR(errarr(1))))
    = MESSAGEBOX(ALLTRIM(errarr(3)), 16, ALLTRIM(STR(errarr(1))))
   EXIT 
  ELSE 
   m.ndtadded = m.ndtadded + 1
  ENDIF 

  REPLACE ntapuid WITH m.ntapuid, ndiaguid WITH m.dsuid
  
  m.SumAdded = m.SumAdded + s_all
  
  IF m.SumAdded>=m.ChoosedSum
   EXIT 
  ENDIF 
 ENDSCAN 
 WAIT CLEAR 
 
 IF m.nNewRecs==0
  MESSAGEBOX('бяе гюохях днаюбкемш пюмее!',0+64,'')
 ELSE 
  MESSAGEBOX('днаюбкемн '+TRANSFORM(m.SumAdded,'99 999 999.99')+' псакеи!',0+64,'')
  
*  MESSAGEBOX('днаюбкемн '+ALLTRIM(STR(m.ntapsadded))+' рюкнмнб!',0+64,'')
*  MESSAGEBOX('днаюбкемн '+ALLTRIM(STR(m.ndsadded))+' дхюцмнгнб!',0+64,'')
*  MESSAGEBOX('днаюбкемн '+ALLTRIM(STR(m.nusladded))+' сяксц!',0+64,'')
*  MESSAGEBOX('днаюбкемн '+ALLTRIM(STR(m.ndtadded))+' сяксц!',0+64,'')
 ENDIF 

 WAIT CLEAR 


 WAIT "нрйкчвемхе нр аюгш дюммшу..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && яНЕДХМЕМХЕ СЯОЕЬМН ГЮЙПШРН
*   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && нЬХАЙЮ СПНБМЪ ЯНЕДХМЕМХЪ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && нЬХАЙЮ СПНБМЪ НЙПСФЕМХЪ
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && рЮЙНЦН АШРЭ МЕ ДНКФМН, МН БЯЕ-РЮЙХ...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 
 WAIT CLEAR 

RELEASE lnhndl, lresult, ln_disconn

USE IN talon 
USE IN people

*MESSAGEBOX('днаюбкеммюъ ясллю: '+TRANSFORM(m.SumOfAcc, '9 999 999.99'),0+64,'')

 
RETURN 