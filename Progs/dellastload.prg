PROCEDURE DelLastLoad
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш унрхре сдюкхрэ онякедмчч гюцпсгйс'+;
  CHR(13)+CHR(10)+'хг ад MS SQL?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'бш юаянкчрмн сбепемш б ябнху деиярбхъу?'+;
  CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 IF !fso.FileExists(tc_basedir+'\Talon.dbf')
  MESSAGEBOX('нрясрярбсер тюик TALON.DBF!',0+16,'')
  RETURN 
 ENDIF 
 
 IF OpenFile(tc_basedir+'\Talon', 'talon', 'shar')>0
  IF USED('talon')
   USE IN talon 
  ENDIF 
  RETURN 
 ENDIF 
 
 SELECT talon 
 m.LoadedRecs = 0
 COUNT  FOR !EMPTY(ntapuid) TO m.LoadedRecs
 IF m.LoadedRecs = 0
  MESSAGEBOX('мевецн сдюкърэ!',0+16,'')
  USE IN talon 
  RETURN 
 ENDIF 
 IF MESSAGEBOX(CHR(13)+CHR(10)+'ашкн гюцпсфемн '+TRANSFORM(m.LoadedRecs,'9999999')+' гюохяеи!'+;
  CHR(13)+CHR(10)+'опнднкфхрэ?'+CHR(13)+CHR(10),4+32,'')==7
  USE IN talon 
  RETURN 
 ENDIF 

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
 
 WAIT "нрйкчвемхе рпхццепю сдюкемхъ TRT_TAPSERVICE_63336_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_63336 disable trigger trt_tapservice_63336_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "сдюкемхе гюохяеи хг рюакхжш T_TAPSERVICE_63336..." WINDOW NOWAIT 
 m.deletedin63336 = 0
 m.skippedin63336 = 0
 SCAN 
  IF EMPTY(ntapuid)
   LOOP 
  ENDIF 
  m.tapuid = ntapuid
  m.dsuid = ndiaguid

  DelCom = "delete from t_tapservice_63336 where diagnosisuid_59030=?m.dsuid"
  lrslt = SQLEXEC(lnHndl, DelCom)
  IF lrslt<=0
   m.skippedin63336 = m.skippedin63336 + 1
   =AERROR(errarr)
*   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
*   EXIT 
  ELSE 
   m.deletedin63336 = m.deletedin63336 + 1
  ENDIF 
 ENDSCAN 

* MESSAGEBOX('сдюкемн хг T_TAPSERVICE_63336 '+ALLTRIM(STR(m.deletedin63336))+' гюохяеи!',0+64,'')
* MESSAGEBOX('опносыемн б T_TAPSERVICE_63336 '+ALLTRIM(STR(m.skippedin63336))+' гюохяеи!',0+64,'')

 WAIT "бйкчвемхе рпхццепю сдюкемхъ TRT_TAPSERVICE_63336_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_63336 enable trigger trt_tapservice_63336_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп бйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ TRT_TAPSERVICE_61560_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_61560 disable trigger trt_tapservice_61560_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "сдюкемхе гюохяеи хг рюакхжш T_TAPSERVICE_61560..." WINDOW NOWAIT 
 m.deletedin61560 = 0
 m.skippedin61560 = 0
 SCAN 
  IF EMPTY(ntapuid)
   LOOP 
  ENDIF 
  m.tapuid = ntapuid
  m.dsuid = ndiaguid

  DelCom = "delete from t_tapservice_61560 where diagnosisuid_34765=?m.dsuid"
  lrslt = SQLEXEC(lnHndl, DelCom)
  IF lrslt<=0
   m.skippedin61560 = m.skippedin61560 + 1
   =AERROR(errarr)
*   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
*   EXIT 
  ELSE 
   m.deletedin61560 = m.deletedin61560 + 1
  ENDIF 
 ENDSCAN 

* MESSAGEBOX('сдюкемн хг T_TAPSERVICE_61560 '+ALLTRIM(STR(m.deletedin61560))+' гюохяеи!',0+64,'')
* MESSAGEBOX('опносыемн б T_TAPSERVICE_61560 '+ALLTRIM(STR(m.skippedin61560))+' гюохяеи!',0+64,'')

 WAIT "бйкчвемхе рпхццепю сдюкемхъ TRT_TAPSERVICE_61560_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_61560 enable trigger trt_tapservice_61560_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп бйкчвем!",0+64,"")
 WAIT CLEAR 
 
 WAIT "нрйкчвемхе рпхццепю сдюкемхъ TRT_TAPDIAGNOSIS_31571_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table T_TAPDIAGNOSIS_31571 disable trigger trt_tapdiagnosis_31571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.tapuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "сдюкемхе гюохяеи хг рюакхжш T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 m.deletedin31571 = 0
 m.skippedin31571 = 0
 SCAN 
  IF EMPTY(ntapuid)
   LOOP 
  ENDIF 
  m.tapuid = ntapuid
  m.dsuid = ndiaguid

  DelCom = "delete from t_tapdiagnosis_31571 where tapuid_30432=?m.tapuid"
  lrslt = SQLEXEC(lnHndl, DelCom)
  IF lrslt<=0
   m.skippedin31571 = m.skippedin31571 + 1
   =AERROR(errarr)
*   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
*   EXIT 
  ELSE 
   m.deletedin31571 = m.deletedin31571 + 1
  ENDIF 
 ENDSCAN 

* MESSAGEBOX('сдюкемн хг T_TAPDIAGNOSIS_31571 '+ALLTRIM(STR(m.deletedin61560))+' гюохяеи!',0+64,'')
* MESSAGEBOX('опносыемн б T_TAPDIAGNOSIS_31571 '+ALLTRIM(STR(m.skippedin61560))+' гюохяеи!',0+64,'')

 WAIT "бйкчвемхе рпхццепю сдюкемхъ TRT_TAPDIAGNOSIS_31571_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapdiagnosis_31571 enable trigger trt_tapdiagnosis_31571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.tapuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп бйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ TRT_TAP_10874_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 disable trigger trt_tap_10874_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.tapuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "сдюкемхе гюохяеи хг рюакхжш T_TAP_10874..." WINDOW NOWAIT 
 m.deletedin10874 = 0
 m.skippedin10874 = 0
 SCAN 
  IF EMPTY(ntapuid)
   LOOP 
  ENDIF 
  m.tapuid = ntapuid
  m.dsuid = ndiaguid

  DelCom = "delete from t_tap_10874 where uid=?m.tapuid"
  lrslt = SQLEXEC(lnHndl, DelCom)
  IF lrslt<=0
   m.skippedin10874 = m.skippedin10874 + 1
   =AERROR(errarr)
*   = MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
*   EXIT 
  ELSE 
   m.deletedin10874 = m.deletedin10874 + 1
  ENDIF 
 ENDSCAN 

* MESSAGEBOX('сдюкемн хг T_TAP_10874 '+ALLTRIM(STR(m.deletedin10874))+' гюохяеи!',0+64,'')
* MESSAGEBOX('опносыемн б T_TAP_10874 '+ALLTRIM(STR(m.skippedin10874))+' гюохяеи!',0+64,'')

 WAIT "бйкчвемхе рпхццепю сдюкемхъ TRT_TAP_10874_DELETE..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 enable trigger trt_tap_10874_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.tapuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп бйкчвем!",0+64,"")
 WAIT CLEAR 

 USE IN talon 

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
 
 MESSAGEBOX('цнрнбн!',0+64,'')

 
RETURN 