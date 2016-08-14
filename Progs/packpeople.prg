PROCEDURE PackPeople
 IF MESSAGEBOX(CHR(13)+CHR(10)+'унрхре союйнбюрэ ад оюжхемрнб?'+CHR(13)+CHR(10),4+16,'')==7
  RETURN 
 ENDIF 

 WAIT "ондйкчвемхе й яепбепс аюгш дюммшу...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  =AERROR(errarr)
  =MESSAGEBOX(CHR(13)+CHR(10)+'Cannot make connection'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
  RETURN 
 ELSE
*  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 WAIT CLEAR 

 lrslt = SQLEXEC(lnHndl, "USE PDPStdStorage")
 IF lrslt<=0
  =AERROR(errarr)
  =MESSAGEBOX(CHR(13)+CHR(10)+'Cannot select database!'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
  RETURN 
 ELSE 
*  = MESSAGEBOX('Database selected!', 48, 'PDPStdStorage')
 ENDIF 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_PATIENT_10905..." WINDOW NOWAIT 
 TriggerOff = "alter table t_patient_10905 disable trigger trt_patient_10905_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_PATIENT_10905..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_PATIENT_10905 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_PATIENT_10905'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_PATIENT_10905..." WINDOW NOWAIT 
 TriggerOff = "alter table t_patient_10905 enable trigger trt_patient_10905_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_PATIENT_45426..." WINDOW NOWAIT 
 TriggerOff = "alter table t_patient_45426 disable trigger trt_patient_45426_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_PATIENT_45426..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_PATIENT_45426 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_PATIENT_45426'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_PATIENT_45426..." WINDOW NOWAIT 
 TriggerOff = "alter table t_patient_45426 enable trigger trt_patient_45426_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_ADDRESS_47256..." WINDOW NOWAIT 
 TriggerOff = "alter table t_address_47256 disable trigger trt_address_47256_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_ADDRESS_47256..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_ADDRESS_47256 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_ADDRESS_47256'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_ADDRESS_47256..." WINDOW NOWAIT 
 TriggerOff = "alter table t_address_47256 enable trigger trt_address_47256_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_POLICY_43176..." WINDOW NOWAIT 
 TriggerOff = "alter table t_policy_43176 disable trigger trt_policy_43176_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_POLICY_43176..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_POLICY_43176 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_POLICY_43176'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_POLICY_43176..." WINDOW NOWAIT 
 TriggerOff = "alter table t_policy_43176 enable trigger trt_policy_43176_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_POLICY_53495..." WINDOW NOWAIT 
 TriggerOff = "alter table t_policy_53495 disable trigger trt_policy_53495_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_POLICY_53495..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_POLICY_53495 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_POLICY_53495'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_POLICY_53495..." WINDOW NOWAIT 
 TriggerOff = "alter table t_policy_53495 enable trigger trt_policy_53495_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_ADDRESSTYPE_39597..." WINDOW NOWAIT 
 TriggerOff = "alter table t_addresstype_39597 disable trigger trt_addresstype_39597_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_ADDRESSTYPE_39597..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_ADDRESSTYPE_39597 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_ADDRESSTYPE_39597'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_ADDRESSTYPE_39597..." WINDOW NOWAIT 
 TriggerOff = "alter table t_addresstype_39597 enable trigger trt_addresstype_39597_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_ADMOKR_44752..." WINDOW NOWAIT 
 TriggerOff = "alter table t_admokr_44752 disable trigger trt_admokr_44752_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_ADMOKR_44752..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_ADMOKR_44752 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_ADMOKR_44752'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_ADMOKR_44752..." WINDOW NOWAIT 
 TriggerOff = "alter table t_admokr_44752 enable trigger trt_admokr_44752_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_BOOK_65067..." WINDOW NOWAIT 
 TriggerOff = "alter table t_book_65067 disable trigger trt_book_65067_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_BOOK_65067..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_BOOK_65067 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_BOOK_65067'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_BOOK_65067..." WINDOW NOWAIT 
 TriggerOff = "alter table t_book_65067 enable trigger trt_book_65067_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_CASEINTERRUPT_12098..." WINDOW NOWAIT 
 TriggerOff = "alter table t_caseinterrupt_12098 disable trigger trt_caseinterrupt_12098_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_CASEINTERRUPT_12098..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_CASEINTERRUPT_12098 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_CASEINTERRUPT_12098'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_CASEINTERRUPT_12098..." WINDOW NOWAIT 
 TriggerOff = "alter table t_caseinterrupt_12098 enable trigger trt_caseinterrupt_12098_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_CASEINTERRUPT_43563..." WINDOW NOWAIT 
 TriggerOff = "alter table t_caseinterrupt_43563 disable trigger trt_caseinterrupt_43563_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_CASEINTERRUPT_43563..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_CASEINTERRUPT_43563 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_CASEINTERRUPT_43563'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_CASEINTERRUPT_43563..." WINDOW NOWAIT 
 TriggerOff = "alter table t_caseinterrupt_43563 enable trigger trt_caseinterrupt_43563_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_DETACHEDSERVICE_58302..." WINDOW NOWAIT 
 TriggerOff = "alter table t_detachedservice_58302 disable trigger trt_detachedservice_58302_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_DETACHEDSERVICE_58302..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_DETACHEDSERVICE_58302 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_DETACHEDSERVICE_58302'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_DETACHEDSERVICE_58302..." WINDOW NOWAIT 
 TriggerOff = "alter table t_detachedservice_58302 enable trigger trt_detachedservice_58302_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_DUL_44571..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dul_44571 disable trigger trt_dul_44571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_DUL_44571..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_DUL_44571 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_DUL_44571'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_DUL_44571..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dul_44571 enable trigger trt_dul_44571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_DULTYPE_18894..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultype_18894 disable trigger trt_dultype_18894_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_DULTYPE_18894..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_DULTYPE_18894 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_DULTYPE_18894'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_DULTYPE_18894..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultype_18894 enable trigger trt_dultype_18894_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_DULTYPE_47004..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultype_47004 disable trigger trt_dultype_47004_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_DULTYPE_47004..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_DULTYPE_47004 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_DULTYPE_47004'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_DULTYPE_47004..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultype_47004 enable trigger trt_dultype_47004_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_DULTYPELINK_03773..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultypelink_03773 disable trigger trt_dultypelink_03773_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_DULTYPELINK_03773..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_DULTYPELINK_03773 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_DULTYPELINK_03773'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_DULTYPELINK_03773..." WINDOW NOWAIT 
 TriggerOff = "alter table t_dultypelink_03773 enable trigger trt_dultypelink_03773_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_EMPLOYEE_22089..." WINDOW NOWAIT 
 TriggerOff = "alter table t_employee_22089 disable trigger trt_employee_22089_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_EMPLOYEE_22089..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_EMPLOYEE_22089 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_EMPLOYEE_22089'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_EMPLOYEE_22089..." WINDOW NOWAIT 
 TriggerOff = "alter table t_employee_22089 enable trigger trt_employee_22089_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "нрйкчвемхе рпхццепю сдюкемхъ T_EMPLOYEE_49492..." WINDOW NOWAIT 
 TriggerOff = "alter table t_employee_49492 disable trigger trt_employee_49492_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
 WAIT CLEAR 

 WAIT "союйнбйю аюгш T_EMPLOYEE_49492..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_EMPLOYEE_49492 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'ме сдюеряъ союйнбюрэ T_EMPLOYEE_49492'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "бйкчвемхе рпхццепю сдюкемхъ T_EMPLOYEE_49492..." WINDOW NOWAIT 
 TriggerOff = "alter table t_employee_49492 enable trigger trt_employee_49492_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, '')
 ENDIF 
*  MESSAGEBOX("рпхццеп нрйкчвем!",0+64,"")
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

RETURN 