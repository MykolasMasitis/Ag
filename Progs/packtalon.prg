PROCEDURE PackTalon
 IF MESSAGEBOX(CHR(13)+CHR(10)+'������ ��������� �� �������?'+CHR(13)+CHR(10),4+16,'')==7
  RETURN 
 ENDIF 

 WAIT "����������� � ������� ���� ������...." WINDOW NOWAIT 
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

 WAIT "���������� �������� �������� T_TAPSERVICE_61560..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_61560 disable trigger trt_tapservice_61560_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� ��������!",0+64,"")
 WAIT CLEAR 

 WAIT "�������� ���� T_TAPSERVICE_61560..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_TAPSERVICE_61560 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ��������� T_TAPSERVICE_61560'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "��������� �������� �������� T_TAPSERVICE_61560..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_61560 enable trigger trt_tapservice_61560_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� �������!",0+64,"")
 WAIT CLEAR 

 WAIT "���������� �������� �������� T_TAPSERVICE_63336..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_63336 disable trigger trt_tapservice_63336_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� ��������!",0+64,"")
 WAIT CLEAR 

 WAIT "�������� ���� T_TAPSERVICE_63336..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_TAPSERVICE_63336 WHERE version_end!=9223372036854775807 or ;
  speccase_39152 is null or speccase_39152='0'"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ��������� T_TAPSERVICE_63336'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "��������� �������� �������� T_TAPSERVICE_63336..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapservice_63336 enable trigger trt_tapservice_63336_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� �������!",0+64,"")
 WAIT CLEAR 

 WAIT "���������� �������� �������� T_TAP_10874..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 disable trigger trt_tap_10874_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� ��������!",0+64,"")
 WAIT CLEAR 

 WAIT "�������� ���� T_TAP_10874..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_TAP_10874 WHERE version_end!=9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ��������� T_TAP_10874'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "��������� �������� �������� T_TAP_10874..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_10874 enable trigger trt_tap_10874_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� �������!",0+64,"")
 WAIT CLEAR 

 WAIT "���������� �������� �������� T_TAP_54502..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_54502 disable trigger trt_tap_54502_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� ��������!",0+64,"")
 WAIT CLEAR 

 WAIT "�������� ���� T_TAP_54502..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_TAP_54502 WHERE version_end!=9223372036854775807 or speccase_39994 is null or;
  speccase_39994='0'"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ��������� T_TAP_54502'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "��������� �������� �������� T_TAP_54502..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tap_54502 enable trigger trt_tap_54502_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� �������!",0+64,"")
 WAIT CLEAR 

 WAIT "���������� �������� �������� T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapdiagnosis_31571 disable trigger trt_tapdiagnosis_31571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� ��������!",0+64,"")
 WAIT CLEAR 

 WAIT "�������� ���� T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 DelCom = "DELETE FROM T_TAPDIAGNOSIS_31571 WHERE version_end<9223372036854775807"
 lrslt = SQLEXEC(lnHndl, DelCom)
 IF lrslt<=0
  =AERROR(errarr)
  = MESSAGEBOX(CHR(13)+CHR(10)+'�� ������� ��������� T_TAPDIAGNOSIS_31571'+CHR(13)+CHR(10)+;
   ALLTRIM(errarr(2)), 16, ALLTRIM(STR(errarr(1))))
 ENDIF 
 WAIT CLEAR 

 WAIT "��������� �������� �������� T_TAPDIAGNOSIS_31571..." WINDOW NOWAIT 
 TriggerOff = "alter table t_tapdiagnosis_31571 enable trigger trt_tapdiagnosis_31571_delete"
 lrslt = SQLEXEC(lnHndl, TriggerOff)
 IF lrslt<=0
  =AERROR(errarr)
  MESSAGEBOX(ALLTRIM(errarr(2)), 16, STR(m.dsuid))
 ENDIF 
*  MESSAGEBOX("������� �������!",0+64,"")
 WAIT CLEAR 

 WAIT "���������� �� ���� ������..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && ���������� ������� �������
*   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && ������ ������ ����������
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && ������ ������ ���������
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && ������ ���� �� ������, �� ���-����...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 
 WAIT CLEAR 

RETURN 