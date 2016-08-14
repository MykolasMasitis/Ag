PROCEDURE SaveUIDInTapDiags
 IF MESSAGEBOX(CHR(13)+CHR(10)+'�� ������ ��������� ������'+;
  CHR(13)+CHR(10)+'������� T_TAPDIAGNOSIS_31571_UID?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 

 WAIT "����������� � ������� ���� ������...." WINDOW NOWAIT 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
 ENDIF
 WAIT CLEAR 
 
 WAIT "����� �� PDPSTDSTORAGE..." WINDOW NOWAIT 
 lrslt = SQLEXEC(lnHndl, "USE PDPSTDSTORAGE")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPSTDSTORAGE')
 ELSE 
 ENDIF 
 WAIT CLEAR 

 WAIT "..." WINDOW NOWAIT 
 SelString = "SELECT * FROM t_tapdiagnosis_31571_uid"

 lrslt = SQLEXEC(lnHndl, SelString)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, '')
 ELSE 
  COPY TO &tc_comdir\tapdiagsuidet
  USE 
 ENDIF 
 WAIT CLEAR 

 WAIT "..." WINDOW NOWAIT 
 SelString = "SELECT spid, CAST(entity as char(50)) as entity, uid FROM last_uid"

 lrslt = SQLEXEC(lnHndl, SelString)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, '')
 ELSE 
  CREATE TABLE &tc_comdir\lastuidet (spid i, entity c(50), uid i)
  USE 
  =OpenFile(tc_comdir+'\lastuidet', 'lastuid', 'shar')
  SELECT sqlresult
  BROWSE 
  SCAN 
   SCATTER MEMVAR 
   m.entity = ALLTRIM(m.entity)
   INSERT INTO lastuid FROM MEMVAR 
  ENDSCAN 
  USE 
  USE IN lastuid
 ENDIF 
 WAIT CLEAR 

 WAIT "���������� �� ���� ������..." WINDOW NOWAIT 
 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && ���������� ������� �������
 *  = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
  CASE ln_disconn == -1 && ������ ������ ����������
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  CASE ln_disconn == -2 && ������ ������ ���������
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
  OTHERWISE && ������ ���� �� ������, �� ���-����...
   = MESSAGEBOX('Cannot close connection', 16, 'SQL Connect Error')
 ENDCASE 

 RELEASE lnhndl, lresult, ln_disconn
 
 WAIT CLEAR 
RETURN 