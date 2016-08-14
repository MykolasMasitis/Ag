PROCEDURE SQLExp
 SET NULL OFF 
 
 WAIT "����������� � ������� ���� ������...." WINDOW NOWAIT 
 
 lnHndl = SQLCONNECT("mssql", "sa", "")
 IF lnHndl <= 0
  = MESSAGEBOX('Cannot make connection', 16, 'SQL Connect Error')
  RETURN 
 ELSE
  = MESSAGEBOX('Connection made', 48, 'SQL Connect Message')
 ENDIF
 
 lrslt = SQLEXEC(lnHndl, "USE PDPStrFacility")
 IF lrslt<=0
  = MESSAGEBOX('Cannot select database!', 16, 'PDPStrFacility')
 ELSE 
  = MESSAGEBOX('Database selected!', 48, 'PDPStrFacility')
 ENDIF 

 WAIT "VERSION_TICK..." WINDOW NOWAIT 

 SelString = "select CAST(version_begin as numeric(20)) as v_beg, ;
  CAST(version_end as numeric(20)) as v_end, tick_begin, tick_end from version_tick"

 lrslt = SQLEXEC(lnHndl, selstring)
 IF lrslt<=0
  = MESSAGEBOX('Cannot process select!', 16, 'version_tick')
 ELSE 
  = MESSAGEBOX('Records selected!', 48, 'version_tick')
  COPY TO &tc_comdir\version_tick
  USE 
 ENDIF 

 WAIT CLEAR 

 WAIT "���������� �� ���� ������..." WINDOW NOWAIT 

 ln_disconn = SQLDISCONNECT(lnHndl)
 DO CASE 
  CASE ln_disconn ==  1 && ���������� ������� �������
   = MESSAGEBOX('Connection closed', 48, 'SQL Connect Message')
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