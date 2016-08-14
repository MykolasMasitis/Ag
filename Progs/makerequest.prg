PROCEDURE MakeRequest
 SET NULL OFF  
 IF !fso.FileExists(tc_comdir+'\People.dbf')
  MESSAGEBOX(CHR(13)+CHR(10)+'Œ“—”“—“¬”≈“ ‘¿…À PEOPLE.DBF!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN 
 ENDIF 

 USE tc_comdir+'\people' IN 0 ALIAS people SHARED 
 IF RECCOUNT('people')<=0
  USE IN people 
  MESSAGEBOX(CHR(13)+CHR(10)+'¬ ‘¿…À≈ PEOPLE Õ≈“ Õ» ŒƒÕŒ… «¿œ»—»!'+CHR(13)+CHR(10),0+16,'people.dbf')
  RETURN
 ENDIF 
 
 DO FORM setperiod

 fzz = 'q_' + PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)+'.dbf'
 
 IF fso.FileExists(tc_comdir+'\'+fzz)
  IF MESSAGEBOX(CHR(13)+CHR(10)+'«¿œ–Œ— «¿ ¬€¡–¿ÕÕ€… œ≈–»Œƒ ”∆≈ ‘Œ–Ã»–Œ¬¿À—ﬂ!'+CHR(13)+CHR(10)+;
   '¬€ ’Œ“»“≈ —‘Œ–Ã»–¬Œ¿“‹ ≈√Œ «¿ÕŒ¬Œ?',4+32,FZZ)=7
    RETURN 
  ELSE 
   iii = 1
   fzzbak = 'q_' + PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)+'.'+PADL(iii,3,'0')
   DO WHILE fso.FileExists(tc_comdir+'\'+fzzbak)
    iii = iii + 1
    fzzbak = 'q_' + PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)+'.'+PADL(iii,3,'0')
   ENDDO 
   fso.CopyFile(tc_comdir+'\'+fzz, tc_comdir+'\'+fzzbak)
   fso.DeleteFile(tc_comdir+'\'+fzz)
  ENDIF 
 ENDIF 

 IF OpenFile(tc_comdir+'\logfile', 'logfile', 'shar', 'id')>0
  USE IN people
  RETURN 
 ENDIF 
 
 m.id = ''

 fso.CopyFile(tc_tmpldir+'\Zapros.dbf', tc_comdir+'\'+fzz, .t.)
 oApp.CodePage(tc_comdir+'\'+fzz, 866, .t.)
 
 WAIT "‘Œ–Ã»–Œ¬¿Õ»≈ «¿œ–Œ—¿..." WINDOW NOWAIT 

 USE tc_comdir+'\'+fzz IN 0 ALIAS zapros SHARED 

 SELECT people
 m.totrecs = 0
 SCAN FOR c_t==77
  SCATTER MEMVAR 
  m.rec = RECNO()

  m.date_in  = m.tdat1
  m.date_out = m.tdat2

  m.recid    = PADL(m.rec,6,'0')
  m.fam      = PROPER(ALLTRIM(m.fam))
  m.im       = PROPER(ALLTRIM(m.im))
  m.ot       = PROPER(ALLTRIM(m.ot))

  m.q        = ''

  IF OCCURS(' ', ALLTRIM(m.sn_pol))>0
   m.s_pol    = SUBSTR(m.sn_pol, 1, AT(' ',m.sn_pol)-1)
   m.n_pol    = SUBSTR(m.sn_pol, AT(' ',m.sn_pol)+1)
  ELSE 
   m.s_pol    = ''
   m.n_pol    = m.sn_pol
  ENDIF 
  DO CASE 
   CASE ISALPHA(m.sn_pol)
    m.tip_d='¬'
   CASE OCCURS(' ', ALLTRIM(m.sn_pol))==0
    m.tip_d='œ'
   OTHERWISE 
    m.tip_d='—'
  ENDCASE 

  m.dr       = DToS(m.dr)

  INSERT INTO Zapros FROM MEMVAR 
  
  m.totrecs = m.totrecs + 1
* IF m.totrecs>103
* EXIT 
* ENDIF 
    
 ENDSCAN 
 USE IN Zapros
 USE IN people 

 =MakeOneRequest()
 INSERT INTO logfile (id, period, sent, tip, totrecs) VALUES ;
  (m.id, STR(m.tyear,4)+PADL(m.tmonth,2,'0'), DATETIME(), 'ERZ_sverka4', m.totrecs)

 USE IN logfile
RETURN 

FUNCTION MakeOneRequest
 m.id  = SYS(3)
 ID     = ALLTRIM(m.id+'.'+m.gcUser+'@gp189.msk.oms')
  
 TFile  = 'terz_' + m.mcod
 BFile  = 'berz_' + m.mcod
 DFile  = 'derz_' + m.mcod

 iii = 1
 DO WHILE fso.FileExists(tc_aisdir+'\'+m.gcUser+'\OUTPUT\'+m.bfile)
  m.tfile  = 'terz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.bfile  = 'berz_' + m.mcod + '_' + PADL(iii,2,'0')
  m.dfile  = 'derz_' + m.mcod + '_' + PADL(iii,2,'0')
  iii = iii + 1
 ENDDO 
   
 fso.CopyFile(tc_comdir+'\'+fzz, tc_aisdir+'\'+m.gcUser+'\OutPut\'+DFile)
* fso.CopyFile(tc_comdir+'\Zapros.dbf', tc_comdir+'\'+DFile)
* fso.DeleteFile(tc_comdir+'\Zapros.dbf')
   
 poi = FCREATE('&tc_aisdir\&gcUser\OutPut\&TFile')
 IF poi != -1
  =FPUTS(poi,'To: erz@mgf.msk.oms')
  =FPUTS(poi,'Message-Id: &ID')
  =FPUTS(poi,'Subject: ERZ_sverka4')
*  fzz = 'q_' + PADL(tmonth,2,'0')+RIGHT(STR(tyear,4),1)+'.dbf'
  =FPUTS(poi,'Attachment: '+DFile+' '+Fzz)
 ENDIF 
 =FCLOSE(poi)
 
* fso.CopyFile('&tc_aisdir\&gcUser\OutPut\&TFile', tc_comdir+'\'+BFile)
 oTFile = fso.GetFile('&tc_aisdir\&gcUser\OutPut\&TFile')
 oTFile.Move('&tc_aisdir\&gcUser\OutPut\&BFile')

 WAIT CLEAR 

 MESSAGEBOX('«¿œ–Œ— Œ“œ–¿¬À≈Õ!', 0+64, '')
RETURN 