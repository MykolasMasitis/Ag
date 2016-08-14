PROCEDURE LookForAnswer
  IF MESSAGEBOX(CHR(13)+CHR(10)+'ПОИСКАТЬ ОТВЕТЫ?'+CHR(13)+CHR(10),4+32,'')==7
  RETURN 
 ENDIF 
 
 IF !fso.FileExists(tc_comdir+'\logfile.dbf')
  RETURN 
 ENDIF 
 
 IF OpenFile(tc_comdir+'\logfile', 'logfile', 'shar', 'id')>0
  RETURN 
 ENDIF 
 
 oMailDir = fso.GetFolder(tc_aisdir+'\'+m.gcuser+'\input')
 MailDirName = oMailDir.Path
 oFilesInMailDir = oMailDir.Files
 nFilesInMailDir = oFilesInMailDir.Count

 MESSAGEBOX('ОБНАРУЖЕНО '+ALLTRIM(STR(nFilesInMailDir))+' ФАЙЛОВ!', 0+64, gcUser)

 IF nFilesInMailDir<=0
  USE IN logfile
  RELEASE omaildir,omaildirname,ofilesinmaildir,nfilesinmaildir
  RETURN 
 ENDIF 

 FOR EACH oFileInMailDir IN oFilesInMailDir
  IF LOWER(oFileInMailDir.Name) = 'b'
   m.cfrom      = ''
   m.cdate      = ''
   m.cmessage   = ''
   m.resmesid   = ''
   m.csubject   = ''
   m.csubject1  = ''
   m.csubject2  = ''
   m.attachment = ''
   m.bodypart   = ''

   m.attaches   = 0 && Сколько присоединенных файлов в одной ИП
   DIMENSION dattaches(10,2)
   dattaches = ''

   m.bparts   = 0 && Сколько присоединенных файлов в одной ИП
   DIMENSION dbparts(10,2)
   dbparts = ''

   m.BFullName = oFileInMailDir.Path
   m.bname     = oFileInMailDir.Name
   m.recieved  = oFileInMailDir.DateLastModified

   CFG = FOPEN(m.BFullName)
   =ReadCFGFile()
   = FCLOSE (CFG)
   
*   MESSAGEBOX(m.resmesid,0+64,oFileInMailDir.Name)

   IF LEFT(UPPER(ALLTRIM(m.csubject)),3) != 'ERZ'
    LOOP 
   ENDIF 
   
*   m.sent = dt2date(m.cdate)

   IF !SEEK(LEFT(m.resmesid,8), 'logfile')
    LOOP 
   ELSE 
*   MESSAGEBOX(m.sent,0+64,oFileInMailDir.Name)
    UPDATE logfile SET recieved=m.recieved WHERE id=LEFT(m.resmesid,8)
*    REPLACE logfile.recieved WITH m.sent IN logfile
   ENDIF 

   fso.CopyFile(m.BFullName, tc_comdir + '\' + m.bname)
   fso.DeleteFile(m.BFullName)
   IF m.attaches>0
    FOR nattach = 1 TO m.attaches
     m.dname   = ALLTRIM(dattaches(nattach, 1))
     m.aattname = ALLTRIM(dattaches(nattach, 2))
     IF !EMPTY(m.dname)
      IF !EMPTY(tc_comdir+'\'+m.aattname)
       fso.CopyFile(MailDirName + '\' + m.dname, tc_comdir+'\'+m.aattname)
       fso.DeleteFile(MailDirName + '\' + m.dname)
      ELSE 
       MESSAGEBOX(CHR(13)+CHR(10)+'ФАЙЛ '+UPPER(m.aattname)+' В ДИРЕКТОРИИ COMMON'+CHR(13)+CHR(10)+;
        'СУЩЕСТВУЕТ!',0+16,'')
      ENDIF 
     ENDIF 
    ENDFOR 
   ENDIF 

   IF m.bparts > 0
    FOR npart = 1 TO m.bparts
     m.bpname   = ALLTRIM(dbparts(npart, 1))
     IF !EMPTY(m.bpname)
      fso.CopyFile(MailDirName + '\' + m.bpname, tc_comdir+'\'+m.bpname, .f.)
      fso.DeleteFile(MailDirName + '\' + m.bpname)
     ENDIF 
    ENDFOR 
   ENDIF 
   
  ENDIF 
 NEXT 
 
 RELEASE omaildir,omaildirname,ofilesinmaildir,nfilesinmaildir
 USE IN logfile
 MESSAGEBOX('ОБРАБОТКА ЗАКОНЧЕНА!',0+64,'')

RETURN 

FUNCTION ReadCFGFile
 DO WHILE NOT FEOF(CFG)
  READCFG = FGETS (CFG)
  DO CASE
   CASE UPPER(READCFG) = 'FROM'
    m.cfrom = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'DATE'
    m.cdate = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'MESSAGE'
    m.cmessage = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'RESENT-MESSAGE-ID'
    m.resmesid = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
   CASE UPPER(READCFG) = 'SUBJECT'
    m.csubject = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    m.csubject1 = LEFT(m.csubject, RAT('#',m.csubject,2))   && Делим subject для последующей вставки кода результата
    m.csubject2 = SUBSTR(m.csubject, RAT('#',m.csubject,1)) && Делим subject для последующей вставки кода результата
   CASE UPPER(READCFG) = 'ATTACHMENT'
    m.attaches   = m.attaches + 1
    m.attachment = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dattaches(m.attaches,1) = ALLTRIM(SUBSTR(m.attachment, 1, AT(" ",m.attachment)-1)) && Название d-файла
    dattaches(m.attaches,2) = ALLTRIM(SUBSTR(m.attachment, AT(" ",m.attachment)+1))    && Фактическое название файла
   CASE UPPER(READCFG) = 'BODYPART'
    m.bparts   = m.bparts + 1
    m.bodypart = ALLTRIM(SUBSTR(READCFG,ATC(':',READCFG)+1))
    dbparts(m.bparts,1) = ALLTRIM(SUBSTR(m.bodypart, 1, AT(" ",m.bodypart)-1))
  ENDCASE
 ENDDO
RETURN 