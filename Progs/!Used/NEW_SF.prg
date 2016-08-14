set talk off
set stat off
set defa to c:\work

* if 3=2

use base\people in 1
wait "Индексация файла people ..." wind nowa
sele 1
 inde on polis to base\people
set inde to base\people
use base\main_add in 2
sele 2

recno=1
param='xxx'
m.wasborn={01.01.1900}
m.pol=.f.
m.found=.f.

sele 2

scan
 wait "Простановка данных в main ..."+str(round((recno/reccount())*100,0),3)+" %" wind nowa
 if polis!=param
  x=polis
  sele 1
  seek x
  if found()
   m.wasborn=wasborn
   m.pol=pol
   m.found=.t.
  else
   m.found=.f. 
  endi
  sele 2
  repl wasborn with m.wasborn, pol with m.pol, found with m.found
 endi 
 repl wasborn with m.wasborn, pol with m.pol, found with m.found
 param=polis
 recno=recno+1
endscan

* endi

* if 3=2

clos data
use base\main_add in 1
use base\men in 2
use base\women in 3
sele 2
inde on year(wasborn) to base\men
set inde to base\men
sele 3
inde on year(wasborn) to base\women
set inde to base\women

set near on

param='xxx'
recno=1

sele 1

scan
 wait "Замена данных (polis) в main ..."+str(round((recno/reccount())*100,0),3)+" %" wind nowa
 if polis!=param 
  x=year(wasborn)
  if pol=.t.
   sele 2
  else
   sele 3
  endi 
  seek x
  if found()
   if used=.t.
    do while used=.t.
     skip
    enddo 
   endi 
   repl used with .t.
   m.polis=polis
   m.karta=poliz
   m.pol=pol
   m.wasborn=wasborn
   m.found=.t.
  else
   m.found=.f.
  endi  
   sele 1
   repl new_polis with m.polis, new_karta with m.karta, new_pol with m.pol, ;
   	new_born with m.wasborn, replace with m.found
 endi
 repl new_polis with m.polis, new_karta with m.karta, new_pol with m.pol, ;
    new_born with m.wasborn, replace with m.found
 param=polis
 recno=recno+1
endscan

set near off

* endi
  
set talk on
set stat on
