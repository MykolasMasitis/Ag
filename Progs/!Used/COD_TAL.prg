set talk off
* set defa to \sfact\work
* use base\main in 1
* sele 1

m.cod_tal=189225769766

scan
perem=dtoc(wdate)+str(cod_doc,5)+dtoc(odate)+polis

do while dtoc(wdate)+str(cod_doc,5)+dtoc(odate)+polis=perem
 repl cod_tal with str(m.cod_tal,12)
 skip
endd

skip -1
m.cod_tal=m.cod_tal+1

ends

set talk on