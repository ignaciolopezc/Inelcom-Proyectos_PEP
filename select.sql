select nombres,
       count(*) as cantidad
  from nombre_tabla
 group by nombres;