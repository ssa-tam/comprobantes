# Notas técnicas

- El Sheet "Accesos" de cada trimestre usa el mismo layout: filas 1-3 son
  título/instrucciones/encabezados, los datos arrancan en la fila 4
  (constante DATA_START).
- Columnas: # | Trimestre | Unidad | Usuario | Contraseña | QNA | Nómina | Tipo | ID Carpeta Drive | Estado
- Las contraseñas de las 41 unidades viven repetidas en cada Sheet de
  trimestre (por diseño del sistema original) — la fuente de verdad más
  confiable es `Control_Acceso` (usuario/estado) + cualquier fila existente
  en `SST_Todos_Trimestres` (usuario/contraseña).
- Carpetas especiales (bonos) NO tienen nivel de "tipo": cuelgan directo de
  la QNA y contienen carpetas de unidad directamente.
- Nóminas normales (Ord/Extras) SÍ tienen nivel de "tipo" (420, 610, Estatal,
  Eventual, For 1/2/3, Homologados, IMSS-Bienestar, Jubilados, Reg N/S,
  Aportación EST) antes del nivel de unidad.
