/* =============================================================================
   Excel Intensivo — Contenido del curso
   Curso intensivo de Excel aplicado a gestión comercial

   Todo contextualizado a una tienda con stands móviles, eventos,
   otras sedes y proveedores en consignación. Excel en español.
   Se prioriza lo moderno (BUSCARX / funciones dinámicas) indicando el
   equivalente en Excel 2016/2019/2021.

   EJERCICIOS INTERACTIVOS — tipos (ver engine.js / app.js):
     mc       — opción múltiple           { q, options:[str], answer:idx, explain }
     multi    — varias correctas          { q, options:[str], answer:[idx], explain }
     order    — ordenar pasos             { q, steps:[str en orden correcto], explain }
     classify — clasificar en categorías  { q, cats:[str], items:[{t,c:idx}], explain }
     build    — armar fórmula con fichas   { q, bank:[tok], answer:[tok], explain }
     formula  — escribir la fórmula        { q, grid, cell, model, (expect), (accept), explain }
     predict  — predecir el resultado      { q, grid, formula, (expect), (fmt), explain }
   Las grillas: cols[] es la fila 1 de Excel; rows[][] son valores crudos
   (número o texto, o "=..." para celdas ya calculadas). money/pct: índices
   de columnas a formatear al mostrar.
   ========================================================================== */

const COURSE = {

  meta: {
    student: "estudiante",
    role: "Gestión comercial",
    priorityOrder: [4, 5, 6, 7, 2, 3, 1]
  },

  modules: [

    /* ===================== MÓDULO 1 ===================== */
    {
      id: 1,
      title: "Fundamentos de Excel",
      short: "Celdas, formato, referencias, ordenar y filtrar",
      goal: "Moverte con soltura en una hoja y entender qué es un dato, un formato y una referencia.",
      lessons: [
        {
          h: "Libro, hoja, celda, rango",
          body:
            "<p>Un <b>libro</b> es el archivo (.xlsx). Dentro tiene una o varias <b>hojas</b> (las pestañas de abajo). Cada hoja es una grilla de <b>celdas</b>; cada celda se nombra por su <b>columna + fila</b>: <code>A1</code>, <code>C15</code>.</p>" +
            "<p>Un <b>rango</b> es un bloque de celdas: <code>A1:D20</code> (esquina superior izquierda : esquina inferior derecha). Una columna entera es <code>D:D</code>, una fila entera <code>7:7</code>.</p>" +
            "<p>El <b>Cuadro de nombres</b> (arriba a la izquierda, al lado de la barra de fórmulas) te dice qué celda está seleccionada y te deja saltar escribiendo una referencia. También sirve para ponerle un <b>nombre</b> a un rango (ej. <code>Alicuota_IVA</code>) y usarlo en fórmulas.</p>"
        },
        {
          h: "El valor de una celda no es su formato",
          body:
            "<p>Esta distinción es la que más se malinterpreta. Una celda guarda un <b>valor real</b> (un número, un texto, una fecha) y por separado tiene un <b>formato</b> (cómo se muestra).</p>" +
            "<p>Si ves <code>$ 9.000,00</code>, el valor real puede ser <code>9000</code>. Si ves <code>15%</code>, el valor real es <code>0,15</code>. Las fechas son números por dentro (01/01/1900 = 1): por eso podés restar dos fechas y te da la cantidad de días.</p>" +
            "<p><b>Truco de diagnóstico:</b> por defecto Excel <b>alinea los números a la derecha y el texto a la izquierda</b>. Si una columna de importes te aparece pegada a la izquierda, son <b>números guardados como texto</b> y van a romper sumas y BUSCARV. Se arreglan con <i>Datos › Texto en columnas</i> o multiplicando por 1.</p>",
          checks: [
            {
              id: "m1c1", type: "mc",
              q: "Una columna de importes muestra los números pegados a la IZQUIERDA de la celda. ¿Qué son?",
              options: [
                "Números con formato de moneda",
                "Texto disfrazado de número: van a romper SUMA y BUSCARV",
                "Números negativos",
                "Celdas protegidas contra edición"
              ],
              answer: 1,
              explain: "<p>Excel alinea los números a la derecha y el texto a la izquierda. Un número pegado a la izquierda es <b>texto</b>. Se arregla multiplicando por 1 o con Datos › Texto en columnas.</p>"
            }
          ]
        },
        {
          h: "Formato de celdas",
          body:
            "<p><b>Ctrl+1</b> abre el cuadro de formato. Los que vas a usar sí o sí:</p>" +
            "<ul>" +
            "<li><b>Número</b> con separador de miles y 0 o 2 decimales.</li>" +
            "<li><b>Moneda / Contabilidad</b> ($). Contabilidad alinea el símbolo a la izquierda y los números por la coma decimal: queda más prolijo en reportes.</li>" +
            "<li><b>Porcentaje</b>: recordá que multiplica la vista por 100. <code>0,3</code> → <code>30%</code>.</li>" +
            "<li><b>Fecha</b>: elegí un formato corto consistente para todo el archivo.</li>" +
            "</ul>" +
            "<p>Cambiar el formato <b>nunca</b> cambia el valor ni el resultado de una fórmula: solo la presentación.</p>"
        },
        {
          h: "Formato condicional",
          body:
            "<p><i>Inicio › Formato condicional</i>. Pinta celdas según una regla, y se recalcula solo cuando cambian los datos.</p>" +
            "<ul>" +
            "<li><b>Resaltar reglas de celdas:</b> «es mayor que», «entre», «el texto contiene». Ej: pintar de rojo el stock que sea <b>menor o igual</b> al punto de reposición.</li>" +
            "<li><b>Escalas de color / barras de datos / conjuntos de iconos:</b> para ver de un vistazo qué punto de venta rinde mejor.</li>" +
            "<li><b>«Usar una fórmula»:</b> la opción potente. Con <code>=$F2&lt;=$G2</code> podés pintar <b>toda la fila</b> cuando el stock (F) cae por debajo del ROP (G).</li>" +
            "</ul>"
        },
        {
          h: "Fórmulas básicas y orden de operaciones",
          body:
            "<p>Toda fórmula empieza con <code>=</code>. Operadores: <code>+ - * / ^</code> y <code>%</code>. El orden es el matemático: primero <code>()</code>, luego <code>^</code>, luego <code>* /</code>, luego <code>+ -</code>. Ante la duda, <b>paréntesis</b>.</p>" +
            "<p>Funciones de arranque:</p>" +
            "<ul>" +
            "<li><code>=SUMA(B2:B50)</code> — total. Atajo: seleccionás la columna y <b>Alt + =</b>.</li>" +
            "<li><code>=PROMEDIO(...)</code>, <code>=MAX(...)</code>, <code>=MIN(...)</code>.</li>" +
            "<li><code>=CONTAR(...)</code> cuenta celdas con número; <code>=CONTARA(...)</code> cuenta celdas no vacías (sirven para «cuántas ventas hay»).</li>" +
            "</ul>",
          sheet: {
            caption: "Hoja: Ventas Tienda — total del día",
            cols: ["Producto", "Unidades", "Importe"],
            rows: [
              ["Llavero", "3", "9.000"],
              ["Libro «Historia 1908»", "1", "18.500"],
              ["Taza esmaltada", "2", "13.000"],
              ["TOTAL", "6", "40.500"]
            ],
            highlight: { row: 5, col: 3 },
            formula: { cell: "C5", text: "=SUMA(C2:C4)", result: "40.500" }
          },
          checks: [
            {
              id: "m1c2", type: "formula",
              q: "Escribí la fórmula que va en C5 para obtener el total de la columna Importe.",
              grid: {
                cols: ["Producto", "Unidades", "Importe"],
                rows: [["Llavero", 3, 9000], ["Libro «1908»", 1, 18500], ["Taza esmaltada", 2, 13000], ["TOTAL", "", ""]],
                money: [2]
              },
              cell: "C5",
              model: "=SUMA(C2:C4)",
              explain: "<p><code>=SUMA(C2:C4)</code> suma el rango. Atajo: seleccionás C5 y apretás <b>Alt + =</b>.</p>"
            }
          ]
        },
        {
          h: "Referencias relativas vs. absolutas ($)",
          body:
            "<p>Cuando copiás una fórmula hacia abajo, Excel <b>corre</b> las referencias: <code>=B2*C2</code> pasa a <code>=B3*C3</code>. Eso es una referencia <b>relativa</b> y casi siempre es lo que querés.</p>" +
            "<p>El signo <b>$</b> <i>ancla</i>. <code>$B$2</code> no se mueve nunca. <code>B$2</code> fija la fila, <code>$B2</code> fija la columna. La tecla <b>F4</b> rota entre las cuatro variantes mientras editás.</p>" +
            "<p><b>Caso típico del puesto:</b> tenés la alícuota de IVA en <code>$B$2</code> y una columna de precios sin IVA. En <code>C2</code> escribís <code>=A2*(1+$B$2)</code> y copiás hacia abajo: <code>A2</code> corre, <code>$B$2</code> queda fijo apuntando siempre a la alícuota.</p>",
          sheet: {
            caption: "Precio final con IVA — la alícuota (0,21) está en B2",
            cols: ["Precio s/IVA", "Alícuota", "Precio c/IVA"],
            rows: [["1.200", "0,21", "1.452"], ["2.500", "", "3.025"], ["4.000", "", "4.840"]],
            formula: { cell: "C2", text: "=A2*(1+$B$2)", result: "1.452" }
          },
          checks: [
            {
              id: "m1c3", type: "build",
              q: "Armá la fórmula para C2 (precio con IVA de A2; la alícuota está una sola vez, en B2) de modo que se pueda copiar hacia abajo sin romperse.",
              bank: ["=", "A2", "*", "(", "1", "+", "B2", "$B$2", ")"],
              answer: ["=", "A2", "*", "(", "1", "+", "$B$2", ")"],
              explain: "<p><code>A2</code> tiene que <b>correr</b> con cada fila; <code>$B$2</code> tiene que quedar <b>fijo</b> apuntando siempre a la alícuota. Si dejás <code>B2</code> sin <code>$</code>, al copiar la fórmula pasa a <code>B3</code>, <code>B4</code>… y trae celdas vacías.</p>"
            }
          ]
        },
        {
          h: "Ordenar, filtrar y convertir en Tabla",
          body:
            "<p><b>Ordenar:</b> <i>Datos › Ordenar</i>. Podés ordenar por varios niveles (primero Punto de venta A→Z, después Importe de mayor a menor). Seleccioná siempre todo el bloque, no una sola columna, o desarmás las filas.</p>" +
            "<p><b>Filtrar:</b> <b>Ctrl+Shift+L</b> pone las flechitas de Autofiltro. Filtros de texto («contiene»), de número («mayor que», «10 mejores») y de fecha («este mes»).</p>" +
            "<p><b>Tabla (Ctrl+T):</b> convierte un rango en un objeto «Tabla». Ventajas: se pinta sola, filtra y ordena desde el encabezado, <b>se expande automáticamente</b> cuando agregás filas, y las fórmulas usan nombres legibles (<code>=[@Unidades]*[@Precio]</code>) en vez de <code>B2*C2</code>. Toda tabla dinámica o BUSCARX que armes sobre una Tabla se actualiza sin que tengas que corregir rangos.</p>",
          checks: [
            {
              id: "m1c4", type: "mc",
              q: "¿Por qué conviene convertir el rango en Tabla (Ctrl+T) antes de armar una tabla dinámica o un BUSCARX?",
              options: [
                "Para que la tabla dinámica se actualice sola al abrir el archivo",
                "Porque la Tabla se expande al agregar filas y todo lo que apunta a ella las toma sin reajustar rangos",
                "Para poder usar más de 65.000 filas",
                "Para convertir las fechas de texto en fechas de verdad"
              ],
              answer: 1,
              explain: "<p>Una Tabla <b>crece sola</b>. Como la tabla dinámica y las fórmulas apuntan al <b>nombre</b> de la Tabla y no a un rango fijo, cuando cargás ventas nuevas alcanza con Actualizar.</p>"
            }
          ]
        },
        {
          h: "Atajos que te hacen ver rápido",
          body:
            "<ul>" +
            "<li><b>Ctrl + flechas</b>: saltás al borde de los datos. <b>Ctrl + Shift + flechas</b>: seleccionás hasta ahí.</li>" +
            "<li><b>Ctrl + T</b>: Tabla · <b>Alt + =</b>: autosuma · <b>F2</b>: editar celda · <b>F4</b>: alternar $ / repetir última acción.</li>" +
            "<li><b>Ctrl + ;</b>: fecha de hoy fija · <b>Ctrl + Shift + L</b>: filtro · <b>Ctrl + 1</b>: formato de celda.</li>" +
            "<li><b>Ctrl + Av/Re Pág</b>: cambiar de hoja · <b>Ctrl + Inicio</b>: volver a A1.</li>" +
            "</ul>"
        }
      ],
      exercises: [
        {
          id: "m1e1", type: "formula",
          q: "Total del día. Escribí en C5 la fórmula que suma la columna Importe.",
          grid: {
            cols: ["Producto", "Unidades", "Importe"],
            rows: [["Llavero", 3, 9000], ["Programa", 5, 20000], ["Disco", 2, 24000], ["TOTAL", "", ""]],
            money: [2]
          },
          cell: "C5", model: "=SUMA(C2:C4)",
          explain: "<p><code>=SUMA(C2:C4)</code> → 53.000. Con <b>Alt + =</b> Excel la escribe sola.</p>"
        },
        {
          id: "m1e2", type: "formula",
          q: "Precio con IVA. La alícuota (0,21) está una sola vez, en B2. Escribí en C2 la fórmula lista para copiar hacia abajo.",
          grid: {
            cols: ["Precio s/IVA", "Alícuota", "Precio c/IVA"],
            rows: [[1200, 0.21, ""], [2500, "", ""], [4000, "", ""]],
            money: [0, 2]
          },
          cell: "C2",
          accept: ["=A2*(1+$B$2)", "=A2*(1+$B2)", "=(1+$B$2)*A2"],
          model: "=A2*(1+$B$2)",
          explain: "<p>Solo B2 tiene la alícuota; las demás filas de la columna B están vacías. Si no anclás con <code>$B$2</code>, al copiar hacia abajo la fórmula multiplica por 1 y no suma el IVA. Resultado en C2: 1.452.</p>"
        },
        {
          id: "m1e3", type: "predict",
          q: "Variación de ventas de la Tienda entre agosto y septiembre.",
          grid: {
            cols: ["Punto de venta", "Agosto", "Septiembre"],
            rows: [["Tienda central", 2900000, 3200000]],
            money: [1, 2]
          },
          cell: "D2", formula: "=(C2-B2)/B2", fmt: "pct",
          explain: "<p>Variación % = (actual − anterior) / anterior = (3.200.000 − 2.900.000) / 2.900.000 ≈ <b>10,3%</b>. Se le da formato de porcentaje.</p>"
        },
        {
          id: "m1e4", type: "classify",
          q: "Clasificá cada referencia según qué le pasa cuando copiás la fórmula a la fila de abajo.",
          cats: ["Se mueve al copiar", "Queda fija al copiar"],
          items: [
            { t: "A1", c: 0 },
            { t: "$A$1", c: 1 },
            { t: "$B$1  (la alícuota, en =A2*(1+$B$1))", c: 1 },
            { t: "A2  (el precio, en =A2*(1+$B$1))", c: 0 }
          ],
          explain: "<p>Sin <code>$</code> la referencia es <b>relativa</b> y corre con la fila. Con <code>$</code> queda <b>anclada</b>. Por eso el precio (que cambia fila a fila) va sin <code>$</code> y la alícuota (única) va con <code>$</code>.</p>"
        },
        {
          id: "m1e5", type: "order",
          q: "Ordená los pasos para ordenar una lista de ventas de mayor a menor importe sin desarmar las filas.",
          steps: [
            "Seleccionar TODO el bloque de datos (todas las columnas, con encabezados)",
            "Datos › Ordenar",
            "Elegir la columna «Importe» y el sentido «De mayor a menor»",
            "Aceptar y verificar que cada fila quedó entera (el producto con su importe)"
          ],
          explain: "<p>El error clásico es seleccionar solo la columna Importe: Excel ordena esa columna y deja el resto quieto, y se te mezclan los datos. Si el rango es una Tabla, alcanza con la flechita del encabezado.</p>"
        }
      ],
      checklist: [
        "Explico la diferencia entre el valor de una celda y su formato",
        "Escribo una fórmula con referencia absoluta ($) y digo para qué sirve",
        "Convierto un rango en Tabla (Ctrl+T) y menciono una ventaja",
        "Ordeno una lista por varios niveles sin desarmar las filas",
        "Aplico formato condicional con una fórmula para pintar una fila entera"
      ],
      keys: [
        { k: "= ", v: "Toda fórmula arranca con el signo igual." },
        { k: "$A$1", v: "Referencia absoluta: no se mueve al copiar. F4 alterna las variantes." },
        { k: "Ctrl+T", v: "Convierte un rango en Tabla: se expande sola y ordena/filtra desde el encabezado." },
        { k: "Alt + =", v: "Autosuma de la columna o fila." },
        { k: "Alineación", v: "Números a la derecha, texto a la izquierda. Un número a la izquierda = texto disfrazado." }
      ]
    },

    /* ===================== MÓDULO 2 ===================== */
    {
      id: 2,
      title: "Fórmulas aplicadas al análisis comercial",
      short: "SI, BUSCARV / BUSCARX, SUMAR.SI.CONJUNTO, márgenes y variaciones",
      goal: "Cruzar tablas por código, sumar con condiciones y calcular los indicadores comerciales básicos.",
      lessons: [
        {
          h: "Funciones lógicas: SI, Y, O, SI.ERROR",
          body:
            "<p><code>=SI(prueba; valor_si_verdadero; valor_si_falso)</code>. Ejemplo: <code>=SI(B2=\"Sala principal\";\"Central\";\"Otra sede\")</code>.</p>" +
            "<p><b>SI anidado:</b> un SI dentro de otro para más de dos caminos. <code>=SI(C2&gt;=100000;\"A\";SI(C2&gt;=50000;\"B\";\"C\"))</code>. Si tenés Excel 365, <code>SI.CONJUNTO</code> (IFS) lo hace más legible.</p>" +
            "<p><code>Y(cond1; cond2; ...)</code> es verdadero si <b>todas</b> se cumplen; <code>O(...)</code> si <b>al menos una</b>. Se usan dentro de un SI: <code>=SI(Y(B2=\"Tienda\";C2&gt;0);\"OK\";\"Revisar\")</code>.</p>" +
            "<p><code>=SI.ERROR(formula; valor_si_falla)</code> reemplaza cualquier error (#N/A, #¡DIV/0!, etc.) por lo que vos quieras. Clave para que un reporte no se muestre lleno de errores: <code>=SI.ERROR(BUSCARX(...);\"Sin dato\")</code>.</p>",
          checks: [
            {
              id: "m2c1", type: "formula",
              q: "En D2 clasificá la función: si el punto de venta (B2) es «Central», poné «Sala principal»; si no, «Otra sede».",
              grid: {
                cols: ["Fecha", "Punto de venta", "Importe", "Tipo de sede"],
                rows: [["03/08", "Central", 90000, ""], ["05/08", "La Rural", 40000, ""]]
              },
              cell: "D2",
              model: "=SI(B2=\"Central\";\"Sala principal\";\"Otra sede\")",
              explain: "<p><code>=SI(B2=\"Central\";\"Sala principal\";\"Otra sede\")</code>. La prueba va primero, después el valor si es verdadero y por último el si es falso.</p>"
            }
          ]
        },
        {
          h: "BUSCARV — el clásico para cruzar tablas",
          body:
            "<p><code>=BUSCARV(valor_buscado; tabla; nº_columna; FALSO)</code></p>" +
            "<p>Busca <b>valor_buscado</b> en la <b>primera columna</b> de <b>tabla</b> y devuelve el dato que está en la <b>columna nº_columna</b> de esa misma fila.</p>" +
            "<ul>" +
            "<li>El 4º argumento siempre <b>FALSO</b> (o <code>0</code>) = coincidencia exacta. Con VERDADERO da resultados raros si la tabla no está ordenada.</li>" +
            "<li><b>nº_columna se cuenta desde la primera columna de la tabla</b>, no desde la A de la hoja.</li>" +
            "<li>Limitaciones: solo mira <b>hacia la derecha</b> (el código tiene que estar en la 1ª columna) y si alguien <b>inserta una columna</b> en el medio, el número queda mal y trae otro dato sin avisar.</li>" +
            "</ul>",
          sheet: {
            caption: "Traer el costo del producto a la hoja de Ventas",
            cols: ["Código", "Producto", "Unid.", "Costo unit."],
            rows: [["LLV-001", "Llavero", "3", "1.700"], ["LIB-014", "Libro «1908»", "1", "9.200"], ["TAZ-007", "Taza esmaltada", "2", "3.400"]],
            formula: { cell: "D2", text: "=BUSCARV(A2; Productos!$A:$D; 4; FALSO)", result: "1.700" }
          },
          checks: [
            {
              id: "m2c2", type: "mc",
              q: "¿Cuál trae el COSTO (que está en la 4ª columna de la tabla Productos) para el código de A2, con coincidencia exacta?",
              options: [
                "=BUSCARV(A2; Productos; 4; VERDADERO)",
                "=BUSCARV(A2; Productos; 4; FALSO)",
                "=BUSCARV(4; Productos; A2; FALSO)",
                "=BUSCARV(A2; Productos; FALSO; 4)"
              ],
              answer: 1,
              explain: "<p>Orden: valor buscado, tabla, número de columna a devolver, y <b>FALSO</b> para coincidencia exacta. La columna se cuenta desde la primera de la tabla.</p>"
            }
          ]
        },
        {
          h: "BUSCARX (XLOOKUP) — la versión moderna",
          body:
            "<p><code>=BUSCARX(valor_buscado; matriz_donde_buscar; matriz_a_devolver; [si_no_se_encuentra])</code></p>" +
            "<p>Hace lo mismo que BUSCARV pero mejor:</p>" +
            "<ul>" +
            "<li>Busca en <b>cualquier dirección</b> (la columna del código puede estar a la derecha del dato que querés).</li>" +
            "<li>No depende de un <b>número</b> de columna: le pasás directamente el rango a devolver, así que insertar columnas no lo rompe.</li>" +
            "<li>Trae incorporado el <b>«si no se encuentra»</b>: <code>=BUSCARX(A2; Productos[Código]; Productos[Costo]; \"Sin dato\")</code>.</li>" +
            "</ul>" +
            "<p><b>Disponibilidad:</b> Excel 365 y Excel para la web. En <b>Excel 2016 / 2019 / 2021</b> (el que tenés localmente) <b>no existe</b> → se resuelve con INDICE + COINCIDIR. En la entrevista conviene mencionarlo: «uso BUSCARX donde está disponible; en versiones sin BUSCARX, INDICE+COINCIDIR».</p>",
          checks: [
            {
              id: "m2c3", type: "mc",
              q: "Estás en Excel 2021 y BUSCARX no aparece. ¿Con qué combinación lográs lo mismo (buscar por código y traer un dato)?",
              options: ["CONCATENAR", "INDICE + COINCIDIR", "SUMAR.SI", "BUSCARV con VERDADERO"],
              answer: 1,
              explain: "<p><code>COINCIDIR</code> devuelve la posición de la fila donde está el código; <code>INDICE</code> trae el dato de esa posición en la columna que quieras. Funciona en cualquier versión y en cualquier dirección.</p>"
            }
          ]
        },
        {
          h: "INDICE + COINCIDIR — el reemplazo universal",
          body:
            "<p><code>=COINCIDIR(valor; rango_de_una_columna; 0)</code> devuelve la <b>posición</b> (número de fila dentro del rango) donde aparece el valor. El <code>0</code> es coincidencia exacta.</p>" +
            "<p><code>=INDICE(rango_a_devolver; posición)</code> devuelve el dato que está en esa posición.</p>" +
            "<p>Combinadas: <code>=INDICE(Productos!$C:$C; COINCIDIR(A2; Productos!$A:$A; 0))</code> → «traé el costo (columna C) de la fila donde el código (columna A) coincide con A2». Funciona en cualquier versión y en cualquier dirección.</p>"
        },
        {
          h: "Agregación condicional: SUMAR.SI y .CONJUNTO",
          body:
            "<p><code>=SUMAR.SI(rango_criterio; criterio; rango_a_sumar)</code> — suma solo lo que cumple una condición. <code>=SUMAR.SI(B:B; \"Tienda\"; E:E)</code> = total vendido por la Tienda.</p>" +
            "<p><code>=CONTAR.SI(rango; criterio)</code> y <code>=PROMEDIO.SI(...)</code> son análogas.</p>" +
            "<p><b>Versiones .CONJUNTO</b> para varios criterios a la vez. <b>Ojo que cambia el orden de argumentos</b>: el rango a sumar va <b>primero</b>.</p>" +
            "<p><code>=SUMAR.SI.CONJUNTO(rango_suma; rango_crit1; crit1; rango_crit2; crit2; ...)</code></p>" +
            "<p>Ejemplo del puesto: total del punto de venta «Tienda» en agosto:<br><code>=SUMAR.SI.CONJUNTO(Importe; PuntoDeVenta; \"Tienda\"; Fecha; \"&gt;=2025-08-01\"; Fecha; \"&lt;=2025-08-31\")</code></p>" +
            "<p>Los criterios de texto y comparadores van entre comillas: <code>\"&gt;100\"</code>, <code>\"&lt;&gt;0\"</code>, <code>\"Tienda*\"</code> (comodín).</p>",
          checks: [
            {
              id: "m2c4", type: "formula",
              q: "En H2 escribí el total de Importe vendido por la «Tienda» (columna B = punto de venta, columna C = importe).",
              grid: {
                cols: ["Fecha", "Punto de venta", "Importe"],
                rows: [
                  ["02/08", "Tienda", 9000], ["02/08", "Stands móviles", 4200], ["03/08", "Tienda", 18500],
                  ["04/08", "Eventos", 30000], ["05/08", "Tienda", 6500]
                ],
                money: [2]
              },
              cell: "H2",
              model: "=SUMAR.SI(B2:B6;\"Tienda\";C2:C6)",
              explain: "<p><code>=SUMAR.SI(B2:B6;\"Tienda\";C2:C6)</code> → 34.000. En SUMAR.SI (sin CONJUNTO) el rango a sumar va al final.</p>"
            }
          ]
        },
        {
          h: "Cálculos comerciales que tenés que saber de memoria",
          body:
            "<ul>" +
            "<li><b>Variación %</b> = (actual − anterior) / anterior. Formato porcentaje. Si el anterior puede ser 0, envolvé con SI.ERROR.</li>" +
            "<li><b>Margen (sobre venta)</b> = (Precio − Costo) / <b>Precio</b>. Es el % de la venta que te queda.</li>" +
            "<li><b>Markup (sobre costo)</b> = (Precio − Costo) / <b>Costo</b>. Es cuánto le agregás al costo. <b>No son lo mismo</b>: un markup del 100% equivale a un margen del 50%.</li>" +
            "<li><b>Participación / mix</b> = venta_de_la_categoría / venta_total (con el total anclado con $).</li>" +
            "<li><b>Acumulado (running total)</b> = <code>=SUMA($E$2:E2)</code> copiado hacia abajo: el inicio fijo, el final corre.</li>" +
            "<li><b>Precio con IVA</b> = Precio × (1 + alícuota). <b>Sacar el IVA</b> de un precio final = Precio / (1 + alícuota).</li>" +
            "</ul>",
          checks: [
            {
              id: "m2c5", type: "predict",
              q: "Un libro se vende a $ 3.000 y a la Tienda le cuesta $ 1.700.",
              grid: { cols: ["Precio", "Costo", "Margen"], rows: [[3000, 1700, ""]], money: [0, 1] },
              cell: "C2", formula: "=(A2-B2)/A2", fmt: "pct",
              explain: "<p>Margen sobre venta = (3.000 − 1.700) / 3.000 ≈ <b>43%</b>. El markup sobre costo sería 1.300 / 1.700 ≈ 76% — distinto número, mismo peso.</p>"
            }
          ]
        },
        {
          h: "Leer y tapar errores",
          body:
            "<ul>" +
            "<li><b>#N/A</b> — «no encontré» (BUSCARV/BUSCARX/COINCIDIR sin coincidencia; a veces por espacios o números-texto).</li>" +
            "<li><b>#¡REF!</b> — la fórmula apunta a una celda que se borró.</li>" +
            "<li><b>#¡VALOR!</b> — tipo de dato incorrecto (texto donde va número).</li>" +
            "<li><b>#¡DIV/0!</b> — división por cero o por celda vacía.</li>" +
            "<li><b>#¿NOMBRE?</b> — nombre de función mal escrito o un texto sin comillas.</li>" +
            "</ul>" +
            "<p>Para diagnosticar de dónde viene: <b>Fórmulas › Evaluar fórmula</b> o seleccioná un pedazo de la fórmula y apretá <b>F9</b> para ver ese trozo calculado. Para presentar: <code>=SI.ERROR(...; \"—\")</code>. Cuidado de no usar SI.ERROR para <i>esconder</i> un problema real de datos.</p>",
          checks: [
            {
              id: "m2c6", type: "classify",
              q: "Emparejá cada error de Excel con su causa.",
              cats: ["No encontró", "Referencia borrada", "Tipo incorrecto", "División por cero"],
              items: [
                { t: "#N/A", c: 0 },
                { t: "#¡REF!", c: 1 },
                { t: "#¡VALOR!", c: 2 },
                { t: "#¡DIV/0!", c: 3 }
              ],
              explain: "<p>En una liquidación, <b>#N/A</b> suele ser una venta con un código sin dar de alta → hay que corregir el dato, no taparlo con SI.ERROR.</p>"
            }
          ]
        }
      ],
      exercises: [
        {
          id: "m2e1", type: "formula",
          q: "Traé el costo unitario a la hoja de Ventas. En G2, buscá el código de A2 en la tabla de la derecha (columnas E y F) y devolvé el costo.",
          grid: {
            cols: ["Código", "Unid.", "Importe", "·", "Cód. prod.", "Costo", "Costo unit."],
            rows: [
              ["LLV-001", 3, 9000, "", "LLV-001", 1700, ""],
              ["LIB-014", 1, 18500, "", "LIB-014", 9200, ""],
              ["TAZ-007", 2, 13000, "", "TAZ-007", 3400, ""]
            ],
            money: [2, 5, 6]
          },
          cell: "G2",
          model: "=BUSCARX(A2;E2:E4;F2:F4)",
          explain: "<p><code>=BUSCARX(A2;E2:E4;F2:F4)</code> → 1.700. En Excel 2021: <code>=INDICE(F2:F4;COINCIDIR(A2;E2:E4;0))</code> (también lo da por válido).</p>"
        },
        {
          id: "m2e2", type: "formula",
          q: "Margen %. En D2 calculá el margen sobre venta (precio en B, costo en C).",
          grid: {
            cols: ["Producto", "Precio", "Costo", "Margen %"],
            rows: [["Llavero", 3000, 1700, ""], ["Remera", 12000, 5400, ""]],
            money: [1, 2]
          },
          cell: "D2", model: "=(B2-C2)/B2", fmt: "pct",
          explain: "<p>Margen sobre venta = (Precio − Costo) / Precio = (3.000 − 1.700) / 3.000 ≈ <b>43%</b>.</p>"
        },
        {
          id: "m2e3", type: "formula",
          q: "Total de la «Tienda» en agosto. En H2, sumá el Importe donde el punto de venta sea «Tienda» y la fecha esté en agosto de 2025.",
          grid: {
            cols: ["Fecha", "Punto de venta", "Importe"],
            rows: [
              ["2025-08-03", "Tienda", 9000],
              ["2025-08-20", "Tienda", 18500],
              ["2025-09-01", "Tienda", 6500],
              ["2025-08-11", "Stands móviles", 4200]
            ],
            money: [2]
          },
          cell: "H2",
          model: "=SUMAR.SI.CONJUNTO(C2:C5;B2:B5;\"Tienda\";A2:A5;\">=2025-08-01\";A2:A5;\"<=2025-08-31\")",
          explain: "<p>Da 27.500 (los dos consumos de la Tienda de agosto; el del 1/9 queda afuera). En SUMAR.SI.CONJUNTO el <b>rango a sumar va primero</b>.</p>"
        },
        {
          id: "m2e4", type: "build",
          q: "Armá el SUMAR.SI.CONJUNTO que suma «Importe» (C2:C6) para el proveedor «Editorial» (A2:A6).",
          bank: ["=", "SUMAR.SI.CONJUNTO(", "C2:C6", "A2:A6", "\"Editorial\"", ";", ")"],
          answer: ["=", "SUMAR.SI.CONJUNTO(", "C2:C6", ";", "A2:A6", ";", "\"Editorial\"", ")"],
          explain: "<p>Orden: <b>rango a sumar</b>, después los pares <b>rango de criterio ; criterio</b>. Es al revés que SUMAR.SI, donde el rango a sumar va al final.</p>"
        },
        {
          id: "m2e5", type: "predict",
          q: "El código «SOU-999» no está dado de alta. ¿Qué muestra la celda?",
          grid: {
            cols: ["Código", "·", "Cód. prod.", "Costo"],
            rows: [["SOU-999", "", "LLV-001", 1700], ["", "", "LIB-014", 9200]],
            money: [3]
          },
          cell: "B2", formula: "=SI.ERROR(BUSCARX(A2;C2:C3;D2:D3);\"Sin dato\")",
          explain: "<p>BUSCARX no encuentra «SOU-999» → devolvería #N/A, pero SI.ERROR lo reemplaza por <b>«Sin dato»</b>. Ojo: para presentar está bien; para una liquidación, ese código hay que darlo de alta.</p>"
        }
      ],
      checklist: [
        "Explico con mis palabras qué hace BUSCARV y por qué el 4º argumento es FALSO",
        "Digo dos ventajas de BUSCARX sobre BUSCARV y qué uso cuando no está disponible",
        "Armo un SUMAR.SI.CONJUNTO con dos o tres criterios (y sé que el rango a sumar va primero)",
        "Calculo variación %, margen y participación, y explico margen vs. markup",
        "Sé qué significa #N/A, cómo diagnosticarlo y cuándo (y cuándo no) taparlo con SI.ERROR"
      ],
      keys: [
        { k: "BUSCARV(val; tabla; nºcol; FALSO)", v: "Busca en la 1ª columna, devuelve de la columna nºcol. Solo mira a la derecha." },
        { k: "BUSCARX(val; buscar_en; devolver; \"si no\")", v: "Cualquier dirección, sin nº de columna. Excel 365. En 2021 no está." },
        { k: "INDICE(devolver; COINCIDIR(val; buscar_en; 0))", v: "Reemplazo de BUSCARX en toda versión." },
        { k: "SUMAR.SI.CONJUNTO(suma; c1; k1; c2; k2)", v: "Suma con varios criterios. El rango a sumar va primero." },
        { k: "Margen", v: "(PV − Costo) / PV. Markup usa / Costo. No son iguales." },
        { k: "Variación %", v: "(actual − anterior) / anterior." },
        { k: "SI.ERROR(fórmula; \"—\")", v: "Reemplaza errores para presentar. No abuses: puede esconder datos malos." }
      ]
    },

    /* ===================== MÓDULO 3 ===================== */
    {
      id: 3,
      title: "Tablas dinámicas y reportes de ventas",
      short: "Resumir miles de filas sin fórmulas y armar el reporte mensual",
      goal: "Construir un reporte de ventas por punto de venta y por mes, con % de participación y variación.",
      lessons: [
        {
          h: "Qué es una tabla dinámica y cuándo usarla",
          body:
            "<p>Una <b>tabla dinámica</b> (TD) resume miles de filas de datos <b>sin escribir una sola fórmula</b>, arrastrando campos a cuatro áreas:</p>" +
            "<ul>" +
            "<li><b>Filas</b> — lo que querés desglosar hacia abajo (Punto de venta, Producto, Proveedor).</li>" +
            "<li><b>Columnas</b> — un segundo corte hacia el costado (Mes, Categoría).</li>" +
            "<li><b>Valores</b> — el número a calcular (Suma de Importe, Recuento de operaciones).</li>" +
            "<li><b>Filtros</b> — un filtro general arriba de todo (Año, Canal).</li>" +
            "</ul>" +
            "<p><b>TD vs. fórmulas:</b> la TD es para <b>explorar y resumir rápido</b>, reordenar y probar cortes. Las <b>fórmulas</b> (SUMAR.SI.CONJUNTO) son para un <b>reporte fijo</b> que siempre tiene la misma forma y se actualiza solo. En la práctica se combinan: TD para analizar, fórmulas para el entregable estándar.</p>",
          checks: [
            {
              id: "m3c1", type: "classify",
              q: "¿A qué área de la tabla dinámica va cada campo si querés «ventas por punto de venta y por mes»?",
              cats: ["Filas o Columnas", "Valores", "Filtros"],
              items: [
                { t: "Punto de venta", c: 0 },
                { t: "Mes", c: 0 },
                { t: "Importe (suma)", c: 1 },
                { t: "Cantidad de operaciones", c: 1 },
                { t: "Año (para ver solo 2025)", c: 2 }
              ],
              explain: "<p>Las <b>categorías</b> (con qué desglosás) van a Filas o Columnas. Los <b>números</b> a medir van a Valores. Un filtro general de toda la tabla va a Filtros.</p>"
            }
          ]
        },
        {
          h: "Preparar los datos (el 80% del trabajo)",
          body:
            "<p>Una TD necesita datos <b>«planos»</b>:</p>" +
            "<ul>" +
            "<li>Una <b>fila = un registro</b> (una venta, un movimiento).</li>" +
            "<li>Encabezados <b>únicos</b> en la primera fila, sin celdas combinadas.</li>" +
            "<li><b>Sin filas ni columnas vacías</b> en el medio, sin subtotales intercalados.</li>" +
            "<li>Cada columna, un solo tipo de dato (fechas reales, no texto).</li>" +
            "</ul>" +
            "<p>Convertí el origen en <b>Tabla (Ctrl+T)</b> antes de crear la TD: cuando agregues ventas nuevas, la TD las toma con solo <b>Actualizar</b>, sin tocar rangos.</p>"
        },
        {
          h: "Crear y configurar",
          body:
            "<p><i>Insertar › Tabla dinámica</i>. Elegí el origen (la Tabla) y una hoja nueva.</p>" +
            "<p>Para el reporte del puesto: arrastrá <b>PuntoDeVenta</b> a Filas, <b>Importe</b> a Valores. Excel pone «Suma de Importe». Agregá <b>Mes</b> a Columnas.</p>" +
            "<p><b>Configuración de campo de valor</b> (clic derecho sobre un número › «Configuración de campo de valor»):</p>" +
            "<ul>" +
            "<li>Pestaña <b>Resumir valores por</b>: Suma, Recuento, Promedio, Máx.</li>" +
            "<li>Pestaña <b>Mostrar valores como</b>: <b>% del total general</b> (participación de cada punto de venta), <b>% del total de la columna</b>, <b>Diferencia respecto de</b> (mes anterior → variación).</li>" +
            "</ul>",
          sheet: {
            caption: "Resultado: ventas por punto de venta y mes",
            cols: ["Punto de venta", "Jul", "Ago", "Sep", "Total"],
            rows: [
              ["Tienda central", "1.20 M", "1.45 M", "1.60 M", "4.25 M"],
              ["Stands móviles", "0.35 M", "0.42 M", "0.38 M", "1.15 M"],
              ["Eventos", "0.60 M", "0.90 M", "1.10 M", "2.60 M"],
              ["Otras sedes", "0.10 M", "0.15 M", "0.12 M", "0.37 M"],
              ["Total general", "2.25 M", "2.92 M", "3.20 M", "8.37 M"]
            ],
            highlight: { row: 6, col: 5 }
          },
          checks: [
            {
              id: "m3c2", type: "mc",
              q: "Querés que cada punto de venta muestre qué % representa sobre la venta total. ¿Qué hacés en la tabla dinámica?",
              options: [
                "Una columna con una fórmula al costado de la tabla dinámica",
                "Configuración de campo de valor › Mostrar valores como › % del total general",
                "Formato de celda › Porcentaje",
                "Un gráfico de torta, en la tabla no se puede"
              ],
              answer: 1,
              explain: "<p>«Mostrar valores como» recalcula el campo: % del total general, % de la columna, diferencia respecto de otro período, etc. Sin escribir fórmulas.</p>"
            }
          ]
        },
        {
          h: "Agrupar fechas y valores",
          body:
            "<p>Si tenés fechas día por día y las ponés en Columnas, la TD te tira 90 columnas. Clic derecho sobre una fecha › <b>Agrupar</b> › elegí <b>Meses</b> (y Trimestres / Años si querés). Ahí aparece «Mes» como campo.</p>" +
            "<p>También podés agrupar <b>importes en rangos</b> (0–5.000, 5.000–10.000...) para ver la distribución del ticket.</p>" +
            "<p>Si en tu Excel las fechas no agrupan, casi seguro son <b>texto</b>: arreglá el origen primero.</p>",
          checks: [
            {
              id: "m3c3", type: "mc",
              q: "Ponés las fechas (día por día) en Columnas y la tabla dinámica genera decenas de columnas. Solución más limpia:",
              options: [
                "Borrar las fechas que sobran",
                "Clic derecho sobre una fecha › Agrupar › Meses",
                "Cambiar el formato de las celdas a «mmmm»",
                "Filtrar y dejar un solo día"
              ],
              answer: 1,
              explain: "<p>Agrupar por Meses (y Trimestres/Años) crea el campo «Mes» y colapsa las columnas. Si Excel no te deja agrupar, las fechas son texto: hay que arreglar el origen.</p>"
            }
          ]
        },
        {
          h: "Segmentaciones y escala de tiempo",
          body:
            "<p><b>Segmentación de datos (slicer):</b> <i>Analizar tabla dinámica › Insertar segmentación</i>. Son botones grandes para filtrar por Categoría, Proveedor, Canal. Más claro que el filtro desplegable y queda lindo para presentar.</p>" +
            "<p><b>Escala de tiempo (timeline):</b> lo mismo pero para fechas, con una barra deslizante por mes/trimestre.</p>" +
            "<p>Una misma segmentación puede controlar <b>varias TD a la vez</b>: clic derecho › <b>Conexiones de informe</b> › tildás las tablas. Así armás un mini-tablero donde un botón filtra todo.</p>"
        },
        {
          h: "Actualizar — el error más común",
          body:
            "<p>La TD <b>no se actualiza sola</b> cuando cambian los datos de origen. Después de cargar ventas nuevas:</p>" +
            "<ul>" +
            "<li>Clic derecho sobre la TD › <b>Actualizar</b>, o <b>Alt + F5</b>.</li>" +
            "<li>Actualizar todas: <b>Ctrl + Alt + F5</b> (o Datos › Actualizar todo).</li>" +
            "<li>Si el origen era un rango fijo y creció, <i>Analizar › Cambiar origen de datos</i>. Con una Tabla no hace falta.</li>" +
            "</ul>" +
            "<p>Opción recomendable: <i>Opciones de tabla dinámica › Datos › «Actualizar al abrir el archivo»</i>.</p>",
          checks: [
            {
              id: "m3c4", type: "mc",
              q: "Cargaste ventas nuevas en el origen pero la tabla dinámica sigue mostrando los totales viejos. ¿Qué pasa?",
              options: [
                "La tabla dinámica está dañada, hay que rehacerla",
                "La tabla dinámica no se refresca sola: hay que Actualizar (Alt+F5)",
                "Falta guardar el archivo",
                "El origen tiene que estar en otra hoja"
              ],
              answer: 1,
              explain: "<p>La TD trabaja sobre una copia en caché. Clic derecho › Actualizar, o <b>Alt+F5</b>. Se puede dejar «Actualizar al abrir el archivo».</p>"
            }
          ]
        },
        {
          h: "Gráficos dinámicos",
          body:
            "<p>Un <b>gráfico dinámico</b> está pegado a la TD: si filtrás la tabla, el gráfico sigue.</p>" +
            "<ul>" +
            "<li><b>Columnas</b> para comparar puntos de venta entre sí.</li>" +
            "<li><b>Líneas</b> para la evolución mes a mes.</li>" +
            "<li><b>Torta</b> solo si son pocas categorías y querés mostrar participación (y ni así es la mejor).</li>" +
            "</ul>" +
            "<p>Reglas de presentación: título que diga la conclusión, ordenar las barras de mayor a menor, sacar líneas de grilla de más, no usar 3D.</p>"
        }
      ],
      exercises: [
        {
          id: "m3e1", type: "order",
          q: "Ordená los pasos para armar el reporte mensual de ventas por punto de venta.",
          steps: [
            "Consolidar la venta en una tabla plana: una fila = una venta",
            "Convertir el rango en Tabla (Ctrl+T)",
            "Insertar › Tabla dinámica",
            "Punto de venta a Filas, Mes a Columnas (agrupando la fecha), Importe a Valores",
            "Agregar una 2ª métrica: Importe mostrado como «% del total general»",
            "Insertar gráfico de columnas, ordenarlo de mayor a menor y titularlo con la conclusión"
          ],
          explain: "<p>Primero los datos limpios, después la Tabla, después la TD, después la presentación. El orden importa: sin datos planos la TD no arranca.</p>"
        },
        {
          id: "m3e2", type: "formula",
          q: "Sin tabla dinámica: en H2, total vendido por «Eventos» (col. B = punto de venta, col. C = importe).",
          grid: {
            cols: ["Fecha", "Punto de venta", "Importe"],
            rows: [
              ["03/08", "Tienda", 12000], ["03/08", "Eventos", 45000],
              ["04/08", "Eventos", 30000], ["05/08", "Stands móviles", 8000]
            ],
            money: [2]
          },
          cell: "H2", model: "=SUMAR.SI(B2:B5;\"Eventos\";C2:C5)",
          explain: "<p><code>=SUMAR.SI(B2:B5;\"Eventos\";C2:C5)</code> → 75.000. Esto es lo que hace por dentro una TD con Punto de venta en Filas.</p>"
        },
        {
          id: "m3e3", type: "formula",
          q: "% de participación de la Tienda. En H3, dividí el total de la «Tienda» por el total general de la columna Importe.",
          grid: {
            cols: ["Fecha", "Punto de venta", "Importe"],
            rows: [
              ["03/08", "Tienda", 30000], ["03/08", "Eventos", 45000],
              ["04/08", "Tienda", 10000], ["05/08", "Stands móviles", 15000]
            ],
            money: [2]
          },
          cell: "H3",
          model: "=SUMAR.SI(B2:B5;\"Tienda\";C2:C5)/SUMA(C2:C5)",
          fmt: "pct",
          explain: "<p>40.000 / 100.000 = <b>40%</b>. En una TD lo hace «Mostrar valores como › % del total general».</p>"
        },
        {
          id: "m3e4", type: "classify",
          q: "¿Conviene tabla dinámica o fórmula (SUMAR.SI.CONJUNTO)?",
          cats: ["Mejor tabla dinámica", "Mejor fórmula"],
          items: [
            { t: "Explorar rápido las ventas probando varios cortes", c: 0 },
            { t: "Un reporte con forma fija que se entrega igual todos los meses", c: 1 },
            { t: "Probar «¿y si agrupo por categoría en vez de por punto de venta?»", c: 0 },
            { t: "Un número puntual que va dentro de otra planilla de cálculo", c: 1 }
          ],
          explain: "<p>La TD es para <b>explorar</b>; la fórmula para el <b>entregable estándar</b> que se recalcula solo. En la práctica se usan las dos.</p>"
        },
        {
          id: "m3e5", type: "mc",
          q: "Te piden «lo mismo pero solo Libros y solo del último trimestre». ¿Cómo lo resolvés sin rehacer la tabla dinámica?",
          options: [
            "Rehago la tabla dinámica desde cero con esos filtros",
            "Inserto una segmentación por Categoría y una escala de tiempo por Fecha, y filtro ahí",
            "Borro las filas que no son Libros del origen",
            "Cambio el origen de datos"
          ],
          answer: 1,
          explain: "<p>La segmentación y la escala de tiempo filtran la TD y el gráfico juntos, sin tocar los datos. Si te lo van a pedir seguido, las dejás fijas arriba del reporte.</p>"
        }
      ],
      checklist: [
        "Creo una TD desde cero y explico las cuatro áreas (Filas, Columnas, Valores, Filtros)",
        "Cambio de Suma a Recuento y muestro un valor como % del total general",
        "Agrupo fechas por mes y sé por qué a veces no deja (texto)",
        "Sé que hay que Actualizar (Alt+F5) y por qué la TD no se refresca sola",
        "Describo paso a paso cómo armaría el reporte mensual de ventas por punto de venta"
      ],
      keys: [
        { k: "4 áreas", v: "Filas / Columnas / Valores / Filtros. Categorías van a Filas o Columnas, números a Valores." },
        { k: "Origen = Tabla", v: "Ctrl+T antes de crear la TD → crece sola al agregar filas." },
        { k: "Mostrar valores como", v: "% del total general, % de la columna, Diferencia respecto del mes anterior." },
        { k: "Agrupar fechas", v: "Clic derecho sobre una fecha › Agrupar › Meses/Trim/Años." },
        { k: "Actualizar", v: "Alt+F5. La TD NO se actualiza sola cuando cambian los datos." },
        { k: "Segmentación", v: "Filtro visual; con «Conexiones de informe» controla varias TD a la vez." }
      ]
    },

    /* ===================== MÓDULO 4 ===================== */
    {
      id: 4,
      title: "Control de stock y consignación",
      short: "Inventario en Excel, punto de reposición, rotación y venta por consignación",
      goal: "Controlar el stock de la Tienda y explicar con claridad la venta por consignación.",
      lessons: [
        {
          h: "La ecuación del inventario",
          body:
            "<p>Todo control de stock se apoya en una sola identidad:</p>" +
            "<p style='text-align:center'><code>Stock final = Stock inicial + Entradas − Salidas</code></p>" +
            "<p>Las <b>entradas</b> son compras / remitos del proveedor / devoluciones de clientes. Las <b>salidas</b> son ventas + ajustes por merma, rotura o robo.</p>" +
            "<p>En Excel se lleva de dos formas:</p>" +
            "<ul>" +
            "<li><b>Foto por producto:</b> una fila por producto con columnas Stock inicial, Entradas, Ventas, Ajustes, Stock final = <code>=C2+D2-E2-F2</code>.</li>" +
            "<li><b>Kardex / movimientos:</b> una fila por movimiento (fecha, tipo, cantidad + o −) y una columna de <b>saldo acumulado</b>. Da trazabilidad total.</li>" +
            "</ul>",
          checks: [
            {
              id: "m4c1", type: "formula",
              q: "En F2 calculá el stock final: inicial (B) + entradas (C) − ventas (D) − ajustes/merma (E).",
              grid: {
                cols: ["Producto", "Inicial", "Entradas", "Ventas", "Ajustes", "Final"],
                rows: [["Llavero", 40, 24, 31, 2, ""], ["Taza esmaltada", 15, 0, 6, 0, ""]]
              },
              cell: "F2", model: "=B2+C2-D2-E2",
              explain: "<p><code>=B2+C2-D2-E2</code> → 40 + 24 − 31 − 2 = <b>31</b>. Los ajustes por merma/rotura también restan.</p>"
            }
          ]
        },
        {
          h: "La ficha de producto",
          body:
            "<p>Datos mínimos por artículo de la Tienda:</p>" +
            "<ul>" +
            "<li><b>Código</b> (único, es la clave de todos los BUSCARX), <b>Nombre</b>, <b>Categoría</b> (Libros, Indumentaria, Souvenirs, Discos…).</li>" +
            "<li><b>Proveedor</b> y <b>Tipo de relación</b>: <i>compra en firme</i> o <i>consignación</i>.</li>" +
            "<li><b>Costo</b> (o precio «a liquidar» si es consignación), <b>PVP</b> (precio de venta al público), <b>% de comisión</b> si aplica.</li>" +
            "<li><b>Stock actual</b>, <b>Punto de reposición (ROP)</b>, <b>Lote de reposición</b> (cuánto pedir), <b>Lead time</b> (días que tarda el proveedor).</li>" +
            "</ul>"
        },
        {
          h: "Punto de reposición y stock de seguridad",
          body:
            "<p>El <b>punto de reposición (ROP)</b> es el nivel de stock que dispara «hay que pedir de nuevo».</p>" +
            "<p style='text-align:center'><code>ROP = demanda diaria promedio × lead time (días) + stock de seguridad</code></p>" +
            "<p>El <b>stock de seguridad</b> es un colchón para cubrir picos de demanda o demoras del proveedor. En una tienda de museo hay <b>estacionalidad</b> fuerte (temporada de ópera y ballet, vacaciones, turismo), así que la demanda diaria conviene calcularla por temporada, no anual.</p>" +
            "<p>Un <b>quiebre de stock</b> (stockout) es quedarse sin producto: cuesta una <b>venta perdida</b> y, en un local premium, imagen.</p>",
          checks: [
            {
              id: "m4c2", type: "predict",
              q: "Un producto vende en promedio 5 unidades por día, el proveedor tarda 7 días en reponer y se quiere un colchón de 10 unidades. En C2: ROP.",
              grid: {
                cols: ["Demanda diaria", "Lead time (días)", "Stock seguridad", "ROP"],
                rows: [[5, 7, 10, ""]]
              },
              cell: "D2", formula: "=A2*B2+C2",
              explain: "<p>ROP = 5 × 7 + 10 = <b>45</b>. Cuando el stock baja a 45, hay que emitir la orden de compra para no quebrar durante los 7 días de espera.</p>"
            }
          ]
        },
        {
          h: "Alertas de reposición en Excel",
          body:
            "<p>Dos maneras, se suelen usar juntas:</p>" +
            "<ul>" +
            "<li><b>Columna de estado:</b> <code>=SI(Stock&lt;=ROP; \"REPONER\"; \"OK\")</code>. Contás cuántos hay que reponer con <code>=CONTAR.SI(Estado; \"REPONER\")</code>.</li>" +
            "<li><b>Formato condicional</b> con fórmula <code>=$E2&lt;=$F2</code> para pintar toda la fila de rojo.</li>" +
            "</ul>" +
            "<p>Para armar la <b>orden de compra</b>: filtrás por «REPONER», y con <code>BUSCARX</code> traés Proveedor, contacto y lote de reposición. Cantidad a pedir = <code>=MAX(LoteReposición; ROP − Stock)</code>.</p>",
          sheet: {
            caption: "Hoja: Stock Tienda — control de reposición",
            cols: ["Código", "Producto", "Stock", "ROP", "Estado"],
            rows: [
              ["LLV-001", "Llavero", "42", "20", "OK"],
              ["LIB-014", "Libro «1908»", "3", "8", "REPONER"],
              ["IND-022", "Remera", "6", "10", "REPONER"],
              ["DIS-005", "Disco «Filarmónica»", "25", "12", "OK"]
            ],
            highlight: { row: 3, col: 5 }
          },
          checks: [
            {
              id: "m4c3", type: "formula",
              q: "En D2 escribí la fórmula del semáforo: si el stock (B) es menor o igual al ROP (C), «REPONER»; si no, «OK».",
              grid: {
                cols: ["Producto", "Stock", "ROP", "Estado"],
                rows: [["Libro «1908»", 3, 8, ""], ["Llavero", 42, 20, ""]]
              },
              cell: "D2", model: "=SI(B2<=C2;\"REPONER\";\"OK\")",
              explain: "<p><code>=SI(B2&lt;=C2;\"REPONER\";\"OK\")</code>. Para la fila 2 da «REPONER» (3 ≤ 8). El mismo criterio se usa en el formato condicional con <code>=$B2&lt;=$C2</code>.</p>"
            }
          ]
        },
        {
          h: "Rotación de stock",
          body:
            "<p>La <b>rotación</b> mide cuántas veces «se da vuelta» el inventario en un período.</p>" +
            "<p style='text-align:center'><code>Rotación = Costo de la mercadería vendida (CMV) / Stock promedio a costo</code></p>" +
            "<p>También en unidades: <code>unidades vendidas / stock promedio en unidades</code>.</p>" +
            "<p><b>Días de inventario</b> = 365 / rotación → cuántos días de venta tenés parados en el depósito.</p>" +
            "<p>Rotación <b>alta</b> = se vende rápido, poco capital inmovilizado. Rotación <b>baja</b> (típico de souvenirs de nicho, libros caros): mucho capital quieto y riesgo de que quede viejo. Si un producto rota poco: revisar precio, ubicación en la vitrina, hacer combo/promo, reducir la compra, o pasarlo a <b>consignación</b> para no inmovilizar plata.</p>",
          checks: [
            {
              id: "m4c4", type: "predict",
              q: "El CMV del año fue $ 1.200.000 y el stock promedio a costo fue $ 300.000. En C2: rotación anual.",
              grid: { cols: ["CMV", "Stock promedio", "Rotación"], rows: [[1200000, 300000, ""]], money: [0, 1] },
              cell: "C2", formula: "=A2/B2",
              explain: "<p>Rotación = 1.200.000 / 300.000 = <b>4</b> veces al año. Días de inventario = 365 / 4 ≈ 91 días parados en depósito.</p>"
            }
          ]
        },
        {
          h: "Inventario físico vs. teórico",
          body:
            "<p>El <b>stock teórico</b> es el que dice tu planilla. El <b>stock físico</b> es el que contás en el mostrador/depósito. Casi nunca coinciden (roturas, robo hormiga, errores de carga, unidades de regalo).</p>" +
            "<p>Proceso: se hace un <b>conteo</b> (total una vez al año o <b>cíclico</b> por categoría), se calcula la <b>diferencia</b> = físico − teórico, se investiga lo relevante y se hace el <b>ajuste</b> en la planilla dejando el motivo. Ese ajuste es una «salida» (o entrada) en la ecuación del inventario.</p>"
        },
        {
          h: "Venta por consignación — el concepto clave del puesto",
          body:
            "<p>En la <b>venta por consignación</b>, el <b>proveedor (consignante)</b> entrega mercadería a la Tienda (<b>consignatario</b>) para que la exhiba y venda, <b>pero la propiedad NO se transfiere</b> hasta que la Tienda la vende al cliente final. Lo que no se vende, <b>se devuelve</b>.</p>" +
            "<p><b>Por qué la Tienda lo usa:</b></p>" +
            "<ul>" +
            "<li>Amplía el surtido <b>sin inmovilizar capital</b> ni asumir el riesgo de que el producto quede sin vender.</li>" +
            "<li>Solo paga lo que <b>efectivamente vendió</b>.</li>" +
            "<li>Al artista / editorial / proveedor le sirve por la <b>exposición</b> en un canal premium.</li>" +
            "</ul>" +
            "<p><b>Diferencia con la compra en firme:</b> en la compra en firme la Tienda <b>compra</b> el stock, lo <b>paga por adelantado</b> y asume el riesgo de rotación y obsolescencia. En consignación, no.</p>" +
            "<p><b>Cómo gana la Tienda:</b> retiene una <b>comisión</b> (un % del PVP) y le <b>liquida</b> el resto al consignante. A veces al revés: el consignante fija un precio «a liquidar» y la Tienda vende al PVP que quiera y se queda con la diferencia.</p>",
          checks: [
            {
              id: "m4c5", type: "mc",
              q: "En la venta por consignación, ¿cuándo se transfiere la propiedad de la mercadería?",
              options: [
                "Cuando el proveedor entrega la mercadería a la Tienda",
                "Cuando la Tienda la vende al cliente final",
                "Al firmar el convenio",
                "A los 30 días de la entrega"
              ],
              answer: 1,
              explain: "<p>La propiedad pasa recién con la venta al cliente final. Lo que no se vende se devuelve al consignante. Por eso el stock consignado no es un activo de la Tienda.</p>"
            },
            {
              id: "m4c6", type: "classify",
              q: "¿Compra en firme o consignación?",
              cats: ["Compra en firme", "Consignación"],
              items: [
                { t: "La Tienda paga el stock por adelantado", c: 0 },
                { t: "La Tienda paga solo lo que efectivamente vendió", c: 1 },
                { t: "Lo no vendido se devuelve al proveedor", c: 1 },
                { t: "La Tienda asume el riesgo de que el producto no rote", c: 0 }
              ],
              explain: "<p>La consignación le permite a la Tienda ampliar el surtido sin inmovilizar capital ni asumir riesgo de obsolescencia.</p>"
            }
          ]
        },
        {
          h: "Controlar el stock en consignación en Excel",
          body:
            "<ul>" +
            "<li><b>Separado del stock propio:</b> una columna «Tipo» (Firme / Consignación) o directamente una hoja aparte. Nunca mezclar, porque lo consignado no es un activo de la Tienda.</li>" +
            "<li>Por cada <b>entrega</b> del consignante: remito, fecha, cantidad, PVP y % de comisión pactado.</li>" +
            "<li>Por cada <b>venta</b>: registrar el código, que ya trae asociado el consignante.</li>" +
            "<li><b>Conciliación por consignante y período:</b> <code>stock inicial consignado + entregas − ventas − devoluciones = stock final consignado</code>. Ese cuadre es lo que después respalda la liquidación (Módulo 5).</li>" +
            "<li><b>Sell-through</b> = unidades vendidas / unidades recibidas del consignante en el período. Te dice si al consignante le conviene seguir y cuánto reponer.</li>" +
            "</ul>",
          checks: [
            {
              id: "m4c7", type: "predict",
              q: "Un consignante tenía 5 remeras del mes anterior, entregó 40, se vendieron 28 y se devolvieron 2 por falla. En E2: sell-through del mes.",
              grid: {
                cols: ["Inicial", "Entregas", "Ventas", "Devol.", "Sell-through"],
                rows: [[5, 40, 28, 2, ""]]
              },
              cell: "E2", formula: "=C2/(A2+B2)", fmt: "pct",
              explain: "<p>Sell-through = vendidas / disponibles = 28 / (5 + 40) ≈ <b>62%</b>. (El stock final consignado sería 5 + 40 − 28 − 2 = 15.)</p>"
            }
          ]
        }
      ],
      exercises: [
        {
          id: "m4e1", type: "formula",
          q: "Stock final. En F2, calculá inicial + entradas − ventas − ajustes.",
          grid: {
            cols: ["Código", "Inicial", "Entradas", "Ventas", "Ajustes", "Final"],
            rows: [["LLV-001", 40, 24, 31, 2, ""], ["LIB-014", 12, 6, 5, 0, ""], ["IND-022", 20, 0, 14, 1, ""]]
          },
          cell: "F2", model: "=B2+C2-D2-E2",
          explain: "<p><code>=B2+C2-D2-E2</code> → 31. Las salidas incluyen ventas y ajustes por merma/rotura.</p>"
        },
        {
          id: "m4e2", type: "formula",
          q: "Estado de reposición ya calculado en la columna D. En G2, contá cuántos productos hay que REPONER.",
          grid: {
            cols: ["Código", "Stock", "ROP", "Estado", "·", "·", "A reponer"],
            rows: [
              ["LLV-001", 42, 20, "=SI(B2<=C2;\"REPONER\";\"OK\")", "", "", ""],
              ["LIB-014", 3, 8, "=SI(B3<=C3;\"REPONER\";\"OK\")", "", "", ""],
              ["IND-022", 6, 10, "=SI(B4<=C4;\"REPONER\";\"OK\")", "", "", ""],
              ["DIS-005", 25, 12, "=SI(B5<=C5;\"REPONER\";\"OK\")", "", "", ""]
            ]
          },
          cell: "G2", model: "=CONTAR.SI(D2:D5;\"REPONER\")",
          explain: "<p><code>=CONTAR.SI(D2:D5;\"REPONER\")</code> → 2. La columna Estado ya trae el <code>=SI(...)</code> de cada fila; acá solo contás.</p>"
        },
        {
          id: "m4e3", type: "formula",
          q: "Stock final consignado. En F2, conciliá: inicial + entregas − ventas − devoluciones.",
          grid: {
            cols: ["Consignante", "Inicial", "Entregas", "Ventas", "Devol.", "Final consig."],
            rows: [["Artesanías Rivadavia", 5, 40, 28, 2, ""]]
          },
          cell: "F2", model: "=B2+C2-D2-E2",
          explain: "<p>5 + 40 − 28 − 2 = <b>15</b>. Este número tiene que coincidir con el conteo físico; es lo que respalda la liquidación del mes.</p>"
        },
        {
          id: "m4e4", type: "order",
          q: "Ordená la conciliación del stock en consignación de un período.",
          steps: [
            "Stock inicial consignado (lo que quedó del período anterior)",
            "+ Entregas del consignante en el período (por remito)",
            "− Unidades vendidas al cliente final",
            "− Devoluciones al consignante",
            "= Stock final consignado (tiene que coincidir con el conteo físico)"
          ],
          explain: "<p>Si el stock final no cuadra con el conteo, puede haber ventas sin registrar o unidades traspapeladas — y eso afecta la liquidación.</p>"
        },
        {
          id: "m4e5", type: "mc",
          q: "Un producto tiene rotación baja (rota 1 vez al año). ¿Qué implica y qué harías?",
          options: [
            "Se vende rapidísimo; hay que comprar mucho más",
            "Mucho capital inmovilizado y riesgo de que quede viejo: revisar precio/ubicación, promo, comprar menos, o pasarlo a consignación",
            "Que el margen es negativo",
            "Que hubo un quiebre de stock"
          ],
          answer: 1,
          explain: "<p>Rotación baja = por cada unidad que vendés, tenés muchas paradas. Típico en souvenirs de nicho y libros caros. Pasarlo a consignación evita inmovilizar plata.</p>"
        }
      ],
      checklist: [
        "Escribo la ecuación del stock final y explico qué entra en «entradas» y «salidas»",
        "Sé qué es el punto de reposición y armo la alerta en Excel (SI + formato condicional)",
        "Explico qué es la venta por consignación y en qué se diferencia de la compra en firme",
        "Sé qué es la rotación de stock y qué haría con un producto que rota poco",
        "Explico por qué el stock en consignación va separado y cómo se concilia por período"
      ],
      keys: [
        { k: "Stock final", v: "= Stock inicial + Entradas − Salidas (ventas + ajustes)." },
        { k: "ROP", v: "Punto de reposición = demanda diaria × lead time + stock de seguridad." },
        { k: "Rotación", v: "CMV / stock promedio a costo. Días de inventario = 365 / rotación." },
        { k: "Consignación", v: "El proveedor entrega mercadería; la propiedad pasa recién al venderse. Lo no vendido se devuelve." },
        { k: "En firme vs. consignación", v: "En firme: la Tienda compra y asume el riesgo. Consignación: paga solo lo vendido." },
        { k: "Sell-through", v: "Unidades vendidas / unidades recibidas en el período." },
        { k: "Conciliación", v: "Stock inicial consignado + entregas − ventas − devoluciones = stock final." }
      ]
    },

    /* ===================== MÓDULO 5 ===================== */
    {
      id: 5,
      title: "Liquidación a proveedores",
      short: "El cálculo mensual de lo que se le paga a cada proveedor bajo convenio",
      goal: "Armar y controlar la liquidación mensual a proveedores en consignación / bajo convenio.",
      lessons: [
        {
          h: "Qué es una liquidación",
          body:
            "<p>La <b>liquidación</b> es el cálculo y el documento <b>periódico</b> (normalmente <b>mensual</b>) que determina <b>cuánto le corresponde cobrar</b> a cada proveedor que trabaja <b>bajo convenio</b> (consignación o condición especial), según lo que <b>efectivamente se vendió</b> de sus productos en el período.</p>" +
            "<p>Es la contracara de la consignación: como la Tienda no compró el stock por adelantado, al cierre de cada mes hay que decirle a cada consignante «esto se vendió tuyo, esta es mi comisión, esto te toca a vos».</p>" +
            "<p>Con ese número, el proveedor <b>emite su factura</b> por el neto y la Tienda la paga según el <b>plazo del convenio</b> (por ejemplo 30 días fecha de factura).</p>",
          checks: [
            {
              id: "m5c1", type: "mc",
              q: "¿Qué es la liquidación mensual a proveedores?",
              options: [
                "El pago de sueldos del área comercial",
                "El cálculo de cuánto le corresponde cobrar a cada proveedor bajo convenio según lo efectivamente vendido en el período",
                "El cierre de caja diario de la Tienda",
                "El inventario físico de fin de mes"
              ],
              answer: 1,
              explain: "<p>Es la contracara de la consignación: al no comprar el stock por adelantado, cada mes se calcula lo vendido de cada proveedor y su neto a cobrar.</p>"
            }
          ]
        },
        {
          h: "Qué información debe contener",
          body:
            "<ul>" +
            "<li><b>Encabezado:</b> proveedor, CUIT / condición fiscal, período liquidado, número de liquidación, fecha de corte.</li>" +
            "<li><b>Detalle por producto:</b> código, descripción, unidades vendidas, PVP, <b>importe bruto vendido</b>.</li>" +
            "<li><b>Comisión de la Tienda:</b> % pactado y monto (o el margen retenido, según cómo esté el convenio).</li>" +
            "<li><b>Deducciones:</b> devoluciones / notas de crédito, mermas o roturas si el convenio las pone a cargo del proveedor, descuentos autorizados.</li>" +
            "<li><b>Neto a liquidar</b> al proveedor = bruto − comisión − deducciones.</li>" +
            "<li><b>Impuestos y retenciones</b> según la condición fiscal de cada parte (IVA, retención de Ganancias / IIBB si corresponde).</li>" +
            "<li><b>Conciliación de stock consignado:</b> stock inicial + entregas − ventas − devoluciones = stock final (informativo, respalda que las unidades vendidas están bien).</li>" +
            "</ul>",
          checks: [
            {
              id: "m5c2", type: "classify",
              q: "¿Forma parte de la liquidación al proveedor?",
              cats: ["Sí, va en la liquidación", "No corresponde"],
              items: [
                { t: "Detalle de unidades vendidas por producto y PVP", c: 0 },
                { t: "Comisión de la Tienda (% y monto)", c: 0 },
                { t: "Devoluciones / notas de crédito del período", c: 0 },
                { t: "Conciliación del stock consignado", c: 0 },
                { t: "El sueldo del vendedor de la Tienda", c: 1 }
              ],
              explain: "<p>La liquidación cubre todo lo necesario para que el proveedor entienda su neto y pueda facturar. Los costos internos de la Tienda no van ahí.</p>"
            }
          ]
        },
        {
          h: "Cómo se arma en Excel — la base",
          body:
            "<p>Punto de partida: la hoja de <b>ventas del período</b> con Fecha, Código, Producto, Proveedor, Unidades, Importe (PVP × unidades).</p>" +
            "<p>Si la venta no trae el Proveedor, lo traés con <code>=BUSCARX(Código; Productos[Código]; Productos[Proveedor])</code>. Lo mismo el <b>% de comisión</b>, que puede estar por proveedor o por convenio.</p>" +
            "<p>Una <b>tabla dinámica</b> te da el resumen en 30 segundos: Filas = Proveedor › Producto, Valores = Suma de Unidades y Suma de Importe. Filtro / escala de tiempo por el mes que estás liquidando.</p>"
        },
        {
          h: "Cómo se arma en Excel — los cálculos",
          body:
            "<p>Sobre el resumen por proveedor:</p>" +
            "<ul>" +
            "<li><b>Bruto vendido</b> = <code>=SUMAR.SI.CONJUNTO(Importe; Proveedor; A2; Fecha; \"&gt;=\"&amp;$F$1; Fecha; \"&lt;=\"&amp;$G$1)</code> (F1/G1 = desde/hasta del período).</li>" +
            "<li><b>% comisión</b> = <code>=BUSCARX(A2; Convenios[Proveedor]; Convenios[Comision])</code>.</li>" +
            "<li><b>Comisión $</b> = <code>=Bruto * %comisión</code>.</li>" +
            "<li><b>Devoluciones</b> = <code>=SUMAR.SI.CONJUNTO(...)</code> sobre la hoja de NC del período.</li>" +
            "<li><b>Neto a liquidar</b> = <code>=Bruto − Comisión − Devoluciones</code>.</li>" +
            "</ul>" +
            "<p>Estructura del archivo: una hoja <b>«Resumen»</b> con una fila por proveedor (para la mirada de gestión) + una hoja de <b>detalle por proveedor</b> (lo que se le manda a cada uno), o un único detalle filtrable con una <b>segmentación</b> por proveedor.</p>",
          sheet: {
            caption: "Resumen de liquidación — agosto 2025",
            cols: ["Proveedor", "Bruto", "Comis. 30%", "Devol.", "Neto a pagar"],
            rows: [
              ["Editorial Central", "1.240.000", "372.000", "0", "868.000"],
              ["Artesanías Rivadavia", "486.000", "145.800", "12.000", "328.200"],
              ["Discográfica Filarmónica", "930.000", "279.000", "0", "651.000"],
              ["TOTAL", "2.656.000", "796.800", "12.000", "1.847.200"]
            ],
            highlight: { row: 5, col: 5 }
          },
          checks: [
            {
              id: "m5c3", type: "formula",
              q: "En E2, calculá el neto a liquidar: bruto (B) − comisión $ (C) − devoluciones (D).",
              grid: {
                cols: ["Proveedor", "Bruto", "Comisión $", "Devol.", "Neto"],
                rows: [["Artesanías Rivadavia", 486000, 145800, 12000, ""], ["Editorial Central", 1240000, 372000, 0, ""]],
                money: [1, 2, 3, 4]
              },
              cell: "E2", model: "=B2-C2-D2",
              explain: "<p>486.000 − 145.800 − 12.000 = <b>328.200</b>. Con eso el proveedor factura y se le paga según el plazo del convenio.</p>"
            }
          ]
        },
        {
          h: "Controles antes de enviar",
          body:
            "<p>Nunca mandes una liquidación sin estos chequeos:</p>" +
            "<ul>" +
            "<li><b>Cierre aritmético:</b> Σ Netos + Σ Comisiones + Σ Deducciones = Bruto total vendido del período. Si no cierra, hay una fila mal.</li>" +
            "<li><b>Productos sin proveedor:</b> filtrá la columna de proveedor por <b>#N/A</b> o vacío. Cada venta tiene que poder asignarse a alguien.</li>" +
            "<li><b>Cuadre de stock consignado</b> por proveedor (la conciliación del Módulo 4).</li>" +
            "<li><b>Comisión aplicada correcta:</b> que el % sea el del convenio vigente de ese proveedor (ojo con los que cambiaron de condición a mitad de mes).</li>" +
            "<li><b>Comparación contra el mes anterior:</b> un salto raro (±50%) casi siempre es un error de datos, no una explosión de ventas.</li>" +
            "</ul>",
          checks: [
            {
              id: "m5c4", type: "formula",
              q: "Control de cierre. En C5, escribí la verificación: suma de la columna Neto MÁS suma de la columna Comisión $ (debería dar el bruto total).",
              grid: {
                cols: ["Proveedor", "Neto", "Comisión $", "Bruto"],
                rows: [
                  ["Editorial", 868000, 372000, 1240000],
                  ["Artesanías", 340200, 145800, 486000],
                  ["Discográfica", 651000, 279000, 930000],
                  ["Verificación →", "", "", ""]
                ],
                money: [1, 2, 3]
              },
              cell: "C5", model: "=SUMA(B2:B4)+SUMA(C2:C4)",
              explain: "<p>Σ netos + Σ comisiones = 1.859.200 + 796.800 = 2.656.000, igual a <code>=SUMA(D2:D4)</code>. (En este caso hay $12.000 de devoluciones que explican la diferencia con el bruto — el control real incluye también las deducciones.)</p>"
            }
          ]
        },
        {
          h: "Buenas prácticas y trazabilidad",
          body:
            "<ul>" +
            "<li><b>Plantilla estandarizada</b>: misma estructura todos los meses, así se puede comparar y automatizar.</li>" +
            "<li><b>Versionado por mes</b>: <code>Liquidacion_2025-08.xlsx</code>. No pisar el archivo del mes anterior.</li>" +
            "<li><b>Congelar los datos</b>: pegá como valores la venta del período liquidado, así una carga tardía no te cambia una liquidación ya enviada.</li>" +
            "<li><b>Trazabilidad</b>: dejá registrado qué rango de comprobantes / fechas entró en cada liquidación.</li>" +
            "<li><b>Oportunidad de mejora</b> (para proponer en la entrevista): hoy suele hacerse a mano; se puede armar una plantilla con Power Query que tome la venta exportada del sistema y devuelva las liquidaciones por proveedor casi sin intervención, y un tablero en Looker Studio para que cada área vea el avance.</li>" +
            "</ul>"
        }
      ],
      exercises: [
        {
          id: "m5e1", type: "formula",
          q: "Bruto vendido por proveedor. En H2, sumá el Importe (col. C) donde el Proveedor (col. B) sea «Editorial Central».",
          grid: {
            cols: ["Fecha", "Proveedor", "Importe"],
            rows: [
              ["03/08", "Editorial Central", 620000],
              ["05/08", "Artesanías Rivadavia", 180000],
              ["12/08", "Editorial Central", 620000],
              ["20/08", "Discográfica Filarmónica", 930000]
            ],
            money: [2]
          },
          cell: "H2", model: "=SUMAR.SI(B2:B5;\"Editorial Central\";C2:C5)",
          explain: "<p><code>=SUMAR.SI(B2:B5;\"Editorial Central\";C2:C5)</code> → 1.240.000.</p>"
        },
        {
          id: "m5e2", type: "formula",
          q: "Neto con la comisión traída por BUSCARX. En C2, calculá bruto (B) menos la comisión, usando el % del convenio (columnas E y F).",
          grid: {
            cols: ["Proveedor", "Bruto", "Neto", "·", "Prov. (convenio)", "Comisión"],
            rows: [
              ["Editorial Central", 1240000, "", "", "Editorial Central", 0.30],
              ["Artesanías Rivadavia", 486000, "", "", "Artesanías Rivadavia", 0.30],
              ["Discográfica Filarmónica", 930000, "", "", "Discográfica Filarmónica", 0.25]
            ],
            money: [1, 2]
          },
          cell: "C2",
          model: "=B2-B2*BUSCARX(A2;E2:E4;F2:F4)",
          explain: "<p><code>=B2-B2*BUSCARX(A2;E2:E4;F2:F4)</code> → 1.240.000 − 30% = <b>868.000</b>. La comisión sale del convenio de cada proveedor, no de un número fijo.</p>"
        },
        {
          id: "m5e3", type: "predict",
          q: "Un consignante vendió $ 486.000 en el mes, tuvo $ 12.000 de devoluciones y la comisión de la Tienda es 30%. En E2: neto a liquidar.",
          grid: {
            cols: ["Bruto", "Comisión %", "Devol.", "·", "Neto"],
            rows: [[486000, 0.30, 12000, "", ""]],
            money: [0, 2, 4]
          },
          cell: "E2", formula: "=A2-A2*B2-C2",
          explain: "<p>Comisión = 486.000 × 30% = 145.800. Neto = 486.000 − 145.800 − 12.000 = <b>328.200</b>. Control: 328.200 + 145.800 + 12.000 = 486.000 ✓.</p>"
        },
        {
          id: "m5e4", type: "order",
          q: "Ordená los pasos para armar la liquidación mensual.",
          steps: [
            "Partir de la venta del período (Código, Proveedor, Unidades, Importe)",
            "Traer Proveedor y % de comisión con BUSCARX si faltan",
            "Calcular el bruto por proveedor (SUMAR.SI.CONJUNTO o tabla dinámica)",
            "Comisión = bruto × %; Neto = bruto − comisión − deducciones",
            "Control: Σ netos + Σ comisiones + Σ deducciones = bruto total; sin proveedores #N/A",
            "Congelar los datos (pegar como valores) y versionar el archivo por mes"
          ],
          explain: "<p>El control aritmético y el congelado de datos son lo que evita que una liquidación ya enviada cambie sola cuando entra una carga tardía.</p>"
        },
        {
          id: "m5e5", type: "mc",
          q: "Antes de enviar una liquidación filtrás la columna «Proveedor» por #N/A / vacío. ¿Para qué?",
          options: [
            "Para borrar esos productos",
            "Para detectar ventas sin proveedor asignado, que hacen que la liquidación no cuadre",
            "Para ordenar alfabéticamente",
            "Para aplicar SI.ERROR y seguir"
          ],
          answer: 1,
          explain: "<p>Cada venta tiene que poder asignarse a un proveedor. Un código sin dar de alta rompe el cierre aritmético: por eso se corrige, no se tapa.</p>"
        }
      ],
      checklist: [
        "Explico qué es una liquidación a proveedores y cada cuánto se hace",
        "Enumero los datos mínimos que lleva (detalle, comisión, deducciones, neto)",
        "Digo qué funciones usaría para armarla (SUMAR.SI.CONJUNTO, BUSCARX, tabla dinámica)",
        "Nombro al menos tres controles para asegurarme de que «cierra»",
        "Conecto consignación → conciliación de stock → liquidación"
      ],
      keys: [
        { k: "Liquidación", v: "Cálculo mensual de lo que se le paga a un proveedor bajo convenio, según lo vendido en el período." },
        { k: "Neto a liquidar", v: "Bruto vendido − comisión de la Tienda − deducciones (devoluciones, NC, mermas)." },
        { k: "Contiene", v: "Detalle por producto, unidades, PVP, comisión %, deducciones, neto, impuestos, conciliación de stock." },
        { k: "Control de cierre", v: "Σ netos + Σ comisiones + Σ deducciones = bruto total del período." },
        { k: "Error frecuente", v: "Productos sin proveedor asignado (BUSCARX #N/A) → la liquidación no cuadra." },
        { k: "Periodicidad", v: "Cierre mensual con fecha de corte; el proveedor factura el neto; pago según plazo del convenio." }
      ]
    },

    /* ===================== MÓDULO 6 ===================== */
    {
      id: 6,
      title: "Vocabulario comercial y licitaciones",
      short: "Licitaciones en el sector público, glosario de gestión comercial y Looker",
      goal: "Manejar el vocabulario del rubro y entender cómo compra una institución pública.",
      lessons: [
        {
          h: "Qué es una licitación",
          body:
            "<p>Una <b>licitación</b> es el procedimiento por el cual un organismo selecciona un proveedor o contratista de forma <b>competitiva y transparente</b>: publica un <b>pliego</b> con las condiciones, recibe <b>ofertas</b>, las compara con reglas conocidas de antemano y adjudica.</p>" +
            "<p><b>Por qué licitan los organismos públicos:</b> un organismo del <b>Estado</b> está sujeto a normas de compras. La normativa de compras del Estado (Ley de Compras y Contrataciones) exige <b>transparencia, concurrencia de oferentes e igualdad de trato</b>. Por encima de ciertos montos <b>no se puede comprar «a dedo»</b>: hay que licitar.</p>",
          checks: [
            {
              id: "m6c1", type: "mc",
              q: "¿Por qué un organismo público compra por licitación en vez de elegir proveedor directamente?",
              options: [
                "Porque siempre sale más barato",
                "Porque es un organismo público y la normativa exige transparencia, concurrencia e igualdad entre oferentes",
                "Porque no tiene presupuesto propio",
                "Porque lo pide cada proveedor"
              ],
              answer: 1,
              explain: "<p>Arriba de ciertos montos el sector público no puede comprar «a dedo»: tiene que abrir el proceso y adjudicar con reglas conocidas de antemano.</p>"
            }
          ]
        },
        {
          h: "Tipos de procedimiento",
          body:
            "<ul>" +
            "<li><b>Licitación pública:</b> convocatoria abierta, cualquier oferente que cumpla puede presentarse. Para los <b>montos más altos</b>. Máxima transparencia.</li>" +
            "<li><b>Licitación privada:</b> se <b>invita</b> a un conjunto de proveedores del rubro. Para <b>montos intermedios</b>.</li>" +
            "<li><b>Contratación directa:</b> se contrata a un proveedor puntual. Solo para <b>montos bajos</b> o casos de <b>excepción</b> justificados: urgencia, exclusividad / único proveedor, obra de arte, fracaso de una licitación previa.</li>" +
            "<li><b>Convenio marco / orden de compra abierta:</b> se licita <b>una vez</b> y después se compra contra ese acuerdo durante un plazo, sin licitar cada pedido.</li>" +
            "</ul>",
          checks: [
            {
              id: "m6c2", type: "classify",
              q: "¿Licitación pública, privada o contratación directa?",
              cats: ["Pública", "Privada", "Contratación directa"],
              items: [
                { t: "Monto alto, convocatoria abierta a cualquier oferente", c: 0 },
                { t: "Monto intermedio, se invita a un grupo de proveedores del rubro", c: 1 },
                { t: "Monto bajo, o urgencia, o único proveedor posible", c: 2 }
              ],
              explain: "<p>A mayor monto, más apertura y transparencia exige la normativa. La contratación directa siempre necesita justificar el caso de excepción.</p>"
            }
          ]
        },
        {
          h: "Las piezas de una licitación",
          body:
            "<ul>" +
            "<li><b>Pliego de bases y condiciones (PByC):</b> qué se compra, especificaciones técnicas, cantidades, requisitos del oferente, criterio de evaluación, plazos.</li>" +
            "<li><b>Presupuesto oficial:</b> la estimación del organismo de cuánto debería salir.</li>" +
            "<li><b>Oferta:</b> parte <b>económica</b> (precios) + parte <b>técnica</b> (qué ofrece, muestras, antecedentes).</li>" +
            "<li><b>Acto de apertura:</b> se abren las ofertas en fecha y hora fijadas.</li>" +
            "<li><b>Cuadro comparativo:</b> se ponen todas las ofertas lado a lado (¡Excel!) para compararlas por precio unitario y cumplimiento técnico.</li>" +
            "<li><b>Preadjudicación → adjudicación → orden de compra:</b> se elige al ganador, se notifica y se emite la OC.</li>" +
            "</ul>",
          checks: [
            {
              id: "m6c3", type: "order",
              q: "Ordená las etapas de una licitación.",
              steps: [
                "Pliego de bases y condiciones (PByC)",
                "Presentación de ofertas por los proveedores",
                "Acto de apertura",
                "Cuadro comparativo de ofertas",
                "Preadjudicación y adjudicación",
                "Orden de compra al adjudicado"
              ],
              explain: "<p>El cuadro comparativo (que suele armar el analista comercial en Excel) es el insumo clave para la preadjudicación.</p>"
            }
          ]
        },
        {
          h: "Qué hace un analista comercial con una licitación",
          body:
            "<ul>" +
            "<li>Arma el <b>cuadro comparativo de ofertas</b> en Excel: precios unitarios, totales, condiciones, y una columna de observaciones técnicas.</li>" +
            "<li><b>Proyecta el consumo</b> (cuántas unidades se van a necesitar) para dimensionar la compra.</li>" +
            "<li>Analiza si los precios ofertados están en línea con el <b>presupuesto oficial</b> y con lo que se pagó antes.</li>" +
            "<li>Una vez adjudicado: controla el <b>cumplimiento del proveedor</b> (entregas en fecha, precios facturados = precios adjudicados, saldo de la orden de compra) y <b>reporta desvíos</b>.</li>" +
            "</ul>"
        },
        {
          h: "Glosario de gestión comercial — rentabilidad",
          body:
            "<ul>" +
            "<li><b>Margen bruto / sobre venta</b> = (PV − Costo) / PV. Porción de la venta que queda.</li>" +
            "<li><b>Markup / sobre costo</b> = (PV − Costo) / Costo. Lo que se recarga sobre el costo.</li>" +
            "<li><b>Contribución marginal</b> = Precio − costo variable unitario. Lo que cada unidad aporta para cubrir costos fijos.</li>" +
            "<li><b>Punto de equilibrio (break-even)</b> = Costos fijos / contribución marginal unitaria (en unidades). En $: CF / (1 − CV/PV).</li>" +
            "<li><b>CMV / COGS</b> = costo de la mercadería vendida en el período.</li>" +
            "<li><b>Rentabilidad</b> = resultado / inversión (o / ventas, según se mire).</li>" +
            "</ul>",
          checks: [
            {
              id: "m6c4", type: "predict",
              q: "Costos fijos mensuales $ 500.000. Cada unidad se vende a $ 1.000 y su costo variable es $ 600. En D2: punto de equilibrio en unidades.",
              grid: {
                cols: ["Costos fijos", "Precio", "Costo var.", "Equilibrio (u.)"],
                rows: [[500000, 1000, 600, ""]],
                money: [0, 1, 2]
              },
              cell: "D2", formula: "=A2/(B2-C2)",
              explain: "<p>Contribución marginal por unidad = 1.000 − 600 = 400. Punto de equilibrio = 500.000 / 400 = <b>1.250 unidades</b> por mes para no ganar ni perder.</p>"
            }
          ]
        },
        {
          h: "Glosario — retail y demanda",
          body:
            "<ul>" +
            "<li><b>Ticket promedio</b> = facturación / cantidad de tickets (operaciones).</li>" +
            "<li><b>UPT (unidades por ticket)</b> = unidades vendidas / tickets.</li>" +
            "<li><b>Tasa de conversión</b> = tickets / visitantes (cuánta gente que entra, compra).</li>" +
            "<li><b>Sell-through</b> = unidades vendidas / unidades recibidas en el período. Clave en consignación.</li>" +
            "<li><b>Rotación de stock</b> y <b>días de inventario</b> (Módulo 4).</li>" +
            "<li><b>Quiebre de stock (stockout)</b> = quedarse sin producto → venta perdida.</li>" +
            "<li><b>Estacionalidad</b> = patrón que se repite por época del año (temporada lírica, receso, turismo).</li>" +
            "<li><b>Análisis ABC / Pareto</b> = el ~20% de los productos explica el ~80% de la venta; se gestionan distinto los A, B y C.</li>" +
            "</ul>",
          checks: [
            {
              id: "m6c5", type: "predict",
              q: "En el mes la Tienda facturó $ 4.200.000 en 350 operaciones. En C2: ticket promedio.",
              grid: { cols: ["Facturación", "Tickets", "Ticket promedio"], rows: [[4200000, 350, ""]], money: [0, 2] },
              cell: "C2", formula: "=A2/B2",
              explain: "<p>Ticket promedio = 4.200.000 / 350 = <b>$ 12.000</b>. Sube subiendo precios o vendiendo más unidades por operación (UPT).</p>"
            }
          ]
        },
        {
          h: "Glosario — planificación y control de gestión",
          body:
            "<ul>" +
            "<li><b>KPI</b> = indicador clave de desempeño; una métrica elegida porque mueve el negocio y tiene meta.</li>" +
            "<li><b>Meta / target</b> = valor objetivo del KPI.</li>" +
            "<li><b>Forecast / pronóstico</b> = estimación de lo que va a pasar (ventas, demanda). <b>Error de forecast</b> = real − pronóstico.</li>" +
            "<li><b>Presupuesto (budget)</b> = plan financiero del período. <b>Desvío</b> = real − presupuesto (y su %).</li>" +
            "<li><b>Real vs. presupuesto vs. año anterior</b>: los tres cortes de todo reporte de gestión.</li>" +
            "<li><b>YTD (year to date)</b> = acumulado del año hasta la fecha.</li>" +
            "</ul>"
        },
        {
          h: "Glosario — administración de ventas y fiscal",
          body:
            "<ul>" +
            "<li><b>PVP</b> = precio de venta al público. <b>Precio de lista</b> vs. <b>descuento / bonificación</b>.</li>" +
            "<li><b>IVA</b> = impuesto al valor agregado (alícuota general 21%). <b>Condición frente al IVA</b>: Responsable Inscripto, Monotributo, Exento.</li>" +
            "<li><b>Remito</b> = documento que acompaña la mercadería (no tiene precio necesariamente).</li>" +
            "<li><b>Orden de compra (OC)</b> = documento con el que el comprador formaliza el pedido al proveedor.</li>" +
            "<li><b>Nota de crédito / débito</b> = ajustan una factura ya emitida (devolución, error, bonificación).</li>" +
            "<li><b>Retenciones / percepciones</b> = adelantos de impuestos que se practican al pagar o cobrar.</li>" +
            "<li><b>Consignación, liquidación, comisión, convenio</b> (Módulos 4 y 5).</li>" +
            "</ul>"
        },
        {
          h: "Looker y Looker Studio",
          body:
            "<p><b>Looker Studio</b> (antes Google Data Studio) es una herramienta <b>gratuita</b> de Google para armar <b>dashboards y reportes interactivos</b> conectados a fuentes de datos: Google Sheets, BigQuery, bases SQL, planillas, etc. <b>Looker</b> (a secas) es la plataforma <b>enterprise</b> de Google con un modelo semántico propio (LookML) donde se definen métricas una sola vez para toda la organización.</p>" +
            "<p><b>Para qué sirve:</b> reportes que <b>se actualizan solos</b>, se comparten por link, tienen filtros interactivos y se ven bien. Evita mandar un Excel distinto por mail cada semana.</p>" +
            "<p><b>Diferencia con Excel:</b> Excel es <b>cálculo, modelado y análisis ad hoc</b> — es donde <i>trabajás</i> los datos. Looker es <b>presentación y monitoreo recurrente de KPIs</b> sobre datos ya preparados — es donde <i>se miran</i>. El flujo típico: los datos se limpian y transforman (Excel / Sheets / SQL / Power Query) y se publican en Looker.</p>" +
            "<p><b>El puente conceptual:</b> Looker trabaja con <b>dimensiones</b> (categorías: punto de venta, mes, categoría de producto) y <b>métricas</b> (números: importe, unidades, ticket). Es <b>la misma idea</b> que Filas/Columnas vs. Valores de una tabla dinámica. Si entendés tablas dinámicas, entendés el modelo mental de Looker.</p>" +
            "<p><b>Qué decir en la entrevista:</b> «Manejo la lógica de dimensiones y métricas y de construir un tablero de KPIs. En Looker Studio me pongo al día rápido porque el modelo mental es el de una tabla dinámica conectada a datos vivos; de hecho puedo conectar la venta desde Sheets y tener el reporte del área siempre actualizado.»</p>",
          checks: [
            {
              id: "m6c6", type: "classify",
              q: "En un tablero de ventas: ¿dimensión (con qué desagregás) o métrica (qué medís)?",
              cats: ["Dimensión", "Métrica"],
              items: [
                { t: "Punto de venta", c: 0 },
                { t: "Mes", c: 0 },
                { t: "Categoría de producto", c: 0 },
                { t: "Importe vendido", c: 1 },
                { t: "Ticket promedio", c: 1 }
              ],
              explain: "<p>Es la misma división que Filas/Columnas vs. Valores en una tabla dinámica. Si tenés claro eso, el modelo mental de Looker ya lo tenés.</p>"
            },
            {
              id: "m6c7", type: "mc",
              q: "¿En qué se diferencia Looker (Studio) de Excel?",
              options: [
                "Looker hace cálculos más complejos que Excel",
                "Excel es para calcular y modelar los datos; Looker es para presentar y monitorear KPIs sobre datos ya preparados, conectado a datos vivos",
                "Son lo mismo con otro nombre",
                "Looker reemplaza a las tablas dinámicas dentro de Excel"
              ],
              answer: 1,
              explain: "<p>Flujo típico: limpiar y transformar en Excel/Sheets/SQL, y publicar el dashboard en Looker para que se vea siempre actualizado.</p>"
            }
          ]
        }
      ],
      exercises: [
        {
          id: "m6e1", type: "predict",
          q: "Un producto se vende a $ 1.000 y a la Tienda le cuesta $ 600. En C2: MARGEN sobre venta.",
          grid: { cols: ["Precio", "Costo", "Margen"], rows: [[1000, 600, ""]], money: [0, 1] },
          cell: "C2", formula: "=(A2-B2)/A2", fmt: "pct",
          explain: "<p>Margen sobre venta = (1.000 − 600) / 1.000 = <b>40%</b>.</p>"
        },
        {
          id: "m6e2", type: "predict",
          q: "Mismo producto ($ 1.000 de precio, $ 600 de costo). En D2: MARKUP sobre costo.",
          grid: { cols: ["Precio", "Costo", "·", "Markup"], rows: [[1000, 600, "", ""]], money: [0, 1] },
          cell: "D2", formula: "=(A2-B2)/B2", fmt: "pct",
          explain: "<p>Markup = (1.000 − 600) / 600 ≈ <b>66,7%</b>. Mismo peso, número distinto: un markup del 66,7% es un margen del 40%.</p>"
        },
        {
          id: "m6e3", type: "predict",
          q: "Costos fijos $ 500.000, precio $ 1.000, costo variable $ 600. En D2: punto de equilibrio en unidades.",
          grid: { cols: ["Costos fijos", "Precio", "Costo var.", "Equilibrio"], rows: [[500000, 1000, 600, ""]], money: [0, 1, 2] },
          cell: "D2", formula: "=A2/(B2-C2)",
          explain: "<p>500.000 / (1.000 − 600) = <b>1.250 unidades</b>. Debajo de eso perdés; arriba, ganás.</p>"
        },
        {
          id: "m6e4", type: "order",
          q: "Ordená las etapas de una licitación.",
          steps: [
            "Pliego de bases y condiciones",
            "Presentación de ofertas",
            "Acto de apertura",
            "Cuadro comparativo de ofertas",
            "Adjudicación",
            "Orden de compra"
          ],
          explain: "<p>El analista comercial suele armar el cuadro comparativo y controlar después el cumplimiento del proveedor adjudicado.</p>"
        },
        {
          id: "m6e5", type: "mc",
          q: "'Sell-through' es:",
          options: [
            "El margen sobre venta",
            "Unidades vendidas / unidades recibidas en el período",
            "Las ventas totales del canal",
            "El descuento promedio aplicado"
          ],
          answer: 1,
          explain: "<p>Mide qué proporción de lo recibido ya vendiste. Clave en consignación para decidir reposición y continuidad del consignante.</p>"
        }
      ],
      checklist: [
        "Explico qué es una licitación y por qué una institución pública compra así",
        "Distingo licitación pública, privada y contratación directa por monto y por caso",
        "Defino margen, markup, contribución marginal, punto de equilibrio, ticket promedio y sell-through",
        "Explico qué es Looker / Looker Studio y en qué se diferencia de Excel",
        "Explico «dimensiones vs. métricas» y lo conecto con las tablas dinámicas"
      ],
      keys: [
        { k: "Licitación", v: "Compra competitiva y transparente: pliego → ofertas → cuadro comparativo → adjudicación → OC." },
        { k: "Pública / privada / directa", v: "Por monto: alta = pública abierta; media = privada por invitación; baja o excepción = directa." },
        { k: "Pliego (PByC)", v: "Documento con las condiciones y el criterio de evaluación de la licitación." },
        { k: "Punto de equilibrio", v: "Costos fijos / contribución marginal unitaria. Ventas para no ganar ni perder." },
        { k: "Ticket promedio", v: "Facturación / cantidad de tickets. Sube con precio o con unidades por ticket (UPT)." },
        { k: "KPI", v: "Métrica elegida porque mueve el negocio y tiene una meta asociada." },
        { k: "Looker vs. Excel", v: "Excel: calcular y modelar. Looker: presentar y monitorear KPIs sobre datos ya preparados." },
        { k: "Dimensión vs. métrica", v: "Dimensión = categoría (mes, punto de venta). Métrica = número (importe). Igual que Filas vs. Valores." }
      ]
    },

    /* ===================== MÓDULO 7 ===================== */
    {
      id: 7,
      title: "Simulacro de entrevista técnica",
      short: "Preguntas reales con respuesta modelo y espacio para practicar la tuya",
      goal: "Poder sostener 10–15 minutos de conversación técnica sobre Excel y gestión comercial sin quedar expuesto.",
      lessons: [
        {
          h: "Cómo usar este módulo",
          body:
            "<p>Andá a la sección <b>Entrevista</b> (menú de arriba): ahí está el <b>banco de preguntas</b> agrupado por tema. Para cada una:</p>" +
            "<ol>" +
            "<li>Leé <b>qué buscan</b> con la pregunta.</li>" +
            "<li><b>Escribí tu respuesta</b> en el recuadro (se guarda sola) <b>antes</b> de mirar la respuesta modelo.</li>" +
            "<li>Abrí la <b>respuesta modelo</b> y compará: no para copiarla, para ver qué te faltó nombrar.</li>" +
            "</ol>" +
            "<p>Al final de esa página hay <b>4 casos «resolvé esto en Excel»</b> que se corrigen solos. Después hacé el <b>Quiz</b> filtrando por Módulos 4, 5 y 6.</p>"
        },
        {
          h: "La pregunta del CV — cómo pararte",
          body:
            "<p>Va a aparecer, directa o indirecta: <b>«¿Qué nivel de Excel tenés?»</b>. El objetivo no es mentir ni disculparse, es <b>reencuadrar</b>:</p>" +
            "<ul>" +
            "<li><b>Base honesta:</b> «Manejo bien fórmulas, búsquedas, tablas dinámicas y armado de reportes. Lo que menos toqué en proyectos grandes es macros/VBA.»</li>" +
            "<li><b>Fortaleza real:</b> «Mi diferencial es que vengo de automatización — Make, n8n, integraciones — así que no solo armo el reporte, lo dejo andando solo. Eso aplica directo a las liquidaciones mensuales y a los reportes del área.»</li>" +
            "<li><b>Velocidad de aprendizaje:</b> «Aprendo muy rápido y me preparé específicamente para esta posición: consignación, liquidaciones, licitaciones y la parte de Excel del día a día.»</li>" +
            "<li><b>Cierre:</b> «En dos semanas estoy a full con los reportes del área y aportando mejoras de proceso.»</li>" +
            "</ul>" +
            "<p>Nunca digas «en realidad mi Excel es básico». Decí qué <b>sí</b> podés hacer y mostralo con un ejemplo concreto.</p>",
          checks: [
            {
              id: "m7c1", type: "order",
              q: "Ordená la estructura de la respuesta a «¿qué nivel de Excel tenés?».",
              steps: [
                "Base honesta: qué SÍ manejo (fórmulas, búsquedas, tablas dinámicas, reportes)",
                "Qué toqué menos (macros/VBA en proyectos grandes)",
                "Fortaleza: vengo de automatización, dejo los reportes andando solos",
                "Aprendo rápido y me preparé puntualmente para este puesto",
                "Cierre: en dos semanas estoy a full con los reportes del área"
              ],
              explain: "<p>Arrancás por lo que sí sabés, reconocés el hueco sin dramatizar, y girás rápido a tu diferencial. Nunca abrís con «es básico».</p>"
            }
          ]
        }
      ],
      exercises: [],
      checklist: [
        "Respondí por escrito la pregunta del CV y la puedo decir en voz alta sin dudar",
        "Puedo explicar BUSCARV/BUSCARX, tablas dinámicas y referencias absolutas con un ejemplo del puesto",
        "Puedo explicar consignación, liquidación y licitación en dos frases cada una",
        "Tengo un ejemplo concreto de un proceso que automaticé o mejoré",
        "Hice el quiz de Módulos 4, 5 y 6 con confianza alta"
      ],
      keys: [
        { k: "Pregunta del CV", v: "Base honesta + fortaleza en automatización + aprendo rápido + ejemplo concreto. Nunca «es básico»." },
        { k: "Estructura de respuesta", v: "Qué es → cómo lo haría en Excel → ejemplo con la tienda." },
        { k: "Si no sabés algo", v: "«No lo usé en producción, pero sé para qué sirve y cómo lo encararía» + lo que sí sabés al lado." },
        { k: "Cierre de entrevista", v: "Preguntá: ¿cómo es hoy el proceso de reportes y liquidaciones? Muestra interés y te da info." }
      ]
    }
  ],

  /* ----------------------------------------------------------------------- */
  /* BANCO DE ENTREVISTA (se muestra en la sección Entrevista)               */
  /* ----------------------------------------------------------------------- */
  interview: {
    groups: [
      {
        title: "Excel — técnico",
        items: [
          {
            id: "iv-cv",
            q: "¿Qué nivel de Excel tenés y qué cosas hiciste con Excel?",
            aim: "Ver si lo del CV se sostiene y cómo reaccionás a una pregunta incómoda. Buscan seguridad, no perfección.",
            model:
              "Manejo bien el Excel del día a día de un analista: fórmulas lógicas y de búsqueda (SI, BUSCARV/BUSCARX, INDICE-COINCIDIR), sumas con condiciones (SUMAR.SI.CONJUNTO), tablas dinámicas para reportes y control con formato condicional. Lo que menos trabajé en proyectos grandes es macros/VBA. Mi diferencial es que vengo de automatización —Make, n8n, integraciones—, así que además de armar el reporte lo dejo funcionando solo; eso aplica directo a las liquidaciones y reportes mensuales del área. Y aprendo rápido: me preparé puntualmente para esta posición."
          },
          {
            id: "iv-buscarv",
            q: "¿Cuál es la diferencia entre BUSCARV y BUSCARX?",
            aim: "Saber si entendés el cruce de tablas, que es el 90% del trabajo con datos.",
            model:
              "Las dos traen un dato de otra tabla a partir de una clave. BUSCARV busca en la primera columna del rango y devuelve el dato de la columna número N; solo mira hacia la derecha y si insertás una columna se rompe. BUSCARX es la versión moderna: busca en cualquier dirección, le pasás directamente el rango a devolver (no un número de columna) y trae incorporado el 'si no se encuentra'. BUSCARX está en Excel 365; en Excel 2019/2021 no está y ahí uso INDICE + COINCIDIR, que hace lo mismo."
          },
          {
            id: "iv-abs",
            q: "¿Referencia relativa vs. absoluta? ¿Para qué sirve el signo $?",
            aim: "Concepto base. Si no lo tenés claro, no podés copiar una fórmula sin romperla.",
            model:
              "Cuando copiás una fórmula, las referencias relativas se corren solas: =B2*C2 pasa a =B3*C3. El $ ancla: $B$2 no se mueve, B$2 fija la fila y $B2 fija la columna. Se usa cuando todas las filas tienen que apuntar a la misma celda, por ejemplo la alícuota de IVA o el total para calcular participación. F4 rota entre las variantes."
          },
          {
            id: "iv-pivot",
            q: "¿Qué es una tabla dinámica y cuándo la usás en vez de fórmulas?",
            aim: "Ver si sabés resumir datos sin ensuciar la hoja con mil fórmulas.",
            model:
              "Una tabla dinámica resume miles de filas sin escribir fórmulas, arrastrando campos a cuatro áreas: Filas y Columnas para las categorías, Valores para el número, Filtros arriba. La uso para explorar y probar cortes rápido —ventas por punto de venta, por mes, participación—. Cuando el reporte ya es fijo y siempre igual, lo paso a fórmulas (SUMAR.SI.CONJUNTO) para que se actualice solo. Ojo: la tabla dinámica no se refresca sola, hay que Actualizar con Alt+F5."
          },
          {
            id: "iv-report",
            q: "¿Cómo armarías el reporte de ventas mensual por punto de venta?",
            aim: "El entregable central del puesto. Quieren escuchar un proceso ordenado.",
            model:
              "Primero consolido la venta en una tabla plana: una fila por venta, con Fecha, Punto de venta, Producto, Código, Unidades, Importe. La convierto en Tabla (Ctrl+T). Armo una tabla dinámica con Punto de venta en Filas, Mes en Columnas (agrupando la fecha), Suma de Importe en Valores, y agrego una segunda métrica como % del total general. Sumo la variación contra el mes anterior. Un gráfico de columnas ordenado de mayor a menor y un título con la conclusión. Dejo una segmentación por categoría para responder repreguntas sin rehacer nada."
          },
          {
            id: "iv-top10",
            q: "Te doy 50.000 filas de ventas y te pido el top 10 de productos. ¿Cómo lo hacés?",
            aim: "Reflejo práctico. No quieren que te asuste el volumen.",
            model:
              "Tabla dinámica: Producto en Filas, Suma de Importe en Valores, ordeno de mayor a menor y aplico un filtro de valor '10 mejores'. Si lo necesito como fórmula viva, con Excel 365 uso ORDENAR y tomo las 10 primeras. Con datos así de grandes conviene Power Query o al menos tener el origen como Tabla."
          },
          {
            id: "iv-error",
            q: "¿Cómo detectás y manejás los errores #N/A en un modelo?",
            aim: "Ver si distinguís entre tapar el error y entender por qué aparece.",
            model:
              "#N/A en un BUSCARX/BUSCARV significa que no encontró la clave: puede ser un código nuevo sin dar de alta, espacios de más o un número guardado como texto. Primero lo diagnostico —filtro los #N/A, reviso esas claves—. Para presentar, envuelvo con SI.ERROR y muestro '—' o 'Sin dato', pero solo después de entender el origen: si son ventas sin proveedor asignado en una liquidación, eso hay que corregirlo, no esconderlo."
          },
          {
            id: "iv-sumifs",
            q: "¿Qué es SUMAR.SI.CONJUNTO y dame un ejemplo del puesto?",
            aim: "Suma con condiciones, base de cualquier reporte.",
            model:
              "Suma los valores de un rango que cumplen varios criterios a la vez. El rango a sumar va primero. Ejemplo: total vendido por la Tienda en agosto = SUMAR.SI.CONJUNTO(Importe; PuntoDeVenta;\"Tienda\"; Fecha;\">=2025-08-01\"; Fecha;\"<=2025-08-31\"). Lo mismo para el bruto vendido por proveedor en una liquidación."
          },
          {
            id: "iv-auto",
            q: "Un reporte se hace a mano cada mes y lleva dos días. ¿Cómo lo automatizarías?",
            aim: "Acá jugás tu fortaleza. Buscan iniciativa de mejora de procesos.",
            model:
              "Primero mapeo el proceso actual: de dónde sale cada dato, qué pasos manuales hay, dónde se rompe. Después: estandarizo la plantilla, muevo la limpieza y consolidación a Power Query para que actualice con un botón, dejo las tablas dinámicas y fórmulas atadas a esa Tabla, y si el dato vive en varios lados uso una automatización (Make/n8n) que junte los archivos y dispare la actualización. El resultado final lo puedo publicar en Looker Studio para que el área lo vea siempre al día. La primera versión la hago conservadora y validando contra el reporte viejo."
          }
        ]
      },
      {
        title: "Stock y consignación",
        items: [
          {
            id: "iv-consig",
            q: "¿Qué es la venta por consignación y en qué se diferencia de la compra en firme?",
            aim: "Concepto central de la gestión comercial. Tiene que salir redondo.",
            model:
              "En consignación el proveedor entrega la mercadería a la Tienda para que la venda, pero la propiedad no se transfiere hasta que se vende al cliente final; lo que no se vende se devuelve. La Tienda solo paga lo que efectivamente vendió y retiene una comisión sobre el PVP. En la compra en firme la Tienda compra el stock, lo paga por adelantado y asume el riesgo de que no rote. La consignación le permite a la Tienda ampliar el surtido sin inmovilizar capital ni asumir riesgo de obsolescencia."
          },
          {
            id: "iv-consig-excel",
            q: "¿Cómo controlarías en Excel el stock en consignación separado del propio?",
            aim: "Bajar el concepto a la práctica.",
            model:
              "Marco cada producto con un tipo (Firme / Consignación) o lo llevo en una hoja aparte, porque el stock consignado no es un activo de la Tienda. Registro cada entrega del consignante con remito, fecha, cantidad, PVP y % de comisión, y cada venta con el código, que ya trae el consignante asociado. Al cierre concilio por consignante: stock inicial + entregas − ventas − devoluciones = stock final. Ese cuadre respalda la liquidación."
          },
          {
            id: "iv-rop",
            q: "¿Qué es el punto de reposición y cómo lo mostrás en Excel?",
            aim: "Gestión de inventario aplicada.",
            model:
              "Es el nivel de stock que dispara 'hay que pedir de nuevo': demanda diaria promedio por el lead time del proveedor más un stock de seguridad. En Excel lo muestro con una columna =SI(Stock<=ROP;\"REPONER\";\"OK\") y formato condicional que pinta la fila. Para la orden de compra filtro por 'REPONER' y traigo el proveedor con BUSCARX."
          },
          {
            id: "iv-rotacion",
            q: "¿Qué es la rotación de stock? Si un producto rota poco, ¿qué hacés?",
            aim: "Indicador + criterio comercial.",
            model:
              "La rotación es cuántas veces se renueva el inventario en el período: costo de la mercadería vendida sobre stock promedio a costo. Días de inventario = 365 sobre rotación. Si un producto rota poco tengo capital parado y riesgo de que quede viejo: reviso precio y ubicación en la vitrina, armo un combo o promo, reduzco la próxima compra, o si el proveedor acepta lo paso a consignación para no inmovilizar plata."
          }
        ]
      },
      {
        title: "Liquidaciones",
        items: [
          {
            id: "iv-liq",
            q: "¿Qué es la liquidación mensual a proveedores y qué información lleva?",
            aim: "Responsabilidad explícita del puesto.",
            model:
              "Es el cálculo mensual de cuánto le corresponde cobrar a cada proveedor bajo convenio según lo que efectivamente se vendió de sus productos en el mes. Lleva: encabezado con proveedor, período y número; detalle por producto con unidades, PVP e importe bruto; la comisión de la Tienda con su porcentaje y monto; deducciones por devoluciones o notas de crédito; el neto a liquidar; impuestos y retenciones según la condición fiscal; y la conciliación del stock consignado. Con eso el proveedor factura el neto y se paga según el plazo del convenio."
          },
          {
            id: "iv-liq-controles",
            q: "¿Qué controles hacés antes de enviar una liquidación?",
            aim: "Perfil orientado al detalle.",
            model:
              "Que cierre la aritmética: suma de netos más comisiones más deducciones igual al bruto total vendido del período. Que no haya productos sin proveedor asignado —filtro los #N/A—. Que el porcentaje de comisión sea el del convenio vigente de cada proveedor. Que el stock consignado concilie. Y comparo contra el mes anterior: un salto del 50% casi siempre es un error de datos, no de ventas."
          },
          {
            id: "iv-liq-excel",
            q: "¿Con qué funciones de Excel armás la liquidación?",
            aim: "Conectar concepto con herramienta.",
            model:
              "Parto de la venta del período. Traigo proveedor y % de comisión con BUSCARX. Calculo el bruto por proveedor con SUMAR.SI.CONJUNTO acotando por fechas, o con una tabla dinámica de Proveedor y Producto. Comisión = bruto por el porcentaje; neto = bruto menos comisión menos deducciones. Armo una hoja resumen con una fila por proveedor y un detalle por proveedor filtrable con segmentación. Congelo los datos del período pegando como valores para que una carga tardía no cambie una liquidación ya enviada."
          }
        ]
      },
      {
        title: "Licitaciones y vocabulario",
        items: [
          {
            id: "iv-lici",
            q: "¿Qué es una licitación? ¿Qué tipos hay?",
            aim: "Concepto pedido por el contacto interno. Contexto sector público.",
            model:
              "Es un procedimiento de compra competitivo y transparente: el organismo publica un pliego con lo que necesita y las condiciones, recibe ofertas de varios proveedores y adjudica con reglas conocidas de antemano. Un organismo público compra así porque es público y la normativa exige transparencia e igualdad entre oferentes. Tipos: licitación pública, abierta a cualquiera, para montos altos; licitación privada, por invitación a un grupo, montos medios; y contratación directa, para montos bajos o casos de excepción como urgencia o proveedor único."
          },
          {
            id: "iv-lici-rol",
            q: "¿Qué haría un analista comercial en una licitación?",
            aim: "Ver que te ubicás en el proceso.",
            model:
              "Armo el cuadro comparativo de ofertas en Excel —precios unitarios, totales, condiciones, observaciones técnicas—, proyecto el consumo para dimensionar la compra y comparo los precios ofertados contra el presupuesto oficial y contra lo que se pagó antes. Una vez adjudicado, controlo que el proveedor cumpla: entregas en fecha, precios facturados iguales a los adjudicados, saldo de la orden de compra, y reporto los desvíos."
          },
          {
            id: "iv-margen",
            q: "¿Qué es el margen? ¿Y el punto de equilibrio?",
            aim: "Vocabulario de rentabilidad.",
            model:
              "El margen sobre venta es la porción del precio que queda después del costo: precio menos costo sobre precio. Distinto del markup, que es sobre el costo. El punto de equilibrio es el nivel de ventas donde no ganás ni perdés: costos fijos divididos por la contribución marginal unitaria, es decir por lo que cada unidad aporta para cubrir los fijos."
          },
          {
            id: "iv-looker",
            q: "¿Conocés Looker? ¿En qué se diferencia de Excel?",
            aim: "Lo valoran especialmente. No hace falta ser experto, sí entender el rol.",
            model:
              "Looker Studio es la herramienta gratuita de Google para dashboards interactivos conectados a datos vivos —Sheets, BigQuery, SQL—; Looker a secas es la plataforma enterprise con modelo semántico. Excel es donde calculás y modelás los datos; Looker es donde se presentan y se monitorean los KPIs, ya preparados. Trabaja con dimensiones y métricas, que es la misma lógica de Filas y Valores de una tabla dinámica, así que el modelo mental lo tengo; me pondría al día rápido conectando la venta desde Sheets."
          }
        ]
      },
      {
        title: "Situacionales y de cierre",
        items: [
          {
            id: "iv-mejora",
            q: "Contame una vez que mejoraste o automatizaste un proceso.",
            aim: "Iniciativa de mejora, competencia clave. Preparala con un caso real tuyo.",
            model:
              "(Caso propio.) Estructura recomendada: cuál era el proceso y cuánto tiempo/errores costaba → qué detecté → qué hice concretamente (herramienta, pasos) → resultado medible (horas ahorradas, menos errores, dato disponible antes). Cerrá con cómo lo trasladarías al nuevo puesto: 'lo mismo aplicaría a la liquidación mensual y a los reportes del área'."
          },
          {
            id: "iv-nocuadra",
            q: "El reporte que tenés que presentar en una hora tiene un número que no cuadra. ¿Qué hacés?",
            aim: "Manejo de presión y criterio. Buscan honestidad y método, no héroes.",
            model:
              "Acoto el problema: comparo el total contra una fuente independiente para saber la magnitud del desvío, reviso las uniones (BUSCARX con #N/A), filas duplicadas, filtros olvidados y fechas fuera de período. Si lo encuentro, lo corrijo y dejo anotado qué era. Si en una hora no lo cierro, presento el reporte marcando explícitamente el punto que estoy revisando y el rango probable, y me comprometo a un horario para el número final. Nunca presento un número que sé que está mal como si estuviera bien."
          },
          {
            id: "iv-porque",
            q: "¿Por qué te interesa este puesto? ¿Por qué esta organización?",
            aim: "Motivación y encaje. Quieren saber qué te motiva.",
            model:
              "(Personalizá.) Combina tres cosas que me interesan: una institución de relevancia internacional donde el trabajo tiene impacto y visibilidad, un área comercial con problemas concretos y variados —tienda, stands, funciones, consignación, liquidaciones, licitaciones— y margen real para aportar desde la automatización y el análisis. Es el tipo de rol donde puedo crecer y a la vez dejar procesos mejores que como los encontré."
          }
        ]
      }
    ],
    cases: [
      {
        id: "case1", type: "formula",
        q: "Caso 1 — Traé el costo y calculá el margen. En G2, traé el costo del código de A2 (tabla de la derecha, columnas E y F).",
        grid: {
          cols: ["Código", "Precio", "Unid.", "·", "Cód. prod.", "Costo", "Costo unit."],
          rows: [
            ["LLV-001", 3000, 3, "", "LLV-001", 1700, ""],
            ["LIB-014", 18500, 1, "", "LIB-014", 9200, ""]
          ],
          money: [1, 5, 6]
        },
        cell: "G2", model: "=BUSCARX(A2;E2:E3;F2:F3)",
        explain: "<p><code>=BUSCARX(A2;E2:E3;F2:F3)</code> → 1.700. En Excel 2021: <code>=INDICE(F2:F3;COINCIDIR(A2;E2:E3;0))</code>.</p>"
      },
      {
        id: "case2", type: "formula",
        q: "Caso 2 — Total vendido por proveedor en el mes. En H2, sumá el Importe (C) de «Editorial Central» (B).",
        grid: {
          cols: ["Fecha", "Proveedor", "Importe"],
          rows: [
            ["03/08", "Editorial Central", 620000],
            ["10/08", "Artesanías Rivadavia", 180000],
            ["21/08", "Editorial Central", 620000]
          ],
          money: [2]
        },
        cell: "H2", model: "=SUMAR.SI(B2:B4;\"Editorial Central\";C2:C4)",
        explain: "<p><code>=SUMAR.SI(B2:B4;\"Editorial Central\";C2:C4)</code> → 1.240.000. La tabla dinámica hace lo mismo con Proveedor en Filas.</p>"
      },
      {
        id: "case3", type: "formula",
        q: "Caso 3 — Marcá los productos a reponer. En D2, «REPONER» si el stock (B) es ≤ al punto de reposición (C); si no, «OK».",
        grid: {
          cols: ["Producto", "Stock", "ROP", "Estado"],
          rows: [["Libro «1908»", 3, 8, ""], ["Llavero", 42, 20, ""]]
        },
        cell: "D2", model: "=SI(B2<=C2;\"REPONER\";\"OK\")",
        explain: "<p><code>=SI(B2&lt;=C2;\"REPONER\";\"OK\")</code> → «REPONER». Después, <code>=CONTAR.SI(D:D;\"REPONER\")</code> te dice cuántos hay que pedir.</p>"
      },
      {
        id: "case4", type: "predict",
        q: "Caso 4 — Neto a liquidar con comisión del 30%. Un consignante vendió $ 486.000 y tuvo $ 12.000 de devoluciones. En E2: neto.",
        grid: {
          cols: ["Bruto", "Comisión %", "Devol.", "·", "Neto"],
          rows: [[486000, 0.30, 12000, "", ""]],
          money: [0, 2, 4]
        },
        cell: "E2", formula: "=A2-A2*B2-C2",
        explain: "<p>486.000 − 145.800 − 12.000 = <b>$ 328.200</b>. Control: 328.200 + 145.800 + 12.000 = 486.000 ✓.</p>"
      }
    ]
  },

  /* ----------------------------------------------------------------------- */
  /* QUIZ — banco de preguntas                                               */
  /* ----------------------------------------------------------------------- */
  quiz: [
    { id: "q1-1", module: 1, q: "En una columna de importes, los números aparecen pegados a la izquierda de la celda. ¿Qué significa?", options: [ { t: "Están formateados como moneda", correct: false }, { t: "Son números guardados como texto y van a romper sumas y BUSCARV", correct: true }, { t: "Tienen demasiados decimales", correct: false }, { t: "Están protegidos", correct: false } ], explain: "Excel alinea los números a la derecha y el texto a la izquierda. Un número a la izquierda es texto disfrazado: se corrige con Texto en columnas o multiplicando por 1." },
    { id: "q1-2", module: 1, q: "Tenés la alícuota de IVA en B2 (una sola celda) y querés calcular el precio con IVA en una columna, copiando la fórmula hacia abajo. ¿Cuál usás en la primera fila?", options: [ { t: "=A2*(1+B2)", correct: false }, { t: "=A2*(1+$B$2)", correct: true }, { t: "=$A$2*(1+B2)", correct: false }, { t: "=A3*(1+B$3)", correct: false } ], explain: "A2 tiene que correr con cada fila (relativa); la alícuota tiene que quedar fija en todas (absoluta, $B$2)." },
    { id: "q1-3", module: 1, q: "¿Qué ventaja da convertir un rango en Tabla (Ctrl+T) antes de armar una tabla dinámica?", options: [ { t: "La tabla dinámica se actualiza sola al abrir", correct: false }, { t: "Cuando agregás filas nuevas, la tabla dinámica las toma con solo Actualizar, sin cambiar el rango", correct: true }, { t: "Permite usar más de 65.000 filas", correct: false }, { t: "Convierte las fechas de texto a fecha automáticamente", correct: false } ], explain: "Una Tabla se expande sola. La tabla dinámica apunta al nombre de la Tabla, así que no hay que reajustar rangos cuando crecen los datos." },
    { id: "q1-4", module: 1, q: "Querés pintar de rojo toda la fila cuando el stock (columna F) cae por debajo del ROP (columna G). ¿Qué fórmula va en el formato condicional?", options: [ { t: "=F2<G2", correct: false }, { t: "=$F2<=$G2", correct: true }, { t: "=$F$2<=$G$2", correct: false }, { t: "=F$2<=G$2", correct: false } ], explain: "Para pintar la fila entera hay que fijar la columna con $ y dejar la fila libre: $F2 y $G2." },

    { id: "q2-1", module: 2, q: "¿Qué función usarías para traer el precio de un producto desde otra tabla a partir de su código?", options: [ { t: "CONCATENAR", correct: false }, { t: "BUSCARV (o BUSCARX)", correct: true }, { t: "CONTAR.SI", correct: false }, { t: "PROMEDIO", correct: false } ], explain: "BUSCARV/BUSCARX cruzan tablas por una clave. BUSCARV busca en la primera columna y devuelve un dato de la misma fila." },
    { id: "q2-2", module: 2, q: "El cuarto argumento de BUSCARV, ¿por qué conviene que sea siempre FALSO (o 0)?", options: [ { t: "Para que ordene la tabla automáticamente", correct: false }, { t: "Para pedir coincidencia exacta; con VERDADERO da resultados incorrectos si la tabla no está ordenada", correct: true }, { t: "Para que ignore mayúsculas y minúsculas", correct: false }, { t: "Para que devuelva un texto en vez de un número", correct: false } ], explain: "FALSO = coincidencia exacta. VERDADERO hace una búsqueda aproximada que solo sirve con la primera columna ordenada y suele traer el dato de al lado." },
    { id: "q2-3", module: 2, q: "Nombrá una ventaja concreta de BUSCARX sobre BUSCARV.", options: [ { t: "Es más rápida de escribir siempre", correct: false }, { t: "Puede buscar en cualquier dirección y no se rompe si insertás una columna", correct: true }, { t: "Funciona en todas las versiones de Excel", correct: false }, { t: "No necesita que la clave sea única", correct: false } ], explain: "BUSCARX no depende de un número de columna y busca a izquierda o derecha. (No está en Excel 2016/2019/2021: ahí va INDICE+COINCIDIR.)" },
    { id: "q2-4", module: 2, q: "En SUMAR.SI.CONJUNTO, ¿qué argumento va primero?", options: [ { t: "El primer rango de criterio", correct: false }, { t: "El rango que se va a sumar", correct: true }, { t: "El criterio de texto", correct: false }, { t: "El rango de fechas", correct: false } ], explain: "A diferencia de SUMAR.SI, en la versión .CONJUNTO el rango a sumar va primero, y después vienen los pares rango_criterio / criterio." },
    { id: "q2-5", module: 2, q: "Un producto se vende a $ 1.000 y cuesta $ 600. ¿Cuál es el margen sobre venta?", options: [ { t: "66,7%", correct: false }, { t: "40%", correct: true }, { t: "60%", correct: false }, { t: "$ 400 de markup, no se puede calcular en %", correct: false } ], explain: "Margen sobre venta = (1000 − 600) / 1000 = 40%. El markup sobre costo sería 400/600 = 66,7%." },
    { id: "q2-6", module: 2, q: "Un BUSCARX devuelve #N/A para varias filas en una liquidación. ¿Cuál es la mejor primera acción?", options: [ { t: "Envolver todo con SI.ERROR y seguir", correct: false }, { t: "Filtrar los #N/A y revisar esas claves: suelen ser códigos sin dar de alta o con espacios", correct: true }, { t: "Borrar esas filas", correct: false }, { t: "Cambiar BUSCARX por BUSCARV", correct: false } ], explain: "En una liquidación, #N/A puede ser ventas sin proveedor asignado: hay que corregir el dato, no taparlo. SI.ERROR viene después, para presentar." },
    { id: "q2-7", module: 2, q: "Ventas de septiembre: $ 3.200.000. Agosto: $ 2.900.000. ¿Variación?", options: [ { t: "+9,4%", correct: false }, { t: "+10,3%", correct: true }, { t: "+30%", correct: false }, { t: "−10,3%", correct: false } ], explain: "(3.200.000 − 2.900.000) / 2.900.000 = 300.000 / 2.900.000 ≈ 10,3%." },

    { id: "q3-1", module: 3, q: "En una tabla dinámica, ¿a qué área va el campo 'Punto de venta' y a cuál 'Importe'?", options: [ { t: "Ambos a Valores", correct: false }, { t: "Punto de venta a Filas (o Columnas); Importe a Valores", correct: true }, { t: "Punto de venta a Filtros; Importe a Filas", correct: false }, { t: "Ambos a Filas", correct: false } ], explain: "Las categorías van a Filas o Columnas; los números a medir van a Valores." },
    { id: "q3-2", module: 3, q: "Cargaste ventas nuevas en el origen pero la tabla dinámica sigue mostrando los totales viejos. ¿Qué pasa?", options: [ { t: "La tabla dinámica está dañada, hay que rehacerla", correct: false }, { t: "La tabla dinámica no se actualiza sola: hay que Actualizar (Alt+F5)", correct: true }, { t: "Falta guardar el archivo", correct: false }, { t: "El origen tiene que estar en otra hoja", correct: false } ], explain: "La tabla dinámica trabaja sobre una copia en caché. Clic derecho › Actualizar, o Alt+F5." },
    { id: "q3-3", module: 3, q: "Querés mostrar qué porcentaje de la venta total representa cada punto de venta. ¿Qué usás?", options: [ { t: "Una columna con fórmula al costado de la tabla dinámica", correct: false }, { t: "Configuración de campo de valor › Mostrar valores como › % del total general", correct: true }, { t: "Formato de celda › Porcentaje", correct: false }, { t: "Un gráfico de torta, no se puede en la tabla", correct: false } ], explain: "'Mostrar valores como' recalcula el campo como % del total general, % de la columna, diferencia respecto de otro período, etc." },
    { id: "q3-4", module: 3, q: "Tenés fechas día por día y al ponerlas en Columnas la tabla dinámica genera decenas de columnas. ¿Solución más limpia?", options: [ { t: "Borrar las fechas que sobran", correct: false }, { t: "Clic derecho sobre una fecha › Agrupar › Meses", correct: true }, { t: "Cambiar el formato de las celdas a 'mmmm'", correct: false }, { t: "Filtrar solo un día", correct: false } ], explain: "Agrupar por Meses (y Trimestres/Años) crea un campo 'Mes' y colapsa las columnas. Si no deja agrupar, las fechas son texto." },
    { id: "q3-5", module: 3, q: "¿Cuándo conviene un reporte con fórmulas (SUMAR.SI.CONJUNTO) en vez de una tabla dinámica?", options: [ { t: "Nunca, la tabla dinámica siempre es mejor", correct: false }, { t: "Cuando el reporte tiene una forma fija y querés que se actualice solo sin tocar nada", correct: true }, { t: "Cuando hay más de 1.000 filas", correct: false }, { t: "Cuando no sabés usar tablas dinámicas", correct: false } ], explain: "La tabla dinámica es para explorar; las fórmulas son para un entregable estándar que se recalcula solo al cambiar los datos." },

    { id: "q4-1", module: 4, q: "¿Cuál es la ecuación del stock final?", options: [ { t: "Stock final = Ventas − Compras", correct: false }, { t: "Stock final = Stock inicial + Entradas − Salidas", correct: true }, { t: "Stock final = Stock inicial × Rotación", correct: false }, { t: "Stock final = Entradas − Ventas + Merma", correct: false } ], explain: "Entradas = compras/remitos/devoluciones de clientes. Salidas = ventas + ajustes por merma, rotura o robo." },
    { id: "q4-2", module: 4, q: "En la venta por consignación, ¿cuándo se transfiere la propiedad de la mercadería?", options: [ { t: "Cuando el proveedor entrega la mercadería a la Tienda", correct: false }, { t: "Cuando la Tienda la vende al cliente final", correct: true }, { t: "Al firmar el convenio", correct: false }, { t: "A los 30 días de la entrega", correct: false } ], explain: "En consignación la propiedad pasa recién con la venta al cliente final. Lo que no se vende se devuelve al consignante." },
    { id: "q4-3", module: 4, q: "Principal diferencia entre compra en firme y consignación para la Tienda:", options: [ { t: "En firme se paga más caro", correct: false }, { t: "En firme la Tienda compra el stock, lo paga por adelantado y asume el riesgo; en consignación paga solo lo que vende", correct: true }, { t: "La consignación no permite fijar el PVP", correct: false }, { t: "No hay diferencia contable", correct: false } ], explain: "La consignación evita inmovilizar capital y asumir riesgo de obsolescencia; el stock consignado no es un activo de la Tienda." },
    { id: "q4-4", module: 4, q: "El punto de reposición (ROP) se calcula como:", options: [ { t: "Stock inicial − ventas del mes", correct: false }, { t: "Demanda diaria promedio × lead time + stock de seguridad", correct: true }, { t: "Costo de la mercadería vendida / stock promedio", correct: false }, { t: "PVP − costo", correct: false } ], explain: "El ROP cubre la demanda esperada durante el tiempo que tarda el proveedor (lead time), más un colchón para picos y demoras." },
    { id: "q4-5", module: 4, q: "Un producto tiene rotación baja. ¿Qué implica y qué harías?", options: [ { t: "Se vende muy rápido; hay que comprar más", correct: false }, { t: "Capital inmovilizado y riesgo de quedar viejo; revisar precio/ubicación, promo, comprar menos o pasarlo a consignación", correct: true }, { t: "Que el margen es negativo", correct: false }, { t: "Que hubo un quiebre de stock", correct: false } ], explain: "Rotación baja = mucho stock parado por cada unidad que se vende. Es típico en souvenirs de nicho y libros caros." },
    { id: "q4-6", module: 4, q: "Un consignante tenía 5 unidades, entregó 40, se vendieron 28 y se devolvieron 2. Stock final consignado y sell-through:", options: [ { t: "15 unidades; 62%", correct: true }, { t: "10 unidades; 70%", correct: false }, { t: "17 unidades; 56%", correct: false }, { t: "13 unidades; 65%", correct: false } ], explain: "Final = 5 + 40 − 28 − 2 = 15. Sell-through = 28 / (5+40) = 28/45 ≈ 62%." },
    { id: "q4-7", module: 4, q: "¿Por qué conviene llevar el stock en consignación separado del stock propio?", options: [ { t: "Para pagar menos impuestos", correct: false }, { t: "Porque el stock consignado no es un activo de la Tienda y hay que conciliarlo por consignante para la liquidación", correct: true }, { t: "Porque ocupa menos espacio en el archivo", correct: false }, { t: "No conviene, es mejor todo junto", correct: false } ], explain: "Mezclar distorsiona el valor del inventario propio y complica la conciliación (stock inicial + entregas − ventas − devoluciones) que respalda cada liquidación." },

    { id: "q5-1", module: 5, q: "¿Qué es la liquidación mensual a proveedores?", options: [ { t: "El pago de sueldos del área comercial", correct: false }, { t: "El cálculo de cuánto le corresponde cobrar a cada proveedor bajo convenio según lo efectivamente vendido en el período", correct: true }, { t: "El cierre de caja diario de la Tienda", correct: false }, { t: "El inventario físico de fin de mes", correct: false } ], explain: "Es la contracara de la consignación: al no comprar el stock por adelantado, cada mes se calcula lo vendido de cada proveedor y su neto." },
    { id: "q5-2", module: 5, q: "El neto a liquidar a un proveedor se calcula como:", options: [ { t: "Bruto vendido + comisión de la Tienda", correct: false }, { t: "Bruto vendido − comisión de la Tienda − deducciones (devoluciones, NC, mermas)", correct: true }, { t: "Unidades vendidas × costo", correct: false }, { t: "PVP × stock final", correct: false } ], explain: "La Tienda retiene su comisión y descuenta lo que corresponda; el proveedor factura el neto." },
    { id: "q5-3", module: 5, q: "Un consignante vendió $ 486.000, con $ 12.000 de devoluciones y comisión de la Tienda del 30%. Neto a liquidar:", options: [ { t: "$ 340.200", correct: false }, { t: "$ 328.200", correct: true }, { t: "$ 316.200", correct: false }, { t: "$ 474.000", correct: false } ], explain: "Comisión = 486.000 × 30% = 145.800. Neto = 486.000 − 145.800 − 12.000 = 328.200." },
    { id: "q5-4", module: 5, q: "¿Cuál es el control aritmético que tiene que dar antes de enviar una liquidación?", options: [ { t: "Σ netos = Σ comisiones", correct: false }, { t: "Σ netos + Σ comisiones + Σ deducciones = bruto total vendido del período", correct: true }, { t: "Σ netos = stock final × PVP", correct: false }, { t: "Σ comisiones = 30% de las unidades", correct: false } ], explain: "Si no cierra, hay una fila mal cargada, un % de comisión equivocado o ventas sin proveedor." },
    { id: "q5-5", module: 5, q: "Antes de enviar una liquidación filtrás la columna de proveedor por #N/A / vacío. ¿Para qué?", options: [ { t: "Para borrar esos productos", correct: false }, { t: "Para detectar ventas sin proveedor asignado, que hacen que la liquidación no cuadre", correct: true }, { t: "Para ordenar alfabéticamente", correct: false }, { t: "Para aplicar SI.ERROR", correct: false } ], explain: "Cada venta tiene que poder asignarse a un proveedor; un código sin dar de alta rompe el cuadre." },
    { id: "q5-6", module: 5, q: "¿Qué función principal usás para calcular el bruto vendido por proveedor en un período?", options: [ { t: "CONCATENAR", correct: false }, { t: "SUMAR.SI.CONJUNTO (o una tabla dinámica)", correct: true }, { t: "CONTARA", correct: false }, { t: "BUSCARV con VERDADERO", correct: false } ], explain: "SUMAR.SI.CONJUNTO(Importe; Proveedor; X; Fecha; \">=desde\"; Fecha; \"<=hasta\"). La TD es la alternativa visual." },

    { id: "q6-1", module: 6, q: "¿Por qué un organismo público compra por licitación en vez de elegir proveedor directamente?", options: [ { t: "Porque siempre sale más barato", correct: false }, { t: "Porque es un organismo público y la normativa exige transparencia, concurrencia e igualdad entre oferentes", correct: true }, { t: "Porque no tiene presupuesto propio", correct: false }, { t: "Porque lo pide cada proveedor", correct: false } ], explain: "Arriba de ciertos montos el sector público no puede comprar 'a dedo': tiene que abrir el proceso y adjudicar con reglas conocidas." },
    { id: "q6-2", module: 6, q: "¿Cuál es la diferencia entre licitación pública y contratación directa?", options: [ { t: "La pública es para servicios y la directa para productos", correct: false }, { t: "La pública es abierta a cualquier oferente y para montos altos; la directa es para montos bajos o casos de excepción (urgencia, proveedor único)", correct: true }, { t: "La directa necesita más firmas", correct: false }, { t: "No hay diferencia real", correct: false } ], explain: "Entre medio está la licitación privada, por invitación a un grupo de proveedores, para montos intermedios." },
    { id: "q6-3", module: 6, q: "Punto de equilibrio (break-even) es:", options: [ { t: "El precio al que el margen es 100%", correct: false }, { t: "El nivel de ventas donde no se gana ni se pierde: costos fijos / contribución marginal unitaria", correct: true }, { t: "El stock mínimo antes de reponer", correct: false }, { t: "La diferencia entre PVP y costo", correct: false } ], explain: "La contribución marginal (precio − costo variable unitario) es lo que cada unidad aporta para cubrir los costos fijos." },
    { id: "q6-4", module: 6, q: "Ticket promedio se calcula como:", options: [ { t: "Facturación / cantidad de tickets (operaciones)", correct: true }, { t: "Unidades vendidas / visitantes", correct: false }, { t: "Costo / margen", correct: false }, { t: "Ventas / stock promedio", correct: false } ], explain: "Sube subiendo precios o vendiendo más unidades por operación (UPT: unidades por ticket)." },
    { id: "q6-5", module: 6, q: "¿En qué se diferencia Looker (Studio) de Excel?", options: [ { t: "Looker hace cálculos más complejos que Excel", correct: false }, { t: "Excel es para calcular y modelar los datos; Looker es para presentar y monitorear KPIs sobre datos ya preparados, conectado a datos vivos", correct: true }, { t: "Son lo mismo con otro nombre", correct: false }, { t: "Looker reemplaza a las tablas dinámicas dentro de Excel", correct: false } ], explain: "El flujo típico: limpiar/transformar en Excel/Sheets/SQL y publicar el dashboard en Looker. Trabaja con dimensiones y métricas, igual que Filas y Valores." },
    { id: "q6-6", module: 6, q: "'Dimensión' y 'métrica' en un tablero equivalen, en una tabla dinámica, a:", options: [ { t: "Filtros y Segmentaciones", correct: false }, { t: "Filas/Columnas (categorías) y Valores (números)", correct: true }, { t: "Encabezado y pie", correct: false }, { t: "Origen y caché", correct: false } ], explain: "Dimensión = con qué desagregás (mes, punto de venta, categoría). Métrica = qué medís (importe, unidades, ticket)." },
    { id: "q6-7", module: 6, q: "'Sell-through' es:", options: [ { t: "El margen sobre venta", correct: false }, { t: "Unidades vendidas / unidades recibidas en el período", correct: true }, { t: "Ventas totales del canal", correct: false }, { t: "El descuento promedio aplicado", correct: false } ], explain: "Mide qué proporción de lo que recibiste ya vendiste. Es clave en consignación para decidir reposición y continuidad del consignante." },
    { id: "q6-8", module: 6, q: "Análisis ABC / Pareto sobre los productos de la Tienda significa:", options: [ { t: "Ordenarlos alfabéticamente en tres grupos", correct: false }, { t: "Que ~20% de los productos explica ~80% de la venta, y se gestionan distinto A, B y C", correct: true }, { t: "Clasificarlos por proveedor", correct: false }, { t: "Un método de conteo de inventario", correct: false } ], explain: "Los 'A' (pocos, mucha venta) merecen control fino de stock y nunca quebrar; los 'C' se gestionan con menos esfuerzo." }
  ],

  /* ----------------------------------------------------------------------- */
  /* CHEAT SHEET — repaso final                                              */
  /* ----------------------------------------------------------------------- */
  cheat: [
    {
      title: "Fórmulas clave",
      items: [
        { k: "BUSCARV(val; tabla; nºcol; FALSO)", v: "Busca en la 1ª columna, devuelve el dato de la columna nºcol. Siempre FALSO = exacto. Solo mira a la derecha." },
        { k: "BUSCARX(val; buscar_en; devolver; \"si no\")", v: "Moderna: cualquier dirección, sin nº de columna. Excel 365. En 2019/2021 no está." },
        { k: "INDICE(devolver; COINCIDIR(val; buscar_en; 0))", v: "Reemplazo de BUSCARX en toda versión. COINCIDIR da la posición, INDICE trae el dato." },
        { k: "SI(prueba; sí; no)", v: "Lógica condicional. Anidar para 3+ caminos. Y()/O() para varias condiciones." },
        { k: "SUMAR.SI(rango; criterio; sumar)", v: "Suma con una condición. Rango de suma al final." },
        { k: "SUMAR.SI.CONJUNTO(sumar; c1; k1; c2; k2)", v: "Varios criterios. El rango a sumar va PRIMERO." },
        { k: "SI.ERROR(fórmula; \"—\")", v: "Tapa errores para presentar. No lo uses para esconder datos mal cargados." }
      ]
    },
    {
      title: "Referencias y datos",
      items: [
        { k: "$A$1 / A$1 / $A1", v: "Absoluta / fila fija / columna fija. F4 rota. Para pintar filas: $A2." },
        { k: "Ctrl+T", v: "Tabla: se expande sola, ordena/filtra desde el encabezado, fórmulas legibles." },
        { k: "Número a la izquierda", v: "= texto disfrazado. Rompe sumas y BUSCARV. Arreglo: ×1 o Texto en columnas." },
        { k: "Formato ≠ valor", v: "El formato solo cambia la vista. Las fechas son números por dentro." }
      ]
    },
    {
      title: "Tablas dinámicas",
      items: [
        { k: "4 áreas", v: "Filas / Columnas = categorías. Valores = números. Filtros = filtro general." },
        { k: "Mostrar valores como", v: "% del total general, % de columna, Diferencia respecto del mes anterior." },
        { k: "Agrupar fechas", v: "Clic derecho sobre una fecha › Agrupar › Meses/Trimestres/Años." },
        { k: "Actualizar", v: "Alt+F5. NO se refresca sola al cambiar los datos." },
        { k: "Segmentación", v: "Filtro visual; con Conexiones de informe controla varias TD." },
        { k: "TD vs. fórmulas", v: "TD para explorar; fórmulas para el reporte fijo que se actualiza solo." }
      ]
    },
    {
      title: "Cálculos comerciales",
      items: [
        { k: "Margen", v: "(PV − Costo) / PV. Lo que queda de la venta." },
        { k: "Markup", v: "(PV − Costo) / Costo. Lo que se recarga al costo. ≠ margen." },
        { k: "Variación %", v: "(actual − anterior) / anterior." },
        { k: "Participación / mix", v: "parte / total (total anclado con $)." },
        { k: "Punto de equilibrio", v: "Costos fijos / contribución marginal unitaria." },
        { k: "Ticket promedio", v: "Facturación / cantidad de tickets." },
        { k: "IVA", v: "c/IVA = precio × 1,21. Sacar IVA = precio / 1,21." }
      ]
    },
    {
      title: "Stock",
      items: [
        { k: "Stock final", v: "Stock inicial + Entradas − Salidas (ventas + ajustes)." },
        { k: "ROP", v: "Demanda diaria × lead time + stock de seguridad." },
        { k: "Rotación", v: "CMV / stock promedio a costo. Días de inv. = 365 / rotación." },
        { k: "Rotación baja", v: "Capital parado + riesgo de obsolescencia → precio, promo, comprar menos, consignación." },
        { k: "Sell-through", v: "Unidades vendidas / unidades recibidas en el período." },
        { k: "Alerta", v: "=SI(Stock<=ROP;\"REPONER\";\"OK\") + formato condicional $F2<=$G2." }
      ]
    },
    {
      title: "Consignación y liquidación",
      items: [
        { k: "Consignación", v: "El proveedor entrega mercadería; la propiedad pasa al venderse. Lo no vendido se devuelve." },
        { k: "En firme vs. consignación", v: "En firme la Tienda compra y asume el riesgo. Consignación: paga solo lo vendido." },
        { k: "Liquidación", v: "Cálculo mensual de lo que se le paga a cada proveedor bajo convenio, según lo vendido." },
        { k: "Neto a liquidar", v: "Bruto vendido − comisión de la Tienda − deducciones." },
        { k: "Lleva", v: "Detalle por producto, unidades, PVP, comisión %, deducciones, neto, impuestos, conciliación de stock." },
        { k: "Control", v: "Σ netos + Σ comisiones + Σ deducciones = bruto total. Sin productos #N/A." }
      ]
    },
    {
      title: "Vocabulario y licitaciones",
      items: [
        { k: "Licitación", v: "Compra pública competitiva y transparente: pliego → ofertas → cuadro comparativo → adjudicación → OC." },
        { k: "Pública / privada / directa", v: "Abierta y monto alto / por invitación y monto medio / monto bajo o excepción." },
        { k: "Pliego (PByC)", v: "Documento con condiciones y criterio de evaluación." },
        { k: "KPI", v: "Métrica que mueve el negocio y tiene meta." },
        { k: "Forecast / presupuesto / desvío", v: "Pronóstico / plan / real − plan." },
        { k: "Looker vs. Excel", v: "Excel calcula y modela; Looker presenta y monitorea KPIs sobre datos vivos." },
        { k: "Dimensión vs. métrica", v: "Categoría vs. número. Igual que Filas vs. Valores en una TD." }
      ]
    },
    {
      title: "Antes de entrar a la entrevista",
      items: [
        { k: "La pregunta del CV", v: "Qué SÍ sé + fortaleza en automatización + aprendo rápido + ejemplo concreto. Nunca «es básico»." },
        { k: "Estructura de respuesta", v: "Qué es → cómo lo hago en Excel → ejemplo con la tienda." },
        { k: "Si no sé algo", v: "«No lo usé en producción pero sé para qué sirve y cómo lo encararía» + lo que sí sé." },
        { k: "Repasar de memoria", v: "BUSCARV/BUSCARX, SI anidado, armar una TD, consignación, liquidación, licitación." },
        { k: "Pregunta para el final", v: "¿Cómo es hoy el proceso de reportes y de liquidación mensual?" }
      ]
    }
  ]
};

if (typeof window !== "undefined") window.COURSE = COURSE;
if (typeof module !== "undefined") module.exports = COURSE;
