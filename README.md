# Excel Intensivo

Curso interactivo de Excel en español (es-AR) con un motor de fórmulas propio: escribís la fórmula y la app la evalúa y corrige en el navegador, sin Excel.

## Qué problema resuelve

Practicar Excel suele requerir tener el programa instalado y alguien que corrija los ejercicios. Esta app estática permite estudiar desde el celular o la PC, resolver ejercicios sobre una grilla simulada y recibir la corrección al instante, con teoría aplicada a casos de gestión comercial.

## Funcionalidades

- 7 módulos con teoría, ejercicios interactivos con solución, checklist y repaso: fundamentos, fórmulas para análisis comercial, tablas dinámicas, stock y consignación, liquidación a proveedores, vocabulario comercial y licitaciones, y simulacro de entrevista técnica.
- Motor de fórmulas propio (`engine.js`) con unas 30 funciones en español: SUMA, SI, BUSCARV, BUSCARX, SUMAR.SI.CONJUNTO, INDICE/COINCIDIR y más. Evalúa lo que el usuario escribe sobre una grilla en memoria.
- Tipos de ejercicio: escribir la fórmula, predecir el resultado, armarla con fichas (pensado para celular), clasificar, ordenar pasos y opción múltiple.
- Quiz con feedback inmediato, filtrable por módulo.
- Banco de preguntas de entrevista con campo para respuestas propias.
- Hoja de referencia rápida (cheat sheet) imprimible.
- Progreso guardado en `localStorage`.

## Stack

| Parte | Tecnología |
|---|---|
| App | HTML, CSS y JavaScript sin frameworks ni dependencias |
| Router | Por hash, en `app.js` |
| Contenido | `content.js` |
| Motor de fórmulas | `engine.js` |
| Persistencia | `localStorage` |

## Arquitectura

- `index.html`: estructura, carga de fuentes y scripts.
- `styles.css`: sistema visual con modo claro y oscuro.
- `content.js`: todo el contenido del curso (módulos, quiz, entrevista, casos).
- `engine.js`: parser y evaluador de fórmulas estilo Excel.
- `app.js`: ruteo, vistas, motor de ejercicios y persistencia del progreso.

## Cómo correrlo localmente

No necesita instalación ni build. Con cualquier servidor estático:

```bash
python -m http.server 8000
# o
npx serve .
```

Abrí `http://localhost:8000`.

## Tests

Este proyecto no tiene tests automatizados. El motor de fórmulas se verifica a través de los ejercicios de la propia app, que comparan el resultado de la fórmula escrita contra el esperado.

## Estado del proyecto

Aplicación estática completa y funcional. El progreso se guarda por dispositivo, así que celular y PC llevan avances independientes.

## Qué aprendí

- Escribir un parser y evaluador de fórmulas desde cero, con las funciones y la sintaxis de Excel en español.
- Diseñar ejercicios que se corrigen solos comparando el resultado de la fórmula con el esperado.
- Construir una SPA sin frameworks, con router por hash y persistencia en `localStorage`.

## English summary

Excel Intensivo is an interactive Excel course in Spanish. Its main feature is a custom formula engine, written in JavaScript, with about 30 Excel functions such as SUMA, SI, BUSCARV, BUSCARX and SUMAR.SI.CONJUNTO. Students type a formula, and the app evaluates it on a simulated grid and checks the answer in the browser. The course has 7 modules, a quiz and an interview question bank, and it saves progress in localStorage. It is a static app with no framework, and it runs with any static web server.