# Plan: Reparación y Modal Genérico de Confirmación

## Objective
Solucionar el error de ejecución de `SpreadsheetApp.getUi().toast` al detener la sincronización y mejorar la UX reemplazando los `ui.alert` nativos por un modal HTML genérico basado en Materialize, que pueda utilizarse tanto para detener la sincronización como para reparar el sistema.

## Key Files & Context
- `src/Activador.js`: Contiene las funciones de eliminación y reparación (`eliminarActivador`, `repararActivador`). Será necesario adaptarlas para lanzar el modal y actuar como callbacks desde el lado cliente.
- `src/Base.js`: Configura el menú. Podría requerir pequeños ajustes si las funciones enrutadas cambian.
- `src/DialogoConfirmacion.html`: Nuevo archivo. Plantilla HTML reutilizable con Materialize para mostrar título, mensaje de advertencia y botones (Cancelar/Aceptar).

## Implementation Steps

### 1. Corrección del Bug y Refactorización del Backend (`Activador.js`)
- **Fix Toast:** Corregir `SpreadsheetApp.getUi().toast` a `SpreadsheetApp.getActive().toast` dentro de la lógica de eliminación.
- **Preparar callbacks:**
  - Crear una función envoltorio `confirmarEliminarActivador()` que reciba la confirmación desde el frontend y ejecute `eliminarActivador(true)` sin pasar por el `ui.alert`.
  - Crear una función envoltorio `confirmarRepararActivador()` que ejecute la reparación de forma directa.
- **Lanzadores del modal:**
  - Crear `mostrarDialogoDetener()`: Renderiza el modal genérico con parámetros específicos (título, mensaje de aviso sobre la detención, acción "detener").
  - Crear `mostrarDialogoReparar()`: Renderiza el modal genérico con parámetros específicos (aviso de destrucción de todos los activadores, acción "reparar").

### 2. Creación del Modal HTML Genérico (`DialogoConfirmacion.html`)
- Crear la estructura con Materialize (contenedor, fila para el texto y botones).
- Usar scripts en línea (`<?= titulo ?>`, `<?= mensaje ?>`, `<?= funcionExito ?>`) para inyectar los valores desde el backend.
- Añadir barra de progreso (por si el proceso tarda unos milisegundos).
- Bloquear botones al confirmar y cerrar el modal tras el éxito, delegando la ejecución final y el toast al script `.gs`.

### 3. Actualización de Menú (`Base.js`)
- Actualizar `onOpen` para enrutar el menú hacia los nuevos lanzadores (`mostrarDialogoDetener` y `mostrarDialogoReparar`) en lugar de a las lógicas directas.

## Verification
- Validar que al pulsar "Detener" o "Reparar", se abre el modal correcto con el diseño de Materialize.
- Confirmar que tras aceptar, se ejecuta el proceso sin generar errores, mostrando el toast correctamente a través de `SpreadsheetApp.getActive().toast`.