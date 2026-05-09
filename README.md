<p align="center">
  <img src="assets/banner.png" alt="Asistente de Grupos Banner">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white" alt="Google Apps Script">
  <img src="https://img.shields.io/badge/Versión-2.0-indigo" alt="Versión">
  <img src="https://img.shields.io/badge/Licencia-GPL_v3-blue" alt="Licencia">
</p>

> 🐙 **Asistente de Grupos**: La herramienta definitiva para mantener tus grupos de Google Workspace en orden, sin perder la cordura en el proceso.

---

### 🎯 ¿De qué va esto?

Si gestionas un dominio de **Google Workspace**, sabrás que mantener los grupos de correo actualizados (claustros, departamentos, alumnos por niveles...) puede convertirse en un "ritual" semanal bastante tedioso. 

Este proyecto nace para automatizar ese proceso. Permite sincronizar listas de miembros (usuarios internos y externos) desde una sencilla hoja de cálculo hacia los grupos del servicio de **Google Groups**, ya sea de forma manual o programando una ejecución automática.

### 📜 Una historia de "cocción lenta" (90% Humana, 10% IA)

Este no es un proyecto de "usar y tirar" generado en 5 minutos por un bot. Su historia tiene solera:

1.  **El origen (Junio 2023):** Fue creado de manera totalmente manual para resolver una necesidad real en el centro educativo donde soy Jefe de Estudios, Tecnología y Calidad.
2.  **La prueba de fuego:** Ha estado funcionando "en la sombra" durante **3 cursos escolares completos** (23/24, 24/25 y 25/26) con una fiabilidad total.
3.  **El empujón final (Mayo 2026):** Siempre quise liberarlo, pero me faltaba tiempo para pulir la interfaz de usuario. Gracias a **Gemini CLI**, en una mañana de sábado, hemos logrado "cablear" ese diálogo de programación granular que tanto tiempo llevaba en mi cabeza.

En resumen: es un código con **90% de ADN humano** y un **10% de potencia IA** para los retoques finales de UX.

---

### ✨ Características principales

1.  **Sincronización Inteligente:** Añade nuevos miembros y elimina los que ya no están en la lista con un solo clic.
2.  **Planificador Granular:** Olvídate de editar constantes en el código. Configura la sincronización cada X horas, ciertos días de la semana (ej. Lunes, Miércoles y Viernes) o cada X días, todo desde un modal visual.
3.  **Registro y Auditoría:** Una hoja de "Registro" detalla cada operación. Si pasas el ratón sobre los números, verás una **nota de celda** con la lista exacta de emails añadidos o eliminados.
4.  **Control Total:** Tú decides si quieres registrar resúmenes o detalles mediante sencillos interruptores (checks) en la propia hoja.
5.  **Directorio a mano:** Descarga la lista de usuarios y grupos de tu dominio con un botón para facilitar la creación de nuevas pestañas de gestión.

---

### 🚀 Cómo empezar

La distribución de este proyecto es ultra-sencilla. No necesitas instalar `clasp` ni pelearte con la consola (a menos que quieras):

1.  **Crea tu copia:** Duplica esta **[Plantilla de Google Sheets](https://docs.google.com/spreadsheets/d/1ZwuDBCEZKd1ELFGHC67WAgTrJvWyHyNrBU0QG0P3Zso/edit?usp=sharing)**.
2.  **Autoriza:** Ve al menú `🐙 Asistente de Grupos` y selecciona cualquier opción. Google te pedirá permisos (es un script que usa el servicio de Administración de Directorio).
3.  **Configura:** Sigue las instrucciones de la hoja "Directorio" para empezar a gestionar tus grupos.

---

### 🤝 Contribuciones

¿Has encontrado un bug o tienes una idea genial? Siéntete libre de abrir un *Issue* o enviar un *Pull Request*. Eso sí, el logo del pulpo es sagrado.

### ✍️ Autoría y agradecimientos

*   **Autor:** Pablo Felip ([@pfelipm](https://twitter.com/pfelipm))
*   **Licencia:** GNU GPL v3. Libertad total para usarlo y modificarlo, siempre que mantengas la autoría y compartas bajo la misma licencia.
*   **Repositorio:** [https://github.com/pfelipm/asistente-grupos](https://github.com/pfelipm/asistente-grupos)

---
<p align="center">Hecho con ❤️, café y un poco de ayuda de Gemini CLI.</p>
