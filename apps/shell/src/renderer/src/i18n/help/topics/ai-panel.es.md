# El panel de IA

Todos los editores pueden invocar el panel de IA: seleccione algo, dé una instrucción y vea el resultado en streaming.

## Asistente compartido de la ventana

Abra el panel compartido con **Asistente de IA** en la barra de pestañas de la ventana principal. Al principio está cerrado. Cree o elija una conversación y seleccione el documento que desea editar y los ámbitos de referencia entre los documentos HWP/HWPX, DOCX/ODT o Markdown/texto abiertos en esa ventana. Cambiar de pestaña o contraer el panel conserva los borradores y las tareas en curso. **Detener** finaliza una tarea. Escribir o cambiar de conversación no envía solicitudes. Los paneles de IA de los editores siguen disponibles.

Si el destino es un documento de Word, Resumir, Mejorar, Formato y Mapa mental rellenan la entrada. Pulsa Enviar para ejecutar; Mejorar usa el texto seleccionado originalmente, si hay una selección. Los mapas permanecen en su conversación y permiten contraer ramas y descargar PNG/SVG. Al pulsar un tema se prepara una pregunta sobre el mismo destino; si cambiaste de destino, vuelve a seleccionar primero el documento original.

Si el destino es un documento Word, el texto largo aparece mientras se escribe. Detener en el progreso de escritura permite conservar o descartar el texto parcial. Detener junto al campo de entrada termina toda la ejecución y elimina el borrador incompleto. La decisión permanece en la conversación original. El texto terminado o conservado se puede revertir con la opción de deshacer esta ejecución.

En la ventana principal, las acciones de IA de Word desde la cinta, los menús, F7 y el menú contextual abren una nueva conversación compartida con el documento y la selección pulsados. Se conservan las confirmaciones y la ejecución inmediata; abrir o contraer el asistente mantiene la conversación actual. Las ventanas separadas o independientes, un anfitrión aún no preparado usan el panel anterior, accesible desde su barra contraída.

En la ventana principal, añadir una edición de selección de Word a la cola abre la conversación compartida del documento. La tarjeta permite localizar pasajes, editar instrucciones, eliminar elementos o confirmar que se descarta toda la cola. Al enviar, el lote pasa a una nueva conversación compartida; si la conexión está bloqueada, queda como borrador. Un fallo de transferencia conserva la cola y un documento modificado no usa posiciones antiguas. Las ventanas separadas e independientes conservan el panel anterior.

En la ventana principal, una copia de plantilla de una habilidad de Word transfiere una sola vez la solicitud y las etiquetas de habilidades a una nueva conversación compartida. Si la conexión está bloqueada, la solicitud queda allí como borrador. Las copias de proyectos usan el contexto de su proyecto original y se guardan en su carpeta con un nombre disponible tras completarse correctamente. Las vistas previas de plantillas en blanco y los anfitriones compartidos no disponibles usan el panel habitual del editor.

## Abrirlo y usarlo

- El panel necesita una IA conectada. Si no hay ninguna, muestra las tarjetas de conexión en lugar de un error; vea [Conexión por suscripción](help://subscription).
- Entradas: el **botón de IA** en la cinta de opciones de cada editor, **Pedir a la IA** en los menús contextuales, o Preguntar a la IA en la barra de marcado.
- Describa la tarea en lenguaje natural (reescribir esto / convertir esta columna en porcentajes / rehacer el diseño de esta página…) y pulse Intro.
- Las respuestas se renderizan **en streaming**; cuando la IA necesita herramientas (leer el documento, editarlo), las ejecuta y sigue hasta terminar.
- **Detener**: interrumpa la vuelta en curso en cualquier momento.
- **Cola**: mientras se genera una respuesta, Enter pone tu mensaje en cola en lugar de descartarlo — una franja sobre el campo de entrada los cuenta; ábrela para editar, quitar, reordenar, pausar o vaciar. El siguiente mensaje se envía solo al terminar el turno; Esc sigue deteniendo el turno.

## Qué IA responde

- La lista de modelos de la parte inferior del panel muestra la IA en uso y permite elegir otra: primero sus suscripciones y después los proveedores con clave que haya configurado.
- Cuando ha usado el 80 % o más del límite de una suscripción, aparece en el panel una línea como «85 % usado · se restablece a las 15:00».
- Si una solicitud se detiene —se agotó el límite, o la sesión ha caducado—, el panel dice por qué y no cambia a otra IA por su cuenta. Cuando su otra suscripción está conectada, un botón **Continuar con …** la ofrece.

## Qué sabe hacer

- **Docs**: reescribir/ampliar/traducir/resumir, insertar tablas e imágenes, ajustar el formato; cada vuelta guarda primero una instantánea.
- **Sheets**: fórmulas, relleno de datos, transformaciones masivas, formato.
- **Slides**: generación de una presentación completa, ajuste del diseño, reescritura de los textos.
- **PDF**: preguntas y resúmenes sobre el texto o las páginas seleccionadas.
- **Markdown / HTML**: reescribir, ampliar, traducir.
- **Inicio ▸ Ask AI** responde desde tus archivos destacados y los cita — el panel de base de conocimiento tiene su propia sección en el capítulo Inicio, y el panel de IA de Docs lleva la misma capacidad.

## Adjuntar documentos abiertos

En el panel de IA de los editores Docs, Sheets, Slides, Hangul, HTML y mapas mentales puedes adjuntar como material otros documentos que tengas abiertos en GenOffice o que hayas abierto hace poco. Un documento adjunto se trata exactamente igual que un archivo elegido en tu PC.

- **`@` en el campo de entrada**: escribe `@` al principio del campo o después de un espacio para abrir una lista de documentos. Si la ventana está dividida, primero aparece el documento del otro panel y luego los demás documentos abiertos y los recientes; el documento en el que estás no aparece. Fíltrala con unas letras del nombre (se admiten espacios) y elige con Enter o con un clic: `@nombre del documento` se escribe en el campo y el documento queda adjunto. **Buscar más…**, al final, abre la ventana de selección de documentos descrita abajo.
- **Adjuntar desde documentos de GenOffice**: este botón, junto al botón de adjuntar, abre la misma lista en una ventana donde puedes marcar varios documentos y adjuntarlos de una vez.
- Los documentos con cambios sin guardar y los documentos nuevos que nunca se guardaron aparecen atenuados y no se pueden elegir. Guárdalos y vuelve a elegirlos: la IA lee el archivo guardado.
- Si dos documentos se llaman igual pero están en carpetas distintas, se añade el nombre de la carpeta inmediatamente superior, como en `@plan.docx(2.º trimestre)`, y ese nombre de carpeta también se le pasa a la IA.
- Los editores de PDF y Markdown todavía no tienen adjuntos, así que ninguna de las dos formas está disponible allí.

## Proyectos

- Haz clic en el nombre del proyecto, a la izquierda y debajo del campo de entrada, para elegir un proyecto (una carpeta de tu PC). Así la IA responde conociendo los archivos, las instrucciones y la memoria de esa carpeta, y el chat se guarda en ella para que puedas seguir en otro editor. Consulta [Proyectos](help://projects).
- Con un proyecto elegido, el título del chat actual aparece arriba del panel. Haz clic en él para ver la lista de chats del proyecto y **Nuevo chat**.
- En la lista de `@` del campo de entrada aparece «Archivos de este proyecto» después de los documentos abiertos.
- Una petición con un proyecto elegido también envía a la IA que estés usando las instrucciones (junto con `AGENTS.md` y `CLAUDE.md` de la raíz de la carpeta, que puedes desactivar en la página del proyecto), la memoria y la lista de nombres de archivo de ese proyecto, un breve fragmento de cada archivo que la IA encuentre al buscar en la carpeta, además del texto de cualquier archivo que la IA lea.

## Retroceso y seguridad

- El panel de Docs conserva una **lista de versiones**: una instantánea por vuelta, puede volver a cualquiera de ellas y ese propio retroceso se puede deshacer con Ctrl+Z. Las instantáneas sobreviven a reabrir el documento.
- Las ediciones de la IA pasan por la misma canalización de edición que las manuales (se pueden deshacer y requieren guardar): nada omite su confirmación de guardado.

## Ocultar una selección a la IA

- En Slides, HTML y Markdown una selección puede quitarse de la vista del modelo: clic derecho ▸ **Ocultar la selección a la IA**, ponle nombre al marcador, y el modelo lee `{{nombre}}` en su lugar. (Markdown arma la opción con el interruptor de ojo tachado de su cinta; la de HTML vive en el panel de código.)
- El archivo conserva el texto real — guardar, exportar e imprimir quedan intactos, y la marca sobrevive a reabrir. La misma opción deshace el ocultamiento (Markdown tiene además **Volver a mostrar la selección a la IA**).
- HTML oculta contenidos de elementos, valores de atributos y literales de cadena de los scripts; la página renderizada nunca cambia.
- Slides también oculta imágenes, vídeos y audios enteros, y una diapositiva con contenido oculto nunca se envía como captura: la revisión automática de diseño se salta la página y lo avisa.
- Esto vale solo para la IA de la aplicación — la CLI, las herramientas MCP y la exportación sin interfaz leen el archivo directamente y siguen viendo el texto.

## Privacidad

- Las instrucciones y el contenido pertinente del documento van a la **IA que esté usando**: a Anthropic u OpenAI con su propia cuenta cuando usa una suscripción, o al proveedor cuya clave introdujo. Si no hay ninguna IA conectada, no se envía nada.
- Que el contenido de una suscripción se use para entrenar IA depende de los ajustes de su propia cuenta en esa empresa.
- Los archivos locales no se suben a ningún otro sitio; las claves de API se guardan en el archivo de ajustes de la aplicación en este equipo y se envían solo en las cabeceras de la petición.
