# Sistema de incurridos
Facilita contabilizar el tiempo dedicado a las tareas de la semana en horas.

<br/>

<img src="https://github.com/bitasuperactive/SistemaDeIncurridos/blob/main/doc/demostracion.gif"/>


## Descripción
<br/>
<img src="https://github.com/bitasuperactive/SistemaDeIncurridos/blob/main/doc/descripcion.png">

<h3>:one: Cronómetro</h3>
Contabiliza el tiempo transcurrido.
<h3>:two: Botón de incurrir</h3>
Inicia el cronómetro para la tarea seleccionada.
<h3>:three: Botón de terminar</h3>
Finaliza el cronómetro y registra el tiempo incurrido de la tarea seleccionada en la celda correspondiente.
<h3>:four: Seleccionador de la tarea a incurrir</h3>
Solo aparecerán aquellas tareas introducidas en el cuadro inferior (5).
<h3>:five: Tareas a contabilizar</h3>
Para añadir nuevas tareas se deben escribir en la columna A del worksheet "Tareas".
<h3>:six: Suma total del tiempo incurrido</h3>
<h3>:seven: Semana del año</h3>
Cuando la semana termina, se borran los tiempos incurridos.
<h3>:hash: Resto de celdas sin resaltar</h3>
Almacenan el tiempo incurrido.

<br/>


## Funcionalidades
 - <b>Incurrimiento facilitado:</b> No es necesario pulsar en terminar cada vez que se finalice una tarea, simplemente con cambiar la tarea incurrida a otra distinta se restablecerá el cronómetro registrando el tiempo incurrido.  
 Nota: Al abrir el desplegable de la tarea a incurrir el cronómetro se congela.
 - <b>Recordatorio:</b> Lanza un mensaje de texto al usuario si lleva más de 10 minutos sin incurrir ninguna tarea.
 - <b>Guardado de seguridad:</b> Si se cierra el workbook mientras se está incurriendo una tarea, esta se registrará en su celda correspondiente evitando la pérdida del conteo.


## Guías

<details><summary> <h3>Cómo añadir nuevas tareas a incurrir</h3></summary>
<img src="https://github.com/bitasuperactive/SistemaDeIncurridos/blob/main/doc/como_a%C3%B1adir_nuevas_tareas.gif">
</details>

<details><summary> <h3>Cómo incurrir esas nuevas tareas</h3></summary>
<img src="https://github.com/bitasuperactive/SistemaDeIncurridos/blob/main/doc/como_incurrir_nuevas_tareas.gif">
</details>
