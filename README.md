
## __DESCRIPCIÓN__:

Este aplicativo desarrollado en Python __automatiza la captura del contenido de rangos de celdas Excel y su colocación en slides PowerPoint__, con coordenadas y dimensiones configurables por el usuario.

En ejecuciones posteriores, __cualquier cambio en los Excel de origen se refleja automáticamente en los pantallazos generados__, manteniendo siempre la disposición configurada.

No todo el reporting se realiza con herramientas como Power BI, Tableau o Qlik.
__En muchas empresas, por limitaciones de licencias, el reporting sigue basándose en Excel y PowerPoint__.

PowerPoint permite vincular presentaciones a Excel, pero cuando los archivos son pesados (por volumen de datos y/o complejidad de fórmulas), esta opción puede resultar inviable. En estos casos, los usuarios suelen recurrir a capturas manuales.
__Este aplicativo nace para automatizar estas tareas manuales y repetitivas__.

El proceso de captura y generación de PowerPoint __se ejecuta en segundo plano__, permitiendo al usuario continuar trabajando con otros Excel, PowerPoint u otras aplicaciones.

El sistema permite configurar múltiples reports, que se almacenan en un __sistema SQLite creado y gestionado desde el propio aplicativo__.

El aplicativo incluye un sistema de __logs de warnings y errores__ que se genera (en caso de haberlos) al finalizar cada proceso.

-------------------------------------------------------------------------------------------------------------------------------------

## Muestra

Tras __importar un sistema SQLite__ y mediante una interfaz gráfica intuitiva, el usuario puede:
  * configurar múltiples reports PowerPoint.
  * asociar varios Excel de origen a cada report.
  * definir múltiples rangos por Excel de origen.

<img width="927" height="560" alt="image" src="https://github.com/user-attachments/assets/016a7c07-7d4f-4bd1-8db4-7938e27efa42" />

Tras estas configuraciones y guardalas en el sistema SQlite, el usuario configura la colocación de los pantallazos en sus reports PowerPoint. 

<img width="961" height="625" alt="image" src="https://github.com/user-attachments/assets/8aaf700a-9cc7-46a4-99a7-e3e5de177c75" />

Se genera un PowerPoint donde:
  * los pantallazos se colocan en las slides configuradas (inicialmente en la __esquina superior izquierda__).
  * el usuario __ajusta posición y tamaño manualmente__ de cada pantallazo
  * el aplicativo __recupera estas coordenadas y dimensiones y las guarda__ en el sistema SQLite

<img width="1154" height="451" alt="image" src="https://github.com/user-attachments/assets/a69b4992-c84d-41d6-b929-68991f9b25f9" />

<img width="1157" height="383" alt="image" src="https://github.com/user-attachments/assets/aa3b1678-73a4-404b-9c17-927da67ab212" />

Tras guardar las coordenadas y dimensiones en el sistema SQlite, el __report PowerPoint queda completamente configurado__ y listo para ejecuciones futuras.

En ejecuciones posteriores, el usuario selecciona el report y se genera automáticamente un PowerPoint donde los pantallazos se posicionan con las coordenadas y dimensiones previamente definidas, reflejando los cambios en los datos de los Excel de origen.

<img width="1326" height="326" alt="image" src="https://github.com/user-attachments/assets/e4f30488-55b7-47a5-9334-4cfb7e38cd65" />


El aplicativo permite __configuraciones incrementales__ de sus reports.
En los reports cuyos rangos de celdas ya se guardaron sus coordenadas y dimensiones, cuando el usuario da de alta
nuevos rangos de celdas __no requiere volver a colocar y dimensionar todo lo anterior ya configurado__:
   * los nuevos rangos de celdas configurados se colocan en la esquina superior izquierda.
   * los rangos de celdas ya configurados mantienen las posiciones y tamaños definidos.
-------------------------------------------------------------------------------------------------------------------------------------
## __DEPENDENCIAS EXTERNAS__:

El aplicativo utiliza __Poppler__ para la conversión de archivos PDF a imágenes (PNG).

Poppler es una librería open-source ampliamente utilizada para el procesamiento de PDF.

La instalación de Poppler se detalla en la __guía de usuario__ (subcarpeta v1.0/documentacion_otra).

Enlace de descarga para Windows (también incluido en la guía de usuario):

https://github.com/oschwartz10612/poppler-windows/releases

El aplicativo permite configurar múltiples rutas de Poppler (locales o en unidades compartidas) que se almacenan en el sistema SQLite.
Al importar dicho sistema, se selecciona automáticamente la primera ruta accesible.

Esto permite adaptar el aplicativo a distintos entornos (semi multiusuario).


## __INTERFAZ GRÁFICA (GUI)__:

La GUI del aplicativo ha sido desarrollada utilizando mi otro proyecto publicado hace unos meses __python_tools_modulares (tkinter_utils, v1.1)__:

https://github.com/JulienBott/python_tools_modulares.git

Es un sistema hibrido entre widgets nativos y otros personalizados de la libreria tkinter. Facilita la configuración de GUI's de forma limpia, dinámica y escalable mediante el uso de kwargs.

Aunque se ha usado con ligeras adaptaciones (ver módulo APP_PPTX_2_GUI_UTILS.py) con el fin de poder generar messagebox personalizados solo para el uso del aplicativo.

## __LIMITACIONES ACTUALES Y FUTURAS MEJORAS__:

__LIMITACIONES__

Las limitaciones que se exponen a continuación estan previstas ser correjidas en futuras versiones:

 * el aplicativo permite configurar múltiples rutas (locales o en unidades compartidas) tanto para Poppler como para la generación temporal de archivos (PDF / PNG).
   En ambos casos, al importar el sistema SQLite, se selecciona automáticamente la primera ruta accesible.

     * __Proceso de generación de Powerpoint__:
       
       Durante cada proceso, se crean subcarpetas temporales (con timestamp) donde se almacenan las capturas de rangos de celdas Excel antes de su inserción en PowerPoint. 
       Estas subcarpetas se eliminan automáticamente al finalizar el proceso.
       
       Si varios usuarios utilizan simultáneamente el mismo sistema SQLite y el aplicativo asigna la misma ruta de trabajo, existe una probabilidad (baja pero     posible) de conflicto si se generan subcarpetas con el mismo nombre en paralelo.
       
     * __Acceso a capturas de muestra en la ventana de configuración de los rangos de celdas__:
       
       En la ventana de configuración de rangos de celdas, el usuario puede acceder a capturas de ejemplo. Estas también requieren crear subcarpetas temporales, que actualmente no se eliminan automáticamente.

       La resolución de este punto requiere una mejora en el sistema de GUI basado en __tkinter_utils__, incorporando gestión de cierre de ventanas (protocol) para ejecutar rutinas de limpieza.

* limitación en la __apertura de Excel de origen__:

  Durante el proceso de captura de rangos de celdas, el usuario debe configurar un tiempo de espera para la apertura de cada Excel de origen, especialmente en archivos voluminosos donde las fórmulas requieren tiempo para actualizarse. Esto implica que el usuario debe conocer previamente el comportamiento de sus Excel.

  Como mejora futura, se prevee incorporar un mecanismo que detecte automáticamente cuándo el Excel está completamente cargado y listo para iniciar la captura de rangos de celdas.

__FUTURAS MEJORAS__

* evolución del __sistema SQLite hacia un entorno multiusuario real__, permitiendo la gestión compartida de reports.

* ejecución en bloque de múltiples reports PowerPoint (actualmente se procesan de uno en uno).


## __CONTENIDO DEL REPOSITORIO GITHUB__:

Nada más acceder al repositorio, se encuentra el __README__ que estas leyendo ahora mismo acompañado de un __contrato MIT Licence__ donde autorizo cualquier tipo de uso del aplicativo y de su código asociado sea a nivel particular o empresarial siempre y cuando se me reconozca la autoria original del aplicativo. Dicho contrato de MIT Licence tiene clausulas añadidas.

Se encuentra una carpeta llamada __v1.0__ y dentro de la misma se encuentran las subcarpetas siguientes:

* __codigo__

  Contiene los 3 módulos de código Python:

  * __APP_PPTX_1_GUI__
  * __APP_PPTX_2_GUI_UTILS__
  * __APP_PPTX_3_BACK_END__
 
* __documentacion_otra__
    
   Ahi se encuentran 2 ficheros:
  * __GUI_USUARIO_V1.0__: es una guia que explica paso a paso como operar con el aplicativo.
  * __MANUAL_PARA_COMPILAR_EL_APP_EN_EJECUTABLE_V1.0__: es un manual que explica como crear un venv (entorno virtual) para luego compilar el aplicativo en ejecutable (exe).


* __documentacion_tecnica__

  Actualmente la carpeta se encuentra vacía.
  
  En futuras versiones se incorporará el diseño funcional del aplicativo, donde se incluirá:
    * descripción de la arquitectura.
    * listado de rutinas y funciones con detalle de argumentos (args y kwargs) y retornos (en caso de funciones).
    * diagnóstico de dependencias de objetos.
    * explicación de las clases propias que componen la GUI.
    * ejemplo detallado del proceso de captura de rangos de celdas mediante Poppler y OpenCV.


* __templates__

  Contiene los archivos que son necesarios para poder ejecutar el aplicativo:

  Son esenciamente archivos .ico, .png, un fichero excel y un pdf.

  Contiene, asimismo, el fichero __APP_SEMI_AUTOMATIZADOR_POWERPOINT.spec__ que se ha de usar para poder compilar el aplicativo en .exe (ver el manual __MANUAL_PARA_COMPILAR_EL_APP_EN_EJECUTABLE_V1.0__)
  
## __REQUISITOS SISTEMA Y LIBRERIAS PYTHON__

El aplicativo se ha desarrollado y probado en entorno Windows (10) usando la versión 3.9.5 de Python. No se ha probado con otros sistemas operativos por lo que podria haber errores.

Librerias que requieren instalación (pip install):

<img width="232" height="252" alt="image" src="https://github.com/user-attachments/assets/68c717f8-1b25-4f43-92fb-6854822edf36" />

Librerias nativas Python:

<img width="129" height="433" alt="image" src="https://github.com/user-attachments/assets/9d1ba2ff-dbce-4c5e-b2b4-b68630263adc" />



__EJECUCIÓN DEL APLICATIVO DESDE LA INTERFAZ DE PROGRAMACIÓN__:

Para ejecutar el aplicativo desde la consola de la interfaz de programación que se use hay que guardar en una misma carpeta en el pc los archivos de las carpetas __codigo__ y __templates__ mencionadas en este README. Una vez guardados, hay que __ejecutar el módulo APP_PPTX_1_GUI.py__.


## __Actualización 06/04/2026__

Estado inicial de la versión 1.0 tras su publicación.

Tareas pendientes (en curso):
 * elaboraración del diseño funcional.
 * corrección de las limitaciones expuestas y desarrollo  de las futuras mejoras comentadas.

## __Actualización 02/05/2026__

Las tareas pendientes expuestas en la actualización del __06/04/2026__ quedan pospuestas.
 
Actualmente, mi foco  está en el desarrollo de un nuevo proyecto que publicare en breve en Github.
 
Una vez finalizado dicho proyecto realizare los upgrades comentados.




