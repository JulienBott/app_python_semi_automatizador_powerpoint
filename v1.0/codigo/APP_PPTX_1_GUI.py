import tkinter as tk
from tkinter import filedialog as fd
from threading import Thread
import time
import os

import APP_PPTX_2_GUI_UTILS as mod_utils
import APP_PPTX_3_BACK_END as mod_back_end



###################################################################################################################################################
###################################################################################################################################################
# clase gui_ventana_inicio
###################################################################################################################################################
###################################################################################################################################################

class gui_ventana_inicio:

    def __init__(self, master, **kwargs_config_gui):


        #se inicializan atributosde la presente clase
        self.master = master
        self.kwargs_config_gui = kwargs_config_gui

        self.clase_gui_actual_nombre = self.__class__.__name__
        self.kwargs_gui_ventana_inicio = {key: dicc for key, dicc in self.kwargs_config_gui[self.clase_gui_actual_nombre].items() if key != "dicc_config_root"}

        self.sistema_sqlite_importado = False
        self.id_pptx_ventana_inicio_selecc = None


        #se crea la variable global global_dicc_tablas_y_campos_sistema (modulo back-end)
        mod_back_end.def_varios("DICC_TABLAS_Y_CAMPOS_SISTEMA")


        #se insertan los widgets dentro del frame_inicio y se almacenan en el diccionario dicc_gui_frame_widgets_objetos
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_gui_ventana_inicio_frame_widgets_objetos = {}
        for frame_contenedor in self.kwargs_gui_ventana_inicio.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_ventana_inicio_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_ventana_inicio.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_ventana_inicio_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_ventana_inicio_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_ventana_inicio[frame_contenedor].items() if widget != "frame"}

            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_ventana_inicio_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ", "").strip()

                widget_objeto = (mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget_ajust, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]
                                else
                                mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "scrolledtext_propio"
                                else
                                mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust == "treeview_propio"
                                else
                                mod_utils.frame_con_scrollbar(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "frame_con_scrollbar"
                                else None)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                self.dicc_gui_ventana_inicio_frame_widgets_objetos.update({frame_contenedor_widget:
                                                                                                    {"widget_objeto": widget_objeto
                                                                                                    , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                                    }
                                                                        })
            

        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.label_warning_path_poppler = self.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_01"]["widget_objeto"]
        self.scrolledtext_ruta_sistema_sqlite = self.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_09"]["widget_objeto"]
        self.treeview_id_pptx = self.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_13"]["widget_objeto"]
        self.combobox_accion_pptx = self.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_15"]["widget_objeto"]
        self.scrolledtext_desc_accion_pptx = self.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_17"]["widget_objeto"]


        #se recuperan los height de los scrolledtext desde los parametros kwargs de la presente clase
        self.scrolledtext_ruta_sistema_sqlite_height = self.kwargs_gui_ventana_inicio["frame_sistema"]["WIDGET_09"]["kwargs_config"].get("height", 1)
        self.scrolledtext_desc_accion_pptx_height = self.kwargs_gui_ventana_inicio["frame_pptx"]["WIDGET_17"]["kwargs_config"].get("height", 1)




    def def_gui_ventana_inicio_threads(self, proceso_selecc):
        #rutina para ejecutar todos los procesos del app para la clase gui_ventana_inicio
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_back_end.global_proceso_en_ejecucion == "SI":
            mensaje = "Espera a que acabe el proceso actualmente en ejecución."
            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))
        else:
            Thread(target = self.def_gui_ventana_inicio_procesos, args = (proceso_selecc,)).start()




    def def_gui_ventana_inicio_procesos(self, proceso_selecc):
        #rutina que permite ejecutar los procesos de la gui tanto de sistema como los asociados a los pptx


        #se comprueba si algun proceso esta seleccionado (tanto de sistema sqlite como de interaccion pptx)
        #en caso contrario se cancela la rutina
        if len(proceso_selecc) == 0:
            mensaje = "No has seleccionado ningún proceso."
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            #se reinicia la variable global global_proceso_en_ejecucion a NO
            mod_back_end.global_proceso_en_ejecucion = "NO"

            return
        



        ######################################################################
        # GUIA USUARIO
        ######################################################################
        if proceso_selecc == mod_back_end.opcion_gui_guia_usuario:

            mensaje = "Se descargara el manual de usuario (pdf) y se abrira en la ruta que especifiques a continuación.\n\nDeseas continuar?"
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

            if msg.valor_boton_pulsado:

                ruta_carpeta_guia_usuario = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres que se guarde la guia de usuario:")

                if ruta_carpeta_guia_usuario:

                    self.master.widget_objeto.config(cursor = "wait")

                    mod_back_end.def_varios_gui_ventana_inicio("DESCARGA_GUIA_USUARIO"
                                                            , ruta_carpeta_guia_usuario = ruta_carpeta_guia_usuario)

                    self.master.widget_objeto.config(cursor = "")

                    time.sleep(2)#para dar tiempo al pdf que se abra y no se ejecute antes del messagebox en la gui

                    mensaje = f"Guia de usuario descargada en: '{ruta_carpeta_guia_usuario}'."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))


        ######################################################################
        # opciones asociadas a acciones de sistema
        ######################################################################
        elif proceso_selecc in [mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["IMPORTAR_SISTEMA"]["OPCION"]
                                    , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["CREAR_SISTEMA"]["OPCION"]
                                    , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["AGREGAR_RUTA_POPPLER"]["OPCION"]
                                    , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["AGREGAR_RUTA_LOCAL"]["OPCION"]
                                    , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["DESCARGAR_SISTEMA_A_XLS"]["OPCION"]]:
            
            

            (tipo_messagebox
            , mensaje_gui) = mod_back_end.def_varios("MESSAGEBOX_PROCESOS_APP", opcion_proceso = proceso_selecc)


            msg = False
            if tipo_messagebox == "showerror":
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
            
            elif tipo_messagebox == "askokcancel":
                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
            

            if msg:

                if msg.valor_boton_pulsado:
                        
                    #se inician los filedialogs y se construye el diccionario kwargs_rutina_sistema_sqlite que se usa como kwargs
                    #en la llamada a la rutina def_config_sistema_sqlite (mod back_end) mas abajo
                    #se informa mensaje_cancelacion_proceso con el mensaje que avisa en la GUI pq el proceso se cancela
                    #(se cancela tan solo si el usuario en los filedialogs que se generan los cierra sin seleccionar nada)
                    #
                    #IMPORTANTE los nombres de variables que se usan para hacer la llamada al filedialog (sea filename o directory)
                    #han de estar escritos exactamente como los kwargs que se usan en la rutina def_config_sistema_sqlite (mod back_end)
                    kwargs_rutina_sistema_sqlite = {}
                    mensaje_cancelacion_proceso = None


                    if proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["CREAR_SISTEMA"]["OPCION"]:

                        directorio_sqlite = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres que se cree el sistema de base de datos sqlite:")
                        directorio_poppler = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona el directorio que contiene Poppler (la capreta deba llamarse 'bin').")
                        directorio_local_screenshots = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona que directorio quieres que se use para guardar temporalemente los pantallazos de rangos de celda excel cuando ejecutes acciones pptx.")
                        directorio_log_errores = directorio_sqlite

                        if directorio_sqlite and directorio_local_screenshots:
                            kwargs_rutina_sistema_sqlite.update({"directorio_sqlite": directorio_sqlite
                                                                , "directorio_poppler": directorio_poppler
                                                                , "directorio_local_screenshots": directorio_local_screenshots
                                                                })
                        else:
                            mensaje_cancelacion_proceso = "Se ha cancelado el proceso porque has cerrado una o las dos ventanas de dialogo, que se generaron, sin seleccionar nada."




                    elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["IMPORTAR_SISTEMA"]["OPCION"]:
                          
                        path_sqlite = fd.askopenfilename(parent = self.master.widget_objeto, title = "Selecciona la ruta de la base de datos sqlite a importar en memoria ram del PC:", filetypes = [("Base de datos SQLite", "*.db")])
                        directorio_log_errores = os.path.normpath(os.path.dirname(path_sqlite))

                        if path_sqlite:
                            kwargs_rutina_sistema_sqlite.update({"path_sqlite": path_sqlite})
                        else:
                            mensaje_cancelacion_proceso = "Se ha cancelado el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."




                    elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["AGREGAR_RUTA_POPPLER"]["OPCION"]:
                          
                        ruta_add = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona la ruta de la carpeta 'bin' de Poppler que deseas agregar:")
                        directorio_log_errores = mod_back_end.global_ruta_local_config_sistema_sqlite

                        if ruta_add:
                            kwargs_rutina_sistema_sqlite.update({"ruta_add": ruta_add})
                        else:
                            mensaje_cancelacion_proceso = "Se ha cancelado el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."




                    elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["AGREGAR_RUTA_LOCAL"]["OPCION"]:
                          
                        ruta_add = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona la ruta local que deseas agregar:")
                        directorio_log_errores = mod_back_end.global_ruta_local_config_sistema_sqlite

                        if ruta_add:
                            kwargs_rutina_sistema_sqlite.update({"ruta_add": ruta_add})
                        else:
                            mensaje_cancelacion_proceso = "Se ha cancelado el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."






                    elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_SISTEMA"]["DESCARGAR_SISTEMA_A_XLS"]["OPCION"]:

                        directorio_excel = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres guardar el excel de configuración:")
                        directorio_log_errores = directorio_excel

                        if directorio_excel:
                            kwargs_rutina_sistema_sqlite.update({"directorio_excel": directorio_excel})
                        else:
                            mensaje_cancelacion_proceso = "Se ha cancelado el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."



                    #se ejecuta el proceso solo en caso de que kwargs_rutina_sistema_sqlite no este vacio
                    #si lo esta es pq el usuario cuando se han generado los filedialog los ha cerrado en vez 
                    #de seleccionar un directorio o un fichero segun el tipo de filedialog
                    if len(kwargs_rutina_sistema_sqlite) == 0:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_cancelacion_proceso}))

                    else:

                        #se ejecuta la rutina def_config_sistema_sqlite
                        self.master.widget_objeto.config(cursor = "wait")

                        mensaje_warning_gui = mod_back_end.def_config_sistema_sqlite(proceso_selecc, **kwargs_rutina_sistema_sqlite)

                        self.master.widget_objeto.config(cursor = "")


                        if mensaje_warning_gui is not None:
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_warning_gui}))

                        else:

                            #en caso de errores se genera un log en formato .txt
                            if len(mod_back_end.global_lista_dicc_errores) != 0:

                                mod_back_end.def_varios("GENERAR_LOG_WARNING_ERRORES_PROCESOS_APP"
                                                        , directorio_log_errores_warning = directorio_log_errores
                                                        , opcion_warning_errores = "ERRORES")

                                mensaje = f"Se han localizado errores, se ha generado un log de errores en la ruta siguiente: {directorio_log_errores}"
                                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))


                            else:

                                #en caso de que el proceso ejecutado sea IMPORTAR_SISTEMA se informa en la GUI:
                                # --> el scrolledtext con la ruta del sistema sqlite importado
                                # --> el treeview con el ID_PPTX y su descripción corta
                                if mod_back_end.def_varios_gui_ventana_inicio("KEY_PROCESO_SISTEMA", opcion_combobox_sistema = proceso_selecc) == "IMPORTAR_SISTEMA":

                                    #ruta sistema sqlite
                                    self.scrolledtext_ruta_sistema_sqlite.config_atributos(**{"bloquear": False})

                                    self.scrolledtext_ruta_sistema_sqlite.modificaciones("borrar_contenido_y_tags")

                                    self.scrolledtext_ruta_sistema_sqlite.modificaciones("agregar_solo_contenido_desde_string"
                                                                                        , string_texto_informar = path_sqlite
                                                                                        , height_scrolledtext = self.scrolledtext_ruta_sistema_sqlite_height)
                            
                                    self.scrolledtext_ruta_sistema_sqlite.config_atributos(**{"bloquear": True})


                                    #treeview de id pptx
                                    self.treeview_id_pptx.acciones("actualizar_desde_df", df_datos = mod_back_end.global_df_treeview_id_pptx)


                                    #se actualiza el atributo self.sistema_sqlite_importado de la clase
                                    self.sistema_sqlite_importado = True



                                mensaje_fin_proceso = ("Sistema sqlite creado."
                                                        if mod_back_end.def_varios_gui_ventana_inicio("KEY_PROCESO_SISTEMA", opcion_combobox_sistema = proceso_selecc) == "CREAR_SISTEMA"
                                                        else
                                                        "Sistema sqlite importado."
                                                        if mod_back_end.def_varios_gui_ventana_inicio("KEY_PROCESO_SISTEMA", opcion_combobox_sistema = proceso_selecc) == "IMPORTAR_SISTEMA"
                                                        else
                                                        "Ruta agregada al sistema sqlite."
                                                        if mod_back_end.def_varios_gui_ventana_inicio("KEY_PROCESO_SISTEMA", opcion_combobox_sistema = proceso_selecc) in ["AGREGAR_RUTA_POPPLER", "AGREGAR_RUTA_LOCAL"]
                                                        else
                                                        None
                                                        )
                                
                            
                                if mensaje_fin_proceso is not None:
                                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje_fin_proceso}))
                                




                                #para DESCARGAR_SISTEMA_A_XLS el mesaje y tipo de messagebox son 
                                if mod_back_end.def_varios_gui_ventana_inicio("KEY_PROCESO_SISTEMA", opcion_combobox_sistema = proceso_selecc) == "DESCARGAR_SISTEMA_A_XLS":

                                    if mod_back_end.global_df_treeview_id_pptx is None:

                                        mensaje = "No has importado previamente en el app el sistema sqlite."
                                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                                    else:
                                        mensaje = (f"Proceso '{proceso_selecc}' ejecutado.\n\nEl excel de configuración esta vacio (configuraciones de rangos de celdas)." if len(mod_back_end.global_df_treeview_id_pptx) == 0
                                                else f"Proceso '{proceso_selecc}' ejecutado.")
                                        
                                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))




        ######################################################################
        # VENTANA CONFIG ID PPTX
        ######################################################################
        elif proceso_selecc == mod_back_end.opcion_gui_config_id_pptx:


            #se crea la variable abrir_ventana_config_id_pptx para determinar si abrir o no la ventana de configuracion de los id pptx
            #(depende de si hay errores en la descarga de los screenshots en la ruta local para servir de muestra)
            #
            #se crea la variable descargar_screenshots_muestra_en_png que se pasa como parametro de la clase gui_config_id_pptx
            #para indicar si el usuario ha optado por descargar los screenshots en png para que salgan en las muestras
            abrir_ventana_config_id_pptx = False
            descargar_screenshots_muestra_en_png = False

            (tipo_messagebox
            , mensaje_gui) = mod_back_end.def_varios("MESSAGEBOX_PROCESOS_APP", opcion_proceso = "VENTANA_CONFIG_ID_PPTX")

            msg = False
            if tipo_messagebox == "showerror":
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
            
            elif tipo_messagebox == "askokcancel_con_o_sin_muestra_screenshots":
                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
            

            if msg:

                if msg.valor_boton_pulsado in [True, False]:

                    #se crea la carpeta con los screenshots de muestra guardados en el sistema sqlite
                    if msg.valor_boton_pulsado:

                        abrir_ventana_config_id_pptx = True
                        descargar_screenshots_muestra_en_png = True

                        #se ejecuta el proceso de descarga de los png
                        self.master.widget_objeto.config(cursor = "wait")
                        mod_back_end.def_config_sistema_sqlite("DESCARGA_SCREENSHOTS_PNG_PARA_MUESTRA_GUI_CONFIG_ID_PPTX")
                        self.master.widget_objeto.config(cursor = "")


                        #se se han localizado errores se notifican y se crea log de errores (en la ruta local configurada en el sistema sqlite)
                        #pero se habilita el acceso a la ventana de configuracion de los id pptx pero sin acceso a las muestras
                        if len(mod_back_end.global_lista_dicc_errores) != 0:

                            mod_back_end.def_varios("GENERAR_LOG_WARNING_ERRORES_PROCESOS_APP"
                                                    , directorio_log_errores_warning = mod_back_end.global_ruta_local_config_sistema_sqlite
                                                    , opcion_warning_errores = "ERRORES")
                            
                            
                            #se deshabilita el acceso a las muestras en la ventana de configuracion de id pptx
                            descargar_screenshots_muestra_en_png = False


                            mensaje1 = f"Se han localizado errores en la descarga de los pantallazos para usarlos como muestra. Se ha generado un fichero .txt en la ruta: \n{mod_back_end.global_ruta_local_config_sistema_sqlite}.\n\n"
                            mensaje2 = "Aun asi, puedes acceder a la ventana de configuración de los id pptx pero sin acceso a las muestras de pantallazos.\n\nDeseas aun asi acceder a dicha ventana?"

                            msg_2 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje1 + mensaje2}))

                            if msg_2.valor_boton_pulsado:
                                abrir_ventana_config_id_pptx = True

                            elif not msg_2.valor_boton_pulsado:
                                abrir_ventana_config_id_pptx = False



                    elif msg.valor_boton_pulsado == False:

                        abrir_ventana_config_id_pptx = True
                        descargar_screenshots_muestra_en_png = False



                if abrir_ventana_config_id_pptx:
                    #se abre una nueva ventana para configurar los id pptx
                    #se pasa el entorno de la clase gui_ventana_inicio como parametro a la clase gui_config_id_pptx
                    #para poder actualizar el treeview de id pptx cuando se guardan co9nfiguraciones nuevas en gui_config_id_pptx
                    kwargs_gui_config_id_pptx_dicc_config_root = self.kwargs_config_gui["gui_config_id_pptx"]["dicc_config_root"]

                    self.toplevel_gui_config_id_pptx = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_config_id_pptx_dicc_config_root)
                    self.toplevel_gui_config_id_pptx.config_atributos(**kwargs_gui_config_id_pptx_dicc_config_root)

                    gui_config_id_pptx(self.toplevel_gui_config_id_pptx
                                        , descargar_screenshots_muestra_en_png = descargar_screenshots_muestra_en_png
                                        , entorno_clase_gui_ventana_inicio = self
                                        , **self.kwargs_config_gui)




        ######################################################################
        # opciones asociadas a los pptx
        ######################################################################
        elif proceso_selecc in [mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]
                                , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_2"]["OPCION"]
                                , mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["EJECUCION"]["OPCION"]]:
            
            ################################################################
            #si Poppler NO esta instalado en el pc se inhabilitan las acciones pptx
            ################################################################
            if mod_back_end.global_poppler_path is None:
                mensaje1 = "Las acciones asociadas al configuración y ejecución de pptx estan inhabilitadas poque no se ha localizado el ejecutable (.exe) de Poppler en tu equipo.\n\n"
                mensaje2 = "En la guia de usuario se explica como proceder para dicha instalación."

                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showwarning", "mensaje": mensaje1 + mensaje2}))


            ################################################################
            #si Poppler SI esta instalado en el pc se habilitan las acciones pptx
            ################################################################   
            else:

                #se calcula la key asociada al stringvar asociado al combobox de seleccion de procesos pptx
                #(la key es inamovible, su valor es configurable)
                opcion_pptx_selecc = self.combobox_accion_pptx.variable_enlace.get()


                (tipo_messagebox
                , mensaje_gui) = mod_back_end.def_varios("MESSAGEBOX_PROCESOS_APP"
                                                        , opcion_proceso = opcion_pptx_selecc
                                                        , id_pptx_selecc = self.id_pptx_ventana_inicio_selecc)


                msg_1 = False
                if tipo_messagebox == "showerror":
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
                
                elif tipo_messagebox in ["askokcancel", "askokcancel_con_opciones_colocacion_pantallazos"]:
                    msg_1 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
                

                if msg_1:

                    kwargs_rutina_interaccion_pptx = {}

                    #se calcula el kwargs opcion_coordenadas_config_paso_1 segun el boton del messagebox que pulse el usuario
                    #para determinar en el caso de CONFIG_PASO_1 si la colocacion de los screenshots se hace todos en la esquina superior-izquierda
                    #o es una colocacion hibrida que recupera coordenadas y dimensiones ya configuradas para algunos screenshots (configuraciones hechas anteriormente)
                    #o son nuevos screenshots (en este caso se colocan en la esquina superior-iaquierda)
                    kwargs_opcion_coordenadas_config_paso_1 = ({"opcion_coordenadas_config_paso_1": "COLOCACION_INICIAL"}
                                                                if proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"] and msg_1.valor_boton_pulsado
                                                                else
                                                                {"opcion_coordenadas_config_paso_1": "COLOCACION_HIBRIDA"}
                                                                if proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"] and not msg_1.valor_boton_pulsado
                                                                else
                                                                None
                                                                )
                    
                    if kwargs_opcion_coordenadas_config_paso_1 is not None:
                        kwargs_rutina_interaccion_pptx.update(kwargs_opcion_coordenadas_config_paso_1)




                    #se inician los filedialogs y se construye el diccionario kwargs_rutina_interaccion_pptx que se usa como kwargs
                    #en la llamada a la rutina def_config_sistema_sqlite (mod back_end) mas abajo
                    #se informa mensaje_cancelacion_proceso con el mensaje que avisa en la GUI pq el proceso se cancela
                    #(se cancela tan solo si el usuario en los filedialogs que se generan los cierra sin seleccionar nada)
                    #
                    #IMPORTANTE los nombres de variables que se usan para hacer la llamada al filedialog (sea filename o directory)
                    #han de estar escritos exactamente como los kwargs que se usan en la rutina def_interaccion_pptx (mod back_end)
                    #
                    #los file dialog se pueden iniciar:
                    # --> CONFIG_PASO_1     al puslar los 2 1eros botones del message box (valores true y false)
                    # --> CONFIG_PASO_2     al pulsar el 1er boton (valor true)
                    # --> EJECUCION         al pulsar el 1er boton (valor true)
                    proceso_cancelado = True
                    mensaje_cancelacion_proceso = None
                    if ((proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]  and msg_1.valor_boton_pulsado)
                        or (proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]  and not msg_1.valor_boton_pulsado)
                        or (proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_2"]["OPCION"]  and msg_1.valor_boton_pulsado)
                        or (proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["EJECUCION"]["OPCION"]  and msg_1.valor_boton_pulsado)):

                        
                        if proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]:

                            mensaje_filedialog = "Selecciona en que directorio quieres que se cree el pptx tras la configuración del paso 1:"
                            
                            directorio_pptx_destino = fd.askdirectory(parent = self.master.widget_objeto, title = mensaje_filedialog)
                            directorio_log_errores = directorio_pptx_destino

                            if directorio_pptx_destino:
                                proceso_cancelado = False
                                kwargs_rutina_interaccion_pptx.update({"directorio_pptx_destino": directorio_pptx_destino})
                            else:
                                mensaje_cancelacion_proceso = "No se puede iniciar el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."




                        elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_2"]["OPCION"]:

                            mensaje_filedialog = "Selecciona el fichero pptx para poder finalizar la configuración del Paso 2:"

                            path_pptx_config = fd.askopenfilename(parent = self.master.widget_objeto, title ="Selecciona el fichero pptx para po0der finalizar la configuración del Paso 2.", filetypes = [("Fichero pptx", "*.pptx")])
                            directorio_log_errores = os.path.normpath(os.path.dirname(path_pptx_config))

                            if path_pptx_config:
                                proceso_cancelado = False
                                kwargs_rutina_interaccion_pptx.update({"path_pptx_config": path_pptx_config})  
                            else:
                                mensaje_cancelacion_proceso = "No se puede iniciar el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."




                        elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["EJECUCION"]["OPCION"]:

                            directorio_pptx_destino = fd.askdirectory(parent = self.master.widget_objeto, title = "Selecciona en que directorio quieres que se cree el pptx definitivo:")
                            directorio_log_errores = directorio_pptx_destino

                            if directorio_pptx_destino:
                                proceso_cancelado = False
                                kwargs_rutina_interaccion_pptx.update({"directorio_pptx_destino": directorio_pptx_destino})  
                            else:
                                mensaje_cancelacion_proceso = "No se puede iniciar el proceso porque has cerrado la ventana de dialogo, que se genero, sin seleccionar nada."





                    if proceso_cancelado:

                        #se ejecuta el proceso solo en caso de que kwargs_rutina_sistema_sqlite no este vacio
                        #si lo esta es pq el usuario cuando se han generado los filedialog los ha cerrado en vez 
                        #de seleccionar un directorio o un fichero segun el tipo de filedialog
                        if mensaje_cancelacion_proceso is not None:
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_cancelacion_proceso}))

                    else:
                        #se generan los kwargs para ejecutar la rutina def_interaccion_pptx (mod back end)
                        #segun que se ejecute el args de proceso o de generacion del log de errores
                        kwargs_rutina_interaccion_pptx_proceso = kwargs_rutina_interaccion_pptx.copy()

                        kwargs_rutina_interaccion_pptx_proceso.update({"id_pptx_selecc": self.id_pptx_ventana_inicio_selecc})
                        


                        #se agrega el kwargs kwargs_rutina_interaccion_pptx_proceso calculado mas arriba
                        if kwargs_opcion_coordenadas_config_paso_1 is not None:
                            kwargs_rutina_interaccion_pptx_proceso.update(kwargs_opcion_coordenadas_config_paso_1)
                        


                        #se ejecuta el proceso seleccionado donde previamente se almacena las configuraciones del sistema sqlite 
                        #en la variable global global_dicc_datos_id_pptx
                        self.master.widget_objeto.config(cursor = "wait")

                        ruta_fichero_interaccion_pptx = mod_back_end.def_interaccion_pptx(proceso_selecc, **kwargs_rutina_interaccion_pptx_proceso)
                        
                        self.master.widget_objeto.config(cursor = "")



                        ############################################
                        # CASO 1 - hay errores / warnings
                        ###########################################
                        #se genera messagebox avisando de la creacion o no de un log de errores
                        #se usa la variable global global_lista_dicc_errores y global_lista_dicc_warning (modulo backend)
                        #para localizar estos posibles errores / warnings
                        if len(mod_back_end.global_lista_dicc_errores) != 0 or len(mod_back_end.global_lista_dicc_warning) != 0:


                            if proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_2"]["OPCION"]:


                                mensaje1 = f"Proceso ejecutado: {proceso_selecc}.\n\n"
                                mensaje2 = "Cuando cierres este mensaje, se abrira el fichero de log de errores / warning."
                                
                                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel_warning", "mensaje": mensaje1 + mensaje2}))

                                if msg.valor_boton_pulsado:

                                    msg = mod_back_end.def_varios("GENERAR_LOG_WARNING_ERRORES_PROCESOS_APP"
                                                                    , directorio_log_errores_warning = directorio_log_errores
                                                                    , opcion_warning_errores = "WARNING_ERRORES")



                            elif (proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]
                                    or proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["EJECUCION"]["OPCION"]):


                                mensaje1 = "Proceso finalizado.\n\nSe han localizado errores y/o warnings.\n\n"
                                mensaje2 = "Cuando cierres este mensaje, se abriran tanto el fichero powerpoint resultante del proceso como el fichero de log de errores / warning."

                                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel_warning", "mensaje": mensaje1 + mensaje2}))

                                if msg.valor_boton_pulsado:


                                    if ruta_fichero_interaccion_pptx is not None:
                                        os.startfile(ruta_fichero_interaccion_pptx)

                                    mod_back_end.def_varios("GENERAR_LOG_WARNING_ERRORES_PROCESOS_APP"
                                                            , directorio_log_errores_warning = directorio_log_errores
                                                            , opcion_warning_errores = "WARNING_ERRORES")
                                    



                        ############################################
                        # CASO 2 - NO hay errores / warnings
                        ###########################################
                        elif len(mod_back_end.global_lista_dicc_errores) == 0 and len(mod_back_end.global_lista_dicc_warning) == 0:


                            if (proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_1"]["OPCION"]
                                    or proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["EJECUCION"]["OPCION"]):

                                    mensaje = f"Proceso ejecutado: {proceso_selecc}.\n\nSe abrira el powerpoint nada mas cierres este mensaje."

                                    msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

                                    if msg.valor_boton_pulsado:

                                        if ruta_fichero_interaccion_pptx is not None:
                                            os.startfile(ruta_fichero_interaccion_pptx)


                            elif proceso_selecc == mod_back_end.dicc_gui_combobox_procesos["COMBOBOX_PPTX"]["CONFIG_PASO_2"]["OPCION"]:

                                    mensaje = f"Proceso ejecutado: {proceso_selecc}."
                                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))




    def def_gui_ventana_inicio_treeview_id_pptx_click_item(self):
        #rutina que permite informar el atributo self.id_pptx_ventana_inicio_selecc de la presente clase
        #cuando se hace click en un item del treeview id xls (es treeview de seleccion simple)
        #se usa el metodo datos_items relacionado al objeto treeview_id_pptx (creado con el modulo utils)
        #es diccionario donde la key lista_datos_items_seleccionados es a su vez lista de lista con una sola
        #sublista de ahi el 1er [0]
        #a modo recordatorio las columnas que forman el treeview de id xls son el id xls y su descripcion
        #de ahi el 2ndo [0]

        self.id_pptx_ventana_inicio_selecc = (self.treeview_id_pptx.datos_items["lista_datos_items_seleccionados"][0][0]
                                                if self.treeview_id_pptx.datos_items["lista_datos_items_seleccionados"] is not None
                                                else None)

        

    def def_gui_ventana_inicio_combobox_accion_pptx(self):
        #rutina para rellenar el scrolledtext con la descripcion del proceso pptx

        opcion_pptx_selecc = self.combobox_accion_pptx.variable_enlace.get()
        

        if opcion_pptx_selecc is not None:
            desc_proceso_pptx = mod_back_end.def_varios_gui_ventana_inicio("DESC_ASOCIADO_ACCION_PPTX", opcion_combobox_sistema = opcion_pptx_selecc)

            self.scrolledtext_desc_accion_pptx.config_atributos(**{"bloquear": False})

            self.scrolledtext_desc_accion_pptx.modificaciones("borrar_contenido_y_tags")
            self.scrolledtext_desc_accion_pptx.modificaciones("agregar_solo_contenido_desde_string"
                                                                , string_texto_informar = desc_proceso_pptx
                                                                , height_scrolledtext = self.scrolledtext_desc_accion_pptx_height)
            
            self.scrolledtext_desc_accion_pptx.config_atributos(**{"bloquear": True})




###################################################################################################################################################
###################################################################################################################################################
# clase gui_config_id_pptx
###################################################################################################################################################
###################################################################################################################################################

class gui_config_id_pptx:

    def __init__(self, master, descargar_screenshots_muestra_en_png = False, entorno_clase_gui_ventana_inicio = None, **kwargs_config_gui):


        #se inicializan atributosde la presente clase
        self.master = master
        self.clase_gui_actual_nombre = self.__class__.__name__
        self.kwargs_config_gui = kwargs_config_gui


        self.id_pptx_selecc = None 
        self.id_xls_selecc = None
        self.id_xls_selecc_antes_click_item = None

        self.widget_seleccion_id_pptx_lista_opciones = None
        self.widget_hojas_xls_lista_opciones = None
        self.widget_slide_pptx_lista_opciones = None
        self.rango_celdas_selecc = None

        self.descargar_screenshots_muestra_en_png = descargar_screenshots_muestra_en_png

        self.entorno_clase_gui_ventana_inicio = entorno_clase_gui_ventana_inicio

        self.lista_procesos_no_requieren_seleccion_id_pptx = ["GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE", "ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO"]



        #se calcula la variable global del modulo back_end que permite preparar los datos del sistema sqlite y adaptarlos a como se presentan en la GUI
        #y realizar dichas actualizaciones de forma flexible en la GUI sin ajustes inecesarios
        mod_back_end.def_varios("DICC_DATOS_ID_PPTX")




        #se extraen los kwargs para crear el frame con scrollbar integrado al root
        self.kwargs_gui_config_id_pptx_frame_integrado_root = self.kwargs_config_gui[self.clase_gui_actual_nombre]["frame_integrado_root"]["frame_inicio"]["kwargs_config"]
        self.frame_inicio = mod_utils.frame_con_scrollbar(self.master.widget_objeto, **self.kwargs_gui_config_id_pptx_frame_integrado_root)


        #se extraen los kwargs de los widgets a integrar en el frame con scrollbar del bloque anterior
        self.kwargs_gui_config_id_pptx = {key: dicc for key, dicc in kwargs_config_gui[self.clase_gui_actual_nombre]["frame_integrado_root"].items() if key != "frame_inicio"}


        #se insertan los widgets dentro del frame_inicio y se almacenan en el diccionario dicc_gui_frame_widgets_objetos
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_gui_config_id_pptx_frame_widgets_objetos = {}
        for frame_contenedor_nivel_1 in self.kwargs_gui_config_id_pptx.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_config_id_pptx_frame_nivel_1 = [dicc["frame"] for frame, dicc in self.kwargs_gui_config_id_pptx.items() if frame == frame_contenedor_nivel_1][0]

            self.frame_nivel_1 = mod_utils.gui_tkinter_widgets(self.frame_inicio.widget_objeto, tipo_widget_param = "frame", **kwargs_gui_config_id_pptx_frame_nivel_1)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_config_id_pptx_frame_widgets_nivel_1 = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_config_id_pptx[frame_contenedor_nivel_1].items() if widget != "frame"}

            for frame_nivel_1_widget, frame_nivel_1_kwargs_widget in kwargs_gui_config_id_pptx_frame_widgets_nivel_1.items():

                try:
                    tipo_widget_nivel_1 = frame_nivel_1_kwargs_widget["tipo_widget"].lower().replace(" ", "").strip()
                    kwargs_config_nivel_1 = frame_nivel_1_kwargs_widget["kwargs_config"]


                    #se crean los widgets (salvo WIDGET_42 que contiene el frame scrollable para los rangos de celdas)
                    widget_objeto_nivel_1 = (mod_utils.gui_tkinter_widgets(self.frame_nivel_1.widget_objeto, tipo_widget_param = tipo_widget_nivel_1, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_1)
                                            if tipo_widget_nivel_1 in ["label", "combobox", "entry", "button", "listbox"]
                                            else
                                            mod_utils.scrolledtext_propio(self.frame_nivel_1.widget_objeto, **kwargs_config_nivel_1)
                                            if tipo_widget_nivel_1 == "scrolledtext_propio"
                                            else
                                            mod_utils.treeview_propio(self.frame_nivel_1.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_1)
                                            if tipo_widget_nivel_1 == "treeview_propio"
                                            else
                                            mod_utils.entry_propio(self.frame_nivel_1.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_1)
                                            if tipo_widget_nivel_1 == "entry_propio"
                                            else
                                            mod_utils.frame_con_scrollbar(self.frame_nivel_1.widget_objeto, **kwargs_config_nivel_1)
                                            if tipo_widget_nivel_1 == "frame_con_scrollbar"
                                            else None)



                    #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                    self.dicc_gui_config_id_pptx_frame_widgets_objetos.update({frame_nivel_1_widget:
                                                                                            {"frame_nivel_1": frame_contenedor_nivel_1
                                                                                            , "frame_nivel_2": None
                                                                                            , "tipo_widget": tipo_widget_nivel_1
                                                                                            , "widget_objeto": widget_objeto_nivel_1
                                                                                            , "widget_variable_enlace": widget_objeto_nivel_1.variable_enlace
                                                                                            , "kwargs_config": kwargs_config_nivel_1
                                                                                            }
                                                                            })
                    
                    
                except KeyError as _:
                    #si dar KeyError es pq es el frame con scrollbar con los rangos de celda (el kwargs tiene un formato distinto)
                    #en este caso se incorpora un frane (nivel 2) dentro de un frame (nivel 1)

                    #se recuperan los kwargs para crear el frame scrollable
                    kwargs_gui_config_id_pptx_frame_nivel_2 = frame_nivel_1_kwargs_widget["frame"]["kwargs_config"]

                    self.frame_nivel_2 = mod_utils.frame_con_scrollbar(self.frame_nivel_1.widget_objeto, **kwargs_gui_config_id_pptx_frame_nivel_2)


                    #se itera por los distintos widgets que componen este frame scrollable para poder crearlos
                    kwargs_gui_config_id_pptx_frame_nivel_2_widgets = {key: valor for key, valor in frame_nivel_1_kwargs_widget.items() if key != "frame"}

                    for frame_nivel_2_widget, frame_nivel_2_kwargs_widget in kwargs_gui_config_id_pptx_frame_nivel_2_widgets.items():

                        tipo_widget_nivel_2 = frame_nivel_2_kwargs_widget["tipo_widget"].lower().replace(" ", "").strip()
                        kwargs_config_nivel_2 = frame_nivel_2_kwargs_widget["kwargs_config"]

                        widget_objeto_nivel_2 = (mod_utils.gui_tkinter_widgets(self.frame_nivel_2.widget_objeto, tipo_widget_param = tipo_widget_nivel_2, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_2)
                                                if tipo_widget_nivel_2 in ["label", "combobox", "entry", "button", "listbox"]
                                                else
                                                mod_utils.scrolledtext_propio(self.frame_nivel_2.widget_objeto, **kwargs_config_nivel_2)
                                                if tipo_widget_nivel_2 == "scrolledtext_propio"
                                                else
                                                mod_utils.treeview_propio(self.frame_nivel_2.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_2)
                                                if tipo_widget_nivel_2 == "treeview_propio"
                                                else
                                                mod_utils.entry_propio(self.frame_nivel_2.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config_nivel_2)
                                                if tipo_widget_nivel_2 == "entry_propio"
                                                else
                                                mod_utils.frame_con_scrollbar(self.frame_nivel_2.widget_objeto, **kwargs_config_nivel_2)
                                                if tipo_widget_nivel_2 == "frame_con_scrollbar"
                                                else None)



                        #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_nivel_2 junto con su stringvar (si lo tiene)
                        self.dicc_gui_config_id_pptx_frame_widgets_objetos.update({frame_nivel_2_widget:
                                                                                                        {"frame_nivel_1": frame_contenedor_nivel_1
                                                                                                        , "frame_nivel_2": "rangos_celdas_xls" #se informa manualmente (frame de nivel 2 solo hay este)
                                                                                                        , "tipo_widget": tipo_widget_nivel_2
                                                                                                        , "widget_objeto": widget_objeto_nivel_2
                                                                                                        , "widget_variable_enlace": widget_objeto_nivel_2.variable_enlace
                                                                                                        , "kwargs_config": kwargs_config_nivel_2
                                                                                                        }
                                                                                    })
                        
                    pass  


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.widget_seleccion_id_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_20"]["widget_objeto"]
        self.widget_id_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_26"]["widget_objeto"]
        self.widget_id_pptx_path = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_31"]["widget_objeto"]
        self.widget_id_pptx_desc = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_34"]["widget_objeto"]
        self.widget_tiempo_apertura_max_id_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_36"]["widget_objeto"]
        self.widget_id_pptx_num_slides = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_38"]["widget_objeto"]

        self.widget_treeview_id_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_40"]["widget_objeto"]
        self.widget_id_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_42"]["widget_objeto"]
        self.widget_id_xls_path = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_47"]["widget_objeto"]
        self.widget_id_xls_desc = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_50"]["widget_objeto"]
        self.widget_tiempo_apertura_max_id_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_52"]["widget_objeto"]
        self.widget_actualizar_vinculos_otros_excel = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_54"]["widget_objeto"]


        self.widget_treeview_rangos_celdas = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_57"]["widget_objeto"]
        self.widget_hojas_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_59"]["widget_objeto"]
        self.widget_rangos_celdas = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_61"]["widget_objeto"]
        self.widget_slide_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_63"]["widget_objeto"]
        self.widget_nombre_pantallazo = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_65"]["widget_objeto"]
        self.widget_coordenadas_pantallazo = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_67"]["widget_objeto"]

        #se recuperan los height de los scrolledtext desde los parametros kwargs de la presente clase
        self.widget_id_pptx_desc_height = self.kwargs_gui_config_id_pptx["frame_pptx_destino"]["WIDGET_34"]["kwargs_config"].get("height", 1)
        self.widget_id_xls_desc_height = self.kwargs_gui_config_id_pptx["frame_xls_origen"]["WIDGET_50"]["kwargs_config"].get("height", 1)


        #se recuperan las variables de enlace (stringvar)
        self.strvar_widget_seleccion_id_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_20"]["widget_variable_enlace"]
        self.strvar_widget_tiempo_apertura_max_id_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_36"]["widget_variable_enlace"]
        self.strvar_widget_tiempo_apertura_max_id_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_52"]["widget_variable_enlace"]
        self.strvar_widget_actualizar_vinculos_otros_excel = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_54"]["widget_variable_enlace"]
        self.strvar_widget_hojas_xls = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_59"]["widget_variable_enlace"]
        self.strvar_widget_rangos_celdas = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_61"]["widget_variable_enlace"]
        self.strvar_widget_slide_pptx = self.dicc_gui_config_id_pptx_frame_widgets_objetos["WIDGET_63"]["widget_variable_enlace"]




        #se actualiza la lista de opciones id pptx en el combobox de seleccion
        self.widget_seleccion_id_pptx_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_ID_PPTX")

        lista_dicc_widgets = [      
                                {"tipo_widget": "combobox"
                                    , "widget_objeto": self.widget_seleccion_id_pptx
                                    , "variable_enlace": self.strvar_widget_seleccion_id_pptx
                                    , "height": None
                                    , "bloquear": False
                                    , "combobox_lista_opciones": self.widget_seleccion_id_pptx_lista_opciones
                                    , "combobox_opciones_editables": False
                                    , "treeview_seleccionar_item": None
                                    , "valor_informar": None
                                    }
                            ]
        
        self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)



        #se vacian y bloquean todos los widgets de inicio
        self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS")




    def def_gui_config_id_pptx_threads(self, proceso_selecc):
        #rutina para ejecutar todos los procesos del app para la clase gui_config_id_pptx
        #se hace por thread para poder "jugar" con la variable global global_proceso_en_ejecucion
        #y asi evitar que mientras se ejecute el proceso actual se pueda ejecutarlo de nuevo al mismo tiempo
        #si se intenta ejecutar mientras el mismo proceso esta en curso sale un warning
        #(cuando acabe la ejecucion del proceso actual la variable global global_proceso_en_ejecucion se renicia a NO)

        if mod_back_end.global_proceso_en_ejecucion == "SI":
            mensaje = "Espera a que acabe el proceso actualmente en ejecución."
            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

        else:
            Thread(target = self.def_gui_config_id_pptx_procesos, args = (proceso_selecc,)).start()




    def def_gui_config_id_pptx_procesos(self, proceso_selecc):
        #rutina que ejecuta las acciones de cada uno de los botones de la clase gui_config_id_pptx
        #todos los procesos (salvo el de crear un nuevo id pptx) requieren tener seleccionado un id pptx en el combobox de seleccion
        #y haber pulsado el boton VER

        id_pptx_combobox = self.strvar_widget_seleccion_id_pptx.get()
        proceso_ejecutable = False if proceso_selecc not in self.lista_procesos_no_requieren_seleccion_id_pptx and len(id_pptx_combobox) == 0 else True


        if not proceso_ejecutable:
            mensaje = "No has seleccionado ningún 'id pptx'"
            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))
            return


        #############################################################################################################
        # SELECCION_ID_PPTX
        #############################################################################################################
        elif proceso_selecc == "SELECCION_ID_PPTX" and proceso_ejecutable:
            #actualiza los widgets segun el valor del atributo propio self.id_pptx_selecc y tras pulsar el boton de seleccion
            # --> al acceder a la ventana que genera la clase gui_config_id_pptx si se ha clicado sobre un id pptx
            #     previamente en el treeview de la clase gui_ventana_inicio
            # --> al seleccionar un id pptx en el combobox de seleccion y tras pulsar el boton de actualizacion


            # se realiza la actualizacion en pantalla tan solo si se cambia el id pptx en el combobox
            if self.id_pptx_selecc == id_pptx_combobox:
                mensaje = "No has seleccionado otro id pptx."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))

            else:

                #se guardan los datos en pantalla de los datos generales del id pptx seleccionado y del id xls seleccionado (si lo esta)
                #en la variable global global_dicc_datos_id_pptx (sin este ajuste si el usuario ha modificado datos en pantalla
                #estos no se guardan en el sistema sqlite)
                if self.id_pptx_selecc is not None:

                    mod_back_end.def_varios_gui_config_id_pptx("UPDATE_EN_MEMORIA_DATOS_EN_PANTALLA_ID_PPTX_Y_ID_XLS"
                                                                , id_pptx = self.id_pptx_selecc
                                                                , id_pptx_desc = self.widget_id_pptx_desc.texto_informado("todo")
                                                                , id_pptx_path = self.widget_id_pptx_path.texto_informado("todo")
                                                                , id_pptx_tiempo_espera_max_apertura = self.strvar_widget_tiempo_apertura_max_id_pptx.get()
                                                                , id_pptx_numero_total_slides = self.widget_id_pptx_num_slides.texto_informado("todo")

                                                                , id_xls = self.id_xls_selecc
                                                                , id_xls_desc = self.widget_id_xls_desc.texto_informado("todo")
                                                                , id_xls_path = self.widget_id_xls_path.texto_informado("todo")
                                                                , id_xls_tiempo_espera_max_apertura = self.strvar_widget_tiempo_apertura_max_id_xls.get()
                                                                , id_xls_actualizar_vinculos_otros_xls = self.strvar_widget_actualizar_vinculos_otros_excel.get()
                                                                )


                #se recuperan los datos del id pptx seleccionado en el combobox
                (selecc_id_pptx
                , selecc_id_pptx_desc
                , selecc_id_pptx_path
                , selecc_id_pptx_tiempo_espera_max_apertura
                , selecc_numero_total_slides
                , selecc_df_widget_treeview_id_xls
                ) = mod_back_end.def_varios_gui_config_id_pptx("SELECCION_ID_PPTX", id_pptx = id_pptx_combobox)


                #se borra todo el contenido de la GUI y se bloquean todos los widgets
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS")


                #se informan y desbloquean los widgets necesarios mediante el uso de un diccionario temporal
                lista_dicc_widgets = [      
                                        {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_pptx
                                            , "variable_enlace": None
                                            , "height": 1
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_id_pptx
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_pptx_path
                                            , "variable_enlace": None
                                            , "height": 1
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_id_pptx_path
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_pptx_desc
                                            , "variable_enlace": None
                                            , "height": self.widget_id_pptx_desc_height
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_id_pptx_desc
                                            }

                                        , {"tipo_widget": "entry_propio"
                                            , "widget_objeto": self.widget_tiempo_apertura_max_id_pptx
                                            , "variable_enlace": self.strvar_widget_tiempo_apertura_max_id_pptx
                                            , "height": 1
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_id_pptx_tiempo_espera_max_apertura
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_pptx_num_slides
                                            , "variable_enlace": None
                                            , "height": 1
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_numero_total_slides
                                            }

                                        , {"tipo_widget": "treeview"
                                            , "widget_objeto": self.widget_treeview_id_xls
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": selecc_df_widget_treeview_id_xls
                                            }
                                    ]
                
                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)


                #se establecen los valores de los atributoa propios
                self.id_pptx_selecc = id_pptx_combobox
                self.id_xls_selecc = None
                self.id_xls_selecc_antes_click_item = None
                self.rango_celdas_selecc = None

                self.widget_hojas_xls_lista_opciones = None
                self.widget_slide_pptx_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_SLIDES_PPTX", id_pptx = self.id_pptx_selecc)



        #############################################################################################################
        # GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE
        #############################################################################################################
        elif proceso_selecc == "GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE" and proceso_ejecutable:
            #guarda las configuraciones en el sistema sqlite basandose en los datos de la variable global global_dicc_datos_id_pptx
            #genera warning en la GUI (no bloquante) si:
            # --> si hay id pptx pero sin ningun id xls
            # --> si hay id xls sin ningun ranfode celdas
            #actualiza la lista de opciones del combobox de seleccion de id pptx en la GUI
            #realiza asismismo la descar en el path local de losscreenshots png para usarlos como muestra en la GUI

            (tipo_messagebox
            , mensaje_gui) = mod_back_end.def_varios("MESSAGEBOX_PROCESOS_APP"
                                                    , opcion_proceso ="GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE")

            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": tipo_messagebox, "mensaje": mensaje_gui}))
            

            if msg.valor_boton_pulsado in [True, False]:

                self.master.widget_objeto.config(cursor = "wait")


                #se guardan los datos en pantalla de los datos generales del id pptx seleccionado y del id xls seleccionado (si lo esta)
                #en la variable global global_dicc_datos_id_pptx (sin este ajuste si el usuario ha modificado datos en pantalla
                #estos no se guardan en el sistema sqlite)
                mod_back_end.def_varios_gui_config_id_pptx("UPDATE_EN_MEMORIA_DATOS_EN_PANTALLA_ID_PPTX_Y_ID_XLS"
                                                            , id_pptx = self.id_pptx_selecc
                                                            , id_pptx_desc = self.widget_id_pptx_desc.texto_informado("todo")
                                                            , id_pptx_path = self.widget_id_pptx_path.texto_informado("todo")
                                                            , id_pptx_tiempo_espera_max_apertura = self.strvar_widget_tiempo_apertura_max_id_pptx.get()
                                                            , id_pptx_numero_total_slides = self.widget_id_pptx_num_slides.texto_informado("todo")

                                                            , id_xls = self.id_xls_selecc
                                                            , id_xls_desc = self.widget_id_xls_desc.texto_informado("todo")
                                                            , id_xls_path = self.widget_id_xls_path.texto_informado("todo")
                                                            , id_xls_tiempo_espera_max_apertura = self.strvar_widget_tiempo_apertura_max_id_xls.get()
                                                            , id_xls_actualizar_vinculos_otros_xls = self.strvar_widget_actualizar_vinculos_otros_excel.get()
                                                            )



                #se reinician las variables globales global_lista_dicc_errores y global_lista_dicc_warning
                mod_back_end.global_lista_dicc_errores = []
                mod_back_end.global_lista_dicc_warning = []


                #se ejecuta el guardado
                if msg.valor_boton_pulsado == False:
                    mod_back_end.def_varios_gui_config_id_pptx("UPDATE_NUMERO_SLIDES_PPTX_Y_LISTA_HOJAS_XLS")

                mod_back_end.def_varios_gui_config_id_pptx("GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE")



                #se borra todo el contenido de la GUI y se bloquean todos los widgets
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS")


                #se informan y desbloquean los widgets necesarios mediante el uso de un diccionario temporal
                self.widget_seleccion_id_pptx_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_ID_PPTX")

                lista_dicc_widgets = [      
                                        {"tipo_widget": "combobox"
                                            , "widget_objeto": self.widget_seleccion_id_pptx
                                            , "variable_enlace": self.strvar_widget_seleccion_id_pptx
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": self.widget_seleccion_id_pptx_lista_opciones
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": None
                                            }
                                    ]
                
                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)


                #se descargan en la ruta local los png de los screenshots para usarlos en GUI (muestra screenshots)
                mod_back_end.def_config_sistema_sqlite("DESCARGA_SCREENSHOTS_PNG_PARA_MUESTRA_GUI_CONFIG_ID_PPTX")


                #se actualiza el treeview en la ventana de inicio
                self.entorno_clase_gui_ventana_inicio.treeview_id_pptx.acciones("actualizar_desde_df", df_datos = mod_back_end.global_df_treeview_id_pptx)


                self.master.widget_objeto.config(cursor = "")


                #se reinicializan los valores de los atributoa propios
                self.id_pptx_selecc = None
                self.id_xls_selecc = None
                self.id_xls_selecc_antes_click_item = None
                self.rango_celdas_selecc = None

                self.widget_hojas_xls_lista_opciones = None
                self.widget_slide_pptx_lista_opciones = None


                #se genera el log de errores / warning
                if len(mod_back_end.global_lista_dicc_errores) != 0 or len(mod_back_end.global_lista_dicc_warning) != 0:

                    mod_back_end.def_varios("GENERAR_LOG_WARNING_ERRORES_PROCESOS_APP"
                                            , directorio_log_errores_warning = mod_back_end.global_ruta_local_config_sistema_sqlite
                                            , opcion_warning_errores = "WARNING_ERRORES")

                    mensaje_1 = "Configuración actualizada en el sistema sqlite.\n\n"
                    mensaje_2 = f"No obstante, se han localizado warnings, se ha generado un log de errores en la ruta siguiente:\n\n {mod_back_end.global_ruta_local_config_sistema_sqlite}"
                    mensaje = mensaje_1 + mensaje_2

                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showwarning", "mensaje": mensaje}))

                else:
                    mensaje = "Configuración actualizada en el sistema sqlite.\n\nNo se han localizado warnings."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showwarning", "mensaje": mensaje}))




        #############################################################################################################
        # ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO
        #############################################################################################################
        elif proceso_selecc == "ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO" and proceso_ejecutable:
            #abre un filedialog para seleccionar una ruta pptx y abre una nueva ventana donde informar los datos gererales asociados al nuevo id pptx
            #se realiza previamente un check al seleccionar una ruta pptx por si la ruta ya esta configurada salga un warning avisando de ello e informando en que id pptx se configuro


            mensaje = "Se abrira una ventana de dialogo para que selecciones la ubicación del pptx de destino que quieres usar y posteriormente se abrira una ventana donde tendras que configurar los datos generales del id pptx a crear."
            msg_1 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

            if msg_1.valor_boton_pulsado:

                #se abre el file dialogparaseleccionar laubicacion del pptx
                id_pptx_nuevo_path = fd.askopenfilename(parent = self.master.widget_objeto, title = "SELECCIONA UNA NUEVA UBICACIÓN DE UN PPTX DE DESTINO:", filetypes = [("Ficheros Powerpoint", "*.ppt;*.pptx")])

                if id_pptx_nuevo_path:

                    #se asigna un nuevo id pptx, se recupera el numero de slides que contiene y se crea el mensaje para el warning
                    #en caso de que la ubicacion este ya asignada a otros id pptx
                    mensaje_warning = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_PATH_ID_PPTX_YA_ASIGNADO", id_pptx_path = id_pptx_nuevo_path)
                    id_pptx_nuevo, id_pptx_nuevo_num_slides = mod_back_end.def_varios_gui_config_id_pptx("ASIGNAR_ID_PPTX_NUEVO", id_pptx_path = id_pptx_nuevo_path)


                    abrir_ventana_cofig_nuevo_id_pptx = False
                    if mensaje_warning is not None:
                        msg_2 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje_warning}))

                        if msg_2.valor_boton_pulsado:
                            abrir_ventana_cofig_nuevo_id_pptx = True

                    else:
                        abrir_ventana_cofig_nuevo_id_pptx = True

                        
                    #se abre una nueva ventana para configurar el id pptx creado
                    if abrir_ventana_cofig_nuevo_id_pptx:
                        kwargs_gui_config_id_pptx_nuevo_id_pptx_dicc_config_root = self.kwargs_config_gui["gui_config_id_pptx_nuevo_id_pptx"]["dicc_config_root"]

                        self.toplevel_gui_config_id_pptx_nuevo_id_pptx = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_config_id_pptx_nuevo_id_pptx_dicc_config_root)
                        self.toplevel_gui_config_id_pptx_nuevo_id_pptx.config_atributos(**kwargs_gui_config_id_pptx_nuevo_id_pptx_dicc_config_root)

                        gui_config_id_pptx_nuevo_id_pptx(self.toplevel_gui_config_id_pptx_nuevo_id_pptx
                                                        , id_pptx_nuevo = id_pptx_nuevo
                                                        , id_pptx_nuevo_path = id_pptx_nuevo_path
                                                        , id_pptx_nuevo_num_slides = id_pptx_nuevo_num_slides
                                                        , entorno_clase_gui_config_id_pptx = self
                                                        , **self.kwargs_config_gui) 




        #############################################################################################################
        # ELIMINAR_ID_PPTX
        #############################################################################################################
        elif proceso_selecc == "ELIMINAR_ID_PPTX" and proceso_ejecutable:
            #elimina en memoria el id pptx y vacia los widgets
            #reactualiza la lista de opciones del combobox de seleccion de id pptx quitando el id pptx)

            mensaje1 = f"Se eliminara el id pptx '{self.id_pptx_selecc}' y sus configuraciones.\n\nLa eliminación tan solo se hara en la memoria del pc.\n\n"
            mensaje2 = "Tendrás que pulsar el botón 'GUARDAR' para que la eliminación también se realice en el sistema sqlite.\n\nDeseas continuar?"
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje1 + mensaje2}))

            if msg.valor_boton_pulsado:

                #se elimina el id pptx de la variable global global_dicc_datos_id_pptx y se recupera la lista de opciones para el combobox de seleccion id pptx
                #(el proceso de ELIMINAR_ID_PPTX funciona a su vez como funcion)
                lista_opciones_seleccion_id_pptx = mod_back_end.def_varios_gui_config_id_pptx("ELIMINAR_ID_PPTX", id_pptx = self.id_pptx_selecc)


                #se informan los cambios en la GUI
                lista_dicc_widgets = [      
                                        {"tipo_widget": "combobox"
                                            , "widget_objeto": self.widget_seleccion_id_pptx
                                            , "variable_enlace": self.strvar_widget_seleccion_id_pptx
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": lista_opciones_seleccion_id_pptx
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": None
                                            }
                                    ]
                
                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)


                #se borra todo el contenido de la GUI y se bloquean todos los widgets
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS")


                #se reinicializan los valores de los atributoa propios
                self.id_pptx_selecc = None
                self.id_xls_selecc = None
                self.id_xls_selecc_antes_click_item = None
                self.rango_celdas_selecc = None

                self.widget_seleccion_id_pptx_lista_opciones = lista_opciones_seleccion_id_pptx
                self.widget_hojas_xls_lista_opciones = None
                self.widget_slide_pptx_lista_opciones = None


                mensaje = "id pptx eliminado en memoria.\n\nLa eliminación surtira efecto en el sistema sqlite cuando pulses el botón GUARDAR."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))



        #############################################################################################################
        # UPDATE_PATH_ID_PPTX
        #############################################################################################################
        elif proceso_selecc == "UPDATE_PATH_ID_PPTX" and proceso_ejecutable:
            #abre un filedialog para seleccionar una ruta xls y se realiza un check por si la ruta ya esta configurada en otro id pptx
            #para que salga un warning no bloquante avisando de ello

            mensaje1 = "Se abrira una ventana de dialogo para que puedes configurar una nueva ubicación de fichero pptx.\n\n"
            mensaje2 = "Se realizara un check que generara un warning (no bloquante) en caso de que la ruta que intentas configurar ya esta asignada a otro id pptx."
            msg_1 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje1 + mensaje2}))

            if msg_1.valor_boton_pulsado:

                ruta_fichero_pptx = fd.askopenfilename(parent = self.master.widget_objeto, title = "SELECCIONA UNA NUEVA UBICACIÓN DE UN PPTX DE DESTINO:", filetypes = [("Ficheros Powerpoint", "*.ppt;*.pptx")])

                if ruta_fichero_pptx:

                    mensaje_warning_1 = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_PATH_ID_PPTX_YA_ASIGNADO"
                                                                                , id_pptx_path = ruta_fichero_pptx
                                                                                , id_pptx_excluir_check_path = self.id_pptx_selecc)



                    ejecutar_update_path_id_pptx = False
                    if mensaje_warning_1 is not None:
                        msg_2 = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje_warning_1}))

                        if msg_2.valor_boton_pulsado:
                            ejecutar_update_path_id_pptx = True

                    else:
                        ejecutar_update_path_id_pptx = True


                    #se ejecuta el update tanto memoria como en la GUI
                    if ejecutar_update_path_id_pptx:

                        #se actualiza en la variable global global_dicc_datos_id_pptx el nuevo path y su respectivo numero de slides
                        (mensaje_warning_2
                        , id_pptx_path_ajust
                        , id_pptx_numero_total_slides
                        ) = mod_back_end.def_varios_gui_config_id_pptx("UPDATE_PATH_ID_PPTX"
                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                        , id_pptx_path = ruta_fichero_pptx)
                        

                        #se actualizan atributos de la clase (se usa en lista_dicc_widgets, bloque siguiente)
                        self.widget_slide_pptx_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_SLIDES_PPTX", id_pptx = self.id_pptx_selecc)
                        

                        #se informan los cambios en la GUI
                        lista_dicc_widgets = [      
                                                {"tipo_widget": "scrolledtext"
                                                    , "widget_objeto": self.widget_id_pptx_path
                                                    , "variable_enlace": None
                                                    , "height": 1
                                                    , "bloquear": True
                                                    , "combobox_lista_opciones": None
                                                    , "combobox_opciones_editables": None
                                                    , "treeview_seleccionar_item": None
                                                    , "valor_informar": id_pptx_path_ajust
                                                    }

                                                , {"tipo_widget": "scrolledtext"
                                                    , "widget_objeto": self.widget_id_pptx_num_slides
                                                    , "variable_enlace": None
                                                    , "height": 1
                                                    , "bloquear": True
                                                    , "combobox_lista_opciones": None
                                                    , "combobox_opciones_editables": None
                                                    , "treeview_seleccionar_item": None
                                                    , "valor_informar": id_pptx_numero_total_slides
                                                    }

                                                , {"tipo_widget": "combobox"
                                                    , "widget_objeto": self.widget_slide_pptx
                                                    , "variable_enlace": self.strvar_widget_slide_pptx
                                                    , "height": None
                                                    , "bloquear": True
                                                    , "combobox_lista_opciones": self.widget_slide_pptx_lista_opciones
                                                    , "combobox_opciones_editables": False
                                                    , "treeview_seleccionar_item": None
                                                    , "valor_informar": self.strvar_widget_slide_pptx.get()
                                                    }
                                            ]
                            
                        self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)


                        if mensaje_warning_2 is not None:
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje_warning_2}))



        #############################################################################################################
        # ABRIR_PPTX
        #############################################################################################################
        elif proceso_selecc == "ABRIR_PPTX" and proceso_ejecutable:
            #abre el fichero pptx asociado al id pptx

            id_pptx_path = self.widget_id_pptx_path.texto_informado("todo").replace("\n", "")


            if len(id_pptx_path) == 0:
                mensaje = "No hay ninguna ruta configurada."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:
                mensaje = "Se abrira el fichero pptx que se usa para colocar los pantallazos de rangos de celdas excel.\n\nDeseas continuar?"
                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

                if msg.valor_boton_pulsado:
                    mensaje_warning = mod_back_end.def_varios_gui_config_id_pptx("ABRIR_FICHERO", ruta_fichero = id_pptx_path)

                    if mensaje_warning is not None:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_warning}))



        #############################################################################################################
        # ABRIR_VENTANA_CREACION_ID_XLS_NUEVO
        #############################################################################################################
        elif proceso_selecc == "ABRIR_VENTANA_CREACION_ID_XLS_NUEVO" and proceso_ejecutable:
            #abre un filedialog para seleccionar una ruta xls y abre una nueva ventana donde informar los datos gernerales asociados al nuevo id xls
            #se realiza previamente un check al seleccionar una ruta xls por si la ruta ya esta configurada salga un warning avisando de ello e informando en que id xls se configuro


            mensaje = "Se abrira una ventana de dialogo para que selecciones la ubicación del excel de origen que quieres usar y posteriormente se abrira una ventana donde tendras que configurar los datos generales del id xls a crear."
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

            if msg.valor_boton_pulsado:

                #se abre el file dialog para seleccionar la ubicacion del excel
                id_xls_nuevo_path = fd.askopenfilename(parent = self.master.widget_objeto, title = "SELECCIONA UNA NUEVA UBICACIÓN DE UN EXCEL DE ORIGEN:", filetypes = [("Ficheros Excel", "*.xls;*.xlsx;*.xlsb;*.xlsm")])

                if id_xls_nuevo_path:

                    #se asigna un nuevo id xls y se crea el mensaje para el warning
                    #en caso de que la ubicacion este ya asignada a otros id pptx
                    mensaje_warning = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_PATH_ID_XLS_YA_ASIGNADO"
                                                                                , id_pptx = self.id_pptx_selecc
                                                                                , id_xls_path = id_xls_nuevo_path)

                    id_xls_nuevo = mod_back_end.def_varios_gui_config_id_pptx("ASIGNAR_ID_XLS_NUEVO", id_pptx = self.id_pptx_selecc)


                    #a diferencia de la opcion ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO no se puede configurar un mismo excel varias veces para un mismo id pptx
                    if mensaje_warning is not None:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_warning}))

                    else:                

                        #se abre una nueva ventana para configurar el id xls creado
                        kwargs_gui_config_id_pptx_nuevo_id_xls_dicc_config_root = self.kwargs_config_gui["gui_config_id_pptx_nuevo_id_xls"]["dicc_config_root"]

                        self.toplevel_gui_config_id_pptx_nuevo_id_xls = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_config_id_pptx_nuevo_id_xls_dicc_config_root)
                        self.toplevel_gui_config_id_pptx_nuevo_id_xls.config_atributos(**kwargs_gui_config_id_pptx_nuevo_id_xls_dicc_config_root)

                        gui_config_id_pptx_nuevo_id_xls(self.toplevel_gui_config_id_pptx_nuevo_id_xls
                                                        , id_pptx = self.id_pptx_selecc
                                                        , id_xls_nuevo = id_xls_nuevo
                                                        , id_xls_nuevo_path = id_xls_nuevo_path
                                                        , entorno_clase_gui_config_id_pptx = self
                                                        , **self.kwargs_config_gui)     



        #############################################################################################################
        # ELIMINAR_ID_XLS
        #############################################################################################################
        elif proceso_selecc == "ELIMINAR_ID_XLS" and proceso_ejecutable:
            #elimina en memoria el id pptx y vacia los widgets
            #reactualiza la lista de opciones del combobox de seleccion de id pptx quitando el id pptx)

            mensaje1 = f"Se eliminara el id xls '{self.id_xls_selecc}' y sus configuraciones.\n\nLa eliminación tan solo se hara en la memoria del pc.\n\n"
            mensaje2 = "Tendrás que pulsar el botón 'GUARDAR' para que la eliminación también se realice en el sistema sqlite.\n\nDeseas continuar?"
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje1 + mensaje2}))

            if msg.valor_boton_pulsado:

                #se ejecuta la eliminacion en memoria del id xls asociado al id pptx
                #funciona tambien como funcion y devuelve el df sin el id xls para poder actualizar el treeview en la GUI
                df_treeview_id_xls_tras_eliminacion = mod_back_end.def_varios_gui_config_id_pptx("ELIMINAR_ID_XLS"
                                                                                                    , id_pptx = self.id_pptx_selecc
                                                                                                    , id_xls = self.id_xls_selecc)
                

                #se borran y se bloquean los widgets de GUI de config (los que afectan a los id xls y los rangos de celdas)
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS")


                #se actualizan los widgets en la ventana GUI de config de los id pptx
                lista_dicc_widgets = [
                                        {"tipo_widget": "treeview"
                                            , "widget_objeto": self.widget_treeview_id_xls
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": df_treeview_id_xls_tras_eliminacion
                                            }
                                    ]
                
                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                , lista_dicc_widgets = lista_dicc_widgets)
                
                #se actualizan atributos de la clase
                self.id_xls_selecc = None
                self.id_xls_selecc_antes_click_item = None
                self.rango_celdas_selecc = None

                if self.id_pptx_selecc is not None and self.id_xls_selecc is not None:
                    self.widget_hojas_xls_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_HOJAS_XLS"
                                                                                                    , id_pptx = self.id_pptx_selecc
                                                                                                    , id_xls = self.id_xls_selecc)


                mensaje = "id xls eliminado en memoria.\n\nLos cambios se reflejaran en el sistema sqlite una vez que pulses el botón GUARDAR."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))




        #############################################################################################################
        # UPDATE_PATH_ID_XLS
        #############################################################################################################
        elif proceso_selecc == "UPDATE_PATH_ID_XLS" and proceso_ejecutable:
            #abre un filedialog para seleccionar una ruta xls y se realiza un check por si la ruta ya esta configurada en otro id xls asociado al mismo id pptx
            #para que salga un warning no bloquante avisando de ello

            mensaje1 = "Se abrira una ventana de dialogo para que puedes configurar una nueva ubicación de fichero excel associado al id xls para el mismo id pptx.\n\n"
            mensaje2 = "Se realizara un check que generara un warning (bloquante) en caso de que la ruta que intentas configurar ya esta asignada a otro id xls para el mismo id pptx."
            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje1 + mensaje2}))

            if msg.valor_boton_pulsado:

                ruta_fichero_xls = fd.askopenfilename(parent = self.master.widget_objeto, title = "SELECCIONA UNA NUEVA UBICACIÓN DE UN EXCEL DE ORIGEN:", filetypes = [("Ficheros Excel", "*.xls;*.xlsx;*.xlsb;*.xlsm")])

                if ruta_fichero_xls:

                    mensaje_warning = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_PATH_ID_XLS_YA_ASIGNADO"
                                                                                    , id_pptx = self.id_pptx_selecc
                                                                                    , id_xls_path = ruta_fichero_xls
                                                                                    , id_xls_excluir_check_path = self.id_xls_selecc)

                    
                    #a diferencia de la opcion ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO no se puede configurar un mismo excel varias veces para un mismo id pptx
                    if mensaje_warning is not None:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_warning}))

                    else:
                        #se actualiza en la variable global global_dicc_datos_id_pptx el nuevo path y su respectivo numero de slides
                        #funciona tambien como funcion y devuelve mensaje warning (no bloquante) avisando que las hojas de rangos de celdas excel
                        #configurados previamente pueden estar desactualizados
                        (mensaje_warning_2
                        , id_pptx_path_ajust
                        , lista_hojas_xls
                        ) = mod_back_end.def_varios_gui_config_id_pptx("UPDATE_PATH_ID_XLS"
                                                                        , id_pptx = self.id_pptx_selecc
                                                                        , id_xls = self.id_xls_selecc
                                                                        , id_xls_path = ruta_fichero_xls)
                        

                        #se actualiza el atributo propio de la presente clase
                        self.widget_hojas_xls_lista_opciones = lista_hojas_xls
                        

                        #se actualizan atributos de la clase (se usan en lista_dicc_widgets, bloque siguiente)
                        if self.id_pptx_selecc is not None and self.id_xls_selecc is not None:
                            self.widget_hojas_xls_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_HOJAS_XLS"
                                                                                                            , id_pptx = self.id_pptx_selecc
                                                                                                            , id_xls = self.id_xls_selecc)
                            

                        #se informan los cambios en la GUI (la ruta del id xls y la lista de opcionesdel combobox de hojas xls)
                        lista_dicc_widgets = [      
                                                {"tipo_widget": "scrolledtext"
                                                    , "widget_objeto": self.widget_id_xls_path
                                                    , "variable_enlace": None
                                                    , "height": 1
                                                    , "bloquear": True
                                                    , "combobox_lista_opciones": None
                                                    , "combobox_opciones_editables": None
                                                    , "treeview_seleccionar_item": None
                                                    , "valor_informar": id_pptx_path_ajust
                                                    }

                                                , {"tipo_widget": "combobox"
                                                    , "widget_objeto": self.widget_hojas_xls
                                                    , "variable_enlace": self.strvar_widget_hojas_xls
                                                    , "height": None
                                                    , "bloquear": True
                                                    , "combobox_lista_opciones": self.widget_hojas_xls_lista_opciones
                                                    , "combobox_opciones_editables": False
                                                    , "treeview_seleccionar_item": None
                                                    , "valor_informar": self.strvar_widget_hojas_xls.get()
                                                    }
                                            ]
                        
                        self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC", lista_dicc_widgets = lista_dicc_widgets)


                        if mensaje_warning_2 is not None:
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showwarning", "mensaje": mensaje_warning_2}))



        #############################################################################################################
        # ABRIR_XLS
        #############################################################################################################
        elif proceso_selecc == "ABRIR_XLS" and proceso_ejecutable:
            #abre el fichero xls asociado al id pptx

            id_xls_path = self.widget_id_xls_path.texto_informado("todo").replace("\n", "")


            if len(id_xls_path) == 0:
                mensaje = "No hay ninguna ruta configurada."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:
                mensaje = "Se abrira el fichero excel.\n\nDeseas continuar?"
                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

                if msg.valor_boton_pulsado:
                    mensaje_warning = mod_back_end.def_varios_gui_config_id_pptx("ABRIR_FICHERO", ruta_fichero = id_xls_path)

                    if mensaje_warning is not None:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_warning}))




        #############################################################################################################
        # LIMPIAR_RANGO_CELDAS
        #############################################################################################################
        elif proceso_selecc == "LIMPIAR_RANGO_CELDAS" and proceso_ejecutable:

            if self.id_xls_selecc is None:
                mensaje = "No has seleccionado ningún excel."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:

                mensaje = "Se limpiarán los widgets de rangos de celdas paraque puedes informar uno nuvo que podrás almacenar luego en memoria cuando pulses el botón AGREGAR.\n\nDeseas continuar?"
                msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

                if msg.valor_boton_pulsado:
                    

                    #se borran y se bloquean los widgets de GUI de config (los que afectan a los los rangos de celdas salvo el treeview)
                    self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW")


                    #se actualizan los widgets en la ventana GUI de config de los id pptx
                    #se usa el mismo metodo que el bloque anterior
                    lista_dicc_widgets = [
                                            {"tipo_widget": "combobox"
                                                , "widget_objeto": self.widget_hojas_xls
                                                , "variable_enlace": self.strvar_widget_hojas_xls
                                                , "height": None
                                                , "bloquear": False
                                                , "combobox_lista_opciones": self.widget_hojas_xls_lista_opciones #se calcula en otras interacciones de la GUI
                                                , "combobox_opciones_editables": False
                                                , "treeview_seleccionar_item": None
                                                , "valor_informar": None
                                                }

                                            , {"tipo_widget": "entry_propio"
                                                , "widget_objeto": self.widget_rangos_celdas
                                                , "variable_enlace": self.strvar_widget_rangos_celdas
                                                , "height": None
                                                , "bloquear": False
                                                , "combobox_lista_opciones": None
                                                , "combobox_opciones_editables": None
                                                , "treeview_seleccionar_item": None
                                                , "valor_informar": None
                                                }

                                            , {"tipo_widget": "combobox"
                                                , "widget_objeto": self.widget_slide_pptx
                                                , "variable_enlace": self.strvar_widget_slide_pptx
                                                , "height": None
                                                , "bloquear": False
                                                , "combobox_lista_opciones": self.widget_slide_pptx_lista_opciones #se calcula en otras interacciones de la GUI
                                                , "combobox_opciones_editables": False
                                                , "treeview_seleccionar_item": None
                                                , "valor_informar": None
                                                }

                                            , {"tipo_widget": "scrolledtext"
                                                , "widget_objeto": self.widget_nombre_pantallazo
                                                , "variable_enlace": None
                                                , "height": None
                                                , "bloquear": True
                                                , "combobox_lista_opciones": None
                                                , "combobox_opciones_editables": None
                                                , "treeview_seleccionar_item": None
                                                , "valor_informar": None
                                                }

                                            , {"tipo_widget": "scrolledtext"
                                                , "widget_objeto": self.widget_coordenadas_pantallazo
                                                , "variable_enlace": None
                                                , "height": self.widget_id_xls_desc_height
                                                , "bloquear": True
                                                , "combobox_lista_opciones": None
                                                , "combobox_opciones_editables": None
                                                , "treeview_seleccionar_item": None
                                                , "valor_informar": None
                                                }

                                        ]
                    
                    self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                    , lista_dicc_widgets = lista_dicc_widgets)



        #############################################################################################################
        # AGREGAR_RANGO_CELDAS
        #############################################################################################################
        elif proceso_selecc == "AGREGAR_RANGO_CELDAS" and proceso_ejecutable:
            #agrega un rango de celdas al treeview y incorpora en la variable global global_dicc_datos_id_pptx
            #se chequea previamente si el rango de celdas ya se incoporo previamente
            #para elo mismo id xls y id pptx y para la misma hoja xls y slide pptx



            hoja_xls = self.strvar_widget_hojas_xls.get()
            rango_celdas = self.strvar_widget_rangos_celdas.get()
            slide_pptx = self.strvar_widget_slide_pptx.get()

            if self.id_xls_selecc is None:
                mensaje = "No has seleccionado ningún excel."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:

                if self.id_xls_selecc is None:
                    mensaje = "No has seleccionado ningún id xls."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                else:
                    if len(hoja_xls) == 0 or len(rango_celdas) == 0 or len(slide_pptx) == 0:
                        mensaje = "La hoja excel, el rango de celdas y el nº de slides pptx son obligatorios."
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                    else:

                        mensaje_gui_warning = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_RANGO_CELDAS_YA_ASIGNADO"
                                                                                            , id_pptx = self.id_pptx_selecc
                                                                                            , id_xls = self.id_xls_selecc
                                                                                            , hoja_xls = hoja_xls
                                                                                            , rango_celdas = rango_celdas
                                                                                            , slide_pptx = slide_pptx
                                                                                            )
                        if mensaje_gui_warning is not None:
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_gui_warning}))

                        else:

                            mensaje = "Se agregara en memoria el rango de celdas (junto con la hoja excel y el nº de slide pptx).\n\nNo se registrarán los cambios en el sistema sqlite hasta que pulses el botón GUARDAR.\n\n"
                            msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel", "mensaje": mensaje}))

                            if msg.valor_boton_pulsado:


                                #se actualiza el treeview con el rango agregado y se crea la lista de lista para dejar seleccionado el item dentro del treeview
                                #lista_rangos_celdas_item_seleccionar_treeview tiene que se lista de lisyas con 1 sola sublista
                                (lista_datos_items_seleccionados
                                , df_widget_treeview_rangos_celdas_nuevo) = mod_back_end.def_varios_gui_config_id_pptx("AGREGAR_RANGO_CELDAS"
                                                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                                                        , id_xls = self.id_xls_selecc
                                                                                                                        , hoja_xls = hoja_xls
                                                                                                                        , rango_celdas = rango_celdas
                                                                                                                        , slide_pptx = slide_pptx
                                                                                                                        )
                                

                                #se borran y se bloquean los widgets de GUI de config (los que afectan a los los rangos de celdas salvo el treeview)
                                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW")


                                #se actualizan los widgets en la ventana GUI de config de los id pptx
                                #se usa el mismo metodo que el bloque anterior
                                lista_dicc_widgets = [
                                                        {"tipo_widget": "treeview"
                                                            , "widget_objeto": self.widget_treeview_rangos_celdas
                                                            , "variable_enlace": None
                                                            , "height": None
                                                            , "bloquear": False
                                                            , "combobox_lista_opciones": None
                                                            , "combobox_opciones_editables": False
                                                            , "treeview_seleccionar_item": lista_datos_items_seleccionados
                                                            , "valor_informar": df_widget_treeview_rangos_celdas_nuevo
                                                            }

                                                        , {"tipo_widget": "combobox"
                                                            , "widget_objeto": self.widget_hojas_xls
                                                            , "variable_enlace": self.strvar_widget_hojas_xls
                                                            , "height": None
                                                            , "bloquear": False
                                                            , "combobox_lista_opciones": self.widget_hojas_xls_lista_opciones #se calcula en otras interacciones de la GUI
                                                            , "combobox_opciones_editables": False
                                                            , "treeview_seleccionar_item": None
                                                            , "valor_informar": hoja_xls
                                                            }

                                                        , {"tipo_widget": "entry_propio"
                                                            , "widget_objeto": self.widget_rangos_celdas
                                                            , "variable_enlace": self.strvar_widget_rangos_celdas
                                                            , "height": None
                                                            , "bloquear": False
                                                            , "combobox_lista_opciones": None
                                                            , "combobox_opciones_editables": None
                                                            , "treeview_seleccionar_item": None
                                                            , "valor_informar": rango_celdas
                                                            }

                                                        , {"tipo_widget": "combobox"
                                                            , "widget_objeto": self.widget_slide_pptx
                                                            , "variable_enlace": self.strvar_widget_slide_pptx
                                                            , "height": None
                                                            , "bloquear": False
                                                            , "combobox_lista_opciones": self.widget_slide_pptx_lista_opciones #se calcula en otras interacciones de la GUI
                                                            , "combobox_opciones_editables": False
                                                            , "treeview_seleccionar_item": None
                                                            , "valor_informar": slide_pptx
                                                            }

                                                        , {"tipo_widget": "scrolledtext"
                                                            , "widget_objeto": self.widget_nombre_pantallazo
                                                            , "variable_enlace": None
                                                            , "height": None
                                                            , "bloquear": True
                                                            , "combobox_lista_opciones": None
                                                            , "combobox_opciones_editables": None
                                                            , "treeview_seleccionar_item": None
                                                            , "valor_informar": None
                                                            }

                                                        , {"tipo_widget": "scrolledtext"
                                                            , "widget_objeto": self.widget_coordenadas_pantallazo
                                                            , "variable_enlace": None
                                                            , "height": self.widget_id_xls_desc_height
                                                            , "bloquear": True
                                                            , "combobox_lista_opciones": None
                                                            , "combobox_opciones_editables": None
                                                            , "treeview_seleccionar_item": None
                                                            , "valor_informar": None
                                                            }

                                                    ]
                                
                                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                                , lista_dicc_widgets = lista_dicc_widgets)


                                #se informa el atributo datos_items del objeto treeview rangos de celdas
                                self.widget_treeview_rangos_celdas.datos_items["lista_datos_items_seleccionados"] = lista_datos_items_seleccionados

                                #see actualiza el atributo rango_celdas_selecc (es el 1er item de la lista lista_datos_items_seleccionados)
                                self.rango_celdas_selecc = lista_datos_items_seleccionados[0]
                                

                                mensaje = "Rango de celdas agregado en la memoria.\n\nTendrás que pulsar el botón 'GUARDAR' para que la agregación también se realice en el sistema sqlite.\n\n"
                                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje})) 



        #############################################################################################################
        # ELIMINAR_RANGO_CELDAS
        #############################################################################################################
        elif proceso_selecc == "ELIMINAR_RANGO_CELDAS" and proceso_ejecutable:
            #elimina el rango de celdas seleecionado en el treeview y lo elimina en la variable global global_dicc_datos_id_pptx

            hoja_xls = self.strvar_widget_hojas_xls.get()
            rango_celdas = self.strvar_widget_rangos_celdas.get()
            slide_pptx = self.strvar_widget_slide_pptx.get()


            if self.id_xls_selecc is None:
                mensaje = "No has seleccionado ningún excel."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:

                if self.id_xls_selecc is None:
                    mensaje = "No has seleccionado ningún id xls."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                else:
                    if len(hoja_xls) == 0 or len(rango_celdas) == 0 or len(slide_pptx) == 0:
                        mensaje = "No has seleccionado ningún rango de celdas."
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                    else:
                        mensaje = "Se eliminara en memoria en rango de celdas seleccionado en el treeview.\n\nNo se registrarán los cambios en el sistema sqlite hasta que pulses el botón GUARDAR.\n\nDeseas continuar?"
                        msg = mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "askokcancel_warning", "mensaje": mensaje}))

                        if msg.valor_boton_pulsado:

                            #se elimina el rango de celdas y se recuperan las listas de opciones para los combobox de hojasxls y de slides pptx
                            #se actualiza el treeview con el rango agregado
                            df_widget_treeview_rangos_celdas_nuevo = mod_back_end.def_varios_gui_config_id_pptx("ELIMINAR_RANGO_CELDAS"
                                                                                                                , id_pptx = self.id_pptx_selecc
                                                                                                                , id_xls = self.id_xls_selecc
                                                                                                                , hoja_xls = hoja_xls
                                                                                                                , rango_celdas = rango_celdas
                                                                                                                , slide_pptx = slide_pptx
                                                                                                                )
                            

                            #se borran y se bloquean los widgets de GUI de config (los que afectan a los los rangos de celdas salvo el treeview)
                            self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW")


                            #se actualizan los widgets en la ventana GUI de config de los id pptx
                            #se usa el mismo metodo que el bloque anterior
                            lista_dicc_widgets = [
                                                    {"tipo_widget": "treeview"
                                                        , "widget_objeto": self.widget_treeview_rangos_celdas
                                                        , "variable_enlace": None
                                                        , "height": None
                                                        , "bloquear": False
                                                        , "combobox_lista_opciones": None
                                                        , "combobox_opciones_editables": False
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": df_widget_treeview_rangos_celdas_nuevo
                                                        }

                                                    , {"tipo_widget": "combobox"
                                                        , "widget_objeto": self.widget_hojas_xls
                                                        , "variable_enlace": self.strvar_widget_hojas_xls
                                                        , "height": None
                                                        , "bloquear": False
                                                        , "combobox_lista_opciones": self.widget_hojas_xls_lista_opciones #se calcula en otras interacciones de la GUI
                                                        , "combobox_opciones_editables": False
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": None
                                                        }

                                                    , {"tipo_widget": "entry_propio"
                                                        , "widget_objeto": self.widget_rangos_celdas
                                                        , "variable_enlace": self.strvar_widget_rangos_celdas
                                                        , "height": None
                                                        , "bloquear": False
                                                        , "combobox_lista_opciones": None
                                                        , "combobox_opciones_editables": None
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": None
                                                        }

                                                    , {"tipo_widget": "combobox"
                                                        , "widget_objeto": self.widget_slide_pptx
                                                        , "variable_enlace": self.strvar_widget_slide_pptx
                                                        , "height": None
                                                        , "bloquear": False
                                                        , "combobox_lista_opciones": self.widget_slide_pptx_lista_opciones #se calcula en otras interacciones de la GUI
                                                        , "combobox_opciones_editables": False
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": None
                                                        }

                                                    , {"tipo_widget": "scrolledtext"
                                                        , "widget_objeto": self.widget_nombre_pantallazo
                                                        , "variable_enlace": None
                                                        , "height": None
                                                        , "bloquear": True
                                                        , "combobox_lista_opciones": None
                                                        , "combobox_opciones_editables": None
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": None
                                                        }

                                                    , {"tipo_widget": "scrolledtext"
                                                        , "widget_objeto": self.widget_coordenadas_pantallazo
                                                        , "variable_enlace": None
                                                        , "height": self.widget_id_xls_desc_height
                                                        , "bloquear": True
                                                        , "combobox_lista_opciones": None
                                                        , "combobox_opciones_editables": None
                                                        , "treeview_seleccionar_item": None
                                                        , "valor_informar": None
                                                        }

                                                ]
                            
                            self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                            , lista_dicc_widgets = lista_dicc_widgets)
                            
                            mensaje = "Rango de celdas eliminado en la memoria.\n\nTendrás que pulsar el botón 'GUARDAR' para que la eliminación también se realice en el sistema sqlite.\n\n"
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje})) 



        #############################################################################################################
        # UPDATE_RANGO_CELDAS
        #############################################################################################################
        elif proceso_selecc == "UPDATE_RANGO_CELDAS" and proceso_ejecutable:
            #actualiza el treeview de rangos en memoria y en la GUI de celdas solo si se modifican o la hoja o el rango de celdas o la slide pptx de destino
            #y selecciona el item modificado en el treeview

            hoja_xls = self.strvar_widget_hojas_xls.get()
            rango_celdas = self.strvar_widget_rangos_celdas.get()
            slide_pptx = self.strvar_widget_slide_pptx.get()



            if self.rango_celdas_selecc is None:
                mensaje = "No has seleccionado ningún rango de celdas."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:
                if len(hoja_xls) == 0 or len(rango_celdas) == 0 or len(slide_pptx) == 0:
                    mensaje = "La hoja excel, el rango de celdas y la slide ppts son obligatorios."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                else:

                    mensaje_gui_warning = mod_back_end.def_varios_gui_config_id_pptx("CHECK_SI_RANGO_CELDAS_YA_ASIGNADO"
                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                        , id_xls = self.id_xls_selecc
                                                                                        , hoja_xls = hoja_xls
                                                                                        , rango_celdas = rango_celdas
                                                                                        , slide_pptx = slide_pptx
                                                                                        )


                    if mensaje_gui_warning is not None:
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje_gui_warning}))

                    else:

                        lista_rangos_celdas_antes_update = [str(item) for item in self.rango_celdas_selecc]
                        lista_rangos_celdas_despues_update = [hoja_xls, rango_celdas, slide_pptx]

                        if lista_rangos_celdas_despues_update == lista_rangos_celdas_antes_update:
                            mensaje = "No has modificado el rango de celdas."
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                        else:

                            self.master.widget_objeto.config(cursor = "wait")

                            (lista_datos_items_seleccionados
                            , df_widget_treeview_rangos_celdas_actualizado) = mod_back_end.def_varios_gui_config_id_pptx("UPDATE_RANGO_CELDAS"
                                                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                                                        , id_xls = self.id_xls_selecc
                                                                                                                        , lista_rangos_celdas_antes_update = self.rango_celdas_selecc
                                                                                                                        , lista_rangos_celdas_despues_update = lista_rangos_celdas_despues_update)
                            

                            #se actualizan el treeview en la GUi y se seleciona el item actualizado
                            lista_dicc_widgets = [
                                                    {"tipo_widget": "treeview"
                                                        , "widget_objeto": self.widget_treeview_rangos_celdas
                                                        , "variable_enlace": None
                                                        , "height": None
                                                        , "bloquear": False
                                                        , "combobox_lista_opciones": None
                                                        , "combobox_opciones_editables": False
                                                        , "treeview_seleccionar_item": lista_datos_items_seleccionados
                                                        , "valor_informar": df_widget_treeview_rangos_celdas_actualizado
                                                        }
                                                ]
                            
                            self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                            , lista_dicc_widgets = lista_dicc_widgets)
                            

                            #se informa el atributo datos_items del objeto treeview rangos de celdas
                            self.widget_treeview_rangos_celdas.datos_items["lista_datos_items_seleccionados"] = lista_datos_items_seleccionados
                            

                            self.master.widget_objeto.config(cursor = "")
                            
                            mensaje = "Rango de celdas actualizado en la memoria.\n\nTendrás que pulsar el botón 'GUARDAR' para que la eliminación también se realice en el sistema sqlite.\n\n"
                            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje})) 








        #############################################################################################################
        # MOSTRAR_SCREENSHOT_RANGO_CELDAS
        #############################################################################################################
        elif proceso_selecc == "MOSTRAR_SCREENSHOT_RANGO_CELDAS" and proceso_ejecutable:
            #abre una nueva ventana y muestra el screenshot realizado en el proceso de configuracion
            #del pptx 'CONFIG_PASO_1' siempre y cuando se localice el screenshot de muestra en la ruta local

            if self.id_xls_selecc is None:
                mensaje = "No has seleccionado ningún excel."
                mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

            else:

                nombre_screenshot = self.widget_nombre_pantallazo.texto_informado("todo")

                #si el usuario ha optado por no descargar las muestras de los screenshots en png o no al acceder a la ventana que genera la presente clase
                #se ejecuta o no este proceso
                #CASO 1 - se opto por la muestras
                if self.descargar_screenshots_muestra_en_png:

                    path_screenshot_png = mod_back_end.def_varios_gui_config_id_pptx("SCREENSHOT_PNG_MUESTRA_PATH"
                                                                                    , nombre_screenshot = nombre_screenshot)
                
                    if path_screenshot_png is not None:

                        kwargs_gui_screenshot_muestra_dicc_config_root = self.kwargs_config_gui["gui_screenshot_muestra"]["dicc_config_root"]

                        self.toplevel_gui_screenshot_muestra = mod_utils.gui_tkinter_widgets(self.master.widget_objeto, tipo_widget_param = "toplevel", **kwargs_gui_screenshot_muestra_dicc_config_root)
                        self.toplevel_gui_screenshot_muestra.config_atributos(**kwargs_gui_screenshot_muestra_dicc_config_root)

                        gui_screenshot_muestra(self.toplevel_gui_screenshot_muestra
                                                , path_screenshot_png = path_screenshot_png
                                                , **self.kwargs_config_gui)
                        
                    else:
                        mensaje = "No se ha localizado ningún pantallazo de muestra."
                        mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))

                #CASO 2 - NO se opto por la muestras
                elif not self.descargar_screenshots_muestra_en_png:
                    mensaje = "Al acceder a esta ventana desde la ventana de inicio optaste por no descargar las muestras de los pantallazos."
                    mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))



    def def_gui_config_id_pptx_treeview_click_item(self, opcion_treeview):
        #rutina que actualiza otros widgets de la ventana al clicar sobre un item de los treeviews


        if opcion_treeview == "ID_XLS":

            #se guardan los cambios en la variable global realizados en los datos del id xls seleccionado antes de hacer
            #click de nuevo en un item por si no se han guardado aun en sitema sqlite sino los datos se pierden
            (lista_datos_items_seleccionados
            , df_widget_treeview_id_xls) = mod_back_end.def_varios_gui_config_id_pptx("UPDATE_EN_MEMORIA_DATOS_GUI_ANTES_HACER_CLICK_EN_OTRO_ID_XLS"
                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                        , id_xls = self.id_xls_selecc_antes_click_item
                                                                                        , id_xls_desc = self.widget_id_xls_desc.texto_informado("todo")
                                                                                        , id_xls_path = self.widget_id_xls_path.texto_informado("todo")
                                                                                        , id_xls_tiempo_espera_max_apertura = self.strvar_widget_tiempo_apertura_max_id_xls.get()
                                                                                        , id_xls_actualizar_vinculos_otros_xls = self.strvar_widget_actualizar_vinculos_otros_excel.get()
                                                                                        )
            

            #se recupera el id xls seleccionado en el treeview
            self.id_xls_selecc = (self.widget_treeview_id_xls.datos_items["lista_datos_items_seleccionados"][0][0]
                                    if self.widget_treeview_id_xls.datos_items["lista_datos_items_seleccionados"] is not None
                                    else None)
            
            self.widget_treeview_id_xls.datos_items["lista_datos_items_seleccionados"] = lista_datos_items_seleccionados


            #se actualiza la GUI (tan solo se se ha hecho click en un item)
            if self.id_xls_selecc is not None:

                (click_item_id_xls_path
                , click_item_id_xls_desc
                , click_item_id_xls_tiempo_espera_max_apertura
                , click_item_id_xls_actualizar_vinculos_otros_xls
                , click_item_df_widget_treeview_rangos_celdas
                ) = mod_back_end.def_varios_gui_config_id_pptx("CLICK_ITEM_TREEVIEW_ID_XLS"
                                                                , id_pptx = self.id_pptx_selecc
                                                                , id_xls = self.id_xls_selecc
                                                                )



                #se crea la lista lista_datos_items_seleccionados_actualizada para poder actualizar el treeview
                #en caso de que el usuario cambie la descripcion y luego clica en otro item y luego vuelve al anterior
                #si no se realiza este ajuste la descripcion en el treeview no se actualiza pero en el widget de la descripcion
                #si aparece el cambio
                #tiene que ser de lista con 1 sola sublista
                lista_datos_items_selecc_update_desc = [[self.id_xls_selecc, click_item_id_xls_desc]]
                self.widget_treeview_id_xls.datos_items["lista_datos_items_seleccionados"] = lista_datos_items_selecc_update_desc

                df_widget_treeview_id_xls_update = mod_back_end.def_varios_gui_config_id_pptx("DF_TREEVIEW_ID_XLS_TRAS_CAMBIOS_EN_DESCRIPCION"
                                                                                              , id_pptx = self.id_pptx_selecc
                                                                                              , lista_datos_items_selecc_update_desc = lista_datos_items_selecc_update_desc)


                #se actulaiza el atributo de la presente clase widget_hojas_xls_lista_opciones
                if self.id_pptx_selecc is not None and self.id_xls_selecc is not None:
                    self.widget_hojas_xls_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_HOJAS_XLS"
                                                                                                        , id_pptx = self.id_pptx_selecc
                                                                                                        , id_xls = self.id_xls_selecc)
                

                #se borran y se bloquean los widgets de GUI de config (los que afectan a los los rangos de celdas salvo el treeview)
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW")


                #se actualizan los widgets en la ventana GUI de config de los id pptx
                #se usa el mismo metodo que el bloque anterior
                lista_dicc_widgets_1 = [
                                        {"tipo_widget": "treeview"
                                            , "widget_objeto": self.widget_treeview_id_xls
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": lista_datos_items_selecc_update_desc
                                            , "valor_informar": df_widget_treeview_id_xls_update
                                            }
                                        ]

                lista_dicc_widgets_2 = [
                                        {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_xls
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": self.id_xls_selecc
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_xls_path
                                            , "variable_enlace": None
                                            , "height": 1
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_id_xls_path
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_id_xls_desc
                                            , "variable_enlace": None
                                            , "height": self.widget_id_xls_desc_height
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_id_xls_desc
                                            }

                                        , {"tipo_widget": "entry_propio"
                                            , "widget_objeto": self.widget_tiempo_apertura_max_id_xls
                                            , "variable_enlace": self.strvar_widget_tiempo_apertura_max_id_xls
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_id_xls_tiempo_espera_max_apertura
                                            }

                                        , {"tipo_widget": "combobox"
                                            , "widget_objeto": self.widget_actualizar_vinculos_otros_excel
                                            , "variable_enlace": self.strvar_widget_actualizar_vinculos_otros_excel
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": mod_back_end.lista_opciones_id_xls_combobox_actualizar_vinculos
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_id_xls_actualizar_vinculos_otros_xls
                                            }

                                        , {"tipo_widget": "treeview"
                                            , "widget_objeto": self.widget_treeview_rangos_celdas
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_df_widget_treeview_rangos_celdas
                                            }
                                    ]
                
                if self.id_xls_selecc_antes_click_item is None:
                
                    self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                    , lista_dicc_widgets = lista_dicc_widgets_2)
                    
                else:
                    self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                , lista_dicc_widgets = lista_dicc_widgets_1)
                
                    self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                    , lista_dicc_widgets = lista_dicc_widgets_2)



            #se actualizan atributos de la clase
            #(self.id_xls_selecc_antes_click_item para usarlo si el usuario hace otro click en el item)
            self.id_xls_selecc_antes_click_item = self.id_xls_selecc



        elif opcion_treeview == "RANGOS_CELDAS":


            #se recupera el rango de celdas seleccionado en el treeview (es lista que contiene la hoja, el rango y la slide)
            #para la slide se realiza un ajuste para que aparezca como entero y no float
            self.rango_celdas_selecc = (self.widget_treeview_rangos_celdas.datos_items["lista_datos_items_seleccionados"][0]
                                        if self.widget_treeview_rangos_celdas.datos_items["lista_datos_items_seleccionados"] is not None
                                        else None)
            

            
            self.rango_celdas_selecc_con_slide_pptx_numerico = ([item if ind != 2 else int(float(item)) for ind, item in enumerate(self.rango_celdas_selecc)]
                                                                if isinstance(self.rango_celdas_selecc, list) else None)
            

            #se actualiza la GUI (tan solo se se ha hecho click en un item)
            if isinstance(self.rango_celdas_selecc_con_slide_pptx_numerico, list) and len(self.rango_celdas_selecc_con_slide_pptx_numerico) != 0:

                click_item_hoja_xls = self.rango_celdas_selecc_con_slide_pptx_numerico[0]
                click_item_rango_celdas = self.rango_celdas_selecc_con_slide_pptx_numerico[1]
                click_item_slide_pptx = self.rango_celdas_selecc_con_slide_pptx_numerico[2]

                (click_item_nombre_screenshot
                , click_item_coordenadas_screenshot
                ) = mod_back_end.def_varios_gui_config_id_pptx("CLICK_ITEM_TREEVIEW_RANGOS_CELDAS"
                                                                , id_pptx = self.id_pptx_selecc
                                                                , id_xls = self.id_xls_selecc
                                                                , hoja_xls = click_item_hoja_xls
                                                                , rango_celdas = click_item_rango_celdas
                                                                , slide_pptx = click_item_slide_pptx
                                                                )
                

                #se borran y se bloquean los widgets de GUI de config (los que afectan a los los rangos de celdas salvo el treeview)
                self.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW")


                #se actualizan los widgets en la ventana GUI de config de los id pptx
                #se usa el mismo metodo que el bloque anterior
                lista_dicc_widgets = [
                                        {"tipo_widget": "combobox"
                                            , "widget_objeto": self.widget_hojas_xls
                                            , "variable_enlace": self.strvar_widget_hojas_xls
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": self.widget_hojas_xls_lista_opciones
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_hoja_xls
                                            }

                                        , {"tipo_widget": "entry_propio"
                                            , "widget_objeto": self.widget_rangos_celdas
                                            , "variable_enlace": self.strvar_widget_rangos_celdas
                                            , "height": None
                                            , "bloquear": False
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_rango_celdas
                                            }

                                        , {"tipo_widget": "combobox"
                                            , "widget_objeto": self.widget_slide_pptx
                                            , "variable_enlace": self.strvar_widget_slide_pptx
                                            , "height": None
                                            , "bloquear": True
                                            , "combobox_lista_opciones": self.widget_slide_pptx_lista_opciones
                                            , "combobox_opciones_editables": False
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_slide_pptx
                                            }

                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_nombre_pantallazo
                                            , "variable_enlace": None
                                            , "height": None
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_nombre_screenshot
                                            }


                                        , {"tipo_widget": "scrolledtext"
                                            , "widget_objeto": self.widget_coordenadas_pantallazo
                                            , "variable_enlace": None
                                            , "height": self.widget_id_xls_desc_height
                                            , "bloquear": True
                                            , "combobox_lista_opciones": None
                                            , "combobox_opciones_editables": None
                                            , "treeview_seleccionar_item": None
                                            , "valor_informar": click_item_coordenadas_screenshot
                                            }

                                    ]
                
                self.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                , lista_dicc_widgets = lista_dicc_widgets)




    def def_gui_config_id_pptx_widgets_actualizar(self, opcion, **kwargs):
        #rutina que permite actualizar los widgets de la ventana de configuracion de los id pptx
        #de forma dinamica según la opcion seleccionada

        #parametros kwargs
        lista_dicc_widgets = kwargs.get("lista_dicc_widgets", None)  


        ######################################################################################
        # BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS
        # BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS
        ######################################################################################
        if opcion in ["BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS"
                      , "BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS"
                      , "BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW"]:
            
            #se borra todo el contenido y se bloquean todos los widgets (salvo los labels y botones)
            # --> BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS                                 se excluye el combobox de seleccion de id pptx
            #
            # --> BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS                 se excluye el combobox de seleccion de id pptx y loswidgets del id pptx
            #                                                                         (se usa cuando se crea un nuevo id xls o se elimina un id xls)
            #
            # --> BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW             se excluye el combobox de seleccion de id pptx, los widgets del id pptx, los widgets del id xls
            #                                                                         y el treeview de rangos de celdas
            #                                                                         (se usa cuando se hace click en un item del treeview por id xls o el el treeview por rangos de celdas)
            #     

            lista_widgets_excluidos = ([self.widget_seleccion_id_pptx]
                                       if opcion == "BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS"
                                       else
                                            [self.widget_seleccion_id_pptx
                                             
                                            , self.widget_id_pptx
                                            , self.widget_id_pptx_path
                                            , self.widget_id_pptx_desc
                                            , self.widget_tiempo_apertura_max_id_pptx
                                            , self.widget_id_pptx_num_slides
                                            ]
                                        if opcion == "BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS"
                                        else
                                            [self.widget_seleccion_id_pptx
                                             
                                            , self.widget_id_pptx
                                            , self.widget_id_pptx_path
                                            , self.widget_id_pptx_desc
                                            , self.widget_tiempo_apertura_max_id_pptx
                                            , self.widget_id_pptx_num_slides

                                            , self.widget_treeview_id_xls
                                            , self.widget_id_xls
                                            , self.widget_id_xls_path
                                            , self.widget_id_xls_desc
                                            , self.widget_tiempo_apertura_max_id_xls
                                            , self.widget_actualizar_vinculos_otros_excel

                                            , self.widget_treeview_rangos_celdas
                                            ]
                                        if opcion == "BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_CLICK_ITEM_TREEVIEW"
                                        else
                                        []
                                       )


            dicc_gui_config_id_pptx_widgets_borrar_y_bloquear = {key: dicc for key, dicc in self.dicc_gui_config_id_pptx_frame_widgets_objetos.items()
                                                                if dicc["widget_objeto"] not in lista_widgets_excluidos
                                                                    and dicc["tipo_widget"] in ["scrolledtext_propio", "treeview_propio", "entry_propio", "combobox"]
                                                                }

            for _, dicc_widget in dicc_gui_config_id_pptx_widgets_borrar_y_bloquear.items():

                tipo_widget = dicc_widget["tipo_widget"]
                widget_objeto = dicc_widget["widget_objeto"]
                widget_variable_enlace = dicc_widget["widget_variable_enlace"]

                if tipo_widget == "combobox":
                    widget_variable_enlace.set("")

                    if widget_objeto != self.widget_actualizar_vinculos_otros_excel:
                        widget_objeto.config_atributos(**{"combobox_lista_opciones": []})


                elif tipo_widget in ["scrolledtext_propio", "treeview_propio", "entry_propio"]:
                    widget_objeto.config_atributos(**{"bloquear": False})

                    if tipo_widget == "scrolledtext_propio":
                        widget_objeto.modificaciones("borrar_contenido_y_tags")

                    elif tipo_widget == "treeview_propio":
                        widget_objeto.acciones("eliminar_contenido")

                    elif tipo_widget == "entry_propio":
                        widget_variable_enlace.set("")


                widget_objeto.config_atributos(**{"bloquear": True})



        ######################################################################################
        # INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC
        ######################################################################################
        elif opcion == "INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC":
            #permite informar y desbloquear widgetsdesde una lista de diccionarios de widgets que se pasa como parametro kwargs


            for dicc in lista_dicc_widgets:

                tipo_widget = dicc["tipo_widget"]
                widget_objeto = dicc["widget_objeto"]
                variable_enlace = dicc["variable_enlace"]
                height = dicc["height"]
                bloquear = dicc["bloquear"]
                combobox_lista_opciones = dicc["combobox_lista_opciones"]
                combobox_opciones_editables = dicc["combobox_opciones_editables"]
                treeview_seleccionar_item = dicc["treeview_seleccionar_item"]
                valor_informar = dicc["valor_informar"]


                widget_objeto.config_atributos(**{"bloquear": False})

                if tipo_widget == "scrolledtext":

                    widget_objeto.modificaciones("borrar_contenido_y_tags")

                    widget_objeto.modificaciones("agregar_solo_contenido_desde_string"
                                                , string_texto_informar = valor_informar if valor_informar is not None else ""
                                                , height_scrolledtext = height)


                elif tipo_widget in ["entry", "entry_propio"]:
                    variable_enlace.set(valor_informar if valor_informar is not None else "")


                elif tipo_widget == "combobox":

                    widget_objeto.config_atributos(**{"combobox_lista_opciones": combobox_lista_opciones})
                    widget_objeto.config_atributos(**{"combobox_opciones_editables": combobox_opciones_editables})
                    variable_enlace.set(valor_informar if valor_informar is not None else "")


                elif tipo_widget == "treeview":
                    widget_objeto.acciones("eliminar_contenido")

                    if valor_informar is not None:
                        widget_objeto.acciones("actualizar_desde_df", df_datos = valor_informar)

                    if treeview_seleccionar_item is not None:
                        widget_objeto.acciones("seleccionar_item", lista_item_seleccionado = treeview_seleccionar_item)

                    
                if tipo_widget in ["scrolledtext", "entry", "entry_propio"]:
                    widget_objeto.config_atributos(**{"bloquear": bloquear})



###################################################################################################################################################
###################################################################################################################################################
# clase gui_config_id_pptx_nuevo_id_pptx
###################################################################################################################################################
###################################################################################################################################################

class gui_config_id_pptx_nuevo_id_pptx():

    def __init__(self, master
                , id_pptx_nuevo = None
                , id_pptx_nuevo_path = None
                , id_pptx_nuevo_num_slides = None
                , entorno_clase_gui_config_id_pptx = None
                , **kwargs_config_gui):


        #se inicializan atributosde la presente clase
        self.master = master
        self.clase_gui_actual_nombre = self.__class__.__name__

        self.id_pptx_nuevo = id_pptx_nuevo
        self.id_pptx_nuevo_path = id_pptx_nuevo_path
        self.id_pptx_nuevo_num_slides = id_pptx_nuevo_num_slides
        self.entorno_clase_gui_config_id_pptx = entorno_clase_gui_config_id_pptx
        
        self.kwargs_config_gui = kwargs_config_gui
        self.kwargs_gui_config_id_pptx_nuevo_id_pptx = {key: dicc for key, dicc in self.kwargs_config_gui[self.clase_gui_actual_nombre].items() if key != "dicc_config_root"}

        
        #se insertan los widgets dentro del frame_inicio y se almacenan en el diccionario dicc_gui_frame_widgets_objetos
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos = {}
        for frame_contenedor in self.kwargs_gui_config_id_pptx_nuevo_id_pptx.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_config_id_pptx_nuevo_id_pptx_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_config_id_pptx_nuevo_id_pptx.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_config_id_pptx_nuevo_id_pptx_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_config_id_pptx_nuevo_id_pptx_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_config_id_pptx_nuevo_id_pptx[frame_contenedor].items() if widget != "frame"}

            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_config_id_pptx_nuevo_id_pptx_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ", "").strip()

                widget_objeto = (mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget_ajust, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]
                                else
                                mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "scrolledtext_propio"
                                else
                                mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust == "entry_propio"
                                else
                                mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust == "treeview_propio"
                                else
                                mod_utils.frame_con_scrollbar(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "frame_con_scrollbar"
                                else None)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos.update({frame_contenedor_widget:
                                                                                                                {"widget_objeto": widget_objeto
                                                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                                                }
                                                                                        })
            

        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.widget_id_pptx = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_74"]["widget_objeto"]
        self.widget_id_pptx_path = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_77"]["widget_objeto"]
        self.widget_id_pptx_desc = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_79"]["widget_objeto"]
        self.widget_tiempo_apertura_max_id_pptx = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_81"]["widget_objeto"]
        self.widget_id_pptx_num_slides = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_83"]["widget_objeto"]

        self.strvar_widget_tiempo_apertura_max_id_pptx = self.dicc_gui_config_id_pptx_nuevo_id_pptx_frame_widgets_objetos["WIDGET_81"]["widget_variable_enlace"]



        #se actualiza el id pptx, su ubicacion y el numero de slides que contiene
        #aqui no se hace con el metodo def_gui_config_id_pptx_widgets_actualizar pq no se define en la clase gui_config_id_pptx_nuevo_id_pptx
        #la actualizacion de los datos se hace manualmente
        self.widget_id_pptx.config_atributos(**{"bloquear": False})
        self.widget_id_pptx_path.config_atributos(**{"bloquear": False})
        self.widget_id_pptx_num_slides.config_atributos(**{"bloquear": False})

        self.widget_id_pptx.modificaciones("agregar_solo_contenido_desde_string"
                                            , string_texto_informar = self.id_pptx_nuevo
                                            , height_scrolledtext = 1)
        
        self.widget_id_pptx_path.modificaciones("agregar_solo_contenido_desde_string"
                                                , string_texto_informar = self.id_pptx_nuevo_path
                                                , height_scrolledtext = 1)
        
        self.widget_id_pptx_num_slides.modificaciones("agregar_solo_contenido_desde_string"
                                                , string_texto_informar = self.id_pptx_nuevo_num_slides
                                                , height_scrolledtext = 1)

        self.widget_id_pptx.config_atributos(**{"bloquear": True})
        self.widget_id_pptx_path.config_atributos(**{"bloquear": True})
        self.widget_id_pptx_num_slides.config_atributos(**{"bloquear": True})



    def def_gui_config_id_pptx_nuevo_id_pptx_guardar(self):
        #rutina que permite traspasar los datos informados en la ventana que genera la presente clase a la ventana anterior
        #(generada con la clase gui_config_id_pptx) modificando asimismo el treeview id xls agregando el nuevo id xls
        #y seleccionandolo en pantalla, ademas de modificar la variable global global_dicc_datos_id_pptx

        id_pptx_nuevo_desc = self.widget_id_pptx_desc.texto_informado("todo")
        id_pptx_nuevo_tiempo_apertura_max = self.strvar_widget_tiempo_apertura_max_id_pptx.get()


        if len(id_pptx_nuevo_desc) == 0 or len(id_pptx_nuevo_tiempo_apertura_max) == 0:
            mensaje = "Todos los campos son obligatorios."
            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))
            
        else:

            #se actualiza en memoria en la variable global global_dicc_datos_id_pptx el id pptx creado
            mod_back_end.def_varios_gui_config_id_pptx("CREAR_NUEVO_ID_PPTX"
                                                        , id_pptx = self.id_pptx_nuevo
                                                        , id_pptx_path = self.id_pptx_nuevo_path
                                                        , id_pptx_desc = id_pptx_nuevo_desc
                                                        , id_pptx_tiempo_espera_max_apertura = id_pptx_nuevo_tiempo_apertura_max
                                                        , id_pptx_numero_total_slides = self.id_pptx_nuevo_num_slides
                                                        )
            
            #se actualiza el atributo widget_slide_pptx_lista_opciones de la clase gui_config_id_pptx
            #(se usa en lista_dicc_widgets, ver mas abajo)
            self.entorno_clase_gui_config_id_pptx.widget_slide_pptx_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_SLIDES_PPTX", id_pptx = self.id_pptx_nuevo)


            #se borran y se bloquean los widgets de GUI de config
            #se usa el metodo de la clase def_gui_config_id_pptx_widgets_actualizar que se recupera al haber pasado en la presente clase
            #por parametro el entorno entorno entorno_clase_gui_config_id_pptx
            self.entorno_clase_gui_config_id_pptx.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS")


            #se actualizan los widgets en la ventana GUI de config de los id pptx
            #se usa el mismo metodo que el bloque anterior
            lista_dicc_widgets = [
                                    {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_pptx
                                        , "variable_enlace": None
                                        , "height": 1
                                        , "bloquear": True
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": None
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": self.id_pptx_nuevo
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_pptx_path
                                        , "variable_enlace": None
                                        , "height": 1
                                        , "bloquear": True
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": None
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": self.id_pptx_nuevo_path
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_pptx_desc
                                        , "variable_enlace": None
                                        , "height": self.entorno_clase_gui_config_id_pptx.widget_id_pptx_desc_height
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": id_pptx_nuevo_desc
                                        }

                                    , {"tipo_widget": "entry_propio"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_tiempo_apertura_max_id_pptx
                                        , "variable_enlace": self.entorno_clase_gui_config_id_pptx.strvar_widget_tiempo_apertura_max_id_pptx
                                        , "height": None
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": id_pptx_nuevo_tiempo_apertura_max
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_pptx_num_slides
                                        , "variable_enlace": None
                                        , "height": None
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": self.id_pptx_nuevo_num_slides
                                        }

                                    , {"tipo_widget": "combobox"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_slide_pptx
                                        , "variable_enlace": self.entorno_clase_gui_config_id_pptx.strvar_widget_slide_pptx
                                        , "height": None
                                        , "bloquear": True
                                        , "combobox_lista_opciones": self.entorno_clase_gui_config_id_pptx.widget_slide_pptx_lista_opciones
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": None
                                        }
                                ]
            
            self.entorno_clase_gui_config_id_pptx.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                                            , lista_dicc_widgets = lista_dicc_widgets)


            
            #se agrega una opcion en la lista de opciones del combobox de seleccion de id pptx y se selecciona en el combobox el nuevo valor
            #se actualiza tambien el atributo propio self.id_pptx_selecc
            seleccion_id_pptx_lista_opciones = self.entorno_clase_gui_config_id_pptx.widget_seleccion_id_pptx_lista_opciones
            seleccion_id_pptx_lista_opciones.append(self.id_pptx_nuevo)
            seleccion_id_pptx_lista_opciones = sorted(seleccion_id_pptx_lista_opciones)#se reordena por si el nuevo id pptx reelnea un hueco de un id pptx eliminado

            self.entorno_clase_gui_config_id_pptx.widget_seleccion_id_pptx_lista_opciones = seleccion_id_pptx_lista_opciones
            self.entorno_clase_gui_config_id_pptx.widget_seleccion_id_pptx.config_atributos(**{"combobox_lista_opciones": seleccion_id_pptx_lista_opciones})

            self.entorno_clase_gui_config_id_pptx.strvar_widget_seleccion_id_pptx.set(self.id_pptx_nuevo)


            #se actualizan atributos de la clase gui_config_id_pptx
            self.entorno_clase_gui_config_id_pptx.id_pptx_selecc = self.id_pptx_nuevo
            self.entorno_clase_gui_config_id_pptx.id_xls_selecc = None
            self.entorno_clase_gui_config_id_pptx.id_xls_selecc_antes_click_item = None

            self.entorno_clase_gui_config_id_pptx.widget_hojas_xls_lista_opciones = None



            #se cierra la ventana generadapor la presente clase
            self.master.widget_objeto.destroy()


            #se genera un messagebox que informa de los pasos a seguir una vez creado el nuevo id pptx
            mensaje = "Una vez creado el nuevo id pptx, tienes que pulsar el botón de agregar un nuevo id xls.\n\n"
            mod_utils.messagebox_propio(self.entorno_clase_gui_config_id_pptx.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))
            



###################################################################################################################################################
###################################################################################################################################################
# clase gui_config_id_pptx_nuevo_id_xls
###################################################################################################################################################
###################################################################################################################################################

class gui_config_id_pptx_nuevo_id_xls():

    def __init__(self, master
                , id_pptx = None
                , id_xls_nuevo = None
                , id_xls_nuevo_path = None
                , entorno_clase_gui_config_id_pptx = None
                , **kwargs_config_gui):


        #se inicializan atributosde la presente clase
        self.master = master
        self.clase_gui_actual_nombre = self.__class__.__name__

        self.id_pptx = id_pptx
        self.id_xls_nuevo = id_xls_nuevo
        self.id_xls_nuevo_path = id_xls_nuevo_path
        self.entorno_clase_gui_config_id_pptx = entorno_clase_gui_config_id_pptx
        
        self.kwargs_config_gui = kwargs_config_gui
        self.kwargs_gui_config_id_pptx_nuevo_id_xls = {key: dicc for key, dicc in self.kwargs_config_gui[self.clase_gui_actual_nombre].items() if key != "dicc_config_root"}

        
        #se insertan los widgets dentro del frame_inicio y se almacenan en el diccionario dicc_gui_frame_widgets_objetos
        #para posterior uso en las rutinas propias de la presente clase
        self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos = {}
        for frame_contenedor in self.kwargs_gui_config_id_pptx_nuevo_id_xls.keys():

            #se crea el frame correspondiente dentro de la GUI
            #(se recuperan el diccionario de parametros creando lista de diccionarios y recuperando el 1er item, es lista de 1 solo item)         
            kwargs_gui_config_id_pptx_nuevo_id_xls_frame_iter = [dicc["frame"] for frame, dicc in self.kwargs_gui_config_id_pptx_nuevo_id_xls.items() if frame == frame_contenedor][0]

            self.objeto_frame_contenedor = mod_utils.gui_tkinter_widgets(self.master, tipo_widget_param = "frame", **kwargs_gui_config_id_pptx_nuevo_id_xls_frame_iter)


            #se crea diccionario con los parametros de los widgets a incluir en el frame de la iteracion
            #y mediante bucle sobre las keys de este diccionario se crean los widgets dinamicamente
            kwargs_gui_config_id_pptx_nuevo_id_xls_frame_iter_widgets = {widget: kwargs_widget for widget, kwargs_widget in self.kwargs_gui_config_id_pptx_nuevo_id_xls[frame_contenedor].items() if widget != "frame"}

            for frame_contenedor_widget, frame_contenedor_kwargs_widget in kwargs_gui_config_id_pptx_nuevo_id_xls_frame_iter_widgets.items():

                tipo_widget = frame_contenedor_kwargs_widget["tipo_widget"].lower().strip()
                kwargs_config = frame_contenedor_kwargs_widget["kwargs_config"]


                #se crean los widgets
                tipo_widget_ajust = tipo_widget.lower().replace(" ", "").strip()

                widget_objeto = (mod_utils.gui_tkinter_widgets(self.objeto_frame_contenedor.widget_objeto, tipo_widget_param = tipo_widget_ajust, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust in ["label", "combobox", "entry", "button", "listbox"]
                                else
                                mod_utils.scrolledtext_propio(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "scrolledtext_propio"
                                else
                                mod_utils.entry_propio(self.objeto_frame_contenedor.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust == "entry_propio"
                                else
                                mod_utils.treeview_propio(self.objeto_frame_contenedor.widget_objeto, entorno_donde_se_llama_la_clase = self, **kwargs_config)
                                if tipo_widget_ajust == "treeview_propio"
                                else
                                mod_utils.frame_con_scrollbar(self.objeto_frame_contenedor.widget_objeto, **kwargs_config)
                                if tipo_widget_ajust == "frame_con_scrollbar"
                                else None)


                #se almacena el widget (objeto) en el diccionario dicc_widgets_frame_contenedor junto con su stringvar (si lo tiene)
                self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos.update({frame_contenedor_widget:
                                                                                                                {"widget_objeto": widget_objeto
                                                                                                                , "widget_variable_enlace": widget_objeto.variable_enlace
                                                                                                                }
                                                                                        })


        #se recuperan los widgets_objetos que se usan en distintas rutinas de la la presente clase
        self.widget_id_xls = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_85"]["widget_objeto"]
        self.widget_id_xls_path = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_88"]["widget_objeto"]
        self.widget_id_xls_desc = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_90"]["widget_objeto"]
        self.widget_tiempo_apertura_max_id_xls = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_92"]["widget_objeto"]
        self.widget_actualizar_vinculos_otros_excel = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_94"]["widget_objeto"]

        self.strvar_widget_tiempo_apertura_max_id_xls = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_92"]["widget_variable_enlace"]
        self.strvar_widget_actualizar_vinculos_otros_excel = self.dicc_gui_config_id_pptx_nuevo_id_xls_frame_widgets_objetos["WIDGET_94"]["widget_variable_enlace"]


        #se actualiza el id xls y su ubicacion
        #se establece la actualizacion e otros vinculos excel a 'No' por defecto
        #aqui no se hace con el metodo def_gui_config_id_pptx_widgets_actualizar pq no se define en la clase gui_config_id_pptx_nuevo_id_xls
        #la actualizacion de los datos se hace manualmente
        self.widget_id_xls.config_atributos(**{"bloquear": False})
        self.widget_id_xls_path.config_atributos(**{"bloquear": False})

        self.widget_id_xls.modificaciones("agregar_solo_contenido_desde_string"
                                            , string_texto_informar = self.id_xls_nuevo
                                            , height_scrolledtext = 1)
        
        self.widget_id_xls_path.modificaciones("agregar_solo_contenido_desde_string"
                                            , string_texto_informar = self.id_xls_nuevo_path
                                            , height_scrolledtext = 1)

        self.widget_id_xls.config_atributos(**{"bloquear": True})
        self.widget_id_xls_path.config_atributos(**{"bloquear": True})

        self.strvar_widget_actualizar_vinculos_otros_excel.set("No")
    


    def def_gui_config_id_pptx_nuevo_id_xls_guardar(self):
        #rutina que permite traspasar los datos informados en la ventana que genera la presente clase a la ventana anterior
        #(generada con la clase gui_config_id_pptx) modificando asimismo el treeview id xls agregando el nuevo id xls
        #y seleccionandolo en pantalla, ademas de modificar la variable global global_dicc_datos_id_pptx

        id_xls_nuevo_desc = self.widget_id_xls_desc.texto_informado("todo")
        id_xls_nuevo_tiempo_apertura_max = self.strvar_widget_tiempo_apertura_max_id_xls.get()
        id_xls_nuevo_actualizar_vinculos_otros_excel = self.strvar_widget_actualizar_vinculos_otros_excel.get()

        
        if len(id_xls_nuevo_desc) == 0 or len(id_xls_nuevo_tiempo_apertura_max) == 0:
            mensaje = "Todos los campos son obligatorios."
            mod_utils.messagebox_propio(self.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showerror", "mensaje": mensaje}))
            

        else:

            #se actualiza en memoria en la variable global global_dicc_datos_id_pptx el id pptx creado
            #(funciona aqui tambien como funcion y deveulve el df para infromar el treeview de id xls)
            (lista_datos_items_seleccionados
            , df_actualizado_treeview) = mod_back_end.def_varios_gui_config_id_pptx("CREAR_NUEVO_ID_XLS"
                                                                                , id_pptx = self.id_pptx
                                                                                , id_xls = self.id_xls_nuevo
                                                                                , id_xls_path = self.id_xls_nuevo_path
                                                                                , id_xls_desc = id_xls_nuevo_desc
                                                                                , id_xls_tiempo_espera_max_apertura = id_xls_nuevo_tiempo_apertura_max
                                                                                , id_xls_actualizar_vinculos_otros_xls = id_xls_nuevo_actualizar_vinculos_otros_excel)

            #se borran y se bloquean los widgets de GUI de config (los que afectan a los id xls y los rangos de celdas)
            #se usa el metodo de la clase def_gui_config_id_pptx_widgets_actualizar que se recupera al haber pasado en la presente clase
            #por parametro el entorno entorno entorno_clase_gui_config_id_pptx
            self.entorno_clase_gui_config_id_pptx.def_gui_config_id_pptx_widgets_actualizar("BORRAR_CONTENIDO_Y_BLOQUEAR_WIDGETS_ACCIONES_ID_XLS")



            #se actualizan atributos de la clase (se usan en lista_dicc_widgets, bloque siguiente)
            #se actualizan atributos de la clase gui_config_id_pptx
            self.entorno_clase_gui_config_id_pptx.id_xls_selecc = self.id_xls_nuevo
            self.entorno_clase_gui_config_id_pptx.id_xls_selecc_antes_click_item = None


            if self.entorno_clase_gui_config_id_pptx.id_pptx_selecc is not None and self.entorno_clase_gui_config_id_pptx.id_xls_selecc is not None:
                self.entorno_clase_gui_config_id_pptx.widget_hojas_xls_lista_opciones = mod_back_end.def_varios_gui_config_id_pptx("COMBOBOX_LISTA_OPCIONES_HOJAS_XLS"
                                                                                                                                    , id_pptx = self.entorno_clase_gui_config_id_pptx.id_pptx_selecc
                                                                                                                                    , id_xls = self.entorno_clase_gui_config_id_pptx.id_xls_selecc)



            #se actualizan los widgets en la ventana GUI de config de los id pptx
            #se usa el mismo metodo que el bloque anterior
            lista_dicc_widgets = [
                                    {"tipo_widget": "treeview"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_treeview_id_xls
                                        , "variable_enlace": None
                                        , "height": None
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": None
                                        , "treeview_seleccionar_item": lista_datos_items_seleccionados
                                        , "valor_informar": df_actualizado_treeview
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_xls
                                        , "variable_enlace": None
                                        , "height": 1
                                        , "bloquear": True
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": None
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": self.id_xls_nuevo
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_xls_path
                                        , "variable_enlace": None
                                        , "height": 1
                                        , "bloquear": True
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": None
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": self.id_xls_nuevo_path
                                        }

                                    , {"tipo_widget": "scrolledtext"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_id_xls_desc
                                        , "variable_enlace": None
                                        , "height": self.entorno_clase_gui_config_id_pptx.widget_id_xls_desc_height
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": id_xls_nuevo_desc
                                        }

                                    , {"tipo_widget": "entry_propio"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_tiempo_apertura_max_id_xls
                                        , "variable_enlace": self.entorno_clase_gui_config_id_pptx.strvar_widget_tiempo_apertura_max_id_xls
                                        , "height": None
                                        , "bloquear": False
                                        , "combobox_lista_opciones": None
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": id_xls_nuevo_tiempo_apertura_max
                                        }

                                    , {"tipo_widget": "combobox"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_actualizar_vinculos_otros_excel
                                        , "variable_enlace": self.entorno_clase_gui_config_id_pptx.strvar_widget_actualizar_vinculos_otros_excel
                                        , "height": None
                                        , "bloquear": False
                                        , "combobox_lista_opciones": mod_back_end.lista_opciones_id_xls_combobox_actualizar_vinculos
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": id_xls_nuevo_actualizar_vinculos_otros_excel
                                        }

                                    , {"tipo_widget": "combobox"
                                        , "widget_objeto": self.entorno_clase_gui_config_id_pptx.widget_hojas_xls
                                        , "variable_enlace": self.entorno_clase_gui_config_id_pptx.strvar_widget_hojas_xls
                                        , "height": None
                                        , "bloquear": True
                                        , "combobox_lista_opciones": self.entorno_clase_gui_config_id_pptx.widget_hojas_xls_lista_opciones
                                        , "combobox_opciones_editables": False
                                        , "treeview_seleccionar_item": None
                                        , "valor_informar": None
                                        }
                                ]
            

            self.entorno_clase_gui_config_id_pptx.def_gui_config_id_pptx_widgets_actualizar("INFORMAR_Y_DESBLOQUEAR_WIDGETS_DESDE_LISTA_DICC"
                                                                                            , lista_dicc_widgets = lista_dicc_widgets)




            #se informa el atributo datos_items del objeto treeview id xls
            self.entorno_clase_gui_config_id_pptx.widget_treeview_id_xls.datos_items["lista_datos_items_seleccionados"] = lista_datos_items_seleccionados



            #se cierra la ventana generada por la presente clase
            self.master.widget_objeto.destroy()


            #se genera un messagebox que informa de los pasos a seguir una vez creado el nuevo id xls y seleccionado en el treeview actualizado
            mensaje = "Una vez creado el nuevo id xls asociado al id pptx, tienes ahora que configurar los rangos de celdas."
            mod_utils.messagebox_propio(self.entorno_clase_gui_config_id_pptx.master.widget_objeto, **mod_back_end.def_varios("KWARGS_PARA_MESSAGEBOX_PROPIO", kwargs_messagebox = {"tipo_messagebox": "showinfo", "mensaje": mensaje}))




###################################################################################################################################################
###################################################################################################################################################
# clase gui_screenshot_muestra
###################################################################################################################################################
###################################################################################################################################################

class gui_screenshot_muestra():

    def __init__(self, master
                , path_screenshot_png = None
                , **kwargs_config_gui):


        #se inicializan atributosde la presente clase
        self.master = master
        self.clase_gui_actual_nombre = self.__class__.__name__

        self.path_screenshot_png = path_screenshot_png
        self.kwargs_config_gui = kwargs_config_gui


        #se extraen los kwargs para crear los frame y el label
        self.kwargs_gui_screenshot_muestra_frame_integrado_root = self.kwargs_config_gui[self.clase_gui_actual_nombre]["frame_integrado_root"]["frame_inicio"]["kwargs_config"]
        self.kwargs_gui_screenshot_muestra_frame_label = self.kwargs_config_gui[self.clase_gui_actual_nombre]["frame_integrado_root"]["frame_label"]["frame"]
        self.kwargs_gui_screenshot_muestra_frame_label_screenshot_label = self.kwargs_config_gui[self.clase_gui_actual_nombre]["frame_integrado_root"]["frame_label"]["WIDGET_95"]["kwargs_config"]

        self.frame_inicio = mod_utils.frame_con_scrollbar(self.master.widget_objeto, **self.kwargs_gui_screenshot_muestra_frame_integrado_root)
        self.frame_label = mod_utils.gui_tkinter_widgets(self.frame_inicio.widget_objeto, tipo_widget_param = "frame", **self.kwargs_gui_screenshot_muestra_frame_label)
        self.label_screenshot = mod_utils.gui_tkinter_widgets(self.frame_label.widget_objeto, tipo_widget_param = "label", **self.kwargs_gui_screenshot_muestra_frame_label_screenshot_label)


        #se inserta en png del screenshot de muestra
        kwargs_config_png = {"dicc_imagen":
                                            {"png_imagen": self.path_screenshot_png
                                            , "tupla_pixeles_imagen": mod_back_end.tupla_resize_pixeles_screenshot_muestra
                                            }
                            }
        
        self.label_screenshot.config_atributos(**kwargs_config_png)





##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################
#se inicia el app
##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################
##########################################################################################################################################################

if __name__ == "__main__":

    #se crea el diccionario kwargs dicc_kwargs_gui que sirve para colocar todos los widgets de la gui
    dicc_kwargs_config_gui = {
                            ###############################################################################################################################################################################################
                            # ventana gui_ventana_inicio
                            ###############################################################################################################################################################################################
                            "gui_ventana_inicio":
                                            {"dicc_config_root":
                                                        {"title": mod_back_end.nombre_app
                                                        , "iconbitmap": mod_back_end.template_ico_app_tapar_pluma_tkinter
                                                        , "tupla_geometry": (520, 530)
                                                        , "resizable": (0, 0)
                                                        }

                                            , "frame_varios":
                                                            {"frame":
                                                                    {"width": 480
                                                                    , "height": 50
                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 0}
                                                                    }

                                                            , "WIDGET_01":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label del warning de Poppler no localizado" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"font": ("Calibri", 12, "bold")
                                                                                            , "width": 40
                                                                                            , "fg": "red"
                                                                                            , "alineacion": "left"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 10}
                                                                                            }
                                                                        }

                                                            , "WIDGET_02":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label de la guia de usuario" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "guia usuario" 
                                                                                            , "font": ("Calibri", 12, "bold")
                                                                                            , "width": 12
                                                                                            , "fg": "black"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 330, "coord_y": 10}
                                                                                            }
                                                                        }

                                                            , "WIDGET_03":
                                                                        {"tipo_widget": "button"
                                                                        , "desc_tipo_widget": "boton asociado a la descarga de la guia de usuario" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"width": 40
                                                                                            , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guia_usuario, "tupla_imagen_resize": (23, 23)}
                                                                                            , "controltiptext": "Descarga en la ruta que indiques la guia de usuario en formato PDF"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 430, "coord_y": 10}
                                                                                            , "dicc_rutina":
                                                                                                            {"rutina": "def_gui_ventana_inicio_threads"

                                                                                                            #parametros_args se asocia al parametro args estatico opcion_gui_guia_usuario (mod_back_end)
                                                                                                            , "parametros_args": (lambda widget: mod_back_end.opcion_gui_guia_usuario,)
                                                                                                            }
                                                                                            }
                                                                        }
                                                            }

                                            , "frame_sistema":
                                                            {"frame":
                                                                    {"width": 480
                                                                    , "height": 110
                                                                    , "bd": 2
                                                                    , "relief": "solid"
                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 50}
                                                                    , "bg": "#ACADB1"
                                                                    }

                                                            , "WIDGET_04":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label titulo del frame 'frame_sistema'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "SISTEMA" 
                                                                                            , "font": ("Calibri", 13, "bold")
                                                                                            , "width": 13
                                                                                            , "bd": 1
                                                                                            , "relief": "solid"
                                                                                            , "bg": "#1F40AD"
                                                                                            , "fg": "white"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                            }
                                                                        }

                                                            , "WIDGET_05":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label combobox opciones del frame 'frame_sistema'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "opciones" 
                                                                                            , "font": ("Calibri", 12, "bold")
                                                                                            , "width": 15
                                                                                            , "bg": "#ACADB1"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 35}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }

                                                            , "WIDGET_06": #el nombre de la key aqui es importante se hace referencia en este mismo diccionario (WIDGET_07)
                                                                        {"tipo_widget": "combobox"
                                                                        , "desc_tipo_widget": "combobox para seleccionar las opciones relacionadas con el sistema sqlite" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"font": ("Calibri", 10)
                                                                                            , "width": 22
                                                                                            , "alineacion": "left"
                                                                                            , "controltiptext": "Permite seleccionar acciones a realizar con el sistema de configuración del app en sqlite"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 40}
                                                                                            , "combobox_lista_opciones": mod_back_end.def_varios_gui_ventana_inicio("LISTA_OPCIONES_COMBOBOX_GUI_SISTEMA")
                                                                                            }
                                                                        }

                                                            , "WIDGET_07":
                                                                        {"tipo_widget": "button"
                                                                        , "desc_tipo_widget": "boton asociado al combobox de seleccion de opciones de sistema" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"width": 40
                                                                                            , "dicc_imagen": {"png_imagen": mod_back_end.template_img_config, "tupla_imagen_resize": (23, 23)}
                                                                                            , "controltiptext": "Ejecuta varias acciones relacionadas con el sistema de configuración del app en sqlite"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 340, "coord_y": 36}
                                                                                            , "dicc_rutina":
                                                                                                            {"rutina": "def_gui_ventana_inicio_threads"
                                                                                                            
                                                                                                            #parametros_args es dinamico y se asocia al valor que toma el combobox WIDGET_06 dentro del frame frame_sistema
                                                                                                            , "parametros_args": (lambda widget: 
                                                                                                                widget.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_06"]["widget_objeto"].widget_objeto.get(),)
                                                                                                                
                                                                                                            }
                                                                                            }
                                                                        }

                                                            , "WIDGET_08":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label ruta sistema del frame 'frame_sistema'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "ubicación sistema" 
                                                                                            , "font": ("Calibri", 12, "bold")
                                                                                            , "width": 15
                                                                                            , "bg": "#ACADB1"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 65}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }

                                                            , "WIDGET_09": #el nombre de la key aqui es importante se hace referencia en la clase gui_app (mas arriba en este modulo)
                                                                        {"tipo_widget": "scrolledtext_propio"
                                                                        , "desc_tipo_widget": "scrolledtext que almacena la ruta sistema" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"font": ("Calibri", 10, "bold")
                                                                                            , "width": 43
                                                                                            , "height": 1
                                                                                            , "bloquear": True
                                                                                            , "wrap": tk.NONE
                                                                                            , "bg": "#B7C3F5"
                                                                                            , "controltiptext": "Ubicación del sistema sqlite"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 150, "coord_y": 70}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }
                                                            }

                                            , "frame_pptx":
                                                            {"frame":
                                                                    {"width": 480
                                                                    , "height": 330
                                                                    , "bd": 2
                                                                    , "relief": "solid"
                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 170}
                                                                    , "bg": "#F38E09"
                                                                    }

                                                            , "WIDGET_10":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label titulo del frame 'frame_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "POWERPOINT" 
                                                                                            , "font": ("Calibri", 13, "bold")
                                                                                            , "width": 13
                                                                                            , "bd": 1
                                                                                            , "relief": "solid"
                                                                                            , "bg": "#1F40AD"
                                                                                            , "fg": "white"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                            }
                                                                        }


                                                            , "WIDGET_11":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label boton ver config ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "configurar" 
                                                                                            , "font": ("Calibri", 12, "bold")
                                                                                            , "width": 13
                                                                                            , "bg": "#F38E09"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 325, "coord_y": 15}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }

                                                            , "WIDGET_12":
                                                                        {"tipo_widget": "button"
                                                                        , "desc_tipo_widget": "boton ver config ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"width": 35
                                                                                            , "dicc_imagen": {"png_imagen": mod_back_end.template_img_boton_ver, "tupla_imagen_resize": (28, 20)}
                                                                                            , "controltiptext": "Abre una nueva ventana para poder configurar los id pptx.\nNo es necesario seleccionar previamente un id pptx en el treeview"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 417, "coord_y": 17}
                                                                                            , "dicc_rutina":
                                                                                                            {"rutina": "def_gui_ventana_inicio_threads"
                                                                                                            
                                                                                                            #parametros_args se asocia al parametro args estatico opcion_gui_guia_usuario (mod_back_end)
                                                                                                            , "parametros_args": (lambda widget: mod_back_end.opcion_gui_config_id_pptx,)
                                                                                                                
                                                                                                            }
                                                                                            }
                                                                        }

                                                            , "WIDGET_13": #el nombre de la key aqui es importante se hace referencia en la clase gui_app (mas arriba en este modulo)
                                                                        {"tipo_widget": "treeview_propio"
                                                                        , "desc_tipo_widget": "treeview de las configuraciones pptx dentro del frame 'frame_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 55}
                                                                                            , "dicc_treeview": {"seleccion_item": "simple"
                                                                                                                , "height": 5
                                                                                                                ,"columnas_df": ["ID_PPTX", "DESC_ID_PPTX"]
                                                                                                                #importante de que los campos aparezcan el diccionario dicc_tabla_config_sistema["DICC_CAMPOS"]
                                                                                                                #del modulo back-end

                                                                                                                , "columnas_treeview": ["id pptx", "descripción"]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                , "width_columnas_treeview": [80, 370]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                }
                                                                                            , "dicc_rutina_click_item": {"rutina": "def_gui_ventana_inicio_treeview_id_pptx_click_item"}
                                                                                            }
                                                                        }

                                                            , "WIDGET_14":
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label combobox ACCION en el 'frame_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": "acciones" 
                                                                                            , "font": ("Calibri", 12, "bold")
                                                                                            , "width": 13
                                                                                            , "bg": "#F38E09"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 195}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }

                                                            , "WIDGET_15": #el nombre de la key aqui es importante se hace referencia en este mismo diccionario (WIDGET_16)
                                                                        {"tipo_widget": "combobox"
                                                                        , "desc_tipo_widget": "combobox para seleccionar las opciones ACCION en el 'frame_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"font": ("Calibri", 10)
                                                                                            , "width": 20
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 80, "coord_y": 197}
                                                                                            , "alineacion": "left"
                                                                                            , "controltiptext": "Permite seleccionar acciones de configuración o ejecución de screenshots en presentaciones pptx"
                                                                                            , "combobox_lista_opciones": mod_back_end.def_varios_gui_ventana_inicio("LISTA_OPCIONES_COMBOBOX_GUI_ACCIONES_PPTX")

                                                                                            , "lista_dicc_rutina_aplicar_eventos_widget":[{"tipo_bind": "<<ComboboxSelected>>"
                                                                                                                                            , "rutina": "def_gui_ventana_inicio_combobox_accion_pptx"
                                                                                                                                            }]
                                                                                            }
                                                                        }

                                                            , "WIDGET_16":
                                                                        {"tipo_widget": "button"
                                                                        , "desc_tipo_widget": "boton asociado al combobox ACCION" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"width": 40
                                                                                            , "dicc_imagen": {"png_imagen": mod_back_end.template_img_accion_pptx, "tupla_imagen_resize": (33, 23)}
                                                                                            , "controltiptext": "Ejecuta la acción pptx seleccionada en el combobox"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 260, "coord_y": 195}
                                                                                            , "dicc_rutina":
                                                                                                            {"rutina": "def_gui_ventana_inicio_threads"
                                                                                                            
                                                                                                            #parametros_args es dinamico y se asocia al valor que toma el combobox WIDGET_15 dentro del frame frame_pptx
                                                                                                            , "parametros_args": (lambda widget: 
                                                                                                                widget.dicc_gui_ventana_inicio_frame_widgets_objetos["WIDGET_15"]["widget_objeto"].widget_objeto.get(),)
                                                                                                                
                                                                                                            }
                                                                                            }
                                                                        }

                                                            , "WIDGET_17":
                                                                        {"tipo_widget": "scrolledtext_propio"
                                                                        , "desc_tipo_widget": "scrolledtext que almacena la descripcion correspondiente a la opcion del combobox ACCION" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"font": ("Calibri", 10, "bold")
                                                                                            , "width": 63
                                                                                            , "height": 5
                                                                                            , "bloquear": True
                                                                                            , "wrap": tk.WORD
                                                                                            , "bg": "#B7C3F5"
                                                                                            , "fg": "black"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 230}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }
                                                            }

                                            , "frame_resolucion_pantalla":
                                                            {"frame":
                                                                    {"width": 480
                                                                    , "height": 60
                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 20, "coord_y": 500}
                                                                    }

                                                            , "WIDGET_18": #el nombre de la key aqui es importante se hace referencia mas abajo despues de declarar este diccionario
                                                                        {"tipo_widget": "label"
                                                                        , "desc_tipo_widget": "label resolucion pantalla'" #key informativa (no se usa en el resto del codigo del app)
                                                                        , "kwargs_config":
                                                                                            {"text": None #se informa al inicializar la clase gui_ventana_inicio
                                                                                            , "font": ("Calibri", 11, "bold")
                                                                                            , "width": 150
                                                                                            , "fg": "black"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 5}
                                                                                            , "alineacion": "left"
                                                                                            }
                                                                        }
                                                            }
                                        }
                            ###############################################################################################################################################################################################
                            # ventana gui_config_id_pptx
                            ###############################################################################################################################################################################################
                            , "gui_config_id_pptx":
                                                    {"dicc_config_root":
                                                                        {"title": "CONFIGURACIÓN PPTX"
                                                                        , "iconbitmap": mod_back_end.template_ico_app_tapar_pluma_tkinter
                                                                        , "tupla_geometry": (650, 740)
                                                                        , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                        , "mantener_nueva_ventana_encima_otras": True
                                                                        , "resizable": (0, 0)
                                                                        }

                                                    , "frame_integrado_root":
                                                                    {"frame_inicio":
                                                                                    {"kwargs_config":
                                                                                                    {"dicc_frame_scrollbar": {"width_visible": 620
                                                                                                                                , "width_total": 620
                                                                                                                                , "height_visible": 740
                                                                                                                                , "height_total": 870
                                                                                                                                , "tupla_coord_place": (0, 0)
                                                                                                                                , "velocidad_scrolling": 1
                                                                                                                                }
                                                                                                    }
                                                                                    }

                                                                    , "frame_seleccion_y_guardar":
                                                                                    {"frame":
                                                                                            {"width": 620
                                                                                            , "height": 50
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                            }

                                                                                    , "WIDGET_19":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label seleccion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "selección id pptx"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "width": 40
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_20":
                                                                                                {"tipo_widget": "combobox"
                                                                                                , "desc_tipo_widget": "combobox seleccion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10)
                                                                                                                    , "width": 10
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 140, "coord_y": 12}
                                                                                                                    , "controltiptext": "Permite seleccionar el id pptx para poder configurarlo.\nUna vez seleccionado hay que pulvar el botón 'VER'"
                                                                                                                    , "alineacion": "center"
                                                                                                                    , "combobox_lista_opciones": []

                                                                                                                    }
                                                                                                }


                                                                                    , "WIDGET_21":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton seleccion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "bg": "#EEECE8"
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_boton_ver, "tupla_imagen_resize": (20, 20)}
                                                                                                                    , "controltiptext": "Muestra las configuraciones del id pptx seleccionado"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 250, "coord_y": 9}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("SELECCION_ID_PPTX",)   
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_22":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label guardar configuracion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "guardar"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "width": 40
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 500, "coord_y": 10}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_23":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "guardar configuracion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "bg": "#EEECE8"
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guardar, "tupla_imagen_resize": (20, 20)}
                                                                                                                    , "controltiptext": "Guarda las configuraciones del id pptx seleccionado en el sistema sqlite"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 565, "coord_y": 9}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("GUARDAR_CONFIGURACIONES_EN_SISTEMA_SQLITE",)   
                                                                                                                                    }
                                                                                                                    }
                                                                                                }
                                                                                    }

                                                                    , "frame_pptx_destino":
                                                                                    {"frame":
                                                                                            {"width": 600
                                                                                            , "height": 190
                                                                                            , "bd": 2
                                                                                            , "relief": "solid"
                                                                                            , "bg": "#71C657"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 45}
                                                                                            }

                                                                                    , "WIDGET_24":
                                                                                                    {"tipo_widget": "label"
                                                                                                    , "desc_tipo_widget": "label titulo del frame 'frame_id_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                                                    , "kwargs_config":
                                                                                                                        {"text": "PPTX DESTINO" 
                                                                                                                        , "font": ("Calibri", 13, "bold")
                                                                                                                        , "width": 13
                                                                                                                        , "bd": 1
                                                                                                                        , "relief": "solid"
                                                                                                                        , "bg": "#1F40AD"
                                                                                                                        , "fg": "white"
                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                                        }
                                                                                                    }

                                                                                    , "WIDGET_25":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "id pptx"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_26":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext con el ID PPTX (bloqueado)" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 20
                                                                                                                    , "height": 1
                                                                                                                    , "bloquear": True
                                                                                                                    , "wrap": tk.NONE
                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                    , "fg": "black"
                                                                                                                    , "controltiptext": "Identificador único (interno al app) del fichero pptx de destino"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 42}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_27":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton add ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guardar_ruta, "tupla_imagen_resize": (21, 18)}
                                                                                                                    , "controltiptext": "Abre una ventana de dialogo para seleccionar la ruta del nuevo fichero pptx de destino.\nAbre una nueva ventana para obligarte a informar los datos generales asociados al nuevo id pptx."
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 270, "coord_y": 40}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ABRIR_VENTANA_CREACION_ID_PPTX_NUEVO",)   
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_28":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton clear excel asociado al ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_eliminar_ruta, "tupla_imagen_resize": (21, 18)}
                                                                                                                    , "controltiptext": "Elimina un fichero excel de origen asociado al id pptx"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 330, "coord_y": 40}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ELIMINAR_ID_PPTX",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_29":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton update pptx ya configurado asociado al ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_update_ruta, "tupla_imagen_resize": (21, 18)}
                                                                                                                    , "controltiptext": "Actualiza la ubicación del fichero pptx de destino asociado al id pptx"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 390, "coord_y": 40}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("UPDATE_PATH_ID_PPTX",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_30":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label ubicacion pptx destino" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "ubicación"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 70}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_31":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext ubicacion pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 60
                                                                                                                    , "height": 1
                                                                                                                    , "wrap": tk.NONE
                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                    , "fg": "black"
                                                                                                                    , "bloquear": True
                                                                                                                    , "controltiptext": "Ubicación del fichero pptx asociado al id pptx"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 72}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }


                                                                                    , "WIDGET_32":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton abrir pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 30
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_abrir_fichero, "tupla_pixeles_imagen": (18, 15)}
                                                                                                                    , "controltiptext": "Abre el fichero pptx por si se desea ver en que slides hay que configurar capturas de rangos excel"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 550, "coord_y": 72}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ABRIR_PPTX",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_33":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label descripcion ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "descripción"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 100}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_34":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext con la descripcion del ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 67
                                                                                                                    , "height": 2
                                                                                                                    , "wrap": tk.WORD
                                                                                                                    , "fg": "black"
                                                                                                                    , "controltiptext": "Descripción del fichero pptx asociado al id pptx"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 102}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_35":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label pptx destino - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "tiempo max"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 145}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_36":
                                                                                                {"tipo_widget": "entry_propio"
                                                                                                , "desc_tipo_widget": "entry pptx destino - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 10
                                                                                                                    , "fg": "black"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 147}
                                                                                                                    , "controltiptext": "Tiempo de espera máximo que se esperara para que se abra el fichero pptx de destino por completo"
                                                                                                                    , "alineacion": "center"
                                                                                                                    , "dicc_entry":
                                                                                                                                {"formato_validacion": "entero_positivo"
                                                                                                                                , "titulo_messagebox_warning": "TIEMPO ESPERA MÁXIMO APERTURA PPTX"
                                                                                                                                }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_37":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label pptx destino - numero slides" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "nº slides"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 460, "coord_y": 145}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_38":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext con el nº slides del pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 5
                                                                                                                    , "height": 1
                                                                                                                    , "wrap": tk.NONE
                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                    , "fg": "black"
                                                                                                                    , "controltiptext": "Número total de slides del fichero pptx asociado al id pptx"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 545, "coord_y": 147}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }

                                        
                                                                                    } 

                                                                    , "frame_xls_origen":
                                                                                    {"frame":
                                                                                            {"width": 600
                                                                                            , "height": 590
                                                                                            , "bd": 2
                                                                                            , "relief": "solid"
                                                                                            , "bg": "#71C657"
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 250}
                                                                                            }

                                                                                    , "WIDGET_39":
                                                                                                    {"tipo_widget": "label"
                                                                                                    , "desc_tipo_widget": "label titulo del frame 'frame_xls'" #key informativa (no se usa en el resto del codigo del app)
                                                                                                    , "kwargs_config":
                                                                                                                        {"text": "EXCEL ORIGEN" 
                                                                                                                        , "font": ("Calibri", 13, "bold")
                                                                                                                        , "width": 13
                                                                                                                        , "bd": 1
                                                                                                                        , "relief": "solid"
                                                                                                                        , "bg": "#1F40AD"
                                                                                                                        , "fg": "white"
                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                                        }
                                                                                                    }

                                                                                    , "WIDGET_40":
                                                                                                {"tipo_widget": "treeview_propio"
                                                                                                , "desc_tipo_widget": "treeview con los excel configurados para el ID PPTX" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                                                                                                    , "dicc_treeview": {"seleccion_item": "simple"
                                                                                                                                        , "height": 5
                                                                                                                                        ,"columnas_df": ["ID_XLS", "DESC_ID_XLS"]
                                                                                                                                        #importante de que los campos aparezcan el diccionario dicc_tabla_config_sistema["DICC_CAMPOS"]
                                                                                                                                        #del modulo back-end

                                                                                                                                        , "columnas_treeview": ["id xls", "descripción"]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                                        , "width_columnas_treeview": [80, 495]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                                        }
                                                                                                                    , "dicc_rutina_click_item": {"rutina": "def_gui_config_id_pptx_treeview_click_item"#la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                                , "parametros_args": ("ID_XLS",)
                                                                                                                                                }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_41":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label xls origen - id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "id xls"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 180}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_42":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 20
                                                                                                                    , "height": 1
                                                                                                                    , "wrap": tk.NONE
                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                    , "fg": "black"
                                                                                                                    , "bloquear": True
                                                                                                                    , "controltiptext": "Identificador único (interno al app) del fichero excel de origen"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 180}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_43":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton add (paso 1) excel asociado al ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guardar_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                    , "controltiptext": "Abre una ventana de dialogo para seleccionar la ruta del nuevo fichero excel de origen.\Abre una nueva ventana para obligarte a informar los datos generales asociados al nuevo id xls."
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 280, "coord_y": 178}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ABRIR_VENTANA_CREACION_ID_XLS_NUEVO",)   
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_44":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton clear excel asociado al ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_eliminar_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                    , "controltiptext": "Elimina un fichero excel de origen asociado al ID PPTX"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 340, "coord_y": 178}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ELIMINAR_ID_XLS",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_45":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton update excel ya configurado asociado al ID PPTX seleccionado" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 35
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_update_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                    , "controltiptext": "Actualiza la ubicación del fichero excel de origen asociado al id xls."
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 400, "coord_y": 178}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("UPDATE_PATH_ID_XLS",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_46":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label xls origen - ubicacion" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "ubicación"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 210}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_47":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext ruta xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 60
                                                                                                                    , "height": 1
                                                                                                                    , "wrap": tk.NONE
                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                    , "fg": "black"
                                                                                                                    , "bloquear": True
                                                                                                                    , "controltiptext": "Ubicación del fichero excel asociado al id xls"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 210}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_48":
                                                                                                {"tipo_widget": "button"
                                                                                                , "desc_tipo_widget": "boton abrir xls origen" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"width": 30
                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_abrir_fichero, "tupla_pixeles_imagen": (18, 15)}
                                                                                                                    , "controltiptext": "Abre el fichero pptx por si se desea ver en que slides hay que configurar capturas de rangos excel"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 550, "coord_y": 210}
                                                                                                                    , "dicc_rutina":
                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                    , "parametros_args": ("ABRIR_XLS",)    
                                                                                                                                    }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_49":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label xls origen - descripcion" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "descripción"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 240}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_50":
                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                , "desc_tipo_widget": "scrolledtext descripcion id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 67
                                                                                                                    , "height": 2
                                                                                                                    , "wrap": tk.WORD
                                                                                                                    , "fg": "black"
                                                                                                                    , "controltiptext": "Descripción del fichero excel asociado al id xls"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 240}
                                                                                                                    , "alineacion": "center"
                                                                                                                    }
                                                                                                }


                                                                                    , "WIDGET_51":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label xls origen - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "tiempo max"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 285}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_52":
                                                                                                {"tipo_widget": "entry_propio"
                                                                                                , "desc_tipo_widget": "entry xls origen - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                    , "width": 10
                                                                                                                    , "fg": "black"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 287}
                                                                                                                    , "alineacion": "center"
                                                                                                                    , "controltiptext": "Tiempo de espera máximo que se esperara para que se abra el fichero excel de origen por completo"
                                                                                                                    , "dicc_entry":
                                                                                                                                {"formato_validacion": "entero_positivo"
                                                                                                                                , "titulo_messagebox_warning": "TIEMPO ESPERA MÁXIMO APERTURA EXCEL"
                                                                                                                                }
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_53":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label xls origen - actualizar vinculos otros excel" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"text": "actualizar vinculos otros excels"
                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                    , "bg": "#71C657"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 290, "coord_y": 285}
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_54":
                                                                                                {"tipo_widget": "combobox"
                                                                                                , "desc_tipo_widget": "combobox para configurar actualizar vinculos con otros excels" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"font": ("Calibri", 10)
                                                                                                                    , "width": 5
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 520, "coord_y": 287}
                                                                                                                    , "controltiptext": "Habilita o no la actualización de vinculos externos cuando se abra el excel seleccionado\ncuando se realizen los screenshots de rangos de celdas"
                                                                                                                    , "alineacion": "left"
                                                                                                                    , "combobox_lista_opciones": mod_back_end.lista_opciones_id_xls_combobox_actualizar_vinculos
                                                                                                                    }
                                                                                                }

                                                                                    , "WIDGET_55":
                                                                                                {"frame":
                                                                                                    {"tipo_widget": "frame_con_scrollbar"
                                                                                                    , "desc_tipo_widget": "frame con scrollbar vertical dentro de 'frame_xls' para los rangos celda xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                    , "kwargs_config":
                                                                                                                    {"width": 555
                                                                                                                    , "height": 150
                                                                                                                    , "bd": 2
                                                                                                                    , "relief": "solid"
                                                                                                                    , "bg": "#A5D3F0"
                                                                                                                    , "dicc_frame_scrollbar": {"width_visible": 575
                                                                                                                                                , "width_total": 575
                                                                                                                                                , "height_visible": 240
                                                                                                                                                , "height_total": 240
                                                                                                                                                , "tupla_coord_place": (10, 330)
                                                                                                                                                , "velocidad_scrolling": 1
                                                                                                                                                }
                                                                                                                    }
                                                                                                    }

                                                                                                    , "WIDGET_56":
                                                                                                                    {"tipo_widget": "label"
                                                                                                                    , "desc_tipo_widget": "label titulo del frame frame rangos xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                    , "kwargs_config":
                                                                                                                                        {"text": "RANGOS CELDAS" 
                                                                                                                                        , "font": ("Calibri", 13, "bold")
                                                                                                                                        , "width": 15
                                                                                                                                        , "bd": 1
                                                                                                                                        , "relief": "solid"
                                                                                                                                        , "bg": "#1F40AD"
                                                                                                                                        , "fg": "white"
                                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                                                                        }
                                                                                                                    }


                                                                                                    , "WIDGET_57":
                                                                                                                {"tipo_widget": "treeview_propio"
                                                                                                                , "desc_tipo_widget": "treeview de las configuraciones pptx dentro del frame 'frame_pptx'" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                                                                                                                    , "dicc_treeview": {"seleccion_item": "simple"
                                                                                                                                                        , "height": 3
                                                                                                                                                        ,"columnas_df": ["HOJA_XLS", "RANGO_XLS", "SLIDE_PPTX"]
                                                                                                                                                        #importante de que los campos aparezcan el diccionario dicc_tabla_config_sistema["DICC_CAMPOS"]
                                                                                                                                                        #del modulo back-end

                                                                                                                                                        , "columnas_treeview": ["hoja xls", "rango celdas xls", "nº slide pptx"]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                                                        , "width_columnas_treeview": [250, 150, 150]#la lista tiene que tener misma longitud que la lista 'columnas_df'
                                                                                                                                                        }
                                                                                                                                    , "dicc_rutina_click_item": {"rutina": "def_gui_config_id_pptx_treeview_click_item"
                                                                                                                                                                , "parametros_args": ("RANGOS_CELDAS",)
                                                                                                                                                                }
                                                                                                                                    }
                                                                                                                }

                                                                                                                 

                                                                                                    , "WIDGET_58":
                                                                                                                    {"tipo_widget": "label"
                                                                                                                    , "desc_tipo_widget": "label seleccion hojas excel disponibles" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                    , "kwargs_config":
                                                                                                                                        {"text": "hoja xls" 
                                                                                                                                        , "font": ("Calibri", 12, "bold")
                                                                                                                                        , "bg": "#A5D3F0"
                                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 140}
                                                                                                                                        }
                                                                                                                    }

                                                                                                    , "WIDGET_59":
                                                                                                                {"tipo_widget": "combobox"
                                                                                                                , "desc_tipo_widget": "combobox seleccion hojas excel disponibles" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"font": ("Calibri", 10)
                                                                                                                                    , "width": 20
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 142}
                                                                                                                                    , "controltiptext": "Hojas existentes en el excel de origen seleccionado"
                                                                                                                                    , "alineacion": "left"
                                                                                                                                    , "combobox_lista_opciones": [] #se configura dentro de la GUI
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_60":
                                                                                                                    {"tipo_widget": "label"
                                                                                                                    , "desc_tipo_widget": "label rangos celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                    , "kwargs_config":
                                                                                                                                        {"text": "rango celdas" 
                                                                                                                                        , "font": ("Calibri", 12, "bold")
                                                                                                                                        , "bg": "#A5D3F0"
                                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 170}
                                                                                                                                        }
                                                                                                                    }

                                                                                                    , "WIDGET_61":
                                                                                                                {"tipo_widget": "entry_propio"
                                                                                                                , "desc_tipo_widget": "entry rangos celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"font": ("Calibri", 10)
                                                                                                                                    , "width": 23
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 172}
                                                                                                                                    , "controltiptext": "Rango de celdas excel. Se realiza validación al salir del entry si el rango corresponde a un rango excel"
                                                                                                                                    , "alineacion": "left"
                                                                                                                                    , "dicc_entry":
                                                                                                                                                    {"funcion_validacion_personalizada": "func_check_rango_xls_correcto"
                                                                                                                                                    , "modulo_python_alias": "mod_back_end"
                                                                                                                                                    , "resultado_funcion_validacion_personalizada_para_bloquear_exit": False
                                                                                                                                                    , "messagebox_warning_validacion_personalizada": "El rango informado no es un rango excel válido."
                                                                                                                                                    , "titulo_messagebox_warning": mod_back_end.nombre_app
                                                                                                                                                    }
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_62":
                                                                                                                    {"tipo_widget": "label"
                                                                                                                    , "desc_tipo_widget": "label seleccion numero slide pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                    , "kwargs_config":
                                                                                                                                        {"text": "slide pptx" 
                                                                                                                                        , "font": ("Calibri", 12, "bold")
                                                                                                                                        , "bg": "#A5D3F0"
                                                                                                                                        , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 200}

                                                                                                                                        }
                                                                                                                    }

                                                                                                    , "WIDGET_63":
                                                                                                                {"tipo_widget": "combobox"
                                                                                                                , "desc_tipo_widget": "combobox seleccion numero slide pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"font": ("Calibri", 10)
                                                                                                                                    , "width": 20
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 202}
                                                                                                                                    , "controltiptext": "Números de slides existentes en el pptx de destino seleccionado"
                                                                                                                                    , "alineacion": "left"
                                                                                                                                    , "combobox_lista_opciones": [] #se configura dentro de la GUI
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_64":
                                                                                                                {"tipo_widget": "label"
                                                                                                                , "desc_tipo_widget": "label xls origen - descripcion" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"text": "pantallazo"
                                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                                    , "bg": "#A5D3F0"
                                                                                                                                    , "alineacion": "left"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 300, "coord_y": 140}
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_65":
                                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                                , "desc_tipo_widget": "scrolledtext nombre pantallazo" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                                    , "width": 21
                                                                                                                                    , "height": 1
                                                                                                                                    , "controltiptext": "Es el nombre con el cual el pantallazo del rango de celdas excel asociado se guardara como shape en la slide\nindicada del pptx de destino."
                                                                                                                                    , "bloquear": True
                                                                                                                                    , "wrap": tk.NONE
                                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 410, "coord_y": 142}
                                                                                                                                    , "alineacion": "center"
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_66":
                                                                                                                {"tipo_widget": "label"
                                                                                                                , "desc_tipo_widget": "label xls origen - descripcion" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"text": "coordenadas"
                                                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                                                    , "bg": "#A5D3F0"
                                                                                                                                    , "alineacion": "left"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 300, "coord_y": 170}
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_67":
                                                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                                                , "desc_tipo_widget": "tupla coordenadas y simensiones del shape en el pptx de destino" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                                                    , "width": 21
                                                                                                                                    , "height": 1
                                                                                                                                    , "controltiptext": "Son las coordenadas y dimensiones de como ha de salir el pantallazo en el pptx.\nEs tupla con coordenadas 'x' y 'y' seguido del ancho ('width') y alto ('height') del shape."
                                                                                                                                    , "bloquear": True
                                                                                                                                    , "wrap": tk.NONE
                                                                                                                                    , "bg": "#B7C3F5"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 410, "coord_y": 172}
                                                                                                                                    , "alineacion": "center"
                                                                                                                                    }
                                                                                                                }


                                                                                                    , "WIDGET_68":
                                                                                                                {"tipo_widget": "button"
                                                                                                                , "desc_tipo_widget": "boton clean rango celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"width": 27
                                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_clean_rango_celdas, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                                    , "font": ("Calibri", 9, "bold")
                                                                                                                                    , "controltiptext": "Limpia los widgets de rangos de celdas para poder crear nuevos rangos"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 320, "coord_y": 202}
                                                                                                                                    , "dicc_rutina":
                                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                                    , "parametros_args": ("LIMPIAR_RANGO_CELDAS",)   
                                                                                                                                                    }
                                                                                                                                    }
                                                                                                                }


                                                                                                    , "WIDGET_69":
                                                                                                                {"tipo_widget": "button"
                                                                                                                , "desc_tipo_widget": "boton add rango celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"width": 27
                                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_boton_add, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                                    , "controltiptext": "Agrega un nuevo de rango al treeview de rangos de celdas"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 370, "coord_y": 202}
                                                                                                                                    , "dicc_rutina":
                                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                                    , "parametros_args": ("AGREGAR_RANGO_CELDAS",)   
                                                                                                                                                    }
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_70":
                                                                                                                {"tipo_widget": "button"
                                                                                                                , "desc_tipo_widget": "boton delete rango celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"width": 27
                                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_boton_clear, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                                    , "controltiptext": "Elimina un rango existente en el treeview de rangos de celdas"
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 420, "coord_y": 202}
                                                                                                                                    , "dicc_rutina":
                                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                                    , "parametros_args": ("ELIMINAR_RANGO_CELDAS",)    
                                                                                                                                                    }
                                                                                                                                    }
                                                                                                                } 

                                                                                                    , "WIDGET_71":
                                                                                                                {"tipo_widget": "button"
                                                                                                                , "desc_tipo_widget": "boton update rango celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"width": 27
                                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_update_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                                    , "controltiptext": "Actualiza el rango de celdas seleccionado."
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 470, "coord_y": 202}
                                                                                                                                    , "dicc_rutina":
                                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                                                    , "parametros_args": ("UPDATE_RANGO_CELDAS",)   
                                                                                                                                                    }
                                                                                                                                    }
                                                                                                                }

                                                                                                    , "WIDGET_72":
                                                                                                                {"tipo_widget": "button"
                                                                                                                , "desc_tipo_widget": "boton que muestra el screenshot correspondiente al rango celdas" #key informativa (no se usa en el resto del codigo del app)
                                                                                                                , "kwargs_config":
                                                                                                                                    {"width": 27
                                                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_accion_pptx, "tupla_pixeles_imagen": (21, 18)}
                                                                                                                                    , "controltiptext": "Abre una nueva ventana con una muestra del screenshot del rango de celdas excel configurado."
                                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 520, "coord_y": 202}
                                                                                                                                    , "dicc_rutina":
                                                                                                                                                    {"rutina": "def_gui_config_id_pptx_threads"
                                                                                                                                                    , "parametros_args": ("MOSTRAR_SCREENSHOT_RANGO_CELDAS",)}
                                                                                                                                    }


                                                                                                                }

                                                                                                }
                                                                                    }
                        
                                                                    }
                                                    }

                            ###############################################################################################################################################################################################
                            # ventana gui_config_id_pptx_nuevo_id_pptx
                            ###############################################################################################################################################################################################
                            , "gui_config_id_pptx_nuevo_id_pptx":
                                                    {"dicc_config_root":
                                                                        {"title": "CONFIGURACIÓN NUEVO ID PPTX"
                                                                        , "iconbitmap": mod_back_end.template_ico_app_tapar_pluma_tkinter
                                                                        , "tupla_geometry": (610, 160)
                                                                        , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                        , "mantener_nueva_ventana_encima_otras": True
                                                                        , "resizable": (0, 0)
                                                                        }
  
                                                    , "frame_inicio":
                                                                    {"frame":
                                                                            {"width": 610
                                                                            , "height": 160
                                                                            , "bg": "#DFDF70"
                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                            }

                                                                    , "WIDGET_73":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label pptx destino - id pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "id pptx"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_74":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext id pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 20
                                                                                                    , "height": 1
                                                                                                    , "wrap": tk.NONE
                                                                                                    , "bg": "#B7C3F5"
                                                                                                    , "fg": "black"
                                                                                                    , "bloquear": True
                                                                                                    , "controltiptext": "Identificador único (interno al app) del fichero pptx de destino"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 10}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_75":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "Actualiza los datos generales del nuevo id pptx en la ventana anterior" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 35
                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guardar_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                    , "controltiptext": "Guarda los datos informados en esta ventana en la memoria del pc y permite en la ventana anterior\nagregar una opción en el comobobox de selección de id pptx y selecciona dicha opción"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 280, "coord_y": 8}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_gui_config_id_pptx_nuevo_id_pptx_guardar" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_76":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label pptx destino - ubicacion" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "ubicación"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_77":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext ruta pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 60
                                                                                                    , "height": 1
                                                                                                    , "wrap": tk.NONE
                                                                                                    , "bg": "#B7C3F5"
                                                                                                    , "fg": "black"
                                                                                                    , "bloquear": True
                                                                                                    , "controltiptext": "Ubicación del fichero pptx asociado al id pptx"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 40}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_78":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label pptx destino - descripcion" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "descripción"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 70}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_79":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext descripcion id pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 67
                                                                                                    , "height": 2
                                                                                                    , "wrap": tk.WORD
                                                                                                    , "fg": "black"
                                                                                                    , "controltiptext": "Descripción del fichero pptx asociado al id pptx"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 70}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }


                                                                    , "WIDGET_80":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label pptx destino - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "tiempo max"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 115}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_81":
                                                                                {"tipo_widget": "entry_propio"
                                                                                , "desc_tipo_widget": "entry pptx destino - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 10
                                                                                                    , "fg": "black"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 117}
                                                                                                    , "alineacion": "center"
                                                                                                    , "controltiptext": "Tiempo de espera máximo que se esperara para que se abra el fichero pptx de destino por completo"
                                                                                                    , "dicc_entry":
                                                                                                                {"formato_validacion": "entero_positivo"
                                                                                                                , "titulo_messagebox_warning": "TIEMPO ESPERA MÁXIMO APERTURA PPTX"
                                                                                                                }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_82":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label pptx destino - numero slides" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "nº slides"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 460, "coord_y": 115}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_83":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext con el nº slides del pptx" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 5
                                                                                                    , "height": 1
                                                                                                    , "wrap": tk.NONE
                                                                                                    , "bg": "#B7C3F5"
                                                                                                    , "fg": "black"
                                                                                                    , "controltiptext": "Número total de slides del fichero pptx asociado al id pptx"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 545, "coord_y": 117}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }
                                                                    }
                                                    }

                            ###############################################################################################################################################################################################
                            # ventana gui_config_id_pptx_nuevo_id_xls
                            ###############################################################################################################################################################################################
                            , "gui_config_id_pptx_nuevo_id_xls":
                                                    {"dicc_config_root":
                                                                        {"title": "CONFIGURACIÓN NUEVO ID XLS"
                                                                        , "iconbitmap": mod_back_end.template_ico_app_tapar_pluma_tkinter
                                                                        , "tupla_geometry": (610, 160)
                                                                        , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                        , "mantener_nueva_ventana_encima_otras": True
                                                                        , "resizable": (0, 0)
                                                                        }
  
                                                    , "frame_inicio":
                                                                    {"frame":
                                                                            {"width": 610
                                                                            , "height": 160
                                                                            , "bg": "#DFDF70"
                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                            }

                                                                    , "WIDGET_84":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label xls origen - id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "id xls"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_85":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 20
                                                                                                    , "height": 1
                                                                                                    , "wrap": tk.NONE
                                                                                                    , "bg": "#B7C3F5"
                                                                                                    , "fg": "black"
                                                                                                    , "bloquear": True
                                                                                                    , "controltiptext": "Identificador único (interno al app) del fichero excel de origen"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 10}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_86":
                                                                                {"tipo_widget": "button"
                                                                                , "desc_tipo_widget": "Actualiza los datos generales del nuevo id xls en la ventana anterior" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"width": 35
                                                                                                    , "dicc_imagen": {"png_imagen": mod_back_end.template_img_guardar_ruta, "tupla_pixeles_imagen": (21, 18)}
                                                                                                    , "controltiptext": "Guarda los datos informados en esta ventana en la memoria del pc y permite en la ventana anterior\ncrear una nueva entrada en el treeview de id xls y seleccionarlo para continuar con su configuración"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 280, "coord_y": 8}
                                                                                                    , "dicc_rutina":
                                                                                                                    {"rutina": "def_gui_config_id_pptx_nuevo_id_xls_guardar" #la rutina tiene que estar definida en la clase gui_config_id_pptx
                                                                                                                    }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_87":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label xls origen - ubicacion" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "ubicación"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 40}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_88":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext ruta xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 60
                                                                                                    , "height": 1
                                                                                                    , "wrap": tk.NONE
                                                                                                    , "bg": "#B7C3F5"
                                                                                                    , "fg": "black"
                                                                                                    , "bloquear": True
                                                                                                    , "controltiptext": "Ubicación del fichero excel asociado al id xls"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 40}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_89":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label xls origen - descripcion" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "descripción"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 70}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_90":
                                                                                {"tipo_widget": "scrolledtext_propio"
                                                                                , "desc_tipo_widget": "scrolledtext descripcion id xls" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 67
                                                                                                    , "height": 2
                                                                                                    , "wrap": tk.WORD
                                                                                                    , "fg": "black"
                                                                                                    , "controltiptext": "Descripción del fichero excel asociado al id xls"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 70}
                                                                                                    , "alineacion": "center"
                                                                                                    }
                                                                                }


                                                                    , "WIDGET_91":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label xls origen - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "tiempo max"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 115}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_92":
                                                                                {"tipo_widget": "entry_propio"
                                                                                , "desc_tipo_widget": "entry xls origen - tiempo espera apertura" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10, "bold")
                                                                                                    , "width": 10
                                                                                                    , "fg": "black"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 110, "coord_y": 117}
                                                                                                    , "alineacion": "center"
                                                                                                    , "controltiptext": "Tiempo de espera máximo que se esperara para que se abra el fichero excel de origen por completo"
                                                                                                    , "dicc_entry":
                                                                                                                {"formato_validacion": "entero_positivo"
                                                                                                                , "titulo_messagebox_warning": "TIEMPO ESPERA MÁXIMO APERTURA EXCEL"
                                                                                                                }
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_93":
                                                                                {"tipo_widget": "label"
                                                                                , "desc_tipo_widget": "label xls origen - actualizar vinculos otros excel" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"text": "actualizar vinculos otros excels"
                                                                                                    , "font": ("Calibri", 12, "bold")
                                                                                                    , "bg": "#DFDF70"
                                                                                                    , "alineacion": "left"
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 290, "coord_y": 115}
                                                                                                    }
                                                                                }

                                                                    , "WIDGET_94":
                                                                                {"tipo_widget": "combobox"
                                                                                , "desc_tipo_widget": "combobox para configurar actualizar vinculos con otros excels" #key informativa (no se usa en el resto del codigo del app)
                                                                                , "kwargs_config":
                                                                                                    {"font": ("Calibri", 10)
                                                                                                    , "width": 5
                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 520, "coord_y": 117}
                                                                                                    , "controltiptext": "Habilita o no la actualización de vinculos externos cuando se abra el excel seleccionado\ncuando se realizen los screenshots de rangos de celdas"
                                                                                                    , "alineacion": "left"
                                                                                                    , "combobox_lista_opciones": mod_back_end.lista_opciones_id_xls_combobox_actualizar_vinculos
                                                                                                    }
                                                                                }

                                                                    }
                                                    }

                            ###############################################################################################################################################################################################
                            # ventana gui_screenshot_muestra
                            ###############################################################################################################################################################################################
                            , "gui_screenshot_muestra":
                                                    {"dicc_config_root":
                                                                        {"title": "MUESTRA SCREENSHOT"
                                                                        , "iconbitmap": mod_back_end.template_ico_app_tapar_pluma_tkinter
                                                                        , "tupla_geometry": (630, 600)
                                                                        , "bloquear_interaccion_nueva_ventana_con_otras": True
                                                                        , "mantener_nueva_ventana_encima_otras": True
                                                                        , "resizable": (1, 1)#aqui si es extensible por el usuario
                                                                        }

                                                    , "frame_integrado_root":
                                                                    {"frame_inicio":
                                                                                    {"kwargs_config":
                                                                                                    {"dicc_frame_scrollbar": {"width_visible": 600
                                                                                                                                , "width_total": 1500
                                                                                                                                , "height_visible": 600
                                                                                                                                , "height_total": 1500
                                                                                                                                , "tupla_coord_place": (0, 0)
                                                                                                                                , "velocidad_scrolling": 1
                                                                                                                                }
                                                                                                    }
                                                                                    }

                                                                    , "frame_label":
                                                                                    {"frame":
                                                                                            {"width": 600
                                                                                            , "height": 600
                                                                                            , "dicc_colocacion": {"metodo": "place", "coord_x": 0, "coord_y": 0}
                                                                                            }

                                                                                    , "WIDGET_95":
                                                                                                {"tipo_widget": "label"
                                                                                                , "desc_tipo_widget": "label con muestra del screensgot (reaqlizada en el paso de CONFIG_PASO_1)" #key informativa (no se usa en el resto del codigo del app)
                                                                                                , "kwargs_config":
                                                                                                                    {"bd": 2
                                                                                                                    , "relief": "solid"
                                                                                                                    , "dicc_colocacion": {"metodo": "place", "coord_x": 10, "coord_y": 10}}
                                                                                                }
                                                                                    }
                                                                    }
                                                    }
                            }


                                

    #se crea el root usando el kwargs dicc_config_root del diccionario creado anteriormente
    #y se inicia la clase gui_gui_app
    dicc_kwargs_gui_ventana_inicio = dicc_kwargs_config_gui["gui_ventana_inicio"]
    dicc_kwargs_gui_ventana_inicio_dicc_config_root = dicc_kwargs_gui_ventana_inicio["dicc_config_root"]

    root = mod_utils.gui_tkinter_widgets(None, tipo_widget_param = "root", **dicc_kwargs_gui_ventana_inicio_dicc_config_root)


    #se localiza si la resolucion de pantalla es la recomendada para poder informarlo en la GUi en caso de que no lo sea
    resolucion_pantalla_actual = str(root.widget_objeto.winfo_screenwidth()) + "x" + str(root.widget_objeto.winfo_screenheight())

    label_resolucion_pantalla = (f"Resolución de pantalla recomendada: '{mod_back_end.resolucion_pantalla_recomendada}', la actual es '{resolucion_pantalla_actual}'"
                                    if resolucion_pantalla_actual != mod_back_end.resolucion_pantalla_recomendada else "")
    

    #se recalculan parmetros de la GUI segun la resolucion de pantalla (el texto del label abajo de la GUI y el height del root)
    #se usa el widget WIDGET_18 del diccionario dicc_kwargs_gui
    dicc_kwargs_config_gui["gui_ventana_inicio"]["frame_resolucion_pantalla"]["WIDGET_18"]["kwargs_config"]["text"] = label_resolucion_pantalla



    #se configura el root y se inicializa la clase gui_app
    root.config_atributos(**dicc_kwargs_gui_ventana_inicio)

    gui_ventana_inicio(root, **dicc_kwargs_config_gui)

    root.widget_objeto.mainloop()


