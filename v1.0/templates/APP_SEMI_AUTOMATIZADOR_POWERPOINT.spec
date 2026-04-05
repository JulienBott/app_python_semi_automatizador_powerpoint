# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
	['APP_PPTX_1_GUI.py'],
	pathex=[r"mi_ruta"],
	binaries=[],
	datas=[(r"mi_ruta\ico_app_tapar_pluma_tkinter.ico", "."),
				(r"mi_ruta\PLANTILLA_CONFIG.xlsx", "."),
				(r"mi_ruta\pdf_guia_usuario.pdf", "."),
				
				(r"mi_ruta\img_guia_usuario.png", "."),
				(r"mi_ruta\img_config.png", "."),
				(r"mi_ruta\img_boton_ver.png", "."),
				(r"mi_ruta\img_accion_pptx.png", "."),
				(r"mi_ruta\img_boton_add.png", "."),
				(r"mi_ruta\img_boton_clear.png", "."),
				(r"mi_ruta\img_guardar.png", "."),
				(r"mi_ruta\img_abrir_fichero.png", "."),
				(r"mi_ruta\img_update_datos_config.png", "."),
				(r"mi_ruta\img_clean_rango_celdas.png", "."),
				(r"mi_ruta\img_guardar_id.png", "."),
				(r"mi_ruta\img_eliminar_id.png", "."),
				(r"mi_ruta\img_guardar_ruta.png", "."),
				(r"mi_ruta\img_eliminar_ruta.png", "."),
				(r"mi_ruta\img_update_ruta.png", "."),
				(r"mi_ruta\img_messagebox_askokcancel.png", "."),
				(r"mi_ruta\img_messagebox_showwarning.png", "."),
				(r"mi_ruta\img_messagebox_showerror.png", "."),
				(r"mi_ruta\img_messagebox_showinfo.png", "."),],
				
		
	hiddenimports=[],
	hookspath=[],
	hooksconfig={},
	runtime_hooks=[],
	excludes=[],
	noarchive=False,
)
pyz = PYZ(a.pure)

	exe = EXE(
		pyz,
		a.scripts,
		a.binaries,
		a.datas,
		[],
		name='APP_SEMI_AUTOMATIZADOR_POWERPOINT',
		debug=False,
		bootloader_ignore_signals=False,
		strip=False,
		upx=True,
		upx_exclude=[],
		runtime_tmpdir=None,
		console=False,
		disable_windowed_traceback=False,
		argv_emulation=False,
		target_arch=None,
		codesign_identity=None,
		entitlements_file=None,
		icon="ico_app.ico",
	)
	
	
	
	
