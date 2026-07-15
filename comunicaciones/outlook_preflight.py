# -*- coding: utf-8 -*-
"""Diagnóstico liviano de entorno Outlook/pywin32."""

import sys


def verificar_entorno_outlook() -> tuple[bool, str]:
    """Verifica si Outlook COM está disponible para crear borradores.

    Retorna ``(ok, mensaje)`` con un texto accionable para GUI/CLI. No envía
    correos ni guarda elementos; solo intenta crear y descartar un item en
    memoria.
    """
    if sys.platform != "win32":
        return False, "Outlook COM solo está disponible en Windows. Usa dry-run HTML."

    pythoncom = None
    com_inicializado = False
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as exc:
        return False, f"pywin32 no está disponible: {exc}. Instala con: pip install pywin32"

    try:
        pythoncom.CoInitialize()
        com_inicializado = True
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = "CSMP RUS - verificación de entorno"
        mail.Close(1)  # olDiscard
        return True, "Outlook disponible para crear borradores."
    except Exception as exc:  # noqa: BLE001 - diagnóstico operativo, no regla de negocio
        return False, f"No se pudo inicializar Outlook: {exc}. Abre Outlook y verifica el perfil."
    finally:
        if pythoncom is not None and com_inicializado:
            try:
                pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001 - limpieza best effort
                pass
