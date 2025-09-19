import hashlib
import os
import json

# Configuración
CONFIG_FILE = "config.json"
PASSWORD_HASH = "c8a6ed3ac08087cc037c2fc7846a7f95976b8f5bfbaf2d9540cf89b74452b034"


def cargar_configuracion():
    """Carga la configuración desde el archivo"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            return {"password_hash": PASSWORD_HASH}
    return {"password_hash": PASSWORD_HASH}


def guardar_configuracion(config):
    """Guarda la configuración en el archivo"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f)


def verificar_password_actual(password_actual):
    """Verifica la contraseña actual"""
    config = cargar_configuracion()
    password_actual_hash = config.get("password_hash", PASSWORD_HASH)

    hash_ingresado = hashlib.sha256(password_actual.encode()).hexdigest()
    return hash_ingresado == password_actual_hash


def cambiar_password_web(password_actual, nueva_password, confirmar_password):
    """Función principal para cambiar la contraseña (versión web)"""
    # Verificar contraseña actual
    if not verificar_password_actual(password_actual):
        return False, "Contraseña actual incorrecta"

    # Verificar que coincidan
    if nueva_password != confirmar_password:
        return False, "Las contraseñas no coinciden"

    # Verificar fortaleza de la contraseña
    if len(nueva_password) < 4:
        return False, "La contraseña debe tener al menos 4 caracteres"

    # Generar hash de la nueva contraseña
    hash_nuevo = hashlib.sha256(nueva_password.encode()).hexdigest()

    # Guardar nueva contraseña
    config = cargar_configuracion()
    config["password_hash"] = hash_nuevo
    guardar_configuracion(config)

    return True, "Contraseña cambiada correctamente"


def generar_hash_password(password):
    """Genera el hash de una contraseña"""
    if password:
        return hashlib.sha256(password.encode()).hexdigest()
    return None
