import hashlib
import json
import os

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


def verificar_password(password):
    """Verifica la contraseña contra el hash almacenado"""
    if not password:
        return False

    config = cargar_configuracion()
    password_hash_almacenado = config.get("password_hash", PASSWORD_HASH)

    hash_password = hashlib.sha256(password.encode()).hexdigest()
    return hash_password == password_hash_almacenado


def hash_password(password):
    """Genera el hash de una contraseña"""
    return hashlib.sha256(password.encode()).hexdigest()