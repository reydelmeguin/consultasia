import hashlib
from tkinter import messagebox

class Autenticacion:
    CREDENCIALES = {
        "admin": "5e884898da28047151d0e56f8dc6292773603d0d6aabbdd62a11ef721d1542d8",  # password
        "analista": "a665a45920422f9d417e4867efdc4fb8a04a1f3fff1fa07e998e86f7f7a27ae3"  # 123
    }

    @staticmethod
    def verificar_usuario(usuario, password):
        if usuario in Autenticacion.CREDENCIALES:
            hashed_pwd = hashlib.sha256(password.encode()).hexdigest()
            return Autenticacion.CREDENCIALES[usuario] == hashed_pwd
        return False