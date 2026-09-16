from django.test import SimpleTestCase

from .sig_sync.client import clasificar_sweetalert


class ClasificarSweetAlertTests(SimpleTestCase):
    def test_registro_usuario_agregado_con_icono_advertencia_es_exito(self):
        """El caso real: éxito con popup tipo 'Registro' e icono de advertencia."""
        self.assertFalse(
            clasificar_sweetalert(
                "Registro", "Usuario agregado", con_advertencia=True
            )
        )

    def test_popup_exito_normal_sin_iconos_es_exito(self):
        self.assertFalse(clasificar_sweetalert("Registro", "Usuario agregado"))

    def test_sin_popup_prevalece_exito(self):
        self.assertFalse(clasificar_sweetalert("", ""))

    def test_advertencia_es_error(self):
        self.assertTrue(
            clasificar_sweetalert("Advertencia", "El usuario ya existe")
        )
        self.assertTrue(
            clasificar_sweetalert("Advertencia", "Faltan datos obligatorios")
        )

    def test_icono_error_es_error(self):
        self.assertTrue(
            clasificar_sweetalert("Registro", "Usuario agregado", con_error=True)
        )

    def test_icono_advertencia_sin_texto_exitoso_es_error(self):
        self.assertTrue(
            clasificar_sweetalert("SIG", "No fue posible continuar", con_advertencia=True)
        )

    def test_icono_advertencia_con_texto_exitoso_es_exito(self):
        for cuerpo in ("Usuario creado", "Correo actualizado", "Usuario deshabilitado"):
            self.assertFalse(
                clasificar_sweetalert("Registro", cuerpo, con_advertencia=True)
            )

    def test_titulo_case_insensitive(self):
        self.assertTrue(clasificar_sweetalert("ADVERTENCIA", "Datos inválidos"))
        self.assertFalse(clasificar_sweetalert("REGISTRO", "Usuario agregado", con_advertencia=True))