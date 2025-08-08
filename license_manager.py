import platform
import getpass
import hashlib
import requests
from datetime import datetime
import threading
import customtkinter as ctk
from tkinter import messagebox


class LicenseManager:
    def __init__(self):
        # URL de tu Firebase Realtime Database
        self.firebase_url = "https://recibos-anses-default-rtdb.firebaseio.com"
        self.machine_id = self.get_machine_id()

    def get_machine_id(self):
        """Genera un ID único para la máquina"""
        try:
            machine_info = f"{platform.node()}-{platform.machine()}-{getpass.getuser()}-{platform.platform()}"
            return hashlib.sha256(machine_info.encode()).hexdigest()
        except:
            return "UNKNOWN_MACHINE"

    def check_license(self):
        """Verifica si esta máquina tiene licencia válida en Firebase"""
        try:
            url = f"{self.firebase_url}/licenses/{self.machine_id}.json"
            response = requests.get(url, timeout=10)

            if response.status_code == 200:
                license_data = response.json()

                if license_data is None:
                    return False, f"Máquina no registrada.\nID: {self.machine_id}"

                if license_data.get('active', False):
                    if 'expires_at' in license_data:
                        try:
                            expires_at = datetime.fromisoformat(license_data['expires_at'])
                            if datetime.now() > expires_at:
                                return False, "Licencia expirada"
                        except:
                            pass

                    self.update_last_used()
                    return True, "Licencia válida, cierre la ventana para continuar"
                else:
                    return False, f"Licencia desactivada.\nID: {self.machine_id}"
            else:
                return False, f"Error de conexión: {response.status_code}"

        except requests.exceptions.RequestException:
            return False, f"Sin conexión a internet.\nID: {self.machine_id}"
        except Exception as e:
            return False, f"Error: {str(e)}"

    def update_last_used(self):
        """Actualiza la fecha de último uso en Firebase"""
        try:
            url = f"{self.firebase_url}/licenses/{self.machine_id}/last_used.json"
            requests.put(url, json=datetime.now().isoformat(), timeout=5)
        except:
            pass  # No es crítico si falla


class LicenseDialog(ctk.CTkToplevel):
    def __init__(self, parent, license_manager):
        super().__init__(parent)
        self.parent = parent
        self.license_manager = license_manager
        self.license_valid = False

        # Variables para threading
        self.license_result = None
        self.checking_license = False

        self.setup_ui()
        self.check_license()

    def setup_ui(self):
        self.title("🔐 Verificación de Licencia")
        self.geometry("600x500")
        self.resizable(False, False)

        # Centrar ventana
        self.transient(self.parent)
        self.grab_set()

        # Frame principal
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Título
        title_label = ctk.CTkLabel(
            main_frame,
            text="🔐 VERIFICACIÓN DE LICENCIA",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=["#1E40AF", "#60A5FA"]
        )
        title_label.pack(pady=(20, 10))

        # Información de la máquina
        info_frame = ctk.CTkFrame(main_frame, fg_color=["#F8FAFC", "#1E293B"])
        info_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            info_frame,
            text="💻 INFORMACIÓN DE LA MÁQUINA",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 10))

        machine_info = [
            f"🖥️ Equipo: {platform.node()}",
            f"👤 Usuario: {getpass.getuser()}",
            f"🔧 Sistema: {platform.system()}",
            f"🆔 ID: {self.license_manager.machine_id[:32]}..."
        ]

        for info in machine_info:
            ctk.CTkLabel(
                info_frame,
                text=info,
                font=ctk.CTkFont(size=12),
                text_color=["#6B7280", "#9CA3AF"]
            ).pack(pady=2)

        ctk.CTkLabel(info_frame, text="", height=10).pack()

        # Estado de la licencia
        self.status_frame = ctk.CTkFrame(main_frame, fg_color=["#FEF2F2", "#1F1F1F"])
        self.status_frame.pack(fill="x", pady=(0, 20))

        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="🔄 Verificando licencia...",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=["#DC2626", "#EF4444"]
        )
        self.status_label.pack(pady=15)

        # Instrucciones
        instructions_frame = ctk.CTkFrame(main_frame, fg_color=["#F0F9FF", "#1E293B"])
        instructions_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(
            instructions_frame,
            text="📋 INSTRUCCIONES",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 10))

        instructions_text = ctk.CTkTextbox(instructions_frame, height=100, wrap="word")
        instructions_text.pack(fill="x", padx=15, pady=(0, 15))

        instructions = """ℹ️ VERIFICACIÓN DE LICENCIA:
Si tu máquina no está autorizada, contacta al administrador con tu ID de máquina.
El administrador debe crear una licencia para tu máquina usando el generador de licencias.

📞 CONTACTO: Proporciona tu ID de máquina completo al administrador."""

        instructions_text.insert("1.0", instructions)
        instructions_text.configure(state="disabled")

        # Botones
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=(0, 20))

        self.refresh_btn = ctk.CTkButton(
            buttons_frame,
            text="🔄 Verificar Nuevamente",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            command=self.check_license
        )
        self.refresh_btn.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.continue_btn = ctk.CTkButton(
            buttons_frame,
            text="✅ Continuar",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            state="disabled",
            fg_color=["#10B981", "#059669"],
            command=self.continue_app
        )
        self.continue_btn.pack(side="right", fill="x", expand=True, padx=(10, 0))

        # ID completo para copiar
        id_frame = ctk.CTkFrame(main_frame, fg_color=["#F8FAFC", "#1E293B"])
        id_frame.pack(fill="x")

        ctk.CTkLabel(
            id_frame,
            text="🆔 ID COMPLETO DE LA MÁQUINA:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(pady=(15, 5))

        self.id_textbox = ctk.CTkTextbox(id_frame, height=60, wrap="word")
        self.id_textbox.pack(fill="x", padx=15, pady=(0, 15))
        self.id_textbox.insert("1.0", self.license_manager.machine_id)
        self.id_textbox.configure(state="disabled")

    def check_license(self):
        """Verifica la licencia"""
        self.refresh_btn.configure(text="🔄 Verificando...", state="disabled")
        self.update_status(False, "🔄 Verificando licencia...")

        self.license_result = None
        self.checking_license = True

        thread = threading.Thread(target=self._check_license_thread)
        thread.daemon = True
        thread.start()

        self._poll_license_result()

    def _check_license_thread(self):
        """Hilo de verificación de licencia"""
        try:
            valid, message = self.license_manager.check_license()
            self.license_result = (valid, message)
            self.checking_license = False
        except Exception as e:
            self.license_result = (False, f"Error en verificación: {str(e)}")
            self.checking_license = False

    def _poll_license_result(self):
        """Verifica periódicamente si hay resultados del hilo de licencia"""
        if hasattr(self, 'checking_license') and not self.checking_license and self.license_result:
            valid, message = self.license_result
            self._update_license_result(valid, message)
            self.license_result = None
            self.checking_license = False
        elif hasattr(self, 'checking_license') and self.checking_license:
            self.after(100, self._poll_license_result)

    def _update_license_result(self, valid, message):
        """Actualiza el resultado de la verificación (ejecutado en hilo principal)"""
        self.refresh_btn.configure(text="🔄 Verificar Nuevamente", state="normal")
        self.update_status(valid, message)

    def update_status(self, valid, message):
        """Actualiza el estado visual de la licencia"""
        self.license_valid = valid

        if valid:
            self.status_frame.configure(fg_color=["#F0FDF4", "#1F2937"])
            self.status_label.configure(
                text=f"✅ {message}",
                text_color=["#059669", "#10B981"]
            )
            self.continue_btn.configure(state="normal")
        else:
            self.status_frame.configure(fg_color=["#FEF2F2", "#1F1F1F"])
            self.status_label.configure(
                text=f"❌ {message}",
                text_color=["#DC2626", "#EF4444"]
            )
            self.continue_btn.configure(state="disabled")

    def continue_app(self):
        """Continúa con la aplicación si la licencia es válida"""
        if self.license_valid:
            self.destroy()
        else:
            messagebox.showerror("Error", "Debes tener una licencia válida para continuar")
