from datetime import datetime
import customtkinter as ctk


class ConsoleWidget(ctk.CTkTextbox):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color="#00FF00",
            fg_color="#1a1a1a",
            border_color="#333333",
            border_width=2
        )
        self.log_count = 0

    def log(self, mensaje, tipo="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_count += 1

        color_map = {
            "info": "#00FF00",
            "warning": "#FFA500",
            "error": "#FF4444",
            "success": "#00FF88",
            "process": "#00AAFF"
        }

        color = color_map.get(tipo, "#00FF00")

        line = f"[{timestamp}] [{self.log_count:03d}] {mensaje}\n"

        self.insert("end", line)
        self.see("end")

        lines = self.get("1.0", "end").split('\n')
        if len(lines) > 100:
            self.delete("1.0", f"{len(lines)-100}.0")
