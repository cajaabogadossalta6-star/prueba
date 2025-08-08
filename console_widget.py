from datetime import datetime
import customtkinter as ctk


class ConsoleWidget(ctk.CTkTextbox):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color="#E5E7EB",
            fg_color="#111827",
            border_color="#374151",
            border_width=2
        )
        self.log_count = 0

    def log(self, mensaje, tipo="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_count += 1

        color_map = {
            "info": "#F8FAFC",
            "warning": "#FBBF24",
            "error": "#F87171",
            "success": "#34D399",
            "process": "#60A5FA"
        }

        color = color_map.get(tipo, "#F8FAFC")

        line = f"[{timestamp}] [{self.log_count:03d}] {mensaje}\n"
        tag = f"log_{self.log_count}"

        self.insert("end", line, tag)
        self.tag_config(tag, foreground=color)
        self.see("end")

        lines = self.get("1.0", "end").split('\n')
        if len(lines) > 100:
            self.delete("1.0", f"{len(lines)-100}.0")
