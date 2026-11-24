#!/usr/bin/env python3
"""
CENTRAL DE AUTOMAÇÕES BI-LEC
Gerenciador Visual Multi-Tarefas (WhatsApp, SQL e mais)
Linguagem: Python 3 (Tkinter 9.0)
Motor de Agendamento: Próprio e 100% Independente do Windows
"""

import os
import sys
import json
import time
import subprocess
import threading
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox

APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR) if os.path.basename(APP_DIR).lower() == "auto_agenda" else APP_DIR
CONFIG_FILE = os.path.join(APP_DIR, "horarios_agendados.json")

PYTHON_EXE = sys.executable.replace("pythonw.exe", "python.exe")
if not os.path.exists(PYTHON_EXE):
    PYTHON_EXE = r"C:\Python314\python.exe"

TAREFAS_DEFAULT = {
    "whatsapp": {
        "id": "whatsapp",
        "nome": "Relatório WhatsApp (Power BI)",
        "icone": "📱",
        "script": "A1_Envio_Zap.py",
        "horarios": ["09:03", "11:03", "13:03", "15:03", "16:03", "17:03", "18:03"],
        "apenas_dias_uteis": True,
        "ativo": True,
        "ultimo_status": "Pronto para iniciar",
        "ultimo_horario": "-"
    },
    "sql": {
        "id": "sql",
        "nome": "Copiar Arquivos SQL",
        "icone": "🗄️",
        "script": "A2_Envio_SQL.py",
        "horarios": ["09:05", "11:05", "13:05", "15:05", "16:05", "17:05", "18:05", "19:05"],
        "apenas_dias_uteis": True,
        "ativo": True,
        "ultimo_status": "Pronto para iniciar",
        "ultimo_horario": "-"
    }
}


def vpn_esta_conectada():
    try:
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "Get-NetAdapter | Where-Object { $_.InterfaceDescription -like '*Fortinet*' -and $_.Status -eq 'Up' }"],
            capture_output=True,
            text=True,
            timeout=5
        )
        return bool(res.stdout.strip())
    except Exception:
        return False


def conectar_vpn_automatica():
    ps1_path = os.path.join(APP_DIR, "Conectar_VPN.ps1")
    if os.path.exists(ps1_path):
        subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-NoProfile", "-File", ps1_path],
            cwd=r"C:\Program Files\Fortinet\FortiClient",
            capture_output=True,
            timeout=35
        )


def carregar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                dados = json.load(f)
            # Se for formato novo com "tarefas"
            if "tarefas" in dados:
                return dados
            # Se for formato antigo (migra preservando horários do WhatsApp)
            cfg_migrada = {"tarefas": {}}
            cfg_migrada["tarefas"]["whatsapp"] = dict(TAREFAS_DEFAULT["whatsapp"])
            cfg_migrada["tarefas"]["sql"] = dict(TAREFAS_DEFAULT["sql"])
            if "horarios" in dados:
                cfg_migrada["tarefas"]["whatsapp"]["horarios"] = dados["horarios"]
            if "apenas_dias_uteis" in dados:
                cfg_migrada["tarefas"]["whatsapp"]["apenas_dias_uteis"] = dados["apenas_dias_uteis"]
            if "ativo" in dados:
                cfg_migrada["tarefas"]["whatsapp"]["ativo"] = dados["ativo"]
            salvar_config(cfg_migrada)
            return cfg_migrada
        except Exception:
            pass

    cfg_nova = {"tarefas": TAREFAS_DEFAULT}
    salvar_config(cfg_nova)
    return cfg_nova


def salvar_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar config: {e}")


class CentralAutomacoesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BI-LEC • Central de Automações")
        self.root.geometry("680x760")
        self.root.minsize(620, 700)
        self.root.configure(bg="#0f172a")

        self.config = carregar_config()
        self.tarefa_selecionada = "whatsapp"
        self.executando_tarefa_id = None
        self.processo_atual = None

        self._configurar_estilos()
        self._construir_interface()

        # Inicia o motor independente do agendador
        self.loop_ativo = True
        self.thread_relogio = threading.Thread(target=self._motor_agendador, daemon=True)
        self.thread_relogio.start()

        self._atualizar_tela_tarefa()
        self.root.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    def _configurar_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Card.TFrame", background="#1e293b", relief="flat")
        style.configure("TLabel", background="#1e293b", foreground="#f8fafc", font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"), foreground="#38bdf8", background="#0f172a")
        style.configure("SubTitle.TLabel", font=("Segoe UI", 9), foreground="#94a3b8", background="#0f172a")
        style.configure("CardTitle.TLabel", font=("Segoe UI", 11, "bold"), foreground="#e2e8f0", background="#1e293b")

        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), background="#059669", foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#10b981"), ("disabled", "#334155")])

        style.configure("Secondary.TButton", font=("Segoe UI", 9), background="#334155", foreground="#f8fafc")
        style.map("Secondary.TButton", background=[("active", "#475569")])

        style.configure("Danger.TButton", font=("Segoe UI", 9), background="#991b1b", foreground="#ffffff")
        style.map("Danger.TButton", background=[("active", "#dc2626")])

    def _construir_interface(self):
        # Header
        header = tk.Frame(self.root, bg="#0f172a")
        header.pack(fill="x", padx=20, pady=(15, 8))

        top_tit_row = tk.Frame(header, bg="#0f172a")
        top_tit_row.pack(fill="x")

        lbl_tit = ttk.Label(top_tit_row, text="⚡ BI-LEC  |  Central de Automações", style="Title.TLabel")
        lbl_tit.pack(side="left")

        self.btn_vpn_status = tk.Button(
            top_tit_row,
            text="🛡️ VPN CONECTADA",
            font=("Segoe UI", 9, "bold"),
            bg="#065f46",
            fg="#34d399",
            bd=0,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self._conectar_vpn_click
        )
        self.btn_vpn_status.pack(side="right")

        lbl_sub = ttk.Label(header, text="Gerencie e agende suas rotinas de forma autônoma e sem limites", style="SubTitle.TLabel")
        lbl_sub.pack(anchor="w")

        # BARRA DE ABAS / SELETOR DE TAREFA
        tab_bar = tk.Frame(self.root, bg="#0f172a")
        tab_bar.pack(fill="x", padx=20, pady=(5, 10))

        self.btn_tab_zap = tk.Button(
            tab_bar,
            text="📱  Relatório WhatsApp",
            font=("Segoe UI", 10, "bold"),
            bg="#2563eb",
            fg="#ffffff",
            activebackground="#3b82f6",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2",
            command=lambda: self._mudar_aba("whatsapp")
        )
        self.btn_tab_zap.pack(side="left", padx=(0, 8))

        self.btn_tab_sql = tk.Button(
            tab_bar,
            text="🗄️  Copiar Arquivos SQL",
            font=("Segoe UI", 10, "bold"),
            bg="#1e293b",
            fg="#94a3b8",
            activebackground="#334155",
            activeforeground="#ffffff",
            bd=0,
            padx=16,
            pady=8,
            cursor="hand2",
            command=lambda: self._mudar_aba("sql")
        )
        self.btn_tab_sql.pack(side="left")

        # CARD 1: Status & Controles da Tarefa Atual
        card_status = ttk.Frame(self.root, style="Card.TFrame", padding=15)
        card_status.pack(fill="x", padx=20, pady=5)

        top_row = tk.Frame(card_status, bg="#1e293b")
        top_row.pack(fill="x")

        self.lbl_tarefa_nome = ttk.Label(top_row, text="STATUS DA TAREFA", style="CardTitle.TLabel")
        self.lbl_tarefa_nome.pack(side="left")

        self.lbl_badge = tk.Label(top_row, text="● ATIVA", font=("Segoe UI", 9, "bold"), bg="#065f46", fg="#34d399", padx=8, pady=2)
        self.lbl_badge.pack(side="right")

        self.lbl_proximo = ttk.Label(card_status, text="Calculando...", font=("Segoe UI", 12, "bold"))
        self.lbl_proximo.pack(anchor="w", pady=(10, 4))

        self.lbl_ultimo = ttk.Label(card_status, text="Última execução: -", font=("Segoe UI", 9), foreground="#94a3b8")
        self.lbl_ultimo.pack(anchor="w")

        # Botões de ação
        ctrl_bar = tk.Frame(card_status, bg="#1e293b")
        ctrl_bar.pack(fill="x", pady=(12, 0))

        self.btn_toggle = ttk.Button(ctrl_bar, text="⏸ Pausar Tarefa", style="Secondary.TButton", command=self._toggle_ativo)
        self.btn_toggle.pack(side="left")

        self.btn_exec_agora = ttk.Button(ctrl_bar, text="▶ Executar Esta Tarefa Agora", style="Primary.TButton", command=self._executar_agora_click)
        self.btn_exec_agora.pack(side="right")

        # CARD 2: Horários Programados
        card_horarios = ttk.Frame(self.root, style="Card.TFrame", padding=15)
        card_horarios.pack(fill="both", expand=True, padx=20, pady=8)

        ttk.Label(card_horarios, text="HORÁRIOS DE DISPARO PROGRAMADOS", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))

        # Adicionar novo horário
        add_bar = tk.Frame(card_horarios, bg="#1e293b")
        add_bar.pack(fill="x", pady=(0, 10))

        ttk.Label(add_bar, text="Novo Horário (HH:MM):", font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        self.entry_hora = tk.Entry(add_bar, width=8, font=("Segoe UI", 11, "bold"), justify="center", bg="#0f172a", fg="#ffffff", insertbackground="#ffffff")
        self.entry_hora.insert(0, "12:00")
        self.entry_hora.pack(side="left", padx=(0, 8))

        btn_add = ttk.Button(add_bar, text="+ Adicionar", style="Primary.TButton", command=self._adicionar_horario)
        btn_add.pack(side="left")

        btn_del = ttk.Button(add_bar, text="✕ Remover Selecionado", style="Danger.TButton", command=self._remover_horario)
        btn_del.pack(side="right")

        # Lista de horários
        list_frame = tk.Frame(card_horarios, bg="#0f172a")
        list_frame.pack(fill="both", expand=True)

        self.lista_box = tk.Listbox(
            list_frame,
            font=("Segoe UI", 12, "bold"),
            bg="#0f172a",
            fg="#f8fafc",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            activestyle="none",
            borderwidth=0,
            highlightthickness=0
        )
        self.lista_box.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        scroll = tk.Scrollbar(list_frame, orient="vertical", command=self.lista_box.yview)
        scroll.pack(side="right", fill="y")
        self.lista_box.config(yscrollcommand=scroll.set)

        # Checkbox dias úteis
        self.var_dias_uteis = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(
            card_horarios,
            text="Executar somente de Segunda a Sexta-feira",
            variable=self.var_dias_uteis,
            command=self._salvar_dias_uteis,
            bg="#1e293b",
            fg="#e2e8f0",
            selectcolor="#0f172a",
            activebackground="#1e293b",
            activeforeground="#ffffff",
            font=("Segoe UI", 9)
        )
        chk.pack(anchor="w", pady=(8, 0))

        # CARD 3: Mini-Log de Execução em Tempo Real
        card_log = ttk.Frame(self.root, style="Card.TFrame", padding=12)
        card_log.pack(fill="x", padx=20, pady=(0, 15))

        top_log = tk.Frame(card_log, bg="#1e293b")
        top_log.pack(fill="x")
        ttk.Label(top_log, text="MONITOR DE ATIVIDADE EM TEMPO REAL", font=("Segoe UI", 9, "bold"), foreground="#94a3b8").pack(side="left")

        self.lbl_log = tk.Label(card_log, text="Sistema pronto e monitorando agendamentos.", font=("Segoe UI", 9), fg="#38bdf8", bg="#1e293b", anchor="w")
        self.lbl_log.pack(fill="x", pady=(4, 0))

    def _mudar_aba(self, tarefa_id):
        self.tarefa_selecionada = tarefa_id
        if tarefa_id == "whatsapp":
            self.btn_tab_zap.config(bg="#2563eb", fg="#ffffff")
            self.btn_tab_sql.config(bg="#1e293b", fg="#94a3b8")
        else:
            self.btn_tab_zap.config(bg="#1e293b", fg="#94a3b8")
            self.btn_tab_sql.config(bg="#2563eb", fg="#ffffff")
        self._atualizar_tela_tarefa()

    def _tarefa_atual(self):
        return self.config["tarefas"].get(self.tarefa_selecionada, {})

    def _atualizar_tela_tarefa(self):
        tarefa = self._tarefa_atual()
        nome = tarefa.get("nome", "Tarefa")
        icone = tarefa.get("icone", "⚡")
        self.lbl_tarefa_nome.config(text=f"{icone} {nome.upper()}")

        # Lista de horários
        self.lista_box.delete(0, tk.END)
        for h in sorted(tarefa.get("horarios", [])):
            self.lista_box.insert(tk.END, f"  ⏰  {h}")

        # Checkbox
        self.var_dias_uteis.set(tarefa.get("apenas_dias_uteis", True))

        # Badge e botão de pausar
        if tarefa.get("ativo", True):
            self.lbl_badge.config(text="● ATIVA", bg="#065f46", fg="#34d399")
            self.btn_toggle.config(text="⏸ Pausar Esta Tarefa")
        else:
            self.lbl_badge.config(text="○ PAUSADA", bg="#78350f", fg="#fcd34d")
            self.btn_toggle.config(text="▶ Retomar Esta Tarefa")

        # Último envio
        ult_h = tarefa.get("ultimo_horario", "-")
        ult_s = tarefa.get("ultimo_status", "Pronto")
        self.lbl_ultimo.config(text=f"Última execução: {ult_h} — {ult_s}")

        # Próximo disparo
        _, txt = self._calcular_proximo_disparo(self.tarefa_selecionada)
        self.lbl_proximo.config(text=f"Próximo disparo: {txt}")

    def _adicionar_horario(self):
        val = self.entry_hora.get().strip()
        try:
            time.strptime(val, "%H:%M")
        except ValueError:
            messagebox.showerror("Formato Inválido", "Por favor digite no formato HH:MM (ex: 09:30)")
            return

        tarefa = self._tarefa_atual()
        if val in tarefa.get("horarios", []):
            messagebox.showinfo("Aviso", f"O horário {val} já está cadastrado nesta tarefa.")
            return

        tarefa["horarios"].append(val)
        tarefa["horarios"] = sorted(tarefa["horarios"])
        salvar_config(self.config)
        self._atualizar_tela_tarefa()

    def _remover_horario(self):
        sel = self.lista_box.curselection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione um horário na lista para remover.")
            return
        idx = sel[0]
        tarefa = self._tarefa_atual()
        del tarefa["horarios"][idx]
        salvar_config(self.config)
        self._atualizar_tela_tarefa()

    def _salvar_dias_uteis(self):
        tarefa = self._tarefa_atual()
        tarefa["apenas_dias_uteis"] = self.var_dias_uteis.get()
        salvar_config(self.config)
        self._atualizar_tela_tarefa()

    def _toggle_ativo(self):
        tarefa = self._tarefa_atual()
        tarefa["ativo"] = not tarefa.get("ativo", True)
        salvar_config(self.config)
        self._atualizar_tela_tarefa()

    def _calcular_proximo_disparo(self, tarefa_id):
        tarefa = self.config["tarefas"].get(tarefa_id, {})
        if not tarefa.get("ativo", True):
            return None, "Tarefa pausada"

        horarios = tarefa.get("horarios", [])
        if not horarios:
            return None, "Nenhum horário cadastrado"

        agora = datetime.now()
        for d in range(8):
            dia_candidato = agora + timedelta(days=d)
            if tarefa.get("apenas_dias_uteis", True) and dia_candidato.weekday() >= 5:
                continue

            for h_str in sorted(horarios):
                hora, minuto = map(int, h_str.split(":"))
                dt_candidato = dia_candidato.replace(hour=hora, minute=minuto, second=0, microsecond=0)
                if dt_candidato > agora:
                    delta = dt_candidato - agora
                    horas_rest, resto = divmod(int(delta.total_seconds()), 3600)
                    mins_rest, _ = divmod(resto, 60)
                    if d == 0:
                        txt_quando = f"Hoje às {h_str} (em {horas_rest}h {mins_rest}m)"
                    elif d == 1:
                        txt_quando = f"Amanhã às {h_str}"
                    else:
                        dias_sem = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
                        txt_quando = f"{dias_sem[dt_candidato.weekday()]} às {h_str}"
                    return dt_candidato, txt_quando

        return None, "Sem horários próximos"

    def _executar_agora_click(self):
        if self.executando_tarefa_id:
            messagebox.showwarning("Aviso", "Já existe uma automação em execução neste momento. Aguarde finalizar.")
            return
        tarefa_id = self.tarefa_selecionada
        threading.Thread(target=self._executar_tarefa, args=(tarefa_id, "Manual"), daemon=True).start()

    def _executar_tarefa(self, tarefa_id, motivo="Agendado"):
        if self.executando_tarefa_id:
            return
        self.executando_tarefa_id = tarefa_id
        self.btn_exec_agora.config(state="disabled")

        tarefa = self.config["tarefas"].get(tarefa_id, {})
        script_nome = tarefa.get("script", "")
        script_path = os.path.join(ROOT_DIR, script_nome)
        if not os.path.exists(script_path):
            script_path = os.path.join(APP_DIR, script_nome)
        if not os.path.exists(script_path):
            self._log_msg(f"❌ Script não encontrado: {script_nome}")
            tarefa["ultimo_status"] = "Arquivo não encontrado ❌"
            salvar_config(self.config)
            return

        icone = tarefa.get("icone", "⚡")
        nome_curto = tarefa.get("nome", tarefa_id)

        hora_inicio = datetime.now().strftime("%H:%M:%S")
        self._log_msg(f"[{hora_inicio}] ⏳ Iniciando {icone} {nome_curto} ({motivo})...")

        # Garante que a VPN está ativa antes de executar qualquer tarefa
        if not vpn_esta_conectada():
            self._log_msg(f"[{hora_inicio}] 🛡️ Conectando VPN Engelmig antes de iniciar...")
            conectar_vpn_automatica()

        try:
            cmd = [PYTHON_EXE, script_path]
            self.processo_atual = subprocess.Popen(
                cmd,
                cwd=ROOT_DIR,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace"
            )

            linhas_log = []
            for linha in self.processo_atual.stdout:
                linha = linha.strip()
                if linha:
                    linhas_log.append(linha)
                    self._log_msg(f"{icone} {linha}")

            self.processo_atual.wait()
            sucesso = (self.processo_atual.returncode == 0)

            hora_fim = datetime.now().strftime("%H:%M:%S")
            if sucesso:
                msg_status = "Sucesso ✅"
                self._log_msg(f"[{hora_fim}] {icone} {nome_curto} concluído com sucesso!")
            else:
                msg_status = f"Falha (Código {self.processo_atual.returncode}) ❌"
                ult_erro = linhas_log[-1] if linhas_log else "Verifique o log"
                self._log_msg(f"[{hora_fim}] Erro: {ult_erro[:75]}")

            tarefa["ultimo_status"] = msg_status
            tarefa["ultimo_horario"] = hora_fim
            salvar_config(self.config)

            self.root.after(0, self._atualizar_tela_tarefa)

        except Exception as e:
            self._log_msg(f"Erro na execução de {nome_curto}: {e}")
        finally:
            self.executando_tarefa_id = None
            self.btn_exec_agora.config(state="normal")
            self.root.after(0, self._atualizar_tela_tarefa)

    def _log_msg(self, texto):
        self.root.after(0, lambda: self.lbl_log.config(text=texto[:95]))

    def _conectar_vpn_click(self):
        self._log_msg("Iniciando conexão automática da VPN Engelmig...")
        threading.Thread(target=self._exec_conectar_vpn, daemon=True).start()

    def _exec_conectar_vpn(self):
        conectar_vpn_automatica()
        self.root.after(0, self._atualizar_status_vpn_ui)

    def _atualizar_status_vpn_ui(self):
        if vpn_esta_conectada():
            self.btn_vpn_status.config(text="🛡️ VPN CONECTADA", bg="#065f46", fg="#34d399")
        else:
            self.btn_vpn_status.config(text="⚠️ CONECTAR VPN", bg="#991b1b", fg="#ffffff")

    def _motor_agendador(self):
        """Thread que monitora o relógio, VPN e dispara cada tarefa no minuto exato."""
        # Ao abrir o aplicativo, se a VPN estiver desconectada, conecta automaticamente
        if not vpn_esta_conectada():
            self._log_msg("Iniciando conexão automática com FortiClient VPN...")
            conectar_vpn_automatica()
        self.root.after(0, self._atualizar_status_vpn_ui)

        ultimo_disparo = {}
        contador_vpn = 0
        ultimo_mtime = 0
        try:
            if os.path.exists(CONFIG_FILE):
                ultimo_mtime = os.path.getmtime(CONFIG_FILE)
        except Exception:
            pass

        while self.loop_ativo:
            time.sleep(2)
            contador_vpn += 1
            if contador_vpn >= 15:  # Atualiza badge de VPN a cada 30s
                contador_vpn = 0
                self.root.after(0, self._atualizar_status_vpn_ui)

            # Recarrega configurações do disco se o arquivo foi alterado externamente
            try:
                if os.path.exists(CONFIG_FILE):
                    mtime_atual = os.path.getmtime(CONFIG_FILE)
                    if mtime_atual > ultimo_mtime:
                        ultimo_mtime = mtime_atual
                        self.config = carregar_config()
                        self.root.after(0, self._atualizar_tela_tarefa)
            except Exception:
                pass

            agora = datetime.now()
            agora_hm = agora.strftime("%H:%M")

            for t_id, tarefa in self.config.get("tarefas", {}).items():
                if not tarefa.get("ativo", True):
                    continue

                if tarefa.get("apenas_dias_uteis", True) and agora.weekday() >= 5:
                    continue

                if agora_hm in tarefa.get("horarios", []) and ultimo_disparo.get(t_id) != agora_hm:
                    if self.executando_tarefa_id is None:
                        ultimo_disparo[t_id] = agora_hm
                        threading.Thread(target=self._executar_tarefa, args=(t_id, f"Agendado {agora_hm}"), daemon=True).start()

            if agora.second == 0:
                self.root.after(0, self._atualizar_tela_tarefa)

    def _ao_fechar(self):
        resp = messagebox.askyesno(
            "Fechar Central de Automações",
            "Deseja fechar a Central de Automações?\n\n"
            "• Clique em NÃO para MINIMIZAR (o programa continuará rodando os agendamentos).\n"
            "• Clique em SIM para ENCERRAR completamente todas as automações."
        )
        if resp:
            self.loop_ativo = False
            self.root.destroy()
        else:
            self.root.iconify()


def main():
    root = tk.Tk()
    app = CentralAutomacoesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
