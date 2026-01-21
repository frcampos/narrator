#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTX Narrator - Aplicação de Narração de Apresentações
Requisitos: pip install customtkinter edge-tts python-pptx pydub pillow

Criado por Fernando Rui Campos
"""

# Versão da aplicação
APP_VERSAO = "1.9.3"
APP_DATA = "2026-01-21"

import os
import sys
import json
import threading
import tempfile
from pathlib import Path
from typing import Optional, List
from datetime import datetime

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
except ImportError:
    os.system(f"{sys.executable} -m pip install customtkinter")
    import customtkinter as ctk
    from tkinter import filedialog, messagebox

from idiomas import GestorIdioma, IDIOMAS_DISPONIVEIS, IDIOMAS_PRINCIPAIS, IDIOMAS_ADICIONAIS, VOZES_EDGE
from tts_engine import MotorTTS, ConfigTTS, DESCRICOES_MOTORES
from pptx_handler import GestorPPTX, ConfigIcone, PPTX_DISPONIVEL
from video_generator import GeradorVideo
from tradutor import Tradutor, GOOGLE_DISPONIVEL, ARGOS_DISPONIVEL
from html_exporter import HTMLExporter, ConfigExportHTML, ConfigLayoutHTML, ConfigAudioHTML, ConfigKaraokeHTML, ConfigNavigationHTML, ConfigAppearanceHTML, SlideDataHTML

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")
CONFIG_FILE = "pptx_narrator_config.json"

# Perfis de configuração predefinidos
PERFIS_CONFIG = {
    "padrao": {
        "nome": "Padrão",
        "descricao": "Configuração equilibrada para uso geral",
        "motor": "edge",
        "resolucao": "1280x720",
        "fps": "12",
        "tempo_extra": "0.5"
    },
    "qualidade": {
        "nome": "Alta Qualidade",
        "descricao": "Melhor qualidade, ficheiros maiores e mais lentos",
        "motor": "edge",
        "resolucao": "1920x1080",
        "fps": "24",
        "tempo_extra": "1.0"
    },
    "rapido": {
        "nome": "Rápido",
        "descricao": "Processamento mais rápido, ficheiros menores",
        "motor": "edge",
        "resolucao": "854x480",
        "fps": "6",
        "tempo_extra": "0.5"
    },
    "offline": {
        "nome": "Offline",
        "descricao": "Funciona sem internet (requer Piper TTS instalado)",
        "motor": "piper",
        "resolucao": "1280x720",
        "fps": "12",
        "tempo_extra": "0.5"
    }
}


class Tooltip:
    """Tooltip simples para widgets"""
    def __init__(self, widget, texto):
        self.widget = widget
        self.texto = texto
        self.tooltip = None
        widget.bind("<Enter>", self.mostrar)
        widget.bind("<Leave>", self.ocultar)
    
    def mostrar(self, event=None):
        x, y = self.widget.winfo_rootx() + 25, self.widget.winfo_rooty() - 10
        self.tooltip = ctk.CTkToplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        self.tooltip.attributes("-topmost", True)
        frame = ctk.CTkFrame(self.tooltip, corner_radius=8, 
                             fg_color=("gray90", "gray15"),
                             border_width=1, border_color=("gray70", "gray30"))
        frame.pack(padx=2, pady=2)
        label = ctk.CTkLabel(frame, text=self.texto,
                             wraplength=280, justify="left",
                             font=ctk.CTkFont(size=12))
        label.pack(padx=10, pady=8)
    
    def ocultar(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


def criar_label_com_ajuda(parent, texto, ajuda, **kwargs):
    """Cria label com botão de ajuda (?) visível"""
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    ctk.CTkLabel(frame, text=texto, **kwargs).pack(side="left")
    btn_ajuda = ctk.CTkButton(frame, text="?", width=22, height=22,
                               corner_radius=11, 
                               fg_color=("gray75", "gray35"),
                               hover_color=("gray65", "gray45"),
                               text_color=("gray20", "gray90"),
                               font=ctk.CTkFont(size=11, weight="bold"))
    btn_ajuda.pack(side="left", padx=(5, 0))
    Tooltip(btn_ajuda, ajuda)
    return frame


class PPTXNarratorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.idioma = GestorIdioma("pt-PT")
        self.tts = MotorTTS()
        self.tts_trad = MotorTTS()  # Motor TTS para tradução
        self.pptx = GestorPPTX()
        self.video = GeradorVideo()
        self.tradutor = Tradutor()
        self.slide_atual = 1
        self.a_processar = False
        self._pasta_saida = ""
        self._pasta_saida_html5 = ""  # v1.9.1: Pasta específica para HTML5
        self._ultimo_ficheiro = ""  # v1.9.3: Último ficheiro gerado
        self._usar_pasta_estruturada = True  # v1.9.3: Organizar em subpastas
        
        # v1.7.2: Inicializar TTS de tradução com idioma inglês por defeito
        self._inicializar_tts_traducao("en")
        
        self.carregar_config()
        self.title(f"{self.idioma.t('app_titulo')} v{APP_VERSAO}")
        self.geometry("1100x750")
        self.minsize(900, 600)
        
        self._criar_menu()
        self._criar_interface()
        self._atualizar_estado_botoes()
    
    def _inicializar_tts_traducao(self, codigo_idioma: str):
        """Inicializa o TTS de tradução com o idioma especificado - v1.7.2"""
        self.tts_trad.config.idioma = codigo_idioma
        vozes = VOZES_EDGE.get(codigo_idioma, [])
        if vozes:
            self.tts_trad.config.voz = vozes[0][0]  # Primeira voz disponível
        self.tradutor.config.idioma_destino = codigo_idioma
    
    def carregar_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                self.idioma.definir_idioma(cfg.get("idioma", "pt-PT"))
                self.tts.config.motor = cfg.get("motor_tts", "edge")
                self.tts.config.voz = cfg.get("voz", "pt-PT-RaquelNeural")
                self.tts.config.velocidade = cfg.get("velocidade", 1.0)
                self.tts_trad.config.velocidade = cfg.get("velocidade_trad", 1.0)
                self.tts.config.idioma = cfg.get("idioma_voz", "pt-PT")
                self.pptx.config_icone.mostrar = cfg.get("mostrar_icone", True)
                self.pptx.config_icone.posicao = cfg.get("posicao_icone", "sup_dir")
                self.pptx.config_icone.tamanho_cm = cfg.get("tamanho_icone", 1.0)
                self.video.config.tempo_extra_slide = cfg.get("tempo_extra_slide", 5.0)
                self._pasta_saida = cfg.get("pasta_saida", "")
                self._pasta_saida_html5 = cfg.get("pasta_saida_html5", "")
                self._usar_pasta_estruturada = cfg.get("usar_pasta_estruturada", True)  # v1.9.3
        except: pass
    
    def guardar_config(self):
        cfg = {
            "idioma": self.idioma.idioma_atual,
            "motor_tts": self.tts.config.motor,
            "voz": self.tts.config.voz,
            "velocidade": self.tts.config.velocidade,
            "velocidade_trad": self.tts_trad.config.velocidade,
            "idioma_voz": self.tts.config.idioma,
            "mostrar_icone": self.pptx.config_icone.mostrar,
            "posicao_icone": self.pptx.config_icone.posicao,
            "tamanho_icone": self.pptx.config_icone.tamanho_cm,
            "tempo_extra_slide": self.video.config.tempo_extra_slide,
            "pasta_saida": self._pasta_saida,
            "pasta_saida_html5": getattr(self, '_pasta_saida_html5', ''),
            "usar_pasta_estruturada": getattr(self, '_usar_pasta_estruturada', True),  # v1.9.3
        }
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, indent=2)
        except: pass
    
    def _criar_menu(self):
        self.menu_bar = ctk.CTkFrame(self, height=35)
        self.menu_bar.pack(fill="x", padx=5, pady=(5, 0))
        
        ctk.CTkButton(self.menu_bar, text="📂 " + self.idioma.t("menu_abrir"), 
                      width=120, height=28, command=self.abrir_pptx).pack(side="left", padx=2)
        ctk.CTkButton(self.menu_bar, text="💾 Guardar Config", 
                      width=120, height=28, command=lambda: [self.guardar_config(), 
                      messagebox.showinfo("", "Configurações guardadas!")]).pack(side="left", padx=2)
        ctk.CTkButton(self.menu_bar, text="🌐 " + self.idioma.t("menu_idioma"), 
                      width=100, height=28, command=self._menu_idioma).pack(side="left", padx=2)
        ctk.CTkButton(self.menu_bar, text="❓ Sobre", 
                      width=80, height=28, command=self._mostrar_sobre).pack(side="left", padx=2)
        ctk.CTkButton(self.menu_bar, text="🚪 Sair", 
                      width=80, height=28, command=self._confirmar_saida).pack(side="left", padx=2)
        
        # Interceptar fecho da janela
        self.protocol("WM_DELETE_WINDOW", self._confirmar_saida)
    
    def _confirmar_saida(self):
        """Confirma saída da aplicação"""
        resposta = messagebox.askquestion(
            "Confirmar Saída",
            "Deseja sair da aplicação?\n\nAs configurações serão guardadas automaticamente.",
            icon='question'
        )
        if resposta == 'yes':
            self.guardar_config()
            self.quit()
            self.destroy()
    
    def _menu_idioma(self):
        menu = ctk.CTkToplevel(self)
        menu.title(self.idioma.t("menu_idioma"))
        menu.geometry("250x180")
        menu.transient(self)
        menu.grab_set()
        menu.geometry(f"+{self.winfo_x()+100}+{self.winfo_y()+70}")
        for codigo, nome in IDIOMAS_DISPONIVEIS.items():
            ctk.CTkButton(menu, text=nome, command=lambda c=codigo, m=menu: 
                          self._mudar_idioma(c, m)).pack(fill="x", padx=10, pady=3)
    
    def _mudar_idioma(self, codigo: str, menu):
        menu.destroy()
        self.idioma.definir_idioma(codigo)
        self.tts.config.idioma = codigo
        self.guardar_config()
        messagebox.showinfo("", "Idioma alterado. Reinicie para aplicar.")
    
    def _mostrar_sobre(self):
        messagebox.showinfo(self.idioma.t("sobre_titulo"), self.idioma.t("sobre_texto"))
    
    def _criar_interface(self):
        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tabview = ctk.CTkTabview(self.main_container)
        self.tabview.pack(fill="both", expand=True)
        
        self.tab_slides = self.tabview.add(self.idioma.t("tab_slides"))
        self.tab_config = self.tabview.add(self.idioma.t("tab_configuracoes"))
        self.tab_html5 = self.tabview.add("HTML5")  # Nova tab
        self.tab_progresso = self.tabview.add(self.idioma.t("tab_progresso"))
        
        self._criar_tab_slides()
        self._criar_tab_config()
        self._criar_tab_html5()  # Nova
        self._criar_tab_progresso()
    
    def _criar_tab_slides(self):
        # Botões
        frame_botoes = ctk.CTkFrame(self.tab_slides)
        frame_botoes.pack(fill="x", pady=(0, 10))
        
        self.btn_abrir = ctk.CTkButton(frame_botoes, text=self.idioma.t("btn_abrir_pptx"),
                                        command=self.abrir_pptx, width=180)
        self.btn_abrir.pack(side="left", padx=5, pady=5)
        
        self.btn_gerar_audio = ctk.CTkButton(frame_botoes, text=self.idioma.t("btn_gerar_audio"),
                                              command=self.gerar_todos_audios, width=150)
        self.btn_gerar_audio.pack(side="left", padx=5, pady=5)
        
        self.btn_gerar_pptx = ctk.CTkButton(frame_botoes, text=self.idioma.t("btn_gerar_pptx"),
                                             command=self.guardar_pptx_audio, width=180)
        self.btn_gerar_pptx.pack(side="left", padx=5, pady=5)
        
        self.btn_gerar_video = ctk.CTkButton(frame_botoes, text=self.idioma.t("btn_gerar_video"),
                                              command=self.gerar_video, width=150)
        self.btn_gerar_video.pack(side="left", padx=5, pady=5)
        
        # Frame principal
        frame_principal = ctk.CTkFrame(self.tab_slides, fg_color="transparent")
        frame_principal.pack(fill="both", expand=True)
        
        # Lista de slides
        frame_lista = ctk.CTkFrame(frame_principal, width=200)
        frame_lista.pack(side="left", fill="y", padx=(0, 10))
        frame_lista.pack_propagate(False)
        
        ctk.CTkLabel(frame_lista, text="Slides", font=ctk.CTkFont(size=14, weight="bold")).pack(pady=5)
        
        self.frame_lista_slides = ctk.CTkScrollableFrame(frame_lista)
        self.frame_lista_slides.pack(fill="both", expand=True, padx=5, pady=5)
        self.lista_slides_btns = []
        
        self.lbl_nenhum = ctk.CTkLabel(self.frame_lista_slides, text=self.idioma.t("lbl_nenhum_pptx"),
                                        text_color="gray", wraplength=170)
        self.lbl_nenhum.pack(pady=50)
        
        # Editor do slide
        frame_editor = ctk.CTkFrame(frame_principal)
        frame_editor.pack(side="left", fill="both", expand=True)
        
        self.lbl_slide_titulo = ctk.CTkLabel(frame_editor, text=f"{self.idioma.t('lbl_slide')} 1",
                                              font=ctk.CTkFont(size=16, weight="bold"))
        self.lbl_slide_titulo.pack(pady=10)
        
        ctk.CTkLabel(frame_editor, text=self.idioma.t("lbl_texto_slide")).pack(anchor="w", padx=10)
        self.txt_slide = ctk.CTkTextbox(frame_editor, height=80, state="disabled")
        self.txt_slide.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(frame_editor, text=self.idioma.t("lbl_notas_originais")).pack(anchor="w", padx=10)
        self.txt_notas = ctk.CTkTextbox(frame_editor, height=80, state="disabled")
        self.txt_notas.pack(fill="x", padx=10, pady=(0, 10))
        
        ctk.CTkLabel(frame_editor, text=self.idioma.t("lbl_texto_narrar"),
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10)
        self.txt_narrar = ctk.CTkTextbox(frame_editor, height=90)
        self.txt_narrar.pack(fill="x", padx=10, pady=(0, 5))
        self.txt_narrar.bind("<FocusOut>", self._guardar_texto_slide)
        
        # Botões de áudio original
        frame_audio = ctk.CTkFrame(frame_editor, fg_color="transparent")
        frame_audio.pack(fill="x", padx=10, pady=2)
        
        self.btn_preview = ctk.CTkButton(frame_audio, text=self.idioma.t("btn_preview"),
                                          command=self.preview_audio, width=140)
        self.btn_preview.pack(side="left", padx=5)
        
        self.btn_parar = ctk.CTkButton(frame_audio, text=self.idioma.t("btn_parar"),
                                        command=self.parar_audio, width=100)
        self.btn_parar.pack(side="left", padx=5)
        
        # Texto traduzido
        ctk.CTkLabel(frame_editor, text=self.idioma.t("lbl_texto_traduzido"),
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 0))
        self.txt_traduzido = ctk.CTkTextbox(frame_editor, height=90)
        self.txt_traduzido.pack(fill="x", padx=10, pady=(0, 5))
        self.txt_traduzido.bind("<FocusOut>", self._guardar_texto_traduzido)
        
        # Botões de áudio traduzido
        frame_audio_trad = ctk.CTkFrame(frame_editor, fg_color="transparent")
        frame_audio_trad.pack(fill="x", padx=10, pady=2)
        
        self.btn_preview_trad = ctk.CTkButton(frame_audio_trad, text="▶️ Preview Tradução",
                                               command=self.preview_audio_traduzido, width=160)
        self.btn_preview_trad.pack(side="left", padx=5)
        
        self.btn_parar_trad = ctk.CTkButton(frame_audio_trad, text="⏹️ Parar",
                                             command=self.parar_audio_traduzido, width=100)
        self.btn_parar_trad.pack(side="left", padx=5)
        
        self.lbl_status_slide = ctk.CTkLabel(frame_editor, text="", text_color="green")
        self.lbl_status_slide.pack(pady=5)
    
    def _criar_tab_config(self):
        scroll = ctk.CTkScrollableFrame(self.tab_config)
        scroll.pack(fill="both", expand=True)
        
        # Frame Perfil de Configuração
        frame_perfil = ctk.CTkFrame(scroll)
        frame_perfil.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_perfil, text="⚡ Configuração Rápida",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fp1 = ctk.CTkFrame(frame_perfil, fg_color="transparent")
        fp1.pack(fill="x", padx=10, pady=5)
        lbl_perfil = criar_label_com_ajuda(fp1, "Perfil:",
            "Escolha um perfil predefinido para aplicar\n"
            "automaticamente as configurações recomendadas.")
        lbl_perfil.pack(side="left")
        
        perfis_nomes = [p["nome"] for p in PERFIS_CONFIG.values()]
        self.combo_perfil = ctk.CTkComboBox(fp1, values=perfis_nomes, width=180)
        self.combo_perfil.set("Padrão")
        self.combo_perfil.pack(side="left", padx=10)
        
        ctk.CTkButton(fp1, text="Aplicar Perfil", width=120,
                      command=self._aplicar_perfil).pack(side="left", padx=5)
        
        self.lbl_desc_perfil = ctk.CTkLabel(frame_perfil, 
            text="Configuração equilibrada para uso geral",
            text_color="gray50", font=ctk.CTkFont(size=11))
        self.lbl_desc_perfil.pack(anchor="w", padx=15, pady=(0, 5))
        
        # Atualizar descrição ao mudar perfil
        self.combo_perfil.configure(command=self._ao_mudar_perfil)
        
        # Frame Voz
        frame_voz = ctk.CTkFrame(scroll)
        frame_voz.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_voz, text=self.idioma.t("frame_voz"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        # Motor TTS com tooltip
        f1 = ctk.CTkFrame(frame_voz, fg_color="transparent")
        f1.pack(fill="x", padx=10, pady=5)
        lbl_motor = criar_label_com_ajuda(f1, self.idioma.t("lbl_motor_tts"),
            "Edge TTS: Vozes naturais da Microsoft (requer internet)\n"
            "Piper TTS: Vozes neurais offline de alta qualidade\n"
            "Sistema: Vozes básicas do sistema operativo")
        lbl_motor.pack(side="left")
        
        # Obter motores disponíveis
        motores_disp = MotorTTS.obter_motores_disponiveis()
        nomes_motores = [m[1] for m in motores_disp]
        self._mapa_motores = {m[1]: m[0] for m in motores_disp}
        
        self.combo_motor = ctk.CTkComboBox(f1, values=nomes_motores,
                                            command=self._ao_mudar_motor, width=250)
        # Selecionar motor atual
        for nome, codigo in self._mapa_motores.items():
            if codigo == self.tts.config.motor:
                self.combo_motor.set(nome)
                break
        self.combo_motor.pack(side="right")
        
        # Idioma da voz com tooltip
        f2 = ctk.CTkFrame(frame_voz, fg_color="transparent")
        f2.pack(fill="x", padx=10, pady=5)
        lbl_idioma = criar_label_com_ajuda(f2, self.idioma.t("lbl_idioma_voz"),
            "Idioma das vozes disponíveis para narração.\n"
            "Pode ser diferente do idioma da interface.")
        lbl_idioma.pack(side="left")
        
        # v1.8: Variável para controlar línguas visíveis
        self.var_mais_linguas_voz = ctk.BooleanVar(value=False)
        self._idiomas_voz_visiveis = list(IDIOMAS_PRINCIPAIS.keys())
        
        self.combo_idioma_voz = ctk.CTkComboBox(f2, values=self._idiomas_voz_visiveis,
                                                 command=self._ao_mudar_idioma_voz, width=200)
        self.combo_idioma_voz.set(self.tts.config.idioma)
        self.combo_idioma_voz.pack(side="right")
        
        # Checkbox "Mais línguas"
        self.chk_mais_linguas_voz = ctk.CTkCheckBox(f2, text="+ Línguas", width=80,
                                                      variable=self.var_mais_linguas_voz,
                                                      command=self._ao_mudar_mais_linguas_voz)
        self.chk_mais_linguas_voz.pack(side="right", padx=5)
        
        # Voz
        f3 = ctk.CTkFrame(frame_voz, fg_color="transparent")
        f3.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f3, text=self.idioma.t("lbl_voz")).pack(side="left")
        self.combo_voz = ctk.CTkComboBox(f3, values=self._obter_lista_vozes(),
                                          command=self._ao_mudar_voz, width=250)
        self.combo_voz.pack(side="right")
        
        # Velocidade com tooltip
        f4 = ctk.CTkFrame(frame_voz, fg_color="transparent")
        f4.pack(fill="x", padx=10, pady=5)
        lbl_vel = criar_label_com_ajuda(f4, self.idioma.t("lbl_velocidade"),
            "1.0x = velocidade normal\n"
            "< 1.0 = mais lento (melhor compreensão)\n"
            "> 1.0 = mais rápido")
        lbl_vel.pack(side="left")
        # Entry para valor preciso de velocidade
        self.var_vel = ctk.StringVar(value=f"{self.tts.config.velocidade:.2f}")
        self.entry_vel = ctk.CTkEntry(f4, textvariable=self.var_vel, width=55, justify="center")
        self.entry_vel.pack(side="right", padx=5)
        self.entry_vel.bind("<Return>", self._ao_digitar_velocidade)
        self.entry_vel.bind("<FocusOut>", self._ao_digitar_velocidade)
        ctk.CTkLabel(f4, text="x").pack(side="right")
        self.slider_vel = ctk.CTkSlider(f4, from_=0.5, to=2.0, number_of_steps=150, command=self._ao_mudar_velocidade, width=180)
        self.slider_vel.set(self.tts.config.velocidade)
        self.slider_vel.pack(side="right")
        
        # Frame Ícone
        frame_icone = ctk.CTkFrame(scroll)
        frame_icone.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_icone, text=self.idioma.t("frame_icone"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        f5 = ctk.CTkFrame(frame_icone, fg_color="transparent")
        f5.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f5, text=self.idioma.t("lbl_mostrar_icone")).pack(side="left")
        self.var_mostrar = ctk.BooleanVar(value=self.pptx.config_icone.mostrar)
        ctk.CTkCheckBox(f5, text="", variable=self.var_mostrar,
                        command=lambda: setattr(self.pptx.config_icone, 'mostrar', self.var_mostrar.get())).pack(side="right")
        
        f6 = ctk.CTkFrame(frame_icone, fg_color="transparent")
        f6.pack(fill="x", padx=10, pady=5)
        lbl_posicao = criar_label_com_ajuda(f6, self.idioma.t("lbl_posicao"),
            "Cantos: Superior/Inferior + Direito/Esquerdo\n"
            "Centro: Posições centrais nas margens")
        lbl_posicao.pack(side="left")
        self.combo_pos = ctk.CTkComboBox(f6, values=[
            "Superior Direito", "Superior Esquerdo", 
            "Inferior Direito", "Inferior Esquerdo",
            "Centro Superior", "Centro Inferior",
            "Centro Esquerdo", "Centro Direito"
        ], command=self._ao_mudar_posicao, width=200)
        self.combo_pos.set("Superior Direito")
        self.combo_pos.pack(side="right")
        
        # Frame Tradução
        frame_traducao = ctk.CTkFrame(scroll)
        frame_traducao.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_traducao, text=self.idioma.t("frame_traducao"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        ft1 = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft1.pack(fill="x", padx=10, pady=5)
        lbl_trad = criar_label_com_ajuda(ft1, self.idioma.t("lbl_traducao_ativa"),
            "Permite criar versões traduzidas da narração.\n"
            "Útil para materiais multilingues.")
        lbl_trad.pack(side="left")
        self.var_traducao_ativa = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(ft1, text="", variable=self.var_traducao_ativa,
                        command=self._ao_mudar_traducao_ativa).pack(side="right")
        
        ft2 = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft2.pack(fill="x", padx=10, pady=5)
        lbl_motor_trad = criar_label_com_ajuda(ft2, self.idioma.t("lbl_motor_traducao"),
            "Google Translate: Melhor qualidade (requer internet)\n"
            "Argos: Funciona offline mas qualidade inferior")
        lbl_motor_trad.pack(side="left")
        motores = []
        if GOOGLE_DISPONIVEL:
            motores.append("Google Translate")
        if ARGOS_DISPONIVEL:
            motores.append("Argos (Offline)")
        if not motores:
            motores = ["Nenhum disponível"]
        self.combo_motor_trad = ctk.CTkComboBox(ft2, values=motores, width=200)
        self.combo_motor_trad.pack(side="right")
        
        ft3 = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft3.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(ft3, text=self.idioma.t("lbl_idioma_destino")).pack(side="left")
        
        # v1.8: Variável para controlar línguas visíveis no destino
        self.var_mais_linguas_dest = ctk.BooleanVar(value=False)
        self._idiomas_dest_visiveis = list(IDIOMAS_PRINCIPAIS.values())
        
        self.combo_idioma_destino = ctk.CTkComboBox(ft3, values=self._idiomas_dest_visiveis,
                                                     command=self._ao_mudar_idioma_destino, width=160)
        self.combo_idioma_destino.set("English (US)")
        self.combo_idioma_destino.pack(side="right")
        
        # Checkbox "Mais línguas" para destino
        self.chk_mais_linguas_dest = ctk.CTkCheckBox(ft3, text="+ Línguas", width=80,
                                                       variable=self.var_mais_linguas_dest,
                                                       command=self._ao_mudar_mais_linguas_dest)
        self.chk_mais_linguas_dest.pack(side="right", padx=5)
        
        # === VOZ PARA LÍNGUA TRADUZIDA (NOVO v1.7) ===
        ft3b = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft3b.pack(fill="x", padx=10, pady=5)
        lbl_voz_trad = criar_label_com_ajuda(ft3b, "Voz tradução:",
            "Voz a usar para gerar áudio na língua traduzida.\n"
            "Automaticamente atualizada quando muda o idioma destino.")
        lbl_voz_trad.pack(side="left")
        # Obter vozes para o idioma de tradução atual
        vozes_trad = VOZES_EDGE.get(self.tts_trad.config.idioma, VOZES_EDGE.get("en", []))
        nomes_vozes_trad = [v[1] for v in vozes_trad] if vozes_trad else ["N/A"]
        self.combo_voz_trad = ctk.CTkComboBox(ft3b, values=nomes_vozes_trad,
                                               command=self._ao_mudar_voz_trad, width=200)
        if nomes_vozes_trad:
            self.combo_voz_trad.set(nomes_vozes_trad[0])
        self.combo_voz_trad.pack(side="right")
        
        # Velocidade da voz traduzida
        ft3c = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft3c.pack(fill="x", padx=10, pady=5)
        lbl_vel_trad = criar_label_com_ajuda(ft3c, "Velocidade tradução:",
            "Velocidade da narração na língua traduzida.\n"
            "Pode ser diferente da velocidade original.")
        lbl_vel_trad.pack(side="left")
        # Entry para valor preciso de velocidade traducao
        self.var_vel_trad = ctk.StringVar(value=f"{self.tts_trad.config.velocidade:.2f}")
        self.entry_vel_trad = ctk.CTkEntry(ft3c, textvariable=self.var_vel_trad, width=55, justify="center")
        self.entry_vel_trad.pack(side="right", padx=5)
        self.entry_vel_trad.bind("<Return>", self._ao_digitar_velocidade_trad)
        self.entry_vel_trad.bind("<FocusOut>", self._ao_digitar_velocidade_trad)
        ctk.CTkLabel(ft3c, text="x").pack(side="right")
        self.slider_vel_trad = ctk.CTkSlider(ft3c, from_=0.5, to=2.0, width=130, number_of_steps=150, command=self._ao_mudar_velocidade_trad)
        self.slider_vel_trad.set(self.tts_trad.config.velocidade)
        self.slider_vel_trad.pack(side="right")
        
        ft4 = ctk.CTkFrame(frame_traducao, fg_color="transparent")
        ft4.pack(fill="x", padx=10, pady=5)
        self.btn_traduzir_todos = ctk.CTkButton(ft4, text=self.idioma.t("btn_traduzir_todos"),
                                                 command=self.traduzir_todos_slides, width=180)
        self.btn_traduzir_todos.pack(side="left", padx=5)
        self.btn_gerar_audio_trad = ctk.CTkButton(ft4, text=self.idioma.t("btn_gerar_audio_trad"),
                                                   command=self.gerar_todos_audios_traduzidos, width=200)
        self.btn_gerar_audio_trad.pack(side="left", padx=5)
        
        # Frame Saída
        frame_saida = ctk.CTkFrame(scroll)
        frame_saida.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_saida, text=self.idioma.t("frame_saida"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        f7 = ctk.CTkFrame(frame_saida, fg_color="transparent")
        f7.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f7, text=self.idioma.t("lbl_audio_junto_pptx")).pack(side="left")
        self.var_audio_junto = ctk.BooleanVar(value=getattr(self, '_audio_junto_pptx', True))
        ctk.CTkCheckBox(f7, text="", variable=self.var_audio_junto,
                        command=lambda: setattr(self, '_audio_junto_pptx', self.var_audio_junto.get())).pack(side="right")
        
        f8 = ctk.CTkFrame(frame_saida, fg_color="transparent")
        f8.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f8, text=self.idioma.t("lbl_pasta_saida")).pack(side="left")
        ctk.CTkButton(f8, text="Escolher...", command=self._escolher_pasta, width=100).pack(side="right")
        self.lbl_pasta = ctk.CTkLabel(f8, text=self._pasta_saida or "(automático)", text_color="gray")
        self.lbl_pasta.pack(side="right", padx=10)
        
        # Opção de confirmação antes de gerar
        f8b = ctk.CTkFrame(frame_saida, fg_color="transparent")
        f8b.pack(fill="x", padx=10, pady=5)
        lbl_confirmar = criar_label_com_ajuda(f8b, "Confirmar antes de gerar:",
            "Se ativo, mostra um resumo das configurações\n"
            "antes de gerar PPTX ou Vídeo.\n"
            "Permite verificar e cancelar se necessário.")
        lbl_confirmar.pack(side="left")
        self.var_confirmar_antes = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(f8b, text="", variable=self.var_confirmar_antes).pack(side="right")
        
        # Frame Vídeo
        frame_video = ctk.CTkFrame(scroll)
        frame_video.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_video, text=self.idioma.t("frame_video"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        f9 = ctk.CTkFrame(frame_video, fg_color="transparent")
        f9.pack(fill="x", padx=10, pady=5)
        lbl_res = criar_label_com_ajuda(f9, self.idioma.t("lbl_resolucao"),
            "1920x1080: Alta definição (ficheiro maior, mais lento)\n"
            "1280x720: Boa qualidade, equilibrado\n"
            "854x480: Ficheiro pequeno, mais rápido")
        lbl_res.pack(side="left")
        self.var_resolucao = ctk.StringVar(value="1280x720")
        ctk.CTkComboBox(f9, values=["1920x1080", "1280x720", "854x480"], 
                        variable=self.var_resolucao, width=120).pack(side="right")
        
        f10 = ctk.CTkFrame(frame_video, fg_color="transparent")
        f10.pack(fill="x", padx=10, pady=5)
        lbl_fps = criar_label_com_ajuda(f10, self.idioma.t("lbl_fps"),
            "Imagens por segundo no vídeo.\n"
            "12: Recomendado (apresentações são estáticas)\n"
            "6: Mais rápido, ficheiro menor")
        lbl_fps.pack(side="left")
        self.var_fps = ctk.StringVar(value="12")
        ctk.CTkComboBox(f10, values=["24", "12", "6"], 
                        variable=self.var_fps, width=80).pack(side="right")
        
        f11 = ctk.CTkFrame(frame_video, fg_color="transparent")
        f11.pack(fill="x", padx=10, pady=5)
        lbl_extra = criar_label_com_ajuda(f11, self.idioma.t("lbl_tempo_extra_slide"),
            "Segundos adicionais após o áudio terminar.\n"
            "Dá tempo ao aluno para observar o slide.")
        lbl_extra.pack(side="left")
        self.var_tempo_extra = ctk.StringVar(value=str(self.video.config.tempo_extra_slide))
        ctk.CTkEntry(f11, textvariable=self.var_tempo_extra, width=80).pack(side="right")
        
        f12 = ctk.CTkFrame(frame_video, fg_color="transparent")
        f12.pack(fill="x", padx=10, pady=5)
        lbl_idioma_vid = criar_label_com_ajuda(f12, self.idioma.t("lbl_idioma_video"),
            "Qual áudio usar no vídeo:\n"
            "Original: narração na língua principal\n"
            "Traduzida: narração traduzida")
        lbl_idioma_vid.pack(side="left")
        self.var_idioma_video = ctk.StringVar(value="original")
        ctk.CTkComboBox(f12, values=[self.idioma.t("opt_lingua_original"), 
                                      self.idioma.t("opt_lingua_traduzida")],
                        variable=self.var_idioma_video, width=180).pack(side="right")
        
        # Frame Legendas
        frame_legendas = ctk.CTkFrame(scroll)
        frame_legendas.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_legendas, text=self.idioma.t("frame_legendas"),
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fl1 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl1.pack(fill="x", padx=10, pady=5)
        lbl_leg_slide = criar_label_com_ajuda(fl1, self.idioma.t("lbl_legenda_slide"),
            "Adiciona uma caixa de texto discreta no fundo\n"
            "de cada slide com o texto traduzido.\n"
            "Útil para alunos que preferem ler.")
        lbl_leg_slide.pack(side="left")
        self.var_legenda_slide = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fl1, text="", variable=self.var_legenda_slide).pack(side="right")
        
        fl2 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl2.pack(fill="x", padx=10, pady=5)
        lbl_leg_notas = criar_label_com_ajuda(fl2, self.idioma.t("lbl_legenda_notas"),
            "Adiciona a tradução nas notas do apresentador.\n"
            "Visível apenas no modo apresentador,\n"
            "não aparece na projeção.")
        lbl_leg_notas.pack(side="left")
        self.var_legenda_notas = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fl2, text="", variable=self.var_legenda_notas).pack(side="right")
        
        fl3 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl3.pack(fill="x", padx=10, pady=5)
        lbl_leg_video = criar_label_com_ajuda(fl3, self.idioma.t("lbl_legenda_video"),
            "Adiciona legendas visíveis no vídeo.\n"
            "O texto aparece na parte inferior do ecrã,\n"
            "ideal para alunos com dificuldades auditivas.")
        lbl_leg_video.pack(side="left")
        self.var_legenda_video = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fl3, text="", variable=self.var_legenda_video).pack(side="right")
        
        # Idioma das legendas no vídeo
        fl3b = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl3b.pack(fill="x", padx=10, pady=5)
        lbl_idioma_leg = criar_label_com_ajuda(fl3b, "Idioma legendas vídeo:",
            "Escolha o idioma das legendas no vídeo.\n"
            "'Igual ao áudio': legendas no mesmo idioma do áudio\n"
            "'Traduzido': legendas na língua traduzida (como subtítulos)")
        lbl_idioma_leg.pack(side="left")
        self.var_idioma_legendas = ctk.StringVar(value="Língua traduzida")
        ctk.CTkComboBox(fl3b, values=["Igual ao áudio", "Língua original", "Língua traduzida"],
                        variable=self.var_idioma_legendas, width=160).pack(side="right")
        
        # Posição legendas no vídeo
        fl3c = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl3c.pack(fill="x", padx=10, pady=5)
        lbl_pos_leg_video = criar_label_com_ajuda(fl3c, "Posição legendas vídeo:",
            "'Sobrepor slide': barra semi-transparente sobre o slide\n"
            "'Área separada': barra preta abaixo do slide")
        lbl_pos_leg_video.pack(side="left")
        self.var_pos_legenda_video = ctk.StringVar(value="Sobrepor slide")
        ctk.CTkComboBox(fl3c, values=["Sobrepor slide", "Área separada"],
                        variable=self.var_pos_legenda_video, width=160).pack(side="right")
        
        # Número de linhas por página de legenda
        fl3d = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl3d.pack(fill="x", padx=10, pady=5)
        lbl_linhas_leg = criar_label_com_ajuda(fl3d, "Linhas por legenda:",
            "Número de linhas visíveis em cada 'página'\n"
            "de legenda (1-8 linhas).")
        lbl_linhas_leg.pack(side="left")
        self.var_linhas_legenda = ctk.StringVar(value="3")
        ctk.CTkComboBox(fl3d, values=["1", "2", "3", "4", "5", "6", "7", "8"],
                        variable=self.var_linhas_legenda, width=80).pack(side="right")
        
        # === MODO KARAOKE ===
        ctk.CTkLabel(frame_legendas, text="─── Modo Karaoke ───",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(anchor="w", padx=10, pady=(10, 5))
        
        # Nota sobre exclusão mútua
        ctk.CTkLabel(frame_legendas, 
                     text="ℹ️ Karaoke e legendas normais são mutuamente exclusivos",
                     text_color="orange", font=ctk.CTkFont(size=10)).pack(anchor="w", padx=10, pady=2)
        
        # Ativar karaoke
        fl_k1 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl_k1.pack(fill="x", padx=10, pady=5)
        lbl_karaoke = criar_label_com_ajuda(fl_k1, "Modo karaoke:",
            "Destaca palavra a palavra sincronizado\n"
            "com a narração do áudio.\n"
            "Ideal para acompanhamento de leitura.\n"
            "NOTA: Desativa legendas normais quando ativo.")
        lbl_karaoke.pack(side="left")
        self.var_karaoke = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fl_k1, text="", variable=self.var_karaoke,
                        command=self._ao_mudar_karaoke).pack(side="right")
        
        # Idioma do Karaoke (NOVO v1.7)
        fl_k1b = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl_k1b.pack(fill="x", padx=10, pady=5)
        lbl_idioma_karaoke = criar_label_com_ajuda(fl_k1b, "Idioma do karaoke:",
            "Qual texto usar no karaoke:\n"
            "'Igual ao áudio': texto na mesma língua do áudio\n"
            "'Língua original': sempre texto original\n"
            "'Língua traduzida': sempre texto traduzido")
        lbl_idioma_karaoke.pack(side="left")
        self.var_idioma_karaoke = ctk.StringVar(value="Igual ao áudio")
        ctk.CTkComboBox(fl_k1b, values=["Igual ao áudio", "Língua original", "Língua traduzida"],
                        variable=self.var_idioma_karaoke, width=160).pack(side="right")
        
        # Modo exibição karaoke
        fl_k2 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl_k2.pack(fill="x", padx=10, pady=5)
        lbl_modo_karaoke = criar_label_com_ajuda(fl_k2, "Exibição karaoke:",
            "'Scroll contínuo': texto faz scroll acompanhando palavra\n"
            "'Por página': mostra N linhas, muda quando termina página")
        lbl_modo_karaoke.pack(side="left")
        self.var_modo_karaoke = ctk.StringVar(value="Scroll contínuo")
        ctk.CTkComboBox(fl_k2, values=["Scroll contínuo", "Por página"],
                        variable=self.var_modo_karaoke, width=140).pack(side="right")
        
        # Cor destaque karaoke
        fl_k3 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl_k3.pack(fill="x", padx=10, pady=5)
        lbl_cor_karaoke = criar_label_com_ajuda(fl_k3, "Cor destaque:",
            "Cor do fundo que destaca a palavra atual.")
        lbl_cor_karaoke.pack(side="left")
        self.var_cor_karaoke = ctk.StringVar(value="Yellow")
        cores_web_safe = ["Yellow", "Cyan", "Lime", "Magenta", "Orange", "Pink", 
                          "Aqua", "Red", "Green", "Blue", "White"]
        ctk.CTkComboBox(fl_k3, values=cores_web_safe,
                        variable=self.var_cor_karaoke, width=120).pack(side="right")
        
        # Transparência destaque karaoke
        fl_k4 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl_k4.pack(fill="x", padx=10, pady=5)
        lbl_transp_karaoke = criar_label_com_ajuda(fl_k4, "Transparência destaque:",
            "0% = totalmente transparente\n"
            "100% = cor sólida")
        lbl_transp_karaoke.pack(side="left")
        self.var_transp_karaoke = ctk.IntVar(value=70)
        self.slider_transp_karaoke = ctk.CTkSlider(fl_k4, from_=0, to=100, 
                                                    variable=self.var_transp_karaoke, width=100)
        self.slider_transp_karaoke.pack(side="right")
        self.lbl_transp_valor = ctk.CTkLabel(fl_k4, text="70%", width=40)
        self.lbl_transp_valor.pack(side="right", padx=5)
        self.slider_transp_karaoke.configure(command=self._ao_mudar_transp_karaoke)
        
        # === GERAR SRT ===
        ctk.CTkLabel(frame_legendas, text=" Ficheiro SRT ",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(anchor="w", padx=10, pady=(10, 5))
        
        fl4 = ctk.CTkFrame(frame_legendas, fg_color="transparent")
        fl4.pack(fill="x", padx=10, pady=5)
        lbl_gerar_srt = criar_label_com_ajuda(fl4, self.idioma.t("lbl_gerar_srt"),
            "Cria um ficheiro .srt com as legendas.\n"
            "Pode ser usado em leitores de vídeo como por exemplo VLC\n"
            "para ativar/desativar legendas.")
        lbl_gerar_srt.pack(side="left")
        
        # Opção para gerar SRT automaticamente
        self.var_gerar_srt_auto = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fl4, text="Auto", variable=self.var_gerar_srt_auto,
                        width=60).pack(side="right", padx=5)
        ctk.CTkButton(fl4, text=self.idioma.t("btn_gerar_srt"),
                      command=self.gerar_ficheiro_srt, width=140).pack(side="right")
    
    def _ao_mudar_karaoke(self):
        """Callback quando ativa/desativa karaoke - NOVO v1.7"""
        if self.var_karaoke.get():
            # Desativar legendas normais (são mutuamente exclusivos)
            self.var_legenda_video.set(False)
    
    def _criar_tab_html5(self):
        """Cria a tab de exportação HTML5"""
        scroll = ctk.CTkScrollableFrame(self.tab_html5)
        scroll.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Título
        ctk.CTkLabel(scroll, text="📦 Exportar para HTML5",
                     font=ctk.CTkFont(size=18, weight="bold")).pack(pady=10)
        
        ctk.CTkLabel(scroll, text="Crie uma versão web interactiva da sua apresentação",
                     text_color="gray").pack(pady=(0, 15))
        
        # === PASTA DE SAÍDA (NOVO v1.9.1) ===
        frame_pasta = ctk.CTkFrame(scroll)
        frame_pasta.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_pasta, text="▼ Pasta de Saída",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fp1 = ctk.CTkFrame(frame_pasta, fg_color="transparent")
        fp1.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fp1, text="Pasta HTML5:").pack(side="left")
        self.lbl_pasta_html5 = ctk.CTkLabel(fp1, text=self._pasta_saida_html5 or "(usar pasta geral)", 
                                             text_color="gray", width=300, anchor="w")
        self.lbl_pasta_html5.pack(side="left", padx=10)
        ctk.CTkButton(fp1, text="Escolher...", width=100, 
                      command=self._escolher_pasta_html5).pack(side="right")
        
        fp2 = ctk.CTkFrame(frame_pasta, fg_color="transparent")
        fp2.pack(fill="x", padx=10, pady=5)
        self.var_html_usar_pasta_geral = ctk.BooleanVar(value=not bool(self._pasta_saida_html5))
        ctk.CTkCheckBox(fp2, text="Usar pasta de saída geral (Vídeo/PPTX)", 
                        variable=self.var_html_usar_pasta_geral,
                        command=self._ao_mudar_usar_pasta_geral).pack(side="left")
        
        fp3 = ctk.CTkFrame(frame_pasta, fg_color="transparent")
        fp3.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fp3, text="ℹ️ Subpastas criadas automaticamente: 'html_unico/' e 'html_pasta/'",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left")
        
        # === FORMATO DE SAÍDA ===
        frame_formato = ctk.CTkFrame(scroll)
        frame_formato.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_formato, text="▼ Formato de Saída",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        f1 = ctk.CTkFrame(frame_formato, fg_color="transparent")
        f1.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f1, text="Formato:").pack(side="left")
        self.var_html_formato = ctk.StringVar(value="folder")
        ctk.CTkComboBox(f1, values=["Pasta com ficheiros", "Ficheiro único (.html)"],
                        variable=self.var_html_formato, width=200,
                        command=self._ao_mudar_formato_html).pack(side="right")
        
        f2 = ctk.CTkFrame(frame_formato, fg_color="transparent")
        f2.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(f2, text="Recursos:").pack(side="left")
        self.var_html_recursos = ctk.StringVar(value="embedded")
        ctk.CTkComboBox(f2, values=["Embedados (offline)", "CDN (ficheiro menor)"],
                        variable=self.var_html_recursos, width=200).pack(side="right")
        
        # === LAYOUT E VISUALIZAÇÃO ===
        frame_layout = ctk.CTkFrame(scroll)
        frame_layout.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_layout, text="▼ Layout e Visualização",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fl1 = ctk.CTkFrame(frame_layout, fg_color="transparent")
        fl1.pack(fill="x", padx=10, pady=5)
        lbl_modo = criar_label_com_ajuda(fl1, "Modo:",
            "• Área separada: Slide em cima, texto em baixo\n"
            "• Sobreposto: Texto sobre o slide\n"
            "• Corpo principal: Texto grande, slide pequeno\n"
            "• Só texto: Apenas texto (foco na leitura)")
        lbl_modo.pack(side="left")
        self.var_html_layout = ctk.StringVar(value="separated")
        ctk.CTkComboBox(fl1, values=["Área separada", "Sobreposto no slide", 
                                      "Corpo principal", "Só texto"],
                        variable=self.var_html_layout, width=200,
                        command=self._ao_mudar_layout_html).pack(side="right")
        
        fl2 = ctk.CTkFrame(frame_layout, fg_color="transparent")
        fl2.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fl2, text="Mostrar slide:").pack(side="left")
        self.var_html_mostrar_slide = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fl2, text="", variable=self.var_html_mostrar_slide).pack(side="left", padx=10)
        ctk.CTkLabel(fl2, text="Tamanho:").pack(side="left", padx=(20, 0))
        self.var_html_tamanho_slide = ctk.StringVar(value="Médio")
        ctk.CTkComboBox(fl2, values=["Grande", "Médio", "Pequeno", "Miniatura"],
                        variable=self.var_html_tamanho_slide, width=120).pack(side="right")
        
        # === ÁUDIO ===
        frame_audio = ctk.CTkFrame(scroll)
        frame_audio.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_audio, text="▼ Áudio",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fa1 = ctk.CTkFrame(frame_audio, fg_color="transparent")
        fa1.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fa1, text="Incluir áudio:").pack(side="left")
        self.var_html_audio = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fa1, text="", variable=self.var_html_audio).pack(side="left", padx=10)
        ctk.CTkLabel(fa1, text="Idioma:").pack(side="left", padx=(20, 0))
        self.var_html_idioma_audio = ctk.StringVar(value="Original")
        ctk.CTkComboBox(fa1, values=["Original", "Traduzido", "Ambos (selector)"],
                        variable=self.var_html_idioma_audio, width=150).pack(side="right")
        
        fa2 = ctk.CTkFrame(frame_audio, fg_color="transparent")
        fa2.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fa2, text="Auto-play:").pack(side="left")
        self.var_html_autoplay = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fa2, text="Avançar automaticamente entre slides",
                        variable=self.var_html_autoplay).pack(side="left", padx=10)
        
        # === KARAOKE / LEGENDAS ===
        frame_karaoke = ctk.CTkFrame(scroll)
        frame_karaoke.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_karaoke, text="▼ Karaoke / Legendas",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fk1 = ctk.CTkFrame(frame_karaoke, fg_color="transparent")
        fk1.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fk1, text="Activar karaoke:").pack(side="left")
        self.var_html_karaoke = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fk1, text="Destacar palavra actual",
                        variable=self.var_html_karaoke).pack(side="left", padx=10)
        
        fk2 = ctk.CTkFrame(frame_karaoke, fg_color="transparent")
        fk2.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fk2, text="Idioma texto:").pack(side="left")
        self.var_html_idioma_texto = ctk.StringVar(value="Igual ao áudio")
        ctk.CTkComboBox(fk2, values=["Igual ao áudio", "Original", "Traduzido"],
                        variable=self.var_html_idioma_texto, width=150).pack(side="right")
        
        fk3 = ctk.CTkFrame(frame_karaoke, fg_color="transparent")
        fk3.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fk3, text="Cor destaque:").pack(side="left")
        self.var_html_cor_karaoke = ctk.StringVar(value="Amarelo")
        ctk.CTkComboBox(fk3, values=["Amarelo", "Verde", "Azul", "Laranja", "Rosa", "Ciano"],
                        variable=self.var_html_cor_karaoke, width=120).pack(side="left", padx=10)
        ctk.CTkLabel(fk3, text="Modo:").pack(side="left", padx=(20, 0))
        self.var_html_modo_karaoke = ctk.StringVar(value="Scroll")
        ctk.CTkComboBox(fk3, values=["Scroll", "Página"],
                        variable=self.var_html_modo_karaoke, width=100).pack(side="right")
        
        # Opacidade do marcador (NOVO)
        fk4 = ctk.CTkFrame(frame_karaoke, fg_color="transparent")
        fk4.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fk4, text="Opacidade marcador:").pack(side="left")
        self.var_html_opacidade_karaoke = ctk.DoubleVar(value=0.4)
        self.slider_opacidade_html = ctk.CTkSlider(fk4, from_=0.1, to=1.0, 
                                                    variable=self.var_html_opacidade_karaoke,
                                                    width=150)
        self.slider_opacidade_html.pack(side="left", padx=10)
        self.lbl_opacidade_html = ctk.CTkLabel(fk4, text="40%", width=40)
        self.lbl_opacidade_html.pack(side="left")
        self.slider_opacidade_html.configure(command=self._ao_mudar_opacidade_html)
        
        # === NAVEGAÇÃO ===
        frame_nav = ctk.CTkFrame(scroll)
        frame_nav.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_nav, text="▼ Navegação e Controlos",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fn1 = ctk.CTkFrame(frame_nav, fg_color="transparent")
        fn1.pack(fill="x", padx=10, pady=5)
        self.var_html_controlos = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fn1, text="Mostrar controlos (play/pause)",
                        variable=self.var_html_controlos).pack(side="left")
        self.var_html_progresso = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fn1, text="Barra de progresso",
                        variable=self.var_html_progresso).pack(side="left", padx=20)
        
        fn2 = ctk.CTkFrame(frame_nav, fg_color="transparent")
        fn2.pack(fill="x", padx=10, pady=5)
        self.var_html_teclado = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fn2, text="Navegação por teclado",
                        variable=self.var_html_teclado).pack(side="left")
        self.var_html_touch = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(fn2, text="Gestos touch (mobile)",
                        variable=self.var_html_touch).pack(side="left", padx=20)
        
        fn3 = ctk.CTkFrame(frame_nav, fg_color="transparent")
        fn3.pack(fill="x", padx=10, pady=5)
        self.var_html_indice = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(fn3, text="Índice de slides (menu lateral)",
                        variable=self.var_html_indice).pack(side="left")
        
        # === APARÊNCIA ===
        frame_aparencia = ctk.CTkFrame(scroll)
        frame_aparencia.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_aparencia, text="▼ Aparência",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=5)
        
        fap1 = ctk.CTkFrame(frame_aparencia, fg_color="transparent")
        fap1.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fap1, text="Tema:").pack(side="left")
        self.var_html_tema = ctk.StringVar(value="Escuro")
        ctk.CTkComboBox(fap1, values=["Escuro", "Claro", "Alto contraste"],
                        variable=self.var_html_tema, width=150).pack(side="left", padx=10)
        
        fap2 = ctk.CTkFrame(frame_aparencia, fg_color="transparent")
        fap2.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(fap2, text="Tamanho fonte:").pack(side="left")
        self.var_html_fonte_tamanho = ctk.StringVar(value="Médio")
        ctk.CTkComboBox(fap2, values=["Pequeno", "Médio", "Grande", "Extra grande"],
                        variable=self.var_html_fonte_tamanho, width=130).pack(side="left", padx=10)
        ctk.CTkLabel(fap2, text="Fonte:").pack(side="left", padx=(20, 0))
        self.var_html_fonte_familia = ctk.StringVar(value="Sans-serif")

        # ctk.CTkComboBox(fap2, values=["Sans-serif", "Serif", "Monospace"],
         #               variable=self.var_html_fonte_familia, width=120).pack(side="right")

         # Fontes padrão + fontes acessíveis para dislexia/baixa visão
        ctk.CTkComboBox(fap2, values=["Sans-serif", "Serif", "Monospace", 
                                       "Atkinson (acessível)", "Lexend (acessível)", "OpenDyslexic"],
                        variable=self.var_html_fonte_familia, width=160).pack(side="right")
        
        # === BOTÕES DE ACÇÃO ===
        frame_acoes = ctk.CTkFrame(scroll)
        frame_acoes.pack(fill="x", padx=10, pady=15)
        
        self.btn_preview_html = ctk.CTkButton(frame_acoes, text="👁 Pré-visualizar",
                                               command=self._preview_html, width=150)
        self.btn_preview_html.pack(side="left", padx=10, pady=10)
        
        self.btn_exportar_html = ctk.CTkButton(frame_acoes, text="📦 Exportar HTML5",
                                                command=self._exportar_html, width=180,
                                                fg_color="#28a745", hover_color="#218838")
        self.btn_exportar_html.pack(side="left", padx=10, pady=10)
        
        # Info sobre atalhos
        frame_info = ctk.CTkFrame(scroll, fg_color="transparent")
        frame_info.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_info, text="ℹ️ Atalhos no player: Espaço=Play, ←→=Navegar, F=Fullscreen, L=Idioma, M=Mute",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w")
    
    def _ao_mudar_formato_html(self, valor):
        """Callback quando muda o formato de exportação HTML"""
        # Pode ser usado para ajustar outras opções baseado no formato
        pass
    
    def _ao_mudar_layout_html(self, valor):
        """Callback quando muda o layout HTML"""
        # Se "Só texto", desactivar opção de mostrar slide
        if "Só texto" in valor:
            self.var_html_mostrar_slide.set(False)
        else:
            self.var_html_mostrar_slide.set(True)
    
    def _escolher_pasta_html5(self):
        """Escolhe pasta específica para exportação HTML5 - v1.9.1"""
        pasta = filedialog.askdirectory(
            title="Escolher pasta para exportação HTML5",
            initialdir=self._pasta_saida_html5 or self._pasta_saida or os.path.expanduser("~")
        )
        if pasta:
            self._pasta_saida_html5 = pasta
            self.lbl_pasta_html5.configure(text=pasta, text_color="white")
            self.var_html_usar_pasta_geral.set(False)
            self.guardar_config()
    
    def _ao_mudar_usar_pasta_geral(self):
        """Callback quando muda checkbox de usar pasta geral - v1.9.1"""
        if self.var_html_usar_pasta_geral.get():
            self._pasta_saida_html5 = ""
            self.lbl_pasta_html5.configure(text="(usar pasta geral)", text_color="gray")
        self.guardar_config()
    
    def _ao_mudar_opacidade_html(self, valor):
        """Callback quando muda slider de opacidade do karaoke HTML5"""
        pct = int(float(valor) * 100)
        self.lbl_opacidade_html.configure(text=f"{pct}%")
    
    def _obter_pasta_saida_html5(self) -> str:
        """Obtém a pasta de saída para HTML5 - v1.9.3: usa pasta estruturada"""
        # Se tem pasta específica HTML5, usar essa
        if self._pasta_saida_html5:
            return self._pasta_saida_html5
        # v1.9.3: Usar pasta estruturada
        return self._obter_pasta_estruturada("html5")
    
    def _preview_html(self):
        """Pré-visualiza a exportação HTML"""
        if not self.pptx.apresentacao:
            messagebox.showwarning("Aviso", "Por favor, abra uma apresentação primeiro.")
            return
        
        # Exportar para pasta temporária
        import tempfile
        import webbrowser
        
        temp_dir = tempfile.mkdtemp(prefix="pptx_html_preview_")
        
        try:
            self._log("A preparar pré-visualização...")
            
            config = self._obter_config_html()
            config.format = "folder"  # Preview sempre como pasta
            
            exporter = HTMLExporter()
            exporter.config = config
            
            # Preparar dados dos slides (com imagens na pasta temporária)
            slides_data = self._preparar_slides_html(temp_dir)
            
            title = Path(self.pptx.apresentacao.caminho).stem if self.pptx.apresentacao else "Preview"
            
            self._log("A gerar HTML...")
            
            if exporter.export(temp_dir, slides_data, title):
                # Abrir no browser
                index_path = os.path.join(temp_dir, "index.html")
                webbrowser.open(f"file://{index_path}")
                self._log("✅ Preview aberto no browser")
            else:
                messagebox.showerror("Erro", "Falha ao gerar preview")
        except Exception as e:
            self._log(f"❌ Erro no preview: {e}")
            messagebox.showerror("Erro", f"Erro no preview: {e}")
    
    def _exportar_html(self):
        """Exporta a apresentação para HTML5 - v1.9.3: usa pasta estruturada"""
        if not self.pptx.apresentacao:
            messagebox.showwarning("Aviso", "Por favor, abra uma apresentação primeiro.")
            return
        
        # Obter configuração
        config = self._obter_config_html()
        
        # v1.9.3: Usar pasta estruturada
        pasta_base = self._obter_pasta_saida_html5()
        nome_base = self._obter_nome_base_apresentacao()
        
        # Definir caminho de saída
        if config.format == "folder":
            output_path = os.path.join(pasta_base, f"{nome_base}_html")
        else:
            output_path = os.path.join(pasta_base, f"{nome_base}.html")
        
        # Confirmar com utilizador
        msg = f"Exportar para:\n{output_path}\n\nContinuar?"
        if not messagebox.askyesno("Confirmar Exportação", msg):
            return
        
        # Mudar para tab de progresso
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        # Exportar em thread
        def exportar():
            try:
                import tempfile
                
                self.a_processar = True
                self._atualizar_estado_botoes()
                self.lbl_estado.configure(text="A exportar HTML5...", text_color="orange")
                self._log("A preparar slides...")
                
                # Criar pasta temporária para imagens
                temp_dir = tempfile.mkdtemp(prefix="pptx_html_export_")
                
                exporter = HTMLExporter()
                exporter.config = config
                exporter.set_progress_callback(self._progresso_html)
                
                slides_data = self._preparar_slides_html(temp_dir)
                title = Path(self.pptx.apresentacao.caminho).stem if self.pptx.apresentacao else "Apresentação"
                
                self._log("A gerar HTML5...")
                
                if exporter.export(output_path, slides_data, title):
                    self.lbl_estado.configure(text="HTML5 exportado!", text_color="green")
                    self._log(f"HTML5 exportado para: {output_path}")
                    # Registar ficheiro para botões de acção
                    if config.format == "folder":
                        index_path = os.path.join(output_path, "index.html")
                        self.after(0, lambda p=index_path: self._registar_ficheiro_gerado(p, "HTML5"))
                    else:
                        self.after(0, lambda p=output_path: self._registar_ficheiro_gerado(p, "HTML5"))
                    self.after(0, lambda: messagebox.showinfo("Sucesso", f"HTML5 exportado para:\n{output_path}"))
                else:
                    self.lbl_estado.configure(text="Erro na exportação", text_color="red")
                    self._log("Falha na exportação HTML5")
                
            except Exception as e:
                self.lbl_estado.configure(text="Erro", text_color="red")
                self._log(f"❌ Erro: {e}")
                self.after(0, lambda: messagebox.showerror("Erro", f"Erro na exportação: {e}"))
            finally:
                self.a_processar = False
                self._atualizar_estado_botoes()
        
        threading.Thread(target=exportar, daemon=True).start()
    
    def _obter_config_html(self) -> ConfigExportHTML:
        """Obtém configuração HTML5 da interface"""
        config = ConfigExportHTML()
        
        # Formato
        config.format = "single" if "único" in self.var_html_formato.get() else "folder"
        config.resources = "cdn" if "CDN" in self.var_html_recursos.get() else "embedded"
        
        # Layout
        modos_layout = {
            "Área separada": "separated",
            "Sobreposto no slide": "overlay",
            "Corpo principal": "main",
            "Só texto": "textonly"
        }
        config.layout.mode = modos_layout.get(self.var_html_layout.get(), "separated")
        config.layout.show_slide = self.var_html_mostrar_slide.get()
        
        tamanhos = {"Grande": "large", "Médio": "medium", "Pequeno": "small", "Miniatura": "thumbnail"}
        config.layout.slide_size = tamanhos.get(self.var_html_tamanho_slide.get(), "medium")
        
        # Áudio
        config.audio.enabled = self.var_html_audio.get()
        idiomas_audio = {"Original": "original", "Traduzido": "translated", "Ambos (selector)": "both"}
        config.audio.language = idiomas_audio.get(self.var_html_idioma_audio.get(), "original")
        config.audio.autoplay = self.var_html_autoplay.get()
        
        # Karaoke
        config.karaoke.enabled = self.var_html_karaoke.get()
        idiomas_texto = {"Igual ao áudio": "same", "Original": "original", "Traduzido": "translated"}
        config.karaoke.text_language = idiomas_texto.get(self.var_html_idioma_texto.get(), "same")
        
        cores = {
            "Amarelo": "#FFFF00", "Verde": "#00FF00", "Azul": "#00BFFF",
            "Laranja": "#FFA500", "Rosa": "#FF69B4", "Ciano": "#00FFFF"
        }
        config.karaoke.highlight_color = cores.get(self.var_html_cor_karaoke.get(), "#FFFF00")
        config.karaoke.highlight_opacity = self.var_html_opacidade_karaoke.get()  # Opacidade do marcador
        config.karaoke.scroll_mode = "scroll" if "Scroll" in self.var_html_modo_karaoke.get() else "page"
        
        # Navegação
        config.navigation.show_controls = self.var_html_controlos.get()
        config.navigation.show_progress = self.var_html_progresso.get()
        config.navigation.keyboard_nav = self.var_html_teclado.get()
        config.navigation.touch_gestures = self.var_html_touch.get()
        config.navigation.slide_index = self.var_html_indice.get()
        
        # Aparência
        temas = {"Escuro": "dark", "Claro": "light", "Alto contraste": "highcontrast"}
        config.appearance.theme = temas.get(self.var_html_tema.get(), "dark")
        
        tamanhos_fonte = {"Pequeno": "small", "Médio": "medium", "Grande": "large", "Extra grande": "xlarge"}
        config.appearance.font_size = tamanhos_fonte.get(self.var_html_fonte_tamanho.get(), "medium")
        
        # Famílias de fonte: padrão + acessíveis
        familias = {
            "Sans-serif": "sans-serif", 
            "Serif": "serif", 
            "Monospace": "monospace",
            "Atkinson (acessível)": "atkinson",
            "Lexend (acessível)": "lexend",
            "OpenDyslexic": "opendyslexic"
        }
        config.appearance.font_family = familias.get(self.var_html_fonte_familia.get(), "sans-serif")
        
        return config
    
    def _preparar_slides_html(self, pasta_temp: str = None) -> list:
        """
        Prepara dados dos slides para exportação HTML.
        v1.9.1: Exporta imagens dos slides para pasta temporária.
        
        Args:
            pasta_temp: Pasta para guardar imagens temporárias
        
        Returns:
            Lista de SlideDataHTML
        """
        import tempfile
        
        slides_data = []
        
        # Criar pasta temporária para imagens se não fornecida
        if not pasta_temp:
            pasta_temp = tempfile.mkdtemp(prefix="pptx_html_slides_")
        
        # Exportar slides como imagens
        imagens = []
        try:
            imagens = self.pptx.exportar_slides_imagens(pasta_temp)
            self._log(f"Exportadas {len(imagens)} imagens de slides")
        except Exception as e:
            self._log(f"⚠️ Aviso: Não foi possível exportar imagens ({e})")
        
        for i in range(1, self.pptx.num_slides + 1):
            slide_info = self.pptx.obter_slide(i)
            if slide_info:
                # Encontrar imagem correspondente (índice i-1)
                image_path = ""
                if i <= len(imagens):
                    image_path = imagens[i - 1]
                
                slide_data = SlideDataHTML(
                    id=i,
                    image_path=image_path,
                    audio_original=slide_info.caminho_audio or "",
                    audio_translated=slide_info.caminho_audio_traduzido or "",
                    duration=slide_info.duracao_audio if slide_info.duracao_audio > 0 else 5.0,
                    text_original=slide_info.texto_narrar or "",
                    text_translated=slide_info.texto_traduzido or ""
                )
                slides_data.append(slide_data)
        
        return slides_data
    
    def _progresso_html(self, atual: int, total: int, msg: str = ""):
        """Callback de progresso para exportação HTML"""
        self.progress_total.set(atual / total)
        self.lbl_pct.configure(text=f"{atual}%")
        if msg:
            self._log(msg)
    
    def _criar_tab_progresso(self):
        frame_estado = ctk.CTkFrame(self.tab_progresso)
        frame_estado.pack(fill="x", padx=10, pady=10)
        ctk.CTkLabel(frame_estado, text=self.idioma.t("lbl_estado"),
                     font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)
        self.lbl_estado = ctk.CTkLabel(frame_estado, text=self.idioma.t("estado_pronto"), text_color="green")
        self.lbl_estado.pack(side="left")
        
        frame_prog = ctk.CTkFrame(self.tab_progresso)
        frame_prog.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(frame_prog, text=self.idioma.t("lbl_progresso_total")).pack(anchor="w", padx=10)
        self.progress_total = ctk.CTkProgressBar(frame_prog)
        self.progress_total.pack(fill="x", padx=10, pady=5)
        self.progress_total.set(0)
        self.lbl_pct = ctk.CTkLabel(frame_prog, text="0%")
        self.lbl_pct.pack(anchor="e", padx=10)
        
        ctk.CTkLabel(self.tab_progresso, text=self.idioma.t("lbl_log"),
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(20, 5))
        self.txt_log = ctk.CTkTextbox(self.tab_progresso, height=250)
        self.txt_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # v1.9.3: Painel de acções rápidas
        frame_acoes = ctk.CTkFrame(self.tab_progresso)
        frame_acoes.pack(fill="x", padx=10, pady=(5, 10))
        
        self.btn_abrir_ficheiro = ctk.CTkButton(
            frame_acoes, 
            text="Abrir Ficheiro",
            command=self._abrir_ultimo_ficheiro,
            width=140,
            state="disabled"
        )
        self.btn_abrir_ficheiro.pack(side="left", padx=5)
        
        self.btn_abrir_pasta = ctk.CTkButton(
            frame_acoes,
            text="Abrir Pasta",
            command=self._abrir_pasta_saida,
            width=140,
            state="disabled"
        )
        self.btn_abrir_pasta.pack(side="left", padx=5)
        
        self.lbl_ultimo_ficheiro = ctk.CTkLabel(
            frame_acoes,
            text="",
            text_color="gray"
        )
        self.lbl_ultimo_ficheiro.pack(side="left", padx=15)
    
    # === CALLBACKS ===
    def _obter_lista_vozes(self) -> List[str]:
        vozes = VOZES_EDGE.get(self.tts.config.idioma, VOZES_EDGE["pt-PT"])
        return [v[1] for v in vozes]
    
    def _ao_mudar_idioma_voz(self, valor):
        self.tts.config.idioma = valor
        self.combo_voz.configure(values=self._obter_lista_vozes())
        vozes = VOZES_EDGE.get(valor, [])
        if vozes:
            self.combo_voz.set(vozes[0][1])
            self.tts.config.voz = vozes[0][0]
    
    def _ao_mudar_motor(self, nome_motor):
        """Callback quando muda o motor TTS"""
        codigo = self._mapa_motores.get(nome_motor, "edge")
        self.tts.config.motor = codigo
        # Atualizar lista de vozes
        self._atualizar_vozes_motor()
    
    def _atualizar_vozes_motor(self):
        """Atualiza lista de vozes baseado no motor selecionado"""
        vozes = self.tts.obter_vozes_disponiveis()
        if vozes:
            nomes = [v[1] for v in vozes]
            self.combo_voz.configure(values=nomes)
            self.combo_voz.set(nomes[0])
            self.tts.config.voz = vozes[0][0]
    
    def _ao_mudar_voz(self, nome):
        # Procurar em vozes do motor atual
        vozes = self.tts.obter_vozes_disponiveis()
        for codigo, n in vozes:
            if n == nome:
                self.tts.config.voz = codigo
                break
    
    def _ao_mudar_velocidade(self, v):
        velocidade = round(v, 2)
        self.tts.config.velocidade = velocidade
        self.var_vel.set(f"{velocidade:.2f}")
    
    def _ao_digitar_velocidade(self, event=None):
        try:
            valor = float(self.var_vel.get().replace(",", "."))
            valor = max(0.5, min(2.0, valor))
            valor = round(valor, 2)
            self.tts.config.velocidade = valor
            self.slider_vel.set(valor)
            self.var_vel.set(f"{valor:.2f}")
        except ValueError:
            self.var_vel.set(f"{self.tts.config.velocidade:.2f}")
    
    def _ao_mudar_mais_linguas_voz(self):
        """Callback quando ativa/desativa mais línguas para voz - v1.8"""
        if self.var_mais_linguas_voz.get():
            # Mostrar todas as línguas
            self._idiomas_voz_visiveis = list(IDIOMAS_DISPONIVEIS.keys())
        else:
            # Mostrar apenas principais
            self._idiomas_voz_visiveis = list(IDIOMAS_PRINCIPAIS.keys())
        
        # Actualizar combobox
        self.combo_idioma_voz.configure(values=self._idiomas_voz_visiveis)
        
        # Se idioma actual não está na lista, seleccionar o primeiro
        if self.tts.config.idioma not in self._idiomas_voz_visiveis:
            self.combo_idioma_voz.set(self._idiomas_voz_visiveis[0])
            self._ao_mudar_idioma_voz(self._idiomas_voz_visiveis[0])
    
    def _ao_mudar_mais_linguas_dest(self):
        """Callback quando ativa/desativa mais línguas para destino - v1.8"""
        if self.var_mais_linguas_dest.get():
            # Mostrar todas as línguas
            self._idiomas_dest_visiveis = list(IDIOMAS_DISPONIVEIS.values())
        else:
            # Mostrar apenas principais
            self._idiomas_dest_visiveis = list(IDIOMAS_PRINCIPAIS.values())
        
        # Actualizar combobox
        self.combo_idioma_destino.configure(values=self._idiomas_dest_visiveis)
        
        # Se idioma actual não está na lista, seleccionar English
        idioma_actual = self.combo_idioma_destino.get()
        if idioma_actual not in self._idiomas_dest_visiveis:
            self.combo_idioma_destino.set("English (US)")
            self._ao_mudar_idioma_destino("English (US)")
    
    def _ao_mudar_transp_karaoke(self, v):
        """Atualiza label de transparência karaoke - CORRIGIDO v1.7"""
        self.lbl_transp_valor.configure(text=f"{int(v)}%")
        # REMOVIDO: self.lbl_vel.configure(...) que estava a atualizar o campo errado
    
    def _ao_mudar_perfil(self, nome_perfil):
        """Atualiza descrição quando muda perfil"""
        for codigo, perfil in PERFIS_CONFIG.items():
            if perfil["nome"] == nome_perfil:
                self.lbl_desc_perfil.configure(text=perfil["descricao"])
                break
    
    def _aplicar_perfil(self):
        """Aplica o perfil selecionado"""
        nome_sel = self.combo_perfil.get()
        for codigo, perfil in PERFIS_CONFIG.items():
            if perfil["nome"] == nome_sel:
                # Aplicar motor
                motor = perfil["motor"]
                self.tts.config.motor = motor
                for nome, cod in self._mapa_motores.items():
                    if cod == motor:
                        self.combo_motor.set(nome)
                        break
                self._atualizar_vozes_motor()
                
                # Aplicar resolução
                self.var_resolucao.set(perfil["resolucao"])
                
                # Aplicar FPS
                self.var_fps.set(perfil["fps"])
                
                # Aplicar tempo extra
                self.var_tempo_extra.set(perfil["tempo_extra"])
                
                self._log(f"Perfil '{nome_sel}' aplicado")
                break
    
    def _ao_mudar_posicao(self, v):
        mapa = {
            "Superior Direito": "sup_dir", 
            "Superior Esquerdo": "sup_esq",
            "Inferior Direito": "inf_dir", 
            "Inferior Esquerdo": "inf_esq",
            "Centro Superior": "centro_sup",
            "Centro Inferior": "centro_inf",
            "Centro Esquerdo": "centro_esq",
            "Centro Direito": "centro_dir"
        }
        self.pptx.config_icone.posicao = mapa.get(v, "sup_dir")
    
    def _ao_mudar_traducao_ativa(self):
        self.tradutor.config.ativo = self.var_traducao_ativa.get()
    
    def _ao_mudar_idioma_destino(self, valor):
        # Converter nome para código
        for codigo, nome in IDIOMAS_DISPONIVEIS.items():
            if nome == valor:
                self.tradutor.config.idioma_destino = codigo
                # Atualizar TTS de tradução
                self.tts_trad.config.idioma = codigo
                vozes = VOZES_EDGE.get(codigo, [])
                if vozes:
                    self.tts_trad.config.voz = vozes[0][0]
                    # Atualizar combobox de voz traduzida
                    nomes_vozes = [v[1] for v in vozes]
                    self.combo_voz_trad.configure(values=nomes_vozes)
                    self.combo_voz_trad.set(nomes_vozes[0])
                break
    
    def _ao_mudar_voz_trad(self, nome):
        """Callback quando muda a voz de tradução - NOVO v1.7"""
        vozes = VOZES_EDGE.get(self.tts_trad.config.idioma, [])
        for codigo, n, _ in vozes:
            if n == nome:
                self.tts_trad.config.voz = codigo
                break
    
    def _ao_mudar_velocidade_trad(self, v):
        velocidade = round(v, 2)
        self.tts_trad.config.velocidade = velocidade
        self.var_vel_trad.set(f"{velocidade:.2f}")
    
    def _ao_digitar_velocidade_trad(self, event=None):
        try:
            valor = float(self.var_vel_trad.get().replace(",", "."))
            valor = max(0.5, min(2.0, valor))
            valor = round(valor, 2)
            self.tts_trad.config.velocidade = valor
            self.slider_vel_trad.set(valor)
            self.var_vel_trad.set(f"{valor:.2f}")
        except ValueError:
            self.var_vel_trad.set(f"{self.tts_trad.config.velocidade:.2f}")
    
    def _escolher_pasta(self):
        pasta = filedialog.askdirectory(title=self.idioma.t("dlg_pasta_titulo"))
        if pasta:
            self._pasta_saida = pasta
            self.lbl_pasta.configure(text=pasta)
    
    def _abrir_ultimo_ficheiro(self):
        """Abre o último ficheiro gerado com a aplicação predefinida do sistema"""
        if not self._ultimo_ficheiro or not os.path.exists(self._ultimo_ficheiro):
            return
        try:
            if sys.platform == 'win32':
                os.startfile(self._ultimo_ficheiro)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.run(['open', self._ultimo_ficheiro])
            else:
                import subprocess
                subprocess.run(['xdg-open', self._ultimo_ficheiro])
        except Exception as e:
            self._log(f"Erro ao abrir ficheiro: {e}")
    
    def _abrir_pasta_saida(self):
        """Abre a pasta de saída no explorador de ficheiros"""
        pasta = ""
        if self._ultimo_ficheiro and os.path.exists(self._ultimo_ficheiro):
            pasta = os.path.dirname(self._ultimo_ficheiro)
        elif self._pasta_saida and os.path.exists(self._pasta_saida):
            pasta = self._pasta_saida
        
        if not pasta:
            return
        
        try:
            if sys.platform == 'win32':
                os.startfile(pasta)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.run(['open', pasta])
            else:
                import subprocess
                subprocess.run(['xdg-open', pasta])
        except Exception as e:
            self._log(f"Erro ao abrir pasta: {e}")
    
    def _registar_ficheiro_gerado(self, caminho: str, tipo: str = ""):
        """Regista o último ficheiro gerado e actualiza interface"""
        if not caminho or not os.path.exists(caminho):
            return
        
        self._ultimo_ficheiro = caminho
        nome_ficheiro = os.path.basename(caminho)
        
        # Actualizar label e botões
        self.lbl_ultimo_ficheiro.configure(text=f"{tipo}: {nome_ficheiro}" if tipo else nome_ficheiro)
        self.btn_abrir_ficheiro.configure(state="normal")
        self.btn_abrir_pasta.configure(state="normal")
    
    def _obter_pasta_base(self) -> str:
        """Obtém a pasta base de saída (pasta do PPTX original ou pasta configurada)"""
        if self._pasta_saida:
            return self._pasta_saida
        elif self.pptx.apresentacao and self.pptx.apresentacao.caminho:
            return os.path.dirname(self.pptx.apresentacao.caminho)
        else:
            return os.path.expanduser("~")
    
    def _obter_pasta_estruturada(self, tipo: str) -> str:
        """
        Obtém pasta estruturada por tipo de ficheiro.
        tipo: 'audio', 'pptx', 'video', 'html5'
        Cria a pasta se não existir.
        """
        pasta_base = self._obter_pasta_base()
        
        if self._usar_pasta_estruturada:
            pasta = os.path.join(pasta_base, tipo)
        else:
            pasta = pasta_base
        
        os.makedirs(pasta, exist_ok=True)
        return pasta
    
    def _obter_nome_base_apresentacao(self) -> str:
        """Obtém o nome base da apresentação (sem extensão)"""
        if self.pptx.apresentacao and self.pptx.apresentacao.caminho:
            return Path(self.pptx.apresentacao.caminho).stem
        return "apresentacao"
    
    def _guardar_texto_slide(self, event=None):
        if self.pptx.apresentacao:
            texto = self.txt_narrar.get("1.0", "end-1c")
            self.pptx.atualizar_texto_narrar(self.slide_atual, texto)
    
    def _guardar_texto_traduzido(self, event=None):
        if self.pptx.apresentacao:
            texto = self.txt_traduzido.get("1.0", "end-1c")
            self.pptx.atualizar_texto_traduzido(self.slide_atual, texto)
    
    def _atualizar_estado_botoes(self):
        tem_pptx = self.pptx.apresentacao is not None
        estado = "normal" if tem_pptx and not self.a_processar else "disabled"
        self.btn_gerar_audio.configure(state=estado)
        self.btn_gerar_pptx.configure(state=estado)
        self.btn_gerar_video.configure(state=estado)
        self.btn_preview.configure(state=estado)
    
    def _log(self, msg: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.insert("end", f"[{timestamp}] {msg}\n")
        self.txt_log.see("end")
    
    def _atualizar_progresso(self, valor: float, msg: str = ""):
        self.progress_total.set(valor / 100)
        self.lbl_pct.configure(text=f"{int(valor)}%")
        if msg:
            self.lbl_estado.configure(text=msg)
    
    # === AÇÕES ===
    def abrir_pptx(self):
        if not PPTX_DISPONIVEL:
            messagebox.showerror("Erro", "python-pptx não instalado!\npip install python-pptx")
            return
        
        caminho = filedialog.askopenfilename(
            title=self.idioma.t("dlg_abrir_titulo"),
            filetypes=[("PowerPoint", "*.pptx"), ("Todos", "*.*")]
        )
        if not caminho:
            return
        
        if self.pptx.abrir(caminho):
            self._log(self.idioma.t("msg_ficheiro_aberto", caminho))
            self._log(self.idioma.t("msg_slides_encontrados", self.pptx.num_slides))
            self._atualizar_lista_slides()
            self._atualizar_estado_botoes()
            if self.pptx.num_slides > 0:
                self._selecionar_slide(1)
        else:
            messagebox.showerror("Erro", self.idioma.t("msg_erro_abrir", caminho))
    
    def _atualizar_lista_slides(self):
        for btn in self.lista_slides_btns:
            btn.destroy()
        self.lista_slides_btns.clear()
        self.lbl_nenhum.pack_forget()
        
        for i in range(1, self.pptx.num_slides + 1):
            slide = self.pptx.obter_slide(i)
            texto = slide.texto_narrar[:30] + "..." if len(slide.texto_narrar) > 30 else slide.texto_narrar
            btn = ctk.CTkButton(
                self.frame_lista_slides,
                text=f"{i}. {texto or '(vazio)'}",
                anchor="w",
                command=lambda n=i: self._selecionar_slide(n),
                width=180, height=30
            )
            btn.pack(fill="x", pady=2)
            self.lista_slides_btns.append(btn)
    
    def _selecionar_slide(self, numero: int):
        self._guardar_texto_slide()
        self.slide_atual = numero
        slide = self.pptx.obter_slide(numero)
        if not slide:
            return
        
        self.lbl_slide_titulo.configure(text=f"{self.idioma.t('lbl_slide')} {numero}")
        
        self.txt_slide.configure(state="normal")
        self.txt_slide.delete("1.0", "end")
        self.txt_slide.insert("1.0", slide.texto_visivel)
        self.txt_slide.configure(state="disabled")
        
        self.txt_notas.configure(state="normal")
        self.txt_notas.delete("1.0", "end")
        self.txt_notas.insert("1.0", slide.notas)
        self.txt_notas.configure(state="disabled")
        
        self.txt_narrar.delete("1.0", "end")
        self.txt_narrar.insert("1.0", slide.texto_narrar)
        
        # Texto traduzido
        self.txt_traduzido.delete("1.0", "end")
        self.txt_traduzido.insert("1.0", slide.texto_traduzido)
        
        # Status
        status_parts = []
        if slide.caminho_audio:
            status_parts.append("✓ Áudio")
        if slide.caminho_audio_traduzido:
            status_parts.append("✓ Áudio traduzido")
        self.lbl_status_slide.configure(text=" | ".join(status_parts))
        
        # Destacar botão selecionado
        for i, btn in enumerate(self.lista_slides_btns):
            if i + 1 == numero:
                btn.configure(fg_color=("gray75", "gray25"))
            else:
                btn.configure(fg_color=("gray85", "gray17"))
    
    def preview_audio(self):
        texto = self.txt_narrar.get("1.0", "end-1c").strip()
        if not texto:
            messagebox.showwarning("", self.idioma.t("msg_sem_texto", self.slide_atual))
            return
        
        self._log(f"A gerar preview para slide {self.slide_atual}...")
        
        def gerar():
            caminho = self.tts.gerar_preview(texto)
            if caminho:
                self.tts.tocar_audio(caminho)
                self._log("Preview a reproduzir")
            else:
                self._log("Erro ao gerar preview")
        
        threading.Thread(target=gerar, daemon=True).start()
    
    def preview_audio_traduzido(self):
        texto = self.txt_traduzido.get("1.0", "end-1c").strip()
        if not texto:
            messagebox.showwarning("", "Sem texto traduzido para este slide.")
            return
        
        self._log(f"A gerar preview traduzido para slide {self.slide_atual}...")
        
        def gerar():
            caminho = self.tts_trad.gerar_preview(texto)
            if caminho:
                self.tts_trad.tocar_audio(caminho)
                self._log("Preview traduzido a reproduzir")
            else:
                self._log("Erro ao gerar preview")
        
        threading.Thread(target=gerar, daemon=True).start()
    
    def parar_audio(self):
        self.tts.parar_audio()
        self.tts_trad.parar_audio()  # Parar ambos por segurança
        self._log("Reprodução parada")
    
    def parar_audio_traduzido(self):
        self.tts_trad.parar_audio()
        self.tts.parar_audio()  # Parar ambos por segurança
        self._log("Reprodução parada")
    
    def gerar_todos_audios(self):
        """Gera áudios - se tradução ativa, gera ambos (original + traduzido)"""
        if not self.pptx.apresentacao:
            return
        
        self._guardar_texto_slide()
        self._guardar_texto_traduzido()
        self.a_processar = True
        self._atualizar_estado_botoes()
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        # Verificar se tradução está ativa
        traducao_ativa = self.var_traducao_ativa.get()
        
        def processar():
            total = self.pptx.num_slides
            passos_totais = total * (2 if traducao_ativa else 1)
            if traducao_ativa:
                passos_totais += total  # Passo extra para tradução de texto
            
            passo_atual = 0
            
            # Determinar pasta de saída - v1.9.3: usar pasta estruturada
            pasta = self._obter_pasta_estruturada("audio")
            
            # PASSO 1: Gerar áudios originais
            self.after(0, lambda: self._log("=== A gerar áudios na língua original ==="))
            for i in range(1, total + 1):
                slide = self.pptx.obter_slide(i)
                passo_atual += 1
                self.after(0, lambda p=passo_atual, t=passos_totais: 
                    self._atualizar_progresso(p * 100 / t, f"Áudio original slide {i}..."))
                
                if not slide.texto_narrar.strip():
                    continue
                
                caminho = os.path.join(pasta, f"slide_{i:02d}.mp3")
                if self.tts.gerar_audio(slide.texto_narrar, caminho):
                    duracao = self.tts.obter_duracao(caminho)
                    self.pptx.definir_audio_slide(i, caminho, duracao)
                    self.after(0, lambda i=i: self._log(f"✓ Áudio original slide {i}"))
            
            # PASSO 2: Se tradução ativa, traduzir textos
            if traducao_ativa:
                self.after(0, lambda: self._log("=== A traduzir textos ==="))
                
                # Configurar tradutor
                motor_sel = self.combo_motor_trad.get()
                if "Google" in motor_sel:
                    self.tradutor.config.motor = "google"
                else:
                    self.tradutor.config.motor = "argos"
                self.tradutor.config.idioma_origem = self.tts.config.idioma
                self.tradutor.config.ativo = True
                
                for i in range(1, total + 1):
                    slide = self.pptx.obter_slide(i)
                    passo_atual += 1
                    self.after(0, lambda p=passo_atual, t=passos_totais:
                        self._atualizar_progresso(p * 100 / t, f"A traduzir slide {i}..."))
                    
                    if not slide.texto_narrar.strip():
                        continue
                    
                    # Só traduzir se ainda não tiver tradução
                    if not slide.texto_traduzido.strip():
                        traducao = self.tradutor.traduzir(slide.texto_narrar)
                        if traducao:
                            self.pptx.atualizar_texto_traduzido(i, traducao)
                            self.after(0, lambda i=i: self._log(f"✓ Traduzido slide {i}"))
                
                # PASSO 3: Gerar áudios traduzidos
                self.after(0, lambda: self._log("=== A gerar áudios traduzidos ==="))
                
                # Obter código do idioma destino
                idioma_dest = self.tradutor.config.idioma_destino
                codigo_idioma = idioma_dest.split('-')[0] if '-' in idioma_dest else idioma_dest[:2]
                
                # Configurar TTS para idioma destino
                from idiomas import VOZES_EDGE
                vozes_dest = VOZES_EDGE.get(idioma_dest, [])
                if vozes_dest:
                    self.tts_trad.config.idioma = idioma_dest
                    self.tts_trad.config.voz = vozes_dest[0][0]
                
                for i in range(1, total + 1):
                    slide = self.pptx.obter_slide(i)
                    passo_atual += 1
                    self.after(0, lambda p=passo_atual, t=passos_totais:
                        self._atualizar_progresso(p * 100 / t, f"Áudio traduzido slide {i}..."))
                    
                    if not slide.texto_traduzido.strip():
                        continue
                    
                    caminho = os.path.join(pasta, f"slide_{i:02d}_{codigo_idioma}.mp3")
                    if self.tts_trad.gerar_audio(slide.texto_traduzido, caminho):
                        duracao = self.tts_trad.obter_duracao(caminho)
                        self.pptx.definir_audio_traduzido_slide(i, caminho, duracao)
                        self.after(0, lambda i=i: self._log(f"✓ Áudio traduzido slide {i}"))
            
            self.after(0, lambda: self._atualizar_progresso(100, self.idioma.t("estado_concluido")))
            self.after(0, lambda: self._log(f"Áudios guardados em: {pasta}"))
            # Registar pasta de áudios para acesso rápido
            self.after(0, lambda: self._registar_ficheiro_gerado(pasta, "Áudios"))
            self.a_processar = False
            self.after(0, self._atualizar_estado_botoes)
            self.after(0, self._atualizar_lista_slides)
            self.after(0, lambda: self._selecionar_slide(self.slide_atual))
        
        threading.Thread(target=processar, daemon=True).start()
    
    def traduzir_todos_slides(self):
        """Traduz o texto de narração de todos os slides"""
        if not self.pptx.apresentacao:
            return
        
        if not self.tradutor.motor_disponivel():
            messagebox.showerror("Erro", "Motor de tradução não disponível.\nInstale: pip install deep-translator")
            return
        
        # Configurar motor
        motor_sel = self.combo_motor_trad.get()
        if "Google" in motor_sel:
            self.tradutor.config.motor = "google"
        else:
            self.tradutor.config.motor = "argos"
        
        # Configurar idioma origem (igual ao TTS principal)
        self.tradutor.config.idioma_origem = self.tts.config.idioma
        self.tradutor.config.ativo = True
        
        self._guardar_texto_slide()
        self.a_processar = True
        self._atualizar_estado_botoes()
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        def processar():
            total = self.pptx.num_slides
            
            for i in range(1, total + 1):
                slide = self.pptx.obter_slide(i)
                self.after(0, lambda i=i: self._atualizar_progresso(
                    (i-1) * 100 / total, self.idioma.t("msg_traduzindo", i)))
                self.after(0, lambda i=i: self._log(self.idioma.t("msg_traduzindo", i)))
                
                if not slide.texto_narrar.strip():
                    continue
                
                traducao = self.tradutor.traduzir(slide.texto_narrar)
                if traducao:
                    self.pptx.atualizar_texto_traduzido(i, traducao)
            
            self.after(0, lambda: self._atualizar_progresso(100, self.idioma.t("msg_traducao_concluida")))
            self.after(0, lambda: self._log(self.idioma.t("msg_traducao_concluida")))
            self.a_processar = False
            self.after(0, self._atualizar_estado_botoes)
            self.after(0, lambda: self._selecionar_slide(self.slide_atual))
        
        threading.Thread(target=processar, daemon=True).start()
    
    def gerar_todos_audios_traduzidos(self):
        """Gera áudio para os textos traduzidos"""
        if not self.pptx.apresentacao:
            return
        
        self._guardar_texto_traduzido()
        self.a_processar = True
        self._atualizar_estado_botoes()
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        def processar():
            total = self.pptx.num_slides
            
            # v1.9.3: Usar pasta estruturada
            pasta = self._obter_pasta_estruturada("audio")
            
            # Código do idioma destino para sufixo
            idioma_dest = self.tradutor.config.idioma_destino[:2]
            
            for i in range(1, total + 1):
                slide = self.pptx.obter_slide(i)
                self.after(0, lambda i=i: self._atualizar_progresso(
                    (i-1) * 100 / total, f"A gerar áudio traduzido slide {i}..."))
                self.after(0, lambda i=i: self._log(f"A gerar áudio traduzido slide {i}..."))
                
                if not slide.texto_traduzido.strip():
                    continue
                
                caminho = os.path.join(pasta, f"slide_{i:02d}_{idioma_dest}.mp3")
                if self.tts_trad.gerar_audio(slide.texto_traduzido, caminho):
                    duracao = self.tts_trad.obter_duracao(caminho)
                    self.pptx.definir_audio_traduzido_slide(i, caminho, duracao)
                    self.after(0, lambda i=i: self._log(f"Áudio traduzido gerado slide {i}"))
            
            self.after(0, lambda: self._atualizar_progresso(100, self.idioma.t("estado_concluido")))
            self.after(0, lambda: self._log(f"Áudios traduzidos guardados em: {pasta}"))
            # Registar pasta de áudios para acesso rápido
            self.after(0, lambda: self._registar_ficheiro_gerado(pasta, "Áudios Traduzidos"))
            self.a_processar = False
            self.after(0, self._atualizar_estado_botoes)
            self.after(0, lambda: self._selecionar_slide(self.slide_atual))
        
        threading.Thread(target=processar, daemon=True).start()
    
    def guardar_pptx_audio(self):
        if not self.pptx.apresentacao:
            return
        
        # Recolher configurações
        try:
            self.pptx.config_icone.tamanho_cm = float(getattr(self, 'var_tamanho', ctk.StringVar(value="1.0")).get())
        except: pass
        
        tem_traducao = any(s.caminho_audio_traduzido for s in self.pptx.apresentacao.slides)
        tem_audios_orig = any(s.caminho_audio for s in self.pptx.apresentacao.slides)
        legenda_slide = self.var_legenda_slide.get()
        legenda_notas = self.var_legenda_notas.get()
        traducao_ativa = self.var_traducao_ativa.get()
        
        # Mostrar confirmação se ativo
        if self.var_confirmar_antes.get():
            # Construir resumo
            posicao_nome = self.combo_pos.get()
            idioma_destino = self.combo_idioma_destino.get() if traducao_ativa else "N/A"
            
            resumo = "═══════════════════════════════════════\n"
            resumo += f"    PPTX Narrator v{APP_VERSAO} - GERAR PPTX\n"
            resumo += "═══════════════════════════════════════\n\n"
            
            resumo += f"📁 Apresentação: {os.path.basename(self.pptx.apresentacao.caminho)}\n"
            resumo += f"📊 Slides: {self.pptx.num_slides}\n\n"
            
            resumo += "── ÁUDIOS ──\n"
            resumo += f"  • Áudios originais: {'✅ Sim' if tem_audios_orig else '❌ Não'}\n"
            resumo += f"  • Áudios traduzidos: {'✅ Sim' if tem_traducao else '❌ Não'}\n\n"
            
            resumo += "── TRADUÇÃO ──\n"
            resumo += f"  • Tradução ativa: {'✅ Sim' if traducao_ativa else '❌ Não'}\n"
            if traducao_ativa:
                resumo += f"  • Idioma destino: {idioma_destino}\n"
            resumo += "\n"
            
            resumo += "── ÍCONES ──\n"
            resumo += f"  • Mostrar ícones: {'✅ Sim' if self.pptx.config_icone.mostrar else '❌ Não'}\n"
            resumo += f"  • Posição: {posicao_nome}\n"
            resumo += f"  • Tamanho: {self.pptx.config_icone.tamanho_cm} cm\n\n"
            
            resumo += "── LEGENDAS NO PPTX ──\n"
            resumo += f"  • Caixa texto no slide: {'✅ Sim' if legenda_slide else '❌ Não'}\n"
            resumo += f"  • Nas notas apresentador: {'✅ Sim' if legenda_notas else '❌ Não'}\n\n"
            
            resumo += "═══════════════════════════════════════\n"
            resumo += "Continuar com estas configurações?"
            
            if not messagebox.askyesno("Confirmar Geração PPTX", resumo):
                return
        
        # v1.9.3: Usar pasta estruturada com nome predefinido
        pasta_pptx = self._obter_pasta_estruturada("pptx")
        nome_base = self._obter_nome_base_apresentacao()
        nome_predefinido = f"{nome_base}_narrado.pptx"
        
        # Pedir caminho
        caminho = filedialog.asksaveasfilename(
            title=self.idioma.t("dlg_guardar_titulo"),
            initialdir=pasta_pptx,
            initialfile=nome_predefinido,
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")]
        )
        if not caminho:
            return
        
        if self.pptx.guardar_com_audio(caminho, incluir_traducao=tem_traducao,
                                        legenda_no_slide=legenda_slide,
                                        legenda_nas_notas=legenda_notas):
            self._log(self.idioma.t("msg_pptx_guardado", caminho))
            self._registar_ficheiro_gerado(caminho, "PPTX")
            messagebox.showinfo("", self.idioma.t("msg_pptx_guardado", caminho))
        else:
            messagebox.showerror("", self.idioma.t("msg_erro_gerar", "PPTX"))
    
    def gerar_ficheiro_srt(self):
        """Gera ficheiro de legendas SRT"""
        if not self.pptx.apresentacao:
            messagebox.showwarning("", "Nenhuma apresentação aberta.")
            return
        
        # Determinar se usa tradução baseado na configuração
        usar_traducao = self.var_traducao_ativa.get()
        
        # Verificar se há textos apropriados
        if usar_traducao:
            tem_texto = any(s.texto_traduzido.strip() for s in self.pptx.apresentacao.slides)
            if not tem_texto:
                messagebox.showwarning("", 
                    "Tradução ativa mas sem textos traduzidos.\n"
                    "Execute 'Traduzir Todos' primeiro ou desative a tradução.")
                return
        else:
            tem_texto = any(s.texto_narrar.strip() for s in self.pptx.apresentacao.slides)
            if not tem_texto:
                messagebox.showwarning("", "Sem texto para gerar legendas.")
                return
        
        # Verificar se existem áudios gerados (necessário para tempos corretos)
        if usar_traducao:
            tem_audios = any(s.caminho_audio_traduzido and os.path.exists(s.caminho_audio_traduzido) 
                            for s in self.pptx.apresentacao.slides)
            tipo_audio = "traduzidos"
        else:
            tem_audios = any(s.caminho_audio and os.path.exists(s.caminho_audio) 
                            for s in self.pptx.apresentacao.slides)
            tipo_audio = "originais"
        
        if not tem_audios:
            # Perguntar se quer gerar os áudios primeiro
            resposta = messagebox.askyesnocancel(
                "Áudios Necessários",
                f"Não foram encontrados áudios {tipo_audio}.\n\n"
                f"O ficheiro SRT precisa dos tempos de áudio para\n"
                f"sincronizar corretamente as legendas.\n\n"
                f"Deseja gerar os áudios {tipo_audio} agora?\n\n"
                f"• Sim: Gera áudios e depois o SRT\n"
                f"• Não: Gera SRT com tempos estimados (3s/slide)\n"
                f"• Cancelar: Volta sem fazer nada"
            )
            
            if resposta is None:  # Cancelar
                return
            elif resposta:  # Sim - gerar áudios primeiro
                self._log(f"A gerar áudios {tipo_audio} para SRT...")
                if usar_traducao:
                    self._gerar_audios_e_depois_srt(usar_traducao=True)
                else:
                    self._gerar_audios_e_depois_srt(usar_traducao=False)
                return
            # Se Não, continua com tempos estimados
        
        # Pedir caminho do ficheiro
        caminho = filedialog.asksaveasfilename(
            title="Guardar ficheiro SRT",
            defaultextension=".srt",
            filetypes=[("Legendas SRT", "*.srt")]
        )
        if not caminho:
            return
        
        # Obter tempo extra das configurações
        try:
            tempo_extra = float(self.var_tempo_extra.get())
        except:
            tempo_extra = 0.5
        
        if self.pptx.gerar_srt(caminho, usar_traducao=usar_traducao, 
                               tempo_extra=tempo_extra, tempo_minimo=3.0):
            self._log(f"Ficheiro SRT gerado: {caminho}")
            messagebox.showinfo("", f"Legendas guardadas em:\n{caminho}")
        else:
            messagebox.showerror("", "Erro ao gerar ficheiro SRT")
    
    def _gerar_audios_e_depois_srt(self, usar_traducao: bool):
        """Gera áudios e depois abre diálogo para guardar SRT"""
        self._guardar_texto_slide()
        if usar_traducao:
            self._guardar_texto_traduzido()
        
        self.a_processar = True
        self._atualizar_estado_botoes()
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        def processar():
            total = self.pptx.num_slides
            pasta = os.path.dirname(self.pptx.apresentacao.caminho)
            
            if usar_traducao:
                # Configurar TTS para idioma de destino
                idioma_dest = self.tradutor.config.idioma_destino
                self.tts_trad.config.idioma = idioma_dest
                vozes = VOZES_EDGE.get(idioma_dest, [])
                if vozes:
                    self.tts_trad.config.voz = vozes[0][0]
                
                for i in range(1, total + 1):
                    slide = self.pptx.obter_slide(i)
                    self.after(0, lambda i=i: self._atualizar_progresso(
                        (i-1) * 100 / total, f"A gerar áudio traduzido slide {i}..."))
                    
                    if not slide.texto_traduzido.strip():
                        continue
                    
                    caminho_audio = os.path.join(pasta, f"slide_{i:02d}_{idioma_dest}.mp3")
                    if self.tts_trad.gerar_audio(slide.texto_traduzido, caminho_audio):
                        duracao = self.tts_trad.obter_duracao(caminho_audio)
                        self.pptx.definir_audio_traduzido_slide(i, caminho_audio, duracao)
            else:
                for i in range(1, total + 1):
                    slide = self.pptx.obter_slide(i)
                    self.after(0, lambda i=i: self._atualizar_progresso(
                        (i-1) * 100 / total, f"A gerar áudio slide {i}..."))
                    
                    if not slide.texto_narrar.strip():
                        continue
                    
                    caminho_audio = os.path.join(pasta, f"slide_{i:02d}.mp3")
                    if self.tts.gerar_audio(slide.texto_narrar, caminho_audio):
                        duracao = self.tts.obter_duracao(caminho_audio)
                        self.pptx.definir_audio_slide(i, caminho_audio, duracao)
            
            self.after(0, lambda: self._atualizar_progresso(100, "Áudios gerados!"))
            self.a_processar = False
            self.after(0, self._atualizar_estado_botoes)
            
            # Agora abrir diálogo para guardar SRT
            self.after(100, lambda: self._finalizar_srt(usar_traducao))
        
        threading.Thread(target=processar, daemon=True).start()
    
    def _finalizar_srt(self, usar_traducao: bool):
        """Finaliza geração do SRT após áudios gerados"""
        caminho = filedialog.asksaveasfilename(
            title="Guardar ficheiro SRT",
            defaultextension=".srt",
            filetypes=[("Legendas SRT", "*.srt")]
        )
        if not caminho:
            return
        
        try:
            tempo_extra = float(self.var_tempo_extra.get())
        except:
            tempo_extra = 0.5
        
        if self.pptx.gerar_srt(caminho, usar_traducao=usar_traducao,
                               tempo_extra=tempo_extra, tempo_minimo=3.0):
            self._log(f"Ficheiro SRT gerado: {caminho}")
            messagebox.showinfo("", f"Legendas guardadas em:\n{caminho}")
        else:
            messagebox.showerror("", "Erro ao gerar ficheiro SRT")
    
    def gerar_video(self):
        if not self.pptx.apresentacao:
            return
        
        if not self.video.disponivel():
            messagebox.showerror("Erro", 
                "MoviePy não disponível!\n\n"
                "Para instalar:\n"
                "pip install moviepy imageio imageio-ffmpeg\n\n"
                "Também necessita FFmpeg instalado:\n"
                "Linux: sudo apt install ffmpeg\n"
                "macOS: brew install ffmpeg\n"
                "Windows: descarregar de ffmpeg.org")
            return
        
        # Recolher configurações primeiro
        try:
            tempo_extra = float(self.var_tempo_extra.get())
        except:
            tempo_extra = 0.5
        
        resolucao = self.var_resolucao.get()
        fps = self.var_fps.get()
        
        usar_traduzido = self.idioma.t("opt_lingua_traduzida") in self.var_idioma_video.get()
        legendas_video = self.var_legenda_video.get()
        idioma_leg = self.var_idioma_legendas.get()
        
        # Verificar áudios disponíveis
        tem_audios_orig = any(s.caminho_audio and os.path.exists(s.caminho_audio) 
                              for s in self.pptx.apresentacao.slides)
        tem_audios_trad = any(s.caminho_audio_traduzido and os.path.exists(s.caminho_audio_traduzido) 
                              for s in self.pptx.apresentacao.slides)
        
        # Calcular duração real lendo os ficheiros de áudio
        duracao_total = 0.0
        for slide in self.pptx.apresentacao.slides:
            caminho_audio = slide.caminho_audio if not usar_traduzido else (slide.caminho_audio_traduzido or slide.caminho_audio)
            if caminho_audio and os.path.exists(caminho_audio):
                try:
                    duracao = self.tts.obter_duracao(caminho_audio)
                    duracao_total += duracao + tempo_extra
                except:
                    duracao_total += 5 + tempo_extra  # Estimativa
            else:
                duracao_total += 5 + tempo_extra  # Estimativa
        
        # Obter novas opções de legendas
        pos_legenda_video = self.var_pos_legenda_video.get()
        linhas_legenda = self.var_linhas_legenda.get()
        traducao_ativa = self.var_traducao_ativa.get()
        idioma_destino = self.combo_idioma_destino.get() if traducao_ativa else "N/A"
        
        # Opções karaoke
        karaoke_ativo = self.var_karaoke.get()
        karaoke_modo = self.var_modo_karaoke.get()
        karaoke_cor = self.var_cor_karaoke.get()
        karaoke_transp = self.var_transp_karaoke.get()
        
        # Mostrar confirmação se ativo
        if self.var_confirmar_antes.get():
            resumo = "═══════════════════════════════════════\n"
            resumo += f"    PPTX Narrator v{APP_VERSAO} - GERAR VÍDEO\n"
            resumo += "═══════════════════════════════════════\n\n"
            
            resumo += f"📁 Apresentação: {os.path.basename(self.pptx.apresentacao.caminho)}\n"
            resumo += f"📊 Slides: {self.pptx.num_slides}\n"
            resumo += f"⏱️ Duração estimada: {int(duracao_total // 60)}m {int(duracao_total % 60)}s\n\n"
            
            resumo += "── VÍDEO ──\n"
            resumo += f"  • Resolução: {resolucao}\n"
            resumo += f"  • FPS: {fps}\n"
            resumo += f"  • Tempo extra/slide: {tempo_extra}s\n\n"
            
            resumo += "── ÁUDIO ──\n"
            resumo += f"  • Idioma áudio: {'Traduzido' if usar_traduzido else 'Original'}\n"
            resumo += f"  • Áudios originais: {'✅ Disponíveis' if tem_audios_orig else '❌ Não encontrados'}\n"
            resumo += f"  • Áudios traduzidos: {'✅ Disponíveis' if tem_audios_trad else '❌ Não encontrados'}\n\n"
            
            resumo += "── TRADUÇÃO ──\n"
            resumo += f"  • Tradução ativa: {'✅ Sim' if traducao_ativa else '❌ Não'}\n"
            if traducao_ativa:
                resumo += f"  • Idioma destino: {idioma_destino}\n"
            resumo += "\n"
            
            resumo += "── LEGENDAS NO VÍDEO ──\n"
            resumo += f"  • Legendas ativas: {'✅ Sim' if legendas_video else '❌ Não'}\n"
            if legendas_video:
                resumo += f"  • Idioma legendas: {idioma_leg}\n"
                resumo += f"  • Posição: {pos_legenda_video}\n"
                resumo += f"  • Linhas: {linhas_legenda}\n"
            resumo += f"  • Gerar ficheiro SRT: {'✅ Sim' if legendas_video else '❌ Não'}\n\n"
            
            resumo += "── MODO KARAOKE ──\n"
            resumo += f"  • Karaoke ativo: {'✅ Sim' if karaoke_ativo else '❌ Não'}\n"
            if karaoke_ativo:
                resumo += f"  • Modo exibição: {karaoke_modo}\n"
                resumo += f"  • Cor destaque: {karaoke_cor}\n"
                resumo += f"  • Transparência: {karaoke_transp}%\n"
            resumo += "\n"
            
            # Avisos
            if usar_traduzido and not tem_audios_trad:
                resumo += "⚠️ AVISO: Selecionou áudio traduzido mas não há\n"
                resumo += "   áudios traduzidos. Será usado áudio original.\n\n"
            
            if karaoke_ativo:
                resumo += "ℹ️ Modo karaoke pode demorar mais a gerar.\n\n"
            
            resumo += "═══════════════════════════════════════\n"
            resumo += "Continuar com estas configurações?"
            
            if not messagebox.askyesno("Confirmar Geração Vídeo", resumo):
                return
        
        # v1.9.3: Usar pasta estruturada com nome predefinido
        pasta_video = self._obter_pasta_estruturada("video")
        nome_base = self._obter_nome_base_apresentacao()
        nome_predefinido = f"{nome_base}.mp4"
        
        # Pedir caminho
        caminho = filedialog.asksaveasfilename(
            title=self.idioma.t("dlg_guardar_titulo"),
            initialdir=pasta_video,
            initialfile=nome_predefinido,
            defaultextension=".mp4",
            filetypes=[("Vídeo MP4", "*.mp4")]
        )
        if not caminho:
            return
        
        # Verificar se os slides têm áudio configurado
        # Se não tiverem, procurar na pasta de áudios estruturada ou na pasta original do PPTX
        pasta_audio = self._obter_pasta_estruturada("audio")
        pasta_pptx_original = os.path.dirname(self.pptx.apresentacao.caminho)
        
        audios_encontrados = 0
        for i in range(1, self.pptx.num_slides + 1):
            slide = self.pptx.obter_slide(i)
            if slide and not slide.caminho_audio:
                # Procurar áudio na pasta do vídeo ou na pasta do PPTX
                for pasta in [pasta_audio, pasta_pptx_original]:
                    caminho_audio = os.path.join(pasta, f"slide_{i:02d}.mp3")
                    if os.path.exists(caminho_audio):
                        from tts_engine import MotorTTS
                        duracao = MotorTTS.obter_duracao(caminho_audio)
                        self.pptx.definir_audio_slide(i, caminho_audio, duracao)
                        audios_encontrados += 1
                        self._log(f"Áudio encontrado para slide {i}: {caminho_audio}")
                        break
        
        if audios_encontrados > 0:
            self._log(f"Total de áudios encontrados: {audios_encontrados}")
        
        # Aplicar configurações de vídeo
        self.video.config.tempo_extra_slide = tempo_extra
        try:
            res = resolucao.split("x")
            self.video.config.largura = int(res[0])
            self.video.config.altura = int(res[1])
        except: pass
        try:
            self.video.config.fps = int(fps)
        except: pass
        
        # Configurar legendas no vídeo
        self.video.config.legendas_embutidas = legendas_video
        
        # Posição das legendas
        if "Sobrepor" in pos_legenda_video:
            self.video.config.legendas_posicao = "sobrepor"
        else:
            self.video.config.legendas_posicao = "separada"
        
        # Número de linhas por página
        try:
            self.video.config.legendas_linhas = int(linhas_legenda)
        except:
            self.video.config.legendas_linhas = 3
        
        # Determinar idioma das legendas
        if "Igual" in idioma_leg:
            self.video.config.legendas_usar_traducao = usar_traduzido
        elif "original" in idioma_leg.lower():
            self.video.config.legendas_usar_traducao = False
        else:
            self.video.config.legendas_usar_traducao = True
        
        # Configurações karaoke
        self.video.config.karaoke_ativo = self.var_karaoke.get()
        if self.video.config.karaoke_ativo:
            self.video.config.karaoke_modo = "scroll" if "Scroll" in self.var_modo_karaoke.get() else "pagina"
            self.video.config.karaoke_cor = self.var_cor_karaoke.get()
            self.video.config.karaoke_transparencia = self.var_transp_karaoke.get()
            # v1.7: Idioma do karaoke
            idioma_karaoke = self.var_idioma_karaoke.get()
            if "Igual" in idioma_karaoke:
                self.video.config.karaoke_usar_traducao = usar_traduzido
            elif "original" in idioma_karaoke.lower():
                self.video.config.karaoke_usar_traducao = False
            else:
                self.video.config.karaoke_usar_traducao = True
            # Ativar legendas embutidas para karaoke funcionar
            self.video.config.legendas_embutidas = True
        
        self._log(f"Áudio: {'traduzido' if usar_traduzido else 'original'}")
        if self.video.config.karaoke_ativo:
            self._log(f"Karaoke: {self.var_modo_karaoke.get()}, idioma: {self.var_idioma_karaoke.get()}, cor: {self.var_cor_karaoke.get()}, transp: {self.var_transp_karaoke.get()}%")
        elif legendas_video:
            self._log(f"Legendas: {idioma_leg}, posição: {pos_legenda_video}, linhas: {linhas_legenda}")
        
        self.a_processar = True
        self._atualizar_estado_botoes()
        self.tabview.set(self.idioma.t("tab_progresso"))
        
        def callback_progresso(atual, total, msg):
            self.after(0, lambda: self._atualizar_progresso(atual, msg))
            self.after(0, lambda: self._log(msg))
        
        def processar():
            self.video.definir_callback_progresso(callback_progresso)
            if self.video.gerar_video(self.pptx, caminho, usar_audio_traduzido=usar_traduzido):
                self.after(0, lambda: messagebox.showinfo("", self.idioma.t("msg_video_guardado", caminho)))
                self.after(0, lambda: self._registar_ficheiro_gerado(caminho, "Video"))
                
                # Gerar SRT automaticamente se opção ativa (v1.7)
                if self.var_gerar_srt_auto.get():
                    srt_path = caminho.replace('.mp4', '.srt')
                    tempo_extra = self.video.config.tempo_extra_slide
                    # Usar idioma correto para SRT
                    if self.video.config.karaoke_ativo:
                        usar_trad_srt = self.video.config.karaoke_usar_traducao
                    else:
                        usar_trad_srt = self.video.config.legendas_usar_traducao
                    if self.pptx.gerar_srt(srt_path, usar_traducao=usar_trad_srt,
                                           tempo_extra=tempo_extra, tempo_minimo=3.0):
                        self.after(0, lambda: self._log(f"Ficheiro SRT gerado: {srt_path}"))
            else:
                self.after(0, lambda: messagebox.showerror("", self.idioma.t("msg_erro_gerar", "Vídeo")))
            self.a_processar = False
            self.after(0, self._atualizar_estado_botoes)
        
        threading.Thread(target=processar, daemon=True).start()


def main():
    app = PPTXNarratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
