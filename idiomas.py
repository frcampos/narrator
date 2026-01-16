#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
idiomas.py - Sistema de traduções para a aplicação
"""

TRADUCOES = {
    "pt-PT": {
        "app_titulo": "PPTX Narrator - Narrador de Apresentações",
        "menu_ficheiro": "Ficheiro",
        "menu_abrir": "Abrir PPTX...",
        "menu_guardar_config": "Guardar Configurações",
        "menu_sair": "Sair",
        "menu_ajuda": "Ajuda",
        "menu_sobre": "Sobre",
        "menu_idioma": "Idioma",
        
        # Painel principal
        "tab_slides": "Slides",
        "tab_configuracoes": "Configurações",
        "tab_progresso": "Progresso",
        
        # Slides
        "btn_abrir_pptx": " Abrir Apresentação",
        "btn_gerar_audio": " Gerar Áudio",
        "btn_gerar_pptx": " Gerar PPTX com Áudio",
        "btn_gerar_video":  " Gerar Vídeo",
        "btn_preview": " Ouvir Preview",
        "btn_parar": " Parar",
        "lbl_slide": "Slide",
        "lbl_texto_narrar": "Texto para narrar:",
        "lbl_notas_originais": "Notas originais:",
        "lbl_texto_slide": "Texto do slide:",
        "lbl_nenhum_pptx": "Nenhuma apresentação aberta.\nClique em 'Abrir Apresentação' para começar.",
        
        # Configurações
        "frame_voz": "Configurações de Voz",
        "lbl_motor_tts": "Motor TTS:",
        "lbl_voz": "Voz:",
        "lbl_velocidade": "Velocidade:",
        "lbl_idioma_voz": "Idioma do Audio:",
        "motor_edge": "Edge TTS (Online - Alta Qualidade)",
        "lbl_resolucao": "Resolução do Vídeo:",
        "lbl_fps": "Frames por segundo:",
        "motor_offline": "Offline (pyttsx3 - Sem Internet)",
        
        "frame_icone": "Ãcone de Áudio no Slide",
        "lbl_mostrar_icone": "Mostrar ícone:",
        "lbl_posicao": "Posição:",
        "lbl_tamanho": "Tamanho (cm):",
        "pos_sup_dir": "Superior Direito",
        "pos_sup_esq": "Superior Esquerdo",
        "pos_inf_dir": "Inferior Direito",
        "pos_inf_esq": "Inferior Esquerdo",
        
        "frame_saida": "Opções de Saída",
        "lbl_pasta_saida": "Pasta de saída:",
        "btn_escolher_pasta": "Escolher...",
        "lbl_guardar_audios": "Guardar Audios separados:",
        "lbl_audio_junto_pptx": "Áudio junto ao PPTX:",
        
        "frame_traducao": "Tradução AutomAtica",
        "lbl_traducao_ativa": "Ativar tradução:",
        "lbl_motor_traducao": "Motor de tradução:",
        "lbl_idioma_destino": "Idioma destino:",
        "btn_traduzir_todos": "  Traduzir Todos",
        "lbl_texto_traduzido": "Texto traduzido:",
        "btn_gerar_audio_trad": " Gerar Áudio Traduzido",
        "msg_traduzindo": "A traduzir slide {0}...",
        "msg_traducao_concluida": "Tradução concluída",
        "lbl_idioma_video": "Idioma do Vídeo:",
        "opt_lingua_original": "Língua original",
        "opt_lingua_traduzida": "Língua traduzida",
        
        "frame_video": "Opções de Vídeo",
        "lbl_tempo_extra_slide": "Tempo extra por slide (s):",
        "lbl_transicao": "Duração da transição (s):",
        
        "frame_legendas": "Legendas e Tradução Visível",
        "lbl_legenda_slide": "Caixa de texto no slide:",
        "lbl_legenda_notas": "Nas notas do apresentador:",
        "lbl_legenda_video": "Legendas no Vídeo:",
        "lbl_gerar_srt": "Gerar ficheiro .srt:",
        "btn_gerar_srt": "📄 Gerar SRT",
        
        # Progresso
        "lbl_estado": "Estado:",
        "estado_pronto": "Pronto",
        "estado_processando": "A processar...",
        "estado_concluido": "Concluído!",
        "estado_erro": "Erro",
        "lbl_progresso_total": "Progresso total:",
        "lbl_slide_atual": "Slide atual:",
        "lbl_log": "Registo de atividade:",
        
        # Mensagens
        "msg_ficheiro_aberto": "Apresentação aberta: {0}",
        "msg_slides_encontrados": "{0} slides encontrados",
        "msg_gerando_audio": "A gerar Audio para slide {0}...",
        "msg_audio_gerado": "Áudio gerado para slide {0}",
        "msg_pptx_guardado": "PPTX guardado em: {0}",
        "msg_video_guardado": "Vídeo guardado em: {0}",
        "msg_erro_abrir": "Erro ao abrir ficheiro: {0}",
        "msg_erro_gerar": "Erro ao gerar: {0}",
        "msg_selecionar_pptx": "Por favor, abra uma apresentação primeiro.",
        "msg_sem_texto": "Slide {0} não tem texto para narrar.",
        "msg_tts_indisponivel": "Motor TTS não disponível. Verifique a ligação à  internet ou instale pyttsx3.",
        
        # DiAlogos
        "dlg_abrir_titulo": "Abrir Apresentação PowerPoint",
        "dlg_guardar_titulo": "Guardar Como",
        "dlg_pasta_titulo": "Escolher Pasta de Saída",
        "dlg_confirmar": "Confirmar",
        "dlg_cancelar": "Cancelar",
        "dlg_sim": "Sim",
        "dlg_nao": "Não",
        
        # Sobre
        "sobre_titulo": "Sobre PPTX Narrator",
        "sobre_texto": """PPTX Narrator v1.0

Aplicação para criação de apresentações PowerPoint com narração automAtica.

Desenvolvido para apoiar docentes na criação de recursos educativos acessíveis.

Funcionalidades:
â€¢ Extração automAtica de texto dos slides
â€¢ Geração de Audio com vozes naturais
â€¢ Inserção de Audio nos slides
â€¢ Criação de Vídeos com narração

Â© 2024 - Uso Educativo""",

        # Vozes
        "voz_feminina": "Feminina",
        "voz_masculina": "Masculino",
    },
    
    "pt-BR": {
        "app_titulo": "PPTX Narrator - Narrador de Apresentações",
        "menu_ficheiro": "Arquivo",
        "menu_abrir": "Abrir PPTX...",
        "menu_guardar_config": "Salvar Configurações",
        "menu_sair": "Sair",
        "menu_ajuda": "Ajuda",
        "menu_sobre": "Sobre",
        "menu_idioma": "Idioma",
        
        "tab_slides": "Slides",
        "tab_configuracoes": "Configurações",
        "tab_progresso": "Progresso",
        
        "btn_abrir_pptx": "ðŸ“‚ Abrir Apresentação",
        "btn_gerar_audio": "ðŸ”Š Gerar Áudio",
        "btn_gerar_pptx": "ðŸ“Š Gerar PPTX com Áudio",
        "btn_gerar_video": "ðŸŽ¬ Gerar Vídeo",
        "btn_preview": "â–¶ï¸ Ouvir Preview",
        "btn_parar": "â¹ï¸ Parar",
        "lbl_slide": "Slide",
        "lbl_texto_narrar": "Texto para narrar:",
        "lbl_notas_originais": "Notas originais:",
        "lbl_texto_slide": "Texto do slide:",
        "lbl_nenhum_pptx": "Nenhuma apresentação aberta.\nClique em 'Abrir Apresentação' para começar.",
        
        "frame_voz": "Configurações de Voz",
        "lbl_motor_tts": "Motor TTS:",
        "lbl_voz": "Voz:",
        "lbl_velocidade": "Velocidade:",
        "lbl_idioma_voz": "Idioma do Audio:",
        "motor_edge": "Edge TTS (Online - Alta Qualidade)",
        "lbl_resolucao": "Resolução do Vídeo:",
        "lbl_fps": "Frames por segundo:",
        "motor_offline": "Offline (pyttsx3 - Sem Internet)",
        
        "frame_icone": "Ãcone de Áudio no Slide",
        "lbl_mostrar_icone": "Mostrar ícone:",
        "lbl_posicao": "Posição:",
        "lbl_tamanho": "Tamanho (cm):",
        "pos_sup_dir": "Superior Direito",
        "pos_sup_esq": "Superior Esquerdo",
        "pos_inf_dir": "Inferior Direito",
        "pos_inf_esq": "Inferior Esquerdo",
        
        "frame_saida": "Opções de Saída",
        "lbl_pasta_saida": "Pasta de saída:",
        "btn_escolher_pasta": "Escolher...",
        "lbl_guardar_audios": "Salvar Audios separados:",
        "lbl_audio_junto_pptx": "Áudio junto ao PPTX:",
        
        "frame_traducao": "Tradução AutomAtica",
        "lbl_traducao_ativa": "Ativar tradução:",
        "lbl_motor_traducao": "Motor de tradução:",
        "lbl_idioma_destino": "Idioma destino:",
        "btn_traduzir_todos": "  Traduzir Todos",
        "lbl_texto_traduzido": "Texto traduzido:",
        "btn_gerar_audio_trad": "ðŸ”Š Gerar Áudio Traduzido",
        "msg_traduzindo": "Traduzindo slide {0}...",
        "msg_traducao_concluida": "Tradução concluída",
        "lbl_idioma_video": "Idioma do Vídeo:",
        "opt_lingua_original": "Língua original",
        "opt_lingua_traduzida": "Língua traduzida",
        
        "frame_video": "Opções de Vídeo",
        "lbl_tempo_extra_slide": "Tempo extra por slide (s):",
        "lbl_transicao": "Duração da transição (s):",
        
        "frame_legendas": "Legendas e Tradução Visível",
        "lbl_legenda_slide": "Caixa de texto no slide:",
        "lbl_legenda_notas": "Nas notas do apresentador:",
        "lbl_legenda_video": "Legendas no Vídeo:",
        "lbl_gerar_srt": "Gerar arquivo .srt:",
        "btn_gerar_srt": "ðŸ“„ Gerar SRT",
        
        "lbl_estado": "Estado:",
        "estado_pronto": "Pronto",
        "estado_processando": "Processando...",
        "estado_concluido": "Concluído!",
        "estado_erro": "Erro",
        "lbl_progresso_total": "Progresso total:",
        "lbl_slide_atual": "Slide atual:",
        "lbl_log": "Registro de atividade:",
        
        "msg_ficheiro_aberto": "Apresentação aberta: {0}",
        "msg_slides_encontrados": "{0} slides encontrados",
        "msg_gerando_audio": "Gerando Audio para slide {0}...",
        "msg_audio_gerado": "Áudio gerado para slide {0}",
        "msg_pptx_guardado": "PPTX salvo em: {0}",
        "msg_video_guardado": "Vídeo salvo em: {0}",
        "msg_erro_abrir": "Erro ao abrir arquivo: {0}",
        "msg_erro_gerar": "Erro ao gerar: {0}",
        "msg_selecionar_pptx": "Por favor, abra uma apresentação primeiro.",
        "msg_sem_texto": "Slide {0} não tem texto para narrar.",
        "msg_tts_indisponivel": "Motor TTS não disponível.",
        
        "dlg_abrir_titulo": "Abrir Apresentação PowerPoint",
        "dlg_guardar_titulo": "Salvar Como",
        "dlg_pasta_titulo": "Escolher Pasta de Saída",
        "dlg_confirmar": "Confirmar",
        "dlg_cancelar": "Cancelar",
        "dlg_sim": "Sim",
        "dlg_nao": "Não",
        
        "sobre_titulo": "Sobre PPTX Narrator",
        "sobre_texto": """PPTX Narrator v1.0

Aplicativo para criação de apresentações PowerPoint com narração automAtica.

Â© 2024 - Uso Educativo""",

        "voz_feminina": "Feminina",
        "voz_masculina": "Masculino",
    },
    
    "en": {
        "app_titulo": "PPTX Narrator - Presentation Narrator",
        "menu_ficheiro": "File",
        "menu_abrir": "Open PPTX...",
        "menu_guardar_config": "Save Settings",
        "menu_sair": "Exit",
        "menu_ajuda": "Help",
        "menu_sobre": "About",
        "menu_idioma": "Language",
        
        "tab_slides": "Slides",
        "tab_configuracoes": "Settings",
        "tab_progresso": "Progress",
        
        "btn_abrir_pptx": "ðŸ“‚ Open Presentation",
        "btn_gerar_audio": "ðŸ”Š Generate Audio",
        "btn_gerar_pptx": "ðŸ“Š Generate PPTX with Audio",
        "btn_gerar_video": "ðŸŽ¬ Generate Video",
        "btn_preview": "â–¶ï¸ Preview",
        "btn_parar": "â¹ï¸ Stop",
        "lbl_slide": "Slide",
        "lbl_texto_narrar": "Text to narrate:",
        "lbl_notas_originais": "Original notes:",
        "lbl_texto_slide": "Slide text:",
        "lbl_nenhum_pptx": "No presentation open.\nClick 'Open Presentation' to start.",
        
        "frame_voz": "Voice Settings",
        "lbl_motor_tts": "TTS Engine:",
        "lbl_voz": "Voice:",
        "lbl_velocidade": "Speed:",
        "lbl_idioma_voz": "Audio language:",
        "lbl_resolucao": "Video resolution:",
        "lbl_fps": "Frames per second:",
        "motor_edge": "Edge TTS (Online - High Quality)",
        "motor_offline": "Offline (pyttsx3 - No Internet)",
        
        "frame_icone": "Audio Icon on Slide",
        "lbl_mostrar_icone": "Show icon:",
        "lbl_posicao": "Position:",
        "lbl_tamanho": "Size (cm):",
        "pos_sup_dir": "Top Right",
        "pos_sup_esq": "Top Left",
        "pos_inf_dir": "Bottom Right",
        "pos_inf_esq": "Bottom Left",
        
        "frame_saida": "Output Options",
        "lbl_pasta_saida": "Output folder:",
        "btn_escolher_pasta": "Browse...",
        "lbl_guardar_audios": "Save separate audio files:",
        "lbl_audio_junto_pptx": "Audio with PPTX:",
        
        "frame_traducao": "Automatic Translation",
        "lbl_traducao_ativa": "Enable translation:",
        "lbl_motor_traducao": "Translation engine:",
        "lbl_idioma_destino": "Target language:",
        "btn_traduzir_todos": "  Translate All",
        "lbl_texto_traduzido": "Translated text:",
        "btn_gerar_audio_trad": "ðŸ”Š Generate Translated Audio",
        "msg_traduzindo": "Translating slide {0}...",
        "msg_traducao_concluida": "Translation completed",
        "lbl_idioma_video": "Video language:",
        "opt_lingua_original": "Original language",
        "opt_lingua_traduzida": "Translated language",
        
        "frame_video": "Video Options",
        "lbl_tempo_extra_slide": "Extra time per slide (s):",
        "lbl_transicao": "Transition duration (s):",
        
        "frame_legendas": "Subtitles and Visible Translation",
        "lbl_legenda_slide": "Text box on slide:",
        "lbl_legenda_notas": "In presenter notes:",
        "lbl_legenda_video": "Subtitles in video:",
        "lbl_gerar_srt": "Generate .srt file:",
        "btn_gerar_srt": "ðŸ“„ Generate SRT",
        
        "lbl_estado": "Status:",
        "estado_pronto": "Ready",
        "estado_processando": "Processing...",
        "estado_concluido": "Completed!",
        "estado_erro": "Error",
        "lbl_progresso_total": "Total progress:",
        "lbl_slide_atual": "Current slide:",
        "lbl_log": "Activity log:",
        
        "msg_ficheiro_aberto": "Presentation opened: {0}",
        "msg_slides_encontrados": "{0} slides found",
        "msg_gerando_audio": "Generating audio for slide {0}...",
        "msg_audio_gerado": "Audio generated for slide {0}",
        "msg_pptx_guardado": "PPTX saved to: {0}",
        "msg_video_guardado": "Video saved to: {0}",
        "msg_erro_abrir": "Error opening file: {0}",
        "msg_erro_gerar": "Error generating: {0}",
        "msg_selecionar_pptx": "Please open a presentation first.",
        "msg_sem_texto": "Slide {0} has no text to narrate.",
        "msg_tts_indisponivel": "TTS engine not available.",
        
        "dlg_abrir_titulo": "Open PowerPoint Presentation",
        "dlg_guardar_titulo": "Save As",
        "dlg_pasta_titulo": "Choose Output Folder",
        "dlg_confirmar": "Confirm",
        "dlg_cancelar": "Cancel",
        "dlg_sim": "Yes",
        "dlg_nao": "No",
        
        "sobre_titulo": "About PPTX Narrator",
        "sobre_texto": """PPTX Narrator v1.0

Application for creating PowerPoint presentations with automatic narration.

Â© 2024 - Educational Use""",

        "voz_feminina": "Female",
        "voz_masculina": "Male",
    },
    
    "es": {
        "app_titulo": "PPTX Narrator - Narrador de Presentaciones",
        "menu_ficheiro": "Archivo",
        "menu_abrir": "Abrir PPTX...",
        "menu_guardar_config": "Guardar ConfiguraciÃ³n",
        "menu_sair": "Salir",
        "menu_ajuda": "Ayuda",
        "menu_sobre": "Acerca de",
        "menu_idioma": "Idioma",
        
        "tab_slides": "Diapositivas",
        "tab_configuracoes": "ConfiguraciÃ³n",
        "tab_progresso": "Progreso",
        
        "btn_abrir_pptx": "ðŸ“‚ Abrir PresentaciÃ³n",
        "btn_gerar_audio": "ðŸ”Š Generar Audio",
        "btn_gerar_pptx": "ðŸ“Š Generar PPTX con Audio",
        "btn_gerar_video": "ðŸŽ¬ Generar Vídeo",
        "btn_preview": "â–¶ï¸ Vista Previa",
        "btn_parar": "â¹ï¸ Parar",
        "lbl_slide": "Diapositiva",
        "lbl_texto_narrar": "Texto para narrar:",
        "lbl_notas_originais": "Notas originales:",
        "lbl_texto_slide": "Texto de la diapositiva:",
        "lbl_nenhum_pptx": "Ninguna presentaciÃ³n abierta.\nHaga clic en 'Abrir PresentaciÃ³n' para comenzar.",
        
        "frame_voz": "ConfiguraciÃ³n de Voz",
        "lbl_motor_tts": "Motor TTS:",
        "lbl_voz": "Voz:",
        "lbl_velocidade": "Velocidad:",
        "lbl_idioma_voz": "Idioma del audio:",
        "lbl_resolucao": "ResoluciÃ³n del Vídeo:",
        "lbl_fps": "Fotogramas por segundo:",
        "motor_edge": "Edge TTS (Online - Alta Calidad)",
        "motor_offline": "Offline (pyttsx3 - Sin Internet)",
        
        "frame_icone": "Icono de Audio en Diapositiva",
        "lbl_mostrar_icone": "Mostrar icono:",
        "lbl_posicao": "PosiciÃ³n:",
        "lbl_tamanho": "TamaÃ±o (cm):",
        "pos_sup_dir": "Superior Derecha",
        "pos_sup_esq": "Superior Izquierda",
        "pos_inf_dir": "Inferior Derecha",
        "pos_inf_esq": "Inferior Izquierda",
        
        "frame_saida": "Opciones de Salida",
        "lbl_pasta_saida": "Carpeta de salida:",
        "btn_escolher_pasta": "Elegir...",
        "lbl_guardar_audios": "Guardar audios separados:",
        "lbl_audio_junto_pptx": "Audio junto al PPTX:",
        
        "frame_traducao": "Tradução Automática",
        "lbl_traducao_ativa": "Activar tradução:",
        "lbl_motor_traducao": "Motor de tradução:",
        "lbl_idioma_destino": "Idioma destino:",
        "btn_traduzir_todos": "  Traducir Todos",
        "lbl_texto_traduzido": "Texto traducido:",
        "btn_gerar_audio_trad": "ðŸ”Š Generar Audio Traducido",
        "msg_traduzindo": "Traduciendo diapositiva {0}...",
        "msg_traducao_concluida": "Tradução completada",
        "lbl_idioma_video": "Idioma del Vídeo:",
        "opt_lingua_original": "Idioma original",
        "opt_lingua_traduzida": "Idioma traducido",
        
        "frame_video": "Opciones de Vídeo",
        "lbl_tempo_extra_slide": "Tiempo extra por diapositiva (s):",
        "lbl_transicao": "DuraciÃ³n de la transiciÃ³n (s):",
        
        "frame_legendas": "Subtítulos y Tradução Visible",
        "lbl_legenda_slide": "Caja de texto en diapositiva:",
        "lbl_legenda_notas": "En notas del presentador:",
        "lbl_legenda_video": "Subtítulos en Vídeo:",
        "lbl_gerar_srt": "Generar archivo .srt:",
        "btn_gerar_srt": "ðŸ“„ Generar SRT",
        
        "lbl_estado": "Estado:",
        "estado_pronto": "Listo",
        "estado_processando": "Procesando...",
        "estado_concluido": "Â¡Completado!",
        "estado_erro": "Error",
        "lbl_progresso_total": "Progreso total:",
        "lbl_slide_atual": "Diapositiva actual:",
        "lbl_log": "Registro de actividad:",
        
        "msg_ficheiro_aberto": "PresentaciÃ³n abierta: {0}",
        "msg_slides_encontrados": "{0} diapositivas encontradas",
        "msg_gerando_audio": "Generando audio para diapositiva {0}...",
        "msg_audio_gerado": "Audio generado para diapositiva {0}",
        "msg_pptx_guardado": "PPTX guardado en: {0}",
        "msg_video_guardado": "Vídeo guardado en: {0}",
        "msg_erro_abrir": "Error al abrir archivo: {0}",
        "msg_erro_gerar": "Error al generar: {0}",
        "msg_selecionar_pptx": "Por favor, abra una presentaciÃ³n primero.",
        "msg_sem_texto": "Diapositiva {0} no tiene texto para narrar.",
        "msg_tts_indisponivel": "Motor TTS no disponible.",
        
        "dlg_abrir_titulo": "Abrir PresentaciÃ³n PowerPoint",
        "dlg_guardar_titulo": "Guardar Como",
        "dlg_pasta_titulo": "Elegir Carpeta de Salida",
        "dlg_confirmar": "Confirmar",
        "dlg_cancelar": "Cancelar",
        "dlg_sim": "Sí",
        "dlg_nao": "No",
        
        "sobre_titulo": "Acerca de PPTX Narrator",
        "sobre_texto": """PPTX Narrator v1.0

AplicaciÃ³n para crear presentaciones PowerPoint con narraciÃ³n automAtica.

Â© 2024 - Uso Educativo""",

        "voz_feminina": "Femenina",
        "voz_masculina": "Masculino",
    },
    
    "fr": {
        "app_titulo": "PPTX Narrator - Narrateur de PrÃ©sentations",
        "menu_ficheiro": "Fichier",
        "menu_abrir": "Ouvrir PPTX...",
        "menu_guardar_config": "Enregistrer Configuration",
        "menu_sair": "Quitter",
        "menu_ajuda": "Aide",
        "menu_sobre": "Ã€ propos",
        "menu_idioma": "Langue",
        
        "tab_slides": "Diapositives",
        "tab_configuracoes": "ParamÃ¨tres",
        "tab_progresso": "Progression",
        
        "btn_abrir_pptx": "ðŸ“‚ Ouvrir PrÃ©sentation",
        "btn_gerar_audio": "ðŸ”Š Gerer Audio",
        "btn_gerar_pptx": "ðŸ“Š GÃ©nÃ©rer PPTX avec Audio",
        "btn_gerar_video": "ðŸŽ¬ GÃ©nÃ©rer VidÃ©o",
        "btn_preview": "â–¶ï¸ Ã‰couter",
        "btn_parar": "â¹ï¸ Arrêter",
        "lbl_slide": "Diapositive",
        "lbl_texto_narrar": "Texte Ã  narrer:",
        "lbl_notas_originais": "Notes originales:",
        "lbl_texto_slide": "Texte de la diapositive:",
        "lbl_nenhum_pptx": "Aucune prÃ©sentation ouverte.\nCliquez sur 'Ouvrir PrÃ©sentation' pour commencer.",
        
        "frame_voz": "ParamÃ¨tres de Voix",
        "lbl_motor_tts": "Moteur TTS:",
        "lbl_voz": "Voix:",
        "lbl_velocidade": "Vitesse:",
        "lbl_idioma_voz": "Langue audio:",
        "lbl_resolucao": "RÃ©solution vidÃ©o:",
        "lbl_fps": "Images par seconde:",
        "motor_edge": "Edge TTS (En ligne - Haute QualitÃ©)",
        "motor_offline": "Hors ligne (pyttsx3 - Sans Internet)",
        
        "frame_icone": "IcÃ´ne Audio sur Diapositive",
        "lbl_mostrar_icone": "Afficher icÃ´ne:",
        "lbl_posicao": "Position:",
        "lbl_tamanho": "Taille (cm):",
        "pos_sup_dir": "Haut Droite",
        "pos_sup_esq": "Haut Gauche",
        "pos_inf_dir": "Bas Droite",
        "pos_inf_esq": "Bas Gauche",
        
        "frame_saida": "Options de Sortie",
        "lbl_pasta_saida": "Dossier de sortie:",
        "btn_escolher_pasta": "Choisir...",
        "lbl_guardar_audios": "Enregistrer audios sÃ©parÃ©s:",
        "lbl_audio_junto_pptx": "Audio avec PPTX:",
        
        "frame_traducao": "Traduction Automatique",
        "lbl_traducao_ativa": "Activer la traduction:",
        "lbl_motor_traducao": "Moteur de traduction:",
        "lbl_idioma_destino": "Langue cible:",
        "btn_traduzir_todos": "  Traduire Tout",
        "lbl_texto_traduzido": "Texte traduit:",
        "btn_gerar_audio_trad": "ðŸ”Š GÃ©nÃ©rer Audio Traduit",
        "msg_traduzindo": "Traduction de la diapositive {0}...",
        "msg_traducao_concluida": "Traduction terminÃ©e",
        "lbl_idioma_video": "Langue de la vidÃ©o:",
        "opt_lingua_original": "Langue originale",
        "opt_lingua_traduzida": "Langue traduite",
        
        "frame_video": "Options VidÃ©o",
        "lbl_tempo_extra_slide": "Temps extra par diapositive (s):",
        "lbl_transicao": "DurÃ©e de la transition (s):",
        
        "frame_legendas": "Sous-titres et Traduction Visible",
        "lbl_legenda_slide": "Zone de texte sur diapositive:",
        "lbl_legenda_notas": "Dans les notes du prÃ©sentateur:",
        "lbl_legenda_video": "Sous-titres dans la vidÃ©o:",
        "lbl_gerar_srt": "GÃ©nÃ©rer fichier .srt:",
        "btn_gerar_srt": "ðŸ“„ GÃ©nÃ©rer SRT",
        
        "lbl_estado": "Ã‰tat:",
        "estado_pronto": "Prêt",
        "estado_processando": "En cours...",
        "estado_concluido": "TerminÃ©!",
        "estado_erro": "Erreur",
        "lbl_progresso_total": "Progression totale:",
        "lbl_slide_atual": "Diapositive actuelle:",
        "lbl_log": "Journal d'activitÃ©:",
        
        "msg_ficheiro_aberto": "PrÃ©sentation ouverte: {0}",
        "msg_slides_encontrados": "{0} diapositives trouvÃ©es",
        "msg_gerando_audio": "GÃ©nÃ©ration audio pour diapositive {0}...",
        "msg_audio_gerado": "Audio gÃ©nÃ©rÃ© pour diapositive {0}",
        "msg_pptx_guardado": "PPTX enregistrÃ© dans: {0}",
        "msg_video_guardado": "VidÃ©o enregistrÃ©e dans: {0}",
        "msg_erro_abrir": "Erreur lors de l'ouverture du fichier: {0}",
        "msg_erro_gerar": "Erreur lors de la gÃ©nÃ©ration: {0}",
        "msg_selecionar_pptx": "Veuillez d'abord ouvrir une prÃ©sentation.",
        "msg_sem_texto": "Diapositive {0} n'a pas de texte Ã  narrer.",
        "msg_tts_indisponivel": "Moteur TTS non disponible.",
        
        "dlg_abrir_titulo": "Ouvrir PrÃ©sentation PowerPoint",
        "dlg_guardar_titulo": "Enregistrer Sous",
        "dlg_pasta_titulo": "Choisir Dossier de Sortie",
        "dlg_confirmar": "Confirmer",
        "dlg_cancelar": "Annuler",
        "dlg_sim": "Oui",
        "dlg_nao": "Non",
        
        "sobre_titulo": "Ã€ propos de PPTX Narrator",
        "sobre_texto": """PPTX Narrator v1.0

Application pour crÃ©er des prÃ©sentations PowerPoint avec narration automatique.

Â© 2024 - Usage Ã‰ducatif""",

        "voz_feminina": "FÃ©minin",
        "voz_masculina": "Masculin",
    }
}

# Vozes disponíveis por idioma
VOZES_EDGE = {
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # LÃNGUAS PRINCIPAIS (sempre visíveis)
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    "pt-PT": [
        ("pt-PT-RaquelNeural", "Raquel (Feminina)", "feminina"),
        ("pt-PT-FernandaNeural", "Fernanda (Feminina)", "feminina"),
        ("pt-PT-DuarteNeural", "Duarte (Masculino)", "masculino"),
    ],
    "pt-BR": [
        ("pt-BR-FranciscaNeural", "Francisca (Feminina)", "feminina"),
        ("pt-BR-AntonioNeural", "AntÃ³nio (Masculino)", "masculino"),
        ("pt-BR-ThalitaNeural", "Thalita (Feminina)", "feminina"),
    ],
    "en": [
        ("en-US-JennyNeural", "Jenny US (Female)", "feminina"),
        ("en-US-GuyNeural", "Guy US (Male)", "masculino"),
        ("en-US-AriaNeural", "Aria US (Female)", "feminina"),
        ("en-US-DavisNeural", "Davis US (Male)", "masculino"),
    ],
    "en-GB": [
        ("en-GB-SoniaNeural", "Sonia UK (Female)", "feminina"),
        ("en-GB-RyanNeural", "Ryan UK (Male)", "masculino"),
        ("en-GB-LibbyNeural", "Libby UK (Female)", "feminina"),
    ],
    "es": [
        ("es-ES-ElviraNeural", "Elvira ES (Femenina)", "feminina"),
        ("es-ES-AlvaroNeural", "Álvaro ES (Masculino)", "masculino"),
        ("es-MX-DaliaNeural", "Dalia MX (Femenina)", "feminina"),
        ("es-MX-JorgeNeural", "Jorge MX (Masculino)", "masculino"),
    ],
    "fr": [
        ("fr-FR-DeniseNeural", "Denise FR (FÃ©minin)", "feminina"),
        ("fr-FR-HenriNeural", "Henri FR (Masculin)", "masculino"),
        ("fr-CA-SylvieNeural", "Sylvie CA (FÃ©minin)", "feminina"),
        ("fr-CA-JeanNeural", "Jean CA (Masculin)", "masculino"),
    ],
    "de": [
        ("de-DE-KatjaNeural", "Katja (Weiblich)", "feminina"),
        ("de-DE-ConradNeural", "Conrad (MÃ¤nnlich)", "masculino"),
        ("de-AT-IngridNeural", "Ingrid AT (Weiblich)", "feminina"),
    ],
    "it": [
        ("it-IT-ElsaNeural", "Elsa (Femminile)", "feminina"),
        ("it-IT-DiegoNeural", "Diego (Maschile)", "masculino"),
        ("it-IT-IsabellaNeural", "Isabella (Femminile)", "feminina"),
    ],
    
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    # LÃNGUAS EUROPEIAS ADICIONAIS (v1.8 - mostrar mais)
    # â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
    "nl-NL": [
        ("nl-NL-ColetteNeural", "Colette (Vrouwelijk)", "feminina"),
        ("nl-NL-MaartenNeural", "Maarten (Mannelijk)", "masculino"),
        ("nl-NL-FennaNeural", "Fenna (Vrouwelijk)", "feminina"),
    ],
    "pl-PL": [
        ("pl-PL-AgnieszkaNeural", "Agnieszka (Kobieta)", "feminina"),
        ("pl-PL-MarekNeural", "Marek (MÄ™Å¼czyzna)", "masculino"),
        ("pl-PL-ZofiaNeural", "Zofia (Kobieta)", "feminina"),
    ],
    "ro-RO": [
        ("ro-RO-AlinaNeural", "Alina (Feminina)", "feminina"),
        ("ro-RO-EmilNeural", "Emil (Masculino)", "masculino"),
    ],
    "uk-UA": [
        ("uk-UA-PolinaNeural", "Polina (Ð–Ñ–Ð½Ð¾Ñ‡Ð°)", "feminina"),
        ("uk-UA-OstapNeural", "Ostap (Ð§Ð¾Ð»Ð¾Ð²Ñ–Ñ‡Ð°)", "masculino"),
    ],
    "el-GR": [
        ("el-GR-AthinaNeural", "Athina (Î“Ï…Î½Î±Î¹ÎºÎµÎ¯Î±)", "feminina"),
        ("el-GR-NestorasNeural", "Nestoras (Î‘Î½Î´ÏÎ¹ÎºÎ®)", "masculino"),
    ],
    
    # ════════════════════════════════════════════════════════
    # LÍNGUAS DO SUL DA ÁSIA (v2.0 - Urdu, Bengali, Hindi, etc.)
    # ════════════════════════════════════════════════════════
    "ur-PK": [
        ("ur-PK-UzmaNeural", "Uzma (Female)", "feminina"),
        ("ur-PK-AsadNeural", "Asad (Male)", "masculino"),
    ],
    "bn-BD": [
        ("bn-BD-NabanitaNeural", "Nabanita (Female)", "feminina"),
        ("bn-BD-PradeepNeural", "Pradeep (Male)", "masculino"),
    ],
    "bn-IN": [
        ("bn-IN-TanishaaNeural", "Tanishaa (Female)", "feminina"),
        ("bn-IN-BashkarNeural", "Bashkar (Male)", "masculino"),
    ],
    "hi-IN": [
        ("hi-IN-SwaraNeural", "Swara (Female)", "feminina"),
        ("hi-IN-MadhurNeural", "Madhur (Male)", "masculino"),
    ],
    "ta-IN": [
        ("ta-IN-PallaviNeural", "Pallavi (Female)", "feminina"),
        ("ta-IN-ValluvarNeural", "Valluvar (Male)", "masculino"),
    ],
    "te-IN": [
        ("te-IN-ShrutiNeural", "Shruti (Female)", "feminina"),
        ("te-IN-MohanNeural", "Mohan (Male)", "masculino"),
    ],
    "mr-IN": [
        ("mr-IN-AarohiNeural", "Aarohi (Female)", "feminina"),
        ("mr-IN-ManoharNeural", "Manohar (Male)", "masculino"),
    ],
    "gu-IN": [
        ("gu-IN-DhwaniNeural", "Dhwani (Female)", "feminina"),
        ("gu-IN-NiranjanNeural", "Niranjan (Male)", "masculino"),
    ],
    "kn-IN": [
        ("kn-IN-SapnaNeural", "Sapna (Female)", "feminina"),
        ("kn-IN-GaganNeural", "Gagan (Male)", "masculino"),
    ],
    "ml-IN": [
        ("ml-IN-SobhanaNeural", "Sobhana (Female)", "feminina"),
        ("ml-IN-MidhunNeural", "Midhun (Male)", "masculino"),
    ],
    "pa-IN": [
        ("pa-IN-GulNeural", "Gul (Female)", "feminina"),
        ("pa-IN-JaswinderNeural", "Jaswinder (Male)", "masculino"),
    ],
}

# Idiomas principais (sempre visíveis na UI)
IDIOMAS_PRINCIPAIS = {
    "pt-PT": "Português (Portugal)",
    "pt-BR": "Português (Brasil)",
    "en": "English (US)",
    "en-GB": "English (UK)",
    "es": "EspaÃ±ol",
    "fr": "Français",
    "de": "Deutsch",
    "it": "Italiano",
}

# Idiomas adicionais (mostrar com opção "mais línguas")
IDIOMAS_ADICIONAIS = {
    "nl-NL": "Nederlands",
    "pl-PL": "Polski",
    "ro-RO": "Romeno",
    "uk-UA": "Ucraniano",
    "el-GR": "Grego",
    # Línguas do Sul da Ásia v2.0
    "ur-PK": "Urdu (Pakistan)",
    "bn-BD": "Bengali (Bangladesh)",
    "bn-IN": "Bengali (India)",
    "hi-IN": "Hindi (India)",
    "ta-IN": "Tamil (India)",
    "te-IN": "Telugu (India)",
    "mr-IN": "Marathi (India)",
    "gu-IN": "Gujarati (India)",
    "kn-IN": "Kannada (India)",
    "ml-IN": "Malayalam (India)",
    "pa-IN": "Punjabi (India)",
}

# Todos os idiomas disponíveis (para compatibilidade)
IDIOMAS_DISPONIVEIS = {**IDIOMAS_PRINCIPAIS, **IDIOMAS_ADICIONAIS}


class GestorIdioma:
    """Gestor de traduções da aplicação"""
    
    def __init__(self, idioma: str = "pt-PT"):
        self.idioma_atual = idioma
        self.traducoes = TRADUCOES.get(idioma, TRADUCOES["pt-PT"])
    
    def definir_idioma(self, idioma: str):
        """Muda o idioma da aplicação"""
        if idioma in TRADUCOES:
            self.idioma_atual = idioma
            self.traducoes = TRADUCOES[idioma]
    
    def t(self, chave: str, *args) -> str:
        """ObtÃ©m tradução para uma chave"""
        texto = self.traducoes.get(chave, chave)
        if args:
            try:
                texto = texto.format(*args)
            except:
                pass
        return texto
    
    def obter_vozes(self) -> list:
        """ObtÃ©m lista de vozes para o idioma atual"""
        return VOZES_EDGE.get(self.idioma_atual, VOZES_EDGE["pt-PT"])
